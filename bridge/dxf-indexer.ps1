param(
  [string]$Root,
  [string]$IndexFile,
  [string]$StatusFile,
  [string]$RequestFile,
  [string]$Role = 'CONTROLLER'
)

$ErrorActionPreference = 'Continue'
$DXF_INDEXER_VERSION = '3.0.0'

function Get-PowerShellExe {
  try {
    if($IsWindows) {
      $candidate = Join-Path $PSHOME 'powershell.exe'
      if(Test-Path -LiteralPath $candidate){ return $candidate }
      return (Get-Command powershell.exe -ErrorAction Stop).Source
    }
  } catch {}
  try {
    $candidate = Join-Path $PSHOME 'pwsh'
    if(Test-Path -LiteralPath $candidate){ return $candidate }
  } catch {}
  return (Get-Command pwsh -ErrorAction Stop).Source
}

function Normalize-Key([string]$s) {
  if($null -eq $s){$s=''}
  $s.ToUpperInvariant() -replace '[^A-Z0-9]',''
}

function Shard-Key([string]$part) {
  $n = Normalize-Key $part
  if([string]::IsNullOrWhiteSpace($n)){ return '__' }
  if($n.Length -ge 3){ return $n.Substring(0,3) }
  while($n.Length -lt 3){ $n += '_' }
  $n
}

function Get-DxfFiles([string]$dir) {
  try {
    if($IsWindows) {
      return [IO.Directory]::EnumerateFiles($dir,'*.dxf',[IO.SearchOption]::TopDirectoryOnly)
    }
  } catch {}
  foreach($candidate in [IO.Directory]::EnumerateFiles($dir,'*',[IO.SearchOption]::TopDirectoryOnly)) {
    try {
      if([string]::Equals([IO.Path]::GetExtension($candidate),'.dxf',[StringComparison]::OrdinalIgnoreCase)){
        $candidate
      }
    } catch {}
  }
}

function Is-ReparseDirectory([string]$path) {
  try {
    $item = Get-Item -LiteralPath $path -Force -ErrorAction Stop
    return [bool]($item.Attributes -band [IO.FileAttributes]::ReparsePoint)
  } catch {
    return $false
  }
}

function Write-JsonFile([string]$Path,$Object) {
  try {
    $parent = Split-Path -Parent $Path
    if($parent){ New-Item -ItemType Directory -Path $parent -Force -ErrorAction SilentlyContinue | Out-Null }
    ($Object | ConvertTo-Json -Depth 12) | Set-Content -LiteralPath $Path -Encoding UTF8
  } catch {}
}

function Read-JsonFile([string]$Path) {
  try {
    if(Test-Path -LiteralPath $Path){
      return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json)
    }
  } catch {}
  $null
}

function Read-ControllerRequest {
  if([string]::IsNullOrWhiteSpace($RequestFile)){ return }
  if(-not(Test-Path -LiteralPath $RequestFile)){ throw "DXF indexer request file does not exist: $RequestFile" }
  Get-Content -LiteralPath $RequestFile -Raw -Encoding UTF8 | ConvertFrom-Json
}

$request = $null
try {
  $request = Read-ControllerRequest
  if($request){
    if($request.Root){$Root=[string]$request.Root}
    if($request.IndexFile){$IndexFile=[string]$request.IndexFile}
    if($request.StatusFile){$StatusFile=[string]$request.StatusFile}
    if($request.Role){$Role=[string]$request.Role}
  }
} finally {
  if($RequestFile){
    Remove-Item -LiteralPath $RequestFile -Force -ErrorAction SilentlyContinue
  }
}

if([string]::IsNullOrWhiteSpace($Root)){ throw 'DXF indexer Root is required.' }
if([string]::IsNullOrWhiteSpace($IndexFile)){ throw 'DXF indexer IndexFile is required.' }
if([string]::IsNullOrWhiteSpace($StatusFile)){ throw 'DXF indexer StatusFile is required.' }

$ROOT_PARENT = Split-Path -Parent $StatusFile
$LOG_FILE = Join-Path $ROOT_PARENT 'dxf-indexer.log'
$STATUS_PARENT = $ROOT_PARENT

function Write-Log([string]$message) {
  try {
    New-Item -ItemType Directory -Path $STATUS_PARENT -Force -ErrorAction SilentlyContinue | Out-Null
    ('['+(Get-Date).ToUniversalTime().ToString('o')+'] '+$message) | Add-Content -LiteralPath $LOG_FILE -Encoding UTF8
  } catch {}
}

function Write-Status($state,$message,[int]$count,[int]$errorCount,[datetime]$started,$finished=$null,[string]$current='',[int]$directoriesVisited=0,[int]$workers=1,[int]$workersCompleted=0,[string]$mode='FULL',[int]$workerCount=1,[double]$elapsedSeconds=0,[string]$workDirectory='',[string]$generatedUtc='') {
  $obj = [ordered]@{
    state=$state
    indexerVersion=$DXF_INDEXER_VERSION
    root=$Root
    message=$message
    mode=$mode
    filesFound=$count
    errors=$errorCount
    directoriesVisited=$directoriesVisited
    workers=$workerCount
    workersCompleted=$workersCompleted
    started=$started.ToString('o')
    finished=$(if($finished){$finished.ToString('o')}else{$null})
    generatedUtc=$(if($generatedUtc){$generatedUtc}else{$null})
    currentPath=$current
    elapsedSeconds=[math]::Round($elapsedSeconds,1)
    pid=$PID
    logFile=$LOG_FILE
    workDirectory=$workDirectory
  }
  Write-JsonFile -Path $StatusFile -Object ([pscustomobject]$obj)
}

if($Role -eq 'WORKER') {
  $started = (Get-Date).ToUniversalTime()
  $workerId = 0
  $workerCount = 1
  $workerStatusFile = ''
  $outputFile = ''
  try {
    if(-not $request) { $request = Read-JsonFile -Path $RequestFile }
    if(-not $request){ throw 'Worker request file could not be read.' }

    $workerId = [int]$request.WorkerId
    $workerCount = [int]$request.WorkerCount
    $workerStatusFile = [string]$request.StatusFile
    $outputFile = [string]$request.OutputFile
    $startItems = @($request.StartItems)

    $parent = Split-Path -Parent $outputFile
    if($parent){ New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    $writer = New-Object IO.StreamWriter($outputFile,$false,(New-Object Text.UTF8Encoding($false)))
    $writer.AutoFlush = $false
    $writer.WriteLine(('PartName'+[char]9+'File'+[char]9+'LastWriteUtc'+[char]9+'Length'))

    $count = 0
    $errors = 0
    $directories = 0
    $lastStatus = Get-Date
    $pending = New-Object 'System.Collections.Generic.Stack[string]'

    Write-Log "Worker $workerId/$workerCount starting with $($startItems.Count) work item(s)."
    Write-JsonFile -Path $workerStatusFile -Object ([pscustomobject]@{
      state='RUNNING'; workerId=$workerId; workerCount=$workerCount; filesFound=0; errors=0
      directoriesVisited=0; currentPath=''; started=$started.ToString('o'); finished=$null; pid=$PID
    })

    try {
      foreach($item in $startItems) {
        $path = [string]$item.path
        $isDirectOnly = [bool]$item.directOnly
        if([string]::IsNullOrWhiteSpace($path)){continue}
        if($isDirectOnly){
          try {
            foreach($file in (Get-DxfFiles -dir $path)){
              $count++
              $partName=[IO.Path]::GetFileNameWithoutExtension($file)
              $safePart=([string]$partName).Replace([char]9,' ')
              $safeFile=([string]$file).Replace([char]9,' ')
              $fi=$null
              try{$fi=Get-Item -LiteralPath $file -Force -ErrorAction Stop}catch{}
              $lastWrite=''
              $length=0
              if($fi){
                $lastWrite=$fi.LastWriteTimeUtc.ToString('o')
                $length=[int64]$fi.Length
              }
              $writer.WriteLine($safePart+[char]9+$safeFile+[char]9+$lastWrite+[char]9+$length)
            }
          } catch {
            $errors++
            Write-Log "Worker $workerId direct-file enumeration failed: $path :: $($_.Exception.Message)"
          }
          continue
        }
        try {
          if(Test-Path -LiteralPath $path -PathType Container){ $pending.Push($path) }
        } catch {
          $errors++
        }
      }

      while($pending.Count -gt 0){
        $dir=$pending.Pop()
        $directories++
        try {
          foreach($subdir in [IO.Directory]::EnumerateDirectories($dir,'*',[IO.SearchOption]::TopDirectoryOnly)){
            try {
              if(Is-ReparseDirectory -path $subdir){continue}
              $pending.Push($subdir)
            } catch {
              $errors++
            }
          }
        } catch {
          $errors++
          Write-Log "Worker $workerId subfolder enumeration failed: $dir :: $($_.Exception.Message)"
        }

        try {
          foreach($file in (Get-DxfFiles -dir $dir)){
            $count++
            $partName=[IO.Path]::GetFileNameWithoutExtension($file)
            $safePart=([string]$partName).Replace([char]9,' ')
            $safeFile=([string]$file).Replace([char]9,' ')
            $fi=$null
            try{$fi=Get-Item -LiteralPath $file -Force -ErrorAction Stop}catch{}
            $lastWrite=''
            $length=0
            if($fi){
              $lastWrite=$fi.LastWriteTimeUtc.ToString('o')
              $length=[int64]$fi.Length
            }
            $writer.WriteLine($safePart+[char]9+$safeFile+[char]9+$lastWrite+[char]9+$length)

            $now=Get-Date
            if(($now-$lastStatus).TotalSeconds -ge 2 -or ($count % 500) -eq 0){
              try{$writer.Flush()}catch{}
              $elapsed=((Get-Date).ToUniversalTime()-$started).TotalSeconds
              Write-JsonFile -Path $workerStatusFile -Object ([pscustomobject]@{
                state='RUNNING'; workerId=$workerId; workerCount=$workerCount; filesFound=$count; errors=$errors
                directoriesVisited=$directories; currentPath=$dir; started=$started.ToString('o'); finished=$null; pid=$PID
                elapsedSeconds=[math]::Round($elapsed,1)
              })
              $lastStatus=$now
            }
          }
        } catch {
          $errors++
          Write-Log "Worker $workerId DXF enumeration failed: $dir :: $($_.Exception.Message)"
        }
      }

      try{$writer.Flush()}catch{}
      if($writer){$writer.Dispose()}
      $finished=(Get-Date).ToUniversalTime()
      $elapsed=($finished-$started).TotalSeconds
      Write-JsonFile -Path $workerStatusFile -Object ([pscustomobject]@{
        state='COMPLETE'; workerId=$workerId; workerCount=$workerCount; filesFound=$count; errors=$errors
        directoriesVisited=$directories; currentPath=''; started=$started.ToString('o'); finished=$finished.ToString('o')
        pid=$PID; elapsedSeconds=[math]::Round($elapsed,1); outputFile=$outputFile
      })
      Write-Log "Worker $workerId complete. Files=$count Dirs=$directories Errors=$errors"
      exit 0
    } catch {
      if($writer){try{$writer.Dispose()}catch{}}
      $finished=(Get-Date).ToUniversalTime()
      Write-JsonFile -Path $workerStatusFile -Object ([pscustomobject]@{
        state='FAILED'; workerId=$workerId; workerCount=$workerCount; filesFound=$count; errors=($errors+1)
        directoriesVisited=$directories; currentPath=''; started=$started.ToString('o'); finished=$finished.ToString('o')
        pid=$PID; elapsedSeconds=[math]::Round(($finished-$started).TotalSeconds,1); message=$_.Exception.Message
        outputFile=$outputFile
      })
      exit 1
    }
  } catch {
    $finished=(Get-Date).ToUniversalTime()
    Write-JsonFile -Path $workerStatusFile -Object ([pscustomobject]@{
      state='FAILED'; workerId=$workerId; workerCount=$workerCount; filesFound=0; errors=1
      directoriesVisited=0; currentPath=''; started=$started.ToString('o'); finished=$finished.ToString('o'); pid=$PID
      message=$_.Exception.Message; outputFile=$outputFile
    })
    exit 1
  }
}

try {
  $started=(Get-Date).ToUniversalTime()
  $mode='FULL'
  $requestedWorkers=0
  if($request){
    if($request.Mode){$mode=[string]$request.Mode.ToUpperInvariant()}
    if($request.WorkerCount){$requestedWorkers=[int]$request.WorkerCount}
  }
  if($mode -notin @('FULL','REFRESH')){$mode='FULL'}

  if(-not(Test-Path -LiteralPath $Root)){throw "DXF root does not exist or is unavailable: $Root"}
  if(-not((Get-Item -LiteralPath $Root -Force).PSIsContainer)){throw "DXF root is not a folder: $Root"}

  $indexParent=Split-Path -Parent $IndexFile
  if($indexParent){New-Item -ItemType Directory -Path $indexParent -Force|Out-Null}

  $existing=Read-JsonFile -Path $IndexFile
  $oldFiles=0
  try{$oldFiles=[int]$existing.filesFound}catch{}

  $workDir=Join-Path $indexParent ('dxf-index-work-'+[Guid]::NewGuid().ToString('N'))
  $stageDir=Join-Path $indexParent ('dxf-index.tmp-'+[Guid]::NewGuid().ToString('N'))
  $finalDir=Join-Path $indexParent 'dxf-index'

  New-Item -ItemType Directory -Path $workDir -Force|Out-Null
  New-Item -ItemType Directory -Path $stageDir -Force|Out-Null

  Write-Status 'RUNNING' ($mode+' DXF index build started with controlled parallel workers.') $oldFiles 0 $started $null $Root 0 0 0 $mode 1 0 $workDir ''
  Write-Log "Controller starting. Mode=$mode Root=$Root WorkersRequested=$requestedWorkers ExistingFiles=$oldFiles"

  $items=New-Object System.Collections.ArrayList
  [void]$items.Add([pscustomobject]@{path=$Root;directOnly=$true})

  $topDirs=@()
  try{$topDirs=@([IO.Directory]::EnumerateDirectories($Root,'*',[IO.SearchOption]::TopDirectoryOnly))}catch{$topDirs=@()}
  foreach($top in $topDirs){
    try{
      if(Is-ReparseDirectory -path $top){continue}
      $children=@([IO.Directory]::EnumerateDirectories($top,'*',[IO.SearchOption]::TopDirectoryOnly))
      if($children.Count -eq 0){
        [void]$items.Add([pscustomobject]@{path=$top;directOnly=$false})
      }else{
        [void]$items.Add([pscustomobject]@{path=$top;directOnly=$true})
        foreach($child in $children){
          try{
            if(Is-ReparseDirectory -path $child){continue}
            $grand=@([IO.Directory]::EnumerateDirectories($child,'*',[IO.SearchOption]::TopDirectoryOnly))
            if($grand.Count -eq 0){
              [void]$items.Add([pscustomobject]@{path=$child;directOnly=$false})
            }else{
              [void]$items.Add([pscustomobject]@{path=$child;directOnly=$true})
              foreach($g in $grand){
                try{
                  if(-not(Is-ReparseDirectory -path $g)){[void]$items.Add([pscustomobject]@{path=$g;directOnly=$false})}
                }catch{}
              }
            }
          }catch{
            [void]$items.Add([pscustomobject]@{path=$child;directOnly=$false})
          }
        }
      }
    }catch{
      [void]$items.Add([pscustomobject]@{path=$top;directOnly=$false})
    }
  }

  $cpu=[Environment]::ProcessorCount
  $workerCount=$requestedWorkers
  if($workerCount -le 0){$workerCount=[math]::Min(8,[math]::Max(2,$cpu))}
  $workerCount=[math]::Max(1,[math]::Min(12,$workerCount))
  if($items.Count -lt $workerCount){$workerCount=[math]::Max(1,$items.Count)}
  if($workerCount -lt 1){$workerCount=1}

  $workerBuckets=@()
  for($i=0;$i -lt $workerCount;$i++){$workerBuckets += ,@()}
  for($i=0;$i -lt $items.Count;$i++){
    $bucket=$i % $workerCount
    $workerBuckets[$bucket] += $items[$i]
  }

  $procs=@()
  for($i=0;$i -lt $workerCount;$i++){
    $workerId=$i+1
    $reqPath=Join-Path $workDir ('worker-request-'+$workerId+'.json')
    $workerStatus=Join-Path $workDir ('worker-status-'+$workerId+'.json')
    $outputFile=Join-Path $workDir ('worker-'+$workerId+'.tsv')
    $workerReq=[ordered]@{
      Role='WORKER'
      Root=$Root
      WorkerId=$workerId
      WorkerCount=$workerCount
      StatusFile=$workerStatus
      OutputFile=$outputFile
      StartItems=@($workerBuckets[$i])
    }
    Write-JsonFile -Path $reqPath -Object ([pscustomobject]$workerReq)

    $psi=New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName=Get-PowerShellExe
    $psi.Arguments='-NoProfile -ExecutionPolicy Bypass -File "'+$PSCommandPath+'" -Role WORKER -RequestFile "'+$reqPath+'"'
    $psi.WorkingDirectory=$ROOT_PARENT
    $psi.UseShellExecute=$false
    $psi.CreateNoWindow=$true
    $psi.RedirectStandardOutput=$false
    $psi.RedirectStandardError=$false
    $proc=[System.Diagnostics.Process]::Start($psi)
    $procs += [pscustomobject]@{id=$workerId;process=$proc;statusFile=$workerStatus;outputFile=$outputFile}
    Write-Log "Worker $workerId launched PID=$($proc.Id) Items=$($workerBuckets[$i].Count)"
  }

  $done=$false
  while(-not $done){
    $completed=0
    $totalFiles=0
    $totalErrors=0
    $totalDirs=0
    $currentPaths=@()

    foreach($w in $procs){
      $st=Read-JsonFile -Path $w.statusFile
      if($st){
        try{$totalFiles += [int]$st.filesFound}catch{}
        try{$totalErrors += [int]$st.errors}catch{}
        try{$totalDirs += [int]$st.directoriesVisited}catch{}
        if($st.currentPath){$currentPaths += [string]$st.currentPath}
        if([string]$st.state -eq 'COMPLETE'){$completed++}
        elseif([string]$st.state -eq 'FAILED'){
          if($w.process.HasExited){$completed++}
        }
      } elseif($w.process.HasExited) {
        $completed++
      }
    }

    $elapsed=((Get-Date).ToUniversalTime()-$started).TotalSeconds
    $current=if($currentPaths.Count){$currentPaths -join ' | '}else{$Root}
    Write-Status 'RUNNING' ($mode+' index: '+$completed+'/'+$workerCount+' workers complete; '+$totalFiles+' DXF files discovered so far.') $totalFiles $totalErrors $started $null $current $totalDirs $completed $completed $mode $workerCount $elapsed $workDir ''
    $done=($completed -ge $workerCount)
    if(-not $done){Start-Sleep -Seconds 1}
  }

  foreach($w in $procs){
    try{$w.process.WaitForExit()}catch{}
  }

  $failedWorkers=@()
  $totalFiles=0
  $totalErrors=0
  $totalDirs=0
  foreach($w in $procs){
    $st=Read-JsonFile -Path $w.statusFile
    if(-not $st -or [string]$st.state -ne 'COMPLETE' -or $w.process.ExitCode -ne 0){
      $failedWorkers += $w.id
    }
    try{$totalFiles += [int]$st.filesFound}catch{}
    try{$totalErrors += [int]$st.errors}catch{}
    try{$totalDirs += [int]$st.directoriesVisited}catch{}
  }
  if($failedWorkers.Count -gt 0){
    throw ('One or more DXF workers failed: '+($failedWorkers -join ', ')+'. The previous index was preserved.')
  }

  Write-Status 'RUNNING' 'All DXF workers completed. Merging their results into searchable shards.' $totalFiles $totalErrors $started $null $Root $totalDirs $workerCount $workerCount $mode $workerCount (((Get-Date).ToUniversalTime()-$started).TotalSeconds) $workDir ''

  $writers=New-Object 'System.Collections.Generic.Dictionary[string,object]'
  try {
    foreach($w in $procs){
      $reader=$null
      try{
        if(-not(Test-Path -LiteralPath $w.outputFile)){continue}
        $reader=New-Object IO.StreamReader($w.outputFile,[Text.Encoding]::UTF8,$true)
        [void]$reader.ReadLine()
        while(($line=$reader.ReadLine()) -ne $null){
          $tab1=$line.IndexOf([char]9)
          if($tab1 -lt 1){continue}
          $partName=$line.Substring(0,$tab1)
          $shard=Shard-Key -part $partName
          $writer=$null
          if($writers.ContainsKey($shard)){
            $writer=$writers[$shard]
          }else{
            $shardPath=Join-Path $stageDir ($shard+'.tsv')
            $writer=New-Object IO.StreamWriter($shardPath,$false,(New-Object Text.UTF8Encoding($false)))
            $writer.AutoFlush=$false
            $writer.WriteLine(('PartName'+[char]9+'File'+[char]9+'LastWriteUtc'+[char]9+'Length'))
            $writers.Add($shard,$writer)
          }
          $writer.WriteLine($line)
        }
      } finally {
        if($reader){$reader.Dispose()}
      }
    }
  } finally {
    foreach($w in $writers.Values){try{$w.Flush();$w.Dispose()}catch{}}
    $writers.Clear()
  }

  $manifest=[ordered]@{
    schema='cl-ws-creator/dxf-index/3'
    indexerVersion=$DXF_INDEXER_VERSION
    root=[IO.Path]::GetFullPath($Root)
    mode=$mode
    filesFound=[int]$totalFiles
    errors=[int]$totalErrors
    directoriesVisited=[int]$totalDirs
    workers=[int]$workerCount
    generatedUtc=(Get-Date).ToUniversalTime().ToString('o')
    durationSeconds=[math]::Round(((Get-Date).ToUniversalTime()-$started).TotalSeconds,1)
    shardDirectory='dxf-index'
  }
  $manifestTmp=$IndexFile+'.tmp'
  Write-JsonFile -Path $manifestTmp -Object ([pscustomobject]$manifest)

  if(Test-Path -LiteralPath $finalDir){Remove-Item -LiteralPath $finalDir -Recurse -Force -ErrorAction SilentlyContinue}
  Move-Item -LiteralPath $stageDir -Destination $finalDir -Force
  Move-Item -LiteralPath $manifestTmp -Destination $IndexFile -Force

  if(-not(Test-Path -LiteralPath $IndexFile)){throw 'DXF index manifest was not created.'}
  if(-not(Test-Path -LiteralPath $finalDir)){throw 'DXF shard directory was not created.'}

  $finished=(Get-Date).ToUniversalTime()
  Write-Status 'COMPLETE' ($mode+' DXF index complete: '+$totalFiles+' file(s) across '+$totalDirs+' folder(s) using '+$workerCount+' worker(s).') $totalFiles $totalErrors $started $finished $Root $totalDirs $workerCount $workerCount $mode $workerCount (($finished-$started).TotalSeconds) $workDir $manifest.generatedUtc
  Write-Log "DXF index complete. Mode=$mode Workers=$workerCount Files=$totalFiles Dirs=$totalDirs Errors=$totalErrors"
  Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
  exit 0
} catch {
  $finished=(Get-Date).ToUniversalTime()
  $message=$_.Exception.Message
  Write-Log "DXF controller FAILED: $message"
  try{Write-Status 'FAILED' $message 0 1 $started $finished $Root 0 0 0 $mode 1 (($finished-$started).TotalSeconds) $workDir ''}catch{}
  if($stageDir -and (Test-Path -LiteralPath $stageDir)){Remove-Item -LiteralPath $stageDir -Recurse -Force -ErrorAction SilentlyContinue}
  Remove-Item -LiteralPath ($IndexFile+'.tmp') -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
  exit 1
}
