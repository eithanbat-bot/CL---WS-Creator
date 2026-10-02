param(
  [string]$Root,
  [string]$IndexFile,
  [string]$StatusFile,
  [string]$RequestFile
)

$ErrorActionPreference = 'Continue'
$DXF_INDEXER_VERSION = '2.1.0'

if(-not [string]::IsNullOrWhiteSpace($RequestFile)){
  try{
    if(-not(Test-Path -LiteralPath $RequestFile)){throw "DXF indexer request file does not exist: $RequestFile"}
    $request=Get-Content -LiteralPath $RequestFile -Raw -Encoding UTF8|ConvertFrom-Json
    if([string]::IsNullOrWhiteSpace($Root)){$Root=[string]$request.Root}
    if([string]::IsNullOrWhiteSpace($IndexFile)){$IndexFile=[string]$request.IndexFile}
    if([string]::IsNullOrWhiteSpace($StatusFile)){$StatusFile=[string]$request.StatusFile}
  }catch{
    throw ("Could not read DXF indexer request file: "+$_.Exception.Message)
  }finally{
    Remove-Item -LiteralPath $RequestFile -Force -ErrorAction SilentlyContinue
  }
}

if([string]::IsNullOrWhiteSpace($Root)){throw 'DXF indexer Root is required.'}
if([string]::IsNullOrWhiteSpace($IndexFile)){throw 'DXF indexer IndexFile is required.'}
if([string]::IsNullOrWhiteSpace($StatusFile)){throw 'DXF indexer StatusFile is required.'}

$STATUS_PARENT = Split-Path -Parent $StatusFile
$LOG_FILE = Join-Path $STATUS_PARENT 'dxf-indexer.log'

function Normalize-Key([string]$s){
  if($null -eq $s){$s=''}
  $s.ToUpperInvariant() -replace '[^A-Z0-9]',''
}

function Shard-Key([string]$part){
  $n=Normalize-Key $part
  if([string]::IsNullOrWhiteSpace($n)){return '__'}
  if($n.Length -ge 2){return $n.Substring(0,2)}
  return $n+'_'
}

function Write-Log([string]$message){
  try{
    New-Item -ItemType Directory -Path $STATUS_PARENT -Force -ErrorAction SilentlyContinue|Out-Null
    ('['+(Get-Date).ToUniversalTime().ToString('o')+'] '+$message)|Add-Content -LiteralPath $LOG_FILE -Encoding UTF8
  }catch{}
}

function Write-Status($state,$message,$count,$errorCount,$started,$finished=$null,$current='',$directoriesVisited=0){
  $obj=[pscustomobject]@{
    state=$state
    indexerVersion=$DXF_INDEXER_VERSION
    root=$Root
    message=$message
    filesFound=[int]$count
    errors=[int]$errorCount
    directoriesVisited=[int]$directoriesVisited
    started=$started
    finished=$finished
    currentPath=$current
    pid=$PID
    logFile=$LOG_FILE
  }
  try{
    New-Item -ItemType Directory -Path $STATUS_PARENT -Force -ErrorAction SilentlyContinue|Out-Null
    ($obj|ConvertTo-Json -Depth 8)|Set-Content -LiteralPath $StatusFile -Encoding UTF8
  }catch{}
}

function Is-ReparseDirectory([string]$path){
  try{
    $item=Get-Item -LiteralPath $path -Force -ErrorAction Stop
    return [bool]($item.Attributes -band [IO.FileAttributes]::ReparsePoint)
  }catch{
    return $false
  }
}

$started=(Get-Date).ToUniversalTime().ToString('o')
try{
  New-Item -ItemType Directory -Path $STATUS_PARENT -Force -ErrorAction SilentlyContinue|Out-Null
  Write-Log "DXF indexer $DXF_INDEXER_VERSION starting. Root=$Root IndexFile=$IndexFile StatusFile=$StatusFile"
  Write-Status 'RUNNING' 'DXF indexer process started; validating the DXF root.' 0 0 $started $null $Root 0

  if(-not(Test-Path -LiteralPath $Root)){
    throw "DXF root does not exist or is unavailable: $Root"
  }
  if(-not((Get-Item -LiteralPath $Root -Force).PSIsContainer)){
    throw "DXF root is not a folder: $Root"
  }
  Write-Log 'DXF root validation succeeded.'

  $parent=Split-Path -Parent $IndexFile
  if($parent -and -not(Test-Path -LiteralPath $parent)){New-Item -ItemType Directory -Path $parent -Force|Out-Null}

  $shardFinal=Join-Path $parent 'dxf-index'
  $shardTmp=Join-Path $parent 'dxf-index.tmp'
  $manifestTmp=$IndexFile+'.tmp'

  Remove-Item -LiteralPath $shardTmp -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $manifestTmp -Force -ErrorAction SilentlyContinue
  New-Item -ItemType Directory -Path $shardTmp -Force|Out-Null
  Write-Log "Temporary index storage ready: $shardTmp"
  Write-Status 'RUNNING' 'Temporary DXF index storage is ready.' 0 0 $started $null $Root 0

  $count=0
  $errors=0
  $directories=0
  $lastStatus=Get-Date
  $pending=New-Object 'System.Collections.Generic.Stack[string]'
  $pending.Push($Root)

  while($pending.Count -gt 0){
    $dir=$pending.Pop()
    $directories++

    Write-Status 'RUNNING' ("Enumerating subfolders in directory "+$directories+". Indexed "+$count+" DXF files so far.") $count $errors $started $null $dir $directories
    Write-Log "Enumerating subfolders: $dir"

    try{
      foreach($subdir in [IO.Directory]::EnumerateDirectories($dir,'*',[IO.SearchOption]::TopDirectoryOnly)){
        try{
          if(Is-ReparseDirectory -path $subdir){
            Write-Log "Skipping reparse-point directory: $subdir"
            continue
          }
          $pending.Push($subdir)
        }catch{
          $errors++
          Write-Log "Could not queue subfolder: $subdir :: $($_.Exception.Message)"
        }
      }
    }catch{
      $errors++
      Write-Log "Subfolder enumeration failed: $dir :: $($_.Exception.Message)"
    }

    Write-Status 'RUNNING' ("Enumerating DXF files in directory "+$directories+". Indexed "+$count+" DXF files so far.") $count $errors $started $null $dir $directories
    Write-Log "Enumerating DXF files: $dir"

    try{
      $writers=New-Object 'System.Collections.Generic.Dictionary[string,object]'
      try{
        foreach($file in [IO.Directory]::EnumerateFiles($dir,'*.dxf',[IO.SearchOption]::TopDirectoryOnly)){
          $count++
          $partName=[IO.Path]::GetFileNameWithoutExtension($file)
          $key=Shard-Key $partName
          $writer=$null

          if($writers.ContainsKey($key)){
            $writer=$writers[$key]
          }else{
            $shardFile=Join-Path $shardTmp ($key+'.tsv')
            $existed=Test-Path -LiteralPath $shardFile
            $writer=New-Object IO.StreamWriter($shardFile,$true,[Text.Encoding]::UTF8)
            if(-not $existed){$writer.WriteLine(('PartName'+[char]9+'File'))}
            $writers.Add($key,$writer)
          }

          $safePart=([string]$partName).Replace([char]9,' ')
          $safeFile=([string]$file).Replace([char]9,' ')
          $writer.WriteLine($safePart+[char]9+$safeFile)

          $now=Get-Date
          if(($now-$lastStatus).TotalSeconds -ge 2 -or ($count % 250) -eq 0){
            foreach($w in $writers.Values){try{$w.Flush()}catch{}}
            Write-Status 'RUNNING' ("Indexed "+$count+" DXF files across "+$directories+" folders.") $count $errors $started $null $dir $directories
            $lastStatus=$now
          }
        }
      }finally{
        foreach($w in $writers.Values){try{$w.Flush();$w.Dispose()}catch{}}
        $writers.Clear()
      }
    }catch{
      $errors++
      Write-Log "DXF enumeration failed: $dir :: $($_.Exception.Message)"
    }

    $now=Get-Date
    if(($now-$lastStatus).TotalSeconds -ge 5){
      foreach($w in $writers.Values){try{$w.Flush()}catch{}}
      Write-Status 'RUNNING' ("Heartbeat: "+$directories+" folders visited; "+$count+" DXF files indexed.") $count $errors $started $null $dir $directories
      $lastStatus=$now
    }
  }

  foreach($w in $writers.Values){try{$w.Flush();$w.Dispose()}catch{}}
  $writers.Clear()

  Write-Log "Directory walk complete. Folders=$directories DXF files=$count Errors=$errors"
  Write-Status 'RUNNING' 'DXF walk complete. Finalizing the index.' $count $errors $started $null $Root $directories

  if(Test-Path -LiteralPath $shardFinal){Remove-Item -LiteralPath $shardFinal -Recurse -Force}
  Move-Item -LiteralPath $shardTmp -Destination $shardFinal -Force

  $manifest=[ordered]@{
    schema='cl-ws-creator/dxf-index/2'
    indexerVersion=$DXF_INDEXER_VERSION
    root=$Root
    filesFound=[int]$count
    errors=[int]$errors
    directoriesVisited=[int]$directories
    generatedUtc=(Get-Date).ToUniversalTime().ToString('o')
    shardDirectory=[IO.Path]::GetFileName($shardFinal)
  }
  ($manifest|ConvertTo-Json -Depth 8)|Set-Content -LiteralPath $manifestTmp -Encoding UTF8
  Move-Item -LiteralPath $manifestTmp -Destination $IndexFile -Force

  if(-not(Test-Path -LiteralPath $IndexFile)){throw 'DXF index manifest was not created.'}
  if(-not(Test-Path -LiteralPath $shardFinal)){throw 'DXF index shard directory was not created.'}

  $finished=(Get-Date).ToUniversalTime().ToString('o')
  Write-Log "DXF index complete. Folders=$directories DXF files=$count Errors=$errors"
  Write-Status 'COMPLETE' ("DXF index complete: "+$count+" file(s) across "+$directories+" folder(s).") $count $errors $started $finished $Root $directories
}catch{
  foreach($w in @($writers.Values)){try{$w.Dispose()}catch{}}
  $finished=(Get-Date).ToUniversalTime().ToString('o')
  $message=$_.Exception.Message
  Write-Log "DXF indexer FAILED: $message"
  Write-Status 'FAILED' $message 0 1 $started $finished $Root 0
  exit 1
}
