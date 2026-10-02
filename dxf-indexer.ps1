param(
  [Parameter(Mandatory=$true)][string]$Root,
  [Parameter(Mandatory=$true)][string]$IndexFile,
  [Parameter(Mandatory=$true)][string]$StatusFile
)

$ErrorActionPreference = 'Continue'
$DXF_INDEXER_VERSION = '2.0.2'
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
    ($obj|ConvertTo-Json -Depth 8)|Set-Content -LiteralPath $StatusFile -Encoding UTF8
  }catch{}
}

$started=(Get-Date).ToUniversalTime().ToString('o')
try{
  Write-Log "DXF indexer $DXF_INDEXER_VERSION starting. Root=$Root IndexFile=$IndexFile StatusFile=$StatusFile"
  Write-Status 'RUNNING' 'DXF indexer process started; validating the Y:\ root.' 0 0 $started $null $Root 0

  if(-not(Test-Path -LiteralPath $Root)){
    throw "DXF root does not exist or is unavailable: $Root"
  }
  Write-Log 'Root Test-Path succeeded.'
  if(-not((Get-Item -LiteralPath $Root).PSIsContainer)){
    throw "DXF root is not a folder: $Root"
  }
  Write-Log 'Root Get-Item succeeded and is a directory.'

  $parent=Split-Path -Parent $IndexFile
  if($parent -and -not(Test-Path -LiteralPath $parent)){
    New-Item -ItemType Directory -Path $parent -Force|Out-Null
  }

  $shardFinal=Join-Path $parent 'dxf-index'
  $shardTmp=Join-Path $parent 'dxf-index.tmp'
  $manifestTmp=$IndexFile+'.tmp'
  Write-Status 'RUNNING' 'Preparing temporary DXF index storage.' 0 0 $started $null $Root 0
  Remove-Item -LiteralPath $shardTmp -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $manifestTmp -Force -ErrorAction SilentlyContinue
  New-Item -ItemType Directory -Path $shardTmp -Force|Out-Null
  Write-Log "Temporary shard directory ready: $shardTmp"

  $count=0
  $errors=0
  $directories=0
  $lastStatus=Get-Date
  $writers=New-Object 'System.Collections.Generic.Dictionary[string,object]'

  $pending=New-Object 'System.Collections.Generic.Stack[string]'
  $pending.Push($Root)
  Write-Status 'RUNNING' 'DXF indexer is walking Y:\ recursively. It reports a heartbeat before each directory enumeration.' 0 0 $started $null $Root 0
  Write-Log 'Initial root pushed onto work stack.'

  try{
    while($pending.Count -gt 0){
      $dir=$pending.Pop()
      $directories++

      Write-Status 'RUNNING' ('Enumerating subfolders in directory '+$directories+'. Indexed '+$count+' DXF files so far.') $count $errors $started $null $dir $directories
      Write-Log "Enumerating subfolders: $dir"

      try{
        foreach($subdir in [IO.Directory]::EnumerateDirectories($dir,'*',[IO.SearchOption]::TopDirectoryOnly)){
          try{
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

      Write-Status 'RUNNING' ('Enumerating DXF files in directory '+$directories+'. '+$count+' DXF files indexed so far.') $count $errors $started $null $dir $directories
      Write-Log "Enumerating DXF files: $dir"

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
            $writer=New-Object IO.StreamWriter($shardFile,$false,[Text.Encoding]::UTF8)
            $writer.WriteLine(('PartName'+[char]9+'File'))
            $writers.Add($key,$writer)
          }

          $writer.WriteLine((([string]$partName).Replace([char]9,' ')+[char]9+([string]$file).Replace([char]9,' ')))

          $now=Get-Date
          if(($now-$lastStatus).TotalSeconds -ge 2 -or ($count % 250) -eq 0){
            foreach($w in $writers.Values){try{$w.Flush()}catch{}}
            Write-Status 'RUNNING' ('Indexed '+$count+' DXF files across '+$directories+' folders.') $count $errors $started $null $dir $directories
            $lastStatus=$now
          }
        }
      }catch{
        $errors++
        Write-Log "DXF enumeration failed: $dir :: $($_.Exception.Message)"
      }

      $now=Get-Date
      if(($now-$lastStatus).TotalSeconds -ge 5){
        foreach($w in $writers.Values){try{$w.Flush()}catch{}}
        Write-Status 'RUNNING' ('Heartbeat: '+$directories+' folders visited; '+$count+' DXF files indexed.') $count $errors $started $null $dir $directories
        $lastStatus=$now
      }
    }
  }finally{
    foreach($w in $writers.Values){
      try{$w.Flush();$w.Dispose()}catch{}
    }
  }

  Write-Log "Directory walk complete. Folders=$directories DXF files=$count Errors=$errors"
  Write-Status 'RUNNING' 'DXF walk complete. Finalizing the index shards.' $count $errors $started $null $Root $directories

  if(Test-Path -LiteralPath $shardFinal){
    Remove-Item -LiteralPath $shardFinal -Recurse -Force
  }
  Move-Item -LiteralPath $shardTmp -Destination $shardFinal -Force

  $manifest=[ordered]@{
    schema='cl-ws-creator/dxf-index/2'
    indexerVersion=$DXF_INDEXER_VERSION
    root=$Root
    filesFound=[int]$count
    errors=[int]$errors
    generatedUtc=(Get-Date).ToUniversalTime().ToString('o')
    shardDirectory=[IO.Path]::GetFileName($shardFinal)
  }
  ($manifest|ConvertTo-Json -Depth 8)|Set-Content -LiteralPath $manifestTmp -Encoding UTF8
  Move-Item -LiteralPath $manifestTmp -Destination $IndexFile -Force

  $finished=(Get-Date).ToUniversalTime().ToString('o')
  Write-Log "DXF index complete. Folders=$directories DXF files=$count Errors=$errors"
  Write-Status 'COMPLETE' ('DXF index complete: '+$count+' file(s) across '+$directories+' folder(s).') $count $errors $started $finished $Root $directories
}catch{
  $finished=(Get-Date).ToUniversalTime().ToString('o')
  Write-Log "DXF indexer FAILED: $($_.Exception.Message)"
  Write-Status 'FAILED' $_.Exception.Message 0 1 $started $finished $Root 0
  exit 1
}
