param(
  [Parameter(Mandatory=$true)][string]$Root,
  [Parameter(Mandatory=$true)][string]$IndexFile,
  [Parameter(Mandatory=$true)][string]$StatusFile
)

$ErrorActionPreference = 'Continue'
$DXF_INDEXER_VERSION = '2.0.1'

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

function Write-Status($state,$message,$count,$errorCount,$started,$finished=$null,$current=''){
  $obj=[pscustomobject]@{
    state=$state
    indexerVersion=$DXF_INDEXER_VERSION
    root=$Root
    message=$message
    filesFound=[int]$count
    errors=[int]$errorCount
    started=$started
    finished=$finished
    currentPath=$current
    pid=$PID
  }
  try{
    ($obj|ConvertTo-Json -Depth 8)|Set-Content -LiteralPath $StatusFile -Encoding UTF8
  }catch{}
}

$started=(Get-Date).ToUniversalTime().ToString('o')
try{
  if(-not(Test-Path -LiteralPath $Root)){throw "DXF root does not exist or is unavailable: $Root"}
  if(-not((Get-Item -LiteralPath $Root).PSIsContainer)){throw "DXF root is not a folder: $Root"}

  $parent=Split-Path -Parent $IndexFile
  if($parent -and -not(Test-Path -LiteralPath $parent)){
    New-Item -ItemType Directory -Path $parent -Force|Out-Null
  }

  $shardFinal=Join-Path $parent 'dxf-index'
  $shardTmp=Join-Path $parent 'dxf-index.tmp'
  $manifestTmp=$IndexFile+'.tmp'
  Remove-Item -LiteralPath $shardTmp -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $manifestTmp -Force -ErrorAction SilentlyContinue
  New-Item -ItemType Directory -Path $shardTmp -Force|Out-Null

  $count=0
  $errors=0
  $directories=0
  $lastStatus=Get-Date
  $writers=New-Object 'System.Collections.Generic.Dictionary[string,object]'

  Write-Status 'RUNNING' 'DXF indexer started. Walking the DXF server recursively; progress is reported while folders are being visited.' 0 0 $started $null $Root

  $pending=New-Object 'System.Collections.Generic.Stack[string]'
  $pending.Push($Root)

  try{
    while($pending.Count -gt 0){
      $dir=$pending.Pop()
      $directories++

      $now=Get-Date
      if(($now-$lastStatus).TotalSeconds -ge 2){
        Write-Status 'RUNNING' ('Scanning '+$directories+' folders; '+$count+' DXF files indexed so far.') $count $errors $started $null $dir
        $lastStatus=$now
      }

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
            Write-Status 'RUNNING' ('Indexed '+$count+' DXF files across '+$directories+' folders.') $count $errors $started $null $dir
            $lastStatus=$now
          }
        }
      }catch{
        $errors++
      }

      try{
        foreach($subdir in [IO.Directory]::EnumerateDirectories($dir,'*',[IO.SearchOption]::TopDirectoryOnly)){
          try{$pending.Push($subdir)}catch{$errors++}
        }
      }catch{
        $errors++
      }

      $now=Get-Date
      if(($now-$lastStatus).TotalSeconds -ge 10){
        foreach($w in $writers.Values){try{$w.Flush()}catch{}}
        Write-Status 'RUNNING' ('Still scanning. '+$directories+' folders visited; '+$count+' DXF files indexed.') $count $errors $started $null $dir
        $lastStatus=$now
      }
    }
  }finally{
    foreach($w in $writers.Values){
      try{$w.Flush();$w.Dispose()}catch{}
    }
  }

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
  Write-Status 'COMPLETE' ('DXF index complete: '+$count+' file(s) across '+$directories+' folder(s).') $count $errors $started $finished $Root
}catch{
  $finished=(Get-Date).ToUniversalTime().ToString('o')
  Write-Status 'FAILED' $_.Exception.Message 0 1 $started $finished $Root
  exit 1
}
