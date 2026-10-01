param(
  [Parameter(Mandatory=$true)][string]$Root,
  [Parameter(Mandatory=$true)][string]$IndexFile,
  [Parameter(Mandatory=$true)][string]$StatusFile
)

$ErrorActionPreference='Continue'

function Write-Status($state,$message,$count,$errorCount,$started,$finished=$null,$current=''){
  $obj=[pscustomobject]@{
    state=$state
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
  if(-not(Test-Path -LiteralPath $Root)){ throw "DXF root does not exist or is unavailable: $Root" }
  if(-not((Get-Item -LiteralPath $Root).PSIsContainer)){ throw "DXF root is not a folder: $Root" }

  $parent=Split-Path -Parent $IndexFile
  if($parent -and -not(Test-Path -LiteralPath $parent)){ New-Item -ItemType Directory -Path $parent -Force|Out-Null }
  $tmp=$IndexFile+'.tmp'
  Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue

  $count=0
  $errors=0
  $directories=0
  $lastStatus=Get-Date
  $lastHeartbeat=$lastStatus

  Write-Status 'RUNNING' 'DXF indexer started. Walking Y:\ recursively. First progress update will include the current folder even before the first DXF is found.' 0 0 $started $null $Root

  $writer=New-Object IO.StreamWriter($tmp,$false,[Text.Encoding]::UTF8)
  try{
    $writer.WriteLine(('PartName'+[char]9+'File'))

    # Use a streaming directory stack instead of Get-ChildItem -Recurse.
    # This avoids waiting for a massive recursive enumeration to materialize
    # and lets us publish progress even when no DXF has been found yet.
    $pending=New-Object 'System.Collections.Generic.Stack[string]'
    $pending.Push($Root)

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
          $line=([string]$partName).Replace([char]9,' ')+[char]9+([string]$file).Replace([char]9,' ')
          $writer.WriteLine($line)

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
        foreach($subdir in [IO.Directory]::EnumerateDirectories($dir,[string]'*',[IO.SearchOption]::TopDirectoryOnly)){
          try{$pending.Push($subdir)}catch{$errors++}
        }
      }catch{
        $errors++
      }

      $now=Get-Date
      if(($now-$lastHeartbeat).TotalSeconds -ge 10){
        $writer.Flush()
        Write-Status 'RUNNING' ('Still scanning. '+$directories+' folders visited; '+$count+' DXF files indexed.') $count $errors $started $null $dir
        $lastHeartbeat=$now
      }
    }
  }finally{
    $writer.Flush()
    $writer.Dispose()
  }

  Move-Item -LiteralPath $tmp -Destination $IndexFile -Force
  $finished=(Get-Date).ToUniversalTime().ToString('o')
  Write-Status 'COMPLETE' ('DXF index complete: '+$count+' file(s) across '+$directories+' folder(s).') $count $errors $started $finished $Root
}catch{
  $finished=(Get-Date).ToUniversalTime().ToString('o')
  Write-Status 'FAILED' $_.Exception.Message 0 1 $started $finished $Root
  exit 1
}
