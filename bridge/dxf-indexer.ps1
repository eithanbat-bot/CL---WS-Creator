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

  Write-Status 'RUNNING' 'Walking the DXF server tree and indexing filenames. This can take a while on a very large server.' 0 0 $started

  $count=0
  $errors=0
  $writer=New-Object IO.StreamWriter($tmp,$false,[Text.Encoding]::UTF8)
  try{
    $writer.WriteLine(('PartName'+[char]9+'File'))
    $scanErrors=@()
    Get-ChildItem -LiteralPath $Root -Recurse -File -Filter '*.dxf' -ErrorVariable +scanErrors -ErrorAction SilentlyContinue |
      ForEach-Object{
        $count++
        $partName=$_.BaseName
        $line=([string]$partName).Replace([char]9,' ')+[char]9+([string]$_.FullName).Replace([char]9,' ')
        $writer.WriteLine($line)
        if(($count % 1000) -eq 0){
          $errors=$scanErrors.Count
          Write-Status 'RUNNING' ('Indexed '+$count+' DXF files...') $count $errors $started $null ([string]$_.DirectoryName)
        }
      }
    $errors=$scanErrors.Count
  }finally{
    $writer.Flush()
    $writer.Dispose()
  }

  Move-Item -LiteralPath $tmp -Destination $IndexFile -Force
  $finished=(Get-Date).ToUniversalTime().ToString('o')
  Write-Status 'COMPLETE' ('DXF index complete: '+$count+' file(s).') $count $errors $started $finished $Root
}catch{
  $finished=(Get-Date).ToUniversalTime().ToString('o')
  Write-Status 'FAILED' $_.Exception.Message 0 1 $started $finished $Root
  exit 1
}