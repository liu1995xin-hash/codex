$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonHome = "C:\Users\ZYB\AppData\Local\Programs\Python\Python314"
$OutputDir = Join-Path $ProjectRoot "outputs"
$WorkDir = Join-Path $ProjectRoot "work\pyinstaller"
$SpecDir = Join-Path $ProjectRoot "work"
$SourceDir = Join-Path $ProjectRoot "src"
$EntryScript = Join-Path $SourceDir "run_expert_id_extractor.py"
$AppName = "$([char]0x63D0)$([char]0x53D6)$([char]0x7B26)$([char]0x5408)$([char]0x6761)$([char]0x4EF6)$([char]0x7684)$([char]0x4E13)$([char]0x5BB6)ID"

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$env:PYTHONHOME = $PythonHome

python -m PyInstaller `
  --onefile `
  --noconsole `
  --clean `
  --name $AppName `
  --distpath $OutputDir `
  --workpath $WorkDir `
  --specpath $SpecDir `
  --paths $SourceDir `
  --exclude-module pandas `
  --exclude-module numpy `
  --exclude-module matplotlib `
  --exclude-module scipy `
  --exclude-module PIL `
  --exclude-module pytest `
  $EntryScript

Write-Host "Build complete:" (Join-Path $OutputDir "$AppName.exe")
