$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonHome = "C:\Users\ZYB\AppData\Local\Programs\Python\Python314"
$OutputDir = Join-Path $ProjectRoot "outputs"
$WorkDir = Join-Path $ProjectRoot "work\pyinstaller-number-based"
$SpecDir = Join-Path $ProjectRoot "work"
$SourceDir = Join-Path $ProjectRoot "src"
$EntryScript = Join-Path $SourceDir "run_number_based_extractor.py"
$AppName = "$([char]0x6309)$([char]0x53F7)$([char]0x7801)$([char]0x63D0)$([char]0x53D6)$([char]0x6570)$([char]0x636E)"

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$env:PYTHONHOME = $PythonHome
$env:TCL_LIBRARY = Join-Path $PythonHome "tcl\tcl8.6"
$env:TK_LIBRARY = Join-Path $PythonHome "tcl\tk8.6"

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
