$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$venvPath = Join-Path $root ".venv-build"
if (-not (Test-Path $venvPath)) {
  python -m venv $venvPath
}

$python = Join-Path $venvPath "Scripts\\python.exe"
& $python -m pip install --upgrade pip
& $python -m pip install -r requirements.txt
& $python -m pip install pyinstaller

# Install Playwright browsers into project-local folder
$env:PLAYWRIGHT_BROWSERS_PATH = "0"
& $python -m playwright install chromium

$distDir = Join-Path $root "dist\\ASOgui"
$vendorTesseractDir = Join-Path $root "vendor\\tesseract"
$vendorTesseractExe = Join-Path $vendorTesseractDir "tesseract.exe"
$vendorTesseractAltDir = Join-Path $vendorTesseractDir "Tesseract-OCR"
$vendorTesseractAltExe = Join-Path $vendorTesseractAltDir "tesseract.exe"
$vendorPoppler = Join-Path $root "vendor\\poppler\\bin"

$tessSourceDir = $null
if (Test-Path $vendorTesseractExe) {
  $tessSourceDir = $vendorTesseractDir
} elseif (Test-Path $vendorTesseractAltExe) {
  $tessSourceDir = $vendorTesseractAltDir
} else {
  Write-Error "Missing vendor tesseract: $vendorTesseractExe (or $vendorTesseractAltExe)"
}
if (-not (Test-Path $vendorPoppler)) {
  Write-Error "Missing vendor poppler bin: $vendorPoppler"
}

& $python -m PyInstaller --onedir --name ASOgui `
  --collect-submodules win32com `
  --hidden-import pythoncom `
  --hidden-import pywintypes `
  --collect-all playwright `
  --collect-submodules playwright `
  --collect-all pdf2image `
  --collect-all pytesseract `
  --collect-all dotenv `
  --collect-all PIL `
  main.py

# Copy tools into dist
$toolsDir = Join-Path $distDir "tools"
New-Item -ItemType Directory -Force -Path (Join-Path $toolsDir "tesseract") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $toolsDir "poppler") | Out-Null
Copy-Item (Join-Path $tessSourceDir "*") -Destination (Join-Path $toolsDir "tesseract") -Recurse -Force
Copy-Item $vendorPoppler -Destination (Join-Path $toolsDir "poppler\\bin") -Recurse -Force

# Copy .env if exists, otherwise create a default
$envFile = Join-Path $root ".env"
if (Test-Path $envFile) {
  Copy-Item $envFile -Destination (Join-Path $distDir ".env") -Force
} else {
  @"
TESSERACT_PATH=tools\\tesseract\\tesseract.exe
POPPLER_PATH=tools\\poppler\\bin
"@ | Set-Content -Path (Join-Path $distDir ".env") -Encoding UTF8
}

# Copy Playwright browsers into dist
$pwLocal = Join-Path $root "playwright-browsers"
$pwAlt = Join-Path $root ".playwright"
$pwSrc = $null
if (Test-Path $pwLocal) { $pwSrc = $pwLocal }
elseif (Test-Path $pwAlt) { $pwSrc = $pwAlt }

if ($pwSrc) {
  Copy-Item $pwSrc -Destination (Join-Path $distDir "playwright-browsers") -Recurse -Force
} else {
  Write-Warning "Playwright browsers not found. Ensure PLAYWRIGHT_BROWSERS_PATH=0 and install chromium."
}

# VERSION.txt
$ver = "0.0.0"
$versionFile1 = Join-Path $root "version.txt"
$versionFile2 = Join-Path $root "VERSION.txt"
if (Test-Path $versionFile1) {
  $ver = (Get-Content $versionFile1 | Select-Object -First 1).Trim()
} elseif (Test-Path $versionFile2) {
  $ver = (Get-Content $versionFile2 | Select-Object -First 1).Trim()
}
Set-Content -Path (Join-Path $distDir "VERSION.txt") -Value $ver -Encoding UTF8

Write-Host "Build completed: $distDir"
