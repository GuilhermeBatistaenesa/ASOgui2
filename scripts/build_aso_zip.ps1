param(

    [string]$Version = ""

    [ValidateSet("major","minor","patch")]

    [string]$Bump = "patch"

)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$distDir = Join-Path $repoRoot "dist"
$buildDir = Join-Path $repoRoot "build"
$outDir = Join-Path $env:TEMP "aso_dist"
$workDir = Join-Path $env:TEMP "aso_build"
$vendorTessDir = Join-Path $repoRoot "vendor\tesseract"
$vendorTessExe = Join-Path $vendorTessDir "tesseract.exe"
$vendorTessAltDir = Join-Path $vendorTessDir "Tesseract-OCR"
$vendorTessAltExe = Join-Path $vendorTessAltDir "tesseract.exe"
$vendorPoppler = Join-Path $repoRoot "vendor\poppler\bin"

$pythonExe = $null
$pythonArgs = @()
if ($env:PYTHON_EXE -and (Test-Path $env:PYTHON_EXE)) {
    $pythonExe = $env:PYTHON_EXE
} else {
    $preferred = @(
        (Join-Path $env:LocalAppData "Programs\\Python\\Python314\\python.exe"),
        (Join-Path $env:LocalAppData "Programs\\Python\\Python313\\python.exe"),
        (Join-Path $env:LocalAppData "Programs\\Python\\Python312\\python.exe")
    )
    $pythonExe = $preferred | Where-Object { Test-Path $_ } | Select-Object -First 1
    $pyCmd = Get-Command py -ErrorAction SilentlyContinue
    if (-not $pythonExe -and $pyCmd) {
        $pythonExe = $pyCmd.Source
        $pythonArgs = @("-3")
    } elseif (-not $pythonExe) {
        $pyCmd = Get-Command python -ErrorAction SilentlyContinue
        if ($pyCmd -and ($pyCmd.Source -notlike "*WindowsApps*")) {
            $pythonExe = $pyCmd.Source
        }
    }
}
if (-not $pythonExe) {
    throw "Python not found. Set PYTHON_EXE env var to your python.exe."
}

$py = $pythonExe
& $py -m pip install --upgrade pip
& $py -m pip install -r (Join-Path $repoRoot "requirements.txt")
& $py -m pip install pyinstaller

# Install Playwright browsers locally into a known path (avoid OneDrive locks)
$pwInstallDir = Join-Path $env:TEMP "pw"
if (Test-Path $pwInstallDir) {
    Remove-Item $pwInstallDir -Recurse -Force
}
$env:PLAYWRIGHT_BROWSERS_PATH = $pwInstallDir
& $py -m playwright install chromium

# Build ASOgui onedir
Push-Location $repoRoot
if (Test-Path $outDir) { Remove-Item $outDir -Recurse -Force }
if (Test-Path $workDir) { Remove-Item $workDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
New-Item -ItemType Directory -Force -Path $workDir | Out-Null

& $py -m PyInstaller --onedir --name ASOgui --noconfirm --distpath $outDir --workpath $workDir main.py `
  --collect-submodules win32com `
  --hidden-import pythoncom `
  --hidden-import pywintypes `
  --collect-all playwright `
  --collect-submodules playwright `
  --collect-all pdf2image `
  --collect-all pytesseract `
  --collect-all dotenv `
  --collect-all PIL
Pop-Location

# Determine version (auto-bump when not provided)
if ([string]::IsNullOrWhiteSpace($Version)) {
    $vf1 = Join-Path $repoRoot "version.txt"
    $vf2 = Join-Path $repoRoot "VERSION.txt"
    if (Test-Path $vf1) { $Version = (Get-Content $vf1 -Raw).Trim() }
    elseif (Test-Path $vf2) { $Version = (Get-Content $vf2 -Raw).Trim() }
    else { $Version = "0.0.0" }

    if ($Version -notmatch ^(\d+)\.(\d+)\.(\d+)$) {
        throw "Invalid version format $Version. Use SemVer like 1.2.3 or pass -Version explicitly."
    }

    $major = [int]$matches[1]
    $minor = [int]$matches[2]
    $patch = [int]$matches[3]
    switch ($Bump) {
        "major" { $major++; $minor = 0; $patch = 0 }
        "minor" { $minor++; $patch = 0 }
        default { $patch++ }
    }
    $Version = "{0}.{1}.{2}" -f $major, $minor, $patch

    # Persist bumped version for next build
    Set-Content -Path $vf1 -Value $Version -Encoding ascii
    Set-Content -Path $vf2 -Value $Version -Encoding ascii
}
# Copy tools
$pkgDir = Join-Path $outDir "ASOgui"
$toolsDir = Join-Path $pkgDir "tools"
$toolsTessDir = Join-Path $toolsDir "tesseract"
$toolsPopplerDir = Join-Path $toolsDir "poppler\bin"

New-Item -ItemType Directory -Force -Path $toolsTessDir | Out-Null
New-Item -ItemType Directory -Force -Path $toolsPopplerDir | Out-Null

$tessSourceDir = $null
if (Test-Path $vendorTessExe) {
    $tessSourceDir = $vendorTessDir
} elseif (Test-Path $vendorTessAltExe) {
    $tessSourceDir = $vendorTessAltDir
} else {
    throw "Missing vendor tesseract at: $vendorTessExe (or $vendorTessAltExe)"
}
Copy-Item (Join-Path $tessSourceDir "*") $toolsTessDir -Recurse -Force

if (-not (Test-Path $vendorPoppler)) {
    throw "Missing vendor poppler bin at: $vendorPoppler"
}
Copy-Item (Join-Path $vendorPoppler "*") $toolsPopplerDir -Recurse -Force

# Copy .env if present
$envFile = Join-Path $repoRoot ".env"
if (Test-Path $envFile) {
    Copy-Item $envFile (Join-Path $pkgDir ".env") -Force
}

# Copy Playwright browsers (use robocopy to avoid long path issues)
$pwTarget = Join-Path $pkgDir "playwright-browsers"
New-Item -ItemType Directory -Force -Path $pwTarget | Out-Null
if (-not (Test-Path $pwInstallDir)) {
    throw "Playwright browsers not found at: $pwInstallDir"
}
$pwDest = $pwTarget
$null = New-Item -ItemType Directory -Force -Path $pwDest
$rc = (Start-Process -FilePath "robocopy.exe" -ArgumentList @(
    $pwInstallDir, $pwDest, "/E", "/NFL", "/NDL", "/NJH", "/NJS", "/NC", "/NS"
) -Wait -PassThru).ExitCode
if ($rc -ge 8) {
    throw "Robocopy failed copying Playwright browsers (exit code $rc)"
}

# Write version in package
Set-Content -Path (Join-Path $pkgDir "version.txt") -Value $Version -Encoding utf8

# Create ZIP
$zipName = "ASOgui_{0}.zip" -f $Version
$zipPath = Join-Path $distDir $zipName
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path (Join-Path $pkgDir "*") -DestinationPath $zipPath

Write-Host "ZIP created: $zipPath"

# Create SHA256
$hash = (Get-FileHash $zipPath -Algorithm SHA256).Hash.ToLower()
$shaName = "ASOgui_{0}.sha256" -f $Version
$shaPath = Join-Path $distDir $shaName
("{0}  {1}" -f $hash, (Split-Path $zipPath -Leaf)) | Set-Content -Path $shaPath -Encoding ascii
Write-Host "SHA256 created: $shaPath"

# Write latest.json (ASCII to avoid BOM)
$latestJson = @"
{
  `"version`": `"$Version`",
  `"package_filename`": `"$zipName`",
  `"sha256_filename`": `"$shaName`"
}
"@
$latestPath = Join-Path $distDir "latest.json"
$latestJson | Set-Content -Path $latestPath -Encoding ascii
Write-Host "latest.json updated: $latestPath"
