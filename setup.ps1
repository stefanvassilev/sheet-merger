
# setup.ps1 - Install dependencies for Excel Sheet Merger
# Checks what is already installed; only asks you to download what is missing.
# Packages are installed from local .whl files (no internet access needed from CLI).
#
# Usage:
#   .\setup.ps1
#   .\setup.ps1 -PythonExe "C:\Python312\python.exe"

param(
    [string]$PythonExe = "python"
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$WheelsDir = Join-Path $ScriptDir "wheels"

Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "  Excel Sheet Merger - Setup" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# -- Check Python --------------------------------------------------------------

try {
    $pyVersionStr = & $PythonExe --version 2>&1
    Write-Host "Python found: $pyVersionStr" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Python not found in PATH." -ForegroundColor Red
    Write-Host ""
    Write-Host "Specify the full path to your Python executable:"
    Write-Host "  .\setup.ps1 -PythonExe 'C:\Python312\python.exe'"
    exit 1
}

# -- Detect Python version tag, architecture, MS Store flag --------------------

$pyTag     = & $PythonExe -c "import sys; print('cp{}{}'.format(sys.version_info.major, sys.version_info.minor))"
$pyArch    = & $PythonExe -c "import struct; print('win_amd64' if struct.calcsize('P')*8==64 else 'win32')"
$pyVer     = & $PythonExe -c "import sys; print('{}.{}'.format(sys.version_info.major, sys.version_info.minor))"
$pyExePath = & $PythonExe -c "import sys; print(sys.executable)"
$msStoreCheck = & $PythonExe -c "import sys; print('1' if ('WindowsApps' in sys.executable or 'PythonSoftwareFoundation' in sys.executable) else '0')"
$isMsStore = ($msStoreCheck -eq "1")

Write-Host "Detected: Python $pyVer  |  tag=$pyTag  |  arch=$pyArch" -ForegroundColor Green

if ($isMsStore) {
    Write-Host "  (Microsoft Store install)" -ForegroundColor Cyan
    if ($PythonExe -eq "python") {
        Write-Host ""
        Write-Host "  Tip: if 'python' opens the Store instead of running, use:" -ForegroundColor Yellow
        Write-Host "    .\setup.ps1 -PythonExe `"$pyExePath`"" -ForegroundColor Cyan
    }
}
Write-Host ""

# -- Check which packages are already installed --------------------------------

function Test-PkgInstalled {
    param($Exe, $Pkg)
    $local:ErrorActionPreference = "SilentlyContinue"
    & $Exe -m pip show $Pkg 2>&1 | Out-Null
    return ($LASTEXITCODE -eq 0)
}

$needPywin32       = -not (Test-PkgInstalled $PythonExe "pywin32")
$needWindowsCurses = -not (Test-PkgInstalled $PythonExe "windows-curses")

if (-not $needPywin32)       { Write-Host "pywin32        already installed - skipping." -ForegroundColor Green }
if (-not $needWindowsCurses) { Write-Host "windows-curses already installed - skipping." -ForegroundColor Green }

if (-not $needPywin32 -and -not $needWindowsCurses) {
    Write-Host ""
    Write-Host "All dependencies already satisfied." -ForegroundColor Green
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor White
    Write-Host "  $PythonExe merger.py" -ForegroundColor Gray
    Write-Host "  $PythonExe merger.py C:\path\to\files" -ForegroundColor Gray
    Write-Host ""
    exit 0
}

Write-Host ""

# -- Look for wheel files for anything still needed ----------------------------

if (-not (Test-Path $WheelsDir)) {
    New-Item -ItemType Directory -Path $WheelsDir | Out-Null
}

function Find-Wheel {
    param($Dir, $Prefix)
    $found = @(Get-ChildItem -Path $Dir -Filter "$Prefix-*.whl" -ErrorAction SilentlyContinue |
               Where-Object { $_.Name -match $pyTag -and $_.Name -match $pyArch })
    if ($found.Count -gt 0) { return $found[0].FullName }
    $found = @(Get-ChildItem -Path $Dir -Filter "$Prefix-*.whl" -ErrorAction SilentlyContinue)
    if ($found.Count -gt 0) { return $found[0].FullName }
    return $null
}

$pywin32Whl        = if ($needPywin32)       { Find-Wheel $WheelsDir "pywin32" }        else { $null }
$windowsCursesWhl  = if ($needWindowsCurses) { Find-Wheel $WheelsDir "windows_curses" } else { $null }

# -- Prompt to download any wheels that are missing ----------------------------

$downloadNeeded = @()
if ($needPywin32       -and -not $pywin32Whl)      { $downloadNeeded += "pywin32" }
if ($needWindowsCurses -and -not $windowsCursesWhl) { $downloadNeeded += "windows_curses" }

if ($downloadNeeded.Count -gt 0) {
    Write-Host "----------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "  ACTION REQUIRED: Download missing wheel file(s)" -ForegroundColor Yellow
    Write-Host "----------------------------------------------------------" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Your Python: $pyVer  ($pyTag-$pyArch)" -ForegroundColor White
    Write-Host ""
    Write-Host "  Save the file(s) into:" -ForegroundColor White
    Write-Host "    $WheelsDir" -ForegroundColor Cyan
    Write-Host ""

    $i = 1
    if ($downloadNeeded -contains "pywin32") {
        Write-Host "  $i. pywin32" -ForegroundColor White
        Write-Host "     Open in Chrome: https://pypi.org/project/pywin32/#files" -ForegroundColor Gray
        Write-Host "     Download: pywin32-*-$pyTag-$pyTag-$pyArch.whl" -ForegroundColor Cyan
        Write-Host ""
        $i++
    }
    if ($downloadNeeded -contains "windows_curses") {
        Write-Host "  $i. windows-curses" -ForegroundColor White
        Write-Host "     Open in Chrome: https://pypi.org/project/windows-curses/#files" -ForegroundColor Gray
        Write-Host "     Download: windows_curses-*-$pyTag-$pyTag-$pyArch.whl" -ForegroundColor Cyan
        Write-Host ""
    }

    Write-Host "  Then re-run this script." -ForegroundColor White
    Write-Host ""
    exit 0
}

# -- pip available? ------------------------------------------------------------

try {
    & $PythonExe -m pip --version | Out-Null
} catch {
    Write-Host "ERROR: pip is not available." -ForegroundColor Red
    Write-Host "Try: $PythonExe -m ensurepip"
    exit 1
}

# -- Install from local wheels -------------------------------------------------

if ($needPywin32) {
    Write-Host "Installing pywin32 ..." -ForegroundColor Yellow
    & $PythonExe -m pip install --no-index "$pywin32Whl"
    if ($LASTEXITCODE -ne 0) { Write-Host "ERROR: Failed to install pywin32." -ForegroundColor Red; exit 1 }
    Write-Host "pywin32 installed." -ForegroundColor Green
    Write-Host ""
}

if ($needWindowsCurses) {
    Write-Host "Installing windows-curses ..." -ForegroundColor Yellow
    & $PythonExe -m pip install --no-index "$windowsCursesWhl"
    if ($LASTEXITCODE -ne 0) { Write-Host "ERROR: Failed to install windows-curses." -ForegroundColor Red; exit 1 }
    Write-Host "windows-curses installed." -ForegroundColor Green
    Write-Host ""
}

# -- pywin32 post-install (COM registration) - only if we just installed it ----

if ($needPywin32) {
    Write-Host "Running pywin32 post-install (COM registration) ..." -ForegroundColor Yellow

    $postInstall = & $PythonExe -c "import sys,os,sysconfig; locs=[os.path.join(b,'Scripts','pywin32_postinstall.py') for b in (sys.prefix,sys.exec_prefix)]+[os.path.join(sysconfig.get_path('scripts',s),'pywin32_postinstall.py') for s in ('nt_user',) if sysconfig.get_path('scripts',s)]; [print(p) for p in locs if os.path.exists(p)][:1]" 2>$null

    if (-not $postInstall) {
        $postInstall = & $PythonExe -c "import site,os; p=os.path.join(site.getuserbase(),'Scripts','pywin32_postinstall.py'); print(p) if os.path.exists(p) else None" 2>$null
    }

    if ($postInstall -and (Test-Path $postInstall)) {
        & $PythonExe $postInstall -install
        if ($LASTEXITCODE -eq 0) {
            Write-Host "COM registration complete." -ForegroundColor Green
        } else {
            Write-Host "Warning: post-install returned non-zero (may be OK if already registered)." -ForegroundColor Yellow
        }
    } else {
        Write-Host "Warning: pywin32_postinstall.py not found - skipping." -ForegroundColor Yellow
    }
    Write-Host ""
}

# -- Verify imports ------------------------------------------------------------

Write-Host "Verifying imports ..." -ForegroundColor Yellow
$check = & $PythonExe -c "import win32com.client, curses; print('OK')" 2>&1
if ($check -match "OK") {
    Write-Host "All imports OK." -ForegroundColor Green
} else {
    Write-Host "Import check: $check" -ForegroundColor Yellow
    Write-Host "Try running merger.py anyway - it may still work." -ForegroundColor Gray
}

# -- Done ----------------------------------------------------------------------

Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "  Setup complete!" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Usage:" -ForegroundColor White
Write-Host "  $PythonExe merger.py                   # scan script directory" -ForegroundColor Gray
Write-Host "  $PythonExe merger.py C:\path\to\files   # scan a specific directory" -ForegroundColor Gray
Write-Host ""
