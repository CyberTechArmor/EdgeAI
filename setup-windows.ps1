# Fractionate Edge -- Windows Setup Script
# Usage: powershell -ExecutionPolicy Bypass -File setup-windows.ps1
#
# This script is idempotent -- safe to run multiple times.
# It installs all dependencies, downloads models, configures services,
# and starts the local AI backend.

$ErrorActionPreference = "Stop"

# -- Log File (write everything to a file so errors survive window closes) --

$FRACTIONATE_HOME = Join-Path $env:USERPROFILE ".fractionate"
if (-not (Test-Path $FRACTIONATE_HOME)) {
    New-Item -ItemType Directory -Path $FRACTIONATE_HOME -Force | Out-Null
}
$LOG_FILE = Join-Path $FRACTIONATE_HOME "setup.log"
# Start transcript -- captures all output to the log file
Start-Transcript -Path $LOG_FILE -Force | Out-Null

# -- Configuration ----------------------------------------------
$MODELS_DIR       = Join-Path $FRACTIONATE_HOME "models"
$FALCON_DIR       = Join-Path $MODELS_DIR "falcon3-7b-1.58bit"
$FLORENCE_DIR     = Join-Path $MODELS_DIR "florence-2-base"
$NGINX_DIR        = Join-Path $FRACTIONATE_HOME "nginx"
$SERVER_DIR       = Join-Path $FRACTIONATE_HOME "server"
$LOGS_DIR         = Join-Path $FRACTIONATE_HOME "logs"
$VENV_DIR         = Join-Path $FRACTIONATE_HOME "venv"
$BITNET_DIR       = Join-Path $FRACTIONATE_HOME "bitnet"
$DB_PATH          = Join-Path $FRACTIONATE_HOME "data.db"
$CONFIG_PATH      = Join-Path $FRACTIONATE_HOME "config.yaml"

$NGINX_VERSION    = "1.26.2"
$NGINX_URL        = "https://nginx.org/download/nginx-$NGINX_VERSION.zip"
$FALCON_REPO      = "microsoft/Falcon3-7B-1.58bit-GGUF"
$FALCON_HF_MODEL  = "tiiuae/Falcon3-7B-1.58bit"           # must match BitNet SUPPORTED_HF_MODELS
$BITNET_REPO      = "https://github.com/microsoft/BitNet.git"

# Colors and progress
$global:_StepNumber = 0
$global:_TotalSteps = 10

function Write-Step {
    param($msg)
    $global:_StepNumber++
    $pct = [math]::Floor(($global:_StepNumber / $global:_TotalSteps) * 100)
    Write-Host ""
    Write-Host "[$global:_StepNumber/$global:_TotalSteps] $msg ($pct%)" -ForegroundColor Cyan
    Write-Host ("=" * 50) -ForegroundColor DarkGray
}
function Write-Ok      { param($msg) Write-Host "   [OK] $msg" -ForegroundColor Green }
function Write-Warn    { param($msg) Write-Host "   [WARN] $msg" -ForegroundColor Yellow }
function Write-Err     { param($msg) Write-Host "   [ERROR] $msg" -ForegroundColor Red }
function Write-Info    { param($msg) Write-Host "   $msg" -ForegroundColor Gray }

function Format-Elapsed {
    param([System.Diagnostics.Stopwatch]$sw)
    $s = [math]::Floor($sw.Elapsed.TotalSeconds)
    if ($s -ge 60) { return "$([math]::Floor($s/60))m $($s%60)s" }
    return "${s}s"
}

function Wait-ProcessWithSpinner {
    param(
        [System.Diagnostics.Process]$Process,
        [string]$Activity
    )
    $spinner = @('|', '/', '-', '\')
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $i = 0
    while (-not $Process.HasExited) {
        $t = Format-Elapsed $sw
        Write-Host "`r   $($spinner[$i % 4]) $Activity... ($t)   " -NoNewline -ForegroundColor Yellow
        Start-Sleep -Milliseconds 250
        $i++
    }
    $sw.Stop()
    $t = Format-Elapsed $sw
    Write-Host "`r   [OK] $Activity ($t)          " -ForegroundColor Green
}

function Wait-JobWithSpinner {
    param(
        [System.Management.Automation.Job]$Job,
        [string]$Activity
    )
    $spinner = @('|', '/', '-', '\')
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $i = 0
    while ($Job.State -eq 'Running') {
        $t = Format-Elapsed $sw
        Write-Host "`r   $($spinner[$i % 4]) $Activity... ($t)   " -NoNewline -ForegroundColor Yellow
        Start-Sleep -Milliseconds 500
        $i++
    }
    $sw.Stop()
    $t = Format-Elapsed $sw
    Write-Host "`r   [OK] $Activity ($t)          " -ForegroundColor Green
    return (Receive-Job -Job $Job -AutoRemoveJob -Wait)
}

function Invoke-NativeCommand {
    # Run an external command with stderr merged into stdout without triggering
    # PowerShell's ErrorActionPreference=Stop on stderr lines.
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [string[]]$Arguments = @(),
        [switch]$ShowOutput,
        [switch]$LastLineOnly,
        [string]$FilterPattern
    )
    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        if ($ShowOutput) {
            & $FilePath @Arguments 2>&1 | ForEach-Object {
                $line = $_.ToString()
                if ($FilterPattern) {
                    if ($line -match $FilterPattern) { Write-Info $line }
                } else {
                    Write-Info $line
                }
            }
        } elseif ($LastLineOnly) {
            & $FilePath @Arguments 2>&1 | Select-Object -Last 1 | ForEach-Object { Write-Info $_.ToString() }
        } else {
            & $FilePath @Arguments 2>&1 | Out-Null
        }
    } finally {
        $ErrorActionPreference = $prevEAP
    }
}
    param(
        [string]$Uri,
        [string]$OutFile,
        [string]$Label
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Host "   - Downloading $Label..." -NoNewline -ForegroundColor Gray
    try {
        # Use BITS for large downloads (shows % in the console title)
        $bitsSupported = Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue
        if ($bitsSupported) {
            Start-BitsTransfer -Source $Uri -Destination $OutFile -Description $Label
        } else {
            Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing
        }
    } catch {
        # Fallback if BITS fails (e.g. HTTPS issues)
        Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing
    }
    $t = Format-Elapsed $sw
    $sizeMB = if (Test-Path $OutFile) { [math]::Round((Get-Item $OutFile).Length / 1MB) } else { "?" }
    Write-Host " done (${sizeMB} MB, $t)" -ForegroundColor Green
}

# -- Elevation Check --------------------------------------------

function Assert-Admin {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Warn "Not running as Administrator. Attempting to elevate..."
        # Auto-elevate: re-launch as admin with -NoExit so the window stays open on errors
        if ($PSCommandPath) {
            Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoExit -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        } else {
            Write-Err "Please run this script as Administrator."
            Write-Err "Right-click PowerShell -> Run as Administrator, then paste the script again."
            Read-Host "Press Enter to close"
        }
        exit
    }
    Write-Ok "Running as Administrator"
}

# -- Prerequisite Checks ---------------------------------------

function Test-Command {
    param([string]$Name)
    return $null -ne (Get-Command $Name -ErrorAction SilentlyContinue)
}

function Test-RealPython {
    # Windows has Store alias stubs for python.exe that open the Microsoft Store
    # instead of running Python. Detect and skip those.
    try {
        $cmd = Get-Command python -ErrorAction SilentlyContinue
        if (-not $cmd) { return $false }

        # Check if the python.exe is a Windows Store alias (AppExecAlias)
        if ($cmd.Source -and $cmd.Source -match 'WindowsApps') {
            return $false
        }

        # Actually try to run it -- Store alias triggers an error or opens Store
        $result = & python --version 2>&1
        if ($LASTEXITCODE -ne 0) { return $false }
        if ($result -match 'Python \d+\.\d+') { return $true }
        return $false
    } catch {
        return $false
    }
}

function Assert-Prerequisites {

    # Python -- must detect and skip the Windows Store alias
    if (Test-RealPython) {
        $pyVer = & python --version 2>&1
        Write-Ok "Python found: $pyVer"
        $versionMatch = [regex]::Match("$pyVer", '(\d+)\.(\d+)')
        if ($versionMatch.Success) {
            $major = [int]$versionMatch.Groups[1].Value
            $minor = [int]$versionMatch.Groups[2].Value
            if ($major -lt 3 -or ($major -eq 3 -and $minor -lt 10)) {
                Write-Err "Python 3.10+ is required. Found $pyVer"
                throw "Python 3.10+ is required. Found $pyVer"
            }
        }
    } else {
        Write-Info "Python not found (or only the Windows Store alias is present)."
        Write-Info "Attempting to install Python 3.12 via winget..."
        Install-WithWinget "Python.Python.3.12" "python"
    }

    # Git
    if (Test-Command "git") {
        Write-Ok "Git found: $(git --version)"
    } else {
        Write-Info "Git not found. Attempting to install..."
        Install-WithWinget "Git.Git" "git"
    }

    # CMake
    if (Test-Command "cmake") {
        Write-Ok "CMake found: $(cmake --version | Select-Object -First 1)"
    } else {
        Write-Info "CMake not found. Attempting to install..."
        Install-WithWinget "Kitware.CMake" "cmake"
    }

    # Clang -- BitNet requires Clang or GCC, NOT MSVC
    if (Test-Command "clang") {
        Write-Ok "Clang found: $(clang --version | Select-Object -First 1)"
    } else {
        Write-Info "Clang not found. Installing LLVM..."
        Install-LLVM
    }

    # Visual Studio Build Tools -- needed for linker and Windows SDK
    $vsWhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vsWhere) {
        $vsPath = & $vsWhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
        if ($vsPath) {
            Write-Ok "Visual Studio Build Tools found"
        } else {
            Write-Info "Installing Visual Studio Build Tools (for linker/SDK)..."
            Install-VsBuildTools
        }
    } else {
        Write-Info "Installing Visual Studio Build Tools (for linker/SDK)..."
        Install-VsBuildTools
    }

    # pip
    if (Test-Command "pip") {
        Write-Ok "pip found"
    } else {
        Write-Info "Installing pip..."
        & python -m ensurepip --upgrade 2>$null
    }
}

function Install-WithWinget {
    param([string]$PackageId, [string]$CommandName)

    if (Test-Command "winget") {
        Write-Info "Installing $PackageId via winget (silent)..."
        winget install --id $PackageId --accept-source-agreements --accept-package-agreements --silent -e
        # Refresh PATH
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        if (Test-Command $CommandName) {
            Write-Ok "$CommandName installed successfully"
        } else {
            Write-Err "Failed to install $CommandName. Please install manually and re-run this script."
            throw "Failed to install $CommandName"
        }
    } else {
        Write-Err "$CommandName is required but not found, and winget is not available."
        Write-Err "Please install $CommandName manually and re-run this script."
        throw "$CommandName is required but not found, and winget is not available"
    }
}

function Install-LLVM {
    # Use winget with --silent (forces NSIS /S flag so no GUI wizard pops up)
    if (Test-Command "winget") {
        Write-Info "Installing LLVM via winget (silent mode)..."
        $llvmJob = Start-Job -ScriptBlock {
            winget install --id LLVM.LLVM --accept-source-agreements --accept-package-agreements --silent -e 2>&1
        }
        Wait-JobWithSpinner -Job $llvmJob -Activity "Installing LLVM/Clang"
    } else {
        Write-Err "winget is required to install LLVM. Please install LLVM manually."
        throw "Cannot install LLVM without winget"
    }

    # LLVM's silent installer does NOT add to PATH by default -- fix that
    $llvmBin = "$env:ProgramFiles\LLVM\bin"
    if (Test-Path $llvmBin) {
        $env:Path = "$llvmBin;$env:Path"
        $machinePath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
        if ($machinePath -notlike "*LLVM*") {
            [System.Environment]::SetEnvironmentVariable("Path", "$machinePath;$llvmBin", "Machine")
            Write-Info "Added LLVM to system PATH"
        }
    }

    # Refresh and verify
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
    if (Test-Command "clang") {
        Write-Ok "LLVM/Clang installed: $(clang --version | Select-Object -First 1)"
    } else {
        Write-Err "LLVM installed but clang not found in PATH."
        Write-Err "Expected at: $llvmBin\clang.exe"
        throw "LLVM installation failed - clang not in PATH"
    }
}

function Install-VsBuildTools {
    $installerUrl = "https://aka.ms/vs/17/release/vs_BuildTools.exe"
    $installerPath = Join-Path $env:TEMP "vs_BuildTools.exe"

    Invoke-DownloadWithProgress -Uri $installerUrl -OutFile $installerPath -Label "VS Build Tools installer"

    $proc = Start-Process -FilePath $installerPath -ArgumentList `
        "--quiet", "--wait", "--norestart",
        "--add", "Microsoft.VisualStudio.Workload.VCTools",
        "--includeRecommended" `
        -PassThru -NoNewWindow
    Wait-ProcessWithSpinner -Process $proc -Activity "Installing Visual Studio Build Tools"

    Remove-Item $installerPath -ErrorAction SilentlyContinue
}

# -- Directory Structure ----------------------------------------

function Initialize-Directories {

    $dirs = @($FRACTIONATE_HOME, $MODELS_DIR, $FALCON_DIR, $FLORENCE_DIR,
              $NGINX_DIR, $SERVER_DIR, $LOGS_DIR, $VENV_DIR)

    foreach ($dir in $dirs) {
        if (-not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            Write-Info "Created: $dir"
        }
    }

    Write-Ok "Directory structure ready"
}

# -- BitNet Build -----------------------------------------------

function Build-BitNet {

    # Check if already built
    $possiblePaths = @(
        (Join-Path $BITNET_DIR "build\bin\Release\llama-server.exe"),
        (Join-Path $BITNET_DIR "build\bin\llama-server.exe"),
        (Join-Path $BITNET_DIR "build\Release\bin\llama-server.exe")
    )
    foreach ($p in $possiblePaths) {
        if (Test-Path $p) {
            Write-Ok "llama-server already built: $p"
            return
        }
    }

    # Clone
    if (-not (Test-Path (Join-Path $BITNET_DIR ".git"))) {
        Write-Info "Cloning BitNet repository..."
        git clone --recursive $BITNET_REPO $BITNET_DIR
    } else {
        Write-Info "BitNet repo exists, pulling latest..."
        Push-Location $BITNET_DIR
        git pull
        Pop-Location
    }

    $venvPython = Join-Path $VENV_DIR "Scripts\python.exe"
    $pipExe = Join-Path $VENV_DIR "Scripts\pip.exe"

    # Install BitNet's own Python dependencies (gguf-py, huggingface-cli)
    $ggufPy = Join-Path $BITNET_DIR "3rdparty\llama.cpp\gguf-py"
    if (Test-Path $ggufPy) {
        Write-Info "Installing BitNet Python dependencies..."
        Invoke-NativeCommand -FilePath $pipExe -Arguments @("install", "-e", $ggufPy) -LastLineOnly
    }
    $bitnetReqs = Join-Path $BITNET_DIR "requirements.txt"
    if (Test-Path $bitnetReqs) {
        Invoke-NativeCommand -FilePath $pipExe -Arguments @("install", "-r", $bitnetReqs) -LastLineOnly
    }
    Invoke-NativeCommand -FilePath $pipExe -Arguments @("install", "huggingface_hub") -LastLineOnly

    # Use BitNet's setup_env.py -- it handles: download model, generate
    # optimized kernels, cmake configure + build, and GGUF conversion.
    # --hf-repo must be a model from SUPPORTED_HF_MODELS (NOT the git URL).
    $setupScript = Join-Path $BITNET_DIR "setup_env.py"
    if (Test-Path $setupScript) {
        Write-Info "Running BitNet setup_env.py (download model + build)..."
        Write-Info "Model: $FALCON_HF_MODEL | Quantization: i2_s"
        Push-Location $BITNET_DIR
        Invoke-NativeCommand -FilePath $venvPython -Arguments @("setup_env.py", "--hf-repo", $FALCON_HF_MODEL, "-q", "i2_s") `
            -ShowOutput -FilterPattern '(Downloading|Converting|Compiling|Building|cmake|error|warning|100%|INFO|model)'
        Pop-Location

        # setup_env.py places the GGUF model under bitnet/models/
        # Copy it to our FALCON_DIR so the rest of the pipeline finds it.
        $bitnetModelsDir = Join-Path $BITNET_DIR "models"
        $ggufFiles = Get-ChildItem -Path $bitnetModelsDir -Recurse -Filter "*.gguf" -ErrorAction SilentlyContinue
        if ($ggufFiles.Count -gt 0) {
            foreach ($gf in $ggufFiles) {
                $dest = Join-Path $FALCON_DIR $gf.Name
                if (-not (Test-Path $dest)) {
                    Copy-Item $gf.FullName $dest -Force
                    Write-Info "Copied GGUF model to: $dest"
                }
            }
            Write-Ok "Falcon3-7B GGUF model ready in $FALCON_DIR"
        }
    } else {
        # Fallback: manual CMake build with Clang (if setup_env.py missing)
        Write-Info "setup_env.py not found, falling back to manual CMake build..."
        $buildDir = Join-Path $BITNET_DIR "build"
        if (Test-Path $buildDir) { Remove-Item $buildDir -Recurse -Force }
        New-Item -ItemType Directory -Path $buildDir -Force | Out-Null

        Push-Location $buildDir

        $clangPath = (Get-Command clang -ErrorAction SilentlyContinue).Source
        if ($clangPath) {
            Write-Info "Using Clang at: $clangPath"
            # Match setup_env.py cmake flags: -T ClangCL, -DBITNET_X86_TL2=OFF
            Invoke-NativeCommand -FilePath "cmake" -Arguments @("..", "-T", "ClangCL", "-DBITNET_X86_TL2=OFF", "-DCMAKE_C_COMPILER=clang", "-DCMAKE_CXX_COMPILER=clang++") -ShowOutput
        } else {
            Write-Err "Clang not found in PATH. Please install LLVM and try again."
            Pop-Location
            throw "Clang not found in PATH"
        }

        $threadCount = [Math]::Max(1, [Environment]::ProcessorCount)
        Write-Info "Building with $threadCount parallel jobs..."
        Invoke-NativeCommand -FilePath "cmake" -Arguments @("--build", ".", "--config", "Release", "-j", "$threadCount") -ShowOutput

        Pop-Location
    }

    # Verify build output
    foreach ($p in $possiblePaths) {
        if (Test-Path $p) {
            Write-Ok "llama-server built successfully: $p"
            return
        }
    }

    # Search recursively as a last resort
    $found = Get-ChildItem -Path $BITNET_DIR -Recurse -Filter "llama-server.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) {
        Write-Ok "llama-server built successfully: $($found.FullName)"
        return
    }

    Write-Err "llama-server build failed. Check the build output above."
    Write-Err "You may need to build manually. See: https://github.com/microsoft/BitNet"
    throw "llama-server build failed"
}

# -- Model Downloads --------------------------------------------

function Get-FalconModel {

    $modelFiles = Get-ChildItem -Path $FALCON_DIR -Filter "*.gguf" -ErrorAction SilentlyContinue
    if ($modelFiles.Count -gt 0) {
        Write-Ok "Falcon model already downloaded: $($modelFiles[0].Name)"
        return
    }

    Write-Info "Downloading from HuggingFace ($FALCON_REPO)..."
    Write-Info "This is ~1.7 GB and may take a few minutes."

    # Use huggingface-cli if available, otherwise direct download
    if (Test-Command "huggingface-cli") {
        & huggingface-cli download $FALCON_REPO --local-dir $FALCON_DIR --include "*.gguf"
    } else {
        # Install huggingface_hub and use it
        & (Join-Path $VENV_DIR "Scripts\pip.exe") install huggingface_hub 2>$null

        $downloadScript = @"
from huggingface_hub import snapshot_download
snapshot_download(
    repo_id="$FALCON_REPO",
    local_dir=r"$FALCON_DIR",
    allow_patterns=["*.gguf"]
)
"@
        & (Join-Path $VENV_DIR "Scripts\python.exe") -c $downloadScript
    }

    $modelFiles = Get-ChildItem -Path $FALCON_DIR -Filter "*.gguf" -ErrorAction SilentlyContinue
    if ($modelFiles.Count -gt 0) {
        Write-Ok "Falcon model downloaded: $($modelFiles[0].Name) ($([math]::Round($modelFiles[0].Length / 1MB)) MB)"
    } else {
        Write-Err "Falcon model download failed."
        throw "Falcon model download failed"
    }
}

function Get-Florence2Model {

    $configFile = Join-Path $FLORENCE_DIR "config.json"
    if (Test-Path $configFile) {
        Write-Ok "Florence-2 model already downloaded"
        return
    }

    Write-Info "Downloading Florence-2-base from HuggingFace..."
    Write-Info "This may take a few minutes."

    $downloadScript = @"
from transformers import AutoProcessor, AutoModelForCausalLM
import os

model_path = r"$FLORENCE_DIR"
print("Downloading Florence-2-base model...")
model = AutoModelForCausalLM.from_pretrained("microsoft/Florence-2-base", trust_remote_code=True)
processor = AutoProcessor.from_pretrained("microsoft/Florence-2-base", trust_remote_code=True)

print("Saving model to", model_path)
model.save_pretrained(model_path)
processor.save_pretrained(model_path)
print("Done!")
"@
    & (Join-Path $VENV_DIR "Scripts\python.exe") -c $downloadScript

    if (Test-Path $configFile) {
        Write-Ok "Florence-2 model downloaded"
    } else {
        Write-Err "Florence-2 model download failed."
        throw "Florence-2 model download failed"
    }
}

# -- Python Virtual Environment ---------------------------------

function Initialize-PythonVenv {

    $venvPython = Join-Path $VENV_DIR "Scripts\python.exe"
    if (Test-Path $venvPython) {
        Write-Ok "Virtual environment exists"
    } else {
        Write-Info "Creating virtual environment..."
        & python -m venv $VENV_DIR
        Write-Ok "Virtual environment created"
    }

    Write-Info "Installing Python dependencies..."
    $venvPython = Join-Path $VENV_DIR "Scripts\python.exe"
    $pipExe = Join-Path $VENV_DIR "Scripts\pip.exe"

    # Use python -m pip for self-upgrade (pip.exe cannot overwrite itself on Windows)
    Invoke-NativeCommand -FilePath $venvPython -Arguments @("-m", "pip", "install", "--upgrade", "pip") -LastLineOnly

    # Install from requirements.txt if available, otherwise install directly
    $reqFile = Join-Path $SERVER_DIR "requirements.txt"
    if (Test-Path $reqFile) {
        Invoke-NativeCommand -FilePath $pipExe -Arguments @("install", "-r", $reqFile) -LastLineOnly
    } else {
        Invoke-NativeCommand -FilePath $pipExe -Arguments @(
            "install",
            "fastapi>=0.104.0",
            "uvicorn[standard]>=0.24.0",
            "transformers>=4.36.0",
            "torch>=2.1.0",
            "torchvision>=0.16.0",
            "pillow>=10.0.0",
            "python-multipart>=0.0.6",
            "aiosqlite>=0.19.0",
            "httpx>=0.25.0",
            "sse-starlette>=1.8.0",
            "PyMuPDF>=1.23.0",
            "python-docx>=1.0.0",
            "pyyaml>=6.0",
            "psutil>=5.9.0"
        ) -LastLineOnly
    }

    Write-Ok "Python dependencies installed"
}

# -- NGINX ------------------------------------------------------

function Install-Nginx {

    $nginxExe = Join-Path $NGINX_DIR "nginx.exe"
    if (Test-Path $nginxExe) {
        Write-Ok "NGINX already installed"
    } else {
        $zipPath = Join-Path $env:TEMP "nginx.zip"

        Invoke-DownloadWithProgress -Uri $NGINX_URL -OutFile $zipPath -Label "NGINX $NGINX_VERSION"

        Write-Info "Extracting..."
        $extractDir = Join-Path $env:TEMP "nginx-extract"
        if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
        Expand-Archive -Path $zipPath -DestinationPath $extractDir

        # Move contents from the nested directory
        $innerDir = Get-ChildItem -Path $extractDir -Directory | Select-Object -First 1
        Copy-Item -Path (Join-Path $innerDir.FullName "*") -Destination $NGINX_DIR -Recurse -Force

        # Cleanup
        Remove-Item $zipPath -ErrorAction SilentlyContinue
        Remove-Item $extractDir -Recurse -ErrorAction SilentlyContinue

        Write-Ok "NGINX installed"
    }

    # Generate config
    Write-Info "Generating NGINX configuration..."
    $confDir = Join-Path $NGINX_DIR "conf"
    if (-not (Test-Path $confDir)) {
        New-Item -ItemType Directory -Path $confDir -Force | Out-Null
    }

    # Detect mode and set CORS origin
    $corsOrigin = "https://edge.fractionate.ai"

    $nginxConf = @"
# Fractionate Edge -- NGINX Configuration (auto-generated)
worker_processes 1;

events {
    worker_connections 64;
}

http {
    server {
        listen 127.0.0.1:8080;

        set `$cors_origin "$corsOrigin";

        location /api/health {
            proxy_pass http://127.0.0.1:8081;
            add_header Access-Control-Allow-Origin `$cors_origin always;
            add_header Access-Control-Allow-Methods "GET, OPTIONS" always;
            add_header Access-Control-Allow-Headers "Content-Type" always;

            if (`$request_method = OPTIONS) {
                return 204;
            }
        }

        location /api/ {
            proxy_pass http://127.0.0.1:8081;
            proxy_http_version 1.1;
            proxy_set_header Connection "";
            proxy_set_header Host `$host;
            proxy_buffering off;
            proxy_cache off;
            proxy_read_timeout 300s;
            client_max_body_size 100m;

            add_header Access-Control-Allow-Origin `$cors_origin always;
            add_header Access-Control-Allow-Methods "GET, POST, PUT, DELETE, OPTIONS" always;
            add_header Access-Control-Allow-Headers "Content-Type, Authorization" always;
            add_header Access-Control-Expose-Headers "Content-Type" always;

            if (`$request_method = OPTIONS) {
                return 204;
            }
        }

        location / {
            return 404;
        }
    }
}
"@

    $nginxConf | Out-File -FilePath (Join-Path $confDir "nginx.conf") -Encoding UTF8 -Force
    Write-Ok "NGINX configuration generated"
}

# -- Deploy Server Files ----------------------------------------

function Deploy-ServerFiles {

    # Copy server files from the repo (if running from the repo directory)
    # Otherwise, they should be downloaded alongside this script
    $sourceFiles = @("main.py", "database.py", "models.py", "requirements.txt")
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
    $serverSourceDir = Join-Path $scriptDir "server"

    foreach ($file in $sourceFiles) {
        $sourcePath = Join-Path $serverSourceDir $file
        $destPath = Join-Path $SERVER_DIR $file

        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $destPath -Force
            Write-Info "Deployed: $file"
        } elseif (-not (Test-Path $destPath)) {
            Write-Warn "Server file not found: $file (expected at $sourcePath)"
        }
    }

    Write-Ok "Server files deployed"
}

# -- Configuration File -----------------------------------------

function Initialize-Config {

    if (Test-Path $CONFIG_PATH) {
        Write-Ok "Configuration file already exists"
        return
    }

    # Detect CPU features
    $cpuCores = [Environment]::ProcessorCount
    $optimalThreads = [Math]::Max(1, [Math]::Floor($cpuCores / 2))

    Write-Info "Detected $cpuCores CPU cores, using $optimalThreads threads for inference"

    # Find the GGUF model file
    $ggufFile = Get-ChildItem -Path $FALCON_DIR -Filter "*.gguf" -ErrorAction SilentlyContinue | Select-Object -First 1
    $ggufPath = if ($ggufFile) { $ggufFile.FullName } else { Join-Path $FALCON_DIR "model.gguf" }

    $configYaml = @"
# Fractionate Edge Configuration (auto-generated)
mode: production
cors_origin: "https://edge.fractionate.ai"

models:
  falcon3_7b:
    path: "$($ggufPath -replace '\\', '/')"
    auto_start: true
    threads: $optimalThreads
    context_size: 4096
  florence2:
    path: "$($FLORENCE_DIR -replace '\\', '/')"
    auto_load: false

database:
  path: "$($DB_PATH -replace '\\', '/')"

logging:
  level: "info"
  path: "$($LOGS_DIR -replace '\\', '/')"
"@

    $configYaml | Out-File -FilePath $CONFIG_PATH -Encoding UTF8 -Force
    Write-Ok "Configuration file generated"
}

# -- Task Scheduler ---------------------------------------------

function Register-AutoStart {

    $venvPython = Join-Path $VENV_DIR "Scripts\python.exe"
    $nginxExe = Join-Path $NGINX_DIR "nginx.exe"
    $nginxConf = Join-Path $NGINX_DIR "conf\nginx.conf"

    # Register NGINX task
    $nginxTaskName = "FractionateEdge-NGINX"
    $existingNginx = Get-ScheduledTask -TaskName $nginxTaskName -ErrorAction SilentlyContinue
    if (-not $existingNginx) {
        try {
            $nginxAction = New-ScheduledTaskAction `
                -Execute $nginxExe `
                -Argument "-c `"$nginxConf`"" `
                -WorkingDirectory $NGINX_DIR
            $nginxTrigger = New-ScheduledTaskTrigger -AtLogon
            $nginxSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
            Register-ScheduledTask -TaskName $nginxTaskName -Action $nginxAction `
                -Trigger $nginxTrigger -Settings $nginxSettings `
                -Description "Fractionate Edge NGINX reverse proxy" `
                -RunLevel Limited -Force
            Write-Ok "NGINX auto-start registered"
        } catch {
            Write-Warn "Could not register NGINX auto-start (may need admin): $_"
        }
    } else {
        Write-Ok "NGINX auto-start already registered"
    }

    # Register FastAPI task
    $fastapiTaskName = "FractionateEdge-FastAPI"
    $existingFastapi = Get-ScheduledTask -TaskName $fastapiTaskName -ErrorAction SilentlyContinue
    if (-not $existingFastapi) {
        try {
            $fastapiAction = New-ScheduledTaskAction `
                -Execute $venvPython `
                -Argument "-m uvicorn main:app --host 127.0.0.1 --port 8081" `
                -WorkingDirectory $SERVER_DIR
            $fastapiTrigger = New-ScheduledTaskTrigger -AtLogon
            $fastapiSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
            Register-ScheduledTask -TaskName $fastapiTaskName -Action $fastapiAction `
                -Trigger $fastapiTrigger -Settings $fastapiSettings `
                -Description "Fractionate Edge FastAPI backend" `
                -RunLevel Limited -Force
            Write-Ok "FastAPI auto-start registered"
        } catch {
            Write-Warn "Could not register FastAPI auto-start (may need admin): $_"
        }
    } else {
        Write-Ok "FastAPI auto-start already registered"
    }
}

# -- Start Services ---------------------------------------------

function Start-Services {

    $nginxExe = Join-Path $NGINX_DIR "nginx.exe"
    $venvPython = Join-Path $VENV_DIR "Scripts\python.exe"

    # Start NGINX
    Write-Info "Starting NGINX..."
    $nginxProc = Get-Process -Name "nginx" -ErrorAction SilentlyContinue
    if ($nginxProc) {
        Write-Info "NGINX is already running"
    } else {
        Push-Location $NGINX_DIR
        Start-Process -FilePath $nginxExe -WindowStyle Hidden
        Pop-Location
        Start-Sleep -Seconds 1
        Write-Ok "NGINX started on 127.0.0.1:8080"
    }

    # Start FastAPI
    Write-Info "Starting FastAPI backend..."
    $env:FRACTIONATE_HOME = $FRACTIONATE_HOME
    Push-Location $SERVER_DIR
    Start-Process -FilePath $venvPython -ArgumentList "-m", "uvicorn", "main:app", "--host", "127.0.0.1", "--port", "8081" `
        -WindowStyle Hidden
    Pop-Location
    Start-Sleep -Seconds 2
    Write-Ok "FastAPI started on 127.0.0.1:8081"
}

# -- Health Check -----------------------------------------------

function Test-HealthCheck {

    $maxRetries = 10
    $retryDelay = 2

    for ($i = 1; $i -le $maxRetries; $i++) {
        try {
            $response = Invoke-RestMethod -Uri "http://127.0.0.1:8080/api/health" -TimeoutSec 5
            if ($response.status -eq "ok") {
                Write-Ok "Health check passed!"
                Write-Info "Database: connected=$($response.database.connected)"
                Write-Info "Falcon3-7B: installed=$($response.models.falcon3_7b.installed), running=$($response.models.falcon3_7b.running)"
                Write-Info "Florence-2: installed=$($response.models.florence2.installed)"
                return $true
            }
        } catch {
            if ($i -lt $maxRetries) {
                Write-Info "Waiting for services to start... (attempt $i/$maxRetries)"
                Start-Sleep -Seconds $retryDelay
            }
        }
    }

    Write-Warn "Health check did not pass after $maxRetries attempts."
    Write-Warn "Services may still be starting. Try opening the web app in a minute."
    return $false
}

# -- Main -------------------------------------------------------

function Main {
    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "   Fractionate Edge -- Windows Setup" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "This script will install and configure the local AI backend."
    Write-Host "It will create files in: $FRACTIONATE_HOME"
    Write-Host ""
    Write-Host "Steps: Admin > Prerequisites > Directories > Python venv >" -ForegroundColor DarkGray
    Write-Host "       Florence-2 > BitNet+Falcon > NGINX > Model wait >" -ForegroundColor DarkGray
    Write-Host "       Config > Launch services" -ForegroundColor DarkGray
    Write-Host ""

    $startTime = Get-Date
    $global:_StepNumber = 0

    # Step 1
    Write-Step "Checking admin privileges"
    Assert-Admin

    # Step 2
    Write-Step "Installing prerequisites"
    Assert-Prerequisites

    # Step 3
    Write-Step "Creating directory structure"
    Initialize-Directories

    # Step 4
    Write-Step "Setting up Python virtual environment"
    Initialize-PythonVenv

    # Step 5 -- kick off Florence-2 download in background while we build BitNet
    Write-Step "Starting Florence-2 download (background)"
    $florenceJob = Start-Job -ScriptBlock {
        param($venvDir, $florenceDir)
        $pythonExe = Join-Path $venvDir "Scripts\python.exe"
        $pipExe = Join-Path $venvDir "Scripts\pip.exe"

        # Check if already downloaded
        $configFile = Join-Path $florenceDir "config.json"
        if (Test-Path $configFile) { return "already_exists" }

        & $pipExe install transformers torch torchvision 2>$null
        $script = @"
from transformers import AutoProcessor, AutoModelForCausalLM
model = AutoModelForCausalLM.from_pretrained('microsoft/Florence-2-base', trust_remote_code=True)
processor = AutoProcessor.from_pretrained('microsoft/Florence-2-base', trust_remote_code=True)
model.save_pretrained(r'$florenceDir')
processor.save_pretrained(r'$florenceDir')
"@
        & $pythonExe -c $script
        return "downloaded"
    } -ArgumentList $VENV_DIR, $FLORENCE_DIR
    Write-Ok "Florence-2 download running in background"

    # Step 6 -- BitNet: downloads Falcon model, generates kernels, builds llama-server
    Write-Step "Building BitNet + downloading Falcon3-7B model"
    Build-BitNet

    # Fallback: if setup_env.py didn't produce a GGUF in FALCON_DIR, download pre-converted GGUF
    $falconFiles = Get-ChildItem -Path $FALCON_DIR -Filter "*.gguf" -ErrorAction SilentlyContinue
    if ($falconFiles.Count -eq 0) {
        Write-Info "GGUF model not found in $FALCON_DIR, downloading pre-converted GGUF..."
        $falconJob = Start-Job -ScriptBlock {
            param($venvDir, $falconDir, $falconRepo)
            $pipExe = Join-Path $venvDir "Scripts\pip.exe"
            $pythonExe = Join-Path $venvDir "Scripts\python.exe"
            & $pipExe install huggingface_hub 2>$null
            $script = "from huggingface_hub import snapshot_download; snapshot_download(repo_id='$falconRepo', local_dir=r'$falconDir', allow_patterns=['*.gguf'])"
            & $pythonExe -c $script
            return "downloaded"
        } -ArgumentList $VENV_DIR, $FALCON_DIR, $FALCON_REPO
    } else {
        $falconJob = $null
        Write-Ok "Falcon3-7B GGUF: $($falconFiles[0].Name) ($([math]::Round($falconFiles[0].Length / 1MB)) MB)"
    }

    # Step 7
    Write-Step "Installing NGINX + deploying server"
    Install-Nginx
    Deploy-ServerFiles

    # Step 8 -- wait for any remaining downloads
    Write-Step "Waiting for model downloads"

    if ($falconJob) {
        Wait-JobWithSpinner -Job $falconJob -Activity "Downloading Falcon3-7B GGUF (~1.7 GB)"
        $falconFiles = Get-ChildItem -Path $FALCON_DIR -Filter "*.gguf" -ErrorAction SilentlyContinue
        if ($falconFiles.Count -gt 0) {
            Write-Ok "Falcon3-7B ready: $($falconFiles[0].Name) ($([math]::Round($falconFiles[0].Length / 1MB)) MB)"
        } else {
            Write-Err "Falcon model download failed. Run the script again to retry."
        }
    } else {
        Write-Ok "Falcon3-7B already ready"
    }

    Wait-JobWithSpinner -Job $florenceJob -Activity "Downloading Florence-2-base"
    if (Test-Path (Join-Path $FLORENCE_DIR "config.json")) {
        Write-Ok "Florence-2 ready"
    } else {
        Write-Err "Florence-2 download failed. Run the script again to retry."
    }

    # Step 9
    Write-Step "Configuring system"
    Initialize-Config
    Register-AutoStart

    # Step 10
    Write-Step "Launching services"
    Start-Services
    $healthy = Test-HealthCheck

    $elapsed = (Get-Date) - $startTime

    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host "   Setup Complete!" -ForegroundColor Green
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "   Time elapsed: $([math]::Round($elapsed.TotalMinutes, 1)) minutes" -ForegroundColor Gray
    Write-Host ""
    Write-Host "   Services:" -ForegroundColor Gray
    Write-Host "     NGINX:   http://127.0.0.1:8080" -ForegroundColor Gray
    Write-Host "     FastAPI: http://127.0.0.1:8081" -ForegroundColor Gray
    Write-Host ""
    Write-Host "   Data directory: $FRACTIONATE_HOME" -ForegroundColor Gray
    Write-Host ""

    if ($healthy) {
        Write-Host "   Open https://edge.fractionate.ai to get started!" -ForegroundColor Cyan
    } else {
        Write-Host "   Services are starting up. Try again in a minute:" -ForegroundColor Yellow
        Write-Host "   https://edge.fractionate.ai" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "   For local development, serve the web app:" -ForegroundColor Gray
    Write-Host "   python -m http.server 3000" -ForegroundColor White
    Write-Host "   Then open http://localhost:3000" -ForegroundColor Gray
    Write-Host ""
}

# Run with error handling so the window never closes silently
try {
    Main
} catch {
    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Red
    Write-Host "   Setup Failed!" -ForegroundColor Red
    Write-Host "=========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "   Error: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "   Location: $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   Full log saved to:" -ForegroundColor Gray
    Write-Host "   $LOG_FILE" -ForegroundColor White
    Write-Host ""
} finally {
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    Write-Host ""
    Write-Host "Log file: $LOG_FILE" -ForegroundColor Gray
    Write-Host ""
    Read-Host "Press Enter to close this window"
}
