# Fractionate Edge — Windows Setup Script
# Usage:
#   Option 1: Save as setup-windows.ps1, then: powershell -ExecutionPolicy Bypass -File setup-windows.ps1
#   Option 2: Copy entire script, paste into PowerShell as Administrator, press Enter
#
# This script is idempotent — safe to run multiple times.
# It installs all dependencies, downloads models, configures services,
# and starts the local AI backend.

#Requires -Version 5.1

# Wrap everything in a scriptblock so paste-into-PowerShell executes correctly
& {

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Configuration ──────────────────────────────────────────────

$FRACTIONATE_HOME = Join-Path $env:USERPROFILE ".fractionate"
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
$BITNET_REPO      = "https://github.com/microsoft/BitNet.git"

# Colors
function Write-Step    { param($msg) Write-Host "`n>> $msg" -ForegroundColor Cyan }
function Write-Ok      { param($msg) Write-Host "   [OK] $msg" -ForegroundColor Green }
function Write-Warn    { param($msg) Write-Host "   [WARN] $msg" -ForegroundColor Yellow }
function Write-Err     { param($msg) Write-Host "   [ERROR] $msg" -ForegroundColor Red }
function Write-Info    { param($msg) Write-Host "   $msg" -ForegroundColor Gray }

# ── Elevation Check ────────────────────────────────────────────

function Assert-Admin {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Warn "This script works best with Administrator privileges."
        Write-Warn "Some steps (like installing build tools) may require elevation."
        Write-Host ""
        $response = Read-Host "Continue without admin? (y/N)"
        if ($response -ne "y" -and $response -ne "Y") {
            Write-Host "Re-launching as Administrator..."
            Start-Process powershell.exe -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`""
            exit
        }
    }
}

# ── Prerequisite Checks ───────────────────────────────────────

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

        # Actually try to run it — Store alias triggers an error or opens Store
        $result = & python --version 2>&1
        if ($LASTEXITCODE -ne 0) { return $false }
        if ($result -match 'Python \d+\.\d+') { return $true }
        return $false
    } catch {
        return $false
    }
}

function Assert-Prerequisites {
    Write-Step "Checking prerequisites"

    # Python — must detect and skip the Windows Store alias
    if (Test-RealPython) {
        $pyVer = & python --version 2>&1
        Write-Ok "Python found: $pyVer"
        $versionMatch = [regex]::Match("$pyVer", '(\d+)\.(\d+)')
        if ($versionMatch.Success) {
            $major = [int]$versionMatch.Groups[1].Value
            $minor = [int]$versionMatch.Groups[2].Value
            if ($major -lt 3 -or ($major -eq 3 -and $minor -lt 10)) {
                Write-Err "Python 3.10+ is required. Found $pyVer"
                exit 1
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

    # Check for C++ compiler (MSVC)
    $vsWhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vsWhere) {
        $vsPath = & $vsWhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
        if ($vsPath) {
            Write-Ok "Visual Studio Build Tools found: $vsPath"
        } else {
            Write-Warn "Visual Studio found but C++ tools not installed."
            Write-Info "Installing C++ Build Tools..."
            Install-VsBuildTools
        }
    } else {
        Write-Warn "Visual Studio Build Tools not found."
        Write-Info "Installing Visual Studio Build Tools..."
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
        Write-Info "Installing $PackageId via winget..."
        winget install --id $PackageId --accept-source-agreements --accept-package-agreements -e
        # Refresh PATH
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        if (Test-Command $CommandName) {
            Write-Ok "$CommandName installed successfully"
        } else {
            Write-Err "Failed to install $CommandName. Please install manually and re-run this script."
            exit 1
        }
    } else {
        Write-Err "$CommandName is required but not found, and winget is not available."
        Write-Err "Please install $CommandName manually and re-run this script."
        exit 1
    }
}

function Install-VsBuildTools {
    $installerUrl = "https://aka.ms/vs/17/release/vs_BuildTools.exe"
    $installerPath = Join-Path $env:TEMP "vs_BuildTools.exe"

    Write-Info "Downloading Visual Studio Build Tools installer..."
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -UseBasicParsing

    Write-Info "Installing C++ Build Tools (this may take several minutes)..."
    Start-Process -FilePath $installerPath -ArgumentList `
        "--quiet", "--wait", "--norestart",
        "--add", "Microsoft.VisualStudio.Workload.VCTools",
        "--includeRecommended" `
        -Wait -NoNewWindow

    Remove-Item $installerPath -ErrorAction SilentlyContinue
    Write-Ok "Visual Studio Build Tools installed"
}

# ── Directory Structure ────────────────────────────────────────

function Initialize-Directories {
    Write-Step "Creating directory structure"

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

# ── BitNet Build ───────────────────────────────────────────────

function Build-BitNet {
    Write-Step "Building BitNet (llama-server)"

    $llamaServer = Join-Path $BITNET_DIR "build\bin\Release\llama-server.exe"
    if (Test-Path $llamaServer) {
        Write-Ok "llama-server already built: $llamaServer"
        return
    }

    # Clone
    if (-not (Test-Path (Join-Path $BITNET_DIR ".git"))) {
        Write-Info "Cloning BitNet repository..."
        git clone $BITNET_REPO $BITNET_DIR
    } else {
        Write-Info "BitNet repo exists, pulling latest..."
        Push-Location $BITNET_DIR
        git pull
        Pop-Location
    }

    # Build
    Write-Info "Configuring CMake build..."
    $buildDir = Join-Path $BITNET_DIR "build"
    if (-not (Test-Path $buildDir)) {
        New-Item -ItemType Directory -Path $buildDir -Force | Out-Null
    }

    Push-Location $buildDir

    # Find vcvarsall.bat for MSVC environment
    $vsWhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    $vsPath = & $vsWhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath

    Write-Info "Building with CMake (this may take several minutes)..."
    cmake .. -DCMAKE_BUILD_TYPE=Release
    cmake --build . --config Release --target llama-server -j $([Math]::Max(1, [Environment]::ProcessorCount / 2))

    Pop-Location

    if (Test-Path $llamaServer) {
        Write-Ok "llama-server built successfully"
    } else {
        # Try alternate path
        $altPath = Join-Path $BITNET_DIR "build\bin\llama-server.exe"
        if (Test-Path $altPath) {
            Write-Ok "llama-server built successfully (alternate path)"
        } else {
            Write-Err "llama-server build failed. Check build output above."
            exit 1
        }
    }
}

# ── Model Downloads ────────────────────────────────────────────

function Get-FalconModel {
    Write-Step "Downloading Falcon3-7B 1.58-bit GGUF model"

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
        exit 1
    }
}

function Get-Florence2Model {
    Write-Step "Downloading Florence-2-base model"

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
        exit 1
    }
}

# ── Python Virtual Environment ─────────────────────────────────

function Initialize-PythonVenv {
    Write-Step "Setting up Python virtual environment"

    $venvPython = Join-Path $VENV_DIR "Scripts\python.exe"
    if (Test-Path $venvPython) {
        Write-Ok "Virtual environment exists"
    } else {
        Write-Info "Creating virtual environment..."
        & python -m venv $VENV_DIR
        Write-Ok "Virtual environment created"
    }

    Write-Info "Installing Python dependencies..."
    $pipExe = Join-Path $VENV_DIR "Scripts\pip.exe"
    & $pipExe install --upgrade pip

    # Install from requirements.txt if available, otherwise install directly
    $reqFile = Join-Path $SERVER_DIR "requirements.txt"
    if (Test-Path $reqFile) {
        & $pipExe install -r $reqFile
    } else {
        & $pipExe install `
            "fastapi>=0.104.0" `
            "uvicorn[standard]>=0.24.0" `
            "transformers>=4.36.0" `
            "torch>=2.1.0" `
            "torchvision>=0.16.0" `
            "pillow>=10.0.0" `
            "python-multipart>=0.0.6" `
            "aiosqlite>=0.19.0" `
            "httpx>=0.25.0" `
            "sse-starlette>=1.8.0" `
            "PyMuPDF>=1.23.0" `
            "python-docx>=1.0.0" `
            "pyyaml>=6.0" `
            "psutil>=5.9.0"
    }

    Write-Ok "Python dependencies installed"
}

# ── NGINX ──────────────────────────────────────────────────────

function Install-Nginx {
    Write-Step "Installing NGINX"

    $nginxExe = Join-Path $NGINX_DIR "nginx.exe"
    if (Test-Path $nginxExe) {
        Write-Ok "NGINX already installed"
    } else {
        $zipPath = Join-Path $env:TEMP "nginx.zip"

        Write-Info "Downloading NGINX $NGINX_VERSION..."
        Invoke-WebRequest -Uri $NGINX_URL -OutFile $zipPath -UseBasicParsing

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
# Fractionate Edge — NGINX Configuration (auto-generated)
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

# ── Deploy Server Files ────────────────────────────────────────

function Deploy-ServerFiles {
    Write-Step "Deploying FastAPI server files"

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

# ── Configuration File ─────────────────────────────────────────

function Initialize-Config {
    Write-Step "Generating configuration"

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

# ── Task Scheduler ─────────────────────────────────────────────

function Register-AutoStart {
    Write-Step "Setting up auto-start on login"

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

# ── Start Services ─────────────────────────────────────────────

function Start-Services {
    Write-Step "Starting services"

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

# ── Health Check ───────────────────────────────────────────────

function Test-HealthCheck {
    Write-Step "Running health check"

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

# ── Main ───────────────────────────────────────────────────────

function Main {
    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "   Fractionate Edge — Windows Setup" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "This script will install and configure the local AI backend."
    Write-Host "It will create files in: $FRACTIONATE_HOME"
    Write-Host ""

    $startTime = Get-Date

    Assert-Admin
    Assert-Prerequisites
    Initialize-Directories
    Initialize-PythonVenv
    Build-BitNet
    Get-FalconModel
    Get-Florence2Model
    Install-Nginx
    Deploy-ServerFiles
    Initialize-Config
    Register-AutoStart
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

# Run
Main

} # End of scriptblock wrapper
