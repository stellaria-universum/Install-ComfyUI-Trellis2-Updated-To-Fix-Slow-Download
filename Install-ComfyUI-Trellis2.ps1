param(
  [string]$InstallDir = "",
  [bool]$DownloadDino = $false,
  [string]$HfToken = ""
)

$ErrorActionPreference = "Stop"
# Disable the progress bar to significantly increase download speeds
$ProgressPreference = 'SilentlyContinue'

# Force TLS 1.2 for older Win/PS combos
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

function Ensure-Dir([string]$Path) {
  if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Force -Path $Path | Out-Null }
}

function Download-File([string]$Url, [string]$OutFile) {
  if (Test-Path $OutFile) { return }
  Write-Host "Downloading: $Url"
  Ensure-Dir (Split-Path $OutFile -Parent)
  try {
    # Using BITS for faster, more reliable transfers
    Start-BitsTransfer -Source $Url -Destination $OutFile -ErrorAction Stop
  } catch {
    # Fallback to Invoke-WebRequest if BITS fails for specific URLs
    Invoke-WebRequest -Uri $Url -OutFile $OutFile
  }
}

function Run-Exe([string]$Exe, [string]$Arguments) {
  if ([string]::IsNullOrWhiteSpace($Arguments)) {
    Write-Host "Running: $Exe"
    $p = Start-Process -FilePath $Exe -Wait -PassThru
  } else {
    Write-Host "Running: $Exe $Arguments"
    $p = Start-Process -FilePath $Exe -ArgumentList $Arguments -Wait -PassThru
  }
  if ($p.ExitCode -ne 0) { throw "Command failed (exit code $($p.ExitCode)): $Exe $Arguments" }
}

function Git-Ok {
  $g = Get-Command git -ErrorAction SilentlyContinue
  return ($null -ne $g)
}

function Prompt-YesNo([string]$Message, [bool]$DefaultNo = $true) {
  $suffix = if ($DefaultNo) { "[y/N]" } else { "[Y/n]" }
  $ans = Read-Host "$Message $suffix"
  if ([string]::IsNullOrWhiteSpace($ans)) { return (-not $DefaultNo) }
  return ($ans.Trim().ToLower() -eq "y" -or $ans.Trim().ToLower() -eq "yes")
}

function Read-SecureToPlain([System.Security.SecureString]$sec) {
  if ($null -eq $sec) { return "" }
  $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
  try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
  finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

Write-Host ""
Write-Host "=== ComfyUI + Trellis2 clean installer (local Python 3.11.9) ==="
Write-Host ""

# -----------------------------
# Ask install directory
# -----------------------------
$defaultDir = Join-Path $env:USERPROFILE "AI\ComfyUI_Trellis2"

if ([string]::IsNullOrWhiteSpace($InstallDir)) {
    $InstallDir = Read-Host "Install folder (press Enter for: $defaultDir)"
    if ([string]::IsNullOrWhiteSpace($InstallDir)) { $InstallDir = $defaultDir }
}

$installDir = [IO.Path]::GetFullPath($InstallDir)
Ensure-Dir $installDir

$downloads = Join-Path $installDir "_downloads"
$tools     = Join-Path $installDir "tools"
Ensure-Dir $downloads
Ensure-Dir $tools

# -----------------------------
# Local Python 3.11.9 (Embeddable - avoids installer conflicts)
# -----------------------------
$pythonDir = Join-Path $tools "python311"
$pythonExe = Join-Path $pythonDir "python.exe"

$pythonZipUrl = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip"
$pythonZip    = Join-Path $downloads "python-3.11.9-embed-amd64.zip"
$getPipUrl    = "https://bootstrap.pypa.io/get-pip.py"
$getPipFile   = Join-Path $downloads "get-pip.py"

if (-not (Test-Path $pythonExe)) {
  Download-File $pythonZipUrl $pythonZip
  
  Write-Host "Extracting Python embeddable package..."
  Ensure-Dir $pythonDir
  Expand-Archive -Path $pythonZip -DestinationPath $pythonDir -Force

  # Install pip
  Download-File $getPipUrl $getPipFile
  Write-Host "Installing pip..."
  Run-Exe $pythonExe $getPipFile
}

# -----------------------------
# Configure Python ._pth file (CRITICAL for embeddable python)
# -----------------------------
$pthFile = Join-Path $pythonDir "python311._pth"
if (Test-Path $pthFile) {
  Write-Host "Configuring python311._pth..."
  $content = @(Get-Content $pthFile)
  
  $content = $content -replace "#import site", "import site"
  
  if ($content -notcontains ".") {
      $content = @(".") + $content
  }

  $content = $content | Where-Object { $_ -notmatch "ComfyUI" }
  
  if ([string]::IsNullOrWhiteSpace($comfyDir)) {
      $comfyDir = Join-Path $installDir "ComfyUI"
  }
  $comfyDir = [IO.Path]::GetFullPath($comfyDir)
  $content += $comfyDir
  
  Set-Content -Path $pthFile -Value $content
}

& $pythonExe -m pip install -U pip setuptools wheel | Out-Null

# -----------------------------
# Git (use existing git if available; else download PortableGit)
# -----------------------------
if (-not (Git-Ok)) {
  $gitDir = Join-Path $tools "git"
  Ensure-Dir $gitDir

  $portableGitUrl  = "https://github.com/git-for-windows/git/releases/download/v2.52.0.windows.1/PortableGit-2.52.0-64-bit.7z.exe"
  $portableGitFile = Join-Path $downloads "PortableGit-2.52.0-64-bit.7z.exe"

  Download-File $portableGitUrl $portableGitFile

  Run-Exe $portableGitFile "-o`"$gitDir`" -y"

  $env:PATH = (Join-Path $gitDir "cmd") + ";" + (Join-Path $gitDir "bin") + ";" + $env:PATH
}

# -----------------------------
# Clone ComfyUI
# -----------------------------
$comfyDir = Join-Path $installDir "ComfyUI"
if (-not (Test-Path $comfyDir)) {
  & git clone "https://github.com/comfyanonymous/ComfyUI.git" $comfyDir
} else {
  Write-Host "ComfyUI folder already exists -> pulling latest..."
  & git -C $comfyDir pull
}

# -----------------------------
# Setup Environment (using Embeddable Python directly)
# -----------------------------
$venvDir = $pythonDir
$venvPy  = $pythonExe

# -----------------------------
# Install PyTorch 2.7.0 + cu128
# -----------------------------
Write-Host ""
Write-Host "Installing PyTorch 2.7.0 + cu128 (GPU build)..."
& $venvPy -m pip install `
  "torch==2.7.0" "torchvision==0.22.0" "torchaudio==2.7.0" `
  --index-url "https://download.pytorch.org/whl/cu128"

# -----------------------------
# Install ComfyUI requirements
# -----------------------------
Write-Host ""
Write-Host "Installing ComfyUI requirements..."
& $venvPy -m pip install -r (Join-Path $comfyDir "requirements.txt")
& $venvPy -m pip install PyYAML

# -----------------------------
# ComfyUI-Manager (optional but recommended)
# -----------------------------
$mgrDir = Join-Path $comfyDir "custom_nodes\ComfyUI-Manager"
if (-not (Test-Path $mgrDir)) {
  Write-Host ""
  Write-Host "Installing ComfyUI-Manager..."
  Ensure-Dir (Split-Path $mgrDir -Parent)
  & git clone "https://github.com/ltdrdata/ComfyUI-Manager.git" $mgrDir
}

# -----------------------------
# Trellis2 custom node
# -----------------------------
$trellisDir = Join-Path $comfyDir "custom_nodes\ComfyUI-Trellis2"
if (-not (Test-Path $trellisDir)) {
  Write-Host ""
  Write-Host "Installing ComfyUI-Trellis2..."
  Ensure-Dir (Split-Path $trellisDir -Parent)
  & git clone "https://github.com/visualbruno/ComfyUI-Trellis2.git" $trellisDir
} else {
  & git -C $trellisDir pull
}

if (Test-Path (Join-Path $trellisDir "requirements.txt")) {
  Write-Host "Installing Trellis2 requirements..."
  & $venvPy -m pip install -r (Join-Path $trellisDir "requirements.txt")
}

Write-Host "Installing Trellis2 Windows wheels (Torch270)..."
$wheelBase = Join-Path $trellisDir "wheels\Windows\Torch270"
$wheelNames = @(
  "cumesh-0.0.1-cp311-cp311-win_amd64.whl",
  "nvdiffrast-0.4.0-cp311-cp311-win_amd64.whl",
  "nvdiffrec_render-0.0.0-cp311-cp311-win_amd64.whl",
  "flex_gemm-0.0.1-cp311-cp311-win_amd64.whl",
  "o_voxel-0.0.1-cp311-cp311-win_amd64.whl"
)

foreach ($w in $wheelNames) {
  $p = Join-Path $wheelBase $w
  if (-not (Test-Path $p)) { throw "Missing wheel: $p (repo layout changed?)" }
  & $venvPy -m pip install $p
}

# -----------------------------
# DINOv3 model (required by Trellis2)
# -----------------------------
Write-Host ""
if (-not $DownloadDino) {
    $DownloadDino = Prompt-YesNo "Download required facebook/dinov3 model now? (Needs HF access + token)"
}

# -----------------------------
# DINOv3 (The Fixed Part)
# -----------------------------
if (-not $DownloadDino) { $DownloadDino = Prompt-YesNo "Download dinov3 model?" }
if ($DownloadDino) {
  & $venvPy -m pip install -U "huggingface_hub[cli]" "hf-transfer"
  if ([string]::IsNullOrWhiteSpace($HfToken)) {
      $tokenSec = Read-Host "HF Token (hidden)" -AsSecureString
      $HfToken = Read-SecureToPlain $tokenSec
  }
  if (-not [string]::IsNullOrWhiteSpace($HfToken)) {
    $env:HF_TOKEN = $HfToken
    $env:HF_HUB_ENABLE_HF_TRANSFER = "1"
    
    # RELIABLE WAY TO FIND HF-CLI IN EMBEDDABLE PYTHON
    $hfCli = Join-Path $pythonDir "Scripts\huggingface-cli.exe"
    if (-not (Test-Path $hfCli)) { $hfCli = Join-Path $pythonDir "huggingface-cli.exe" }
    
    # FINAL FALLBACK: Call it via python module if exe is missing
    if (Test-Path $hfCli) {
        & $hfCli download "facebook/dinov3-vitl16-pretrain-lvd1689m" --local-dir (Join-Path $comfyDir "models\facebook\dinov3-vitl16-pretrain-lvd1689m") --local-dir-use-symlinks False
    } else {
        Write-Host "Exe not found, trying python module execution..."
        & $venvPy -m huggingface_hub.commands.huggingface_cli download "facebook/dinov3-vitl16-pretrain-lvd1689m" --local-dir (Join-Path $comfyDir "models\facebook\dinov3-vitl16-pretrain-lvd1689m") --local-dir-use-symlinks False
    }
  }
}

# -----------------------------
# Trellis2 Checkpoints
# -----------------------------
$trellisCkptDir = Join-Path $comfyDir "models\trellis2\ckpts"
Ensure-Dir $trellisCkptDir

$trellisModels = @(
    @{ Url="https://huggingface.co/microsoft/TRELLIS-image-large/resolve/main/ckpts/ss_dec_conv3d_16l8_fp16.json"; File="ss_dec_conv3d_16l8_fp16.json" },
    @{ Url="https://huggingface.co/microsoft/TRELLIS-image-large/resolve/main/ckpts/ss_dec_conv3d_16l8_fp16.safetensors"; File="ss_dec_conv3d_16l8_fp16.safetensors" },
    @{ Url="https://huggingface.co/microsoft/TRELLIS.2-4B/resolve/main/ckpts/ss_flow_img_dit_1_3B_64_bf16.json"; File="ss_flow_img_dit_1_3B_64_bf16.json" },
    @{ Url="https://huggingface.co/microsoft/TRELLIS.2-4B/resolve/main/ckpts/ss_flow_img_dit_1_3B_64_bf16.safetensors"; File="ss_flow_img_dit_1_3B_64_bf16.safetensors" },
    @{ Url="https://huggingface.co/microsoft/TRELLIS.2-4B/resolve/main/ckpts/shape_dec_next_dc_f16c32_fp16.json"; File="shape_dec_next_dc_f16c32_fp16.json" },
    @{ Url="https://huggingface.co/microsoft/TRELLIS.2-4B/resolve/main/ckpts/shape_dec_next_dc_f16c32_fp16.safetensors"; File="shape_dec_next_dc_f16c32_fp16.safetensors" },
    @{ Url="https://huggingface.co/microsoft/TRELLIS.2-4B/resolve/main/ckpts/slat_flow_img2shape_dit_1_3B_512_bf16.json"; File="slat_flow_img2shape_dit_1_3B_512_bf16.json" },
    @{ Url="https://huggingface.co/microsoft/TRELLIS.2-4B/resolve/main/ckpts/slat_flow_img2shape_dit_1_3B_512_bf16.safetensors"; File="slat_flow_img2shape_dit_1_3B_512_bf16.safetensors" },
    @{ Url="https://huggingface.co/microsoft/TRELLIS.2-4B/resolve/main/ckpts/slat_flow_img2shape_dit_1_3B_1024_bf16.json"; File="slat_flow_img2shape_dit_1_3B_1024_bf16.json" },
    @{ Url="https://huggingface.co/microsoft/TRELLIS.2-4B/resolve/main/ckpts/slat_flow_img2shape_dit_1_3B_1024_bf16.safetensors"; File="slat_flow_img2shape_dit_1_3B_1024_bf16.safetensors" },
    @{ Url="https://huggingface.co/microsoft/TRELLIS.2-4B/resolve/main/ckpts/tex_dec_next_dc_f16c32_fp16.json"; File="tex_dec_next_dc_f16c32_fp16.json" },
    @{ Url="https://huggingface.co/microsoft/TRELLIS.2-4B/resolve/main/ckpts/tex_dec_next_dc_f16c32_fp16.safetensors"; File="tex_dec_next_dc_f16c32_fp16.safetensors" },
    @{ Url="https://huggingface.co/microsoft/TRELLIS.2-4B/resolve/main/ckpts/slat_flow_imgshape2tex_dit_1_3B_512_bf16.json"; File="slat_flow_imgshape2tex_dit_1_3B_512_bf16.json" },
    @{ Url="https://huggingface.co/microsoft/TRELLIS.2-4B/resolve/main/ckpts/slat_flow_imgshape2tex_dit_1_3B_512_bf16.safetensors"; File="slat_flow_imgshape2tex_dit_1_3B_512_bf16.safetensors" },
    @{ Url="https://huggingface.co/microsoft/TRELLIS.2-4B/resolve/main/ckpts/slat_flow_imgshape2tex_dit_1_3B_1024_bf16.json"; File="slat_flow_imgshape2tex_dit_1_3B_1024_bf16.json" },
    @{ Url="https://huggingface.co/microsoft/TRELLIS.2-4B/resolve/main/ckpts/slat_flow_imgshape2tex_dit_1_3B_1024_bf16.safetensors"; File="slat_flow_imgshape2tex_dit_1_3B_1024_bf16.safetensors" }
)

Write-Host "Downloading Trellis2 checkpoints..."
foreach ($m in $trellisModels) {
    $dest = Join-Path $trellisCkptDir $m.File
    Download-File $m.Url $dest
}

# -----------------------------
# Create launchers
# -----------------------------
$startBat = Join-Path $installDir "Start_ComfyUI.bat"
@"
@echo off
setlocal
cd /d "%~dp0ComfyUI"
set PYTHONPATH=%~dp0ComfyUI;%PYTHONPATH%
"%~dp0tools\python311\python.exe" main.py
pause
"@ | Set-Content -Encoding ASCII $startBat

$updateBat = Join-Path $installDir "Update_ComfyUI.bat"
@"
@echo off
setlocal
cd /d "%~dp0ComfyUI"
git pull
pause
"@ | Set-Content -Encoding ASCII $updateBat

$debugBat = Join-Path $installDir "Debug_ComfyUI.bat"
@"
@echo off
setlocal
cd /d "%~dp0ComfyUI"
set PYTHONPATH=%~dp0ComfyUI;%PYTHONPATH%
"%~dp0tools\python311\python.exe" -c "import sys; print('Sys.path:'); print(sys.path); import comfy; print('Comfy found!')"
pause
"@ | Set-Content -Encoding ASCII $debugBat

Write-Host ""
Write-Host "DONE."
Write-Host "Run: $startBat"
Write-Host "If it fails, try: $debugBat"
Write-Host "If Trellis2 complains about missing DINOv3, rerun installer and say YES to the model download."
Write-Host ""