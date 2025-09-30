# Script by: hong_test, Teiji
# v1.02 (Final version with overwrite protection)

# LADA Launcher with CUDA and TVAI Support
param()

# Function to format elapsed time
function Format-ElapsedTime {
    param([TimeSpan]$TimeSpan)
    
    $hours = [int][Math]::Floor($TimeSpan.TotalHours)
    $minutes = [int]$TimeSpan.Minutes
    $seconds = [int]$TimeSpan.Seconds
    
    return "{0:D2}h, {1:D2}m, {2:D2}s" -f $hours, $minutes, $seconds
}

# Function to validate file path
function Test-ValidFilePath {
    param([string]$FilePath)
    
    # Check if path is null, empty, or whitespace
    if ([string]::IsNullOrWhiteSpace($FilePath)) {
        return $false
    }
    
    # Check for invalid characters that might cause issues
    try {
        # Try to get the full path - this will fail if path contains invalid characters
        $FullPath = [System.IO.Path]::GetFullPath($FilePath)
        
        # Additional check: ensure the path doesn't contain problematic characters
        $InvalidChars = [System.IO.Path]::GetInvalidPathChars()
        foreach ($char in $InvalidChars) {
            if ($FilePath.Contains($char)) {
                return $false
            }
        }
        
        # Test if the file actually exists
        return (Test-Path -LiteralPath $FilePath -PathType Leaf)
    }
    catch {
        # Any exception means the path is invalid
        return $false
    }
}

# Get script directory  
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$PythonExe = Join-Path $ScriptDir "python\python.exe"
$LadaCli = Join-Path $ScriptDir "python\Scripts\lada-cli.exe"

# TVAI Configuration (modify these values as needed)
$TvaiPath = "C:\Program Files\Topaz Labs LLC\Topaz Video AI"
$TvaiModel = "iris-2"
$TvaiScale = "2"
$PreBlur = "0"
$Noise = "0"
$Details = "0"
$Halo = "0"
$Blur = "0"
$Compression = "0"
$Blend = "0"
$Vram = "1"
$Instances = "1"

# Check if files exist
if (-not (Test-Path $PythonExe)) {
    Write-Host "ERROR: Python not found at $PythonExe" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit
}

if (-not (Test-Path $LadaCli)) {
    Write-Host "ERROR: lada-cli not found at $LadaCli" -ForegroundColor Red
    Read-Host "Press Enter to exit" 
    exit
}

# Create output directory
if (-not (Test-Path "output")) {
    New-Item -ItemType Directory -Path "output" | Out-Null
}

# Main processing loop
$DetectChoice = $null
$TvaiChoice = $null
$Quality = $null
$DeviceChoice = $null
$TvaiDevice = $null
do {
    Clear-Host
    Write-Host "====================================================================================="
    Write-Host "LADA Launcher v1.02" -ForegroundColor Green
    Write-Host "Script by: hong_test, Teiji" -ForegroundColor Green
    Write-Host "====================================================================================="

    # Check for CUDA support
    Write-Host ""
    Write-Host "Checking CUDA availability..." -ForegroundColor Cyan

    # Check if NVIDIA GPU is present
    $NvidiaGPU = $false
    try {
        $GPUInfo = Get-CimInstance -Class Win32_VideoController | Where-Object { $_.Name -like "*NVIDIA*" }
        if ($GPUInfo) {
            $NvidiaGPU = $true
            Write-Host "Found NVIDIA GPU: $($GPUInfo.Name)" -ForegroundColor Green
        } else {
            Write-Host "No NVIDIA GPU detected." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Could not detect GPU information." -ForegroundColor Yellow
    }
	Write-Host ""

    # Check PyTorch CUDA support
    $PytorchCuda = $false
    $CudaDeviceCount = 0
    try {
        Write-Host "Checking PyTorch CUDA support..." -ForegroundColor Cyan
        
        # Run Python command to check PyTorch CUDA
        $PytorchOutput = & $PythonExe -c "
import torch
print('CUDA_AVAILABLE:', torch.cuda.is_available())
if torch.cuda.is_available():
    print('CUDA_COUNT:', torch.cuda.device_count())
    print('CUDA_VERSION:', torch.version.cuda)
    print('GPU_NAME:', torch.cuda.get_device_name(0))
else:
    print('CUDA_COUNT: 0')
    print('CUDA_VERSION: None')
" 2>$null

        if ($LASTEXITCODE -eq 0) {
            $PytorchOutput | ForEach-Object {
                if ($_ -match "CUDA_AVAILABLE: True") {
                    $PytorchCuda = $true
                    Write-Host "PyTorch CUDA support: Available" -ForegroundColor Green
                }
                elseif ($_ -match "CUDA_COUNT: (\d+)") {
                    $CudaDeviceCount = [int]$matches[1]
                    if ($CudaDeviceCount -gt 0) {
                        Write-Host "CUDA devices detected: $CudaDeviceCount" -ForegroundColor Green
                    }
                }
                elseif ($_ -match "CUDA_VERSION: (.+)") {
                    $PytorchCudaVersion = $matches[1]
                    if ($PytorchCudaVersion -ne "None") {
                        Write-Host "PyTorch built with CUDA: $PytorchCudaVersion" -ForegroundColor Green
                    }
                }
                elseif ($_ -match "GPU_NAME: (.+)") {
                    $GpuName = $matches[1]
                    Write-Host "GPU Device: $GpuName" -ForegroundColor Green
                }
            }
            
            if (-not $PytorchCuda) {
                Write-Host "PyTorch CUDA support: Not available" -ForegroundColor Yellow
            }
        } else {
            Write-Host "Could not check PyTorch CUDA support" -ForegroundColor Red
        }
    } catch {
        Write-Host "Error checking PyTorch: $_" -ForegroundColor Red
    }

	Write-Host ""
	
    # Determine device choice automatically
    if ($PytorchCuda -and $CudaDeviceCount -gt 0) {
        $DeviceChoice = "cuda"
        $TvaiDevice = "-2"  # Auto device selection for TVAI
        Write-Host "CUDA is fully supported! Using GPU for processing." -ForegroundColor Green
    } else {
        $DeviceChoice = "cpu"
        $TvaiDevice = "-1"  # CPU for TVAI
        Write-Host "CUDA not available. Using CPU for processing." -ForegroundColor Yellow
        if (-not $PytorchCuda) {
            Write-Host "To enable CUDA (GPU acceleration), run [1] SETUP_PYTORCH.bat." -ForegroundColor Cyan
        }
    }
	Write-Host ""
	
    # Check TVAI availability
    $TvaiAvailable = $false
    $TvaiFFmpegPath = Join-Path $TvaiPath "ffmpeg.exe"
    if (Test-Path $TvaiFFmpegPath) {
        $TvaiAvailable = $true
        Write-Host "Topaz Video AI found at: $TvaiPath" -ForegroundColor Green
    } else {
        Write-Host "Topaz Video AI not found at: $TvaiPath" -ForegroundColor Yellow
    }
	
    Write-Host ""
    Write-Host "====================================================================================="
	Write-Host ""
	
    # Get video file with proper error handling loop
    $VideoFile = $null
    do {
        $InputPath = Read-Host "Enter input video file path (eg, C:\Lada\test.mp4)"
        
        # Validate the input path
        if (-not (Test-ValidFilePath $InputPath)) {
            if ([string]::IsNullOrWhiteSpace($InputPath)) {
                Write-Host "ERROR: Please enter a file path!" -ForegroundColor Red
            } else {
                Write-Host "ERROR: Invalid file name or file not found!" -ForegroundColor Red
                Write-Host "Please check that:" -ForegroundColor Yellow
                Write-Host "  - The file path exists and is correct" -ForegroundColor Yellow
                Write-Host "  - The file path doesn't contain invalid characters" -ForegroundColor Yellow
                Write-Host "  - You have permission to access the file" -ForegroundColor Yellow
            }
            Write-Host ""
        } else {
            $VideoFile = $InputPath
        }
    } while (-not $VideoFile)

    # Get model choices (only on first run or if not previously set)
    if (-not $DetectChoice) {
        Write-Host ""
        Write-Host "Detection Models:" -ForegroundColor Cyan
        Write-Host "  [1] v2 (Best mosaic detection accuracy but slower speed)"
        Write-Host "  [2] v3.1-accurate (Good balance of speed and detection accuracy)"
        Write-Host "  [3] v3.1-fast (Fastest speed at the cost of lower detection accuracy)"
        do {
            $DetectChoice = Read-Host "Choose detection model (default: 1)"
            if ([string]::IsNullOrEmpty($DetectChoice)) { $DetectChoice = "1" }
        } while ($DetectChoice -notin @("1", "2", "3"))

        # Ask about TVAI processing
        if ($TvaiAvailable) {
            Write-Host ""
            Write-Host "Topaz Video AI Enhancement (${TvaiModel}:${TvaiScale}X):" -ForegroundColor Cyan
            Write-Host "  [1] Yes - Apply TVAI enhancement after restoration"
            Write-Host "  [2] No  - Skip TVAI processing"
            do {
                $TvaiChoice = Read-Host "Apply TVAI enhancement? (default: 2)"
                if ([string]::IsNullOrEmpty($TvaiChoice)) { $TvaiChoice = "2" }
            } while ($TvaiChoice -notin @("1", "2"))
        } else {
            $TvaiChoice = "2"
            Write-Host ""
            Write-Host "TVAI not available - skipping enhancement step" -ForegroundColor Yellow
        }

        # Ask for quality parameter
        Write-Host ""
        Write-Host "Quality Parameter (5-30, lower = better quality, larger file):" -ForegroundColor Cyan
        do {
            $Quality = Read-Host "Enter quality (default: 15)"
            if ([string]::IsNullOrEmpty($Quality)) { $Quality = "15" }
            $QualityInt = 0
            if ([int]::TryParse($Quality, [ref]$QualityInt)) {
                if ($QualityInt -ge 5 -and $QualityInt -le 30) {
                    break
                }
            }
            Write-Host "Please enter a number between 5 and 30." -ForegroundColor Yellow
        } while ($true)
    }

    # Set model paths
    if ($DetectChoice -eq "1") {
        $DetectModel = "model_weights/lada_mosaic_detection_model_v2.pt"
    } elseif ($DetectChoice -eq "2") {
        $DetectModel = "model_weights/lada_mosaic_detection_model_v3.1_accurate.pt"
    } else {
        $DetectModel = "model_weights/lada_mosaic_detection_model_v3.1_fast.pt"
    }

    # Set restoration model to BVPP 1.2 (only option)
    $RestoreModel = "model_weights/lada_mosaic_restoration_model_generic_v1.2.pth"

    # Check if model files exist
    if (-not (Test-Path $DetectModel)) {
        Write-Host "ERROR: Detection model not found at $DetectModel" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        continue
    }

    if (-not (Test-Path $RestoreModel)) {
        Write-Host "ERROR: Restoration model not found at $RestoreModel" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        continue
    }
    
    # Get detection model suffix
    $DetectSuffix = ""
    if ($DetectChoice -eq "1") {
        $DetectSuffix = "D1"
    } elseif ($DetectChoice -eq "2") {
        $DetectSuffix = "D2"
    } else {
        $DetectSuffix = "D3"
    }
    
    # Prepare output paths with a numerical suffix and detection model suffix
    $FileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($VideoFile)
    $FileExt = [System.IO.Path]::GetExtension($VideoFile)
    
    $BaseOutputFileNoExt = "output\${FileNameWithoutExt}_lada_${DetectSuffix}Q${Quality}"
    $Suffix = 0
    $OutputFile = "${BaseOutputFileNoExt}${FileExt}"
    while (Test-Path $OutputFile) {
        $Suffix++
        $OutputFile = "${BaseOutputFileNoExt}_${Suffix}${FileExt}"
    }
    
    # Prepare TVAI output paths with a numerical suffix if file exists
    $TvaiBaseOutputFileNoExt = "output\${FileNameWithoutExt}_lada_${DetectSuffix}Q${Quality}+${TvaiModel}${TvaiScale}"
    $Suffix = 0
    $TvaiOutputFile = "${TvaiBaseOutputFileNoExt}${FileExt}"
    while (Test-Path $TvaiOutputFile) {
        $Suffix++
        $TvaiOutputFile = "${TvaiBaseOutputFileNoExt}_${Suffix}${FileExt}"
    }

    Write-Host ""
    Write-Host "====================================================================================="
    Write-Host ""
	Write-Host "Start processing at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')..." -ForegroundColor Green
    Write-Host "Device: $DeviceChoice" -ForegroundColor $(if($DeviceChoice -eq "cuda"){"Green"}else{"Yellow"})
    if ($TvaiChoice -eq "1") {
        Write-Host "TVAI Device: $TvaiDevice" -ForegroundColor $(if($TvaiDevice -eq "-2"){"Green"}else{"Yellow"})
    }
    Write-Host "Input: $VideoFile"
    Write-Host "Output: $OutputFile"

    Write-Host "Detection model: $(if($DetectChoice -eq "1"){"v2"}elseif($DetectChoice -eq "2"){"v3.1-accurate"}else{"v3.1-fast"})"
    Write-Host "Restoration model: BVPP 1.2"
    Write-Host "Quality: $Quality"
    if ($TvaiChoice -eq "1") {
        Write-Host "TVAI Enhancement: Enabled"
		Write-Host "TVAI Output: $TvaiOutputFile"
    } else {
        Write-Host "TVAI Enhancement: Disabled"
    }

    # Change to script directory and execute LADA
    Set-Location $ScriptDir

    # Build arguments array for LADA
    $Args = @(
        $LadaCli, "--codec", "hevc_nvenc", "--crf", $Quality,
        "--input", $VideoFile,
        "--output", $OutputFile,
        "--mosaic-detection-model-path", $DetectModel,
        "--mosaic-restoration-model-path", $RestoreModel,
        "--device", $DeviceChoice
    )

    # Initialize timing variables
    $TotalStartTime = Get-Date
    $RestoreStartTime = $null
    $RestoreEndTime = $null
    $TvaiStartTime = $null
    $TvaiEndTime = $null
    $RestoreElapsed = $null
    $TvaiElapsed = $null

    # Execute LADA restoration
    try {
		Write-Host ""
        Write-Host "Running LADA restoration..." -ForegroundColor Cyan
        Write-Host "Executing: python lada-cli --device $DeviceChoice ..." -ForegroundColor Gray
        
        $RestoreStartTime = Get-Date
        & $PythonExe @Args
        $RestoreEndTime = Get-Date
        $RestoreElapsed = $RestoreEndTime - $RestoreStartTime
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host ""
            Write-Host "LADA restoration completed successfully!" -ForegroundColor Green
            Write-Host "Restoration time: $(Format-ElapsedTime $RestoreElapsed)" -ForegroundColor Cyan
            
            # Run TVAI if requested and available
            if ($TvaiChoice -eq "1" -and $TvaiAvailable) {
                # Build TVAI filter complex (simplified to single model)
                $FilterComplex = "tvai_up=model=${TvaiModel}:scale=${TvaiScale}:preblur=${PreBlur}:noise=${Noise}:details=${Details}:halo=${Halo}:blur=${Blur}:compression=${Compression}:blend=${Blend}:device=${TvaiDevice}:vram=${Vram}:instances=${Instances}"
                $DisplayString = "${TvaiModel}:${TvaiScale}X"
                
                Write-Host ""
                Write-Host "Running TVAI enhancement (${DisplayString})..." -ForegroundColor Cyan
                
                # Execute TVAI
                $TvaiFFmpegPath = Join-Path $TvaiPath "ffmpeg.exe"
                $TvaiArgs = @(
                    "-hide_banner",
                    "-nostdin",
                    "-y",
                    "-i", $OutputFile,
                    "-sws_flags", "spline+accurate_rnd+full_chroma_int",
                    "-filter_complex", $FilterComplex,
                    "-c:v", "hevc_nvenc", # FFmpeg video codec
                    "-profile:v", "main",
                    "-pix_fmt", "yuv420p",
                    "-b_ref_mode", "disabled", # Controls B-frame reference mode in NVENC (whether B-frames can be used as references for other frames). Can be "disabled", "each", "middle".
                    "-tag:v", "hvc1",
                    "-g", "30", # Sets the GOP (Group of Pictures) size (eg, keyframes: smaller = easier seeking/editing but larger file size, larger = better compression but harder seeking)
                    "-rc", "constqp", # rate control mode: cbr = Constant Bitrate, vbr = Variable Bitrate, constqp = Constant QP (quality fixed, bitrate varies)
                    "-qp", $Quality,  # Use QP for constqp mode
                    "-preset", "p6", # NVENC preset (speed/quality trade-off; p1 fastest, p7 slowest/best quality).
                    "-map", "0:a?",
                    "-map_metadata:s:a:0", "0:s:a:0",
                    "-c:a", "copy",
                    "-bsf:a:0", "aac_adtstoasc",
                    "-map_metadata", "0",
                    "-map_metadata:s:v", "0:s:v",
                    "-fps_mode:v", "passthrough",
                    "-movflags", "frag_keyframe+empty_moov+delay_moov+use_metadata_tags+write_colr",
                    "-bf", "0", # Sets the number of B-frames to use (0 = less compression efficiency but lower latency, higher = better compression but more latency.)
                    $TvaiOutputFile
                )
                
                $TvaiStartTime = Get-Date
                & $TvaiFFmpegPath @TvaiArgs
                $TvaiEndTime = Get-Date
                $TvaiElapsed = $TvaiEndTime - $TvaiStartTime
                
                if ($LASTEXITCODE -eq 0) {
                    Write-Host ""
                    Write-Host "TVAI enhancement completed successfully!" -ForegroundColor Green
                    Write-Host "TVAI enhancement time: $(Format-ElapsedTime $TvaiElapsed)" -ForegroundColor Cyan
                } else {
                    Write-Host ""
                    Write-Host "TVAI enhancement failed!" -ForegroundColor Red
                    
                    # Delete the failed TVAI output file if it exists
                    if (Test-Path $TvaiOutputFile) {
                        try {
                            Remove-Item $TvaiOutputFile -Force
                            Write-Host "Deleted failed TVAI output file: $TvaiOutputFile" -ForegroundColor Yellow
                        } catch {
                            Write-Host "Warning: Could not delete failed TVAI output file: $_" -ForegroundColor Yellow
                        }
                    }
                }
            } else {
                # do nothing
            }
            
        } else {
            Write-Host ""
            Write-Host "LADA restoration failed!" -ForegroundColor Red
            if ($RestoreElapsed) {
                Write-Host "Time before failure: $(Format-ElapsedTime $RestoreElapsed)" -ForegroundColor Yellow
            }
        }
        
    } catch {
        Write-Host ""
        Write-Host "Error during processing: $_" -ForegroundColor Red
        if ($RestoreStartTime) {
            $ErrorTime = Get-Date
            $ErrorElapsed = $ErrorTime - $RestoreStartTime
            Write-Host "Time before error: $(Format-ElapsedTime $ErrorElapsed)" -ForegroundColor Yellow
        }
    }

    # Calculate and display total processing time
    $TotalEndTime = Get-Date
    $TotalElapsed = $TotalEndTime - $TotalStartTime

    Write-Host ""
    Write-Host "====================================================================================="
    Write-Host "PROCESSING SUMMARY" -ForegroundColor Green
    Write-Host "====================================================================================="
    Write-Host "Started at:             $(Get-Date $TotalStartTime -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Green
    
    if ($RestoreElapsed) {
        Write-Host "Restoration time:           $(Format-ElapsedTime $RestoreElapsed)" -ForegroundColor Cyan
    }
    
    if ($TvaiElapsed) {
        Write-Host "TVAI enhancement time:      $(Format-ElapsedTime $TvaiElapsed)" -ForegroundColor Cyan
    }
    
    Write-Host "Total processing time:      $(Format-ElapsedTime $TotalElapsed)" -ForegroundColor Cyan
    Write-Host "Completed at:               $(Get-Date $TotalEndTime -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Green
    Write-Host "====================================================================================="
    Write-Host ""
	
    # Exit after processing is complete
    Read-Host "Press Enter to exit"
    exit
    
} while ($true)