<# 
  TimeLapse-v1.0.0.ps1

  Enhanced PowerShell Time-Lapse Script (All Lossless)
#>

###############################################################################
#                       GLOBAL SETTINGS & INITIALIZATION                        #
###############################################################################
$ErrorActionPreference = "Continue"
$global:closingNow = $false
$global:NoFrameWarningCount = 0
# (The CameraReinitAttempts function has been removed per your request)

###############################################################################
#          SET BASE DIRECTORY AND CONFIG FILE PATH (user–writable location)     #
###############################################################################
$baseDir = Join-Path $env:USERPROFILE "Timelapse"
if (-not (Test-Path $baseDir)) {
    New-Item -ItemType Directory -Path $baseDir | Out-Null
}
$configPath = Join-Path $baseDir "config.json"
$global:LogFile = Join-Path (Join-Path $baseDir "Logs") "TimelapseLog.txt"
$scriptPath = $MyInvocation.MyCommand.Path

###############################################################################
#                               USER CONFIGURATION                            #
###############################################################################
$CameraWidth  = 2560
$CameraHeight = 1440
$CameraFps    = 30
$FfmpegPath   = "ffmpeg"  # FFmpeg is used (third-party open-source software)
$OutputFps    = 30
$CrfValue     = 18  
$PresetVal    = "slow"
$DefaultDarkThreshold = 20
$global:CapturingDurationHoursDefault   = "0"
$global:CapturingDurationMinutesDefault = "0"
$StartupTaskNameDefault = "RunTimelapseStartupScript"
$AdminTaskNameDefault   = "RunTimelapseAdminScript"
$StartupTaskName        = $StartupTaskNameDefault
$AdminTaskName          = $AdminTaskNameDefault
$global:SelectedWidth = $CameraWidth
$global:SelectedHeight = $CameraHeight
$global:EnableTimestamp = $true

###############################################################################
#                        HELPER: CLEAR CURRENT CONSOLE LINE                   #
###############################################################################
function Clear-CurrentLine {
    $width = [Console]::WindowWidth
    [Console]::SetCursorPosition(0, [Console]::CursorTop)
    [Console]::Write((" " * $width))
    [Console]::SetCursorPosition(0, [Console]::CursorTop)
}

###############################################################################
#                          TASK SCHEDULER HELPER                              #
###############################################################################
function Set-TaskState {
    param(
        [string]$TaskName,
        [bool]$Enable
    )
    $action = if ($Enable) { "enable" } else { "disable" }
    try {
        schtasks /change /tn $TaskName /$action | Out-Null
        Write-Log "Task '$TaskName' has been ${action}d." "Info"
    }
    catch {
        Write-Log "Failed to $action task '$TaskName': $_" "Error"
    }
}

###############################################################################
#                             STATUS TEXT FUNCTION                            #
###############################################################################
function Set-StatusText($newStatus) {
    if ($global:closingNow) {
        Write-Log "Skipping status update (closingNow=true): $newStatus" -NoConsole
        return
    }
    if ($global:statusLabel -and $global:statusLabel.IsHandleCreated -and -not $global:statusLabel.Disposing -and -not $global:statusLabel.IsDisposed) {
        $action = [System.Windows.Forms.MethodInvoker]{
            param($status)
            $global:statusLabel.Text = $status
            $global:statusLabel.Refresh()
        }
        try {
            $global:statusLabel.Invoke($action, @($newStatus))
        }
        catch {
            Write-Log "Error updating status text: $_" "Error"
        }
    }
}

###############################################################################
#                                WRITE-LOG                                    #
###############################################################################
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "Info",
        [switch]$NoConsole
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] - $Message"
    if (-not $NoConsole) {
        switch ($Level.ToLower()) {
            "info"    { Write-Host "$logMessage" -ForegroundColor White }
            "warning" { Write-Host "$logMessage" -ForegroundColor Yellow }
            "error"   { Write-Host "$logMessage" -ForegroundColor Red }
            default   { Write-Host "$logMessage" -ForegroundColor White }
        }
    }
    $logFolder = Join-Path $baseDir "Logs"
    if (-not (Test-Path $logFolder)) {
        try {
            New-Item -ItemType Directory -Path $logFolder | Out-Null
            Write-Host "Created log folder at ${logFolder}" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to create log folder at ${logFolder}: $($_)" -ForegroundColor Red
        }
    }
    $logFile = Join-Path $logFolder "TimelapseLog.txt"
    try {
        if (-not $global:logMutex) {
            try {
                $global:logMutex = New-Object System.Threading.Mutex($false, "TimeLapseLogMutex")
            }
            catch {
                Write-Host "Failed to create mutex: $_" -ForegroundColor Red
                $global:logMutex = $null
            }
        }
        if ($global:logMutex) {
            $global:logMutex.WaitOne() | Out-Null
        }
        $logMessage | Out-File -FilePath $logFile -Append -Encoding utf8
    }
    catch {
        Write-Host "Failed to write to log file: $($_)" -ForegroundColor Red
    }
    finally {
        if ($global:logMutex) {
            try {
                [void]$global:logMutex.ReleaseMutex()
            }
            catch { }
        }
    }
}

###############################################################################
#                            SAFE-INVOKE HELPER                               #
###############################################################################
function SafeInvoke($control, [ScriptBlock]$action) {
    if ($control -and $control.IsHandleCreated -and -not $control.Disposing -and -not $control.IsDisposed) {
        $delegate = [System.Windows.Forms.MethodInvoker] { & $action }
        [void]$control.Invoke($delegate)
    }
    else {
        $action.Invoke()
    }
}

###############################################################################
#                         CONFIGURATION MANAGEMENT                            #
###############################################################################
function Load-Config {
    if (Test-Path $configPath) {
        try {
            $configContent = Get-Content -Path $configPath -Raw | ConvertFrom-Json
            $global:ImageFolder = if ($configContent.ImageFolder -and $configContent.ImageFolder.Trim() -ne "") { $configContent.ImageFolder } else { "C:\Users\$env:USERNAME\Timelapse\Images" }
            $global:VideoFolder = if ($configContent.VideoFolder -and $configContent.VideoFolder.Trim() -ne "") { $configContent.VideoFolder } else { "C:\Users\$env:USERNAME\Timelapse\Videos" }
            $global:CaptureMinutes = if ($configContent.CaptureMinutes -ge 0) { [int]$configContent.CaptureMinutes } else { 0 }
            $global:CaptureSeconds = if ($configContent.CaptureSeconds -ge 0) { [int]$configContent.CaptureSeconds } else { 1 }
            $global:SelectedCameraIndex = if ($configContent.SelectedCameraIndex -ge 0) { [int]$configContent.SelectedCameraIndex } else { -1 }
            $global:AutoRestart = if ($configContent.AutoRestart -eq $true) { $true } else { $false }
            $global:FfmpegPath = if ($configContent.FfmpegPath -and $configContent.FfmpegPath.Trim() -ne "") { $configContent.FfmpegPath } else { "ffmpeg" }
            $global:OutputFps = if ($configContent.OutputFps -and [int]$configContent.OutputFps -gt 0) { [int]$configContent.OutputFps } else { $OutputFps }
            $global:AutoRestartHours = if ($configContent.AutoRestartHours -ge 0) { [int]$configContent.AutoRestartHours } else { 0 }
            $global:AutoRestartMinutes = if ($configContent.AutoRestartMinutes -ge 0) { [int]$configContent.AutoRestartMinutes } else { 0 }
            $global:AutoRestartDays = if ($configContent.AutoRestartDays -ge 0) { [int]$configContent.AutoRestartDays } else { 0 }
            $global:CapturingDurationHours = if ($configContent.CapturingDurationHours -ge 0) { [int]$configContent.CapturingDurationHours } else { 0 }
            $global:CapturingDurationMinutes = if ($configContent.CapturingDurationMinutes -ge 0) { [int]$configContent.CapturingDurationMinutes } else { 0 }
            if ($configContent.DarkThreshold -ne $null -and ([int]$configContent.DarkThreshold) -ge 0 -and ([int]$configContent.DarkThreshold) -le 255) {
                $global:DarkThreshold = [int]$configContent.DarkThreshold
            }
            else {
                $global:DarkThreshold = $DefaultDarkThreshold
            }
            $global:DontRemoveOldImages = if ($configContent.DontRemoveOldImages -eq $true) { $true } else { $false }
            $global:DontRemoveOldVideos = if ($configContent.DontRemoveOldVideos -eq $true) { $true } else { $false }
            $global:AutoMergeVideoLimit = if ($configContent.AutoMergeVideoLimit -ge 0) { [int]$configContent.AutoMergeVideoLimit } else { 0 }
            if ($configContent.CameraWidth -and $configContent.CameraHeight) {
                $global:SelectedWidth = [int]$configContent.CameraWidth
                $global:SelectedHeight = [int]$configContent.CameraHeight
            } else {
                $global:SelectedWidth = $CameraWidth
                $global:SelectedHeight = $CameraHeight
            }
            if ($configContent.EnableTimestamp -ne $null) {
                $global:EnableTimestamp = [bool]$configContent.EnableTimestamp
            }
            else {
                $global:EnableTimestamp = $true
            }
            Write-Log "Configuration loaded from $configPath"
        }
        catch {
            Write-Log "Failed to load configuration: $_" "Error"
            $global:ImageFolder = "C:\Users\$env:USERNAME\Timelapse\Images"
            $global:VideoFolder = "C:\Users\$env:USERNAME\Timelapse\Videos"
            $global:CaptureMinutes = 0
            $global:CaptureSeconds = 1
            $global:SelectedCameraIndex = -1
            $global:AutoRestart = $false
            $global:FfmpegPath = "ffmpeg"
            $global:OutputFps = $OutputFps
            $global:AutoRestartHours = 0
            $global:AutoRestartMinutes = 0
            $global:AutoRestartDays = 0
            $global:CapturingDurationHours = 0
            $global:CapturingDurationMinutes = 0
            $global:DarkThreshold = $DefaultDarkThreshold
            $global:DontRemoveOldImages = $false
            $global:DontRemoveOldVideos = $false
            $global:AutoMergeVideoLimit = 0
            $global:SelectedWidth = $CameraWidth
            $global:SelectedHeight = $CameraHeight
            $global:EnableTimestamp = $true
        }
    }
    else {
        Write-Log "Configuration file not found. Using default settings." "Warning"
        $global:ImageFolder = "C:\Users\$env:USERNAME\Timelapse\Images"
        $global:VideoFolder = "C:\Users\$env:USERNAME\Timelapse\Videos"
        $global:CaptureMinutes = 0
        $global:CaptureSeconds = 1
        $global:SelectedCameraIndex = -1
        $global:AutoRestart = $false
        $global:FfmpegPath = "ffmpeg"
        $global:OutputFps = $OutputFps
        $global:AutoRestartHours = 0
        $global:AutoRestartMinutes = 0
        $global:AutoRestartDays = 0
        $global:CapturingDurationHours = 0
        $global:CapturingDurationMinutes = 0
        $global:DarkThreshold = $DefaultDarkThreshold
        $global:DontRemoveOldImages = $false
        $global:DontRemoveOldVideos = $false
        $global:AutoMergeVideoLimit = 0
        $global:SelectedWidth = $CameraWidth
        $global:SelectedHeight = $CameraHeight
        $global:EnableTimestamp = $true
    }
}

function Save-Config {
    $config = @{
        ImageFolder                 = $global:ImageFolder
        VideoFolder                 = $global:VideoFolder
        CaptureMinutes              = $global:CaptureMinutes
        CaptureSeconds              = $global:CaptureSeconds
        SelectedCameraIndex         = $global:SelectedCameraIndex
        AutoRestart                 = $global:AutoRestart
        FfmpegPath                  = $global:FfmpegPath
        OutputFps                   = $global:OutputFps
        CapturingDurationHours      = $global:CapturingDurationHours
        CapturingDurationMinutes    = $global:CapturingDurationMinutes
        DarkThreshold               = $global:DarkThreshold
        DontRemoveOldImages         = $global:DontRemoveOldImages
        DontRemoveOldVideos         = $global:DontRemoveOldVideos
        AutoMergeVideoLimit         = $global:AutoMergeVideoLimit
        AutoRestartHours            = $global:AutoRestartHours
        AutoRestartMinutes          = $global:AutoRestartMinutes
        AutoRestartDays             = $global:AutoRestartDays
        CameraWidth                 = $global:SelectedWidth
        CameraHeight                = $global:SelectedHeight
        EnableTimestamp             = $global:EnableTimestamp
    }
    $config | ConvertTo-Json -Depth 10 | Set-Content -Path $configPath
    Write-Log "Configuration saved to $configPath"
}

###############################################################################
#          GETNEXTIMAGENUMB - returns a single [int]
###############################################################################
function GetNextImageNumber {
    param([string]$ImgFolder)
    Write-Log "Scanning folder for images: $ImgFolder" -NoConsole | Out-Null
    $pattern = "^capture_(\d{6})\.png$"
    $existingImages = Get-ChildItem -Path $ImgFolder -Filter "capture_*.png" -File -ErrorAction SilentlyContinue
    if ($existingImages -and $existingImages.Count -gt 0) {
        Write-Log "Found images: $($existingImages.Name -join ', ')" -NoConsole | Out-Null
    }
    else {
        Write-Log "No images found in folder." -NoConsole | Out-Null
    }
    $numbers = @()
    foreach ($image in $existingImages) {
        if ($image.Name -match $pattern) {
            $numbers += [int]$matches[1]
        }
    }
    if ($numbers.Count -gt 0) {
        $maxNum = $numbers | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
        if (-not $maxNum) { $maxNum = 0 }
        $nextNum = [int]($maxNum + 1)
        Write-Log "Returning next image number: $nextNum" -NoConsole | Out-Null
        return $nextNum
    }
    else {
        Write-Log "No images matching the pattern. Starting at 1." -NoConsole | Out-Null
        return 1
    }
}

###############################################################################
#             ARCHIVE RENAME HELPERS (for Old Images / Old Videos)
###############################################################################
function GetNextArchiveIndex {
    param(
        [string]$FolderPath,
        [string]$FilePrefix,
        [string]$Ext
    )
    $escPrefix = [Regex]::Escape($FilePrefix)
    $escExt    = [Regex]::Escape($Ext)
    $regex     = "^$escPrefix(\d{6})$escExt$"
    $files = Get-ChildItem -Path $FolderPath -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $regex }
    if ($files.Count -gt 0) {
        $numbers = @()
        foreach ($f in $files) {
            if ($f.Name -match $regex) {
                $numbers += [int]$Matches[1]
            }
        }
        if ($numbers.Count -gt 0) {
            $maxNum = $numbers | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
            return ($maxNum + 1)
        }
    }
    return 1
}

function ArchiveFileWithIncrement {
    param(
        [System.IO.FileInfo]$File,
        [string]$OldFolder,
        [string]$FilePrefixForOld
    )
    try {
        if (-not (Test-Path $OldFolder)) {
            New-Item -ItemType Directory -Path $OldFolder | Out-Null
            Write-Log "Created Old folder at $OldFolder" -NoConsole
        }
        $ext = $File.Extension
        $nextIdx = [int](GetNextArchiveIndex -FolderPath $OldFolder -FilePrefix $FilePrefixForOld -Ext $ext)
        $nextIdxStr = $nextIdx.ToString("D6")
        $destFileName = "$FilePrefixForOld$nextIdxStr$ext"
        $destPath = Join-Path $OldFolder $destFileName
        Move-Item -Path $File.FullName -Destination $destPath -Force
        Write-Log "Archived -> $destPath" -NoConsole
    }
    catch {
        Write-Log "Error archiving $($File.FullName) -> $OldFolder : $_" "Error"
    }
}

###############################################################################
#                            ARCHIVING & REMOVALS                              #
###############################################################################
function RemoveOrArchiveImages {
    param([System.IO.FileInfo[]]$Files, [string]$Reason)
    if (-not $Files -or $Files.Count -eq 0) {
        Write-Log "No images to remove for reason: $Reason" -NoConsole
        return
    }
    if ($global:DontRemoveOldImages) {
        $oldImgFolder = Join-Path $global:ImageFolder "Old Images"
        foreach ($file in $Files) {
            ArchiveFileWithIncrement -File $file -OldFolder $oldImgFolder -FilePrefixForOld "capture_old_"
        }
    }
    else {
        foreach ($file in $Files) {
            try {
                Write-Log "Deleting image -> $($file.FullName)" -NoConsole
                Remove-Item -Path $file.FullName -Force
            }
            catch {
                Write-Log "Error deleting image -> $($file.FullName): $_" "Error"
            }
        }
    }
}

function RemoveOrArchiveVideos {
    param([System.IO.FileInfo[]]$Files, [string]$Reason)
    if (-not $Files -or $Files.Count -eq 0) {
        Write-Log "No videos to remove for reason: $Reason" -NoConsole
        return
    }
    if ($global:DontRemoveOldVideos) {
        $oldVidFolder = Join-Path $global:VideoFolder "Old Videos"
        foreach ($file in $Files) {
            ArchiveFileWithIncrement -File $file -OldFolder $oldVidFolder -FilePrefixForOld "video_old_"
        }
    }
    else {
        foreach ($file in $Files) {
            try {
                Write-Log "Deleting video -> $($file.FullName)" -NoConsole
                Remove-Item -Path $file.FullName -Force
            }
            catch {
                Write-Log "Error deleting video -> $($file.FullName): $_" "Error"
            }
        }
    }
}

###############################################################################
#                            IMAGE NORMALIZATION                               #
###############################################################################
function NormalizeImageFilenames {
    param([string]$Folder)
    Write-Log "Normalizing image filenames in $Folder." -NoConsole
    $images = Get-ChildItem -Path $Folder -Filter "capture_*.png" | Sort-Object Name
    $counter = 1
    foreach ($image in $images) {
        $newName = "capture_$($counter.ToString("D6")).png"
        Write-Log "Renamed $($image.Name) -> $newName" -NoConsole
        Rename-Item -Path $image.FullName -NewName $newName -Force | Out-Null
        $counter++
    }
}

###############################################################################
#                        NEW: IMAGE OVERLAY - TIMESTAMP                      #
###############################################################################
function Add-TimestampToImage {
    param(
        [Parameter(Mandatory=$true)]
        [System.Object]$bitmap
    )
    if ($bitmap -is [System.Array]) { $bitmap = $bitmap[0] }
    if (-not ($bitmap -is [System.Drawing.Bitmap])) { return $bitmap }
    try {
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
        $timestamp = (Get-Date).ToString("hh:mm tt MM/dd/yy")
        $font = New-Object System.Drawing.Font("Arial", 48, [System.Drawing.FontStyle]::Bold)
        $shadowBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(128, 0, 0, 0))
        $redBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Red)
        $shadowPoint = New-Object System.Drawing.PointF(21, 21)
        $textPoint = New-Object System.Drawing.PointF(20, 20)
        $graphics.DrawString($timestamp, $font, $shadowBrush, $shadowPoint)
        $graphics.DrawString($timestamp, $font, $redBrush, $textPoint)
        $graphics.Dispose()
    }
    catch {
        Write-Log "Error overlaying timestamp: $_" "Error"
    }
    return $bitmap
}

###############################################################################
#                           VIDEO MERGE & DARK FRAME                           #
###############################################################################
function RemoveDarkImages {
    param([string]$Folder, [int]$Threshold)
    Write-Log "Checking for near-black images in: $Folder with threshold: $Threshold" -NoConsole
    $files = Get-ChildItem -Path $Folder -Filter "capture_*.png" -File -ErrorAction SilentlyContinue
    if (-not $files) {
        Write-Log "No images found to check for darkness." -NoConsole
        return
    }
    $removedCount = 0
    foreach ($file in $files) {
        try {
            $bmp = [System.Drawing.Bitmap]::FromFile($file.FullName)
            $sum = 0; $count = 0
            $w = $bmp.Width; $h = $bmp.Height
            $rowStep = [Math]::Max([Math]::Floor($h / 100), 1)
            $colStep = [Math]::Max([Math]::Floor($w / 100), 1)
            for ($y = 0; $y -lt $h; $y += $rowStep) {
                for ($x = 0; $x -lt $w; $x += $colStep) {
                    $p = $bmp.GetPixel($x, $y)
                    $avg = ($p.R + $p.G + $p.B) / 3
                    $sum += $avg
                    $count++
                }
            }
            $bmp.Dispose()
            if ($count -gt 0) {
                $brightness = $sum / $count
                if ($brightness -lt $Threshold) {
                    Write-Log "Removed near-black image -> $($file.FullName)" -NoConsole
                    Remove-Item $file.FullName -Force
                    $removedCount++
                }
            }
        }
        catch {
            Write-Log "Error checking brightness for $($file.FullName): $_" "Warning"
        }
    }
    if ($removedCount -gt 0) {
        Write-Log "Removed $removedCount near-black images." -NoConsole
    }
}

function MergeImagesIntoVideo {
    param(
        [string]$Folder,
        [string]$VidFolder,
        [int]$FPS,
        [int]$CrfVal,
        [string]$PresetVal,
        [int]$DarkThr,
        [int]$OutW,
        [int]$OutH
    )
    Write-Log "Starting merge into sequentially numbered video." "Info"
    RemoveDarkImages -Folder $Folder -Threshold $DarkThr
    $files = Get-ChildItem -Path $Folder -Filter "capture_*.png" -File -ErrorAction SilentlyContinue | Sort-Object Name
    if (-not $files) {
        Write-Log "No valid images after removing dark frames. No video created." "Info"
        Set-StatusText "Status: No valid images to merge."
        return
    }
    NormalizeImageFilenames -Folder $Folder
    $files = Get-ChildItem -Path $Folder -Filter "capture_*.png" -File -ErrorAction SilentlyContinue | Sort-Object Name
    $existingVideos = Get-ChildItem -Path $VidFolder -Filter "video_*.mp4" | Sort-Object Name
    if ($existingVideos.Count -gt 0) {
        $vidNums = $existingVideos | ForEach-Object { $_.BaseName -replace "^video_" -as [int] }
        $maxVid = $vidNums | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
        if (-not $maxVid) { $maxVid = 0 }
        $nextVideoNumber = [int]($maxVid + 1)
    }
    else {
        $nextVideoNumber = 1
    }
    $outVid = Join-Path $VidFolder ("video_$nextVideoNumber.mp4")
    Write-Log "Merging into -> $outVid" "Info"
    $ffFilter = "scale=${OutW}:${OutH},format=yuv420p"
    $ffmpegArgs = @(
        "-loglevel", "error",
        "-hide_banner",
        "-y",
        "-framerate", "$FPS",
        "-i", "`"$Folder\capture_%06d.png`"",
        "-c:v", "libx264",
        "-preset", "veryslow",
        "-qp", "0",
        "-vf", $ffFilter,
        "-pix_fmt", "yuv420p",
        "`"$outVid`""
    )
    try {
        Write-Log "Running FFmpeg to create video." "Info"
        $process = Start-Process -FilePath $global:FfmpegPath -ArgumentList $ffmpegArgs -NoNewWindow -Wait -PassThru
        if ($process.ExitCode -eq 0) {
            Write-Log "Video created: $outVid" "Info"
            Set-StatusText "Video created: $(Split-Path $outVid -Leaf)"
            RemoveOrArchiveImages -Files $files -Reason "After video creation"
        }
        else {
            Write-Log "FFmpeg merge failed with exit code $($process.ExitCode)." "Error"
            Set-StatusText "Status: FFmpeg merge failed."
        }
    }
    catch {
        Write-Log "FFmpeg merge failed: $_" "Error"
        Set-StatusText "Status: FFmpeg merge failed."
    }
}

###############################################################################
#                     MASTER MERGE WITH ORDERED RENAME
###############################################################################
function MergeAllVideos {
    param([string]$VidFolder, [string]$MasterVideoPath)
    Write-Log "Starting advanced merge with ordered renaming." "Info"
    $tempFiles = @()
    $renameIndex = 1
    if (Test-Path $MasterVideoPath) {
        $tempName = "temp_{0:D6}.mp4" -f $renameIndex
        $oldMasterTemp = Join-Path $VidFolder $tempName
        try {
            Rename-Item -Path $MasterVideoPath -NewName $tempName -Force
            $tempFiles += $oldMasterTemp
            $renameIndex++
        }
        catch {
            Write-Log "Failed to rename existing master: $_" "Error"
            Set-StatusText "Status: Error renaming master."
            return
        }
    }
    $videoFiles = Get-ChildItem -Path $VidFolder -Filter "video_*.mp4" -File -ErrorAction SilentlyContinue
    if ($videoFiles) { $videoFiles = $videoFiles | Sort-Object Name }
    if (-not $videoFiles -or $videoFiles.Count -eq 0) {
        if ($tempFiles.Count -eq 0) {
            Write-Log "No videos to merge." "Info"
            Set-StatusText "Status: Nothing to merge."
            return
        }
        Write-Log "No new segments found; only old master remains. Creating master from that alone." "Info"
        Rename-Item -Path $tempFiles[0] -NewName (Split-Path $MasterVideoPath -Leaf) -Force
        Write-Log "Master unchanged." "Info"
        Set-StatusText "Status: Master unchanged."
        return
    }
    $tempFiles += foreach ($v in $videoFiles) {
        $tempName = "temp_{0:D6}.mp4" -f $renameIndex
        $tempPath = Join-Path $VidFolder $tempName
        try {
            Rename-Item -Path $v.FullName -NewName $tempName -Force
            $renameIndex++
            $tempPath
        }
        catch {
            Write-Log "Error renaming segment '$($v.Name)': $_" "Error"
        }
    }
    if ($tempFiles.Count -eq 0) {
        Write-Log "No valid files to merge after rename step." "Info"
        Set-StatusText "Status: Nothing to merge."
        return
    }
    $listFile = Join-Path $VidFolder "video_list.txt"
    if (Test-Path $listFile) { Remove-Item $listFile -Force }
    foreach ($tf in ($tempFiles | Sort-Object)) {
        "file '$tf'" | Out-File -FilePath $listFile -Append -Encoding UTF8
    }
    Write-Log "Created list file with $($tempFiles.Count) item(s)." "Info"
    $ffmpegArgs = @(
        "-loglevel", "error",
        "-hide_banner",
        "-y",
        "-f", "concat",
        "-safe", "0",
        "-i", "`"$listFile`"",
        "-c", "copy",
        "`"$MasterVideoPath`""
    )
    Write-Log "Running FFmpeg final merge => $MasterVideoPath" "Info"
    try {
        $process = Start-Process -FilePath $global:FfmpegPath -ArgumentList $ffmpegArgs -NoNewWindow -Wait -PassThru
        if ($process.ExitCode -eq 0) {
            Write-Log "Successfully created/updated master timelapse => $MasterVideoPath" "Info"
            Set-StatusText "Status: Master video updated."
            $tempFileInfos = $tempFiles | ForEach-Object { New-Object System.IO.FileInfo($_) }
            RemoveOrArchiveVideos -Files $tempFileInfos -Reason "After master merge"
        }
        else {
            Write-Log "FFmpeg merge failed => exit code $($process.ExitCode)" "Error"
            Set-StatusText "Status: Failed to merge videos."
            if ($tempFiles.Count -gt 0 -and (Test-Path $tempFiles[0])) {
                try { Rename-Item -Path $tempFiles[0] -NewName (Split-Path $MasterVideoPath -Leaf) -Force }
                catch {}
            }
        }
    }
    catch {
        Write-Log "Error merging final timelapse => $_" "Error"
        Set-StatusText "Status: Error merging final timelapse."
        if ($tempFiles.Count -gt 0 -and (Test-Path $tempFiles[0])) {
            try { Rename-Item -Path $tempFiles[0] -NewName (Split-Path $MasterVideoPath -Leaf) -Force }
            catch {}
        }
    }
    finally {
        if (Test-Path $listFile) {
            Remove-Item $listFile -Force
            Write-Log "Deleted $listFile" -NoNewline
        }
    }
}

###############################################################################
#                               CAMERA HANDLING                                #
###############################################################################
function SwitchCamera {
    param([int]$DeviceIndex)
    if ($global:isSwitchingCamera) {
        Write-Log "SwitchCamera is already running. Skipping." "Warning"
        return
    }
    $global:isSwitchingCamera = $true
    try {
        Write-Log "SwitchCamera called with DeviceIndex: $DeviceIndex" "Info"
        if ($DeviceIndex -lt 0 -or $DeviceIndex -ge $devices.Count) {
            Write-Log "Invalid device index selected." "Warning"
            return
        }
        Write-Log "Switching to camera: $($devices[$DeviceIndex].Name)" "Info"
        if ($global:videoSrc -and $global:videoSrc.IsRunning) {
            $global:videoSrc.SignalToStop()
            $global:videoSrc.WaitForStop()
        }
        if ($global:videoSrc) {
            if ($global:videoSrc -is [System.IDisposable]) {
                try { $global:videoSrc.Dispose() }
                catch { Write-Log "Error disposing videoSrc: $_" "Error" }
            }
        }
        $global:videoSrc = New-Object AForge.Video.DirectShow.VideoCaptureDevice($devices[$DeviceIndex].MonikerString)
        if (-not $global:videoSrc) {
            Write-Log "Failed to initialize the video source." "Error"
            return
        }
        Write-Log "videoSrc initialized successfully." "Info"
        $allCaps = $global:videoSrc.VideoCapabilities
        if ($allCaps.Count -gt 0) {
            $maxCap = $allCaps | Sort-Object { $_.FrameSize.Width * $_.FrameSize.Height } -Descending | Select-Object -First 1
            if ($maxCap) {
                $global:videoSrc.VideoResolution = $maxCap
                $global:SelectedWidth = $maxCap.FrameSize.Width
                $global:SelectedHeight = $maxCap.FrameSize.Height
            }
        }
        $global:cmbCameraResolution.Items.Clear()
        foreach ($capability in $global:videoSrc.VideoCapabilities) {
            $resString = "$($capability.FrameSize.Width) x $($capability.FrameSize.Height) @ $($capability.MaximumFrameRate) FPS"
            $itemObj = New-Object PSObject -Property @{ 
                Width     = $capability.FrameSize.Width; 
                Height    = $capability.FrameSize.Height; 
                FrameRate = $capability.MaximumFrameRate; 
                Display   = $resString 
            }
            $global:cmbCameraResolution.Items.Add($itemObj) | Out-Null
        }
        $global:cmbCameraResolution.DisplayMember = "Display"
        if ($global:cmbCameraResolution.Items.Count -gt 0) {
            $matchIndex = -1
            for ($i = 0; $i -lt $global:cmbCameraResolution.Items.Count; $i++) {
                $item = $global:cmbCameraResolution.Items[$i]
                if (($item.Width -eq $global:SelectedWidth) -and ($item.Height -eq $global:SelectedHeight)) {
                    $matchIndex = $i
                    break
                }
            }
            if ($matchIndex -ge 0) { $global:cmbCameraResolution.SelectedIndex = $matchIndex }
            else { $global:cmbCameraResolution.SelectedIndex = 0 }
            Save-Config
        }
        else {
            Write-Log "No resolution capabilities found for this device." "Error"
        }
        $global:videoSrc.DesiredFrameRate = $CameraFps
        $videoPlayer.VideoSource = $global:videoSrc
        if (-not $videoPlayer.IsRunning) { $videoPlayer.Start() }
        $videoPlayer.BringToFront()
        Start-Sleep -Milliseconds 500
        $videoPlayer.Refresh()
        if ($videoPlayer.IsRunning) {
            Write-Log "Camera switched and video preview started successfully." "Info"
            Set-StatusText "Status: Preview Running"
        }
        else {
            Write-Log "Video preview failed to start." "Error"
            Set-StatusText "Status: Preview Failed"
            # (Call to reinitialize has been removed.)
        }
        $global:SelectedCameraIndex = $DeviceIndex
        Save-Config
    }
    catch {
        Write-Log "Error switching camera: $_" "Error"
    }
    finally {
        $global:isSwitchingCamera = $false
    }
}

###############################################################################
#                             TIMELAPSE CALCULATOR                             #
###############################################################################
function Update-TimelapseCalculator {
    $hoursParsed = 0
    $minutesParsed = 0
    if ($global:cmbCapturingDurationHours -and -not $global:cmbCapturingDurationHours.IsDisposed) {
        $hoursParsed = [int]$global:cmbCapturingDurationHours.SelectedItem
    }
    if ($global:cmbCapturingDurationMinutes -and -not $global:cmbCapturingDurationMinutes.IsDisposed) {
        $minutesParsed = [int]$global:cmbCapturingDurationMinutes.SelectedItem
    }
    $totalShootingSeconds = ($hoursParsed * 3600) + ($minutesParsed * 60)
    $captureInterval = ($global:CaptureMinutes * 60) + $global:CaptureSeconds
    if ($captureInterval -le 0) { $captureInterval = 5 }
    $totalFrames = [math]::Floor($totalShootingSeconds / $captureInterval)
    if ($global:txtTotalFrames -and -not $global:txtTotalFrames.IsDisposed) {
        SafeInvoke $global:txtTotalFrames { $global:txtTotalFrames.Text = "$totalFrames" }
    }
    $finalVideoSeconds = if ($global:OutputFps -gt 0) { $totalFrames / $global:OutputFps } else { 0 }
    if ($finalVideoSeconds -le 0) { $formatted = "0" }
    else {
        $hrs = [int]([math]::Floor($finalVideoSeconds / 3600))
        $remaining = $finalVideoSeconds % 3600
        $mins = [int]([math]::Floor($remaining / 60))
        $secs = [int]([math]::Floor($remaining % 60))
        $formatted = "$hrs hr : $mins min : $secs sec"
    }
    if ($global:txtFinalVideoDuration -and -not $global:txtFinalVideoDuration.IsDisposed) {
        SafeInvoke $global:txtFinalVideoDuration { $global:txtFinalVideoDuration.Text = $formatted }
    }
    $estimatedSizeBytes = $global:SelectedWidth * $global:SelectedHeight * 1.5 * $totalFrames
    $estimatedSizeMB = $estimatedSizeBytes / (1024*1024)
    if ($estimatedSizeMB -ge 1024) {
        $estimatedSizeGB = [math]::Round($estimatedSizeMB / 1024, 2)
        $sizeString = "$estimatedSizeGB GB"
    } else {
        $estimatedSizeMB = [math]::Round($estimatedSizeMB, 2)
        $sizeString = "$estimatedSizeMB MB"
    }
    if ($global:txtEstimatedVideoSize -and -not $global:txtEstimatedVideoSize.IsDisposed) {
        SafeInvoke $global:txtEstimatedVideoSize { $global:txtEstimatedVideoSize.Text = $sizeString }
    }
}

###############################################################################
#                           FULLSCREEN PREVIEW FUNCTION                      #
###############################################################################
function Show-FullscreenPreview {
    $fsForm = New-Object System.Windows.Forms.Form
    $fsForm.Text = "MasterLapse   |   Timelapse Application   |   Fullscreen Preview"
    $fsForm.FormBorderStyle = "Sizable"
    $fsForm.WindowState = "Maximized"
    $fsForm.BackColor = [System.Drawing.Color]::Black
    if (Test-Path $iconPath) { $fsForm.Icon = New-Object System.Drawing.Icon($iconPath) }
    $fsVideoPlayer = New-Object AForge.Controls.VideoSourcePlayer
    $fsVideoPlayer.Dock = [System.Windows.Forms.DockStyle]::Fill
    $fsVideoPlayer.BackColor = [System.Drawing.Color]::Black
    $fsVideoPlayer.add_Paint({
        param($sender, $e)
        if ($fsVideoPlayer.Image) {
            $img = $fsVideoPlayer.Image
            $rect = $sender.ClientRectangle
            $aspectRatio = $img.Width / $img.Height
            $newWidth = $rect.Width
            $newHeight = [int]($rect.Width / $aspectRatio)
            if ($newHeight -gt $rect.Height) {
                $newHeight = $rect.Height
                $newWidth = [int]($rect.Height * $aspectRatio)
            }
            $x = ($rect.Width - $newWidth) / 2
            $y = ($rect.Height - $newHeight) / 2
            $e.Graphics.Clear($sender.BackColor)
            $e.Graphics.DrawImage($img, $x, $y, $newWidth, $newHeight)
            $e.Handled = $true
        }
    })
    $fsVideoPlayer.Add_Resize({ $fsVideoPlayer.Invalidate() })
    $fsForm.Controls.Add($fsVideoPlayer)
    if ($global:videoSrc) {
        $fsVideoPlayer.VideoSource = $global:videoSrc
        $fsVideoPlayer.Start()
    }
    $fsForm.ShowDialog() | Out-Null
    try {
        if ($fsVideoPlayer.IsRunning) {
            try { $fsVideoPlayer.Stop() }
            catch {
                if ($_.Exception.Message -match "Thread abort") { }
                else { Write-Log "Error stopping fullscreen preview: $($_.Exception.Message)" "Warning" }
            }
        }
    }
    catch { Write-Log "Error in stopping fullscreen preview: $_" "Warning" }
}

###############################################################################
#                               MAIN INITIALIZATION                              #
###############################################################################
Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.Drawing       | Out-Null
$AForgePath = "C:\Program Files (x86)\AForge.NET\Framework\Release"
try {
    Add-Type -Path (Join-Path $AForgePath "AForge.Controls.dll") | Out-Null
    Add-Type -Path (Join-Path $AForgePath "AForge.dll")           | Out-Null
    Add-Type -Path (Join-Path $AForgePath "AForge.Video.dll")     | Out-Null
    Add-Type -Path (Join-Path $AForgePath "AForge.Video.DirectShow.dll") | Out-Null
    Write-Log "AForge.NET assemblies loaded successfully." "Info"
}
catch {
    Write-Log "Error loading AForge.NET assemblies: $_" "Error"
    exit
}
$devices = New-Object AForge.Video.DirectShow.FilterInfoCollection([AForge.Video.DirectShow.FilterCategory]::VideoInputDevice)
if ($devices.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("No webcam detected!","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    Write-Log "No video devices detected." "Error"
    exit
}
else {
    Write-Log "Detected $($devices.Count) video device(s)." "Info"
}
$global:isInitializing    = $true
$global:isSwitchingCamera = $false
Load-Config
foreach ($dir in @($global:VideoFolder, $global:ImageFolder)) {
    if (-not (Test-Path $dir)) {
        Write-Log "Creating directory: $($dir)"
        try {
            New-Item -ItemType Directory -Path $dir | Out-Null
            Write-Log "Directory created: $($dir)"
        }
        catch {
            Write-Log "Failed to create directory $($dir): $_" "Error"
            exit
        }
    }
    else {
        Write-Log "Directory exists: $($dir)" "Info"
    }
}
try {
    $nextNumber = GetNextImageNumber -ImgFolder $global:ImageFolder
    [int]$global:imgCounter = $nextNumber
    Write-Log "Initialized image counter to: $($global:imgCounter)" "Info"
}
catch {
    Write-Log "Error initializing image counter: $_" "Error"
    [int]$global:imgCounter = 1
}
$global:sessionImageCount = 0
$global:captureTimer = New-Object System.Windows.Forms.Timer
$global:captureTimer.Enabled = $false
if (($global:CaptureMinutes -ne $null) -and ($global:CaptureSeconds -ne $null)) {
    $defaultInterval = (($global:CaptureMinutes * 60) + $global:CaptureSeconds)
    if ($defaultInterval -le 0) { $defaultInterval = 5 }
    $global:captureTimer.Interval = $defaultInterval * 1000
}
else {
    $global:captureTimer.Interval = 5000
}
$global:autoCloseTimer = New-Object System.Windows.Forms.Timer

###############################################################################
# CAPTURE TIMER (Image every Interval)
###############################################################################
$global:captureTimer.Add_Tick({
    Write-Log "Capture timer tick event fired." "Info" -NoConsole
    [System.Windows.Forms.Application]::DoEvents()
    if ($videoPlayer.IsRunning) {
        try {
            $frame = $videoPlayer.GetCurrentVideoFrame()
            if (-not $frame) {
                Write-Log "No frame captured; skipping." "Warning"
                $global:NoFrameWarningCount++
                return
            }
            else { $global:NoFrameWarningCount = 0 }
            if (-not ($frame -is [System.Drawing.Bitmap])) {
                Write-Log "Captured frame is not a valid bitmap; skipping." "Warning"
                return
            }
            if ($global:EnableTimestamp) {
                $frame = Add-TimestampToImage $frame
                if (-not ($frame -is [System.Drawing.Bitmap])) { return }
            }
            $filename = Join-Path $global:ImageFolder ("capture_{0:D6}.png" -f $global:imgCounter)
            $frame.Save($filename, [System.Drawing.Imaging.ImageFormat]::Png)
            $frame.Dispose()
            $global:sessionImageCount++
            Write-Host ("Capturing Images: {0}`r" -f $global:sessionImageCount) -ForegroundColor Green -NoNewline
            $global:imgCounter++
        }
        catch {
            Write-Log "Error capturing image: $_" "Error"
        }
    }
    else {
        Write-Log "videoPlayer is not running. Can't capture image." "Warning"
    }
})

###############################################################################
# AUTO-CLOSE TIMER (Triggered after Auto-Restart interval)
###############################################################################
$global:autoCloseTimer.Add_Tick({
    Write-Log "Auto close timer triggered. Checking conditions..." "Info"
    $global:autoCloseTimer.Stop()
    if ($global:captureTimer -and $global:captureTimer.Enabled) {
        Write-Log "Auto close event: capture is in progress. Stopping capture timer." "Info"
        $global:captureTimer.Stop()
        Write-Host "`nCaptured Images: $global:sessionImageCount" -ForegroundColor Green
        Set-StatusText "Capture Stopped. Total Images Captured: $global:sessionImageCount"
        Write-Log "Auto close event: merging captured images into a video." "Info"
        MergeImagesIntoVideo -Folder $global:ImageFolder -VidFolder $global:VideoFolder `
            -FPS $global:OutputFps -CrfVal $CrfValue -PresetVal $PresetVal `
            -DarkThr $global:DarkThreshold -OutW $global:SelectedWidth -OutH $global:SelectedHeight
        try {
            [int]$global:imgCounter = (GetNextImageNumber -ImgFolder $global:ImageFolder)
            Write-Log "Reset image counter to $global:imgCounter after auto-close merge." "Info"
        }
        catch {
            Write-Log "Error recalculating image counter after auto-close merge: $_" "Error"
            [int]$global:imgCounter = 1
        }
    }
    if ($global:AutoMergeVideoLimit -gt 0) {
        $videoFiles = Get-ChildItem -Path $global:VideoFolder -Filter "video_*.mp4" -File -ErrorAction SilentlyContinue
        if ($videoFiles.Count -ge $global:AutoMergeVideoLimit) {
            Write-Log "Auto close event: we have $($videoFiles.Count) videos (>= limit $($global:AutoMergeVideoLimit)). Merging all." "Info"
            Set-StatusText "AutoRestart: Merging all videos..."
            MergeAllVideos -VidFolder $global:VideoFolder -MasterVideoPath (Join-Path $global:VideoFolder "MasterLapse.mp4")
        }
    }
    Set-StatusText "AutoRestart: Exiting and restarting..."
    Write-Log "Closing form due to auto-close event." "Info"
    $form.Close()
})

function Update-AutoCloseTimer {
    if (-not $global:AutoRestart) {
        Write-Log "AutoRestart is off. Stopping autoCloseTimer." "Info"
        $global:autoCloseTimer.Stop()
        return
    }
    $totalMinutes = ($global:AutoRestartDays * 24 * 60) + ($global:AutoRestartHours * 60) + $global:AutoRestartMinutes
    if ($totalMinutes -le 0) {
        Write-Log "No auto-restart time set. Stopping timer." "Info"
        $global:autoCloseTimer.Stop()
        return
    }
    [long]$intervalMs = $totalMinutes * 60 * 1000
    if ($intervalMs -gt [int]::MaxValue) { $intervalMs = [int]::MaxValue }
    $global:autoCloseTimer.Interval = [int]$intervalMs
    $global:autoCloseTimer.Start()
    Write-Log "AutoCloseTimer set to single-shot $totalMinutes minute(s)." "Info"
}

function Immediate-SaveCaptureSeconds {
    param([string]$newValue)
    if ($newValue -match '^\d+$') {
        $global:CaptureSeconds = [int]$newValue
        Write-Log "CaptureSeconds changed to $($global:CaptureSeconds)"
        Save-Config
        Update-TimelapseCalculator
    }
}
function Immediate-SaveCaptureMinutes {
    param([string]$newValue)
    if ($newValue -match '^\d+$') {
        $global:CaptureMinutes = [int]$newValue
        Write-Log "CaptureMinutes changed to $($global:CaptureMinutes)"
        Save-Config
        Update-TimelapseCalculator
    }
}

###############################################################################
#                                 GUI SETUP                                    #
###############################################################################
Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.Drawing | Out-Null
[void][System.Windows.Forms.Application]::EnableVisualStyles()
$form = New-Object System.Windows.Forms.Form
$form.Text          = "MasterLapse   |   Timelapse Application"
$form.StartPosition = "CenterScreen"
$form.Size          = New-Object System.Drawing.Size(1200,1000)
$form.MinimumSize   = New-Object System.Drawing.Size(1200,900)
$form.BackColor     = [System.Drawing.ColorTranslator]::FromHtml("#2E2E2E")
$iconPath = "C:\Program Files (x86)\Timelapse\Icon.ico"
if (Test-Path $iconPath) { $form.Icon = New-Object System.Drawing.Icon($iconPath) }
$commonFont = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Regular)
$videoPlayer = New-Object AForge.Controls.VideoSourcePlayer
$videoPlayer.Dock      = [System.Windows.Forms.DockStyle]::Fill
$videoPlayer.BackColor = [System.Drawing.Color]::Black
$form.Controls.Add($videoPlayer) | Out-Null
[void]$videoPlayer.Handle
$videoPlayer.Add_Resize({ $videoPlayer.Invalidate() })
$videoPlayer.add_Paint({
    param($sender, $e)
    if ($videoPlayer.Image) {
        $img = $videoPlayer.Image
        $rect = $sender.ClientRectangle
        $aspectRatio = $img.Width / $img.Height
        $newWidth = $rect.Width
        $newHeight = [int]($rect.Width / $aspectRatio)
        if ($newHeight -gt $rect.Height) {
            $newHeight = $rect.Height
            $newWidth = [int]($rect.Height * $aspectRatio)
        }
        $x = ($rect.Width - $newWidth) / 2
        $y = ($rect.Height - $newHeight) / 2
        $e.Graphics.Clear($sender.BackColor)
        $e.Graphics.DrawImage($img, $x, $y, $newWidth, $newHeight)
        $e.Handled = $true
    }
})
$panelBottom = New-Object System.Windows.Forms.Panel
$panelBottom.Dock      = [System.Windows.Forms.DockStyle]::Bottom
$panelBottom.Height    = 500
$panelBottom.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#1E1E2E")
$form.Controls.Add($panelBottom) | Out-Null
$bottomLayout = New-Object System.Windows.Forms.TableLayoutPanel
$bottomLayout.Dock       = [System.Windows.Forms.DockStyle]::Fill
$bottomLayout.RowCount   = 2
$bottomLayout.ColumnCount= 1
$bottomLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,80)))
$bottomLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)))
$panelBottom.Controls.Add($bottomLayout) | Out-Null
$statusPanel = New-Object System.Windows.Forms.TableLayoutPanel
$statusPanel.Dock       = [System.Windows.Forms.DockStyle]::Fill
$statusPanel.RowCount   = 1
$statusPanel.ColumnCount= 2
$statusPanel.ColumnStyles.Clear()
$statusPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50)))
$statusPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50)))
$recordingLabelBottom = New-Object System.Windows.Forms.Label
$recordingLabelBottom.Dock      = [System.Windows.Forms.DockStyle]::Fill
$recordingLabelBottom.ForeColor = [System.Drawing.Color]::Red
$recordingLabelBottom.TextAlign = 'MiddleLeft'
$recordingLabelBottom.Text      = "Recording: OFF"
$statusPanel.Controls.Add($recordingLabelBottom,0,0)
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Dock      = [System.Windows.Forms.DockStyle]::Fill
$statusLabel.ForeColor = [System.Drawing.Color]::White
$statusLabel.TextAlign = 'MiddleLeft'
$statusLabel.Text      = "Status: Ready"
$statusPanel.Controls.Add($statusLabel,1,0)
$global:statusLabel = $statusLabel
$bottomLayout.Controls.Add($statusPanel,0,1)
$threePanel = New-Object System.Windows.Forms.TableLayoutPanel
$threePanel.Dock        = [System.Windows.Forms.DockStyle]::Fill
$threePanel.RowCount    = 1
$threePanel.ColumnCount = 3
for ($i=0; $i -lt 3; $i++){
    $threePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,33.3)))
}
$bottomLayout.Controls.Add($threePanel,0,0)

###############################################################################
#                        GROUPBOX 1: Capture Settings
###############################################################################
$gbCapture = New-Object System.Windows.Forms.GroupBox
$gbCapture.Text      = "Capture Settings"
$gbCapture.ForeColor = [System.Drawing.Color]::White
$gbCapture.Font      = $commonFont
$gbCapture.Dock      = [System.Windows.Forms.DockStyle]::Fill
$gbCapture.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#1E1E1E")
$threePanel.Controls.Add($gbCapture,0,0)
$captureLayout = New-Object System.Windows.Forms.TableLayoutPanel
$captureLayout.Dock        = [System.Windows.Forms.DockStyle]::Fill
$captureLayout.RowCount    = 10
$captureLayout.ColumnCount = 1
$captureLayout.Padding     = '5,5,5,5'
$captureLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)))
$captureLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,60)))
$captureLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,40)))
$captureLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,40)))
$captureLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,40)))
$captureLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,40)))
$captureLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)))
$captureLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)))
$captureLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,80)))
$captureLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,40)))
$captureLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
$gbCapture.Controls.Add($captureLayout)
$lblIntervalTitle = New-Object System.Windows.Forms.Label
$lblIntervalTitle.Text      = "Intervals Between Captures"
$lblIntervalTitle.Font      = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$lblIntervalTitle.ForeColor = [System.Drawing.Color]::White
$lblIntervalTitle.TextAlign = "MiddleCenter"
$lblIntervalTitle.Dock      = [System.Windows.Forms.DockStyle]::Fill
$captureLayout.Controls.Add($lblIntervalTitle,0,0)
$tlpIntervals = New-Object System.Windows.Forms.TableLayoutPanel
$tlpIntervals.Dock        = [System.Windows.Forms.DockStyle]::Fill
$tlpIntervals.RowCount    = 2
$tlpIntervals.ColumnCount = 2
$tlpIntervals.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)))
$tlpIntervals.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)))
$tlpIntervals.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50)))
$tlpIntervals.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50)))
$lblSecLabel = New-Object System.Windows.Forms.Label
$lblSecLabel.Text      = "Intervals in Seconds"
$lblSecLabel.ForeColor = [System.Drawing.Color]::White
$lblSecLabel.Font      = $commonFont
$lblSecLabel.TextAlign = 'MiddleCenter'
$lblSecLabel.Dock      = [System.Windows.Forms.DockStyle]::Fill
$tlpIntervals.Controls.Add($lblSecLabel,0,0)
$lblMinLabel = New-Object System.Windows.Forms.Label
$lblMinLabel.Text      = "Intervals in Minutes"
$lblMinLabel.ForeColor = [System.Drawing.Color]::White
$lblMinLabel.Font      = $commonFont
$lblMinLabel.TextAlign = 'MiddleCenter'
$lblMinLabel.Dock      = [System.Windows.Forms.DockStyle]::Fill
$tlpIntervals.Controls.Add($lblMinLabel,1,0)
# Replace textboxes with ComboBoxes for interval selection.
$cmbSec = New-Object System.Windows.Forms.ComboBox
$cmbSec.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbSec.Font = $commonFont
$cmbSec.Dock = [System.Windows.Forms.DockStyle]::Fill
for ($i = 0; $i -le 60; $i++) {
    $null = $cmbSec.Items.Add($i)
}
$cmbSec.SelectedItem = $global:CaptureSeconds
$tlpIntervals.Controls.Add($cmbSec,0,1)
$cmbMin = New-Object System.Windows.Forms.ComboBox
$cmbMin.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbMin.Font = $commonFont
$cmbMin.Dock = [System.Windows.Forms.DockStyle]::Fill
for ($i = 0; $i -le 999; $i++) {
    $null = $cmbMin.Items.Add($i)
}
$cmbMin.SelectedItem = $global:CaptureMinutes
$tlpIntervals.Controls.Add($cmbMin,1,1)
# Replace textchanged events with SelectedIndexChanged events.
$cmbSec.Add_SelectedIndexChanged({ Immediate-SaveCaptureSeconds $cmbSec.SelectedItem; Update-TimelapseCalculator })
$cmbMin.Add_SelectedIndexChanged({ Immediate-SaveCaptureMinutes $cmbMin.SelectedItem; Update-TimelapseCalculator })
$captureLayout.Controls.Add($tlpIntervals,0,1)
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text      = "Start Capture"
$btnStart.Font      = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$btnStart.ForeColor = [System.Drawing.Color]::White
$btnStart.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#3C3C3C")
$btnStart.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnStart.Dock      = [System.Windows.Forms.DockStyle]::Fill
$captureLayout.Controls.Add($btnStart,0,2)
$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text      = "Stop Capture"
$btnStop.Font      = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$btnStop.ForeColor = [System.Drawing.Color]::White
$btnStop.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#3C3C3C")
$btnStop.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnStop.Dock      = [System.Windows.Forms.DockStyle]::Fill
$captureLayout.Controls.Add($btnStop,0,3)
$btnForce = New-Object System.Windows.Forms.Button
$btnForce.Text      = "Force Video"
$btnForce.Font      = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$btnForce.ForeColor = [System.Drawing.Color]::White
$btnForce.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#3C3C3C")
$btnForce.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnForce.Dock      = [System.Windows.Forms.DockStyle]::Fill
$captureLayout.Controls.Add($btnForce,0,4)
$btnMergeAllVideos = New-Object System.Windows.Forms.Button
$btnMergeAllVideos.Text      = "Merge All Videos"
$btnMergeAllVideos.Font      = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$btnMergeAllVideos.ForeColor = [System.Drawing.Color]::White
$btnMergeAllVideos.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#3C3C3C")
$btnMergeAllVideos.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnMergeAllVideos.Dock      = [System.Windows.Forms.DockStyle]::Fill
$captureLayout.Controls.Add($btnMergeAllVideos,0,5)
$chkDontRemoveImages = New-Object System.Windows.Forms.CheckBox
$chkDontRemoveImages.Text      = "Don't remove Old Images"
$chkDontRemoveImages.ForeColor = [System.Drawing.Color]::White
$chkDontRemoveImages.Font      = $commonFont
$chkDontRemoveImages.Dock      = [System.Windows.Forms.DockStyle]::Fill
$chkDontRemoveImages.Checked   = $global:DontRemoveOldImages
$captureLayout.Controls.Add($chkDontRemoveImages,0,6)
$chkDontRemoveImages.Add_CheckedChanged({
    $global:DontRemoveOldImages = $chkDontRemoveImages.Checked
    Write-Log "DontRemoveOldImages -> $($global:DontRemoveOldImages)" "Info"
    Save-Config
})
$chkDontRemoveVideos = New-Object System.Windows.Forms.CheckBox
$chkDontRemoveVideos.Text      = "Don't remove Old Videos"
$chkDontRemoveVideos.ForeColor = [System.Drawing.Color]::White
$chkDontRemoveVideos.Font      = $commonFont
$chkDontRemoveVideos.Dock      = [System.Windows.Forms.DockStyle]::Fill
$chkDontRemoveVideos.Checked   = $global:DontRemoveOldVideos
$captureLayout.Controls.Add($chkDontRemoveVideos,0,7)
$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text = 'Note: If "Don''t remove Old Images/Videos" are checked on, the images/videos get saved to a sub-folder inside your "Set Folders". If they are unchecked, the images/videos are auto-removed once no longer needed.'
$lblNote.ForeColor = [System.Drawing.Color]::White
$lblNote.Font = $commonFont
$lblNote.TextAlign = 'MiddleCenter'
$lblNote.Dock = [System.Windows.Forms.DockStyle]::Fill
$captureLayout.Controls.Add($lblNote,0,8)
$btnOpenLog = New-Object System.Windows.Forms.Button
$btnOpenLog.Text = "Open Log"
$btnOpenLog.Font = $commonFont
$btnOpenLog.ForeColor = [System.Drawing.Color]::White
$btnOpenLog.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#3C3C3C")
$btnOpenLog.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnOpenLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnOpenLog.Add_Click({
    if (Test-Path $global:LogFile) { Start-Process notepad.exe -ArgumentList $global:LogFile }
    else { [System.Windows.Forms.MessageBox]::Show("Log file not found.","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) }
})
$captureLayout.Controls.Add($btnOpenLog,0,9)

###############################################################################
#                  GROUPBOX 2: Camera
###############################################################################
$gbCamera = New-Object System.Windows.Forms.GroupBox
$gbCamera.Text      = "Camera"
$gbCamera.ForeColor = [System.Drawing.Color]::White
$gbCamera.Font      = $commonFont
$gbCamera.Dock      = [System.Windows.Forms.DockStyle]::Fill
$gbCamera.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#1E1E1E")
$threePanel.Controls.Add($gbCamera,1,0)
$cameraLayout = New-Object System.Windows.Forms.TableLayoutPanel
$cameraLayout.Dock         = [System.Windows.Forms.DockStyle]::Fill
$cameraLayout.RowCount     = 14
$cameraLayout.ColumnCount  = 2
$cameraLayout.Padding      = '5,5,5,5'
for ($r=0; $r -lt 14; $r++) {
    $cameraLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)))
}
$cameraLayout.RowStyles[7].Height = 60
for ($c=0; $c -lt 2; $c++) {
    $cameraLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50)))
}
$gbCamera.Controls.Add($cameraLayout)
$lblSelectCamera = New-Object System.Windows.Forms.Label
$lblSelectCamera.Text      = "Select Camera:"
$lblSelectCamera.ForeColor = [System.Drawing.Color]::White
$lblSelectCamera.Font      = $commonFont
$lblSelectCamera.TextAlign = "MiddleRight"
$lblSelectCamera.Dock      = [System.Windows.Forms.DockStyle]::Fill
$cameraLayout.Controls.Add($lblSelectCamera,0,0)
$cmbSelectCamera = New-Object System.Windows.Forms.ComboBox
$cmbSelectCamera.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbSelectCamera.Font          = $commonFont
$cmbSelectCamera.Dock          = [System.Windows.Forms.DockStyle]::Fill
$cameraLayout.Controls.Add($cmbSelectCamera,1,0)
$cmbSelectCamera.Add_SelectedIndexChanged({
    if (-not $global:isInitializing) {
        $selectedIndex = $cmbSelectCamera.SelectedIndex
        if ($selectedIndex -lt 0) {
            Write-Log "No camera selected." "Info"
            $videoPlayer.VideoSource = $null
            try {
                if ($videoPlayer.IsRunning) { 
                    try { $videoPlayer.Stop() }
                    catch {
                        if ($_.Exception.Message -notmatch "Thread abort is not supported") {
                            Write-Log "Error stopping videoPlayer: $_" "Warning"
                        }
                    }
                }
            }
            catch { }
            Set-StatusText "Status: No camera selected."
            $global:SelectedCameraIndex = -1
            Save-Config
        }
        else {
            $global:SelectedCameraIndex = $selectedIndex
            SwitchCamera -DeviceIndex $selectedIndex
        }
    }
}) | Out-Null
$btnProps = New-Object System.Windows.Forms.Button
$btnProps.Text      = "Camera Properties"
$btnProps.Font      = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$btnProps.ForeColor = [System.Drawing.Color]::White
$btnProps.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#3C3C3C")
$btnProps.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnProps.Dock      = [System.Windows.Forms.DockStyle]::Fill
$btnProps.Add_Click({
    if ($global:videoSrc -and $form -and $form.Handle) {
        Write-Log "Opening camera property page..." "Info"
        try { $global:videoSrc.DisplayPropertyPage($form.Handle) }
        catch { Write-Log "Error opening camera properties: $_" "Error" }
    }
    else { Write-Log "Camera property page not available." "Warning" }
}) | Out-Null
$cameraLayout.Controls.Add($btnProps,0,1)
$cameraLayout.SetColumnSpan($btnProps,2)
$lblCameraResolution = New-Object System.Windows.Forms.Label
$lblCameraResolution.Text      = "Camera Resolution:"
$lblCameraResolution.ForeColor = [System.Drawing.Color]::White
$lblCameraResolution.TextAlign = 'MiddleRight'
$lblCameraResolution.Dock      = [System.Windows.Forms.DockStyle]::Fill
$cameraLayout.Controls.Add($lblCameraResolution,0,2)
$cmbCameraResolution = New-Object System.Windows.Forms.ComboBox
$cmbCameraResolution.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbCameraResolution.Font = $commonFont
$cmbCameraResolution.Dock = [System.Windows.Forms.DockStyle]::Fill
$cmbCameraResolution.DisplayMember = "Display"
$cameraLayout.Controls.Add($cmbCameraResolution,1,2)
$global:cmbCameraResolution = $cmbCameraResolution
$cmbCameraResolution.Add_SelectedIndexChanged({
    if ($global:videoSrc) {
        $selectedRes = $global:cmbCameraResolution.SelectedItem
        if ($selectedRes -ne $null) {
            Write-Log "Changing resolution to $($selectedRes.Width)x$($selectedRes.Height)" "Info"
            if ($videoPlayer.IsRunning) { 
                try { $videoPlayer.Stop() }
                catch {
                    if ($_.Exception.Message -notmatch "Thread abort is not supported") {
                        Write-Log "Error stopping videoPlayer: $_" "Warning"
                    }
                }
            }
            $global:videoSrc.SignalToStop()
            $global:videoSrc.WaitForStop()
            $global:videoSrc.VideoResolution = $global:videoSrc.VideoCapabilities | Where-Object { $_.FrameSize.Width -eq $selectedRes.Width -and $_.FrameSize.Height -eq $selectedRes.Height } | Select-Object -First 1
            $global:SelectedWidth = $selectedRes.Width
            $global:SelectedHeight = $selectedRes.Height
            Save-Config
            $videoPlayer.VideoSource = $global:videoSrc
            $videoPlayer.Start()
            Write-Log "Resolution changed and preview restarted." "Info"
            Update-TimelapseCalculator
        }
    }
})
$lblVideoFramerate = New-Object System.Windows.Forms.Label
$lblVideoFramerate.Text      = "Final Video Framerate:"
$lblVideoFramerate.ForeColor = [System.Drawing.Color]::White
$lblVideoFramerate.TextAlign = 'MiddleRight'
$lblVideoFramerate.Dock      = [System.Windows.Forms.DockStyle]::Fill
$cameraLayout.Controls.Add($lblVideoFramerate,0,3)
$cmbVideoFramerate = New-Object System.Windows.Forms.ComboBox
$cmbVideoFramerate.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbVideoFramerate.Font          = $commonFont
$cmbVideoFramerate.Dock          = [System.Windows.Forms.DockStyle]::Fill
$popularRates = @(10,15,24,25,30,48,50,60,90,120,144,240,360)
foreach ($rate in $popularRates) { $null = $cmbVideoFramerate.Items.Add($rate) }
$cmbVideoFramerate.SelectedItem = $global:OutputFps
$cameraLayout.Controls.Add($cmbVideoFramerate,1,3)
$cmbVideoFramerate.Add_SelectedIndexChanged({
    $global:OutputFps = [int]$cmbVideoFramerate.SelectedItem
    Write-Log "Output framerate manually set to $global:OutputFps" "Info"
    Save-Config
    Update-TimelapseCalculator
})
$lblDarkSensitivity = New-Object System.Windows.Forms.Label
$lblDarkSensitivity.Text      = "Dark Image Sensitivity:"
$lblDarkSensitivity.ForeColor = [System.Drawing.Color]::White
$lblDarkSensitivity.TextAlign = 'MiddleRight'
$lblDarkSensitivity.Dock      = [System.Windows.Forms.DockStyle]::Fill
$cameraLayout.Controls.Add($lblDarkSensitivity,0,4)
$cmbDarkSensitivity = New-Object System.Windows.Forms.ComboBox
$cmbDarkSensitivity.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbDarkSensitivity.Font          = $commonFont
$cmbDarkSensitivity.Dock          = [System.Windows.Forms.DockStyle]::Fill
for ($i=0; $i -le 255; $i++) { $null = $cmbDarkSensitivity.Items.Add($i) }
$cmbDarkSensitivity.SelectedItem = $global:DarkThreshold
$cameraLayout.Controls.Add($cmbDarkSensitivity,1,4)
$cmbDarkSensitivity.Add_SelectedIndexChanged({
    $global:DarkThreshold = [int]$cmbDarkSensitivity.SelectedItem
    Write-Log "Dark Image Sensitivity updated to: $global:DarkThreshold" "Info"
    Save-Config
})
$lblTimelapseCalcHeader = New-Object System.Windows.Forms.Label
$lblTimelapseCalcHeader.Text      = "Timelapse Calculator"
$lblTimelapseCalcHeader.ForeColor = [System.Drawing.Color]::White
$lblTimelapseCalcHeader.Font      = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$lblTimelapseCalcHeader.TextAlign = 'MiddleCenter'
$lblTimelapseCalcHeader.Dock      = [System.Windows.Forms.DockStyle]::Fill
$cameraLayout.Controls.Add($lblTimelapseCalcHeader,0,6)
$cameraLayout.SetColumnSpan($lblTimelapseCalcHeader,2)
$lblCapturingDuration = New-Object System.Windows.Forms.Label
$lblCapturingDuration.Text      = "Capturing Duration:"
$lblCapturingDuration.ForeColor = [System.Drawing.Color]::White
$lblCapturingDuration.TextAlign = 'MiddleRight'
$lblCapturingDuration.Dock      = [System.Windows.Forms.DockStyle]::Fill
$cameraLayout.Controls.Add($lblCapturingDuration,0,7)
$tlpDuration = New-Object System.Windows.Forms.TableLayoutPanel
$tlpDuration.RowCount    = 2
$tlpDuration.ColumnCount = 2
$tlpDuration.Dock        = [System.Windows.Forms.DockStyle]::Fill
$tlpDuration.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,20)))
$tlpDuration.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)))
$tlpDuration.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50)))
$tlpDuration.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50)))
$lblHr = New-Object System.Windows.Forms.Label
$lblHr.Text      = "Hr"
$lblHr.TextAlign = 'MiddleCenter'
$lblHr.Dock      = [System.Windows.Forms.DockStyle]::Fill
$tlpDuration.Controls.Add($lblHr,0,0)
$lblMin = New-Object System.Windows.Forms.Label
$lblMin.Text      = "Min"
$lblMin.TextAlign = 'MiddleCenter'
$lblMin.Dock      = [System.Windows.Forms.DockStyle]::Fill
$tlpDuration.Controls.Add($lblMin,1,0)
$cmbCapturingDurationHours = New-Object System.Windows.Forms.ComboBox
$cmbCapturingDurationHours.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbCapturingDurationHours.Dock = [System.Windows.Forms.DockStyle]::Fill
for ($i=0; $i -le 500; $i++) { $null = $cmbCapturingDurationHours.Items.Add($i) }
$cmbCapturingDurationHours.SelectedItem = $global:CapturingDurationHours
$tlpDuration.Controls.Add($cmbCapturingDurationHours,0,1)
$cmbCapturingDurationMinutes = New-Object System.Windows.Forms.ComboBox
$cmbCapturingDurationMinutes.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbCapturingDurationMinutes.Dock = [System.Windows.Forms.DockStyle]::Fill
for ($i=0; $i -le 60; $i++) { $null = $cmbCapturingDurationMinutes.Items.Add($i) }
$cmbCapturingDurationMinutes.SelectedItem = $global:CapturingDurationMinutes
$tlpDuration.Controls.Add($cmbCapturingDurationMinutes,1,1)
$cameraLayout.Controls.Add($tlpDuration,1,7)
$global:cmbCapturingDurationHours = $cmbCapturingDurationHours
$global:cmbCapturingDurationMinutes = $cmbCapturingDurationMinutes
$cmbCapturingDurationHours.Add_SelectedIndexChanged({
    $global:CapturingDurationHours = [int]$cmbCapturingDurationHours.SelectedItem
    Save-Config
    Update-TimelapseCalculator
})
$cmbCapturingDurationMinutes.Add_SelectedIndexChanged({
    $global:CapturingDurationMinutes = [int]$cmbCapturingDurationMinutes.SelectedItem
    Save-Config
    Update-TimelapseCalculator
})
$lblTotalFrames = New-Object System.Windows.Forms.Label
$lblTotalFrames.Text = "Total Frames:"
$lblTotalFrames.ForeColor = [System.Drawing.Color]::White
$lblTotalFrames.TextAlign = 'MiddleRight'
$lblTotalFrames.Dock = [System.Windows.Forms.DockStyle]::Fill
$cameraLayout.Controls.Add($lblTotalFrames,0,8)
$txtTotalFrames = New-Object System.Windows.Forms.TextBox
$txtTotalFrames.ReadOnly = $true
$txtTotalFrames.Font     = $commonFont
$txtTotalFrames.Dock     = [System.Windows.Forms.DockStyle]::Fill
$txtTotalFrames.Text     = "0"
$txtTotalFrames.Enabled  = $false
$global:txtTotalFrames = $txtTotalFrames
$cameraLayout.Controls.Add($txtTotalFrames,1,8)
$lblFinalVideoDuration = New-Object System.Windows.Forms.Label
$lblFinalVideoDuration.Text      = "Estimated Video Length:"
$lblFinalVideoDuration.ForeColor = [System.Drawing.Color]::White
$lblFinalVideoDuration.TextAlign = 'MiddleRight'
$lblFinalVideoDuration.Dock      = [System.Windows.Forms.DockStyle]::Fill
$cameraLayout.Controls.Add($lblFinalVideoDuration,0,9)
$txtFinalVideoDuration = New-Object System.Windows.Forms.TextBox
$txtFinalVideoDuration.ReadOnly = $true
$txtFinalVideoDuration.Font     = $commonFont
$txtFinalVideoDuration.Dock     = [System.Windows.Forms.DockStyle]::Fill
$txtFinalVideoDuration.Text     = "0"
$txtFinalVideoDuration.Enabled  = $false
$global:txtFinalVideoDuration = $txtFinalVideoDuration
$cameraLayout.Controls.Add($txtFinalVideoDuration,1,9)
$lblEstimatedVideoSize = New-Object System.Windows.Forms.Label
$lblEstimatedVideoSize.Text = "Estimated Video Size:"
$lblEstimatedVideoSize.ForeColor = [System.Drawing.Color]::White
$lblEstimatedVideoSize.TextAlign = 'MiddleRight'
$lblEstimatedVideoSize.Dock = [System.Windows.Forms.DockStyle]::Fill
$cameraLayout.Controls.Add($lblEstimatedVideoSize,0,10)
$txtEstimatedVideoSize = New-Object System.Windows.Forms.TextBox
$txtEstimatedVideoSize.ReadOnly = $true
$txtEstimatedVideoSize.Font     = $commonFont
$txtEstimatedVideoSize.Dock     = [System.Windows.Forms.DockStyle]::Fill
$txtEstimatedVideoSize.Text     = "0"
$txtEstimatedVideoSize.Enabled  = $false
$global:txtEstimatedVideoSize = $txtEstimatedVideoSize
$cameraLayout.Controls.Add($txtEstimatedVideoSize,1,10)

###############################################################################
#                  GROUPBOX 3: Auto-Restart / Directory folders
###############################################################################
$gbAuto = New-Object System.Windows.Forms.GroupBox
$gbAuto.Text      = "Auto-Restart  /  Directory folders"
$gbAuto.ForeColor = [System.Drawing.Color]::White
$gbAuto.Font      = $commonFont
$gbAuto.Dock      = [System.Windows.Forms.DockStyle]::Fill
$gbAuto.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#1E1E1E")
$threePanel.Controls.Add($gbAuto,2,0)
$autoLayout = New-Object System.Windows.Forms.TableLayoutPanel
$autoLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$autoLayout.RowCount = 12
$autoLayout.ColumnCount = 2
$autoLayout.Padding = '5,5,5,5'
$autoLayout.RowStyles.Clear()
$autoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)))
$autoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,70)))
$autoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)))
$autoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)))
$autoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)))
$autoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)))
$autoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,80)))
$autoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)))
$autoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)))
$autoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)))
$autoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
$autoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,40)))
for ($c=0; $c -lt 2; $c++) {
    $autoLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50)))
}
$gbAuto.Controls.Add($autoLayout)
$chkAutoRestart = New-Object System.Windows.Forms.CheckBox
$chkAutoRestart.Text = "Auto Restart"
$chkAutoRestart.ForeColor = [System.Drawing.Color]::White
$chkAutoRestart.Font = $commonFont
$chkAutoRestart.AutoSize = $true
$chkAutoRestart.Checked = $global:AutoRestart
$chkAutoRestart.Add_CheckedChanged({
    $global:AutoRestart = $chkAutoRestart.Checked
    Write-Log "AutoRestart set to: $global:AutoRestart" "Info"
    if ($global:AutoRestart) { Set-TaskState -TaskName $StartupTaskName -Enable $true }
    else { Set-TaskState -TaskName $StartupTaskName -Enable $false }
    $cmbAutoMinutes.Enabled = $chkAutoRestart.Checked
    $cmbAutoHours.Enabled = $chkAutoRestart.Checked
    $cmbAutoDays.Enabled = $chkAutoRestart.Checked
    $cmbAutoMergeVideos.Enabled = $chkAutoRestart.Checked
    Save-Config
}) | Out-Null
$autoLayout.Controls.Add($chkAutoRestart,0,0)
$autoLayout.SetColumnSpan($chkAutoRestart,2)
$lblAutoNote = New-Object System.Windows.Forms.Label
$lblAutoNote.Text = 'Note: If "Auto-Restart" is checked on, Minutes, Hours, Days Until Restart and Auto-Merge Video are Enabled. If it''s unchecked, they are Disabled... And/Or Set to Zero they are also Disabled.'
$lblAutoNote.ForeColor = [System.Drawing.Color]::White
$lblAutoNote.Font = $commonFont
$lblAutoNote.TextAlign = 'MiddleCenter'
$lblAutoNote.Dock = [System.Windows.Forms.DockStyle]::Fill
$autoLayout.Controls.Add($lblAutoNote,0,1)
$autoLayout.SetColumnSpan($lblAutoNote,2)
$lblAutoMin = New-Object System.Windows.Forms.Label
$lblAutoMin.Text = "Minutes Until Restart:"
$lblAutoMin.ForeColor = [System.Drawing.Color]::White
$lblAutoMin.Font = $commonFont
$lblAutoMin.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblAutoMin.TextAlign = "MiddleRight"
$autoLayout.Controls.Add($lblAutoMin,0,2)
$cmbAutoMinutes = New-Object System.Windows.Forms.ComboBox
$cmbAutoMinutes.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbAutoMinutes.Dock = [System.Windows.Forms.DockStyle]::Fill
for ($i=0; $i -le 60; $i++) { $null = $cmbAutoMinutes.Items.Add($i) }
$cmbAutoMinutes.SelectedItem = $global:AutoRestartMinutes
$autoLayout.Controls.Add($cmbAutoMinutes,1,2)
$cmbAutoMinutes.Add_SelectedIndexChanged({
    $global:AutoRestartMinutes = [int]$cmbAutoMinutes.SelectedItem
    Write-Log "AutoRestartMinutes updated to: $global:AutoRestartMinutes" "Info"
    Save-Config
    Update-AutoCloseTimer
}) | Out-Null
$lblHours = New-Object System.Windows.Forms.Label
$lblHours.Text = "Hours Until Restart:"
$lblHours.ForeColor = [System.Drawing.Color]::White
$lblHours.Font = $commonFont
$lblHours.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblHours.TextAlign = "MiddleRight"
$autoLayout.Controls.Add($lblHours,0,3)
$cmbAutoHours = New-Object System.Windows.Forms.ComboBox
$cmbAutoHours.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbAutoHours.Dock = [System.Windows.Forms.DockStyle]::Fill
for ($i=0; $i -le 24; $i++) { $null = $cmbAutoHours.Items.Add($i) }
$cmbAutoHours.SelectedItem = $global:AutoRestartHours
$autoLayout.Controls.Add($cmbAutoHours,1,3)
$cmbAutoHours.Add_SelectedIndexChanged({
    $global:AutoRestartHours = [int]$cmbAutoHours.SelectedItem
    Write-Log "AutoRestartHours updated to: $global:AutoRestartHours" "Info"
    Save-Config
    Update-AutoCloseTimer
}) | Out-Null
$lblDays = New-Object System.Windows.Forms.Label
$lblDays.Text = "Days Until Restart:"
$lblDays.ForeColor = [System.Drawing.Color]::White
$lblDays.Font = $commonFont
$lblDays.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblDays.TextAlign = "MiddleRight"
$autoLayout.Controls.Add($lblDays,0,4)
$cmbAutoDays = New-Object System.Windows.Forms.ComboBox
$cmbAutoDays.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbAutoDays.Dock = [System.Windows.Forms.DockStyle]::Fill
for ($i=0; $i -le 24; $i++) { $null = $cmbAutoDays.Items.Add($i) }
$cmbAutoDays.SelectedItem = $global:AutoRestartDays
$autoLayout.Controls.Add($cmbAutoDays,1,4)
$cmbAutoDays.Add_SelectedIndexChanged({
    $global:AutoRestartDays = [int]$cmbAutoDays.SelectedItem
    Write-Log "AutoRestartDays updated to: $global:AutoRestartDays" "Info"
    Save-Config
    Update-AutoCloseTimer
}) | Out-Null
$lblMergeLimit = New-Object System.Windows.Forms.Label
$lblMergeLimit.Text = "Auto-Merge Video Limit:"
$lblMergeLimit.ForeColor = [System.Drawing.Color]::White
$lblMergeLimit.Font = $commonFont
$lblMergeLimit.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblMergeLimit.TextAlign = "MiddleRight"
$autoLayout.Controls.Add($lblMergeLimit,0,5)
$cmbAutoMergeVideos = New-Object System.Windows.Forms.ComboBox
$cmbAutoMergeVideos.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbAutoMergeVideos.Font = $commonFont
$cmbAutoMergeVideos.Dock = [System.Windows.Forms.DockStyle]::Fill
for ($i=0; $i -le 60; $i++) { $null = $cmbAutoMergeVideos.Items.Add($i) }
$cmbAutoMergeVideos.SelectedItem = $global:AutoMergeVideoLimit
$autoLayout.Controls.Add($cmbAutoMergeVideos,1,5)
$cmbAutoMergeVideos.Add_SelectedIndexChanged({
    $global:AutoMergeVideoLimit = [int]$cmbAutoMergeVideos.SelectedItem
    Write-Log "AutoMergeVideoLimit updated to: $global:AutoMergeVideoLimit" "Info"
    Save-Config
}) | Out-Null
$folderButtonsPanel = New-Object System.Windows.Forms.TableLayoutPanel
$folderButtonsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$folderButtonsPanel.RowCount = 2
$folderButtonsPanel.ColumnCount = 2
$folderButtonsPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,50)))
$folderButtonsPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,50)))
$folderButtonsPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50)))
$folderButtonsPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50)))
$btnSelectImageFolder = New-Object System.Windows.Forms.Button
$btnSelectImageFolder.Text = "Set Image Folder"
$btnSelectImageFolder.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$btnSelectImageFolder.ForeColor = [System.Drawing.Color]::White
$btnSelectImageFolder.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#3C3C3C")
$btnSelectImageFolder.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSelectImageFolder.Dock = [System.Windows.Forms.DockStyle]::Fill
$folderButtonsPanel.Controls.Add($btnSelectImageFolder,0,0)
$btnSelectVideoFolder = New-Object System.Windows.Forms.Button
$btnSelectVideoFolder.Text = "Set Video Folder"
$btnSelectVideoFolder.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$btnSelectVideoFolder.ForeColor = [System.Drawing.Color]::White
$btnSelectVideoFolder.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#3C3C3C")
$btnSelectVideoFolder.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSelectVideoFolder.Dock = [System.Windows.Forms.DockStyle]::Fill
$folderButtonsPanel.Controls.Add($btnSelectVideoFolder,1,0)
$btnOpenImageFolder = New-Object System.Windows.Forms.Button
$btnOpenImageFolder.Text = "Open Image Folder"
$btnOpenImageFolder.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$btnOpenImageFolder.ForeColor = [System.Drawing.Color]::White
$btnOpenImageFolder.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#3C3C3C")
$btnOpenImageFolder.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnOpenImageFolder.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnOpenImageFolder.Add_Click({
    if (Test-Path $global:ImageFolder) { Start-Process explorer.exe -ArgumentList $global:ImageFolder }
})
$folderButtonsPanel.Controls.Add($btnOpenImageFolder,0,1)
$btnOpenVideoFolder = New-Object System.Windows.Forms.Button
$btnOpenVideoFolder.Text = "Open Video Folder"
$btnOpenVideoFolder.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$btnOpenVideoFolder.ForeColor = [System.Drawing.Color]::White
$btnOpenVideoFolder.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#3C3C3C")
$btnOpenVideoFolder.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnOpenVideoFolder.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnOpenVideoFolder.Add_Click({
    if (Test-Path $global:VideoFolder) { Start-Process explorer.exe -ArgumentList $global:VideoFolder }
})
$folderButtonsPanel.Controls.Add($btnOpenVideoFolder,1,1)
$autoLayout.Controls.Add($folderButtonsPanel,0,6)
$autoLayout.SetColumnSpan($folderButtonsPanel,2)
$lblImageFolder = New-Object System.Windows.Forms.Label
$lblImageFolder.Text = "Image Folder: $global:ImageFolder"
$lblImageFolder.ForeColor = [System.Drawing.Color]::White
$lblImageFolder.Font = $commonFont
$lblImageFolder.Dock = [System.Windows.Forms.DockStyle]::Fill
$autoLayout.Controls.Add($lblImageFolder,0,7)
$autoLayout.SetColumnSpan($lblImageFolder,2)
$lblVideoFolder = New-Object System.Windows.Forms.Label
$lblVideoFolder.Text = "Video Folder: $global:VideoFolder"
$lblVideoFolder.ForeColor = [System.Drawing.Color]::White
$lblVideoFolder.Font = $commonFont
$lblVideoFolder.Dock = [System.Windows.Forms.DockStyle]::Fill
$autoLayout.Controls.Add($lblVideoFolder,0,8)
$autoLayout.SetColumnSpan($lblVideoFolder,2)
$chkEnableTimestamp = New-Object System.Windows.Forms.CheckBox
$chkEnableTimestamp.Text = "Enable Timestamp"
$chkEnableTimestamp.ForeColor = [System.Drawing.Color]::White
$chkEnableTimestamp.Font = $commonFont
$chkEnableTimestamp.AutoSize = $true
$chkEnableTimestamp.Checked = $global:EnableTimestamp
$chkEnableTimestamp.Add_CheckedChanged({
    $global:EnableTimestamp = $chkEnableTimestamp.Checked
    Write-Log "EnableTimestamp set to: $global:EnableTimestamp" "Info"
    Save-Config
})
$autoLayout.Controls.Add($chkEnableTimestamp,0,9)
$autoLayout.SetColumnSpan($chkEnableTimestamp,2)
$btnFullscreenPreview = New-Object System.Windows.Forms.Button
$btnFullscreenPreview.Text = "Fullscreen Preview"
$btnFullscreenPreview.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$btnFullscreenPreview.ForeColor = [System.Drawing.Color]::White
$btnFullscreenPreview.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#3C3C3C")
$btnFullscreenPreview.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnFullscreenPreview.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnFullscreenPreview.Add_Click({ Show-FullscreenPreview })
$autoLayout.Controls.Add($btnFullscreenPreview,0,11)
$autoLayout.SetColumnSpan($btnFullscreenPreview,2)

###############################################################################
#                            EVENT HANDLERS: Folders                           #
###############################################################################
function Select-Folder([string]$Title) {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = $Title
    $folderDialog.ShowNewFolderButton = $true
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $folderDialog.SelectedPath }
    return $null
}

$btnSelectImageFolder.Add_Click({
    $selectedPath = Select-Folder -Title "Select Image Folder"
    if ($selectedPath) {
        $global:ImageFolder = $selectedPath
        Write-Log "Image folder set to: $global:ImageFolder" "Info"
        Save-Config
        if (-not (Test-Path $global:ImageFolder)) {
            Write-Log "Creating new ImageFolder: $global:ImageFolder" "Info"
            try { New-Item -ItemType Directory -Path $global:ImageFolder | Out-Null }
            catch { Write-Log "Failed to create ImageFolder $global:ImageFolder: $_" "Error" }
        }
        SafeInvoke $lblImageFolder { $lblImageFolder.Text = "Image Folder: $global:ImageFolder" }
    }
}) | Out-Null

$btnSelectVideoFolder.Add_Click({
    $selectedPath = Select-Folder -Title "Select Video Folder"
    if ($selectedPath) {
        $global:VideoFolder = $selectedPath
        Write-Log "Video folder set to: $global:VideoFolder" "Info"
        Save-Config
        if (-not (Test-Path $global:VideoFolder)) {
            Write-Log "Creating new VideoFolder: $global:VideoFolder" "Info"
            try { New-Item -ItemType Directory -Path $global:VideoFolder | Out-Null }
            catch { Write-Log "Failed to create VideoFolder $global:VideoFolder: $_" "Error" }
        }
        SafeInvoke $lblVideoFolder { $lblVideoFolder.Text = "Video Folder: $global:VideoFolder" }
    }
}) | Out-Null

###############################################################################
#                        EVENT HANDLERS: Camera Panel                        #
###############################################################################
function InitializeSelectCamera {
    $cmbSelectCamera.Items.Clear()
    foreach ($device in $devices) { $cmbSelectCamera.Items.Add($device.Name) | Out-Null }
    if ($cmbSelectCamera.Items.Count -gt 0) {
        if ($global:SelectedCameraIndex -lt 0) { $global:SelectedCameraIndex = 0 }
        $cmbSelectCamera.SelectedIndex = $global:SelectedCameraIndex
        SwitchCamera -DeviceIndex $global:SelectedCameraIndex
    }
}

###############################################################################
#                        EVENT HANDLERS: Capture Buttons                        #
###############################################################################
$btnStart.Add_Click({
    if ($global:captureTimer -and -not $global:captureTimer.Enabled) {
        if (-not $videoPlayer.IsRunning -and $videoPlayer.VideoSource) { $videoPlayer.Start() }
        # Since the interval controls are now combo boxes, their values come from SelectedItem.
        $secondsVal = [int]$cmbSec.SelectedItem
        $minutesVal = [int]$cmbMin.SelectedItem
        $total = ($minutesVal * 60) + $secondsVal
        if ($total -le 0) { $total = 5 }
        Save-Config
        try {
            [int]$global:imgCounter = (GetNextImageNumber -ImgFolder $global:ImageFolder)
            Write-Log "Dynamically set image counter -> $global:imgCounter" "Info"
        }
        catch {
            Write-Log "Error setting image counter at start: $_" "Error"
            [int]$global:imgCounter = 1
        }
        $global:sessionImageCount = 0
        Set-StatusText "Capture Started. Capturing Images: 0"
        $global:captureTimer.Interval = $total * 1000
        Write-Log "Capture interval set to $total second(s)." "Info"
        $global:captureTimer.Start()
        SafeInvoke $recordingLabelBottom { 
            $recordingLabelBottom.Text = "Recording: ON"
            $recordingLabelBottom.ForeColor = [System.Drawing.Color]::Lime
        }
        Write-Log "Start Capture button pressed. Capture started." "Info"
        Write-Host ("Capturing Images: 0`r") -ForegroundColor Green -NoNewline
    }
    else {
        Write-Log "Capture timer is already running." "Warning"
    }
}) | Out-Null

$btnStop.Add_Click({
    if ($global:captureTimer -and $global:captureTimer.Enabled) {
        Write-Log "Stopping capture..." "Info"
        $global:captureTimer.Stop()
        Write-Host "`nCaptured Images: $global:sessionImageCount" -ForegroundColor Green
        Set-StatusText "Capture Stopped. Total Images Captured: $global:sessionImageCount"
        SafeInvoke $recordingLabelBottom { 
            $recordingLabelBottom.Text = "Recording: OFF"
            $recordingLabelBottom.ForeColor = [System.Drawing.Color]::Red
        }
        Write-Log "Stop Capture button pressed. Capture stopped." "Info"
        Set-StatusText "Status: Merging..."
        MergeImagesIntoVideo -Folder $global:ImageFolder -VidFolder $global:VideoFolder `
            -FPS $global:OutputFps -CrfVal $CrfValue -PresetVal $PresetVal `
            -DarkThr $global:DarkThreshold -OutW $global:SelectedWidth -OutH $global:SelectedHeight
        try {
            [int]$global:imgCounter = (GetNextImageNumber -ImgFolder $global:ImageFolder)
            Write-Log "Reset image counter to $global:imgCounter after stop merge." "Info"
        }
        catch {
            Write-Log "Error recalculating image counter after stop merge: $_" "Error"
            [int]$global:imgCounter = 1
        }
        Set-StatusText "Status: Preview Only"
    }
    else {
        Write-Log "Capture timer is not running." "Warning"
    }
}) | Out-Null

$btnForce.Add_Click({
    Write-Log "Forcing video creation from leftover images..." "Info"
    Set-StatusText "Status: Force Merge in progress..."
    MergeImagesIntoVideo -Folder $global:ImageFolder -VidFolder $global:VideoFolder `
        -FPS $global:OutputFps -CrfVal $CrfValue -PresetVal $PresetVal `
        -DarkThr $global:DarkThreshold -OutW $global:SelectedWidth -OutH $global:SelectedHeight
    Set-StatusText "Status: Preview Only"
}) | Out-Null

$btnMergeAllVideos.Add_Click({
    Write-Log "Initiating master video merge..." "Info"
    Set-StatusText "Status: Merging all videos into master video..."
    $masterVideo = Join-Path $global:VideoFolder "MasterLapse.mp4"
    MergeAllVideos -VidFolder $global:VideoFolder -MasterVideoPath $masterVideo
}) | Out-Null

$form.Add_Shown({
    Write-Log "Initializing camera selection..." "Info"
    $cmbSelectCamera.Items.Clear()
    InitializeSelectCamera
    Write-Log "Camera preview started." "Info"
    $statusLabel.Text = "Status: Camera preview active."
    Update-TimelapseCalculator
    $cmbAutoMinutes.Enabled = $chkAutoRestart.Checked
    $cmbAutoHours.Enabled = $chkAutoRestart.Checked
    $cmbAutoDays.Enabled = $chkAutoRestart.Checked
    $cmbAutoMergeVideos.Enabled = $chkAutoRestart.Checked
    Update-AutoCloseTimer
    $global:isInitializing = $false
    if ($global:AutoRestart) {
        Write-Log "AutoRestart is enabled. Setting auto start capture timer for 10 seconds." "Info"
        $global:autoStartTimer = New-Object System.Windows.Forms.Timer
        $global:autoStartTimer.Interval = 10000
        $global:autoStartTimer.Add_Tick({
            $global:autoStartTimer.Stop()
            $global:autoStartTimer.Dispose()
            $global:autoStartTimer = $null
            if ($global:AutoRestart -and (-not $global:captureTimer.Enabled)) {
                $btnStart.PerformClick()
                Write-Log "Auto Start Capture triggered after AutoRestart." "Info"
                Set-StatusText "Status: Auto-Recording Started."
            }
        })
        $global:autoStartTimer.Start()
    }
    else {
        Write-Log "AutoRestart is not enabled. No auto start timer will be set." "Info"
    }
    $form.Activate()
}) | Out-Null

###############################################################################
#                NEW: FORM CLOSED EVENT FOR AUTO-RESTART TRIGGER
###############################################################################
$form.Add_FormClosed({
    $global:closingNow = $true
    if ($global:AutoRestart) {
         Write-Log "AutoRestart enabled. Enabling admin scheduled task and triggering restart." "Info"
         try {
             Set-TaskState -TaskName $AdminTaskName -Enable $true
             Start-Sleep -Seconds 2
             schtasks /run /tn $AdminTaskName | Out-Null
         }
         catch {
             Write-Log "Error triggering admin scheduled task for auto restart: $_" "Error"
         }
         [System.Environment]::Exit(0)
    }
})

[void][System.Windows.Forms.Application]::Run($form)

###############################################################################
# CLEANUP: RELEASE CAMERA RESOURCES
###############################################################################
if ($global:videoSrc) {
    if ($global:videoSrc.IsRunning) {
        Write-Log "Stopping video source during cleanup." "Info"
        $global:videoSrc.SignalToStop()
        $global:videoSrc.WaitForStop()
    }
    try {
        if ($global:videoSrc -is [System.IDisposable]) {
            Write-Log "Disposing video source during cleanup." "Info"
            $global:videoSrc.Dispose()
        }
    }
    catch {
        Write-Log "Error disposing video source: $_" "Error"
    }
}
if ($videoPlayer -and $videoPlayer.IsRunning) {
    Write-Log "Stopping video preview during cleanup." "Info"
    try { $videoPlayer.Stop() }
    catch {
         if ($_.Exception.Message -notmatch "Thread abort is not supported") {
              Write-Log "Error stopping videoPlayer during cleanup: $_" "Warning"
         }
    }
}
Write-Log "Script exiting." "Info"