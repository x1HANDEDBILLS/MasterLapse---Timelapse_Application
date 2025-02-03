###############################################################################
# TimeLapse Task Setup Script
# This script:
#   1. Ensures it is running in Windows PowerShell 5.1
#   2. Prompts for Administrator privileges if needed
#   3. Installs FFmpeg if not already installed
#   4. Installs PowerShell 7 if not already installed
#   5. Installs AForge.NET Framework (2.2.5) if not already installed
#   6. Creates and configures two scheduled tasks:
#       - RunTimelapseStartupScript (disabled by default)
#       - RunTimelapseAdminScript (enabled when running)
###############################################################################

# Function to display messages with timestamps
function Show-Message {
    param (
        [string]$Message, 
        [ConsoleColor]$Color = "Gray"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "$timestamp - $Message" -ForegroundColor $Color
}

# Ensure the script is running in Windows PowerShell 5.1
if ($PSVersionTable.PSVersion.Major -ge 7) {
    Show-Message "Script detected running in PowerShell 7. Restarting in Windows PowerShell 5.1..." "Yellow"
    Write-Host "Restarting in Windows PowerShell 5.1..." -ForegroundColor Yellow
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Show-Message "Script started in Windows PowerShell 5.1."

# Check for Administrator Privileges
$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$Principal = New-Object Security.Principal.WindowsPrincipal($CurrentUser)
$IsAdmin = $Principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Show-Message "Not running as Administrator. Attempting to relaunch with admin privileges..." "Yellow"
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Show-Message "Script is running as Administrator." "Green"
Write-Host "Script is running as Administrator!" -ForegroundColor Green

###############################################################################
# 1. Install FFmpeg using winget if not already installed
###############################################################################
function Is-ProgramInstalled {
    param (
        [string]$Name
    )
    $programs = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, `
                                  HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* `
                                  -ErrorAction SilentlyContinue
    return $programs.DisplayName -like "*$Name*"
}

if (-not (Is-ProgramInstalled -Name "FFmpeg")) {
    Show-Message "Installing FFmpeg..." "Yellow"
    Write-Host "Executing: winget install ffmpeg" -ForegroundColor Yellow
    try {
        Start-Process -FilePath "winget" -ArgumentList "install ffmpeg --silent" -NoNewWindow -Wait -ErrorAction Stop
        Show-Message "FFmpeg installation completed successfully." "Green"
    } catch {
        Show-Message "FFmpeg installation failed. Error: $($_.Exception.Message)" "Red"
        Write-Host "FFmpeg installation failed. Error: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
} else {
    Show-Message "FFmpeg is already installed. Skipping installation." "Green"
    Write-Host "FFmpeg is already installed. Skipping installation." -ForegroundColor Green
}

###############################################################################
# 2. Check and install PowerShell 7 if not already installed
###############################################################################
$PowerShell7 = "C:\Program Files\PowerShell\7\pwsh.exe"

if (-not (Test-Path $PowerShell7)) {
    Show-Message "Installing PowerShell 7..." "Yellow"
    Write-Host "Executing: winget install Microsoft.Powershell --silent" -ForegroundColor Yellow
    try {
        Start-Process -FilePath "winget" -ArgumentList "install Microsoft.Powershell --silent" -NoNewWindow -Wait -ErrorAction Stop
        Show-Message "PowerShell 7 installation completed successfully." "Green"
    } catch {
        Show-Message "PowerShell 7 installation failed. Error: $($_.Exception.Message)" "Red"
        Write-Host "PowerShell 7 installation failed. Error: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
} else {
    Show-Message "PowerShell 7 is already installed." "Green"
    Write-Host "PowerShell 7 is already installed." -ForegroundColor Green
}

###############################################################################
# 3. Locate the TimeLapse Script (TimeLapse-v1.0.0.ps1)
#    --- UPDATED to search multiple paths including Program Files (x86) ---
###############################################################################

# A list of possible directories where TimeLapse-v1.0.0.ps1 might be installed:
$PossiblePaths = @(
    "C:\Program Files\TimeLapse",
    "C:\Program Files (x86)\TimeLapse",
    "C:\Users\$env:USERNAME\Desktop"
)

Show-Message "Searching for TimeLapse-v1.0.0.ps1 in known possible paths..." "Gray"
$TargetScript = $null
$ScriptName   = "TimeLapse-v1.0.0.ps1"

foreach ($Path in $PossiblePaths) {
    Show-Message "Checking path: $Path" "Gray"
    if (Test-Path $Path) {
        # Search (with -Filter) to find TimeLapse-v1.0.0.ps1
        $foundScripts = Get-ChildItem -Path $Path -Filter $ScriptName -File -Recurse -ErrorAction SilentlyContinue
        if ($foundScripts) {
            # If multiple scripts found, select the newest
            $TargetScript = $foundScripts | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            Show-Message "Found $ScriptName at $($TargetScript.FullName)" "Gray"
            break
        }
    }
}

if (-not $TargetScript) {
    # Handle the error if the file isn't found
    Show-Message "ERROR: $ScriptName not found in any of the known paths. Exiting." "Red"
    Write-Host "ERROR: $ScriptName not found. Please ensure it is placed in one of the defined paths." -ForegroundColor Red
    exit 1
}

# Store the full path to the TimeLapse script for the rest of the code:
$TargetScriptPath = $TargetScript.FullName
Show-Message "Using TimeLapse script at: $TargetScriptPath" "Gray"

###############################################################################
# 4. Install AForge.NET Framework (2.2.5) if not already installed
###############################################################################
function Is-AForgeInstalled {
    $aForgeInstalled = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, `
                                     HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* `
                                     -ErrorAction SilentlyContinue `
                      | Where-Object { $_.DisplayName -like "*AForge.NET Framework*" }
    return $aForgeInstalled -ne $null
}

if (-not (Is-AForgeInstalled)) {
    $AForgeInstallerURL = "https://storage.googleapis.com/google-code-archive-downloads/v2/code.google.com/aforge/AForge.NET%20Framework-2.2.5.exe"
    $InstallerPath = "C:\Temp\AForge.NET_Framework_2.2.5.exe"

    # Ensure the Temp directory exists
    if (-not (Test-Path "C:\Temp")) {
        New-Item -Path "C:\" -Name "Temp" -ItemType Directory -Force | Out-Null
    }

    # Remove existing installer if present
    if (Test-Path $InstallerPath) {
        try {
            Remove-Item -Path $InstallerPath -Force -ErrorAction Stop
            Show-Message "Existing AForge.NET installer removed." "Gray"
        } catch {
            Show-Message "Cannot remove existing installer. It might be in use. Error: $($_.Exception.Message)" "Red"
            Write-Host "Cannot remove existing installer. It might be in use. Please ensure no other installations are running." -ForegroundColor Red
            exit 1
        }
    }

    Show-Message "Attempting to download AForge.NET Framework installer from $AForgeInstallerURL..." "Yellow"
    try {
        Write-Host "Downloading AForge.NET Framework installer... Please wait." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $AForgeInstallerURL -OutFile $InstallerPath -UseBasicParsing -ErrorAction Stop

        Show-Message "AForge.NET Framework installer downloaded successfully to $InstallerPath." "Green"
        Write-Host "Download completed successfully! The file is saved at $InstallerPath." -ForegroundColor Green

        # Verify if the file is accessible
        if (-not (Test-Path $InstallerPath)) {
            throw "Installer file does not exist after download."
        }

        # Automatically run the installer
        Show-Message "Launching AForge.NET Framework installer..." "Yellow"
        Start-Process -FilePath $InstallerPath -Wait -ErrorAction Stop
        Show-Message "AForge.NET Framework installation process launched successfully." "Gray"
        Write-Host "The installation process for AForge.NET Framework has started. Please complete the setup." -ForegroundColor Yellow
    } catch {
        Show-Message "Download or installation failed. Error: $($_.Exception.Message)" "Red"
        Write-Host "Download or installation failed. Error: $($_.Exception.Message)" -ForegroundColor Red

        # Fallback: Open the AForge.NET download page in a browser
        $AForgeDownloadPage = "https://code.google.com/archive/p/aforge/downloads"
        Show-Message "Opening fallback download page: $AForgeDownloadPage" "Yellow"
        Start-Process $AForgeDownloadPage
        Write-Host "Please manually download version 2.2.5 from the AForge.NET Downloads page." -ForegroundColor Yellow
    }
} else {
    Show-Message "AForge.NET Framework is already installed. Skipping installation." "Green"
    Write-Host "AForge.NET Framework is already installed. Skipping installation." -ForegroundColor Green
}

Show-Message "AForge.NET Framework installation process handled." "Gray"

###############################################################################
# 5. Task Scheduler Setup
###############################################################################
Show-Message "Starting Task Scheduler setup..." "Gray"

# Retrieve the current user's SID
$UserSID = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value

# Task Names
$TaskNameStartup = "RunTimelapseStartupScript"
$TaskNameAdmin   = "RunTimelapseAdminScript"

# Create XML Configuration Function with Specific Settings
function CreateTaskXML {
    param (
        [string]$TaskName,
        [string]$ScriptPath,
        [string]$UserSID,
        [string]$PowerShellPath
    )

    switch ($TaskName) {
        "RunTimelapseStartupScript" {
$Settings = @"
            <Settings>
                <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
                <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
                <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
                <AllowHardTerminate>true</AllowHardTerminate>
                <StartWhenAvailable>false</StartWhenAvailable>
                <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
                <IdleSettings>
                    <StopOnIdleEnd>true</StopOnIdleEnd>
                    <RestartOnIdle>false</RestartOnIdle>
                </IdleSettings>
                <AllowStartOnDemand>true</AllowStartOnDemand>
                <Enabled>false</Enabled>
                <Hidden>false</Hidden>
                <RunOnlyIfIdle>false</RunOnlyIfIdle>
                <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
                <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
                <WakeToRun>false</WakeToRun>
                <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
                <Priority>7</Priority>
                <RestartOnFailure>
                    <Interval>PT1M</Interval>
                    <Count>3</Count>
                </RestartOnFailure>
            </Settings>
"@
$Triggers = @"
            <Triggers>
                <BootTrigger>
                    <Enabled>true</Enabled>
                    <Delay>PT1M</Delay>
                </BootTrigger>
            </Triggers>
"@
        }
        "RunTimelapseAdminScript" {
$Settings = @"
            <Settings>
                <MultipleInstancesPolicy>StopExisting</MultipleInstancesPolicy>
                <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
                <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
                <AllowHardTerminate>true</AllowHardTerminate>
                <StartWhenAvailable>false</StartWhenAvailable>
                <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
                <IdleSettings>
                    <StopOnIdleEnd>false</StopOnIdleEnd>
                    <RestartOnIdle>false</RestartOnIdle>
                </IdleSettings>
                <AllowStartOnDemand>true</AllowStartOnDemand>
                <Enabled>true</Enabled>
                <Hidden>false</Hidden>
                <RunOnlyIfIdle>false</RunOnlyIfIdle>
                <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
                <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
                <WakeToRun>false</WakeToRun>
                <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
                <Priority>7</Priority>
                <RestartOnFailure>
                    <Interval>PT1M</Interval>
                    <Count>3</Count>
                </RestartOnFailure>
            </Settings>
"@
$Triggers = @"
            <Triggers />
"@
        }
        default {
            Write-Host "Unknown TaskName: $TaskName" -ForegroundColor Red
            throw "Unknown TaskName: $TaskName"
        }
    }

    return @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Author>$env:USERNAME</Author>
    <Description>Runs the TimeLapse PowerShell script.</Description>
    <URI>\$TaskName</URI>
  </RegistrationInfo>
  $Triggers
  <Principals>
    <Principal id="Author">
      <UserId>$UserSID</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  $Settings
  <Actions Context="Author">
    <Exec>
      <Command>"$PowerShellPath"</Command>
      <Arguments>-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`"</Arguments>
    </Exec>
  </Actions>
</Task>
"@
}

Show-Message "Creating XML for Startup Task: $TaskNameStartup" "Gray"
$StartupTaskXML = CreateTaskXML -TaskName $TaskNameStartup -ScriptPath $TargetScriptPath -UserSID $UserSID -PowerShellPath $PowerShell7
$StartupTaskXMLPath = "C:\Users\$env:USERNAME\Desktop\$TaskNameStartup.xml"
try {
    $StartupTaskXML | Out-File -FilePath $StartupTaskXMLPath -Encoding Unicode -ErrorAction Stop
    Show-Message "Startup Task XML created at $StartupTaskXMLPath" "Gray"
} catch {
    Show-Message "Failed to create Startup Task XML. Error: $($_.Exception.Message)" "Red"
    Write-Host "Failed to create Startup Task XML. Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Show-Message "Creating XML for Admin Task: $TaskNameAdmin" "Gray"
$AdminTaskXML = CreateTaskXML -TaskName $TaskNameAdmin -ScriptPath $TargetScriptPath -UserSID $UserSID -PowerShellPath $PowerShell7
$AdminTaskXMLPath = "C:\Users\$env:USERNAME\Desktop\$TaskNameAdmin.xml"
try {
    $AdminTaskXML | Out-File -FilePath $AdminTaskXMLPath -Encoding Unicode -ErrorAction Stop
    Show-Message "Admin Task XML created at $AdminTaskXMLPath" "Gray"
} catch {
    Show-Message "Failed to create Admin Task XML. Error: $($_.Exception.Message)" "Red"
    Write-Host "Failed to create Admin Task XML. Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Show-Message "Creating Scheduled Tasks using schtasks..." "Gray"

function CreateAndModifyTask {
    param (
        [string]$TaskName,
        [string]$XMLPath,
        [string]$ModifyAction  # "Enable" or "Disable"
    )

    Write-Host "Creating Task: $TaskName" -ForegroundColor Yellow
    $createResult = schtasks /Create /TN $TaskName /XML $XMLPath /F 2>&1
    if ($LASTEXITCODE -eq 0) {
        Show-Message "Scheduled Task '$TaskName' created successfully." "Green"
    } else {
        Show-Message "ERROR: Failed to create Scheduled Task '$TaskName'. Details: $createResult" "Red"
        Write-Host "ERROR: Failed to create Scheduled Task '$TaskName'. Check console for details." -ForegroundColor Red
        return
    }

    if ($ModifyAction -eq "Disable") {
        Write-Host "Disabling Task: $TaskName" -ForegroundColor Yellow
        $modifyResult = schtasks /Change /TN $TaskName /Disable 2>&1
    } elseif ($ModifyAction -eq "Enable") {
        Write-Host "Enabling Task: $TaskName" -ForegroundColor Yellow
        $modifyResult = schtasks /Change /TN $TaskName /Enable 2>&1
    } else {
        Show-Message "Invalid ModifyAction: $ModifyAction" "Red"
        return
    }

    if ($LASTEXITCODE -eq 0) {
        Show-Message "Scheduled Task '$TaskName' $ModifyAction successfully." "Green"
    } else {
        Show-Message "ERROR: Failed to $ModifyAction Scheduled Task '$TaskName'. Details: $modifyResult" "Red"
        Write-Host "ERROR: Failed to $ModifyAction Scheduled Task '$TaskName'. Check console for details." -ForegroundColor Red
    }
}

# Create and Disable Startup Task
CreateAndModifyTask -TaskName $TaskNameStartup -XMLPath $StartupTaskXMLPath -ModifyAction "Disable"

# Create and Enable Admin Task
CreateAndModifyTask -TaskName $TaskNameAdmin -XMLPath $AdminTaskXMLPath -ModifyAction "Enable"

# Remove XML Files if they exist
if ( (Test-Path $StartupTaskXMLPath) -or (Test-Path $AdminTaskXMLPath) ) {
    Show-Message "Removing temporary XML files..." "Gray"
    try {
        Remove-Item -Path $StartupTaskXMLPath, $AdminTaskXMLPath -Force -ErrorAction Stop
        Show-Message "Temporary XML files removed." "Gray"
    } catch {
        Show-Message "Failed to remove temporary XML files. Error: $($_.Exception.Message)" "Red"
        Write-Host "Failed to remove temporary XML files. Error: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Show-Message "No temporary XML files to remove." "Gray"
}

Show-Message "Task Scheduler setup completed successfully." "Green"

###############################################################################
# 6. Final Completion Messages
###############################################################################
Write-Host "`nFFmpeg Installed, PowerShell 7 Installed, and Tasks Scheduled Successfully!" -ForegroundColor Green
Write-Host "Press any key to exit..." -ForegroundColor Cyan
[System.Console]::ReadKey($true) | Out-Null
