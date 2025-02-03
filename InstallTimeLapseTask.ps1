# Function to log messages
function Write-Log {
    param ([string]$Message)
}

# Path to the secondary script (CreateTimeLapseTasks.ps1)
$SecondaryScript = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "CreateTimeLapseTasks.ps1"

# Ensure the secondary script exists
if (-not (Test-Path $SecondaryScript)) {
    Write-Host "ERROR: Secondary script not found! Check the logs for details." -ForegroundColor Red
    exit 1
}

# Launch the secondary script in admin mode
Write-Host "Attempting to launch secondary script as Administrator: $SecondaryScript"
try {
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$SecondaryScript`"" -Verb RunAs -Wait
    Write-Host "Secondary script executed successfully!" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to execute the secondary script. Check logs for details." -ForegroundColor Red
    exit 1
}

# Completion Message
Write-Host "TimeLapse Task setup completed successfully!" -ForegroundColor Green
