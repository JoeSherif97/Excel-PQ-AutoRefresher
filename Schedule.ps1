# === CONFIG ===
$ScriptPath = "C:\ScriptPS\Work\ExcelRefresh.ps1"

# === ACTION ===
$Action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`""

# === TRIGGER ===
# Start in 1 minute, repeat every 3 hours, for 1 year
$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
    -RepetitionInterval (New-TimeSpan -Hours 3) `
    -RepetitionDuration (New-TimeSpan -Days 365)

# === SETTINGS ===
$Settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -DontStopOnIdleEnd `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1)

# === REGISTER TASK ===
Register-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings `
    -TaskName "AutoRefreshExcelFiles" `
    -Description "Refresh Excel files with Power Query every 3 hours" `
    -User $env:USERNAME

# Optional: start the task immediately for testing
Start-ScheduledTask -TaskName "AutoRefreshExcelFiles"
