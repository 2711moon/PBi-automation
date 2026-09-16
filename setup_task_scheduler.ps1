# setup_task_scheduler.ps1
# Registers the Power BI Report Automation as a Windows Scheduled Task.
# Run this script as Administrator (right-click -> Run with PowerShell as Admin).
#
# To change the schedule time: edit SCHEDULE_HOUR/SCHEDULE_MINUTE in config.py
# then re-run this script.

param(
    [int]$Hour   = 9,
    [int]$Minute = 0
)

$TaskName   = "PowerBI_Report_Automation_Kisna"
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$BatchFile  = Join-Path $ScriptDir "run.bat"

if (-not (Test-Path $BatchFile)) {
    Write-Error "run.bat not found at: $BatchFile"
    exit 1
}

$Action = New-ScheduledTaskAction `
    -Execute "cmd.exe" `
    -Argument "/c `"$BatchFile`" >> `"$ScriptDir\automation.log`" 2>&1" `
    -WorkingDirectory $ScriptDir

$Trigger = New-ScheduledTaskTrigger -Daily -At "$($Hour):$('{0:D2}' -f $Minute)"

$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit  (New-TimeSpan -Hours 2) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -DontStopOnIdleEnd `
    -Hidden   # Runs with no visible window

# Register (or update if already exists)
Register-ScheduledTask `
    -TaskName $TaskName `
    -Action   $Action `
    -Trigger  $Trigger `
    -Settings $Settings `
    -RunLevel Highest `
    -Force    | Out-Null

Write-Host ""
Write-Host "Task registered: $TaskName"
Write-Host "Runs daily at: $($Hour):$('{0:D2}' -f $Minute)"
Write-Host ""
Write-Host "To change schedule: edit SCHEDULE_HOUR in config.py then re-run this script."
Write-Host "To view/edit task:  Task Scheduler -> Task Scheduler Library -> $TaskName"
Write-Host ""
