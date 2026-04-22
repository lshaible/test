# Setup-ReportScheduledTask.ps1
# Creates a Windows Task Scheduler job to run the Azure VM report generation daily

param(
    [string]$TaskName = "Generate Azure VM Report",
    [string]$SubscriptionId,
    [string]$OutputDirectory = "C:\Reports",
    [string]$LogFile = "C:\Logs\AzureVMReport.log",
    [datetime]$DailyTime = (Get-Date -Hour 2 -Minute 0 -Second 0),
    [switch]$RunImmediately
)

# Requires admin privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script requires administrator privileges. Please run as Administrator."
    exit 1
}

try {
    # Get the script directory
    $scriptDir = Split-Path $MyInvocation.MyCommand.Path
    $reportScript = Join-Path $scriptDir "Generate-AzureVMReport-Scheduled.ps1"
    
    if (-not (Test-Path $reportScript)) {
        Write-Error "Generate-AzureVMReport-Scheduled.ps1 not found at: $reportScript"
        exit 1
    }
    
    Write-Host "Creating scheduled task: $TaskName" -ForegroundColor Cyan
    
    # Create task arguments
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$reportScript`""
    if ($SubscriptionId) {
        $arguments += " -SubscriptionId `"$SubscriptionId`""
    }
    $arguments += " -OutputDirectory `"$OutputDirectory`" -LogFile `"$LogFile`""
    
    # Create the task
    $trigger = New-ScheduledTaskTrigger -Daily -At $DailyTime
    $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument $arguments
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable
    $principal = New-ScheduledTaskPrincipal -UserId (whoami)
    
    # Remove existing task if it exists
    $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($existingTask) {
        Write-Host "Removing existing task: $TaskName" -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    }
    
    # Register the task
    Register-ScheduledTask -TaskName $TaskName `
        -Trigger $trigger `
        -Action $action `
        -Settings $settings `
        -Principal $principal `
        -Force
    
    Write-Host "Task created successfully!" -ForegroundColor Green
    Write-Host "Task Name: $TaskName" -ForegroundColor Cyan
    Write-Host "Scheduled Time: $($DailyTime.ToString('HH:mm:ss'))" -ForegroundColor Cyan
    Write-Host "Output Directory: $OutputDirectory" -ForegroundColor Cyan
    Write-Host "Log File: $LogFile" -ForegroundColor Cyan
    
    # Run immediately if requested
    if ($RunImmediately) {
        Write-Host "Starting task immediately..." -ForegroundColor Yellow
        Start-ScheduledTask -TaskName $TaskName
        Write-Host "Task started. Check the output directory for the report." -ForegroundColor Green
    }
    else {
        Write-Host "The task will run daily at $($DailyTime.ToString('HH:mm:ss'))" -ForegroundColor Green
    }
}
catch {
    Write-Error "Error creating scheduled task: $_"
    exit 1
}
