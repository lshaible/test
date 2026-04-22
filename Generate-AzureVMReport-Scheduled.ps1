# Generate-AzureVMReport-Scheduled.ps1
# This script can be scheduled as a Windows Task to generate reports regularly

param(
    [string]$SubscriptionId,
    [string]$OutputDirectory = "C:\Reports",
    [string]$LogFile = "C:\Logs\AzureVMReport.log"
)

# Create directories if they don't exist
if (-not (Test-Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

if (-not (Test-Path (Split-Path $LogFile))) {
    New-Item -ItemType Directory -Path (Split-Path $LogFile) -Force | Out-Null
}

# Start logging
Start-Transcript -Path $LogFile -Append

try {
    Write-Host "===== Azure VM Report Generation Started =====" -ForegroundColor Cyan
    Write-Host "Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
    
    # Build output path
    $filename = "Azure_Windows_VM_Licensing_vCPU_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    $outputPath = Join-Path $OutputDirectory $filename
    
    # Run the main report generation script
    $scriptPath = Join-Path (Split-Path $MyInvocation.MyCommand.Path) "Generate-AzureVMReport.ps1"
    
    if (Test-Path $scriptPath) {
        & $scriptPath -SubscriptionId $SubscriptionId -OutputPath $outputPath
        Write-Host "Report saved to: $outputPath" -ForegroundColor Green
    }
    else {
        Write-Error "Generate-AzureVMReport.ps1 not found at: $scriptPath"
    }
}
catch {
    Write-Error "Error during report generation: $_"
}
finally {
    Write-Host "===== Report Generation Completed =====" -ForegroundColor Cyan
    Write-Host "Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
    Stop-Transcript
}
