# Examples-AzureVMReport.ps1
# This file contains example usage for the Azure VM Report scripts

<#
EXAMPLE 0: Generate SQL Server VMs only (default behavior)
---
.\Generate-AzureVMReport.ps1

# By default, reports only include VMs with SQL Server installed
#>

<#
EXAMPLE 0B: Generate ALL VMs (not just SQL Server)
---
.\Generate-AzureVMReport.ps1 -SQLServerOnly:$false

# This includes all VMs in the report, regardless of SQL Server installation
#>

<#
EXAMPLE 1: Generate report for current subscription
---
.\Generate-AzureVMReport.ps1

Output: Creates a SQL Server VM report with a timestamp in the current directory
#>

<#
EXAMPLE 2: Generate report for specific subscription and custom output path
---
.\Generate-AzureVMReport.ps1 -SubscriptionId "12345678-1234-1234-1234-123456789012" -OutputPath "C:\Reports\MyReport.xlsx"

Output: Creates a SQL Server VM report at the specified path
#>

<#
EXAMPLE 3: Set up daily scheduled task (requires Administrator)
---
# First, open PowerShell as Administrator, then:
.\Setup-ReportScheduledTask.ps1 `
    -SubscriptionId "12345678-1234-1234-1234-123456789012" `
    -OutputDirectory "C:\Reports" `
    -DailyTime (Get-Date -Hour 2 -Minute 0 -Second 0) `
    -RunImmediately

Output: 
- Creates scheduled task to run daily at 2:00 AM
- Runs immediately and generates first report
- SQL Server VM reports saved to C:\Reports
- Logs saved to C:\Logs\AzureVMReport.log
#>

<#
EXAMPLE 4: Generate report and capture in variable for further processing
---
$reportPath = "C:\Reports\report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
.\Generate-AzureVMReport.ps1 -OutputPath $reportPath

if (Test-Path $reportPath) {
    Write-Host "Report successfully created at: $reportPath"
    # Open report in Excel (optional)
    & $reportPath
}
#>

<#
EXAMPLE 5: Generate reports for multiple subscriptions
---
$subscriptions = @(
    "sub-id-1",
    "sub-id-2",
    "sub-id-3"
)

foreach ($sub in $subscriptions) {
    Write-Host "Generating report for: $sub"
    $outPath = "C:\Reports\Report_$($sub)_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    .\Generate-AzureVMReport.ps1 -SubscriptionId $sub -OutputPath $outPath
    Start-Sleep -Seconds 2
}

Output: Creates separate SQL Server VM reports for each subscription
#>

<#
EXAMPLE 6: Scheduled task with custom time (3:30 AM)
---
# Run as Administrator
.\Setup-ReportScheduledTask.ps1 `
    -TaskName "Daily Azure Report" `
    -DailyTime (Get-Date -Hour 3 -Minute 30 -Second 0) `
    -OutputDirectory "D:\AzureReports"

Output: Task scheduled for 3:30 AM daily
#>

<#
EXAMPLE 7: View scheduled task details
---
Get-ScheduledTask -TaskName "Generate Azure VM Report" | Select-Object -Property TaskName, State, @{Name="NextRunTime"; Expression={$_.Triggers[0]}}
#>

<#
EXAMPLE 8: Manually run scheduled task
---
Start-ScheduledTask -TaskName "Generate Azure VM Report"

# Check status
Get-ScheduledTask -TaskName "Generate Azure VM Report" | Select-Object -Property TaskName, State, LastRunTime
#>

<#
EXAMPLE 9: Remove scheduled task
---
Unregister-ScheduledTask -TaskName "Generate Azure VM Report" -Confirm:$false
#>

<#
EXAMPLE 10: Generate report and email it
---
$reportPath = "C:\Reports\report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
.\Generate-AzureVMReport.ps1 -OutputPath $reportPath

if (Test-Path $reportPath) {
    $emailParams = @{
        To = "admin@company.com"
        From = "reports@company.com"
        Subject = "Daily Azure VM Report - $(Get-Date -Format 'yyyy-MM-dd')"
        Body = "Please see attached Azure VM licensing report."
        Attachments = $reportPath
        SmtpServer = "smtp.company.com"
    }
    Send-MailMessage @emailParams
}
#>

<#
EXAMPLE 11: Generate report and filter for VMs with SQL Server
---
$reportPath = "C:\Reports\report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
.\Generate-AzureVMReport.ps1 -OutputPath $reportPath

# The default report already contains only SQL Server VMs.
# This example shows how to further inspect the SQL-specific rows.
$excelData = Import-Excel -Path $reportPath -WorksheetName "VMs"
$sqlVMs = $excelData | Where-Object { $_.'Has SQL Server' -eq 'Yes' }
Write-Host "Found $($sqlVMs.Count) VMs with SQL Server"
$sqlVMs | Select-Object 'VM Name', 'SQL Version', 'SQL Edition', 'SQL License', 'SQL Enterprise Required'
#>

<#
EXAMPLE 12: Generate report and filter for SQL Server Enterprise instances
---
$reportPath = "C:\Reports\report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
.\Generate-AzureVMReport.ps1 -OutputPath $reportPath

# Find VMs flagged as SQL Enterprise Required = Yes (all Enterprise edition VMs by default).
# Reviewers can override individual rows to No using the Excel dropdown before exporting.
$excelData = Import-Excel -Path $reportPath -WorksheetName "VMs"
$enterpriseSQL = $excelData | Where-Object { $_.'SQL Enterprise Required' -eq 'Yes' }
Write-Host "Enterprise SQL Server instances requiring review: $($enterpriseSQL.Count)"
$enterpriseSQL | Select-Object 'VM Name', 'vCPU Count', 'SQL Version', 'SQL Edition', 'SQL License', 'SQL Enterprise Required'
#>

<#
EXAMPLE 13: Tag a VM with SQL Server licensing information
---
$vmResourceId = "/subscriptions/12345678-1234-1234-1234-123456789012/resourceGroups/myRG/providers/Microsoft.Compute/virtualMachines/myVM"

# Tag with BYOL (Bring Your Own License)
az resource tag --ids $vmResourceId --tags SqlLicenseType=BYOL SqlServerLicense="Enterprise BYOL"

# Now regenerate report to reflect the BYOL tag
.\Generate-AzureVMReport.ps1
#>

<#
EXAMPLE 14: Generate SQL Server licensing summary
---
$reportPath = "C:\Reports\report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
.\Generate-AzureVMReport.ps1 -OutputPath $reportPath

$excelData = Import-Excel -Path $reportPath -WorksheetName "VMs"

Write-Host "=== SQL Server Licensing Summary ===" -ForegroundColor Cyan
Write-Host "Total SQL Server VMs: $(($excelData | Where-Object { $_.'Has SQL Server' -eq 'Yes' }).Count)"
Write-Host "Enterprise Editions: $(($excelData | Where-Object { $_.'SQL Edition' -eq 'Enterprise' }).Count)"
Write-Host "Standard Editions: $(($excelData | Where-Object { $_.'SQL Edition' -eq 'Standard' }).Count)"
Write-Host "Developer Editions: $(($excelData | Where-Object { $_.'SQL Edition' -eq 'Developer' }).Count)"
Write-Host "Express Editions: $(($excelData | Where-Object { $_.'SQL Edition' -eq 'Express' }).Count)"

$totalVcpuWithSQL = ($excelData | Where-Object { $_.'Has SQL Server' -eq 'Yes' } | Measure-Object -Property 'vCPU Count' -Sum).Sum
Write-Host "Total vCPUs with SQL Server: $totalVcpuWithSQL"
#>

<#
EXAMPLE 15: Export SQL Server details to separate file
---
$reportPath = "C:\Reports\report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
.\Generate-AzureVMReport.ps1 -OutputPath $reportPath

$excelData = Import-Excel -Path $reportPath -WorksheetName "VMs"
$sqlVMs = $excelData | Where-Object { $_.'Has SQL Server' -eq 'Yes' } | 
    Select-Object 'VM Name', 'vCPU Count', 'SQL Version', 'SQL Edition', 'SQL License', 'SQL Enterprise Required', 'Resource Group'

$sqlReportPath = "C:\Reports\SQLServerOnly_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
$sqlVMs | Export-Excel -Path $sqlReportPath -WorksheetName "SQL VMs" -AutoSize -TableName "SQLServers" -TableStyle "Light10"
Write-Host "SQL Server details exported to: $sqlReportPath"
#>

<#
EXAMPLE 16: Run from Azure Cloud Shell (default SQL Server-only behavior)
---
# Upload Run-In-CloudShell.ps1 to Cloud Shell first, then run:
.\Run-In-CloudShell.ps1

Output: Saves the report to ~/clouddrive/azure-reports/
#>

<#
EXAMPLE 17: Run from Azure Cloud Shell and include all VMs
---
.\Run-In-CloudShell.ps1 -SQLServerOnly:$false

Output: Saves an all-VM report to ~/clouddrive/azure-reports/
#>

# Copy and paste the examples above to execute them in PowerShell
# Modify parameters as needed for your environment
