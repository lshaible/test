# Azure VM Licensing & SQL Server Report - PowerShell Scripts

This collection of PowerShell scripts generates Azure VM licensing and SQL Server licensing reports in Excel format using Azure CLI data.

Supports both **local execution** and **Azure Cloud Shell execution**.

## Features

- **SQL Server VMs Only**: By default, reports only include VMs with SQL Server installed.
- **Windows OS Licensing Detection**: Identifies Windows VMs and licensing requirements.
- **SQL Server Detection**: Automatically detects SQL Server installations on VMs.
- **SQL Server Edition Recognition**: Identifies SQL Server editions (Enterprise, Standard, Express, Developer, Web).
- **SQL Server Version Tracking**: Detects SQL Server versions (2016, 2017, 2019, 2022, 2025).
- **Licensing Type Classification**: Categorizes licenses (License Required, BYOL, Free, etc.).
- **Comprehensive Metrics**: Generates summary statistics for SQL Server deployments.
- **Cloud Shell Support**: Run reports directly from Azure Cloud Shell without local tools.
- **Flexible Filtering**: Option to include all VMs or SQL Server only.

## Scripts

### 1. Generate-AzureVMReport.ps1

Main script that generates the Excel report with VM details. **By default, only includes VMs with SQL Server installed.**

**Supports**: Local PowerShell, Azure Cloud Shell

**Usage:**

```powershell
# Local execution - SQL Server VMs only (default)
.\Generate-AzureVMReport.ps1 -SubscriptionId "your-subscription-id" -OutputPath "C:\Reports\report.xlsx"

# Cloud Shell execution - SQL Server VMs only (default)
.\Generate-AzureVMReport.ps1 -SubscriptionId "your-subscription-id" -Environment CloudShell

# Include all VMs (not just SQL Server)
.\Generate-AzureVMReport.ps1 -SQLServerOnly:$false
```

**Parameters:**

- `-SubscriptionId` (optional): Azure subscription ID. If not provided, uses the current subscription.
- `-OutputPath` (optional): Path where the Excel file will be saved.
  - Local: Defaults to `Azure_Windows_VM_Licensing_vCPU_Report_yyyyMMdd_HHmmss.xlsx` in the script directory.
  - Cloud Shell: Defaults to `~/CloudShell/Azure_Windows_VM_Licensing_vCPU_Report_yyyyMMdd_HHmmss.xlsx`.
  - If you supply `-OutputPath`, the file is saved with exactly the name you provide.
- `-Environment` (optional): `Local` or `CloudShell` (auto-detected if not specified).
- `-SQLServerOnly` (optional): Filter to only include VMs with SQL Server (default: `$true`).

**Output:**

Creates an Excel file with two worksheets:

- **VMs**: Detailed list of all VMs with columns:
  - Subscription
  - Resource Group
  - VM Name
  - VM Size
  - vCPU Count
  - OS Type
  - Publisher
  - Offer
  - Windows License (License Required, BYOL, N/A)
  - Has SQL Server (Yes/No)
  - SQL Version (SQL Server 2016/2017/2019/2022/2025)
  - SQL Edition (Enterprise, Standard, Express, Developer, Web)
  - SQL License (License Required, BYOL, Free, etc.)
  - SQL Enterprise Required (Yes/No)
  - Provisioning State
  - Scan Date
- **Summary**: Key metrics including:
  - Total VMs
  - Total vCPUs
  - Windows VMs count
  - Linux VMs count
  - VMs with SQL Server count
  - SQL Server 2025 instances
  - SQL Server 2022 instances
  - SQL Server 2019 instances
  - SQL Server Enterprise edition instances
  - SQL Enterprise Required (Yes) count
  - SQL Server Standard edition instances
  - SQL Server Developer edition instances
  - SQL Server Express edition instances
  - Report generation time

### 2. Generate-AzureVMReport-Scheduled.ps1

Wrapper script designed for scheduled execution with logging.

**Usage:**

```powershell
.\Generate-AzureVMReport-Scheduled.ps1 -SubscriptionId "your-subscription-id" -OutputDirectory "C:\Reports" -LogFile "C:\Logs\AzureVMReport.log"
```

**Parameters:**

- `-SubscriptionId` (optional): Azure subscription ID.
- `-OutputDirectory`: Directory where reports will be saved (default: `C:\Reports`).
- `-LogFile`: Path to log file (default: `C:\Logs\AzureVMReport.log`).

### 3. Setup-ReportScheduledTask.ps1

Creates a Windows Task Scheduler job to run reports automatically.

**Requirements:**

- Must be run as Administrator.
- Local PowerShell only (not supported in Cloud Shell).

**Usage:**

```powershell
# Run as Administrator
.\Setup-ReportScheduledTask.ps1 -SubscriptionId "your-subscription-id" -DailyTime "02:00:00" -RunImmediately
```

**Parameters:**

- `-TaskName`: Name of the scheduled task (default: `Generate Azure VM Report`).
- `-SubscriptionId` (optional): Azure subscription ID.
- `-OutputDirectory`: Directory where reports will be saved (default: `C:\Reports`).
- `-LogFile`: Path to log file (default: `C:\Logs\AzureVMReport.log`).
- `-DailyTime`: Time to run daily (default: 2:00 AM).
- `-RunImmediately`: Run the task immediately after creation (switch).

### 4. CloudShell-Setup.ps1

Prepares and uploads scripts to your Azure Cloud Shell storage.

**Requirements:**

- Run from local PowerShell (not from Cloud Shell).
- Storage account access permissions.

**Usage:**

```powershell
.\CloudShell-Setup.ps1 -StorageAccountName "mystorageacct" -ResourceGroupName "myResourceGroup" -SubscriptionId "your-subscription-id"
```

**What it does:**

1. Creates a file share in your Cloud Shell storage account.
2. Uploads all report scripts to the file share.
3. Makes scripts available in Cloud Shell at `~/clouddrive/scripts/azure-reports/`.

### 5. Run-In-CloudShell.ps1

Optimized script for running directly in Azure Cloud Shell. Single, self-contained file with all dependencies included. **By default, only includes VMs with SQL Server installed.**

**Usage in Cloud Shell:**

#### Uploading the Script to Cloud Shell

##### Option 1: Use the Cloud Shell Upload Button

1. Go to <https://shell.azure.com> and switch to **PowerShell**.
2. Click the **Upload/Download files** icon in the Cloud Shell toolbar.
3. Click **Upload**.
4. Browse to `C:\SQL-scripts\Run-In-CloudShell.ps1` on your local machine and select it.
5. The file uploads to your home directory (`~/`).
6. Run it:

```powershell
.\Run-In-CloudShell.ps1
```

##### Option 2: Use the Azure Portal File Share

1. Go to <https://portal.azure.com>.
2. Navigate to the **Storage Account** linked to your Cloud Shell, typically in the Cloud Shell resource group and often named `cloud-shell-storage-*`.
3. Click **File shares** and open the share named `cs-*`.
4. Upload `Run-In-CloudShell.ps1` using the **Upload** button.
5. In Cloud Shell, copy it from `~/clouddrive/` and run it:

```powershell
cp ~/clouddrive/Run-In-CloudShell.ps1 ~/
.\Run-In-CloudShell.ps1
```

##### Option 3: Paste Script Content Directly

1. Open `Run-In-CloudShell.ps1` in a text editor locally.
2. Copy all content.
3. In Cloud Shell, create the file and paste:

```powershell
code Run-In-CloudShell.ps1
# Paste content, save with Ctrl+S, close with Ctrl+Q
.\Run-In-CloudShell.ps1
```

#### Downloading the Report from Cloud Shell

##### Option 1: Use the Cloud Shell Download Button

1. In Cloud Shell, note the report path printed after the script runs, for example `~/clouddrive/azure-reports/Azure_VM_Report_20260421_120000.xlsx`.
1. Click the **Upload/Download files** icon in the Cloud Shell toolbar.
1. Click **Download**.
1. Type the full path to the report file, for example:

```text
/home/yourname/clouddrive/azure-reports/Azure_VM_Report_20260421_120000.xlsx
```

1. Click **Download**. The file saves to your local Downloads folder.

##### Option 2: Download via Azure Portal File Share

1. Go to <https://portal.azure.com>.
2. Navigate to your Cloud Shell **Storage Account**, then **File shares**, then open the `cs-*` share.
3. Browse to the `azure-reports` folder.
4. Click the report file and select **Download**.

##### Option 3: List Available Reports in Cloud Shell

```powershell
# See all generated reports
ls ~/clouddrive/azure-reports/

# Get the most recent report name
Get-ChildItem ~/clouddrive/azure-reports/ | Sort-Object LastWriteTime -Descending | Select-Object -First 1
```

**Parameters:**

- `-SubscriptionId` (optional): Azure subscription ID.
- `-OutputStorageAccount` (optional): Storage account to upload reports.
- `-OutputStorageContainer`: Storage container name (default: `reports`).
- `-SaveToCloudDrive`: Save to Cloud Shell persistent storage (default: `$true`).
- `-SQLServerOnly` (optional): Filter to only include VMs with SQL Server (default: `$true`).

## Prerequisites

### Common Requirements

1. **Azure CLI**: Must be installed and authenticated.

   ```powershell
   az login
   ```

2. **Permissions**: Must have permissions to:
   - List VMs in the Azure subscription.
   - Read VM properties.

### For Local Execution

- PowerShell 5.1 or higher.
- Windows, macOS, or Linux.
- ImportExcel module (auto-installed on first run).

### For Azure Cloud Shell Execution

- No local tools needed.
- Access to <https://shell.azure.com>.
- Cloud Shell storage account (auto-created by Azure).
- PowerShell environment in Cloud Shell.

## Execution Methods

### Method 1: Local Execution (Windows/Mac/Linux)

Run scripts directly on your machine with local PowerShell.

**Advantages:**

- Scheduled task automation.
- Local report storage.
- Full control over execution.

**Steps:**

```powershell
cd C:\SQL-scripts
.\Generate-AzureVMReport.ps1 -SubscriptionId "your-sub-id"
```

### Method 2: Cloud Shell (Browser-based)

Run scripts directly in Azure Cloud Shell via browser.

**Advantages:**

- No local tools installation needed.
- Access from anywhere.
- Integrated with Azure.
- Persistent storage in Cloud Shell.

**Steps:**

```powershell
# Option A: Direct execution
cd ~/clouddrive/azure-reports
.\Run-In-CloudShell.ps1

# Option B: Using pre-uploaded scripts
.\Generate-AzureVMReport.ps1 -Environment CloudShell
```

**To set up Cloud Shell storage first (run from local PowerShell):**

```powershell
.\CloudShell-Setup.ps1 -StorageAccountName "mystorageacct" -ResourceGroupName "myRG" -SubscriptionId "my-sub-id"
```

### Method 3: Automated Scheduling (Local Windows)

Schedule automated daily reports on Windows.

**Advantages:**

- Automatic daily execution.
- No user intervention needed.
- Audit logging.

**Steps:**

```powershell
# Run as Administrator
.\Setup-ReportScheduledTask.ps1 -SubscriptionId "your-sub-id" -RunImmediately
```

## Quick Start Guide

### Quick Start: Cloud Shell (Recommended for Cloud-Only)

1. Go to <https://shell.azure.com>.
2. Switch to PowerShell.
3. Upload `Run-In-CloudShell.ps1` using one of the methods in the **Run-In-CloudShell.ps1** section above.
4. Run the script from Cloud Shell:

   ```powershell
   # Run the report for SQL Server VMs only (default)
   .\Run-In-CloudShell.ps1
   ```

5. Report saved to `~/clouddrive/azure-reports/`.

### Quick Start: Local (Recommended for Windows with Scheduling)

1. Save scripts to `C:\SQL-scripts\`.
2. Open PowerShell.
3. Run:

   ```powershell
   cd C:\SQL-scripts
   # Generate report for SQL Server VMs only (default)
   .\Generate-AzureVMReport.ps1
   ```

4. Report in current directory.

### Quick Start: Include All VMs

```powershell
# To report on ALL VMs (not just SQL Server):
.\Generate-AzureVMReport.ps1 -SQLServerOnly:$false
```

### Validate All-VM Mode

Use this command to confirm the report includes both SQL and non-SQL VMs:

```powershell
.\Generate-AzureVMReport.ps1 -SQLServerOnly:$false

$latest = Get-ChildItem .\Azure_Windows_VM_Licensing_vCPU_Report_*.xlsx |
   Sort-Object LastWriteTime -Descending |
   Select-Object -First 1

$rows = Import-Excel -Path $latest.FullName -WorksheetName 'VMs'
$total = $rows.Count
$sqlCount = ($rows | Where-Object { $_.'Has SQL Server' -eq 'Yes' }).Count
$nonSqlCount = ($rows | Where-Object { $_.'Has SQL Server' -ne 'Yes' }).Count

Write-Host "Report: $($latest.Name)"
Write-Host "TotalVMs=$total SQLVMs=$sqlCount NonSQLVMs=$nonSqlCount"
```

Expected behavior:

- `TotalVMs` includes all VMs in the subscription.
- `SQLVMs` includes only rows where `Has SQL Server = Yes`.
- `NonSQLVMs` includes rows where `Has SQL Server = No`.
- SQL-specific fields (`SQL Version`, `SQL Edition`, `SQL License`) show `N/A` for non-SQL VMs.
- `SQL Enterprise Required` shows `Yes` for all Enterprise edition VMs regardless of license type; reviewers can override individual rows to `No` using the dropdown in Excel.

### Quick Start: Set Up Cloud Shell (One-time Setup)

1. From local PowerShell:

   ```powershell
   .\CloudShell-Setup.ps1 -StorageAccountName "mystorageacct" -ResourceGroupName "myRG" -SubscriptionId "12345..."
   ```

2. Go to <https://shell.azure.com>.
3. Navigate to `~/clouddrive/scripts/azure-reports`.
4. Run `.\Generate-AzureVMReport.ps1 -Environment CloudShell`.

## Troubleshooting

### Local Execution Issues

#### Issue: ImportExcel Module Not Found

**Solution:** Run the main script. It will attempt to install the module automatically.

#### Issue: Access Denied When Creating Scheduled Task

**Solution:** Run PowerShell as Administrator.

#### Issue: Not Authenticated with Azure CLI

**Solution:** Run `az login` in PowerShell before running the report script.

### Cloud Shell Execution Issues

#### Issue: ImportExcel Module Not Found in Cloud Shell

**Solution:** The script will attempt to install it. If it fails:

```powershell
Update-Module ImportExcel -Force
```

#### Issue: Storage Account Not Found in CloudShell-Setup

**Solution:** Ensure the storage account name, resource group, and subscription ID are correct:

```powershell
az storage account list --query "[].name" -o table
```

#### Issue: Cannot Access `~/clouddrive`

**Solution:** Ensure Cloud Shell storage account is initialized:

1. Go to <https://shell.azure.com>.
2. Click **Create storage** if prompted.
3. Select a resource group and storage account.

#### Issue: Scripts Not Visible in Cloud Shell After Upload

**Solution:** Refresh Cloud Shell or run:

```powershell
ls ~/clouddrive/scripts/azure-reports/
```

#### Issue: Permission Denied on Scripts in Cloud Shell

**Solution:** Make scripts executable:

```powershell
chmod +x ~/clouddrive/scripts/azure-reports/*.ps1
```

### General Issues

#### Issue: No VMs Found

**Possible causes:**

- Wrong subscription selected.
- No VMs in the subscription.
- Insufficient permissions.

**Solution:** Verify subscription with `az account show` and check permissions.

#### Issue: SQL Server Not Detected

**Possible causes:**

- SQL Server is custom-installed and not from a marketplace image.
- Image offer name does not contain `SQL`.
- Image is from a custom gallery.

**Solution:** The script detects SQL using image publisher, offer, and SKU metadata from Azure Marketplace images. For custom installations, manually tag the VM with SQL licensing tags for accurate reporting.

## SQL Server License Tag Reference

To override auto-detection or provide custom licensing info, use Azure resource tags:

```powershell
# Example: Merge SQL BYOL tags onto a VM without replacing existing tags
az tag update --resource-id /subscriptions/{sub}/resourceGroups/{rg}/providers/Microsoft.Compute/virtualMachines/{vmName} `
   --operation merge `
    --tags SqlLicenseType=BYOL SqlServerLicense="Enterprise BYOL"
```

Supported tag keys:

- `SqlLicenseType`: `BYOL` or `PAY_AS_YOU_GO`.
- `SqlServerLicense`: Custom license description.
- `WindowsLicenseType`: `BYOL` for Windows OS.

## Cloud Shell Notes

### Setting Up Cloud Shell with Persistent Storage

Persistent storage ensures your scripts and reports survive between Cloud Shell sessions. Without it, all files are lost when the session ends.

#### First-Time Setup (Mounting Persistent Storage)

1. Go to <https://shell.azure.com>.
2. Select **PowerShell** as your shell type.
3. If prompted that no storage is mounted, choose **Show advanced settings** to control the storage account, or choose **Create storage** to let Azure create one automatically.
4. If using advanced settings, fill in:
   - **Subscription**: your Azure subscription.
   - **Cloud Shell region**: choose a nearby region.
   - **Resource group**: create new or use existing, for example `cloud-shell-rg`.
   - **Storage account**: create new or use existing, for example `cloudshellstorage[yourname]`.
   - **File share**: create new or use existing, for example `cloudshell`.
5. Click **Attach storage**. Cloud Shell mounts the file share to `~/clouddrive/`.

#### Validating Whether Persistent Storage Is Active

Run these commands in Cloud Shell:

```powershell
# Check 1: Verify clouddrive is mounted
df -h | grep clouddrive

# Check 2: View the storage mount details
cat ~/.cloud_drive

# Check 3: Write a test file and verify it exists
"Storage test $(Get-Date)" | Out-File ~/clouddrive/storage-test.txt
ls ~/clouddrive/

# Check 4: View Cloud Shell environment info
env | grep CLOUD
```

If storage is mounted, you should see a mounted path under `/home/.../clouddrive`, the `.cloud_drive` file should exist, and the test file should appear in `~/clouddrive/`.

#### If Persistent Storage Is Not Mounted

Cloud Shell will show a warning that your cloud drive is not mounted and that files saved outside `~/clouddrive` will not persist.

To fix it:

```text
Use the Cloud Shell Settings menu to reset user settings and re-attach storage,
or restart Cloud Shell from the Azure portal and attach storage when prompted.
```

#### Verifying Reports Are Saved to Persistent Storage

After running the report script, confirm the output landed in the persistent directory:

```powershell
# List all reports saved to persistent cloud drive
Get-ChildItem ~/clouddrive/azure-reports/ | Sort-Object LastWriteTime -Descending

# Confirm the most recent report exists and has content
$latest = Get-ChildItem ~/clouddrive/azure-reports/*.xlsx | Sort-Object LastWriteTime -Descending | Select-Object -First 1
Write-Host "Latest report: $($latest.FullName)"
Write-Host "Size: $($latest.Length) bytes"
Write-Host "Created: $($latest.LastWriteTime)"
```

### Persistent Storage

- Reports saved to `~/clouddrive/azure-reports/` persist across Cloud Shell sessions.
- Cloud Shell storage uses your subscription's storage account.
- Files in Cloud Shell are accessible from your local machine via Azure Portal.

### Uploading Reports to Storage

To automatically upload reports to Azure Storage:

```powershell
.\Run-In-CloudShell.ps1 -OutputStorageAccount "myaccount" -OutputStorageContainer "reports"
```

### Cloud Shell Limitations

- Scripts must use path separators compatible with both Windows and Linux. Use `/`.
- Scheduled tasks are not supported. Use local execution for scheduling.
- Task Scheduler features are not available in Cloud Shell.

### Recommended Cloud Shell Workflow

1. Run reports monthly or weekly via Cloud Shell.
2. Store reports in Cloud Drive for persistence.
3. Optionally upload to Azure Storage Blob for archive.
4. Download reports locally as needed.

## Excel Output Features

- **Formatted Headers**: Blue header row with white text.
- **Auto-sized Columns**: Columns automatically sized for content.
- **Frozen Panes**: Header row frozen for easy scrolling.
- **Summary Sheet**: Quick overview metrics.
- **Data Table**: Formatted as Excel table for sorting and filtering.
- **Timestamp**: Report includes scan date for audit trail.

## Data Dictionary - SQL Server Columns

| Column | Definition |
| --- | --- |
| **Has SQL Server** | `Yes` if VM has SQL Server installed (detected from marketplace image), `No` otherwise. |
| **SQL Version** | SQL Server version (2016, 2017, 2019, 2022, 2025) or `N/A` for non-SQL VMs. |
| **SQL Edition** | SQL Server edition (Enterprise, Standard, Express, Developer, Web) or `N/A` for non-SQL VMs. |
| **SQL License** | Licensing status: `License Required` (pay-per-license), `BYOL` (Bring Your Own License), `Free (Limited)` (Express), `Free (Dev/Test)` (Developer), or `N/A` for non-SQL VMs. |
| **SQL Enterprise Required** | `Yes` if VM has Enterprise SQL Server edition (regardless of license type); `No` for all other editions or non-SQL VMs. Reviewers can override individual rows to `No` using the Excel dropdown. Used to flag VMs that may need compliance review. |

## Notes

- **SQL Server Filtering**: By default, only VMs with SQL Server are included in reports. Use `-SQLServerOnly:$false` to include all VMs.
- **vCPU counts**: Looked up from a built-in map and, when needed, resolved from Azure VM SKUs by VM location.
- **SQL Server detection**: Based on Azure Marketplace image publisher, offer, and SKU metadata.
- **SQL Server editions**: Parsed from image SKU first, with offer-name fallback (Enterprise, Standard, Express, Developer, Web).
- **Licensing overrides**: Can be overridden using Azure tags on the VM resource.
- **Windows licensing**: Detected from VM `licenseType` and tags, with image metadata fallback.
- **BYOL support**: SQL Server BYOL can be detected from VM metadata and tags (`SqlLicenseType=BYOL`), and Windows BYOL can be detected from VM `licenseType=Windows_Server` or `WindowsLicenseType=BYOL` tags.
- **All timestamps**: In local time zone.
- **Reports**: Cumulative. Each run generates a new file with a timestamp.
- **Free editions**: Express and Developer editions are flagged as free or dev-only licenses.
- **Cloud Shell**: Automatically detects the Cloud Shell environment and adapts path handling.
- **File paths**: Local scripts use Windows paths (`C:\...`); Cloud Shell uses Unix paths (`~/`).
- **Module installation**: Both environments support automatic ImportExcel module installation.

## Support

For issues or feature requests, check:

1. Azure CLI documentation: <https://docs.microsoft.com/cli/azure/>
2. ImportExcel documentation: <https://github.com/dfinke/ImportExcel>
3. Azure PowerShell documentation: <https://docs.microsoft.com/powershell/azure/>
