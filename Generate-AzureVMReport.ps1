param(
    [string]$SubscriptionId,
    [string]$OutputPath,
    [ValidateSet('Local', 'CloudShell')]
    [string]$Environment = 'Local',
    [bool]$SQLServerOnly = $true
)

# Auto-detect Cloud Shell environment if not specified
if (-not $Environment) {
    if ($env:CLOUD_SHELL -eq 'true' -or (Test-Path '/bin/bash')) {
        $Environment = 'CloudShell'
    }
    else {
        $Environment = 'Local'
    }
}

# Set default output path based on environment
if (-not $OutputPath) {
    if ($Environment -eq 'CloudShell') {
        $OutputPath = "$HOME/CloudShell/Azure_Windows_VM_Licensing_vCPU_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
        # Create CloudShell directory if it doesn't exist
        if (-not (Test-Path "$HOME/CloudShell")) {
            New-Item -ItemType Directory -Path "$HOME/CloudShell" -Force | Out-Null
        }
    }
    else {
        $OutputPath = "$PSScriptRoot/Azure_Windows_VM_Licensing_vCPU_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    }
}

# Function to check if ImportExcel module is installed
function Test-ImportExcelModule {
    if (Get-Module -ListAvailable -Name ImportExcel) {
        return $true
    }
    return $false
}

# Function to install ImportExcel module
function Install-ImportExcelModule {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Yellow
    try {
        # In Cloud Shell, use Force to bypass confirmation
        if ($Environment -eq 'CloudShell') {
            Install-Module -Name ImportExcel -Force -Scope CurrentUser -SkipPublisherCheck -AllowClobber -ErrorAction Stop
        }
        else {
            Install-Module -Name ImportExcel -Force -Scope CurrentUser -ErrorAction Stop
        }
        Write-Host "ImportExcel module installed successfully." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to install ImportExcel module: $_"
        return $false
    }
}

# Check and install ImportExcel if needed
if (-not (Test-ImportExcelModule)) {
    Write-Host "ImportExcel module not found." -ForegroundColor Yellow
    if ($Environment -eq 'CloudShell') {
        Write-Host "Cloud Shell environment detected. Attempting to install ImportExcel..." -ForegroundColor Cyan
    }
    
    if (-not (Install-ImportExcelModule)) {
        Write-Error "ImportExcel module is required but could not be installed."
        Write-Error "In Cloud Shell: You may need to use 'Update-Module ImportExcel' or reinstall Cloud Shell."
        exit 1
    }
}

Import-Module ImportExcel

# Check if SqlServer module is available (needed for database counting)
$sqlServerModuleAvailable = Get-Module -ListAvailable -Name SqlServer
if (-not $sqlServerModuleAvailable) {
    Write-Host "Note: SqlServer PowerShell module not found. Database counting will be skipped." -ForegroundColor Yellow
    Write-Host "To enable database discovery, install with: Install-Module -Name SqlServer -Force" -ForegroundColor Cyan
}

# Function to get VM vCPU count
function Get-VMvCPUCount {
    param(
        [string]$VMSize,
        [string]$Location
    )
    
    # Common Azure VM sizes and their vCPU counts
    $vmSizeMap = @{
        "Standard_B1s" = 1
        "Standard_B1ms" = 1
        "Standard_B2s" = 2
        "Standard_B2ms" = 2
        "Standard_B4ms" = 4
        "Standard_B12ms" = 12
        "Standard_D1_v2" = 1
        "Standard_D2_v2" = 2
        "Standard_D3_v2" = 4
        "Standard_D4_v2" = 8
        "Standard_D5_v2" = 16
        "Standard_E2_v3" = 2
        "Standard_E4_v3" = 4
        "Standard_E8_v3" = 8
        "Standard_E16_v3" = 16
        "Standard_E32_v3" = 32
        "Standard_F2s_v2" = 2
        "Standard_F4s_v2" = 4
        "Standard_F8s_v2" = 8
        "Standard_F16s_v2" = 16
    }
    
    if ($vmSizeMap.ContainsKey($VMSize)) {
        return $vmSizeMap[$VMSize]
    }
    # Lookup from Azure for sizes not in the static map.
    if ($Location) {
        try {
            $skuInfo = az vm list-skus --location $Location --resource-type virtualMachines --query "[?name=='$VMSize']" -o json 2>$null | ConvertFrom-Json
            if ($skuInfo -and $skuInfo.Count -gt 0) {
                $vCpuCapability = $skuInfo[0].capabilities | Where-Object { $_.name -eq 'vCPUs' } | Select-Object -First 1
                if ($vCpuCapability -and ($vCpuCapability.value -as [int])) {
                    return [int]$vCpuCapability.value
                }
            }
        }
        catch { }
    }

    return 0
}

# Function to detect SQL Server installation
function Get-SQLServerInfo {
    param(
        [string]$ImagePublisher,
        [string]$ImageOffer,
        [string]$ImageSku,
        [string]$VmLicenseType,
        [hashtable]$Tags
    )
    
    $sqlInfo = @{
        'HasSQL' = $false
        'SQLEdition' = 'N/A'
        'SQLVersion' = 'N/A'
        'SQLLicense' = 'N/A'
    }
    
    # Detect SQL Server from publisher/offer/sku.
    if (($ImagePublisher -like "*microsoftsqlserver*") -or ($ImageOffer -match "sql") -or ($ImageSku -match "sql")) {
        $sqlInfo['HasSQL'] = $true
        
        # Parse SQL version from offer or SKU.
        if ($ImageOffer -match "SQL2025") {
            $sqlInfo['SQLVersion'] = "SQL Server 2025"
        }
        elseif ($ImageOffer -match "SQL2022") {
            $sqlInfo['SQLVersion'] = "SQL Server 2022"
        }
        elseif ($ImageOffer -match "SQL2019") {
            $sqlInfo['SQLVersion'] = "SQL Server 2019"
        }
        elseif ($ImageOffer -match "SQL2017") {
            $sqlInfo['SQLVersion'] = "SQL Server 2017"
        }
        elseif ($ImageOffer -match "SQL2016") {
            $sqlInfo['SQLVersion'] = "SQL Server 2016"
        }
        elseif ($ImageSku -match "2025") {
            $sqlInfo['SQLVersion'] = "SQL Server 2025"
        }
        elseif ($ImageSku -match "2022") {
            $sqlInfo['SQLVersion'] = "SQL Server 2022"
        }
        elseif ($ImageSku -match "2019") {
            $sqlInfo['SQLVersion'] = "SQL Server 2019"
        }
        elseif ($ImageSku -match "2017") {
            $sqlInfo['SQLVersion'] = "SQL Server 2017"
        }
        elseif ($ImageSku -match "2016") {
            $sqlInfo['SQLVersion'] = "SQL Server 2016"
        }
        else {
            $sqlInfo['SQLVersion'] = "SQL Server (Version Unknown)"
        }
        
        # Determine edition from SKU first, then fallback to offer.
        if ($ImageSku -match "enterprise" -or $ImageOffer -match "Enterprise") {
            $sqlInfo['SQLEdition'] = "Enterprise"
            $sqlInfo['SQLLicense'] = "License Required"
        }
        elseif ($ImageSku -match "standard" -or $ImageOffer -match "Standard") {
            $sqlInfo['SQLEdition'] = "Standard"
            $sqlInfo['SQLLicense'] = "License Required"
        }
        elseif ($ImageSku -match "express" -or $ImageOffer -match "Express") {
            $sqlInfo['SQLEdition'] = "Express"
            $sqlInfo['SQLLicense'] = "Free (Limited)"
        }
        elseif ($ImageSku -match "web" -or $ImageOffer -match "Web") {
            $sqlInfo['SQLEdition'] = "Web"
            $sqlInfo['SQLLicense'] = "License Required"
        }
        elseif ($ImageSku -match "developer" -or $ImageOffer -match "Developer") {
            $sqlInfo['SQLEdition'] = "Developer"
            $sqlInfo['SQLLicense'] = "Free (Dev/Test)"
        }
        else {
            $sqlInfo['SQLEdition'] = "Edition Unknown"
            $sqlInfo['SQLLicense'] = "License Required"
        }
        
        # Check explicit BYOL hints from VM license type or tags.
        if ($VmLicenseType -match "SQL_Server_BYOL") {
            $sqlInfo['SQLLicense'] = "BYOL"
        }
        elseif ($Tags -and $Tags.ContainsKey("SqlLicenseType") -and $Tags["SqlLicenseType"] -eq "BYOL") {
            $sqlInfo['SQLLicense'] = "BYOL"
        }
        elseif ($Tags -and $Tags.ContainsKey("SqlServerLicense")) {
            $sqlInfo['SQLLicense'] = $Tags["SqlServerLicense"]
        }
    }
    
    return $sqlInfo
}

# Function to determine Windows licensing type
function Get-WindowsLicensingType {
    param(
        [string]$OSType,
        [string]$ImagePublisher,
        [string]$ImageOffer,
        [string]$VmLicenseType,
        [hashtable]$Tags
    )
    
    if ($OSType -ne "Windows") {
        return "N/A"
    }

    # Windows BYOL (Azure Hybrid Benefit) can come from VM metadata or tags.
    if ($VmLicenseType -eq "Windows_Server") {
        return "BYOL"
    }

    if ($Tags -and $Tags.ContainsKey("WindowsLicenseType") -and $Tags["WindowsLicenseType"] -eq "BYOL") {
        return "BYOL"
    }

    # For Windows marketplace images without BYOL indicators, treat as license included/required.
    if ($ImagePublisher -like "*Microsoft*" -and ($ImageOffer -like "*Windows*" -or $ImageOffer -like "*Server*" -or $ImageOffer -like "*SQL*")) {
        return "License Required"
    }
    
    return "N/A"
}

# Function to find VMs with SQL IaaS Agent Extension (discovers custom/non-marketplace SQL)
function Get-VMsWithSQLExtension {
    Write-Host "Scanning for VMs with SQL IaaS Agent Extension..." -ForegroundColor Yellow
    
    try {
        $query = "resources | where type =~ 'microsoft.compute/virtualmachines/extensions' | where name has_any ('SqlIaasAgent','SqlIaas') | project vmId=split(id, '/')[8], vmName=split(id, '/')[8], extensionName=name, resourceGroup, subscriptionId"
        $queryResult = az graph query -q $query --query "data" --only-show-errors -o json 2>&1
        
        # Check if result looks like an error (not JSON)
        if ($queryResult -match '^\s*(ERROR|error|Error|Request)' -or $queryResult -notmatch '^\s*(\[|\{)') {
            Write-Host "SQL IaaS Agent Extension discovery not available (insufficient permissions or unsupported query)" -ForegroundColor Yellow
            return @()
        }
        
        if (-not $queryResult -or $queryResult -eq '[]' -or [string]::IsNullOrWhiteSpace($queryResult)) {
            Write-Host "No VMs with SQL IaaS Agent Extension found" -ForegroundColor Green
            return @()
        }
        
        $sqlExtensionVMs = $queryResult | ConvertFrom-Json -ErrorAction Stop
        $count = if ($sqlExtensionVMs -is [array]) { $sqlExtensionVMs.Count } else { if ($sqlExtensionVMs) { 1 } else { 0 } }
        
        Write-Host "Found $count VMs with SQL IaaS Agent Extension" -ForegroundColor Green
        return $sqlExtensionVMs
    }
    catch {
        Write-Host "SQL IaaS Agent Extension discovery skipped (error: $($_.Exception.Message))" -ForegroundColor Yellow
        return @()
    }
}

# Function to discover Azure SQL managed services (Database servers and Managed Instances)
function Get-AzureSQLResources {
    Write-Host "Scanning for Azure SQL Database servers and Managed Instances..." -ForegroundColor Yellow
    
    try {
        $query = "resources | where type in~ ('microsoft.sql/servers','microsoft.sql/managedinstances') | project name, type, location, resourceGroup, subscriptionId, fullyQualifiedDomainName=tostring(properties.fullyQualifiedDomainName), adminLogin=tostring(properties.administratorLogin)"
        $queryResult = az graph query -q $query --query "data" --only-show-errors -o json 2>&1
        
        if (-not $queryResult -or $queryResult -eq '[]' -or [string]::IsNullOrWhiteSpace($queryResult)) {
            Write-Host "No Azure SQL resources found" -ForegroundColor Green
            return @()
        }
        
        $sqlResources = $queryResult | ConvertFrom-Json -ErrorAction Stop
        $count = if ($sqlResources -is [array]) { $sqlResources.Count } else { if ($sqlResources) { 1 } else { 0 } }
        
        Write-Host "Found $count Azure SQL resources" -ForegroundColor Green
        return $sqlResources
    }
    catch {
        Write-Host "Note: Azure SQL resource discovery unavailable (error: $($_.Exception.Message))." -ForegroundColor Yellow
        return @()
    }
}

# Function to count databases in a SQL Server instance
function Get-SqlDatabaseCount {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ServerInstance,
        
        [string]$Database = "master",
        [string]$Username,
        [SecureString]$Password,
        [string]$AccessToken
    )
    
    try {
        $query = "SELECT COUNT(*) AS database_count FROM sys.databases WHERE database_id > 4;"
        
        if ($AccessToken) {
            # Use Entra/AAD token (Azure SQL auth)
            $result = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -AccessToken $AccessToken -Query $query -ErrorAction Stop 2>$null
        }
        elseif ($Username -and $Password) {
            # Use SQL authentication with a secure password input.
            $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
            try {
                $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
                $result = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Username $Username -Password $plainPassword -Query $query -ErrorAction Stop 2>$null
            }
            finally {
                if ($bstr -ne [IntPtr]::Zero) {
                    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
                }
                $plainPassword = $null
            }
        }
        else {
            # Use integrated/Windows authentication
            $result = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query $query -ErrorAction Stop 2>$null
        }
        
        return $result.database_count
    }
    catch {
        Write-Host "  Warning: Could not connect to $ServerInstance - $_" -ForegroundColor Yellow
        return "N/A"
    }
}

# Main script
Write-Host "Starting Azure VM Report Generation..." -ForegroundColor Cyan
Write-Host "Environment: $Environment" -ForegroundColor Cyan
Write-Host "SQL Server Only: $(if ($SQLServerOnly) { 'Yes' } else { 'No' })" -ForegroundColor Cyan
Write-Host "Output Path: $OutputPath" -ForegroundColor Cyan

try {
    # Set subscription if provided
    if ($SubscriptionId) {
        Write-Host "Setting subscription to: $SubscriptionId" -ForegroundColor Yellow
        az account set --subscription $SubscriptionId
    }
    
    # Get current subscription info
    $subscriptionInfo = az account show | ConvertFrom-Json
    Write-Host "Using subscription: $($subscriptionInfo.name) ($($subscriptionInfo.id))" -ForegroundColor Green
    
    # Get all VMs
    Write-Host "Retrieving Azure VMs..." -ForegroundColor Yellow
    $vms = az vm list --query '[*].[id,name,resourceGroup,hardwareProfile.vmSize,storageProfile.osDisk.osType]' -o json | ConvertFrom-Json
    
    $vmDetails = @()
    
    foreach ($vm in $vms) {
        $vmId = $vm[0]
        $vmName = $vm[1]
        $resourceGroup = $vm[2]
        $vmSize = $vm[3]
        $osType = $vm[4]
        
        Write-Host "Processing VM: $vmName" -ForegroundColor Gray
        
        # Get detailed VM info
        $vmFullInfo = az vm show --ids $vmId --query '[name,hardwareProfile.vmSize,storageProfile.imageReference.publisher,storageProfile.imageReference.offer,storageProfile.imageReference.sku,provisioningState,tags,location,licenseType]' -o json | ConvertFrom-Json

        # az CLI returns tags as PSCustomObject after ConvertFrom-Json; normalize to hashtable for tag lookups.
        $vmTags = @{}
        if ($null -ne $vmFullInfo[6]) {
            if ($vmFullInfo[6] -is [hashtable]) {
                $vmTags = $vmFullInfo[6]
            }
            else {
                foreach ($prop in $vmFullInfo[6].PSObject.Properties) {
                    $vmTags[$prop.Name] = $prop.Value
                }
            }
        }
        
        # Get SQL Server info
        $sqlInfo = Get-SQLServerInfo -ImagePublisher $vmFullInfo[2] -ImageOffer $vmFullInfo[3] -ImageSku $vmFullInfo[4] -VmLicenseType $vmFullInfo[8] -Tags $vmTags
        
        # Skip non-SQL VMs if filtering for SQL only
        if ($SQLServerOnly -and -not $sqlInfo['HasSQL']) {
            Write-Host "  (Skipped - No SQL Server)" -ForegroundColor DarkGray
            continue
        }
        
        $vCPUCount = Get-VMvCPUCount -VMSize $vmSize -Location $vmFullInfo[7]
        $windowsLicense = Get-WindowsLicensingType -OSType $osType -ImagePublisher $vmFullInfo[2] -ImageOffer $vmFullInfo[3] -VmLicenseType $vmFullInfo[8] -Tags $vmTags
        
        # Attempt to get database count if SQL Server is detected and module is available
        $databaseCount = 'N/A'
        if ($sqlInfo['HasSQL'] -and (Get-Module -ListAvailable -Name SqlServer)) {
            # Try to count databases using the server instance from VM name or hostname
            # For simplicity, use the VM's internal name; in production, you'd connect via FQDN
            $databaseCount = Get-SqlDatabaseCount -ServerInstance $vmName -ErrorAction SilentlyContinue
            if ($null -eq $databaseCount) { $databaseCount = 'Unable to connect' }
        }
        
        $vmDetails += [PSCustomObject]@{
            'Subscription' = $subscriptionInfo.name
            'Resource Group' = $resourceGroup
            'VM Name' = $vmName
            'VM Size' = $vmSize
            'vCPU Count' = $vCPUCount
            'OS Type' = $osType
            'Publisher' = $vmFullInfo[2]
            'Offer' = $vmFullInfo[3]
            'Windows License' = $windowsLicense
            'Has SQL Server' = if ($sqlInfo['HasSQL']) { 'Yes' } else { 'No' }
            'SQL Version' = $sqlInfo['SQLVersion']
            'SQL Edition' = $sqlInfo['SQLEdition']
            'SQL License' = $sqlInfo['SQLLicense']
            'Database Count' = $databaseCount
            'SQL Enterprise Required' = if ($sqlInfo['HasSQL'] -and $sqlInfo['SQLEdition'] -eq 'Enterprise') { 'Yes' } else { 'No' }
            'Provisioning State' = $vmFullInfo[5]
            'Scan Date' = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
    }
    
    # Discover SQL servers beyond VM marketplace images
    Write-Host "`n=== Enhanced SQL Server Discovery ===" -ForegroundColor Cyan
    $vmsWithExtension = Get-VMsWithSQLExtension
    $azureSqlResources = Get-AzureSQLResources
    
    if ($vmDetails.Count -eq 0) {
        Write-Host "No VMs found in the subscription." -ForegroundColor Yellow
    }
    else {
        Write-Host "Found $($vmDetails.Count) VMs. Creating Excel report..." -ForegroundColor Yellow
        
        # Create Excel file
        $excelParams = @{
            Path            = $OutputPath
            WorksheetName   = "VMs"
            TableName       = "VMReport"
            AutoSize        = $true
            TableStyle      = "Light10"
            PassThru        = $true
        }
        
        $excel = $vmDetails | Export-Excel @excelParams
        
        # Format the Excel workbook
        $ws = $excel.Workbook.Worksheets["VMs"]
        
        # Add header formatting
        $headerRange = $ws.Cells["A1:P1"]
        $headerRange.Style.Font.Bold = $true
        $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(70, 120, 180))
        $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
        
        # Set column widths
        $ws.Column(1).Width = 20   # Subscription
        $ws.Column(2).Width = 20   # Resource Group
        $ws.Column(3).Width = 25   # VM Name
        $ws.Column(4).Width = 18   # VM Size
        $ws.Column(5).Width = 12   # vCPU Count
        $ws.Column(6).Width = 12   # OS Type
        $ws.Column(7).Width = 20   # Publisher
        $ws.Column(8).Width = 20   # Offer
        $ws.Column(9).Width = 18   # Windows License
        $ws.Column(10).Width = 14  # Has SQL Server
        $ws.Column(11).Width = 18  # SQL Version
        $ws.Column(12).Width = 16  # SQL Edition
        $ws.Column(13).Width = 18  # SQL License
        $ws.Column(14).Width = 18  # Provisioning State
        $ws.Column(15).Width = 20  # Scan Date
        
        # Add freeze panes
        $ws.View.FreezePanes(2, 1)

        # Add Yes/No dropdown on SQL Enterprise Required column (N) so reviewers can override
        $sqlErValidation = $ws.DataValidations.AddListValidation("N2:N1048576")
        $sqlErValidation.ShowErrorMessage = $false
        $sqlErValidation.ShowInputMessage = $true
        $sqlErValidation.PromptTitle = "SQL Enterprise Required"
        $sqlErValidation.Prompt = "Select Yes or No"
        $sqlErValidation.Formula.Values.Add("Yes")
        $sqlErValidation.Formula.Values.Add("No")
        
        # Add summary sheet
        $summaryLabel = if ($SQLServerOnly) { 'SQL Server VMs' } else { 'Total VMs' }
        $summaryData = @(
            [PSCustomObject]@{
                'Metric' = $summaryLabel
                'Value' = $vmDetails.Count
            },
            [PSCustomObject]@{
                'Metric' = 'Total vCPUs'
                'Value' = ($vmDetails | Measure-Object -Property 'vCPU Count' -Sum).Sum
            },
            [PSCustomObject]@{
                'Metric' = 'Windows VMs'
                'Value' = ($vmDetails | Where-Object { $_.'OS Type' -eq 'Windows' }).Count
            },
            [PSCustomObject]@{
                'Metric' = 'Linux VMs'
                'Value' = ($vmDetails | Where-Object { $_.'OS Type' -eq 'Linux' }).Count
            },
            [PSCustomObject]@{
                'Metric' = 'SQL Server 2025 Instances'
                'Value' = ($vmDetails | Where-Object { $_.'SQL Version' -like "*2025*" }).Count
            },
            [PSCustomObject]@{
                'Metric' = 'SQL Server 2022 Instances'
                'Value' = ($vmDetails | Where-Object { $_.'SQL Version' -like "*2022*" }).Count
            },
            [PSCustomObject]@{
                'Metric' = 'SQL Server 2019 Instances'
                'Value' = ($vmDetails | Where-Object { $_.'SQL Version' -like "*2019*" }).Count
            },
            [PSCustomObject]@{
                'Metric' = 'SQL Server Enterprise Edition'
                'Value' = ($vmDetails | Where-Object { $_.'SQL Edition' -eq 'Enterprise' }).Count
            },
            [PSCustomObject]@{
                'Metric' = 'SQL Enterprise Required (Yes)'
                'Value' = ($vmDetails | Where-Object { $_.'SQL Enterprise Required' -eq 'Yes' }).Count
            },
            [PSCustomObject]@{
                'Metric' = 'SQL Server Standard Edition'
                'Value' = ($vmDetails | Where-Object { $_.'SQL Edition' -eq 'Standard' }).Count
            },
            [PSCustomObject]@{
                'Metric' = 'SQL Server Developer Edition'
                'Value' = ($vmDetails | Where-Object { $_.'SQL Edition' -eq 'Developer' }).Count
            },
            [PSCustomObject]@{
                'Metric' = 'SQL Server Express Edition'
                'Value' = ($vmDetails | Where-Object { $_.'SQL Edition' -eq 'Express' }).Count
            },
            [PSCustomObject]@{
                'Metric' = 'Report Generated'
                'Value' = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
        )
        
        $summaryData | Export-Excel -ExcelPackage $excel -WorksheetName "Summary" -AutoSize -TableName "Summary" -TableStyle "Light10" -PassThru > $null
        
        $summarySh = $excel.Workbook.Worksheets["Summary"]
        $summaryHeaders = $summarySh.Cells["A1:B1"]
        $summaryHeaders.Style.Font.Bold = $true
        $summaryHeaders.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $summaryHeaders.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(70, 120, 180))
        $summaryHeaders.Style.Font.Color.SetColor([System.Drawing.Color]::White)
        
        # Add worksheet for VMs with SQL IaaS Agent Extension (custom/non-marketplace SQL)
        if ($vmsWithExtension) {
            $extCount = if ($vmsWithExtension -is [array]) { $vmsWithExtension.Count } else { 1 }
            if ($extCount -gt 0) {
                Write-Host "Adding $extCount SQL IaaS Agent Extension discoveries..." -ForegroundColor Gray
                $vmsWithExtension | Export-Excel -ExcelPackage $excel -WorksheetName "SQL IaaS Extensions" -AutoSize -TableName "SQLExtensions" -TableStyle "Light10" -PassThru > $null
                
                $extSh = $excel.Workbook.Worksheets["SQL IaaS Extensions"]
                $extHeaders = $extSh.Cells["A1:E1"]
                $extHeaders.Style.Font.Bold = $true
                $extHeaders.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $extHeaders.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(70, 120, 180))
                $extHeaders.Style.Font.Color.SetColor([System.Drawing.Color]::White)
            }
        }
        
        # Add worksheet for Azure SQL Database Servers and Managed Instances
        if ($azureSqlResources) {
            $sqlCount = if ($azureSqlResources -is [array]) { $azureSqlResources.Count } else { 1 }
            if ($sqlCount -gt 0) {
                Write-Host "Adding $sqlCount Azure SQL resource discoveries..." -ForegroundColor Gray
                $azureSqlResources | Export-Excel -ExcelPackage $excel -WorksheetName "Azure SQL Resources" -AutoSize -TableName "AzureSQLResources" -TableStyle "Light10" -PassThru > $null
                
                $sqlSh = $excel.Workbook.Worksheets["Azure SQL Resources"]
                $sqlHeaders = $sqlSh.Cells["A1:H1"]
                $sqlHeaders.Style.Font.Bold = $true
                $sqlHeaders.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $sqlHeaders.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(70, 120, 180))
                $sqlHeaders.Style.Font.Color.SetColor([System.Drawing.Color]::White)
            }
        }
        
        # Save and close
        $excel.Save()
        $excel.Dispose()
        
        Write-Host "Report generated successfully!" -ForegroundColor Green
        Write-Host "Output file: $OutputPath" -ForegroundColor Cyan
    }
}
catch {
    Write-Error "Error generating report: $_"
    exit 1
}
