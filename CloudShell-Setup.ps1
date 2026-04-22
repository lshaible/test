# CloudShell-Setup.ps1
# This script prepares and uploads scripts to Azure Cloud Shell
# Run this from your local machine to set up Cloud Shell for running reports

param(
    [Parameter(Mandatory = $true)]
    [string]$StorageAccountName,
    
    [Parameter(Mandatory = $true)]
    [string]$ResourceGroupName,
    
    [Parameter(Mandatory = $true)]
    [string]$SubscriptionId,
    
    [string]$FileShareName = "scripts",
    [string]$DirectoryName = "azure-reports"
)

Write-Host "===== Azure Cloud Shell Setup =====" -ForegroundColor Cyan
Write-Host "This script will upload report scripts to your Cloud Shell storage" -ForegroundColor Cyan

try {
    # Set subscription
    Write-Host "Setting subscription..." -ForegroundColor Yellow
    az account set --subscription $SubscriptionId
    
    # Get storage account key
    Write-Host "Retrieving storage account details..." -ForegroundColor Yellow
    $storageKey = az storage account keys list -n $StorageAccountName -g $ResourceGroupName --query "[0].value" -o tsv
    
    if (-not $storageKey) {
        Write-Error "Could not retrieve storage account key. Check account name, resource group, and subscription."
        exit 1
    }
    
    # Set environment variables for Azure CLI storage commands
    $env:AZURE_STORAGE_ACCOUNT = $StorageAccountName
    $env:AZURE_STORAGE_KEY = $storageKey
    
    # Create file share if it doesn't exist
    Write-Host "Checking/creating file share: $FileShareName" -ForegroundColor Yellow
    az storage share create --name $FileShareName --fail-on-exist 2>$null
    
    # Create directory in file share
    Write-Host "Creating directory: $DirectoryName" -ForegroundColor Yellow
    az storage directory create --share-name $FileShareName --name $DirectoryName 2>$null
    
    # Upload scripts
    $scriptFiles = @(
        "Generate-AzureVMReport.ps1",
        "Generate-AzureVMReport-Scheduled.ps1",
        "README.md",
        "Examples-AzureVMReport.ps1"
    )
    
    $scriptDir = Split-Path $MyInvocation.MyCommand.Path
    
    foreach ($file in $scriptFiles) {
        $filePath = Join-Path $scriptDir $file
        
        if (Test-Path $filePath) {
            Write-Host "Uploading: $file" -ForegroundColor Gray
            az storage file upload `
                --share-name $FileShareName `
                --source $filePath `
                --path "$DirectoryName/$file" `
                --force
            Write-Host "  ✓ Uploaded" -ForegroundColor Green
        }
        else {
            Write-Warning "File not found: $filePath"
        }
    }
    
    Write-Host "`n===== Setup Complete =====" -ForegroundColor Green
    Write-Host "Scripts uploaded to: \\$StorageAccountName.file.core.windows.net\$FileShareName\$DirectoryName" -ForegroundColor Cyan
    
    Write-Host "`nNext steps in Azure Cloud Shell:" -ForegroundColor Yellow
    Write-Host "1. Go to https://shell.azure.com" -ForegroundColor White
    Write-Host "2. Make sure you're using PowerShell environment" -ForegroundColor White
    Write-Host "3. Navigate to the scripts:" -ForegroundColor White
    Write-Host "   cd clouddrive/scripts/$DirectoryName" -ForegroundColor Gray
    Write-Host "4. Run the report script:" -ForegroundColor White
    Write-Host "   .\Generate-AzureVMReport.ps1 -Environment CloudShell" -ForegroundColor Gray
    
    # Create a connection string for Cloud Shell access
    $connectionString = "DefaultEndpointProtocol=https;AccountName=$StorageAccountName;AccountKey=$storageKey;EndpointSuffix=core.windows.net"
    Write-Host "`nFor reference, your Cloud Shell files are mounted at: /home/[username]/clouddrive" -ForegroundColor Cyan
}
catch {
    Write-Error "Error during Cloud Shell setup: $_"
    exit 1
}
finally {
    # Clear sensitive environment variables
    Remove-Item env:AZURE_STORAGE_KEY -ErrorAction SilentlyContinue
}
