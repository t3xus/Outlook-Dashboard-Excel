# File 1: Outlook_Statistics_Preparation.ps1

# Check for ImportExcel module and install if missing
if (!(Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing now..." -ForegroundColor Yellow
    try {
        Install-Module -Name ImportExcel -Force -Scope CurrentUser
        Import-Module -Name ImportExcel
        Write-Host "ImportExcel module installed successfully." -ForegroundColor Green
    } catch {
        Write-Host "Failed to install ImportExcel module. Please install it manually and try again." -ForegroundColor Red
        exit
    }
} else {
    Write-Host "ImportExcel module is already installed. Proceeding..." -ForegroundColor Green
}

# Initialize Outlook Application
try {
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    Write-Host "Outlook initialized successfully." -ForegroundColor Green
} catch {
    Write-Host "Failed to initialize Outlook. Ensure Outlook is installed and configured." -ForegroundColor Red
    exit
}

# Prompt user to select a folder
Write-Host "Available Outlook folders:" -ForegroundColor Yellow
$folders = @{}
$index = 0
foreach ($folder in $Namespace.Folders) {
    $index++
    $folders[$index] = $folder.Name
    Write-Host "$index: $($folder.Name)"
}
$selectedFolderIndex = Read-Host "Enter the number corresponding to the folder you want to analyze"
if (!$folders[$selectedFolderIndex]) {
    Write-Host "Invalid selection. Exiting." -ForegroundColor Red
    exit
}
$selectedFolder = $Namespace.Folders[$folders[$selectedFolderIndex]]

# Prompt user to enter date range
$startDate = Read-Host "Enter start date (MM/DD/YYYY)"
$endDate = Read-Host "Enter end date (MM/DD/YYYY)"

# Save selection and parameters for the next script
$params = @{
    FolderPath = $selectedFolder.FullFolderPath
    StartDate = $startDate
    EndDate = $endDate
}
$params | Export-Clixml -Path "$env:USERPROFILE\Documents\OutlookStatsParams.xml"

# Call the second script
Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$PSScriptRoot\Outlook_Statistics_Analysis.ps1`""

Write-Host "Preparation complete. Proceeding to analysis..." -ForegroundColor Green
