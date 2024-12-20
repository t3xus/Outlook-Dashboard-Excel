# File 2: Outlook_Statistics_Analysis.ps1

# Import parameters from preparation script
$params = Import-Clixml -Path "$env:USERPROFILE\Documents\OutlookStatsParams.xml"

# Validate parameters
if (!$params) {
    Write-Host "No parameters found. Ensure the preparation script was run successfully." -ForegroundColor Red
    exit
}

# Load required variables
$folderPath = $params.FolderPath
$startDate = $params.StartDate
$endDate = $params.EndDate

# Define Excel output file path
$outputFile = "$env:USERPROFILE\Desktop\OutlookDashboard.xlsx"

# Initialize Outlook Application
try {
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    $selectedFolder = $Namespace.Folders | Where-Object { $_.FullFolderPath -eq $folderPath }
    if (!$selectedFolder) {
        Write-Host "Unable to locate the selected folder: $folderPath" -ForegroundColor Red
        exit
    }
} catch {
    Write-Host "Failed to initialize Outlook. Ensure Outlook is installed and configured." -ForegroundColor Red
    exit
}

# Your existing analysis and Excel export logic would go here.

Write-Host "Analysis complete. Dashboard saved to $outputFile" -ForegroundColor Green
