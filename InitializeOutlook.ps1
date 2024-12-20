# InitializeOutlook.ps1: Initializes the Outlook COM object

try {
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    Write-Host "Outlook initialized successfully." -ForegroundColor Green
} catch {
    Write-Host "Failed to initialize Outlook. Ensure Outlook is installed and configured." -ForegroundColor Red
    exit
}
