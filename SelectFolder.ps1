# SelectFolder.ps1: Handles folder selection in Outlook

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
