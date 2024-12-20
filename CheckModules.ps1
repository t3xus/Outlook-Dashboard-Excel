# CheckModules.ps1: Ensures all required modules are installed

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
