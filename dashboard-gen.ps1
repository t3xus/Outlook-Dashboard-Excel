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

# Define the Excel output file path
$outputFile = "$env:USERPROFILE\Desktop\OutlookDashboard.xlsx"

# Initialize Outlook Application
try {
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    $Inbox = $Namespace.GetDefaultFolder(6) # 6 refers to Inbox
    Write-Host "Outlook initialized successfully." -ForegroundColor Green
} catch {
    Write-Host "Failed to initialize Outlook. Ensure Outlook is installed and configured." -ForegroundColor Red
    exit
}

# Collect email statistics
$senders = @{}
$receivers = @{}

# Process emails in the Inbox
Write-Host "Processing emails in the Inbox..." -ForegroundColor Yellow
foreach ($item in $Inbox.Items) {
    if ($item -is [Microsoft.Office.Interop.Outlook.MailItem]) {
        $sender = $item.SenderName
        $receiver = $item.To

        # Count sender occurrences
        if ($senders.ContainsKey($sender)) {
            $senders[$sender] += 1
        } else {
            $senders[$sender] = 1
        }

        # Count receiver occurrences
        $recipients = $receiver -split ";"
        foreach ($recipient in $recipients) {
            $trimRecipient = $recipient.Trim()
            if ($receivers.ContainsKey($trimRecipient)) {
                $receivers[$trimRecipient] += 1
            } else {
                $receivers[$trimRecipient] = 1
            }
        }
    }
}

# Prepare data for Excel
Write-Host "Preparing data for Excel dashboard..." -ForegroundColor Yellow
$topSenders = $senders.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 10
$topReceivers = $receivers.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 10

$sendersData = @()
foreach ($sender in $topSenders) {
    $sendersData += [PSCustomObject]@{
        Name   = $sender.Key
        Count  = $sender.Value
    }
}

$receiversData = @()
foreach ($receiver in $topReceivers) {
    $receiversData += [PSCustomObject]@{
        Name   = $receiver.Key
        Count  = $receiver.Value
    }
}

# Export to Excel
$sendersSheet = "Top Senders"
$receiversSheet = "Top Receivers"

$sendersData | Export-Excel -Path $outputFile -WorksheetName $sendersSheet -AutoSize -TableName Senders
$receiversData | Export-Excel -Path $outputFile -WorksheetName $receiversSheet -AutoSize -TableName Receivers

# Add a pie chart for top senders
Add-ExcelChart -Path $outputFile -WorksheetName $sendersSheet -ChartType PieExploded3D -Range "B1:B11" -Header "A1:A11" -Title "Top 10 Senders"

# Open the Excel file
Write-Host "Opening the Excel dashboard..." -ForegroundColor Green
Invoke-Item $outputFile
