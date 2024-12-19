# Outlook Statistics Dashboard Generator (API Version)

# Function to get an access token
Function Get-AccessToken {
    $tenantId = "YOUR_TENANT_ID"
    $clientId = "YOUR_CLIENT_ID"
    $clientSecret = "YOUR_CLIENT_SECRET"

    $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }

    $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body
    return $response.access_token
}

# Function to fetch email data
Function Get-EmailData {
    param ($accessToken)

    $url = "https://graph.microsoft.com/v1.0/me/messages"
    $headers = @{ Authorization = "Bearer $accessToken" }

    $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
    return $response.value
}

# Main script
Write-Host "Fetching access token..." -ForegroundColor Yellow
$accessToken = Get-AccessToken

Write-Host "Fetching email data from Microsoft Graph API..." -ForegroundColor Yellow
$emailData = Get-EmailData -accessToken $accessToken

# Collect statistics
$senders = @{}
$receivers = @{}

foreach ($email in $emailData) {
    $sender = $email.sender.emailAddress.address
    $receiver = $email.toRecipients | ForEach-Object { $_.emailAddress.address }

    # Count senders
    if ($senders.ContainsKey($sender)) {
        $senders[$sender] += 1
    } else {
        $senders[$sender] = 1
    }

    # Count receivers
    foreach ($recipient in $receiver) {
        if ($receivers.ContainsKey($recipient)) {
            $receivers[$recipient] += 1
        } else {
            $receivers[$recipient] = 1
        }
    }
}

# Prepare data for Excel
$topSenders = $senders.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 10
$sendersData = $topSenders | ForEach-Object { [PSCustomObject]@{ Name = $_.Key; Count = $_.Value } }

# Export to Excel
$outputFile = "$env:USERPROFILE\Desktop\OutlookDashboard_API.xlsx"
$sendersData | Export-Excel -Path $outputFile -WorksheetName "Top Senders" -AutoSize

Write-Host "Dashboard generated successfully. Opening Excel..." -ForegroundColor Green
Invoke-Item $outputFile
