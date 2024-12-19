![Author Badge](https://img.shields.io/badge/Author-Jgooch-1F4D37)
![Distribution Badge](https://img.shields.io/badge/Distribution-PowerShell-blue)
![Target Badge](https://img.shields.io/badge/Target-Microsoft%20365%20API-0078D7)

# Outlook Statistics Dashboard Generator (API Version)

## Overview

The **Outlook Statistics Dashboard Generator (API Version)** uses the Microsoft Graph API to fetch email data from a user's Microsoft 365 account and generate a professional Excel dashboard. It includes advanced metrics like top senders, receivers, and attachments, displayed with visualizations like pie charts. This API-based approach supports remote execution without relying on the local Outlook application.

---

## Features

### Core Functionality

- **Top Senders and Receivers:**
  - Fetches the most frequent email senders and receivers from your Microsoft 365 mailbox.

- **Excel Dashboard Generation:**
  - Creates a structured Excel file with detailed email statistics.
  - Includes visualizations like pie charts for the top 10 senders.

- **Remote Compatibility:**
  - Runs independently of the local Outlook installation by leveraging the Microsoft Graph API.

### Advanced Metrics

- **Email Trends:**
  - Analyze email activity over time (e.g., daily, weekly).

- **Attachment Insights:**
  - Track the number and size of attachments sent and received.

- **Customizable Filters:**
  - Filter email statistics by date ranges or subject keywords.

---

## Prerequisites

### Microsoft 365 API Setup

1. **Register an Application in Azure Active Directory:**
   - Go to [Azure Portal](https://portal.azure.com).
   - Navigate to `Azure Active Directory > App Registrations`.
   - Click `New Registration` and provide a name (e.g., "OutlookDashboardApp").
   - Set the `Redirect URI` to `http://localhost` and click `Register`.

2. **Configure API Permissions:**
   - In the `API Permissions` section of the registered app:
     - Add `Mail.Read` (delegated) permission.
     - Grant admin consent for the permissions.

3. **Generate a Client Secret:**
   - Navigate to `Certificates & Secrets`.
   - Click `New Client Secret`, provide a description, and set an expiration period.
   - Save the generated client secret securely.

4. **Obtain the Following Details:**
   - **Tenant ID:** Found in the Azure Active Directory overview.
   - **Client ID:** Found in the app's overview.
   - **Client Secret:** Generated in the previous step.

### PowerShell Environment

| Dependency       | Description                                     | Installation Command                         |
|-------------------|-------------------------------------------------|---------------------------------------------|
| **PowerShell 5.1+** | Required for script execution                   | Pre-installed on most Windows systems       |
| **ImportExcel Module** | Enables Excel file creation and charting    | `Install-Module -Name ImportExcel -Force`   |

---

## Installation and Usage

### Step 1: Clone the Repository

Download the script to your local machine:

```bash
git clone https://github.com/yourusername/OutlookDashboardGenerator-API.git
cd OutlookDashboardGenerator-API
```

### Step 2: Configure the Script

1. Open the script file (`OutlookDashboardGenerator_API.ps1`).
2. Replace the placeholders with your Azure app details:
   - **Tenant ID**: Replace `YOUR_TENANT_ID`.
   - **Client ID**: Replace `YOUR_CLIENT_ID`.
   - **Client Secret**: Replace `YOUR_CLIENT_SECRET`.

### Step 3: Run the Script

Execute the script in PowerShell:

```powershell
.\OutlookDashboardGenerator_API.ps1
```

### Step 4: View Results

1. The script fetches email data using the Microsoft Graph API.
2. An Excel file named `OutlookDashboard_API.xlsx` is saved to your desktop.
3. The file opens automatically upon completion.

---

## Output Example

### Excel Dashboard

- **Worksheet 1:** `Top Senders`
  - Lists the top 10 senders with email counts.

- **Worksheet 2:** `Top Receivers`
  - Lists the top 10 receivers with email counts.

- **Worksheet 3:** `Top Attachments`
  - Lists the top 10 senders of attachments and their sizes.

- **Pie Chart:**
  - A 3D exploded pie chart visualizing the top 10 senders.

---

## Script Logic

1. **Access Token Retrieval:**
   - The script authenticates using the client ID, tenant ID, and client secret to obtain an access token.

2. **Email Data Fetching:**
   - Fetches email data from the `/me/messages` endpoint of the Microsoft Graph API.

3. **Data Preparation:**
   - Aggregates statistics for senders, receivers, and attachments.

4. **Excel Dashboard Creation:**
   - Exports the aggregated data into structured worksheets.
   - Adds visualizations like pie charts.

5. **File Handling:**
   - Saves the Excel file to the user's desktop and opens it automatically.

---

## Troubleshooting

| Issue                         | Solution                                                                 |
|-------------------------------|-------------------------------------------------------------------------|
| **Invalid Credentials**       | Verify the tenant ID, client ID, and client secret in the script.       |
| **Permission Denied**         | Ensure admin consent is granted for the app's API permissions.          |
| **ImportExcel Not Installed** | Run `Install-Module -Name ImportExcel -Force` in PowerShell.            |
| **Script Permissions**        | Run the script with administrative privileges if required.              |

---

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

