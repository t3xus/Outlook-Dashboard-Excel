![Author Badge](https://img.shields.io/badge/Author-Jgooch-1F4D37)
![Distribution Badge](https://img.shields.io/badge/Distribution-PowerShell-blue)
![Target Badge](https://img.shields.io/badge/Target-Windows-0078D7)

# Outlook Statistics Dashboard

## Overview

The **Outlook Statistics Dashboard** is a two-part PowerShell script suite designed to analyze email data from Microsoft Outlook and generate an Excel-based dashboard. The scripts provide insights into various aspects of email communication, including top senders, receivers, attachment statistics, and trends over time. It uses the `ImportExcel` module to create visually appealing Excel reports.

---

## Features

### Core Functionality

- **Top Senders and Receivers**:
  - Identify the most frequent email senders and receivers.

- **Email Trends**:
  - Analyze email activity trends over selected timeframes.

- **Attachment Statistics**:
  - Track the number and size of email attachments.

- **Interactive Folder Selection**:
  - Choose specific Outlook folders (e.g., Inbox, Sent Items) for analysis.

### Additional Features

- **Customizable Date Range**:
  - Analyze emails within specific start and end dates.

- **Color-Coded Excel Output**:
  - Conditional formatting for flagged, unread, or high-priority emails.

- **Multi-Folder Analysis**:
  - Support for analyzing multiple Outlook folders in one session.

- **Summary Report**:
  - Key metrics, including total emails processed, unique senders, and flagged emails.

---

## Prerequisites

### System Requirements

- Windows operating system with PowerShell 5.1 or higher.
- Microsoft Outlook installed and configured.

### PowerShell Modules

| Module       | Purpose                                      | Installation Command                         |
|--------------|----------------------------------------------|---------------------------------------------|
| **ImportExcel** | For Excel file creation and charting        | `Install-Module -Name ImportExcel -Force`   |

---

## Installation and Setup

1. **Clone or Download** the repository containing the script files.
   ```bash
   git clone https://github.com/yourusername/OutlookStatsDashboard.git
   cd OutlookStatsDashboard
   ```

2. **Run the Preparation Script**:
   - Launch PowerShell as Administrator.
   - Execute the `Outlook_Statistics_Preparation.ps1` script:
     ```powershell
     .\Outlook_Statistics_Preparation.ps1
     ```

3. **Follow Prompts**:
   - Select the desired Outlook folder.
   - Specify the date range for analysis.

4. **Analysis Script**:
   - The preparation script automatically calls `Outlook_Statistics_Analysis.ps1` to process emails and generate the dashboard.

---

## How the Scripts Work Together

### Files and Their Roles

- **`Outlook_Statistics_Preparation.ps1`**:
  - This is the main script that the user runs first. It handles all setup tasks, such as:
    - Ensuring required modules are installed.
    - Initializing the Outlook COM object.
    - Allowing the user to select a folder and date range for analysis.
    - Saving these selections and calling the analysis script automatically.

- **`Outlook_Statistics_Analysis.ps1`**:
  - This script is automatically triggered by the preparation script.
  - It reads the saved selections and performs detailed email analysis based on the folder and date range.
  - Processes email data and generates an Excel dashboard.

- **Function Scripts**:
  - Each function is modularized into its own file and dynamically loaded by the main scripts as needed:
    - `CheckModules.ps1`: Verifies and installs required PowerShell modules.
    - `InitializeOutlook.ps1`: Initializes the Outlook COM object for interaction.
    - `SelectFolder.ps1`: Handles folder selection in Outlook.
    - `SelectDateRange.ps1`: Prompts the user for start and end dates.

### Workflow

1. **User Executes `Outlook_Statistics_Preparation.ps1`**:
   - This script guides the user through selecting folders and date ranges.
   - Calls the necessary function scripts to handle each task.
   - Saves the user's inputs and calls `Outlook_Statistics_Analysis.ps1`.

2. **`Outlook_Statistics_Analysis.ps1`** Runs Automatically:
   - Reads saved inputs and processes email data from the selected folder.
   - Generates the Excel dashboard and saves it to the desktop.

3. **Excel Dashboard**:
   - Users can open the generated dashboard to view insights, charts, and statistics about their email activity.

---

## Output

### Excel Dashboard

- **Worksheets**:
  - `Top Senders`: Lists the top 10 email senders with email counts.
  - `Top Receivers`: Lists the top 10 email receivers with email counts.
  - `Attachment Statistics`: Lists attachment details by sender.
  - `Email Trends`: Line chart showing email activity over time.
  - `Summary`: Highlights key metrics like total emails processed.

- **Charts**:
  - Pie chart for top senders.
  - Line chart for email trends.

---

## Troubleshooting

| Issue                  | Solution                                                                 |
|------------------------|-------------------------------------------------------------------------|
| **Module Not Found**    | Ensure the `ImportExcel` module is installed. Use the provided command. |
| **Outlook Not Initialized** | Check that Microsoft Outlook is installed and configured properly.   |
| **Permission Issues**   | Run PowerShell as Administrator.                                       |
| **Empty Results**       | Verify the selected folder and date range contain emails.             |

---

## Contributing

Contributions are welcome! Please open an issue or submit a pull request on the repository.

---

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

---

## Contact

**Author**: James Christopher Gooch  
**Email**: jamescgooch@me.com  
**GitHub**: [Your GitHub Profile](https://github.com/yourusername)

