![Author Badge](https://img.shields.io/badge/Author-Jgooch-1F4D37)
![Distribution Badge](https://img.shields.io/badge/Distribution-PowerShell-blue)
![Target Badge](https://img.shields.io/badge/Target-Windows-0078D7)

# Outlook Statistics Dashboard Generator

## Overview

The **Outlook Statistics Dashboard Generator** is a PowerShell script designed to analyze your Outlook mailbox and generate a professional Excel dashboard. It extracts data about the top email senders and receivers, creates a summary table, and generates a pie chart for the top 10 senders. The script opens the generated Excel file for immediate use.

---

## Features

### Core Functionality

- **Top Senders and Receivers:**
  - Analyzes your Outlook Inbox to find the most frequent email senders and receivers.

- **Excel Dashboard Generation:**
  - Creates a clean and structured Excel file listing the top senders and receivers.
  - Includes a pie chart visualizing the top 10 senders.

- **Automation:**
  - Automatically opens the Excel dashboard upon completion.

---

### Prerequisites

| Dependency       | Description                                     | Installation Command                         |
|-------------------|-------------------------------------------------|---------------------------------------------|
| **PowerShell 5.1** | Required for script execution                   | Pre-installed on most Windows systems       |
| **Microsoft Outlook** | Access to your email data                     | Ensure Outlook is installed and configured  |
| **ImportExcel Module** | Enables Excel file creation and charting    | `Install-Module -Name ImportExcel -Force`   |

---

## Installation and Usage

### Step 1: Clone the Repository

Download the script to your local machine:

```bash
git clone https://github.com/yourusername/OutlookDashboardGenerator.git
cd OutlookDashboardGenerator
