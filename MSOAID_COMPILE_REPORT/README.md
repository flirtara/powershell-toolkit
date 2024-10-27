# MSOAID Log Data Analyzer

## Overview
The **MSOAID Log Data Analyzer** is a PowerShell script designed to analyze MSOAID log files located in the user's Downloads directory. It compiles reports from log files, specifically focusing on error entries, and generates a summary report for easy review.

## Features
- Automatically navigates to the Downloads directory.
- Displays a menu for selecting MSOAID data directories.
- Compiles reports from log files, including:
  - `dsregcmdDebugOutput.txt`
  - Excel files with `.xlsx` extension.
- Filters and displays error entries from the logs.
- Generates a report file summarizing the findings.

## Requirements
- PowerShell (version 5.1 or later)
- Import-Excel module (for handling Excel files)

## Usage
1. **Clone the repository**:
   ```bash
   git clone https://github.com/yourusername/MSOAID-Log-Analyzer.git
   cd MSOAID-Log-Analyzer
