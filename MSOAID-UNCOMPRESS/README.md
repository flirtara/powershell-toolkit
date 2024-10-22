Here's a GitHub README file for your PowerShell script:

```markdown
# PowerShell Script

## Description
This PowerShell script performs the following tasks:
1.  Checks for MSOAID zip files in Download folder.
2.  If data folder doesn't exist in Download folder it creates it.
3.  If data folder exists delete contents of data folder.
4.  Searches for a `Screenshots.zip` file and if exists unzip and unblock file for viewing images.
5.  Searches for events log files in the `AADLogs` directory.
6.  Conversts all event log files to excel files with a subset of columns for data.
7.  Saves all new event log excel files in data directory.
8.  Copies all text files in AADLogs folder to data directory.

## Usage Example:
1. Launch Windows PowerShell as an Administrator, and wait for the PS> prompt to appear
2. Navigate within PowerShell to the directory where the script lives:
      PS> cd C:\my_path\yada_yada\ (enter)
3. Execute the script:
      PS> .\msoaid_uncompress.ps1 (enter)

Or: you can run the PowerShell script from the Command Prompt (cmd.exe) like this:

      powershell -noexit "& ""C:\my_path\yada_yada\msoaid_uncompress.ps1""" (enter)


## Functions

### 1. `ScreenShots`
- Searches for `Screenshots.zip` and processes it.
- If the `Screenshots` folder exists, it unblocks `.mht` files.
- Removes the `Screenshots.zip` file.

### 2. `GetEvents`
- Searches for error event log files in the `AADLogs` folder.
- Converts each event log file to an excel file.
- Only saves a subset of columns in the output file.

## Configuration Items you may want to change.
Path where MSOAID zip files are locate:  Line 4:  $downloadsPath = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path
Path where  text files are copies and excel files saved:  Line 7:  $extract_path = $downloadsPath + "\data\"
Columns that are pulled and saved in the Excel File:  Line 111: Select-Object 'TimeCreated', 'RecordId', 'MachineName', 'UserId', 'TaskDisplayName', 'LevelDisplayName', 'Message' 

## Requirements
- PowerShell 5.0 or higher

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
```

Feel free to customize the README file further to suit your needs! If you need any more help, just let me know.
