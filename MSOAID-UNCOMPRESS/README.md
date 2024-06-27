Here's a GitHub README file for your PowerShell script:

```markdown
# PowerShell Script

## Description
This PowerShell script performs the following tasks:
1. Searches for a `Screenshots.zip` file and processes it.
2. Searches for error events in the `AADLogs` folder.

## Usage
1. Run the script in a PowerShell environment.
2. Follow the prompts to narrow down the search by date (optional).
3. The script will create output files containing error logs.

## Functions

### 1. `ScreenShots`
- Searches for `Screenshots.zip` and processes it.
- If the `Screenshots` folder exists, it unblocks `.mht` files.
- Removes the `Screenshots.zip` file.

### 2. `GetEvents`
- Searches for error event log files in the `AADLogs` folder.
- Allows narrowing down the search by date (optional).
- Creates output files containing error logs.

## Usage Example
```powershell
# Set the directory where the screenshots and logs are located
$msoaidDir = "C:\Path\To\Directory"

# Run the ScreenShots function
ScreenShots -msoaidDir $msoaidDir

# Run the GetEvents function
GetEvents -msoaidDir $msoaidDir
```

## Requirements
- PowerShell 5.0 or higher
- Administrative privileges to access event logs

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
```

Feel free to customize the README file further to suit your needs! If you need any more help, just let me know.