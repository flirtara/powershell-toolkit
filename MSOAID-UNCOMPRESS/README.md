# MSOAID Zip Processor

This PowerShell script processes MSOAID zip files in the Downloads folder. It uncompresses the zip files, checks for a `Screenshots.zip` file, and processes it if found.

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## Installation

1. **Download the Script**: Save the script to your local machine.
2. **Open PowerShell**: Open PowerShell with administrative privileges.
3. **Navigate to Script Location**: Use `cd` to navigate to the directory where the script is saved.

## Usage

1. **Run the Script**: Execute the script by running `.\scriptname.ps1` in PowerShell.
2. **Process Zip Files**: The script will automatically:
   - Navigate to the Downloads folder.
   - Identify all zip files containing `MSOAID` in their names.
   - Uncompress each zip file.
   - Check for a `Screenshots.zip` file within each uncompressed folder.
   - Uncompress `Screenshots.zip` if it exists and unblock any `.mht` files.

### Example

```powershell
.\MSOAIDZipProcessor.ps1
```

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request.

## License

This project is licensed under the MIT License.

Feel free to customize this README file further to better suit your project's needs. If you have any additional details or sections you'd like to include, let me know!
