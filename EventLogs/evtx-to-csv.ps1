<#
.SYNOPSIS
   This script converts a Microsoft event log with file extension evtx to a csv file.

.DESCRIPTION
    This script just takes the source file and converts it to a csv file.

.PARAMETER Param1
    source:  Full path to a Microsoft event file with extension evtx.

.PARAMETER Param2
    target:  Full path to a csv file for output.

.EXAMPLE
      If you are in the directory with the PowerShell Script.
        .\evtx-to-csv.ps1 "C:\Users\v-kreddick\Downloads\application.evtx" "C:\Users\v-kreddick\Downloads\application.csv"
      
      If you are not in the directory where the PowerShell Script is located.
        C:\user\v-kreddick\myScripts\evtx-to-csv.ps1 "C:\Users\v-kreddick\Downloads\application.evtx" "C:\Users\v-kreddick\Downloads\application.csv"

#>


#================================================
# usage
#================================================
if (!$args) { 
   The source and target files are missing."
   
   Usage Example: C:\Users\v-kreddick\myapps\evtx-to-csv.ps1 "C:\Users\v-kreddick\Downloads\application.evtx" "C:\Users\v-kreddick\Downloads\application.csv"  
}

$source = $args[0]
$target = $args[1]
  
#================================================
# convert to csv
#================================================
$source | Select-Object 'TimeCreated', 'RecordId', 'MachineName', 'UserId', 'TaskDisplayName', 'LevelDisplayName', 'Message' | Export-Excel -Path $target -AutoSize -WorksheetName "Sheet1" -FreezeTopRow -TableStyle Light1
