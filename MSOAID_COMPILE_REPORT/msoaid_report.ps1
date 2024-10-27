#================================================
# global variables
#================================================
$downloadsPath = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path
cd -Path $downloadsPath

$data_list = New-Object System.Collections.Generic.List[System.Object]
$seperator_line = "-----------------------------------------------"
$blank_line = " "

#================================================
# Function Show Menu
#================================================
function Show-Menu
{
     param (
           [string]$Title = ‘MSOAID Log Data Available To Analyze’
     )
     cls
     write-host “================ $Title ================”
     $count = 0
     ForEach ($item in $msoaid_data_options)
     {
        $count += 1
        write-host “Press $count for: $item.”
     }
     Write-Host “Press Q: to quit.”
}
#================================================
# Function Compile Report
#================================================
function Compile-Report
{
   param (
          [string]$msoaid_data_dir = $data_to_process
    )
    cd -Path $msoaid_data_dir
    cd -Path "Data"
    $report_file = $msoaid_data_dir + "_report.txt"
    
    #================================================
    # Write Report Header
    #================================================
    "This Report Analyzes MSOAID Log Files from $msoaid_data_dir" > $report_file
    $blank_line >> $report_file
      
    #================================================
    # Get dsregcmdDebugOutput.txt file if exists
    #================================================
    if("dsregcmdDebugOutput.txt")
    {
        "Data From dsregcmdDebugOutput.txt file" >> $report_file
        $seperator_line >> $report_file
        # Append to an existing file
        cat "dsregcmdDebugOutput.txt" >> $report_file
        $blank_line >> $report_file
        $blank_line >> $report_file
    }
  
    #================================================
    # Get List Of Event Files with extension xlsx
    #================================================
    $eventFiles = @(Get-ChildItem -File -Name -Filter *.xlsx)
    ForEach ($file in $eventFiles)
    {
        write-host "Reading Event Log Entries For:  $file"
        "Reading Event Log Entries For:  $file" >> $report_file

        $data = Import-Excel -Path $file -WorksheetName Sheet1

        # Filter the data based on your conditions
        $filteredData = $data | Where-Object { $_.LevelDisplayName -eq "Error" }
        $filteredData

        # Get the last number of rows
        $lastRows = $filteredData | Select-Object -Last 3

        # Display the last 5 rows
        $seperator_line >> $report_file
        $lastRows >> $report_file
        $seperator_line >> $report_file
        $blank_line >> $report_file
        $blank_line >> $report_file
    }
}

#================================================
# Menu
#================================================
$msoaid_data_options = Get-ChildItem -LiteralPath $downloadsPath -Filter MSOAID* -Directory
if( $msoaid_data_options.count -eq 1)
{
    $data_to_process = $msoaid_data_options
    write-host " "
    write-host "Only 1 Option Available: $data_to_process"
    Compile-Report
}
else
{
    do
    {
        Show-Menu
        $input = Read-Host “Please make a selection”
        $data_to_process = $msoaid_data_options[$input]
        write-host "$data_to_process"
        Compile-Report
        pause
    }
    until ($input -eq ‘q’)
}