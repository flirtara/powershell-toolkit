#================================================
# get list of event log files
#================================================
$eventFiles = @(Get-ChildItem -File -Name -Filter *.evtx)

Write-Output "----------------------------------------------------------"

#================================================
# get date frame to pull errors for
#================================================
$getDate = Read-Host -Prompt "Do you want to narrow by date? [Y/n] "
if($getDate -eq "Y" -or $getDate -eq "y")
{
    Write-Host " "
    $startDate = Read-Host -Prompt "Enter Start Date [ 6/18/2024 ] "
    Write-Host " "
    $stopDate = Read-Host -Prompt "Enter Stop Date [ 6/18/2024 ] "
    Write-Host " "
}
#================================================
# process each file
#================================================
ForEach ($file in $eventFiles)
{
    Write-Host "-------------------------------------------------"
    #================================================
    # create output filename
    #================================================
    $outputFile = $file -replace '.evtx', '.txt'

    #================================================
    # pull errors and put in output file
    #================================================
    #Level (1=FATAL, 2=ERROR, 3=Warning, 4=Information, 5=DEBUG, 6=TRACE, 0=Info)

    Write-Host "Pulling Errors and Saving in $outputFile..."
    if ($StartDate)
    {
        try 
        { 
            Get-WinEvent -FilterHashTable @{ Path=$file;Level=2;StartTime=$startDate;EndTime=$stopDate } -ErrorAction Stop > $outputFile
        }
        catch 
        {
            if ($_.Exception -match "No events were found that match the specified selection criteria")
            {
                Write-Host "No events found";
            }
        }
    }
    else
    {
         try 
        { 
            Get-WinEvent -FilterHashTable @{ Path=$file;Level=2 } -ErrorAction Stop > $outputFile 
        }
        catch 
        {
            if ($_.Exception -match "No events were found that match the specified selection criteria.") 
            {
                Write-Host "No events found";
            }
        } 
    }
}
