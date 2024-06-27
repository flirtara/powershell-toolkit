#================================================
# global variables
#================================================
$downloadsPath = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path
Write-Host $downloadsPath
#$downloadFolder = "C:\Users\v-kreddick\Downloads"
cd -Path $downloadsPath

#================================================
# search for screenshot zip and process 
#================================================
function ScreenShots {
    param (
        $msoaidDir 
    )
    write-Host " "
    Write-Host "Running Function Screenshots.."
    cd -Path $msoaidDir
    if (Test-Path "Screenshots.zip")
    { 
        write-host " "
        write-host "Screenshots File Exists."
        write-host " "
        Write-Host "Uncompressing: Screenshots.zip.."
        Expand-Archive -Path Screenshots.zip -Force
        if (Test-Path "Screenshots")
        { 
            write-host " "
            write-host "Screenshots Folder Exists."

            write-host " "
            write-host "Removing Screenshots.zip file"
            Remove-Item ScreenShots.zip

            write-host " "
            write-host "Unblocking Screenshots mht file"
            Unblock-File -Path *.mht      
        } 
        else 
        { 
            write-host " "
            write-host "Screenshots Folder DOES NOT Exist"
        }      
    } 
    else 
    { 
        write-host " "
        write-host "Screenshots.zip File DOES NOT Exist"
    }
    cd -Path $downloadsPath
}
#================================================
# search for error events
#================================================
function GetEvents {
    param (
        $msoaidDir 
    )
    write-Host " "
    Write-Host "Running Function GetEvents.."
    cd -Path $msoaidDir
    cd -Path "AADLogs"

    #================================================
    # get list of event log files
    #================================================
    $eventFiles = @(Get-ChildItem -File -Name -Filter *.evtx)

    Write-Host "----------------------------------------------------------"
    Write-Host " "

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
    cd -Path $downloadsPath
}

#================================================
# Main Program
#================================================

#================================================
# get list of msoaid zip files
#================================================
$zipFiles = @(Get-ChildItem -File -Filter *MSOAID*.zip)

#================================================
# process each zip file
#================================================
ForEach ($file in $zipFiles) {
    $folderName = $file -replace '.zip'
    write-Host " "
    write-Host " "
    Write-Host "=================================================="
    Write-Host "Uncompressing: $($file.Name).."
    Write-Host "=================================================="
    Expand-Archive -Path $file -Force
    
    #================================================
    # verify zip was decompressed
    #================================================
    write-host " "
    Write-Host "Verifying $folderName.."
    if (Test-Path $folderName)
    { 
        write-host " "
        write-host "Path exists."
        #Remove-Item $file
        #================================================
        # see if screenshots zip exists
        #================================================
        ScreenShots $foldername
        #================================================
        # pull errors from event logs
        #================================================
        Write-Host "Got Here!"
        GetEvents $foldername
    } 
    else 
    { 
        write-host " "
        write-host "Path doesn't exist!"
    }
}
