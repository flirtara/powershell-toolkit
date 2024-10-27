#================================================
# global variables
#================================================
$downloadsPath = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path
write-host $downloadsPath
cd -Path $downloadsPath

#================================================
# search for screenshot zip and process 
#================================================
function ScreenShots {
    param (
        $msoaidDir 
    )
    write-Host " "
    Write-Host "Checking for Screenshots zip file..."
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
# convert event logs to csv
#================================================
function GetEvents {
    param (
        $msoaidDir 
    )
    write-host "msoaidDir:  $msoaidDir"
    write-host " "
    Write-host "Running Function GetEvents.."
    cd -Path $msoaidDir
    cd -Path "AADLogs"
    $extract_path = $downloadsPath + "\" + $msoaidDir + "\Data\"
    write-host "extract_path: $extract_path" 
   
    #================================================
    # check if data directory exists
    #================================================
    If(!(test-path -PathType container $extract_path))
    {
        New-Item -ItemType Directory -Path $extract_path
    }

    #================================================
    # get list of event log files
    #================================================
    $eventFiles = @(Get-ChildItem -File -Name -Filter *.evtx)
    $textFiles = @(Get-ChildItem -File -Name -Filter *.txt)

    write-host "----------------------------------------------------------"
    write-host "Moving Text Files from AADLogs folder..."
    
    ForEach ($file in $textFiles)
    {
        cp $file $extract_path
    }
    #================================================
    # process each event log file
    #================================================

    write-host "----------------------------------------------------------"
    write-host "Converting Event Logs from AADLogs folder into CSV Files..."
 
    #================================================
    # process each event log file
    #================================================
    ForEach ($file in $eventFiles)
    {
        write-host "-------------------------------------------------"
        #================================================
        # create output filename
        #================================================
        $outputFile = $file -replace '.evtx', '.xlsx'
        $source_path = $downloadsPath + "\" + $msoaidDir + "\AADLogs\" + $file
        $target_path = $extract_path + $outputFile
        write-host "source_path:  $source_path"
        write-host "target_path:  $target_path"
        
        #================================================
        # convert to excel
        #================================================
        $events = Get-WinEvent -Path $source_path 
        $events | Select-Object 'TimeCreated', 'RecordId', 'MachineName', 'UserId', 'TaskDisplayName', 'LevelDisplayName', 'Message' | Export-Excel -Path $target_path  -AutoSize -WorksheetName "Sheet1" -FreezeTopRow -TableStyle Light1
    }
}
#================================================
# Main Program
#================================================

#================================================
# get list of msoaid zip files
#================================================
write-host " "
write-host "Getting list of MSOAID zip files"
$zipFiles = @(Get-ChildItem -File -Filter *MSOAID*.zip)

#================================================
# process each zip file
#================================================
ForEach ($file in $zipFiles) {
    $folderName = $file -replace '.zip'
    write-host " "
    write-host " "
    write-host "=================================================="
    write-host "Uncompressing: $($file.Name).."
    write-host "=================================================="
    Expand-Archive -Path $file -Force
    
    #================================================
    # verify zip was decompressed
    #================================================
    write-host " "
    write-host "Verifying $folderName was created..."
    if (Test-Path $folderName)
    { 
        write-host " "
        write-host "$folderName exists."
        write-host "removing zip file..."
        Remove-Item $file
        #================================================
        # see if screenshots zip exists
        #================================================
        ScreenShots $foldername
        #================================================
        # pull errors from event logs
        #================================================
        GetEvents $foldername
    } 
    else 
    { 
        write-host " "
        write-host "Path doesn't exist!"
    }
}
