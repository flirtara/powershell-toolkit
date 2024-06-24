#================================================
# global variables
#================================================
$downloadsPath = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path
#$downloadFolder = "C:\Users\v-kreddick\Downloads"
cd -Path $downloadsPath

#================================================
# get list of msoaid zip files
#================================================
$zipFiles = @(Get-ChildItem -File -Filter *MSOAID*.zip)

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
# Iterate through Files
#================================================
ForEach ($file in $zipFiles) {
    # Perform actions on each file
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
        Remove-Item $file
        #================================================
        # see if screenshots zip exists
        #================================================
        ScreenShots $foldername
    } 
    else 
    { 
        write-host " "
        write-host "Path doesn't exist!"
    }
}
