<#
.Synopsis
    Downloads the newest file from an SFTP server and stores it in a local path.

    2022 Shippingport

.EXAMPLE
   ./sftp-downloadscript.ps1
#>

# Define some parameters for use later
Param(
    [string] $script:SFTP_KeyFilePath = "..\secure\.ssh\keys\key_for_winscp_use.ppk",
    [string] $script:SFTP_HostFingerprint = "<fingerprint>",
    [string] $script:SFTP_Host = "some.ftp.server.com",
    [string] $script:SFTP_User = "remoteftpuser",
    [string] $script:SFTP_Remote_Path = "/usr/share/files",
    [string] $script:LOCAL_Storage_Path = "C:\Users\someone\Files\",

    # If using email alerts, enter the recipient email address here
    [string] $script:EMAIL_To = "someone@somewhere.com",
    
    # Where is our WinSCP library located?
    [string] $script:DLL_LibraryPath = "$PSScriptRoot\WinSCPnet.dll"
)

# Dot source our secure email parameters
# This file should contain securely hashed connection information, such as the sender address and password.
# Required only when using Send-MailMessage
. ..\readonly\secure-email-parameters.ps1

Function preflight()
{
    # Check if we can load the WinSCP .NET library
    Write-Host $DLL_LibraryPath
    if (Test-Path -path $DLL_LibraryPath) {
        Write-Host "WinSCP .NET library exists." -ForegroundColor Green
        Add-Type -Path $DLL_LibraryPath
    } 
    else {
        Write-Host "WinSCP .NET library not found!" -ForegroundColor Red
        Write-Host "Download it from here: https://winscp.net/eng/downloads.php"
        Write-Host "Or check if the DLL_LibraryPath parameter is set correctly."
        break
    }
    Write-Host "Preflight checks done!" -ForegroundColor Green

    # Optional: check if OneDrive is running
    # if((Get-Process "OneDrive" -ea SilentlyContinue) -eq $Null){ 
    #     echo "OneDrive not running. Attempting to start OneDrive..." 
    #     Start-Process "C:\Program Files\Microsoft OneDrive\onedrive.exe" -ArgumentList "/background" > OneDrive_startfile.txt -ErrorAction SilentlyContinue
    # }
}

Function Download-SFTPFiles
{
        try
        {
            # Create a session
            $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
                Protocol = [WinSCP.Protocol]::Sftp
                HostName = $SFTP_Host
                UserName = $SFTP_User
                SshPrivateKeyPath = $SFTP_KeyFilePath
                SshHostKeyFingerprint = $SFTP_HostFingerprint
            }

            $session = New-Object WinSCP.Session

            try
            {
                $localPath = $LOCAL_Storage_Path
                $remotePath = $SFTP_Remote_Path

                # Connect
                $session.Open($sessionOptions)

                # Get list of files in the directory
                $directoryInfo = $session.ListDirectory($remotePath)
            
                # Select the most recent file
                $latest =
                    $directoryInfo.Files |
                    Where-Object { -Not $_.IsDirectory } |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 1

                # Any files at all?
                if (!$latest)
                {
                    Write-Host "Could not find downloadable files in $SFTP_Remote_Path."
                    $Subject = "No files found."
                    $Body = "No eligible files found for download. Script halted."
                    Send-MailMessage -From $PSMAILParam_From -To $To -Subject $Subject -Body $Body -Credential $PSMAILParam_Creds -SmtpServer 'smtp.office365.com' -Port 587 -UseSsl
    
                    exit 1
                }
            
                # Download the selected file
                $session.GetFileToDirectory($latest.FullName, $localPath)
                $Subject = "Downloaded files from remote server"
                $Body = "File $latest was succesfully downloaded and is now available for use in $LOCAL_Storage_Path\NewestFile.json."
                Send-MailMessage -From $PSMAILParam_From -To $To -Subject $Subject -Body $Body -Credential $PSMAILParam_Creds -SmtpServer 'smtp.office365.com' -Port 587 -UseSsl

            }    
            catch
            {
                Write-Host "Error: $($_.Exception.Message)"
                $Subject = "Error getting files from remote FTP server"
                $Body = "An error occurred while getting files from the remote FTP server:`n`n$($_.Exception.Message)"
                Send-MailMessage -From $PSMAILParam_From -To $To -Subject $Subject -Body $Body -Credential $PSMAILParam_Creds -SmtpServer 'smtp.office365.com' -Port 587 -UseSsl

                exit 1
            }
        # Disconnect, clean up
        $session.Dispose()
        # exit 0
    }
    catch [Exception]
    {
        Write-Host "Error: $($_.Exception.Message)"

        $Subject = "Error getting files from remote FTP server"
        $Body = "An error occurred while getting files from the remote FTP server:`n`n$($_.Exception.Message)"
        Send-MailMessage -From $PSMAILParam_From -To $To -Subject $Subject -Body $Body -Credential $PSMAILParam_Creds -SmtpServer 'smtp.office365.com' -Port 587 -UseSsl

        exit 1
    }
}

Function ArchiveAndDownloadFiles
{
    if(!(Test-Path -Path "$LOCAL_Storage_Path\NewestFile.json"))
    {
        # No files in directory yet, so this might be the first time running this script. Download the first file first.
        Download-SFTPFiles
    }
    else
    {
        # Rename, move and archive the old file before downloading the newest file.
        $dateCode = Get-Date -Format yyyy-MM-dd
        $archiveFilename = "Data_" + $dateCode + ".json"
        Move-Item -Path "$LOCAL_Storage_Path\NewestFile.json" -Destination "$LOCAL_Storage_Path\Archive\$archiveFilename"
        Download-SFTPFiles

        $fileToRename = (Get-ChildItem -Path $LOCAL_Storage_Path -Recurse -Filter "Data_*").FullName[0]
        Rename-Item -Path $fileToRename -NewName "NewestFile.json"
    }
}

preflight
ArchiveAndDownloadFiles
