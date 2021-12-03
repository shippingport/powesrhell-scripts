<#
.Synopsis
    Downloads and installs printer drivers from a web location, then creates and installs printers.
    Especially useful when deploying printers without follow-me printing using Intune

    2021 Shippingport

.EXAMPLE
   ./Install-PrinterUsingRemoteDriver.ps1
#>
Function downloadAndInstallDriver
{
    if(!(Get-Path -Path "C:\temp\printers"))
    {
        # Create a download directory
        New-Item -ItemType "directory" -Path "C:\temp\printers"

        if(!(Get-Path -Path "C:\temp\drivers\pcl6"))
        {
            # Create the directory we'll expand the archive into
            New-Item -ItemType Directory -Path "C:\temp\drivers\pcl6"
        }
    }

    # Download the zip containing our drivers
    $URL = "https://somewhereontheinterwebs.com/print/drivers/sharp-pcl6.zip"
    $Output = "C:\temp\printers\sharp-pcl6.zip"
    $webclient = New-Object System.Net.WebClient
    $webclient.DownloadFile($URL, $Output)

    # Extract zip archive
    Expand-Archive -Path C:\temp\printers\sharp-pcl6.zip -DestinationPath "C:\temp\drivers\pcl6"
                                            
    # Install the drivers to the DriverStore
    Get-ChildItem "C:\temp\drivers\pcl6" -Recurse -Filter "*inf" | ForEach-Object { PNPUtil.exe /add-driver $_.FullName /install }

    # Add drivers to Windows print subsystem
    Add-PrinterDriver -Name "SHARP MX-C303W PCL6" # Or some other model, get name from the inf file
    Add-PrinterDriver -Name "SHARP MX-2651 PCL6"

}

Function installPrinters
{
    # Add printer ports
    Add-PrinterPort -Name "IP_192.168.214.2" -PrinterHostAddress "192.168.214.2"
    Add-PrinterPort -Name "IP_10.38.25.2" -PrinterHostAddress "10.38.25.2"

    # Grote printer, MX-C303W
    Add-Printer -DriverName "SHARP MX-C303W PCL6" -Name "Logistics printer" -PortName "IP_192.168.214.2"

    # Kleine printer, MX-2651
    Add-Printer -DriverName "SHARP MX-2651 PCL6" -Name "Jim's personal printer" -PortName "IP_10.38.25.2"
}


# Check if the drivers are already installed using the driver's GUID. You can get this GUID from your printer's driver file
if(!(Get-WmiObject Win32_PnPSignedDriver | Where-Object {$_.ClassGUID -eq "{4D36E979-E325-11CE-BFC1-08002BE10318}"}))
{
    # Driver is not installed to DriverStore, install that first
    downloadAndInstallDriver
    # Driver should now be installed, continue creation process
    # Optionally check for driver again here, we might have errored out somewhere
    installPrinters
}
else
{
    # Driver is present in DriverStore, create printers
    # Might want to check for the printer objects here so the script can abort, but I'm feeling lazy
    installPrinters
}
