<#
.Synopsis
    Version 2.0 of Autopilot Assistant, a complete rewrite from version 1.x, now with a GUI!

    Version history
    1.0 - 26 april 2021 - First version
    1.1 - 24 november 2021 - Fixed some small issues, added AIK-enrollment test
    2.0 - 8 augustus 2022 - Rewrite with PowerShell GUI.

    Shippingport 2022

.EXAMPLE
    ./AutopilotAssistant.ps1
#>

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$Window = New-Object Windows.Forms.Form
$Window.text = "Autopilot Assistant"

# Change both to (600,300) if enabling the PictureBox element
$Window.minimumSize = New-Object System.Drawing.Size(345,300)
$Window.maximumSize = New-Object System.Drawing.Size(345,300)

# Add a close button
$CloseButton = New-Object Windows.Forms.Button
$CloseButton.Text = "Close"
$CloseButton.Location = New-Object System.Drawing.Size(15,225)
$CloseButton.Size = New-Object System.Drawing.Size(145,30)
$CloseButton.Add_Click({$Window.close()})
$Window.Controls.Add($CloseButton)

# Shutdown buttion
$PowerDownButton = New-Object Windows.Forms.Button
$PowerDownButton.Text = "Shutdown"
$PowerDownButton.Location = New-Object system.Drawing.Size(165,225)
$PowerDownButton.Size = New-Object System.Drawing.Size(70,30)
$PowerDownButton.Add_Click({Stop-Computer -ComputerName localhost})
$Window.Controls.Add($PowerDownButton)

# Restart button
$RestartButton = = New-Object Windows.Forms.Button
$RestartButton.Text = "Restart"
$RestartButton.Location = New-Object System.Drawing.Size(240,225)
$RestartButton.Size = New-Object System.Drawing.Size(70,30)
$RestartButton.Add_Click({Restart-Computer -ComputerName localhost})
$Window.Controls.Add($RestartButton)

# Add and connect to WiFi
$WifiAdd = New-Object Windows.Forms.Button
$WifiAdd.Text = "Import && connect to WiFi"
$WifiAdd.Location = New-Object System.Drawing.Size(15,15)
$WifiAdd.Size = New-Object System.Drawing.Size(300,30)
$WifiAdd.Add_Click({Start-Process powershell ".\Scripts\Connect-WiFi.ps1"})
$Window.Controls.Add($WifiAdd)

# RunWU
$RunWU = New-Object Windows.Forms.Button
$RunWU.Text = "Run Windows Update"
$RunWU.Location = New-Object system.Drawing.Size(15,55)
$RunWU.Size = New-Object System.Drawing.Size(300,30)
$RunWU.Add_Click({powershell ".\Scripts\Run-WUAgent.ps1"})
$Window.Controls.Add($RunWU)

# Add to production
$AddToTenantSterker = New-Object Windows.Forms.Button
$AddToTenantSterker.Text = "Add to production environment"
$AddToTenantSterker.Location = New-Object system.Drawing.Size(15,95)
$AddToTenantSterker.Size = New-Object System.Drawing.Size(300,30)
$AddToTenantSterker.Add_Click({powershell ".\Scripts\Add-ProdTenant.ps1"})
$Window.Controls.Add($AddToTenantSterker)

# Add to testing environment
$AddToTenantSTH = New-Object Windows.Forms.Button
$AddToTenantSTH.Text = "Add to testing environment"
$AddToTenantSTH.Location = New-Object system.Drawing.Size(15,135)
$AddToTenantSTH.Size = New-Object System.Drawing.Size(300,30)
$AddToTenantSTH.Add_Click({powershell ".\Scripts\Add-TestTenant.ps1"})
$Window.Controls.Add($AddToTenantSTH)


# Load company logo
# $Picture = (Get-Item (".\Resources\logo.png"))
# $img = [System.Drawing.Image]::Fromfile($Picture)
# $pictureBox = New-Object Windows.Forms.PictureBox
# $pictureBox.Width =  256
# $pictureBox.Height =  256
# $pictureBox.Image = $img
# $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
# $pictureBox.Anchor = [System.Windows.Forms.AnchorStyles]::Right
# $pictureBox.location = New-Object system.Drawing.Size(325,0)
# $Window.controls.add($pictureBox)

Function Test-ConnectionsForIntuneProvisioning {
    # verbindingstest van wifi connector
    $WiFiStatusLabel = New-Object System.Windows.Forms.Label

    if((Get-NetAdapter -Name Wi-Fi).Status -eq "UP")
    {    
        $WiFiStatusLabel.Text = "[OK] WiFi is connected!"
        $WiFiStatusLabel.ForeColor = "Green"   
    }
    elseif ((Get-NetAdapter -Name Wi-Fi).Status -ne "UP") {
        $WiFiStatusLabel.Text = "[ERROR] WiFi is NOT connected!"
        $WiFiStatusLabel.ForeColor = "Red" 
    }
    $WiFiStatusLabel.Location = New-Object system.Drawing.Size(15,170) #205
    $WiFiStatusLabel.Size = New-Object System.Drawing.Size(250,15)
    $Window.Controls.Add($WiFiStatusLabel)

    # verbindingstest van Intune servers
    $MEMPingTest = New-Object System.Windows.Forms.Label
    if(Test-Connection -ComputerName "endpoint.microsoft.com" -Count 1) # Ideally you should change this URL to point to your enterpriseenrollment CNAME record, e.g. enterpriseenrollment.company.com (make sure it responds to icmp, of use another way of comfirming operation)
    {    
        $MEMPingTest.Text = "[OK] Intune servers reachable!"
        $MEMPingTest.ForeColor = "Green"   
    }
    else {
        $MEMPingTest.Text = "[ERROR] Cannot reach Intune servers!"
        $MEMPingTest.ForeColor = "Red" 
    }

    $MEMPingTest.Location = New-Object system.Drawing.Size(15,185) #205
    $MEMPingTest.Size = New-Object System.Drawing.Size(250,15)
    $Window.Controls.Add($MEMPingTest)   
}

Function Test-AikCertificateEnrollment {

    # Check is we can get an AIK certificate, as this will be required to TPM attestation later on.
    # Intel Tiger Lake systems seem to have LOTS of issues with attestation, so checking this here may prevent headaches later on.
    # If the AIK certificate was not issued, make sure you're running the latest BIOS and TPM firmware versions. Microsoft seems
    # to be denying AIK enrollment on systems with known security issues in the BIOS of TPM firmware.

    $AikCertEnroll = New-Object System.Windows.Forms.Label

    $certreq = (certreq -enrollaik -config '""') # We only want the contents of automatic var LASTEXITCODE, so $certreq goes unused for now

    if(($LASTEXITCODE -eq 0))
    {    
        $AikCertEnroll.Text = "[OK] AIK certificate was issued!"
        $AikCertEnroll.ForeColor = "Green"   
    }
    elseif (($LASTEXITCODE -eq "-2145844848")) {
        $AikCertEnroll.Text = "[ERROR] Certreq: ERROR_WINHTTP_INVALID_URL"
        $AikCertEnroll.ForeColor = "Red" 
    }
    elseif (($LASTEXITCODE -eq "-2147024809")) {    # Make sure the double quotes are surrounded by singel qoutes, e.g. '""'
        $AikCertEnroll.Text = "[ERROR] Certreq: ERROR_INVALID_PARAMETER"
        $AikCertEnroll.ForeColor = "Red" 
    }
    elseif (($LASTEXITCODE -eq "-2147024891")) {    # This really shouldn't happen as we should be running as SYSTEM
        $AikCertEnroll.Text = "[ERROR] Certreq: ACCESS_DENIED"
        $AikCertEnroll.ForeColor = "Red" 
    }
    else {
        $AikCertEnroll.Text = "[ERROR] AIK certificaat was NOT issued ($LASTEXITCODE)"
        $AikCertEnroll.ForeColor = "Red"
    }
    $AikCertEnroll.Location = New-Object System.Drawing.Size(15,200) #205
    $AikCertEnroll.Size = New-Object System.Drawing.Size(300,20)
    $Window.Controls.Add($AikCertEnroll)
}

Test-ConnectionsForIntuneProvisioning
Test-AikCertificateEnrollment

$Window.ShowDialog()
