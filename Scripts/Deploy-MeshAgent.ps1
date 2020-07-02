<#
.Synopsis
   Installs MeshCentral agent using an Active Directory security group.
.INPUTS
    -deployTo <array or string>
.EXAMPLE
   ./Deploy-MeshAgent.ps1
.EXAMPLE
    ./Deploy-MeshAgent.ps1 -deployTo PC01,PC02 -credential JohnDoe
#>

Param (
    [PSObject] $deployTo,
    [System.Management.Automation.PSCredential] $credential
)

$script:credential = [PSCredential]
$script:MeshTargets = "MeshAgent_install_group"
$script:installSource = "\\softwaredepot\support$\Meshcentral"
$script:installerFile = "meshagent.exe"

Function MeshInstall()
{ 
    Clear-Host

    Write-Host
    Write-Host -ForegroundColor Green "   MeshCentral Agent deployment tool"
    Write-Host -ForegroundColor Green "   2020 Shippingport"
    Write-Host

    if(!$credential.IsPresent)
    {
        if(!(Test-Path Variable::global:credential) -or ($script:credential -isnot [PSCredential]))
        {
            Set-Variable -Name credential -Value (Get-Credential -Message "Enter credentials to use during deployment:") -Scope Script 
        }
        $credential = $script:credential
    }

    if($deployTo)
    {
        $targetArray = $deployTo

        Write-Host "Deploying MeshCentral agent to named targets:"
        foreach($target in $targetArray)
        {
            if($computer = Get-ADComputer $target -ErrorAction SilentlyContinue)
            {
                Write-Host -ForegroundColor Green "-" $computer.Name
            }
        }

        foreach($target in $targetArray)
        {
            try
            {
                if($computer = Get-ADComputer $target -ErrorAction SilentlyContinue)
                {                  
                    $match = ($computer.objectClass)
                    if($match -eq 'computer')
                    {
                        SingleTargetedInstall($computer)
                    }
                    else
                    {
                        Write-Host -ForegroundColor Red "Non-computer object: " $computer.Name "- Cannot deploy to this target."
                    }
                }
                else
                {
                    Write-Host -ForegroundColor Red "No targets given for deployTo parameter. Stopping execution."
                    Exit
                }
            }
            catch
            {
                Write-Host $_.Exception.Message
            }
        }
    }
    else
    {
        # Install with AD group
        Write-Host "Going for AD targeted install using security group $MeshTargets."
        $count = $null
        $nonComputerCount = $null
        $installTargets = Get-ADGroupMember -Identity $MeshTargets
        foreach($target in $installTargets)
        {
            # Count the number or computer objects in security group. Non-computer objects will soft fail during deployment.
            $match = ($target.objectClass)
            if($match -eq 'computer')
            {
                $count++
            }
            else
            {
                $nonComputerCount++
            }
        }
        
        if($nonComputerCount -eq !$null)
        {
            Write-Host -ForegroundColor Yellow "Warning: $nonComputerCount non-computer object(s) in security group $MeshTargets. These objects may generate warnings during deployment but shouldn't cause the deployment to fail."
        }
        elseif(!$count)
        {
            # Count was null, nothing to do
            Write-Host -ForegroundColor Red "Couldn't find any computer objects in security group $MeshTargets. Quitting..."
            Write-Host
            Exit
        }

        Write-Host -ForegroundColor Green "$count target(s) selected."
        Read-Host "Press ENTER to continue or control-c to abort"
        ADTargetedInstall

    }
}

Function ADTargetedInstall()
{ 
    $installTargets = Get-ADGroupMember -Identity $MeshTargets
    Write-Host "Starting installation on" $count "target(s)."

    foreach($target in $installTargets)
    {
        if(Test-Connection -ComputerName $target.Name -Count 1 -Quiet)
        {
            Write-Host -ForegroundColor Green $target.Name "is online!"

            $match = ($target.objectClass)
            if($match -eq 'computer')
            {
                Invoke-Command -ComputerName $target.Name -ScriptBlock ${function:Deploy} -ArgumentList $target -Credential $credential -Authentication CredSSP
            }
        }
        else
        {
                Write-Host -ForegroundColor Yellow $target.Name "is offline"
        }
    }
}

Function SingleTargetedInstall($targets)
{
    foreach($target in $targets)
    {
        Invoke-Command -ComputerName $target.Name -ScriptBlock ${function:Deploy} -ArgumentList $target -Credential $credential -Authentication CredSSP
    }
}
Function Deploy($target){
    
    $previousExecutionPolicy = Get-ExecutionPolicy
    Write-Host "Setting execution policy to Bypass temporarily... " -NoNewline
    Set-ExecutionPolicy -ExecutionPolicy Bypass
    Write-Host -ForegroundColor Green "[OK]"

    Write-Host "Starting deployment on host" $target.Name

    New-Item -ItemType "directory" -Path C:\temp_meshcentral -Force | Out-Null
    Copy-Item -Path $installSource\$installerFile -Destination C:\temp_meshcentral -Force | Out-Null

    $installer = "C:\temp_meshcentral\" + $installerFile
    $parameters = "-fullinstall"

    & $installer $parameters

    Remove-Item -Path C:\temp_meshcentral -Recurse | Out-Null

    Write-Host "Reverting execution policy to Default... " -NoNewline
    Set-ExecutionPolicy -ExecutionPolicy $previousExecutionPolicy | Out-Null
    Write-Host -ForegroundColor Green "[OK]"
}

MeshInstall
