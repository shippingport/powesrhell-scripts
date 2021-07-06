<#
.Synopsis
   Removes all distribution group memberships for a specified user.
   2021 Shippingport
   .INPUTS
    -Identity <upn>
.EXAMPLE
   ./Remove-UserFromAllGroupMemberships.ps1
.EXAMPLE
    ./Remove-UserFromAllGroupMemberships.ps1 -Identity someone@somewhere.com
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$Identity
)

# Check if an EXO session is already in place, and if not, connect to EXO.
if (!(((Get-PSSession | Select-Object -Property State, Name) -like '@{State=Opened; Name=ExchangeOnlineInternalSession}').Count -gt 0)) {
    Connect-ExchangeOnline
}

Function Remove-UserFromAllGroupMemberships {
   
    if(!(Get-Mailbox -Identity $Identity)){
        # User not found.
        Break
    }

    # Get the DistinguishedName for the selected user
    $userDistinguishedName = (Get-Mailbox -Identity $Identity) | Select-Object DistinguishedName
    $distributionGroups= Get-DistributionGroup -ResultSize unlimited
    
    # Print the user passed for verification purposes
    Write-Host "User:" (Get-Mailbox -Identity $Identity) | Select-Object Name
    
    # Loop through all DGs
    foreach($distributionGroup in $distributionGroups) {
    
        # Add all DG members to an array
        $DistributionGroupMembers = Get-DistributionGroupMember -Identity $distributionGroup.Name
    
        foreach($distributionGroupMember in $DistributionGroupMembers){
            # See if our using is present in the DG array
            if($distributionGroupMember.DistinguishedName -eq $userDistinguishedName.DistinguishedName) {
                # User is present, remove user
                Remove-DistributionGroupMember $distributionGroup.Name -Member $Identity -confirm:$false
                Write-host -ForegroundColor Red "User removed from:" $distributionGroup.DisplayName
            }
        }
    }
}

Remove-UserFromAllGroupMemberships($Identity)
