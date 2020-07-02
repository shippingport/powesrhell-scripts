<#
.Synopsis
   Quickly set UPN as primary SMTP address for AD users.
#>

$searchBase = "OU=Users,OU=Accounts,DC=ad,DC=domain,DC=com"

Function Start-UserTransformation()
{
    $tranformationUsers = Get-ADUser -filter * -searchbase $searchBase

    foreach ($user in $tranformationUsers)
    {
        if(!($user.proxyAddresses))
        {
            # Users with domain prefix 'ad' shouldn't have mailboxes, so don't add any proxyAddresses
            if(!($user.UserPrincipalName -match [regex]::new('(@ad\.)')))
            {
                $userProxyAddresses = ("SMTP:"+$user.UserPrincipalName)
                Set-ADUser -Identity $user -add @{proxyAddresses = $userProxyAddresses}
                Write-Host "Set" $userProxyAddresses "for user" $user.Name
            } 
        }
    }
}

Start-UserTransformation
