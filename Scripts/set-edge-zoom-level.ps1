<#
.Synopsis
    Sets the zoom leven for a specific website in Edge to a predefined level. Restart Edge or reload the profile to see changes.

    Set this script to run as USER in Intune.

    2024 github.com/shippingport
.EXAMPLE
   ./set-edge-zoom-level.ps1
#>

# Select hte Preferences file for the right Edge profile
$EdgePreferencesJsonFilePath = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Preferences"

# The JSON below sets the default zoom level for somewebsite.com to 90%
$JsonToAdd = @'
{"partition": {
        "per_host_zoom_levels": {
            "x": {
                "somewebsite.com": {
                     "last_modified": "13375791867649712",
                     "zoom_level": -0.5778829311823857
                }
            }
        }
    }
}
'@

$PreferencesJson = ConvertFrom-Json -InputObject (Get-Content $EdgePreferencesJsonFilePath).ToString() -AsHashtable

if($PreferencesJson.partition.per_host_zoom_levels.x.'somewebsite.com') {
    # The key for this website already exists, let's check if the zoom level was alreay set correctly or not.
    if($PreferencesJson.partition.per_host_zoom_levels.x.'somewebsite.com'.zoom_level -eq -0.5778829311823857) {
        # Zoom level was already set correctly, do nothing
        break
    } else {
        # Nope, need to set the zoom level.
        $PreferencesJson.partition.per_host_zoom_levels.x.'somewebsite.com'.zoom_level = -0.5778829311823857
    }
} else {
    # Key did not exist in the Preferences file, add it here.
    $JsonAddition = ConvertFrom-Json -InputObject $JsonToAdd

    foreach ($key in $JsonAddition.PSObject.Properties) {
        $PreferencesJson | Add-Member -MemberType NoteProperty -Name $key.Name -Value $key.Value -Force
    }
    # And finally save the object back to JSON.
    ConvertTo-Json $PreferencesJson -Depth 30 | Out-File -FilePath $EdgePreferencesJsonFilePath
}
