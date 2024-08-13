<#
.Synopsis
    This function allows you to unpin apps from the taskbar. Run within user context.

    2024 Christian Wedema

.EXAMPLE
   Remove-TaskbarPin()
#>

Function Remove-TaskbarPin([System.__ComObject]$App) {
    $App.InvokeVerb('taskbarunpin')
}

# Load all apps from the shell location
$Apps = (New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items()

# Unpin an app, for example all apps matching HP*
foreach($App in $Apps) {
    if($App.Name -like "HP*") {
        Remove-TaskBarPin($App)
    }
}
