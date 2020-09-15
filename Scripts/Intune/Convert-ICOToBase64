<#
.Synopsis
   Converts an ICO file to base64. Useful for embedding ICO resources in PowerShell scripts.
.INPUTS
    -file <file>
.EXAMPLE
   ./Convert-ToBase64.ps1
.EXAMPLE
    ./Convert-ToBase64.ps1 -file icon.ico
#>

Param (
    [string] $file

)

Function Get-File($initialDirectory)
{   
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Multiselect = $true
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "ICO-files |*.ICO;"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileNames
}

Function Convert-ToBase64($path)
{
    $FileToConvert = Get-Content -path $path -encoding Byte
    $base64 = [System.Convert]::ToBase64String($FileToConvert)
    $base64 | Out-file $env:HOMEPATH\Desktop\ico_base64.txt
}

if([string]::IsNullOrEmpty($file)) {
    $file = Get-File -initialDirectory "C:"
    Convert-ToBase64($file)
}
else {
    Convert-ToBase64($file)
}
