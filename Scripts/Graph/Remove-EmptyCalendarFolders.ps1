# This script removals all user calendars that have no content.
# First connect to Graph with the following scropes:
# Connect-MgGraph -Scopes "User.Read.All" , "Calendars.ReadWrite"

$user = Read-Host "Welke user?"
$agendaFolders = Get-MailboxFolderStatistics -Identity $user -FolderScope Calendar | Select-Object ItemsInFolderAndSubfolders,Name
$userId = Get-MgUser -UserId $user | Select-Object Id

foreach($agenda in $agendaFolders) {
    If($agenda.ItemsInFolderAndSubfolders -eq 0) {
        Write-Host "Lege map: " $agenda.Name
        $switch = Read-Host "Remove this calendar?"
        If($switch -eq "Y|y") {
            $AgendaGraphId = Get-MgUserCalendar -UserId $user | Where-Object -Propert Name -eq $agenda.Name | Select-Object Id
            $AgendaGraphId = $AgendaGraphId.Id
            $UserId = $UserId.Id
            $UserCalendarFolderUriConcat = "/v1.0/users/$UserId/calendars/$AgendaGraphId"

            # Write-Host $UserCalendarFolderUriConcat
            Invoke-MgGraphRequest -Method DELETE -Uri $UserCalendarFolderUriConcat
        }
    }
}
