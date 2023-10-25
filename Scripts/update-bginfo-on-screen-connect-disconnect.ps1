<#
.Synopsis
   Script used for creating a scheduled task that runs when a monitor is connected or disconnected. This ensures the desktop wallpaper set using BgInfo is correctly positioned/arranged when the resolution or DPI changes.

   The Xpath filters used are:

    *[System[(EventID=6416)]] and *[EventData[Data[@Name='ClassName'] and (Data='Monitor')]]
    
    This listens for event ID 6416 in the 'Microsoft Windows security auditing' event log, but only when the registered device has a ClassName of 'monitor'. This triggers when a new monitor is connected.
    The other filter used is:

    *[System[EventID=1010]]

    For the Windows-Kernel-PnP/Device Management log. This last filter is inefficient because it will trigger on any PnP device removal, but the implementation of Xpath in Task Scheduler doesn't implement the contains() method. This makes it nigh on impossible to filter DeviceInstanceId in event ID 1010 for "DISPLAY", which is how it should ideally work.
    Unless you're wildly unplugging devices all the time this shouldn't really matter /that/ much. You could also set DisallowStartIfOnBatteries to true if this really bothers you on laptops.

    This task uses advanced security auditing, which is only available on Pro and Enterprise versions of Windows.

   2023 Christian Wedema
.EXAMPLE
   ./Importer.ps1
#>

[string]$Task = @'
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2023-09-27T11:33:13.5219985</Date>
    <Author>AzureAD\ChristianWedema</Author>
    <URI>MEM\Trigger BgInfo screen update</URI>
  </RegistrationInfo>
  <Triggers>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="Microsoft-Windows-Kernel-PnP/Device Management"&gt;&lt;Select Path="Microsoft-Windows-Kernel-PnP/Device Management"&gt;*[System[EventID=1010]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
    </EventTrigger>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="Security"&gt;&lt;Select Path="Security"&gt;*[System[(EventID=6416)]] and *[EventData[Data[@Name='ClassName'] and (Data='Monitor')]]	
	&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
    </EventTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <GroupId>S-1-5-32-545</GroupId>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>StopExisting</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>C:\Management\bg\Bginfo64.exe</Command>
      <Arguments>C:\Management\bg\default.bgi /timer:0 /silent /nolicprompt</Arguments>
    </Exec>
  </Actions>
</Task>
'@

Register-ScheduledTask -Xml $Task -TaskName "Update BgInfo when new monitors are connected" -TaskPath "\MEM"
