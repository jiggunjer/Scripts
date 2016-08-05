#--Create task for checking repo X
#Input parameters
$gitpath        = path/to/git
$gitcommand     = ' commit -a .'
$checkerscript  = path/to/script.ps1
$repopath       = /path/to/repo

#Create Task
$command = "checkerscript " + "repopath" + "gitpath" + "gitcommand"
$Action = New-ScheduledTaskAction -Execute 'C:WindowsSystem32WindowsPowerShellv1.0powershell.exe' -Argument "-NonInteractive -NoLogo -NoProfile -File '$command'"
$Trigger = New-ScheduledTaskTrigger -Daily -At '3AM'
$Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings (New-ScheduledTaskSettingsSet)
$Task | Register-ScheduledTask -TaskName 'My PowerShell script'