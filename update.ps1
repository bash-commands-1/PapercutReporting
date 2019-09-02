Start-Sleep -Seconds 3

#Download updated script
$dl = (New-Object System.Net.WebClient).Downloadstring('https://raw.githubusercontent.com/ctmatt/PapercutReporting/master/PapercutNotifications.ps1')

if ($dl -eq $null)
{
    #Download failed exiting updater
    exit
}


try 
{
    Remove-Item "$($PWD.Path)\PapercutNotifications.ps1"
    $dl | Out-File "$($PWD.Path)\PapercutNotifications.ps1"
    Start-Process PowerShell -Arg "$($PWD.Path)\PapercutNotifications.ps1" -notify
}
catch [System.Exception] {
    #Failed to update exiting
    exit
}
