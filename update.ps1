Start-Sleep -Seconds 3

#Download updated script
$dl = (New-Object System.Net.WebClient).Downloadstring('https://raw.githubusercontent.com/ctmatt/PapercutReporting/master/PapercutNotification.ps1')

if ($dl -eq $null)
{
    #Download failed exiting updater
    exit
}


try 
{
    Remove-Item "$($PWD.Path)\PapercutNotification.ps1"
    $dl | Out-File "$($PWD.Path)\PapercutNotification.ps1"
    Start-Process PowerShell -Arg "$($PWD.Path)\PapercutNotification.ps1" -notify
}
catch [System.Exception] {
    #Failed to update exiting
    exit
}
