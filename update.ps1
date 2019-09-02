#Download updated script
$dl = (New-Object System.Net.WebClient).Downloadstring('https://raw.githubusercontent.com/curi0usJack/luckystrike/master/luckystrike.ps1')

if ($dl -eq $null)
{
    #Download failed exiting updater
    exit
}


try 
{
    Remove-Item "$($PWD.Path)\PapercutNotification.ps1"
    $dl | Out-File "$($PWD.Path)\PapercutNotification.ps1"
}
catch [System.Exception] {
    #Failed to update exitting
    exit
}