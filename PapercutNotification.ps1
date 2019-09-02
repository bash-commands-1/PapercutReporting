param([bool]$notify = $true)

$version = 2.4

$config = Get-Content "$($PWD.Path)\config.xml"

$Company = $config.company.name
$PapercutAuthURL = $config.company.PapercutApiURL
$NotificationChannelID = $config.ct.notificationid
$CheckinChannelID = $config.ct.checkinid
$ReportUserLicense = $config.vars.reportuserlicense
$ReportSupportLicense = $config.vars.reportsupportlicense

$githubver = "https://raw.githubusercontent.com/ctmatt/PapercutReporting/master/versioncheck.txt?token=ANCDFFKSC33O4YNMS5HIJPK5NT2DM"
$updatefile = "https://raw.githubusercontent.com/ctmatt/PapercutReporting/master/update.ps1?token=ANCDFFPX6PEJYNCMDEDRRM25NT2H4"

$day = (Get-WmiObject Win32_LocalTime).day
$reportingserver = $env:computername



#Update from github

function UpdateCheck()
{
	$updateavailable = $false
	$nextversion = $null
	try
	{
		$nextversion = (New-Object System.Net.WebClient).DownloadString($githubver).Trim([Environment]::NewLine)
	}
	catch [System.Exception] 
	{
	}
		
	if ($nextversion -ne $null -and $version -ne $nextversion)
	{
		#An update is most likely available, but make sure
		$updateavailable = $false
		$curr = $version.Split('.')
		$next = $nextversion.Split('.')
		for($i=0; $i -le ($curr.Count -1); $i++)
		{
			if ([int]$next[$i] -gt [int]$curr[$i])
			{
				$updateavailable = $true
				break
			}
		}
	}
	UpdateScript
}

function UpdateScript()
{
	if (Test-Connection 8.8.8.8 -Count 1 -Quiet)
		{
			$updatepath = "$($PWD.Path)\update.ps1"
			if (Test-Path -Path $updatepath)	
			{
				Remove-Item $updatepath
			}
			
				(New-Object System.Net.Webclient).DownloadFile($updatefile, $updatepath)
				Start-Process PowerShell -Arg $updatepath
				exit
			}
}




#Grab variables from Papercut
Try
{
    $request = Invoke-WebRequest -Uri $PapercutAuthURL | ConvertFrom-Json
		$papercutversion = $request.applicationServer.systemInfo.version
		$remainingdays = $request.license.upgradeAssuranceRemainingDays
		$totalusers = $request.license.users.licensed
		$currentusers = $request.license.users.used
		$remainingusers = $request.license.users.remaining
}
Catch [System.Net.WebException]
{
    $Fact1 = New-TeamsFact -Name 'Remaining Licenses' -Value "**$remainingusers**"
    $CurrentDate = Get-Date
    $Section = New-TeamsSection `
        -ActivityTitle "Papercut Script Error" `
        -ActivitySubtitle "$CurrentDate" `
        -ActivityImage Add `
        -ActivityText "Unable to access Papercut API (Papercut might be down 😢). Reported by $reportingserver"
    Send-TeamsMessage `
        -URI $NotificationChannelID `
        -MessageTitle $Company `
        -MessageText "" `
        -Color Red `
        -Sections $Section
}
Catch
{
    
}

#Check in to teams (once a month)
if( ($day -le 7) -Or ($notify -eq $true) )
{
    write-host test
    $Fact1 = New-TeamsFact -Name 'Script Version' -Value "**$version**"
		$Fact2 = New-TeamsFact -Name 'Papercut Version' -Value "**$papercutversion**"
		$CurrentDate = Get-Date
    $Section = New-TeamsSection `
        -ActivityTitle "Check-In" `
        -ActivitySubtitle "$CurrentDate" `
        -ActivityImage Add `
        -ActivityText "Papercut for $Company is reporting in 👋 from $reportingserver"  `
        -ActivityDetails $Fact1, $Fact2
    Send-TeamsMessage `
        -URI $CheckinChannelID `
        -MessageTitle $Company `
        -MessageText "" `
        -Color Green `
        -Sections $Section
}

#Check remaining licenses
if( ($remainingusers -lt 1) -And ($ReportUserLicense -eq $true) -And ((get-date).DayOfWeek -eq "Monday") )
{
    $Fact1 = New-TeamsFact -Name 'Remaining Licenses' -Value "**$remainingusers**"
    $Fact2 = New-TeamsFact -Name 'Current Licenses' -Value "**$currentusers**"
    $Fact3 = New-TeamsFact -Name 'Total Licenses' -Value "**$totalusers**"
    $CurrentDate = Get-Date
    $Section = New-TeamsSection `
        -ActivityTitle "Papercut User Limit Notification" `
        -ActivitySubtitle "$CurrentDate" `
        -ActivityImage Add `
        -ActivityText "Papercut for $Company is completely depleted, there are $remainingusers licenses left. Reported by $reportingserver" `
        -ActivityDetails $Fact1, $Fact2, $Fact3
    Send-TeamsMessage `
        -URI $NotificationChannelID `
        -MessageTitle $Company `
        -MessageText "" `
        -Color Red `
        -Sections $Section
}
elseif( ($remainingusers -lt 10) -And ($ReportUserLicense -eq $true) -And ((get-date).DayOfWeek -eq "Monday") )
{
    $Fact1 = New-TeamsFact -Name 'Remaining Licenses' -Value "**$remainingusers**"
    $Fact2 = New-TeamsFact -Name 'Current Licenses' -Value "**$currentusers**"
    $Fact3 = New-TeamsFact -Name 'Total Licenses' -Value "**$totalusers**"
    $CurrentDate = Get-Date
    $Section = New-TeamsSection `
        -ActivityTitle "Papercut User Limit Notification" `
        -ActivitySubtitle "$CurrentDate" `
        -ActivityImage Add `
        -ActivityText "Papercut for $Company is low, there are only $remainingusers licenses left for users. Reported by $reportingserver" `
        -ActivityDetails $Fact1, $Fact2, $Fact3
    Send-TeamsMessage `
        -URI $NotificationChannelID `
        -MessageTitle $Company `
        -MessageText "" `
        -Color Orange `
        -Sections $Section

}

#Check remaining support days
if( ($remainingdays -lt 1) -And ($ReportSupportLicense -eq $true) -And ((get-date).DayOfWeek -eq "Monday") )
{
    $Fact1 = New-TeamsFact -Name 'Support Days Remaining' -Value "**$remainingdays**"
    $CurrentDate = Get-Date
    $Section = New-TeamsSection `
        -ActivityTitle "Papercut Support Expired Notification" `
        -ActivitySubtitle "$CurrentDate" `
        -ActivityImage Add `
        -ActivityText "Papercut Support for $Company has expired by $remainingdays days. Reported by $reportingserver" `
        -ActivityDetails $Fact1
    Send-TeamsMessage `
        -URI $NotificationChannelID `
        -MessageTitle $Company `
        -MessageText "" `
        -Color Red `
        -Sections $Section
}
elseif( ($remainingdays -lt 45) -And ($ReportSupportLicense -eq $true) -And ((get-date).DayOfWeek -eq "Monday") )
{
    $Fact1 = New-TeamsFact -Name 'Support Days Remaining' -Value "**$remainingdays**"
    $CurrentDate = Get-Date
    $Section = New-TeamsSection `
        -ActivityTitle "Papercut Support Expiring Notification" `
        -ActivitySubtitle "$CurrentDate" `
        -ActivityImage Add `
        -ActivityText "Papercut Support for $Company is low, there is only $remainingdays days left before expiry. Reported by $reportingserver" `
        -ActivityDetails $Fact1
    Send-TeamsMessage `
        -URI $NotificationChannelID `
        -MessageTitle $Company `
        -MessageText "" `
        -Color Orange `
        -Sections $Section

}



UpdateCheck
