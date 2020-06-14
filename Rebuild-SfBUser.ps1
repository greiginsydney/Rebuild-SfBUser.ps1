<#
.SYNOPSIS
    The script enables one or more users for SfB based on a csv input file. It also assigns all the relevant policies as well as assigning their line uri.

.DESCRIPTION

	This script will run through a csv input file and generate the required commands to enable users for SfB & assign (grant) policies.
	The script will also report on any users in the csv file that were not enabled for SfB.
    Each user in the csv file will need a valid AD user account.
	The csv file column headers are used as variables so need to be unchanged.
	The script can be set to use various column headers as the "-Identity" of the user being created.

.Format
	The csv file will need the following column headers as a minimum:-
	SamAccountName,SIPAddress,RegistrarPool
	If any policy is to remain as default then simply edit out the policy in the script below, the labels provide for easy recognition within the script.
    Make sure that the data within each row matches already configured, valid policies

.Example of input file format
SamAccountName,SIPAddress,LineUri,LocationPolicy,DialPlanPolicy,VoicePolicy,ConferencingPolicy,ExternalPolicy,ClientPolicy,MobilityPolicy,RegistrarPool
PanP,peter.pan@SfBsorted.co.nz,+6499702700,Auckland,Auckland_Dial_Plan,NZ_International,Audio_Only,Full External_Access,Photo_Control,SfBpool01.SfBsorted.co.nz

.NOTES

	Updated by Greig Sheridan (greiginsydney.com) 14th June 2020.
	Now takes a user dump from Lync/SfB and recreates users from the file. Any policies in the file that are blank will be skipped, thus leaving the default assigned

	This is the way to create a sample CSV file:
	get-csuser -filter {RegistrarPool -ne $null} | select SamAccountName,SipAddress,RegistrarPool,DialPlan,LineURI,EnterpriseVoiceEnabled,*policy | export-csv -NoTypeInformation -path <FILENAME.CSV>

	BASED ON:
	File Name: EnableLyncusers.ps1
	Version: 0.4
	Last Update: 17-May-2014
	Author: Paul Bloem, http://ucsorted.com

    The script is provided “AS IS” with no guarantees, no warranties, USE AT YOUR OWN RISK.
#>

[cmdletbinding()]
param(
	[parameter (ValuefromPipeline = $false, ValueFromPipelineByPropertyName = $false, Mandatory = $true)]
	[string]$InputFile,
	[validateSet("DisplayName", "SamAccountName", "UserPrincipalName")]
	[string]$IdentityField="DisplayName",
	[Parameter(ParameterSetName='Default', Mandatory = $false)]
	[switch]$SkipUpdateCheck
)


#--------------------------------
# START FUNCTIONS ---------------
#--------------------------------

function Get-UpdateInfo
{
  <#
	.SYNOPSIS
	Queries an online XML source for version information to determine if a new version of the script is available.
	*** This version customised by Greig Sheridan. @greiginsydney https://greiginsydney.com ***

	.DESCRIPTION
	Queries an online XML source for version information to determine if a new version of the script is available.

	.NOTES
	Version				: 1.2 - See changelog at https://ucunleashed.com/3168 for fixes & changes introduced with each version
	Wish list				: Better error trapping
	Rights Required		: N/A
	Sched Task Required	: No
	Lync/Skype4B Version	: N/A
	Author/Copyright		: © Pat Richard, Office Servers and Services (Skype for Business) MVP - All Rights Reserved
	Email/Blog/Twitter	: pat@innervation.com  https://ucunleashed.com  @patrichard
	Donations				: https://www.paypal.me/PatRichard
	Dedicated Post		: https://ucunleashed.com/3168
	Disclaimer			: You running this script/function means you will not blame the author(s) if this breaks your stuff. This script/function
							is provided AS IS without warranty of any kind. Author(s) disclaim all implied warranties including, without limitation,
							any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use
							or performance of the sample scripts and documentation remains with you. In no event shall author(s) be held liable for
							any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss
							of business information, or other pecuniary loss) arising out of the use of or inability to use the script or
							documentation. Neither this script/function, nor any part of it other than those parts that are explicitly copied from
							others, may be republished without author(s) express written permission. Author(s) retain the right to alter this
							disclaimer at any time. For the most up to date version of the disclaimer, see https://ucunleashed.com/code-disclaimer.
	Acknowledgements		: Reading XML files
							http://stackoverflow.com/questions/18509358/how-to-read-xml-in-powershell
							http://stackoverflow.com/questions/20433932/determine-xml-node-exists
	Assumptions			: ExecutionPolicy of AllSigned (recommended), RemoteSigned, or Unrestricted (not recommended)
	Limitations			:
	Known issues			:

	.EXAMPLE
	Get-UpdateInfo -Title 'Rebuild-SfBUsers.ps1'

	Description
	-----------
	Runs function to check for updates to script called 'Rebuild-SfBUsers.ps1'.

	.INPUTS
	None. You cannot pipe objects to this script.
  #>
	[CmdletBinding(SupportsShouldProcess = $true)]
	param (
	[string] $title
	)
	try
	{
		[bool] $HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)
		if ($HasInternetAccess)
		{
			write-verbose -message 'Performing update check'
			# ------------------ TLS 1.2 fixup from https://github.com/chocolatey/choco/wiki/Installation#installing-with-restricted-tls
			$securityProtocolSettingsOriginal = [Net.ServicePointManager]::SecurityProtocol
			try {
			# Set TLS 1.2 (3072). Use integers because the enumeration values for TLS 1.2 won't exist in .NET 4.0, even though they are
			# addressable if .NET 4.5+ is installed (.NET 4.5 is an in-place upgrade).
			[Net.ServicePointManager]::SecurityProtocol = 3072
			} catch {
			write-verbose -message 'Unable to set PowerShell to use TLS 1.2 due to old .NET Framework installed.'
			}
			# ------------------ end TLS 1.2 fixup
			[xml] $xml = (New-Object -TypeName System.Net.WebClient).DownloadString('https://greiginsydney.com/wp-content/version.xml')
			[Net.ServicePointManager]::SecurityProtocol = $securityProtocolSettingsOriginal #Reinstate original SecurityProtocol settings
			$article  = select-XML -xml $xml -xpath ("//article[@title='{0}']" -f ($title))
			[string] $Ga = $article.node.version.trim()
			if ($article.node.changeLog)
			{
				[string] $changelog = 'This version includes: ' + $article.node.changeLog.trim() + "`n`n"
			}
			if ($Ga -gt $ScriptVersion)
			{
				$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
				$updatePrompt = $wshell.Popup(("Version {0} is available.`n`n{1}Would you like to download it?" -f ($ga), ($changelog)),0,'New version available',68)
				if ($updatePrompt -eq 6)
				{
					Start-Process -FilePath $article.node.downloadUrl
					write-warning -message "Script is exiting. Please run the new version of the script after you've downloaded it."
					exit
				}
				else
				{
					write-verbose -message ('Upgrade to version {0} was declined' -f ($ga))
				}
			}
			elseif ($Ga -eq $ScriptVersion)
			{
				write-verbose -message ('Script version {0} is the latest released version' -f ($Scriptversion))
			}
			else
			{
				write-verbose -message ('Script version {0} is newer than the latest released version {1}' -f ($Scriptversion), ($ga))
			}
		}
		else
		{
		}
	} # end function Get-UpdateInfo
	catch
	{
		write-verbose -message 'Caught error in Get-UpdateInfo'
		if ($Global:Debug)
		{
			$Global:error | Format-List -Property * -Force #This dumps to screen as white for the time being. I haven't been able to get it to dump in red
		}
	}
}

function Grant-CsPolicy
{
	param ([string]$User, [string]$Policy, [string]$value)
	if ([string]::IsNullOrEmpty($value)) { $value = "`$null" }
	try
	{
		invoke-expression "(Grant-Cs$($Policy) -Identity ""$($User)"" -PolicyName ""$($value)"")"
	}
	catch
	{
		return "Error"
	}
}

#--------------------------------
# END FUNCTIONS -----------------
#--------------------------------


$ScriptVersion = "1.0"
$Error.Clear()
$scriptpath = $myInvocation.MyCommand.Path
$dir = split-path -path $scriptpath
$Log = New-Item -ItemType File -Path "$($dir)\Rebuild-SfBUser-Log.txt" -Force


if ($skipupdatecheck)
{
	write-verbose -message 'Skipping update check'
}
else
{
	write-progress -id 1 -Activity 'Initialising' -Status 'Performing update check' -PercentComplete (2)
	Get-UpdateInfo -title 'Rebuild-SfBUsers.ps1'
	write-progress -id 1 -Activity 'Initialising' -Status 'Back from performing update check' -PercentComplete (2)
}


#Import csv
$usercsv = Import-Csv -path $InputFile -Delimiter ','
#Check if user file is empty.
if ($Usercsv -eq $null)
{
	write-host "No Users Found in the Input File"
	exit 0
}

#Get the number of users in CSV file and begin proccessing.

$count = $Usercsv | Measure-Object | Select-Object -expand count

Write-Host "Found " $count " users to enable for SfB."
Write-Host "Enabling Users.....`n"
$index = 0

# First run through we enable the users:
ForEach ($user in $usercsv)
{
	$index++
	Write-Host "Enabling User " $index " of " $count
	if ([string]::IsNullOrEmpty($user.$IdentityField))
	{
		Add-Content -Path $Log -Value "Skipping. User from CSV has no $($IdentityField). $(Get-Date)"
		continue
	}
	else
	{
		$Identity = $user.$IdentityField
	}
	Write-Host "Enabling user $($Identity)" -Foregroundcolor Green
	Enable-CsUser -Identity "$($Identity)" -RegistrarPool $user.RegistrarPool -SipAddress "$($user.SipAddress)"
	#Check if the previous command failed. If it did then write that to the log file.
	if(!$?)
	{
		Add-Content -Path $Log -Value "$(Get-Date) Failed to enable $($Identity). $(Get-Date)$($error[0])"
		continue
	}
}

$index = 0
# now we add all the policies and other settings:
ForEach ($user in $usercsv)
{
	$index++
	#Set Enterprise Voice (if required)
	if (-not ([string]::IsNullOrEmpty($user.EnterpriseVoiceEnabled)))
	{
		Set-CsUser -identity "$($user.SipAddress)" -EnterpriseVoiceEnabled ([System.Convert]::ToBoolean("$($user.EnterpriseVoiceEnabled)"))
		#Check if previous command failed. If it did then write that to the log file.
		if(!$?)
		{
			Add-Content -Path $Log -Value "$(Get-Date) Failed to enable $($user.SipAddress) for Enterprise Voice. $($error[0])"
		}
	}
	#Set Enterprise Voice (if required)
	if (-not ([string]::IsNullOrEmpty($user.HostedVoiceMail)))
	{
		Set-CsUser -identity "$($user.SipAddress)" -HostedVoiceMail "$($user.HostedVoiceMail)"
		#Check if previous command failed. If it did then write that to the log file.
		if(!$?)
		{
			Add-Content -Path $Log -Value "$(Get-Date) Failed to enable $($user.SipAddress) for Hosted voicemail. $($error[0])"
		}
	}
	#Add the Line URI
	if (-not ([string]::IsNullOrEmpty($user.LineURI)))
	{
		Set-CsUser -identity "$($user.SipAddress)" -LineURI "$($user.LineURI)"
		#Check if previous command failed. If it did then write that to the log file.
		if(!$?)
		{
			Add-Content -Path $Log -Value "$(Get-Date) Failed to set    $($user.SipAddress) LineURI $($user.LineURI). $($error[0])"
		}
	}
	#Add the OnPremLine URI
	if (-not ([string]::IsNullOrEmpty($user.OnPremLineURI)))
	{
		Set-CsUser -identity "$($user.SipAddress)" -OnPremLineURI "$($user.OnPremLineURI)"
		#Check if previous command failed. If it did then write that to the log file.
		if(!$?)
		{
			Add-Content -Path $Log -Value "$(Get-Date) Failed to set    $($user.SipAddress) OnPremLineURI $($user.OnPremLineURI). $($error[0])"
		}
	}
	#Assign the defined Dial Plan
	if (-not ([string]::IsNullOrEmpty($user.DialPlan)))
	{
		$result = Grant-CsPolicy "$($user.SipAddress)" "DialPlan" $user.DialPlan
		if (-not ([string]::IsNullOrEmpty($result)))
		{
			Add-Content -Path $Log -Value "$(Get-Date) Failed to set    $($user.SipAddress) Dial Plan $($user.DialPlan). $($error[0])"
		}
	}
	#Now loop through all the policies:
	foreach ($columnTitle in $usercsv[0].psobject.properties.name)
	{
		if (($columnTitle -like "*policy") -and ($columnTitle -ne "ExchangeArchivingPolicy"))
		{
			$result = Grant-CsPolicy "$($user.SipAddress)" $($user.psobject.properties["$($columnTitle)"].name) $($user.psobject.properties["$($columnTitle)"].value)
			if (-not ([string]::IsNullOrEmpty($result)))
			{
				Add-Content -Path $Log -Value "$(Get-Date) Failed to grant  $($user.SipAddress) a $($user.psobject.properties["$($columnTitle)"].name) of $($user.psobject.properties["$($columnTitle)"].value). $($error[0])"
			}
		}
	}
}

Write-Host "Provisioning of users to SfB from CSV has completed!"
Write-Host ""
Write-Host ""
Write-Host ""

#Update the AddressBook Service
sleep -seconds 10
update-csuserdatabase -verbose
