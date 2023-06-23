#This File is in Unicode format.  Do not edit in an ASCII editor.

<#
.SYNOPSIS
	Creates a Summary Report of the inventory of a Citrix XenApp 6.5 farm using Microsoft Word.
.DESCRIPTION
	Creates a Summary Report of the inventory of a Citrix XenApp 6.5 farm using Microsoft Word.l.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish
.EXAMPLE
	PS C:\PSScript > .\XA65_SummaryReport.ps1
	
	Runs and creates a one page report.
	
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word document.
.NOTES
	NAME: XA65_SummaryReport.ps1
	VERSION: 1.00
	AUTHOR: Carl Webster
	LASTEDIT: November 2, 2013
#>

Set-StrictMode -Version 2

#the following values were attained from 
#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
[int]$wdMove = 0
[int]$wdSeekMainDocument = 0
[int]$wdStory = 6
[int]$wdWord2007 = 12
[int]$wdWord2010 = 14
[int]$wdWord2013 = 15
[string]$RunningOS = (Get-WmiObject -class Win32_OperatingSystem).Caption

$hash = @{}

# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
$wdStyleHeading1 = -2
$wdStyleHeading2 = -3
$wdStyleHeading3 = -4
$wdStyleHeading4 = -5
$wdStyleNoSpacing = -158

$myHash = $hash

$myHash.Word_NoSpacing = $wdStyleNoSpacing
$myHash.Word_Heading1 = $wdStyleheading1
$myHash.Word_Heading2 = $wdStyleheading2
$myHash.Word_Heading3 = $wdStyleheading3
$myHash.Word_Heading4 = $wdStyleheading4

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		Write-Host "This script directly outputs to Microsoft Word, please install Microsoft Word"
		exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		Write-Host "Please close all instances of Microsoft Word before running this report."
		exit
	}
}

#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}

Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
{
	Param([int]$style=0, [int]$tabs = 0, [string]$name = '', [string]$value = '', [string]$newline = "'n", [Switch]$nonewline)
	[string]$output = ""
	#Build output style
	Switch ($style)
	{
		0 {$Selection.Style = $myHash.Word_NoSpacing}
		1 {$Selection.Style = $myHash.Word_Heading1}
		2 {$Selection.Style = $myHash.Word_Heading2}
		3 {$Selection.Style = $myHash.Word_Heading3}
		4 {$Selection.Style = $myHash.Word_Heading4}
		Default {$Selection.Style = $myHash.Word_NoSpacing}
	}
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
		
	#output the rest of the parameters.
	$output += $name + $value
	$Selection.TypeText($output)
	
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Selection.TypeParagraph()
	}
}

Function Check-LoadedModule
#Function created by Jeff Wouters
#@JeffWouters on Twitter
#modified by Michael B. Smith to handle when the module doesn't exist on server
#modified by @andyjmorgan
#bug fixed by @schose
#This Function handles all three scenarios:
#
# 1. Module is already imported into current session
# 2. Module is not already imported into current session, it does exists on the server and is imported
# 3. Module does not exist on the server

{
	Param([parameter(Mandatory = $True)][alias("Module")][string]$ModuleName)
	#$LoadedModules = Get-Module | Select Name
	#following line changed at the recommendation of @andyjmorgan
	$LoadedModules = Get-Module |% { $_.Name.ToString() }
	#bug reported on 21-JAN-2013 by @schose 
	#the following line did not work if the citrix.grouppolicy.commands.psm1 module
	#was manually loaded from a non Default folder
	#$ModuleFound = (!$LoadedModules -like "*$ModuleName*")
	[string]$ModuleFound = ($LoadedModules -like "*$ModuleName*")
	If(!$ModuleFound) 
	{
		$module = Import-Module -Name $ModuleName -PassThru -EA 0
		If($module -and $?)
		{
			# module imported properly
			Return $True
		}
		Else
		{
			# module import failed
			Return $False
		}
	}
	Else
	{
		#module already imported into current session
		Return $True
	}
}

Function Check-NeededPSSnapins
{
	Param([parameter(Mandatory = $True)][alias("Snapin")][string[]]$Snapins)

	#Function specifics
	$MissingSnapins = @()
	[bool]$FoundMissingSnapin = $False
	$LoadedSnapins = @()
	$RegisteredSnapins = @()

	#Creates arrays of strings, rather than objects, we're passing strings so this will be more robust.
	$loadedSnapins += get-pssnapin | % {$_.name}
	$registeredSnapins += get-pssnapin -Registered | % {$_.name}

	ForEach($Snapin in $Snapins)
	{
		#check if the snapin is loaded
		If(!($LoadedSnapins -like $snapin))
		{
			#Check if the snapin is missing
			If(!($RegisteredSnapins -like $Snapin))
			{
				#set the flag if it's not already
				If(!($FoundMissingSnapin))
				{
					$FoundMissingSnapin = $True
				}
				#add the entry to the list
				$MissingSnapins += $Snapin
			}
			Else
			{
				#Snapin is registered, but not loaded, loading it now:
				Write-Host "Loading Windows PowerShell snap-in: $snapin"
				Add-PSSnapin -Name $snapin -EA 0
			}
		}
	}

	If($FoundMissingSnapin)
	{
		Write-Warning "Missing Windows PowerShell snap-ins Detected:"
		$missingSnapins | % {Write-Warning "($_)"}
		return $False
	}
	Else
	{
		Return $True
	}
}

Function ProcessCitrixPolicies
{
	Param([string]$xDriveName)

	If($xDriveName -eq "")
	{
		$Policies = Get-CtxGroupPolicy -EA 0 | Sort-Object Type,Priority
	}
	Else
	{
		$Policies = Get-CtxGroupPolicy -DriveName $xDriveName -EA 0 | Sort-Object Type,Priority
	}
	If($?)
	{
		ForEach($Policy in $Policies)
		{
			Write-Host "$(Get-Date): `tStarted $($Policy.PolicyName)`t$($Policy.Type)"
			If($xDriveName -eq "")
			{
				$Global:TotalIMAPolicies++
			}
			Else
			{
				$Global:TotalADPolicies++
			}

			If($Policy.Type -eq "Computer")
			{
				$Global:TotalComputerPolicies++
			}
			Else
			{
				$Global:TotalUserPolicies++
			}
		}
		
	}
	Else 
	{
		Write-Warning "Citrix Policy information could not be retrieved."
	}
	$Policies = $Null
	If($xDriveName -ne "")
	{
		Write-Host "$(Get-Date): `tRemoving ADGpoDrv PSDrive"
		Remove-PSDrive ADGpoDrv -EA 0
		Write-Host "$(Get-Date): "
	}
}

Function GetCtxGPOsInAD
{
	#thanks to the Citrix Engineering Team for pointers and for Michael B. Smith for creating the function
	$root = [ADSI]"LDAP://RootDSE"
	$domainNC = $root.defaultNamingContext.ToString()
	$root = $null
	$xArray = @()

	$domain = $domainNC.Replace( 'DC=', '').Replace( ',', '.')
	$sysvolFiles = dir -Recurse ( '\\' + $domain  + '\sysvol\' + $domain + '\Policies')
	ForEach($file in $sysvolFiles)
	{
		If(-not $file.PSIsContainer)
		{
			#$file.FullName  ### name of the policy file
			If($file.FullName -like "*\Citrix\GroupPolicy\Policies.gpf")
			{
				#"have match " + $file.FullName ### name of the Citrix policies file
				$array = $file.FullName.Split( '\')
				If($array.Length -gt 7)
				{
					$gp = $array[ 6 ].ToString()
					$gpObject = [ADSI]( "LDAP://" + "CN=" + $gp + ",CN=Policies,CN=System," + $domainNC)
					$xArray += $gpObject.DisplayName	### name of the group policy object
				}
			}
		}
	}
	Return ,$xArray | Sort
}

Function AbortScript
{
	$Word.quit()
	Write-Host "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
	Remove-Variable -Name word -Scope Global -EA 0
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Host "$(Get-Date): Script has been aborted"
	Exit
}

#Script begins

If(!(Check-NeededPSSnapins "Citrix.Common.Commands","Citrix.XenApp.Commands"))
{
    #We're missing Citrix Snapins that we need
    Write-Error "Missing Citrix PowerShell Snap-ins Detected, check the console above for more information. Are you sure you are running this script on a XenApp 6.5 Server? Script will now close."
    Exit
}

CheckWordPreReq

[bool]$Remoting = $False
$RemoteXAServer = Get-XADefaultComputerName -EA 0
If(![String]::IsNullOrEmpty($RemoteXAServer))
{
	$Remoting = $True
}

If($Remoting)
{
	Write-Host "$(Get-Date): Remoting is enabled to XenApp server $RemoteXAServer"
}
Else
{
	Write-Host "$(Get-Date): Remoting is not being used"
	
	#now need to make sure the script is not being run on a session-only host
	$ServerName = (Get-Childitem env:computername).value
	$Server = Get-XAServer -ServerName $ServerName -EA 0
	If($Server.ElectionPreference -eq "WorkerMode")
	{
		Write-Warning "This script cannot be run on a Session-only Host Server if Remoting is not enabled."
		Write-Warning "Use Set-XADefaultComputerName XA65ControllerServerName or run the script on a controller."
		Write-Error "Script cannot continue.  See messages above."
		Exit
	}
}

# Get farm information
Write-Host "$(Get-Date): Getting Farm data"
$farm = Get-XAFarm -EA 0

If($?)
{
	Write-Host "$(Get-Date): Verify farm version"
	#first check to make sure this is a XenApp 6.5 farm
	If($Farm.ServerVersion.ToString().SubString(0,3) -eq "6.5")
	{
		#this is a XenApp 6.5 farm, script can proceed
	}
	Else
	{
		#this is not a XenApp 6.5 farm, script cannot proceed
		Write-Warning "This script is designed for XenApp 6.5 and should not be run on previous versions of XenApp"
		Return 1
	}
	[string]$FarmName = $farm.FarmName
	[string]$filename1 = "$($pwd.path)\Summary Report for $($FarmName).docx"
} 
Else 
{
	Write-Warning "Farm information could not be retrieved"
	If($Remoting)
	{
		Write-Error "A remote connection to $RemoteXAServer could not be established.  Script cannot continue."
	}
	Else
	{
		Write-Error "Farm information could not be retrieved.  Script cannot continue."
	}
	Exit
}
$farm = $Null

Write-Host "$(Get-Date): Setting up Word"

# Setup word for output
Write-Host "$(Get-Date): Create Word comObject.  If you are not running Word 2007, ignore the next message."
$Word = New-Object -comobject "Word.Application" -EA 0

If(!$? -or $Word -eq $Null)
{
	Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
	Write-Error "The Word object could not be created.  You may need to repair your Word installation.  Script cannot continue."
	Exit
}

[int]$WordVersion = [int]$Word.Version
If($WordVersion -eq $wdWord2013)
{
	$WordProduct = "Word 2013"
}
ElseIf($WordVersion -eq $wdWord2010)
{
	$WordProduct = "Word 2010"
}
ElseIf($WordVersion -eq $wdWord2007)
{
	$WordProduct = "Word 2007"
}
Else
{
	Write-Error "You are running an untested or unsupported version of Microsoft Word.  Script will end.  Please send info on your version of Word to webster@carlwebster.com"
	AbortScript
}

Write-Host "$(Get-Date): Running Microsoft $WordProduct"
$Word.Visible = $False

Write-Host "$(Get-Date): Create empty word doc"
$Doc = $Word.Documents.Add()
If($Doc -eq $Null)
{
	Write-Host "$(Get-Date): "
	Write-Error "An empty Word document could not be created.  Script cannot continue."
	AbortScript
}

$Selection = $Word.Selection
If($Selection -eq $Null)
{
	Write-Host "$(Get-Date): "
	Write-Error "An unknown error happened selecting the entire Word document for default formatting options.  Script cannot continue."
	AbortScript
}

#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
#36 = .50"
$Word.ActiveDocument.DefaultTabStop = 36

#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
Write-Host "$(Get-Date): Disable grammar and spell checking"
$Word.Options.CheckGrammarAsYouType = $False
$Word.Options.CheckSpellingAsYouType = $False

#return focus to main document
Write-Host "$(Get-Date): Return focus to main document"
$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

#move to the end of the current document
Write-Host "$(Get-Date): Move to the end of the current document"
Write-Host "$(Get-Date):"
$selection.EndKey($wdStory,$wdMove) | Out-Null
#end of Jeff Hicks 

Write-Host "$(Get-Date): Processing Configuration Logging"
[bool]$ConfigLog = $False
$ConfigurationLogging = Get-XAConfigurationLog -EA 0

If($?)
{
	If($ConfigurationLogging.LoggingEnabled) 
	{
		$ConfigLog = $True
	}
}
Else 
{
	Write-Warning  "Configuration Logging could not be retrieved"
}
$ConfigurationLogging = $Null
Write-Host "$(Get-Date): Finished Configuration Logging"
Write-Host "$(Get-Date): "

Write-Host "$(Get-Date): Processing Administrators"
Write-Host "$(Get-Date): `tSetting summary variables"
[int]$TotalFullAdmins = 0
[int]$TotalViewAdmins = 0
[int]$TotalCustomAdmins = 0

Write-Host "$(Get-Date): `tRetrieving Administrators"
$Administrators = Get-XAAdministrator -EA 0 | Sort-Object AdministratorName

If($?)
{
	ForEach($Administrator in $Administrators)
	{
		Write-Host "$(Get-Date): `t`tProcessing administrator $($Administrator.AdministratorName)"
		Switch ($Administrator.AdministratorType)
		{
			"Unknown"  {}
			"Full"     {$TotalFullAdmins++}
			"ViewOnly" {$TotalViewAdmins++}
			"Custom"   {$TotalCustomAdmins++}
			Default    {}
		}
	}
}
Else 
{
	Write-Warning "Administrator information could not be retrieved"
}
$Administrators = $Null
Write-Host "$(Get-Date): Finished Processing Administrators"
Write-Host "$(Get-Date): "

Write-Host "$(Get-Date): Processing Applications"
[int]$TotalPublishedApps = 0
[int]$TotalPublishedContent = 0
[int]$TotalPublishedDesktops = 0
[int]$TotalStreamedApps = 0

Write-Host "$(Get-Date): `tRetrieving Applications"
$Applications = Get-XAApplication -EA 0 | Sort-Object FolderPath, DisplayName

If($? -and $Applications -ne $Null)
{

	ForEach($Application in $Applications)
	{
		Write-Host "$(Get-Date): `t`tProcessing application $($Application.BrowserName)"
		
		#type properties
		Switch ($Application.ApplicationType)
		{
			"Unknown"                            {}
			"ServerInstalled"                    {$TotalPublishedApps++}
			"ServerDesktop"                      {$TotalPublishedDesktops++}
			"Content"                            {$TotalPublishedContent++}
			"StreamedToServer"                   {$TotalStreamedApps++}
			"StreamedToClient"                   {$TotalStreamedApps++}
			"StreamedToClientOrInstalled"        {$TotalStreamedApps++}
			"StreamedToClientOrStreamedToServer" {$TotalStreamedApps++}
			Default {}
		}
	}
}
ElseIf($Applications -eq $Null)
{
	Write-Host "$(Get-Date): There are no Applications published"
}
Else 
{
	Write-Warning "Application information could not be retrieved.  Do you have any published applications?"
}
$Applications = $Null
Write-Host "$(Get-Date): Finished Processing Applications"
Write-Host "$(Get-Date): "

[int]$TotalConfigLogItems = 0

Write-Host "$(Get-Date): Processing Configuration Logging/History Report"
If($ConfigLog)
{
	#history AKA Configuration Logging report
	#only process if $ConfigLog = $True and XA65ConfigLog.udl file exists
	#build connection string
	#User ID is account that has access permission for the configuration logging database
	#Initial Catalog is the name of the Configuration Logging SQL Database
	If(Test-Path “$($pwd.path)\XA65ConfigLog.udl”)
	{
		$ConnectionString = Get-Content “$($pwd.path)\XA65ConfigLog.udl” | select-object -last 1
		$ConfigLogReport = get-CtxConfigurationLogReport -connectionstring $ConnectionString -EA 0

		If($? -and $ConfigLogReport)
		{
			Write-Host "$(Get-Date): `tProcessing $($ConfigLogReport.Count) history items"
			ForEach($ConfigLogItem in $ConfigLogReport)
			{
				$TotalConfigLogItems++
			}
		} 
		Else 
		{
			Write-Warning "History information could not be retrieved"
		}
		$ConnectionString = $Null
		$ConfigLogReport = $Null
	}
	Else 
	{
		Write-Warning "Configuration Logging is enabled but the XA65ConfigLog.udl file was not found"
	}
}

Write-Host "$(Get-Date): Finished Processing Configuration Logging/History Report"
Write-Host "$(Get-Date): "

#load balancing policies
Write-Host "$(Get-Date): Processing Load Balancing Policies"
[int]$TotalLBPolicies = 0

Write-Host "$(Get-Date): `tRetrieving Load Balancing Policies"
$LoadBalancingPolicies = Get-XALoadBalancingPolicy -EA 0 | Sort-Object PolicyName

If($? -and $LoadBalancingPolicies -ne $Null)
{
	ForEach($LoadBalancingPolicy in $LoadBalancingPolicies)
	{
		$TotalLBPolicies++
		Write-Host "$(Get-Date): `t`tProcessing Load Balancing Policy $($LoadBalancingPolicy.PolicyName)"
	}
}
Elseif($LoadBalancingPolicies -eq $Null)
{
	Write-Host "$(Get-Date): There are no Load balancing policies created"
}
Else 
{
	Write-Warning "Load balancing policy information could not be retrieved.  "
}
$LoadBalancingPolicies = $Null
Write-Host "$(Get-Date): Finished Processing Load Balancing Policies"
Write-Host "$(Get-Date): "

#load evaluators
Write-Host "$(Get-Date): Processing Load Evaluators"
[int]$TotalLoadEvaluators = 0

Write-Host "$(Get-Date): `tRetrieving Load Evaluators"
$LoadEvaluators = Get-XALoadEvaluator -EA 0 | Sort-Object LoadEvaluatorName

If($?)
{
	ForEach($LoadEvaluator in $LoadEvaluators)
	{
		$TotalLoadEvaluators++
		Write-Host "$(Get-Date): `t`tProcessing Load Evaluator $($LoadEvaluator.LoadEvaluatorName)"
	}
}
Else 
{
	Write-Warning "Load Evaluator information could not be retrieved"
}
$LoadEvaluators = $Null
Write-Host "$(Get-Date): Finished Processing Load Evaluators"
Write-Host "$(Get-Date): "

#servers
Write-Host "$(Get-Date): Processing Servers"
[int]$TotalControllers = 0
[int]$TotalWorkers = 0

Write-Host "$(Get-Date): `tRetrieving Servers"
$servers = Get-XAServer -EA 0 | Sort-Object FolderPath, ServerName

If($?)
{
	ForEach($server in $servers)
	{
		Write-Host "$(Get-Date): `t`tProcessing server $($server.ServerName)"
		Switch ($server.ElectionPreference)
		{
			"Unknown"           {}
			"MostPreferred"     {$TotalControllers++}
			"Preferred"         {$TotalControllers++}
			"DefaultPreference" {$TotalControllers++}
			"NotPreferred"      {$TotalControllers++}
			"WorkerMode"        {$TotalWorkers++}
			Default {}
		}
	}
}
Else 
{
	Write-Warning "Server information could not be retrieved"
}
$servers = $Null
Write-Host "$(Get-Date): Finished Processing Servers"
Write-Host "$(Get-Date): "

#worker groups
Write-Host "$(Get-Date): Processing Worker Groups"
[int]$TotalWGByServerName = 0
[int]$TotalWGByServerGroup = 0
[int]$TotalWGByOU = 0

Write-Host "$(Get-Date): `tRetrieving Worker Groups"
$WorkerGroups = Get-XAWorkerGroup -EA 0 | Sort-Object WorkerGroupName

If($? -and $WorkerGroups -ne $Null)
{
	ForEach($WorkerGroup in $WorkerGroups)
	{
		Write-Host "$(Get-Date): `t`tProcessing Worker Group $($WorkerGroup.WorkerGroupName)"
		If($WorkerGroup.ServerNames)
		{
			$TotalWGByServerName++
		}
		If($WorkerGroup.ServerGroups)
		{
			$TotalWGByServerGroup++
		}
		If($WorkerGroup.OUs)
		{
			$TotalWGByOU++
		}
	}
}
ElseIf($WorkerGroups -eq $Null)
{

	Write-Host "$(Get-Date): There are no Worker Groups created"
}
Else 
{
	Write-Warning "Worker Group information could not be retrieved"
}
$WorkerGroups = $Null
Write-Host "$(Get-Date): Finished Processing Worker Groups"
Write-Host "$(Get-Date): "

#zones
Write-Host "$(Get-Date): Processing Zones"
[int]$TotalZones = 0

Write-Host "$(Get-Date): `tRetrieving Zones"
$Zones = Get-XAZone -EA 0 | Sort-Object ZoneName
If($?)
{
	ForEach($Zone in $Zones)
	{
		$TotalZones++
	}
}
Else 
{
	Write-Warning "Zone information could not be retrieved"
}
$Servers = $Null
$Zones = $Null
Write-Host "$(Get-Date): Finished Processing Zones"
Write-Host "$(Get-Date): "

[int]$Global:TotalComputerPolicies = 0
[int]$Global:TotalUserPolicies = 0
[int]$Global:TotalIMAPolicies = 0
[int]$Global:TotalADPolicies = 0
[int]$Global:TotalADPoliciesNotProcessed = 0

#if remoting is enabled, the citrix.grouppolicy.commands module does not work with remoting so skip it
If($Remoting)
{
	Write-Warning "Remoting is enabled."
	Write-Warning "The Citrix.GroupPolicy.Commands module does not work with Remoting."
	Write-Warning "Citrix Policy documentation will not take place."
}
Else
{
	#make sure Citrix.GroupPolicy.Commands module is loaded
	If(!(Check-LoadedModule "Citrix.GroupPolicy.Commands"))
	{
		Write-Warning "The Citrix Group Policy module Citrix.GroupPolicy.Commands.psm1 could not be loaded `nPlease see http://tinyurl.com/XenApp6PSPolicies `nCitrix Policy documentation will not take place"
		Write-Host "$(Get-Date): "
	}
	Else
	{
		Write-Host "$(Get-Date): Processing Citrix IMA Policies"
		Write-Host "$(Get-Date): `tRetrieving IMA Farm Policies"
		ProcessCitrixPolicies	
		Write-Host "$(Get-Date): Finished Processing Citrix IMA Policies"
		Write-Host "$(Get-Date): "
		
		#thanks to the Citrix Engineering Team for helping me solve processing Citrix AD based Policies
		Write-Host "$(Get-Date): See if there are any Citrix AD based policies to process"
		$CtxGPOArray = @()
		$CtxGPOArray = GetCtxGPOsInAD
		If($CtxGPOArray -is [Array] -and $CtxGPOArray.Count -gt 0)
		{
			Write-Host "$(Get-Date): There are $($CtxGPOArray.Count) Citrix AD based policies to process"
			
			ForEach($CtxGPO in $CtxGPOArray)
			{
				Write-Host "$(Get-Date): Creating ADGpoDrv PSDrive"
				New-PSDrive -Name ADGpoDrv -PSProvider CitrixGroupPolicy -Root \ -DomainGpo $($CtxGPO) -Scope "Global" | out-null
				If(Get-PSDrive ADGpoDrv -EA 0)
				{
					ProcessCitrixPolicies "ADGpoDrv"
				}
				Else
				{
					$Global:TotalADPoliciesNotProcessed++
				}
			}
		
			Write-Host "$(Get-Date): Finished Processing Citrix AD Policies"
			Write-Host "$(Get-Date): "
		}
		Else
		{
			Write-Host "$(Get-Date): There are no Citrix AD based policies to process"
			Write-Host "$(Get-Date): "
		}
		Write-Host "$(Get-Date): Finished Processing Citrix Policies"
		Write-Host "$(Get-Date): "
	}
}

#summary page
Write-Host "$(Get-Date): Create Summary Report"
WriteWordLine 1 0 "Summary Report for the $($FarmName) Farm"
Write-Host "$(Get-Date): `tAdd administrator summary info"
WriteWordLine 0 0 "Administrators"
WriteWordLine 0 1 "Total Full Administrators`t: " $TotalFullAdmins
WriteWordLine 0 1 "Total View Administrators`t: " $TotalViewAdmins
WriteWordLine 0 1 "Total Custom Administrators`t: " $TotalCustomAdmins
WriteWordLine 0 2 "Total Administrators`t: " ($TotalFullAdmins + $TotalViewAdmins + $TotalCustomAdmins)
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): `tAdd application summary info"
WriteWordLine 0 0 "Applications"
WriteWordLine 0 1 "Total Published Applications`t: " $TotalPublishedApps
WriteWordLine 0 1 "Total Published Content`t`t: " $TotalPublishedContent
WriteWordLine 0 1 "Total Published Desktops`t: " $TotalPublishedDesktops
WriteWordLine 0 1 "Total Streamed Applications`t: " $TotalStreamedApps
WriteWordLine 0 2 "Total Applications`t: " ($TotalPublishedApps + $TotalPublishedContent + $TotalPublishedDesktops + $TotalStreamedApps)
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): `tAdd configuration logging summary info"
WriteWordLine 0 0 "Configuration Logging"
WriteWordLine 0 1 "Total Config Log Items`t`t: " $TotalConfigLogItems 
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): `tAdd load balancing policies summary info"
WriteWordLine 0 0 "Load Balancing Policies"
WriteWordLine 0 1 "Total Load Balancing Policies`t: " $TotalLBPolicies
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): `tAdd load evaluator summary info"
WriteWordLine 0 0 "Load Evaluators"
WriteWordLine 0 1 "Total Load Evaluators`t`t: " $TotalLoadEvaluators
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): `tAdd server summary info"
WriteWordLine 0 0 "Servers"
WriteWordLine 0 1 "Total Controllers`t`t: " $TotalControllers
WriteWordLine 0 1 "Total Workers`t`t`t: " $TotalWorkers
WriteWordLine 0 2 "Total Servers`t`t: " ($TotalControllers + $TotalWorkers)
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): `tAdd worker group summary info"
WriteWordLine 0 0 "Worker Groups"
WriteWordLine 0 1 "Total WGs by Server Name`t: " $TotalWGByServerName
WriteWordLine 0 1 "Total WGs by Server Group`t: " $TotalWGByServerGroup
WriteWordLine 0 1 "Total WGs by AD Container`t: " $TotalWGByOU
WriteWordLine 0 2 "Total Worker Groups`t: " ($TotalWGByServerName + $TotalWGByServerGroup + $TotalWGByOU)
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): `tAdd zone summary info"
WriteWordLine 0 0 "Zones"
WriteWordLine 0 1 "Total Zones`t`t`t: " $TotalZones
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): `tAdd policy summary info"
WriteWordLine 0 0 "Policies"
WriteWordLine 0 1 "Total Computer Policies`t`t: " $Global:TotalComputerPolicies
WriteWordLine 0 1 "Total User Policies`t`t: " $Global:TotalUserPolicies
WriteWordLine 0 2 "Total Policies`t`t: " ($Global:TotalComputerPolicies + $Global:TotalUserPolicies)
WriteWordLine 0 0 ""
WriteWordLine 0 1 "IMA Policies`t`t`t: " $Global:TotalIMAPolicies
WriteWordLine 0 1 "Citrix AD Policies Processed`t: $($Global:TotalADPolicies)`t(AD Policies can contain multiple Citrix policies)"
WriteWordLine 0 1 "Citrix AD Policies not Processed`t: " $Global:TotalADPoliciesNotProcessed
Write-Host "$(Get-Date): Finished Create Summary Page"
Write-Host "$(Get-Date): "
Write-Host "$(Get-Date): Finishing up Word document"

Write-Host "$(Get-Date): Save and Close document and Shutdown Word"
If($WordVersion -eq $wdWord2007)
{
	Write-Host "$(Get-Date): Running Word 2007 and detected operating system $($RunningOS)"
	If($RunningOS.Contains("Server 2008 R2"))
	{
		$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type] 
		$doc.SaveAs($filename1, $SaveFormat)
	}
	Else
	{
		#works for Server 2008 and Windows 7
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$doc.SaveAs([REF]$filename1, [ref]$SaveFormat)
	}
}
Else
{
	#the $saveFormat below passes StrictMode 2
	#I found this at the following two links
	#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
	#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
	$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
	$doc.SaveAs([REF]$filename1, [ref]$SaveFormat)
}

Write-Host "$(Get-Date): Closing Word"
$doc.Close()
$Word.Quit()
Write-Host "$(Get-Date): System Cleanup"
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
Remove-Variable -Name word -Scope Global -EA 0
$SaveFormat = $Null
[gc]::collect() 
[gc]::WaitForPendingFinalizers()
Write-Host "$(Get-Date): Script has completed"
Write-Host "$(Get-Date): "

Write-Host "$(Get-Date): $($filename1) is ready for use"
Write-Host "$(Get-Date): "

# SIG # Begin signature block
# MIIiywYJKoZIhvcNAQcCoIIivDCCIrgCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUUG9pJU+r80BUsZxx3hgSOl+7
# iw2ggh41MIIDtzCCAp+gAwIBAgIQDOfg5RfYRv6P5WD8G/AwOTANBgkqhkiG9w0B
# AQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAwWhcNMzExMTEwMDAwMDAwWjBlMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3Qg
# Q0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCtDhXO5EOAXLGH87dg
# +XESpa7cJpSIqvTO9SA5KFhgDPiA2qkVlTJhPLWxKISKityfCgyDF3qPkKyK53lT
# XDGEKvYPmDI2dsze3Tyoou9q+yHyUmHfnyDXH+Kx2f4YZNISW1/5WBg1vEfNoTb5
# a3/UsDg+wRvDjDPZ2C8Y/igPs6eD1sNuRMBhNZYW/lmci3Zt1/GiSw0r/wty2p5g
# 0I6QNcZ4VYcgoc/lbQrISXwxmDNsIumH0DJaoroTghHtORedmTpyoeb6pNnVFzF1
# roV9Iq4/AUaG9ih5yLHa5FcXxH4cDrC0kqZWs72yl+2qp/C3xag/lRbQ/6GW6whf
# GHdPAgMBAAGjYzBhMA4GA1UdDwEB/wQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBRF66Kv9JLLgjEtUYunpyGd823IDzAfBgNVHSMEGDAWgBRF66Kv9JLL
# gjEtUYunpyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEAog683+Lt8ONyc3pklL/3
# cmbYMuRCdWKuh+vy1dneVrOfzM4UKLkNl2BcEkxY5NM9g0lFWJc1aRqoR+pWxnmr
# EthngYTffwk8lOa4JiwgvT2zKIn3X/8i4peEH+ll74fg38FnSbNd67IJKusm7Xi+
# fT8r87cmNW1fiQG2SVufAQWbqz0lwcy2f8Lxb4bG+mRo64EtlOtCt/qMHt1i8b5Q
# Z7dsvfPxH2sMNgcWfzd8qVttevESRmCD1ycEvkvOl77DZypoEd+A5wwzZr8TDRRu
# 838fYxAe+o0bJW1sj6W3YQGx0qMmoRBxna3iw/nDmVG3KwcIzi7mULKn+gpFL6Lw
# 8jCCBmowggVSoAMCAQICEAOf7e3LeVuN7TIMiRnwNokwDQYJKoZIhvcNAQEFBQAw
# YjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBD
# QS0xMB4XDTEzMDUyMTAwMDAwMFoXDTE0MDYwNDAwMDAwMFowRzELMAkGA1UEBhMC
# VVMxETAPBgNVBAoTCERpZ2lDZXJ0MSUwIwYDVQQDExxEaWdpQ2VydCBUaW1lc3Rh
# bXAgUmVzcG9uZGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAumlK
# gU1vpRQWqorNZ75Lv8Zpj1gc4HnoHp1YJpjaXNR8o/nbK4wSNsP8+WQGsbvCqJgK
# Fw3hletAtOuWbZi/po95z7yKknttnBgGUdilGFMyAScZYeiEQd/G8OjK/netX9ie
# e4xgb4VcRr1r5w+AzucDw3wxz7dlVcb74JkI5HNa+5fa0Ey+tLbGD38mkqm4/Dju
# tOQ6pEjQTOqpRidbz5IRk5wWp/7SrR8ixR6swXHvvErbAQlE35gcLWe6qIoDM8lR
# tfcCTQmkTf6AXsXXRcN9CKoBM8wz2E8wFuT/IjIu63478PkeMuuVJdLy/m1UhLrV
# 5dTR3RuvvVl7lIUwAQIDAQABo4IDNTCCAzEwDgYDVR0PAQH/BAQDAgeAMAwGA1Ud
# EwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwggG/BgNVHSAEggG2MIIB
# sjCCAaEGCWCGSAGG/WwHATCCAZIwKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRp
# Z2ljZXJ0LmNvbS9DUFMwggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAAdQBz
# AGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAAYwBv
# AG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8AZgAg
# AHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4AZAAg
# AHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUAZQBt
# AGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwAaQB0
# AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQAIABo
# AGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJYIZIAYb9
# bAMVMB8GA1UdIwQYMBaAFBUAEisTmLKZB+0e36K+Vw0rZwLNMB0GA1UdDgQWBBRj
# L8nfeZJ7tSPKu+Gk7jN+4+Kd+jB9BgNVHR8EdjB0MDigNqA0hjJodHRwOi8vY3Js
# My5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDA4oDagNIYy
# aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5j
# cmwwdwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRENBLTEuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQCr
# dL1AAEx2FSVXPdMcA/99RchFEmbnKGVg2N87s/oNwawzj/SBuWHxnfuYVdfeR0O6
# gD3xSMw/ZzBWH8700EyEvYeknsXhD6gGXdAvbl7cGejwh+rgTq89bCCOc29+1ocY
# 4IbTmvye6oxy6UEPuHG1OCz4KbLVHKKdG+xfKrjcNyDhy7vw0GxspbPLn0r2VOMm
# ND0uuMErHLf2wz3+0S0eUPSUyPj97nPbSbUb9PX/pZDBORQb2O1xG2qY+/pAmkSp
# KQ5VXni4t6SDw3AB8GZA5a55NOErTQOhLebbVGIY7dUJi6Kq1gzITxq+mSV4aZmJ
# 1FmJ3t+I8NNnXnSlnaZEMIIGkDCCBXigAwIBAgIQBKVRftX3ANDrw0+OjYS9xjAN
# BgkqhkiG9w0BAQUFADBvMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQg
# SW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMS4wLAYDVQQDEyVEaWdpQ2Vy
# dCBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQS0xMB4XDTExMDkzMDAwMDAwMFoX
# DTE0MTAwODEyMDAwMFowXDELMAkGA1UEBhMCVVMxCzAJBgNVBAgTAlROMRIwEAYD
# VQQHEwlUdWxsYWhvbWExFTATBgNVBAoTDENhcmwgV2Vic3RlcjEVMBMGA1UEAxMM
# Q2FybCBXZWJzdGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAz2g4
# Kup2X6Mscbuq96HnetDDiITbncV1LtQ8Rxf8ZtN00+O/TliIZsWtufMq7GsLj1D8
# ikWfcgWGqMngWMsVYB4vdr1B8aQuHmKWld7W+j8FhKp3l+rNuFviTGa62sR6fEVW
# 1N6lDtJJHpfSIg/FUFfAqOKl0gFc45PU7iWCh08+oG5FJdhZ3WY0SosS1QujKEA4
# riSjeXPV6XSLsAHTE/fmHlGuu7NzJyMUzNNz2gPOFxYupHygbduhM5aAItD6GJ1h
# ajlovRt71tAMyeIPWNjj9B2luXxfRbgO9eufw91uFrXnougBPa7/eQ25YdW3NcGf
# tosYjvVI6Ptw/AaSiQIDAQABo4IDOTCCAzUwHwYDVR0jBBgwFoAUe2jOKarAF75J
# euHlP9an90WPNTIwHQYDVR0OBBYEFMHndyU+4pRT+JRECX9EG4y1laDkMA4GA1Ud
# DwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzBzBgNVHR8EbDBqMDOgMaAv
# hi1odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vYXNzdXJlZC1jcy0yMDExYS5jcmww
# M6AxoC+GLWh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9hc3N1cmVkLWNzLTIwMTFh
# LmNybDCCAcQGA1UdIASCAbswggG3MIIBswYJYIZIAYb9bAMBMIIBpDA6BggrBgEF
# BQcCARYuaHR0cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5
# Lmh0bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAg
# AHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0
# AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABE
# AGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABS
# AGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3
# AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBk
# ACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBu
# ACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjCBggYIKwYBBQUHAQEEdjB0MCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTAYIKwYBBQUHMAKG
# QGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENv
# ZGVTaWduaW5nQ0EtMS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQUFAAOC
# AQEAm1zhveo2Zy2lp8UNpR2E2CE8/NvEk0NDLszcBBuMda3N8Du23CikXCgrVvE0
# 3mMaeu/cIMDVU01ityLaqvDuovmTsvAKqaSJNztV9yTeWK9H4+h+35UEIU5TvYLs
# uzEW+rI5M2KcCXR6/LF9ZPmnBf9hHnK44hweHpmDWbo8HPqMatnIo7ideucuDn/D
# BM6s63eTMsFQCPYwte5vxuyVLqodOubLvIOMezZzByrpvJp9+gWAL151CE4qR6xQ
# jpgk5KqSkkkyvl72D+3PhNwZuxZDbZil5PIcrjmaBYoG8wfJzoNrtPFq3aG8dnQr
# xjXJjl+IN1iHYehBAUoBX98EozCCBqMwggWLoAMCAQICEA+oSQYV1wCgviF2/cXs
# bb0wDQYJKoZIhvcNAQEFBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGln
# aUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTExMDIxMTEyMDAwMFoXDTI2MDIx
# MDEyMDAwMFowbzELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZ
# MBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEuMCwGA1UEAxMlRGlnaUNlcnQgQXNz
# dXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EtMTCCASIwDQYJKoZIhvcNAQEBBQADggEP
# ADCCAQoCggEBAJx8+aCPCsqJS1OaPOwZIn8My/dIRNA/Im6aT/rO38bTJJH/qFKT
# 53L48UaGlMWrF/R4f8t6vpAmHHxTL+WD57tqBSjMoBcRSxgg87e98tzLuIZARR9P
# +TmY0zvrb2mkXAEusWbpprjcBt6ujWL+RCeCqQPD/uYmC5NJceU4bU7+gFxnd7XV
# b2ZklGu7iElo2NH0fiHB5sUeyeCWuAmV+UuerswxvWpaQqfEBUd9YCvZoV29+1aT
# 7xv8cvnfPjL93SosMkbaXmO80LjLTBA1/FBfrENEfP6ERFC0jCo9dAz0eotyS+BW
# tRO2Y+k/Tkkj5wYW8CWrAfgoQebH1GQ7XasCAwEAAaOCA0MwggM/MA4GA1UdDwEB
# /wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDAzCCAcMGA1UdIASCAbowggG2MIIB
# sgYIYIZIAYb9bAMwggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRpZ2ljZXJ0
# LmNvbS9zc2wtY3BzLXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIwggFWHoIB
# UgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQAaQBmAGkA
# YwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUAcAB0AGEA
# bgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMAUAAvAEMA
# UABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEAcgB0AHkA
# IABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkAdAAgAGwA
# aQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8AcgBwAG8A
# cgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIAZQBuAGMA
# ZQAuMBIGA1UdEwEB/wQIMAYBAf8CAQAweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUF
# BzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6
# Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5j
# cnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmw0LmRp
# Z2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwHQYDVR0OBBYE
# FHtozimqwBe+SXrh5T/Wp/dFjzUyMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA0GCSqGSIb3DQEBBQUAA4IBAQB7ch1k/4jIOsG36eepxIe725SS15BZ
# M/orh96oW4AlPxOPm4MbfEPE5ozfOT7DFeyw2jshJXskwXJduEeRgRNG+pw/alE4
# 3rQly/Cr38UoAVR5EEYk0TgPJqFhkE26vSjmP/HEqpv22jVTT8nyPdNs3CPtqqBN
# ZwnzOoA9PPs2TJDndqTd8jq/VjUvokxl6ODU2tHHyJFqLSNPNzsZlBjU1ZwQPNWx
# HBn/j8hrm574rpyZlnjRzZxRFVtCJnJajQpKI5JA6IbeIsKTOtSbaKbfKX8GuTwO
# vZ/EhpyCR0JxMoYJmXIJeUudcWn1Qf9/OXdk8YSNvosesn1oo6WQsQz/MIIGzTCC
# BbWgAwIBAgIQBv35A5YDreoACus/J7u6GzANBgkqhkiG9w0BAQUFADBlMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0Ew
# HhcNMDYxMTEwMDAwMDAwWhcNMjExMTEwMDAwMDAwWjBiMQswCQYDVQQGEwJVUzEV
# MBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29t
# MSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQDogi2Z+crCQpWlgHNAcNKeVlRcqcTSQQaPyTP8
# TUWRXIGf7Syc+BZZ3561JBXCmLm0d0ncicQK2q/LXmvtrbBxMevPOkAMRk2T7It6
# NggDqww0/hhJgv7HxzFIgHweog+SDlDJxofrNj/YMMP/pvf7os1vcyP+rFYFkPAy
# IRaJxnCI+QWXfaPHQ90C6Ds97bFBo+0/vtuVSMTuHrPyvAwrmdDGXRJCgeGDboJz
# PyZLFJCuWWYKxI2+0s4Grq2Eb0iEm09AufFM8q+Y+/bOQF1c9qjxL6/siSLyaxhl
# scFzrdfx2M8eCnRcQrhofrfVdwonVnwPYqQ/MhRglf0HBKIJAgMBAAGjggN6MIID
# djAOBgNVHQ8BAf8EBAMCAYYwOwYDVR0lBDQwMgYIKwYBBQUHAwEGCCsGAQUFBwMC
# BggrBgEFBQcDAwYIKwYBBQUHAwQGCCsGAQUFBwMIMIIB0gYDVR0gBIIByTCCAcUw
# ggG0BgpghkgBhv1sAAEEMIIBpDA6BggrBgEFBQcCARYuaHR0cDovL3d3dy5kaWdp
# Y2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5Lmh0bTCCAWQGCCsGAQUFBwICMIIB
# Vh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkA
# ZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAA
# dABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAA
# LwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIA
# dAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQA
# IABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIA
# cABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUA
# bgBjAGUALjALBglghkgBhv1sAxUwEgYDVR0TAQH/BAgwBgEB/wIBADB5BggrBgEF
# BQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBD
# BggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDig
# NoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNybDAdBgNVHQ4EFgQUFQASKxOYspkH7R7for5XDStnAs0wHwYDVR0jBBgw
# FoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQEFBQADggEBAEZQPsm3
# KCSnOB22WymvUs9S6TFHq1Zce9UNC0Gz7+x1H3Q48rJcYaKclcNQ5IK5I9G6OoZy
# rTh4rHVdFxc0ckeFlFbR67s2hHfMJKXzBBlVqefj56tizfuLLZDCwNK1lL1eT7EF
# 0g49GqkUW6aGMWKoqDPkmzmnxPXOHXh2lCVz5Cqrz5x2S+1fwksW5EtwTACJHvzF
# ebxMElf+X+EevAJdqP77BzhPDcZdkbkPZ0XN1oPt55INjbFpjE/7WeAjD9KqrgB8
# 7pxCDs+R1ye3Fu4Pw718CqDuLAhVhSK46xgaTfwqIa1JMYNHlXdx3LEbS0scEJx3
# FMGdTy9alQgpECYxggQAMIID/AIBATCBgzBvMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMS4wLAYD
# VQQDEyVEaWdpQ2VydCBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQS0xAhAEpVF+
# 1fcA0OvDT46NhL3GMAkGBSsOAwIaBQCgQDAZBgkqhkiG9w0BCQMxDAYKKwYBBAGC
# NwIBBDAjBgkqhkiG9w0BCQQxFgQURj5FMinMhSIb/+/cpHo4BgFzhI0wDQYJKoZI
# hvcNAQEBBQAEggEAxtk+ph0c/PtAJVHK1CFmOiL77JkrsmhiV3bDOCatk9crh1SO
# wRm4Mpp/8yra79K8zwik7M9x5WMc3rLzfyLdvwnjbAXnIW2yW5j+ZCQU3/b1YKiy
# aL+MepRyVrFIYV412aYJrqcWc+1CNtqIBawAXkyxAOWumV47v/if3jJtvJxVKaBS
# wRqFqLNnVSVZEhli/nEze0HJ4RC6TBg3arABtoEpkDZTQfdxWKvHJSuSzMv4IRaH
# Yr0tqEnlFP6PRTdMwoZs64RH2iV9nzpSjw4XG+Rh1CjHCYyOM6MJdd789jVpYkSB
# P4BEZsHzUoZCCB/eQhNv6ewQGSlYOn8oDfPfHqGCAg8wggILBgkqhkiG9w0BCQYx
# ggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IEFzc3VyZWQgSUQgQ0EtMQIQA5/t7ct5W43tMgyJGfA2iTAJBgUrDgMCGgUAoF0w
# GAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTMxMDI4
# MjE1ODM5WjAjBgkqhkiG9w0BCQQxFgQU3YlkECduG5oZY0gn8ZoaU88XQtwwDQYJ
# KoZIhvcNAQEBBQAEggEAn26pqPerViICP9X8ykQI5ZtVfFqSMPrLOX+6DOeC+rnf
# xpNl9oBohuO5cKM7cfPuuuNXnLZXoBO6O+bDwaNhOZJA93HocE4eZN3gQm5sBU0j
# Qnhm81wyf2ulrubOT0hTCg6DURQI5fkntApn3ZTkUSYqG8f/RIKbk+pXVgMOkzxm
# wbElyqjO4RCYpc5YoO+7s7cz6MuBWkr7qYQR/ORnC1Z6iG46RVWAQr5L5RLSPjxd
# T65fUKTDE8j+HOmzI5dDwleI8K6plPfjQJkH/3/RQBpBx1l53Vm4akP6dV2iIcMc
# 1ymyYQcwzhdZdfpN4OWjAbpK1tCLsM/yY1HaAKycOg==
# SIG # End signature block
