#Requires -Version 3.0
#This File is in Unicode format.  Do not edit in an ASCII editor.

#region help text

<#
.SYNOPSIS
	Creates an inventory of a XenServer 8.2 CU1 Pool.
.DESCRIPTION
	Creates a complete inventory of a XenServer 8.2 CU1 Pool using Microsoft Word, PDF, formatted 
	text, HTML, and PowerShell.
	
	The script requires at least PowerShell version 4 but runs best in version 5.

	Word is NOT needed to run the script. This script outputs in Text and HTML.
	The default output format is HTML.
	
	Creates an output file named Poolname.<fileextension>.
	
	Word and PDF documents include a Cover Page, Table of Contents, and Footer.
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

.PARAMETER ServerName
	Specifies which XenServer Pool to use to run the script against.
	
	You can enter the ServerName as the NetBIOS name, FQDN, or IP Address.
	
	If entered as an IP address, the script attempts o determine and use the actual 
	pool or Pool Master name.
	
	ServerName should be the Pool Master. If you use a Slave host, the script attempts 
	to determine the Pool Master and then makes a connection attempt to the Pool Master. 
	If successful, the script continues. If not successful, the script ends.
.PARAMETER User
	Username to use for the connection to the XenServer Pool.
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	
	HTML is the default report format.
	
	This parameter is set to True   If no other output format is selected.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	Text formatting is based on the default tab spacing of 8 by Microsoft Notepad.
	
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER Section
	Processes one or more sections of the report.
	Valid options are:
		Pool
		Host
		VM (Virtual Machines)
		All

	This parameter defaults to All sections.
	
	A comma separates multiple sections. -Section host, pool
.PARAMETER NoPoolMemory
	Excludes Pool Memory information from the output document.
	
	This Switch is helpful in large XenServer pools, where there may be many hosts.
	
	This parameter is disabled by default.
	This parameter has an alias of NPM.
.PARAMETER NoPoolStorage
	Excludes Pool Storage information from the output document.
	
	This Switch is helpful in large XenServer pools, where there may be many storage 
	repositories and hosts.
	
	This parameter is disabled by default.
	This parameter has an alias of NPS.
.PARAMETER NoPoolNetworking
	Excludes Pool Networking information from the output document.
	
	This Switch is helpful in large XenServer pools, where there may be many hosts.
	
	This parameter is disabled by default.
	This parameter has an alias of NPN.
.PARAMETER AddDateTime
	Adds a date timestamp to the end of the file name.
	
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2024, at 6 PM is 2024-06-01_1800.
	
	The output filename will be ReportName_2024-06-01_1800.<ext>.
	
	This parameter is disabled by default.
	This parameter has an alias of ADT.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER Log
	Generates a log file for troubleshooting.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.PARAMETER ReportFooter
	Outputs a footer section at the end of the report.

	This parameter has an alias of RF.
	
	Report Footer
		Report information:
			Created with: <Script Name> - Release Date: <Script Release Date>
			Script version: <Script Version>
			Started on <Date Time in Local Format>
			Elapsed time: nn days, nn hours, nn minutes, nn.nn seconds
			Ran from domain <Domain Name> by user <Username>
			Ran from the folder <Folder Name>

	Script Name and Script Release date are script-specific variables.
	Start Date Time in Local Format is a script variable.
	Elapsed time is a calculated value.
	Domain Name is $env:USERDNSDOMAIN.
	Username is $env:USERNAME.
	Folder Name is a script variable.
.PARAMETER MSWord
	SaveAs DOCX file
	
	Microsoft Word is no longer the default report format.
	This parameter is disabled by default.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	
	This parameter requires Microsoft Word to be installed.
	This parameter uses Word's SaveAs PDF capability.

	This parameter is disabled by default.
.PARAMETER CompanyAddress
	Company Address to use for the Cover Page   If the Cover Page has the Address field.
	
	The following Cover Pages have an Address field:
		Banded (Word 2013/2016)
		Contrast (Word 2010)
		Exposure (Word 2010)
		Filigree (Word 2013/2016)
		Ion (Dark) (Word 2013/2016)
		Retrospect (Word 2013/2016)
		Semaphore (Word 2013/2016)
		Tiles (Word 2010)
		ViewMaster (Word 2013/2016)
		
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CA.
.PARAMETER CompanyEmail
	Company Email to use for the Cover Page   If the Cover Page has the Email field. 
	
	The following Cover Pages have an Email field:
		Facet (Word 2013/2016)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CE.
.PARAMETER CompanyFax
	Company Fax to use for the Cover Page   If the Cover Page has the Fax field. 
	
	The following Cover Pages have a Fax field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CF.
.PARAMETER CompanyName
	Company Name to use for the Cover Page. 
	The default value is contained in 
	HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated 
	on the computer running the script.

	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CN.
.PARAMETER CompanyPhone
	Company Phone to use for the Cover Page if the Cover Page has the Phone field. 
	
	The following Cover Pages have a Phone field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CPh.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010, 2013, and 2016 are supported.
	(default cover pages in Word en-US)

	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly 
		works in 2010, but Subtitle/Subject & Author fields need moving
		after the title box is moved up)
		Banded (Word 2013/2016. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works   If you like looking sideways)
		Facet (Word 2013/2016. Works)
		Filigree (Word 2013/2016. Works)
		Grid (Word 2010/2013/2016. Works in 2010)
		Integral (Word 2013/2016. Works)
		Ion (Dark) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Ion (Light) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013/2016. Works   If the top date is manually changed 
		to 36 point)
		Newsprint (Word 2010. Works but the date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit; box needs to be manually 
		resized or font changed to 14 point)
		Retrospect (Word 2013/2016. Works)
		Semaphore (Word 2013/2016. Works)
		Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016. Works in 
		2010)
		Slice (Dark) (Word 2013/2016. Doesn't work)
		Slice (Light) (Word 2013/2016. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013/2016. Works)
		Whisp (Word 2013/2016. Works)

	The default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER UserName
	Username to use for the Cover Page and Footer.
	The default value is contained in $env:username
	This parameter has an alias of UN.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER SmtpPort
	Specifies the SMTP port for the SmtpServer. 
	The default is 25.
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report(s). 
	
	If From or To are used, this is a required parameter.
.PARAMETER From
	Specifies the username for the From email address.
	
	If SmtpServer or To are used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	
	If SmtpServer or From are used, this is a required parameter.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	The default is False.
.EXAMPLE
	PS C:\PSScript > .\XS_Inventory.ps1
	
	Outputs, by default, to HTML.
	Prompts for the XenServer Pool and login credentials.
.EXAMPLE
	PS C:\PSScript .\XS_Inventory.ps1 -MSWord -CompanyName "Carl Webster 
	Consulting" -CoverPage "Mod" -UserName "Carl Webster" -ServerName XS01

	Uses:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
		XenServer host named XS01 for the ServerName.

	Outputs to Microsoft Word.
	Prompts for the XenServer Pool login credentials.
.EXAMPLE
	PS C:\PSScript .\XS_Inventory.ps1 -PDF -CN "Carl Webster Consulting" -CP 
	"Mod" -UN "Carl Webster"

	Uses:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).

	Outputs to PDF.
	Prompts for the XenServer Pool and login credentials.
.EXAMPLE
	PS C:\PSScript .\XS_Inventory.ps1 -CompanyName "Sherlock Holmes 
	Consulting" -CoverPage Exposure -UserName "Dr. Watson" -CompanyAddress "221B Baker 
	Street, London, England" -CompanyFax "+44 1753 276600" -CompanyPhone "+44 1753 276200" 
	-MSWord
	
	Uses:
		Sherlock Holmes Consulting for the Company Name.
		Exposure for the Cover Page format.
		Dr. Watson for the User Name.
		221B Baker Street, London, England, for the Company Address.
		+44 1753 276600 for the Company Fax.
		+44 1753 276200 for the Company Phone.

	Outputs to Microsoft Word.
	Prompts for the XenServer Pool and login credentials.
.EXAMPLE
	PS C:\PSScript .\XS_Inventory.ps1 -CompanyName "Sherlock Holmes 
	Consulting" -CoverPage Facet -UserName "Dr. Watson" -CompanyEmail 
	SuperSleuth@SherlockHolmes.com
	-PDF

	Uses:
		Sherlock Holmes Consulting for the Company Name.
		Facet for the Cover Page format.
		Dr. Watson for the User Name.
		SuperSleuth@SherlockHolmes.com for the Company Email.

	Outputs to PDF.
	Prompts for the XenServer Pool and login credentials.
.EXAMPLE
	PS C:\PSScript >.\XS_Inventory.ps1 -Dev -ScriptInfo -Log
	
	Creates an HTML report.
	
	Creates a text file named XSInventoryScriptErrors_yyyyMMddTHHmmssffff.txt that 
	contains up to the last 250 errors reported by the script.
	
	Creates a text file named XSInventoryScriptInfo_yyyy-MM-dd_HHmm.txt that 
	contains all the script parameters and other basic information.
	
	Creates a text file for transcript logging named 
	XSDocScriptTranscript_yyyyMMddTHHmmssffff.txt.
	Prompts for the XenServer Pool and login credentials.
.EXAMPLE
	PS C:\PSScript >.\XS_Inventory.ps1 -Section Pool
	
	Creates an HTML report that contains only Pool information.
	Processes only the Pool section of the report.
	Prompts for the XenServer Pool and login credentials.
.EXAMPLE
	PS C:\PSScript >.\XS_Inventory.ps1 -ServerName PoolMaster.domain.com -Section Pool 
	-NoPoolMemory -NoPoolStorage -NoPoolNetworking
	
	Creates an HTML report that contains only Pool information but with no Memory, Storage, 
	or Networking data.
	
	Processes only the Pool section of the report.
	Prompts for the XenServer Pool and login credentials.
.EXAMPLE
	PS C:\PSScript > .\XS_Inventory.ps1 -Section Pool, Host

	Creates an HTML report.

	The report includes only the Pool and Host sections.
	Prompts for the XenServer Pool and login credentials.
.EXAMPLE
	PS C:\PSScript > .\XS_Inventory.ps1 -SmtpServer mail.domain.tld -From 
	XSAdmin@domain.tld -To ITGroup@domain.tld -Text

	The script uses the email server mail.domain.tld, sending from XSAdmin@domain.tld 
	and sending to ITGroup@domain.tld.

	The script uses the default SMTP port 25 and does not use SSL.

	If the current user's credentials are not valid to send an email, the script prompts 
	the user to enter valid credentials.

	Outputs to a text file.
	Prompts for the XenServer Pool and login credentials.
.EXAMPLE
	PS C:\PSScript > .\XS_Inventory.ps1 -SmtpServer mailrelay.domain.tld -From 
	Anonymous@domain.tld -To ITGroup@domain.tld	

	***SENDING UNAUTHENTICATED EMAIL***

	The script uses the email server mailrelay.domain.tld, sending from 
	anonymous@domain.tld and sending to ITGroup@domain.tld.

	To send an unauthenticated email using an email relay server requires the From email 
	account to use the name Anonymous.

	The script uses the default SMTP port 25 and does not use SSL.
	
	***GMAIL/G SUITE SMTP RELAY***
	https://support.google.com/a/answer/2956491?hl=en
	https://support.google.com/a/answer/176600?hl=en

	To send an email using a Gmail or G-suite account, you may have to turn ON the "Less 
	secure app access" option on your account.
	***GMAIL/G SUITE SMTP RELAY***

	The script generates an anonymous, secure password for the anonymous@domain.tld 
	account.

	Outputs, by default, to HTML.
	Prompts for the XenServer Pool and login credentials.
.EXAMPLE
	PS C:\PSScript > .\XS_Inventory.ps1 -SmtpServer 
	labaddomain-com.mail.protection.outlook.com -UseSSL -From 
	SomeEmailAddress@labaddomain.com -To ITGroupDL@labaddomain.com	

	***OFFICE 365 Example***

	https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-office-3
	
	This uses Option 2 from the above link.
	
	***OFFICE 365 Example***

	The script uses the email server labaddomain-com.mail.protection.outlook.com, sending 
	from SomeEmailAddress@labaddomain.com and sending to ITGroupDL@labaddomain.com.

	The script uses the default SMTP port 25 and SSL.

	Outputs, by default, to HTML.
	Prompts for the XenServer Pool and login credentials.
.EXAMPLE
	PS C:\PSScript > .\XS_Inventory.ps1 -SmtpServer smtp.office365.com -SmtpPort 587
	-UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	The script uses the email server smtp.office365.com on port 587 using SSL, sending from 
	webster@carlwebster.com and sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send an email, the script prompts 
	the user to enter valid credentials.

	Outputs, by default, to HTML.
	Prompts for the XenServer Pool and login credentials.
.EXAMPLE
	PS C:\PSScript > .\XS_Inventory.ps1 -SmtpServer smtp.gmail.com -SmtpPort 587
	-UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	*** NOTE ***
	To send an email using a Gmail or G-suite account, you may have to turn ON the "Less 
	secure app access" option on your account.
	*** NOTE ***
	
	The script uses the email server smtp.gmail.com on port 587 using SSL, sending from 
	webster@gmail.com and sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send an email, the script prompts 
	the user to enter valid credentials.

	Outputs, by default, to HTML.
	Prompts for the XenServer Pool and login credentials.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script. This script creates a Word, PDF, HTML, or plain 
	text document.
.NOTES
	NAME: XS_Inventory.ps1
	VERSION: 0.021
	AUTHOR: Carl Webster and John Billekens along with help from Michael B. Smith, Guy Leech, and the XenServer team
	LASTEDIT: July 28, 2023
#>

#endregion

#region script parameters
#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Word") ]

Param(
	[parameter(Mandatory = $True)] 
	[string]$ServerName = "",
	
	[parameter(Mandatory = $False)] 
	[string]$User = "",
	
	[parameter(Mandatory = $False)] 
	[Switch]$HTML = $False,

	[parameter(Mandatory = $False)] 
	[Switch]$Text = $False,

	[parameter(Mandatory = $False)] 
	[string]$Folder = "",
	
	[ValidateSet('All', 'Pool', 'Host', 'VM')]
	[parameter(Mandatory = $False)] 
	[String[]] $Section = 'All',
	
	[parameter(Mandatory = $False)] 
	[Alias("NPM")]
	[Switch]$NoPoolMemory = $False,	
	
	[parameter(Mandatory = $False)] 
	[Alias("NPS")]
	[Switch]$NoPoolStorage = $False,	
	
	[parameter(Mandatory = $False)] 
	[Alias("NPN")]
	[Switch]$NoPoolNetworking = $False,	
	
	[parameter(Mandatory = $False)] 
	[Alias("ADT")]
	[Switch]$AddDateTime = $False,
	
	[parameter(Mandatory = $False)] 
	[Switch]$Dev = $False,
	
	[parameter(Mandatory = $False)] 
	[Switch]$Log = $False,
	
	[parameter(Mandatory = $False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo = $False,
	
	[parameter(Mandatory = $False)] 
	[Alias("RF")]
	[Switch]$ReportFooter = $False,

	[parameter(ParameterSetName = "WordPDF", Mandatory = $False)] 
	[Switch]$MSWord = $False,

	[parameter(ParameterSetName = "WordPDF", Mandatory = $False)] 
	[Switch]$PDF = $False,

	[parameter(ParameterSetName = "WordPDF", Mandatory = $False)] 
	[Alias("CA")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyAddress = "",
    
	[parameter(ParameterSetName = "WordPDF", Mandatory = $False)] 
	[Alias("CE")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyEmail = "",
    
	[parameter(ParameterSetName = "WordPDF", Mandatory = $False)] 
	[Alias("CF")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyFax = "",
    
	[parameter(ParameterSetName = "WordPDF", Mandatory = $False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName = "",
    
	[parameter(ParameterSetName = "WordPDF", Mandatory = $False)] 
	[Alias("CPh")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyPhone = "",
    
	[parameter(ParameterSetName = "WordPDF", Mandatory = $False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage = "Sideline", 

	[parameter(ParameterSetName = "WordPDF", Mandatory = $False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName = $env:username,

	[parameter(Mandatory = $False)] 
	[int]$SmtpPort = 25,

	[parameter(Mandatory = $False)] 
	[string]$SmtpServer = "",

	[parameter(Mandatory = $False)] 
	[string]$From = "",

	[parameter(Mandatory = $False)] 
	[string]$To = "",

	[parameter(Mandatory = $False)] 
	[switch]$UseSSL = $False
	
)
#endregion

#region script change log	
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on June 27, 2023
#
#.021
#	Updated Function OutputHostGPUProperties with code from the XS team (Webster)
#
#.020
#	Added Function OutputHostGeneralOverview (Webster)
#		This function is for what you see when looking at Server General Properties, not a host's Properties, General
#		Function OutputHostGeneral is for the host's Properties, General
#	Added Function OutputHostUpdates (Webster)
#	Added Function OutputHostMemoryOverview (Webster)
#		This function is for what you see when looking at Server General Properties, not a host's Properties, General
#		Function OutputHostMemory is for the host's Properties, General
#	Added Function OutputHostGPUProperties (Webster)
#		This function is for what you see when looking at a host's Properties, GPU
#	In Function OutputPoolStorage, change the following: (JohnB)
#       	Changed OutputHostStorage to gather the data but not show output when specified $NoPoolStorage
#	In Function OutputPoolNetworking, change the following: (JohnB)
#       	Changed OutputHostNetworking to gather the data but not show output when specified $NoPoolNetworking
#	In Function OutputPoolMemory, change the following: (JohnB)
#       	Moved logic from OutputHostMemory to OutputPoolMemory changed it for the pool and save it as script variable
#       	Changed OutputHostMemory to get data from Script variable as it's the same data
#	Reordered the OutputHost____ Functions to match the order in the console (Webster)
#	Updated Function OutputVMCPU to handle the case where the VM isn't storing the default CPU priority value of 256 (Webster)
#
#.019
#	Updated the help text by fixing grammar and spelling issues and adding a new example (Webster)
#	Updated the ReadMe file by adding a "What the script documents" section (Webster)
#
#.018
#	Added Switch Parameters NoPoolMemory, NoPoolStorage, and NoPoolNetworking (Webster)
#	Updated Function OutputPoolUsers with data (Webster and JohnB)
#	Updated Function OutputPoolHA with data (Webster)
#	Updated Function OutputPoolWLB with data I hope is correct since I don't have a WLB appliance (Webster)
#	In Function OutputPoolStorage, change the following: (JohnB)
#       	Moved logic from OutputHostStorage to OutputPoolStorage changed it for the pool and save it as script variable
#       	Changed OutputHostStorage to get data from Script variable as it's the same data
#	In Function OutputPoolNetworking, change the following: (JohnB)
#       	Moved logic from OutputHostNetworking to OutputPoolNetworking changed it for the pool and save it as script variable
#       	Changed OutputHostNetworking to get data from Script variable as it's the same data
#
#.017
#	Add Function OutputPoolGeneralOverview
#		This function is for what you see when looking at Pool General Properties, not a pool's Properties, General
#		Function OutputPoolGeneral is for the pool's Properties, General
#	Rearranged the Pool output functions to the order seen in the Console:
#		Pool General Properties
#			General
#			Updates
#			Management Interfaces
#		Pool Properties
#			General
#			Custom Fields
#			Email Options
#			Power On
#			Live Patching
#			Network Options
#			Clustering
#	Added stub Functions for the remaining Pool tabs
#		Function OutputPoolMemory
#		Function OutputPoolStorage
#		Function OutputPoolNetworking
#		Function OutputPoolGPU
#		Function OutputPoolHA
#		Function OutputPoolWLB
#		Function OutputPoolUsers
#	Updated Functions OutputPoolLivePatching, OutputPoolNetworkOptions, and OutputPoolClustering 
#		to change the MSWord/PDF/HTML output to tables
#	Updated Function OutputPoolEmailOptions for better output
#	Updated the help text for the ServerName parameter
#
#.016
#	Change the Disconnect message from "Disconnect from XenServer to
#		Disconnect from Pool Master $Script:ServerName (Webster)
#	In Function ProcessScriptSetup, handle the scenario where a non-Pool Master is entered (Webster)
#
#.015
#	In Function OutputVMHomeServer, handle a VM whose power_state is not running (Webster)
#		Add the message: VM's power state is $($vm.power_state). Unable to determine the running host.
#	In Function OutputVMStorage, change the following: (Webster)
#		$storages = $storages | Sort-Object -Property Position, Name to
#		$storages = @($storages | Sort-Object -Property Position, Name)
#		To prevent the error:
#			The property 'Count' cannot be found on this object. Verify that the property exists.
#			$storageCount = $storages.Count
#	In Function OutputVMStorage, change the following: (JohnB)
#           Configured the priority value
#
#.014
#   Modified the following Functions
#       OutputHostNICs (added fcoe and sriov, JohnB)
#       OutputVMNIC (folowed XenCenter output, JohnB)
#       OutputHostNetworking (Sorting adjusted, JohnB)
#       OutputVM (small change to core output, JohnB)
#       OutputVMCPU (small change to core output, JohnB)
#.013
#	Updated Function OutputVMBootOptions with data (Webster)
#		Some code written by John borrowed from Function OutputVM
#	Updated Function OutputVMCPU with data (Webster)
#		Some code written by John borrowed from Function OutputVM
#	Updated Function OutputVMHomeServer with the correct data (Webster)
#
#.012
#	Added testing notes to Function OutputVMCPU (Webster)
#	Updated Function OutputVMAdvancedOptions with data (Webster)
#
#.011
#	Updated Function OutputVMHomeServer with data (Webster)
#
#.010
#	Add these Functions (Webster)
#		OutputVMCPU
#		OutputVMBootOptions
#		OutputVMStartOptions (with some data that I can find)
#		OutputVMAlerts (with data)
#		OutputVMHomeServer
#		OutputVMAdvancedOptions
#	Rearrange the order of the VM functions
#
#.009
#	Add output to Function OutputHostAlerts (Webster)
#	Add output to Function OutputHostLogDestination (Webster)
#
#.008
#	Minor cleanup of console output (Webster)
#	Minor updates to the HTML, Text, and MSWord/PDF output (Webster)
#		In MSWord/PDF output, start 2nd+ Hosts and VMs on a new page (Webster)
#	Updated the help text (Webster)
#	Added the following Functions:
#       OutputVMSnapshots (JohnB)
#
#.007
#	Added data to the following Functions:
#		OutputHostLicense
#		OutputHostVersion
#		OutputHostManagement
#   Modified the following Functions
#       OutputPoolUpdates
#       OutputPoolGeneral
#       OutputHostGeneral
#	Added the following Functions:
#       Get-CustomFields
#		OutputHostNICs
#		OutputHostNetworking
#       OutputHostStorage
#       OutputVMStorage
#
#.006
#	Added Citrix CTA John Billekens as a script coauthor
#	Added data to the following Functions:
#		OutputHostGeneral
#		OutputPoolClustering
#		OutputHostGPU
#		ProcessVMs
#		Output VM
#	Added the following Functions:
#		OutputHostPIF
#		OutputVMGPU
#		OutputVMNIC
#
#.005
#	For the Host report section
#		Renamed Function OutputHost to OutputHostGeneral
#		Put as much data as I can find in the Host General section
#
#.004
#	For the Pool report section
#		Added Function OutputPoolPowerOn
#		Added Function OutputPoolLivePatching
#		Added Function OutputPoolNetworkOptions
#		Added Function OutputPoolClustering
#	For the Host report section
#		Added Function OutputHostPowerOn
#		Added the following placeholder functions
#			OutputHostAlerts
#			OutputHostMultipathing
#			OutputHostLogDestination
#			OutputHostGPU
#	In Function ProcessScriptSetup
#		Add getting the pool and hosts by the session's opaque_ref to get all hosts associated with the pool
#	
#.003
#	For the Pool report section
#		Renamed Fnction OutputPool to OutputPoolGeneral
#		Added Function OutputPoolCustomFields
#		Added Function OutputPoolEmailOptions
#		Added Function OutputPoolManagementInterfaces
#			Includes sample code from the XenServer team
#	For the Host report section
#		Added Function OutputHostCustomFields
#	For the VM report section
#		Added Function OutputVMCustomFields
#	In Function ProcessScriptSetup
#		Sort Hosts and VMs by Name_Label so output is in sorted order
#	Some general console and report output cleanup
#
#.002
#	Pool section, 
#		Added "(version x)" to the update name. Example CH82ECU1 (version 1.0)
#		Fixed handling Tags where the default is "<None>"
#.001 - initial version create from the May 2015 attempt
#endregion

#region basics
Function AbortScript
{
	If ($MSWord -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): System Cleanup"
		If (Test-Path variable:global:word)
		{
			$Script:Word.quit()
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
			Remove-Variable -Name word -Scope Global 4>$Null
		}
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()

	If ($MSWord -or $PDF)
	{
		#is the winword Process still running? kill it

		#find out our session (usually "1" except on TS/RDC or Citrix)
		$SessionID = (Get-Process -PID $PID).SessionId

		#Find out   If winword running in our session
		$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object { $_.SessionId -eq $SessionID }) | Select-Object -Property Id 
		If ( $wordprocess -and $wordprocess.Id -gt 0)
		{
			Write-Verbose "$(Get-Date -Format G): WinWord Process is still running. Attempting to stop WinWord Process # $($wordprocess.Id)"
		}
	}
	
	Write-Verbose "$(Get-Date -Format G): Script has been aborted"
	#stop transcript logging
	If ($Log -eq $True)
	{
		If ($Script:StartLog -eq $True)
		{
			try
			{
				Stop-Transcript | Out-Null
				Write-Verbose "$(Get-Date -Format G): $Script:LogPath is ready for use"
			}
			catch
			{
				Write-Verbose "$(Get-Date -Format G): Transcript/log stop failed"
			}
		}
	}
	$ErrorActionPreference = $SaveEAPreference
	Exit
}

Set-StrictMode -Version 4

#force  on
$PSDefaultParameterValues = @{"*:Verbose" = $True }
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'
$Error.Clear()

$Script:emailCredentials = $Null
$script:MyVersion = '0.020'
$Script:ScriptName = "XS_Inventory.ps1"
$tmpdate = [datetime] "07/28/2023"
$Script:ReleaseDate = $tmpdate.ToUniversalTime().ToShortDateString()

If ($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$HTML = $True
}

If ($MSWord)
{
	Write-Verbose "$(Get-Date -Format G): MSWord is set"
}
If ($PDF)
{
	Write-Verbose "$(Get-Date -Format G): PDF is set"
}
If ($Text)
{
	Write-Verbose "$(Get-Date -Format G): Text is set"
}
If ($HTML)
{
	Write-Verbose "$(Get-Date -Format G): HTML is set"
}

$ValidSection = $False
#there are no Break statements since there can be multiple sections entered
Switch ($Section)
{
	"Pool"	{ $ValidSection = $True }
	"Host"	{ $ValidSection = $True }
	"VM"	{ $ValidSection = $True }
	"All"	{ $ValidSection = $True }
}

If ($ValidSection -eq $False)
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Error -Message "
	`n`n
	`t`t
	The Section parameter specified, $Section, is an invalid Section option.
	`n`n
	`t`t
	Valid options are:

	`tPool
	`tHost
	`tVM (Virtual Machines)
	`tAll
	
	`t`t
	Script cannot continue.
	`n`n
	"
	Exit
}

If ($Folder -ne "")
{
	Write-Verbose "$(Get-Date -Format G): Testing folder path"
	#does it exist
	If (Test-Path $Folder -EA 0)
	{
		#it exists, now check to see   If it is a folder and not a file
		If (Test-Path $Folder -PathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Verbose "$(Get-Date -Format G): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
			#Do not indent the following write-error lines. Doing so will mess up the console formatting of the error message.
			Write-Error "
			`n`n
	Folder $Folder is a file, not a folder.
			`n`n
	Script cannot continue.
			`n`n"
			Exit
		}
	}
	Else
	{
		#does not exist
		Write-Error "
		`n`n
	Folder $Folder does not exist.
		`n`n
	Script cannot continue.
		`n`n
		"
		Exit
	}
}

If ($Folder -eq "")
{
	$Script:pwdpath = $pwd.Path
}
Else
{
	$Script:pwdpath = $Folder
}

If ($Script:pwdpath.EndsWith("\"))
{
	#remove the trailing \
	$Script:pwdpath = $Script:pwdpath.SubString(0, ($Script:pwdpath.Length - 1))
}

If ($Log)
{
	#start transcript logging
	$Script:LogPath = "$Script:pwdpath\XSDocScriptTranscript_$(Get-Date -f FileDateTime).txt"
	
	try
	{
		Start-Transcript -Path $Script:LogPath -Force -Verbose:$false | Out-Null
		Write-Verbose "$(Get-Date -Format G): Transcript/log started at $Script:LogPath"
		$Script:StartLog = $true
	}
 catch
	{
		Write-Verbose "$(Get-Date -Format G): Transcript/log failed at $Script:LogPath"
		$Script:StartLog = $false
	}
}

If ($Dev)
{
	$Error.Clear()
	$Script:DevErrorFile = "$Script:pwdpath\XSInventoryScriptErrors_$(Get-Date -f FileDateTime).txt"
}

If (![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer but did not include a From or To email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If (![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a To email address but did not include a From email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If (![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($To) -and ![String]::IsNullOrEmpty($From))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a From email address but did not include a To email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If (![String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified From and To email addresses but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If (![String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified a From email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If (![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified a To email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
#endregion

#region initialize variables for Word, HTML, and text
[string]$Script:RunningOS = (Get-WmiObject -Class Win32_OperatingSystem -EA 0).Caption

If ($MSWord -or $PDF)
{
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	#[int]$wdColorBlack            = 0
	#[int]$wdColorGray05           = 15987699 
	[int]$wdColorGray15 = 14277081
	#[int]$wdColorRed              = 255
	#[int]$wdColorWhite            = 16777215
	#[int]$wdColorYellow           = 65535
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdWord2016 = 16
	[int]$wdFormatDocumentDefault = 16
	[int]$wdFormatPDF = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	#[int]$wdAlignParagraphLeft = 0
	#[int]$wdAlignParagraphCenter = 1
	#[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	#[int]$wdCellAlignVerticalTop = 0
	#[int]$wdCellAlignVerticalCenter = 1
	#[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
	#[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	#[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
	#[int]$wdAdjustFirstColumn = 2
	#[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
	#[int]$Indent1TabStops = 1 * $PointsPerTabStop
	#[int]$Indent2TabStops = 2 * $PointsPerTabStop
	#[int]$Indent3TabStops = 3 * $PointsPerTabStop
	#[int]$Indent4TabStops = 4 * $PointsPerTabStop

	# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleHeading5 = -6
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155
	#[int]$wdTableLightListAccent3 = -206

	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	#[int]$wdHeadingFormatFalse = 0 
}

If ($HTML)
{
	$Script:htmlredmask = "#FF0000" 4>$Null
	$Script:htmlcyanmask = "#00FFFF" 4>$Null
	$Script:htmlbluemask = "#0000FF" 4>$Null
	$Script:htmldarkbluemask = "#0000A0" 4>$Null
	$Script:htmllightbluemask = "#ADD8E6" 4>$Null
	$Script:htmlpurplemask = "#800080" 4>$Null
	$Script:htmlyellowmask = "#FFFF00" 4>$Null
	$Script:htmllimemask = "#00FF00" 4>$Null
	$Script:htmlmagentamask = "#FF00FF" 4>$Null
	$Script:htmlwhitemask = "#FFFFFF" 4>$Null
	$Script:htmlsilvermask = "#C0C0C0" 4>$Null
	$Script:htmlgraymask = "#808080" 4>$Null
	$Script:htmlblackmask = "#000000" 4>$Null
	$Script:htmlorangemask = "#FFA500" 4>$Null
	$Script:htmlmaroonmask = "#800000" 4>$Null
	$Script:htmlgreenmask = "#008000" 4>$Null
	$Script:htmlolivemask = "#808000" 4>$Null

	$Script:htmlbold = 1 4>$Null
	$Script:htmlitalics = 2 4>$Null
	$Script:htmlred = 4 4>$Null
	$Script:htmlcyan = 8 4>$Null
	$Script:htmlblue = 16 4>$Null
	$Script:htmldarkblue = 32 4>$Null
	$Script:htmllightblue = 64 4>$Null
	$Script:htmlpurple = 128 4>$Null
	$Script:htmlyellow = 256 4>$Null
	$Script:htmllime = 512 4>$Null
	$Script:htmlmagenta = 1024 4>$Null
	$Script:htmlwhite = 2048 4>$Null
	$Script:htmlsilver = 4096 4>$Null
	$Script:htmlgray = 8192 4>$Null
	$Script:htmlolive = 16384 4>$Null
	$Script:htmlorange = 32768 4>$Null
	$Script:htmlmaroon = 65536 4>$Null
	$Script:htmlgreen = 131072 4>$Null
	$Script:htmlblack = 262144 4>$Null

	$Script:htmlsb = ( $htmlsilver -bor $htmlBold ) ## point optimization

	$Script:htmlColor = 
	@{
		$htmlred       = $htmlredmask
		$htmlcyan      = $htmlcyanmask
		$htmlblue      = $htmlbluemask
		$htmldarkblue  = $htmldarkbluemask
		$htmllightblue = $htmllightbluemask
		$htmlpurple    = $htmlpurplemask
		$htmlyellow    = $htmlyellowmask
		$htmllime      = $htmllimemask
		$htmlmagenta   = $htmlmagentamask
		$htmlwhite     = $htmlwhitemask
		$htmlsilver    = $htmlsilvermask
		$htmlgray      = $htmlgraymask
		$htmlolive     = $htmlolivemask
		$htmlorange    = $htmlorangemask
		$htmlmaroon    = $htmlmaroonmask
		$htmlgreen     = $htmlgreenmask
		$htmlblack     = $htmlblackmask
	}
}
#endregion

#region word specific functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. Smith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese
	
	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2'; Break }
			'da-'	{ 'Automatisk tabel 2'; Break }
			#'de-'	{ 'Automatische Tabelle 2'; Break }
			'de-'	{ 'Automatisches Verzeichnis 2'; Break } #changed 6-feb-2022 rene bigler
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automática 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
			'fr-'	{ 'Table automatique 2'; Break }
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			'sv-'	{ 'Automatisk innehållsförteckn2'; Break }
			'zh-'	{ '自动目录 2'; Break }
		}
	)

	$Script:myHash = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1 = $wdStyleheading1
	$Script:myHash.Word_Heading2 = $wdStyleheading2
	$Script:myHash.Word_Heading3 = $wdStyleheading3
	$Script:myHash.Word_Heading4 = $wdStyleheading4
	$Script:myHash.Word_Heading5 = $wdStyleheading5
	$Script:myHash.Word_TableGrid = $wdTableGrid
}

Function GetCulture
{
	Param([int]$WordValue)
	
	#codes obtained from http://support.microsoft.com/kb/221435
	#http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$ChineseArray = 2052, 3076, 5124, 4100
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese

	Switch ($WordValue)
	{
		{ $CatalanArray -contains $_ }	{ $CultureCode = "ca-" }
		{ $ChineseArray -contains $_ }	{ $CultureCode = "zh-" }
		{ $DanishArray -contains $_ } { $CultureCode = "da-" }
		{ $DutchArray -contains $_ } { $CultureCode = "nl-" }
		{ $EnglishArray -contains $_ }	{ $CultureCode = "en-" }
		{ $FinnishArray -contains $_ }	{ $CultureCode = "fi-" }
		{ $FrenchArray -contains $_ } { $CultureCode = "fr-" }
		{ $GermanArray -contains $_ } { $CultureCode = "de-" }
		{ $NorwegianArray -contains $_ }	{ $CultureCode = "nb-" }
		{ $PortugueseArray -contains $_ }	{ $CultureCode = "pt-" }
		{ $SpanishArray -contains $_ }	{ $CultureCode = "es-" }
		{ $SwedishArray -contains $_ }	{ $CultureCode = "sv-" }
		Default { $CultureCode = "en-" }
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'
		{
			If ($xWordVersion -eq $wdWord2016)
			{
				$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
			}
			ElseIf ($xWordVersion -eq $wdWord2013)
			{
				$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
			}
			ElseIf ($xWordVersion -eq $wdWord2010)
			{
				$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
			}
		}

		'da-'
		{
			If ($xWordVersion -eq $wdWord2016)
			{
				$xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
			}
			ElseIf ($xWordVersion -eq $wdWord2013)
			{
				$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
			}
			ElseIf ($xWordVersion -eq $wdWord2010)
			{
				$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
			}
		}

		'de-'
		{
			If ($xWordVersion -eq $wdWord2016)
			{
				$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "Rückblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
			}
			ElseIf ($xWordVersion -eq $wdWord2013)
			{
				$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
			}
			ElseIf ($xWordVersion -eq $wdWord2010)
			{
				$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
			}
		}

		'en-'
		{
			If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
			{
				$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
			}
			ElseIf ($xWordVersion -eq $wdWord2010)
			{
				$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
			}
		}

		'es-'
		{
			If ($xWordVersion -eq $wdWord2016)
			{
				$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
					"Semáforo", "Slice (luz)", "Vista principal", "Whisp")
			}
			ElseIf ($xWordVersion -eq $wdWord2013)
			{
				$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
			}
			ElseIf ($xWordVersion -eq $wdWord2010)
			{
				$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
			}
		}

		'fi-'
		{
			If ($xWordVersion -eq $wdWord2016)
			{
				$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
			}
			ElseIf ($xWordVersion -eq $wdWord2013)
			{
				$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
			}
			ElseIf ($xWordVersion -eq $wdWord2010)
			{
				$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
			}
		}

		'fr-'
		{
			If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
			{
				$xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
					"Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
					"Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
			}
			ElseIf ($xWordVersion -eq $wdWord2010)
			{
				$xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
					"Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latérale", "Moderne", 
					"Mosaïques", "Mots croisés", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
			}
		}

		'nb-'
		{
			If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
			{
				$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
			}
			ElseIf ($xWordVersion -eq $wdWord2010)
			{
				$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
			}
		}

		'nl-'
		{
			If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
			{
				$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
			}
			ElseIf ($xWordVersion -eq $wdWord2010)
			{
				$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
			}
		}

		'pt-'
		{
			If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
			{
				$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
			}
			ElseIf ($xWordVersion -eq $wdWord2010)
			{
				$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
			}
		}

		'sv-'
		{
			If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
			{
				$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
			}
			ElseIf ($xWordVersion -eq $wdWord2010)
			{
				$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
			}
		}

		'zh-'
		{
			If ($xWordVersion -eq $wdWord2010 -or $xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
			{
				$xArray = ('奥斯汀', '边线型', '花丝', '怀旧', '积分',
					'离子(浅色)', '离子(深色)', '母版型', '平面', '切片(浅色)',
					'切片(深色)', '丝状', '网格', '镶边', '信号灯',
					'运动型')
			}
		}

		Default
		{
			If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
			{
				$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
			}
			ElseIf ($xWordVersion -eq $wdWord2010)
			{
				$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
			}
		}
	}
	
	If ($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function CheckWordPrereq
{
	If ((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		Write-Host "
		`n
		This script directly outputs to Microsoft Word, please install Microsoft Word
		`n"
		AbortScript
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out   If winword is running in our session
	#fixed by MBS
	[bool]$wordrunning = $null –ne ((Get-Process 'WinWord' -ea 0) | Where-Object { $_.SessionId -eq $SessionID })
	If ($wordrunning)
	{
		Write-Host "
		`n
		Please close all instances of Microsoft Word before running this report.
		`n"
		AbortScript
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If ($xResult)
	{
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If ($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

Function Check-LoadedModule
#Function created by Jeff Wouters
#@JeffWouters on Twitter
#modified by Michael B. Smith to handle when the module doesn't exist on server
#modified by @andyjmorgan
#bug fixed by @schose
#bug fixed by Peter Bosen
#This Function handles all three scenarios:
#
# 1. Module is already imported into current session
{
 # 2. Module is not already imported into current session, it does exists on the server and is imported
 # 3. Module does not exist on the server
	Param([parameter(Mandatory = $True)][alias("Module")][string]$ModuleName)
	#$LoadedModules = Get-Module | Select Name
	#following line changed at the recommendation of @andyjmorgan
	$LoadedModules = Get-Module | ForEach-Object { $_.Name.ToString() }
	#bug reported on 21-JAN-2013 by @schose 
	#the following line did not work   If the citrix.grouppolicy.commands.psm1 module
	#was manually loaded from a non Default folder
	#$ModuleFound = (!$LoadedModules -like "*$ModuleName*")
	
	[string]$ModuleFound = ($LoadedModules -like "*$ModuleName*")
	If ($ModuleFound -ne $ModuleName)
	{
		$module = Import-Module -Name $ModuleName -PassThru -EA 0 4>$Null
		If ($module -and $?)
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

Function Set-DocumentProperty
{
	<#
	.SYNOPSIS
	Function to set the Title Page document properties in MS Word
	.DESCRIPTION
	Long description
	.PARAMETER Document
	Current Document Object
	.PARAMETER DocProperty
	Parameter description
	.PARAMETER Value
	Parameter description
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value 'MyTitle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value 'MyCompany'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value 'Jim Moyle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value 'MySubjectTitle'
	.NOTES
	Function Created by Jim Moyle June 2017
	Twitter : @JimMoyle
	#>
	param (
		[object]$Document,
		[String]$DocProperty,
		[string]$Value
	)
	try
	{
		$binding = "System.Reflection.BindingFlags" -as [type]
		$builtInProperties = $Document.BuiltInDocumentProperties
		$property = [System.__ComObject].invokemember("item", $binding::GetProperty, $null, $BuiltinProperties, $DocProperty)
		[System.__ComObject].invokemember("value", $binding::SetProperty, $null, $property, $Value)
	}
 catch
	{
		Write-Warning "Failed to set $DocProperty to $Value"
	}
}

Function FindWordDocumentEnd
{
	#Return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory, $wdMove) | Out-Null
}

Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If ( $object )
	{
		If ((Get-Member -Name $topLevel -InputObject $object))
		{
			If ((Get-Member -Name $secondLevel -InputObject $object.$topLevel))
			{
				Return $True
			}
		}
	}
	Return $False
}

Function validObject( [object] $object, [string] $topLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If ( $object )
	{
		If ((Get-Member -Name $topLevel -InputObject $object))
		{
			Return $True
		}
	}
	Return $False
}

Function SetupWord
{
	Write-Verbose "$(Get-Date -Format G): Setting up Word"
    
	If (!$AddDateTime)
	{
		[string]$Script:WordFileName = "$($Script:pwdpath)\$($OutputFileName).docx"
		If ($PDF)
		{
			[string]$Script:PDFFileName = "$($Script:pwdpath)\$($OutputFileName).pdf"
		}
	}
	ElseIf ($AddDateTime)
	{
		[string]$Script:WordFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
		If ($PDF)
		{
			[string]$Script:PDFFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
		}
	}

	# Setup word for output
	Write-Verbose "$(Get-Date -Format G): Create Word comObject."
	$Script:Word = New-Object -ComObject "Word.Application" -EA 0 4>$Null

	#Do not indent the following write-error lines. Doing so will mess up the console formatting of the error message.
	If (!$? -or $Null -eq $Script:Word)
	{
		Write-Warning "The Word object could not be created. You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	The Word object could not be created. You may need to repair your Word installation.
		`n`n
	Script cannot Continue.
		`n`n"
		AbortScript
	}

	Write-Verbose "$(Get-Date -Format G): Determine Word language value"
	If ( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If (!($Script:WordLanguageValue -gt -1))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	Unable to determine the Word language value. You may need to repair your Word installation.
		`n`n
	Script cannot Continue.
		`n`n
		"
		AbortScript
	}
	Write-Verbose "$(Get-Date -Format G): Word language value is $($Script:WordLanguageValue)"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If ($Script:WordVersion -eq $wdWord2016)
	{
		$Script:WordProduct = "Word 2016"
	}
	ElseIf ($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf ($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf ($Script:WordVersion -eq $wdWord2007)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	Microsoft Word 2007 is no longer supported.`n`n`t`tScript will end.
		`n`n
		"
		AbortScript
	}
	ElseIf ($Script:WordVersion -eq 0)
	{
		Write-Error "
		`n`n
	The Word Version is 0. You should run a full online repair of your Office installation.
		`n`n
	Script cannot Continue.
		`n`n
		"
		AbortScript
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	You are running an untested or unsupported version of Microsoft Word.
		`n`n
	Script will end.
		`n`n
	Please send info on your version of Word to webster@carlwebster.com
		`n`n
		"
		AbortScript
	}

	#only validate CompanyName   If the field is blank
	If ([String]::IsNullOrEmpty($CompanyName))
	{
		Write-Verbose "$(Get-Date -Format G): Company name is blank. Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If ([String]::IsNullOrEmpty($TmpName))
		{
			Write-Host "
		Company Name is blank so Cover Page will not show a Company Name.
		Check HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value.
		You may want to use the -CompanyName parameter   If you need a Company Name on the cover page.
			" -ForegroundColor White
			$Script:CoName = $TmpName
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date -Format G): Updated company name to $($Script:CoName)"
		}
	}
	Else
	{
		$Script:CoName = $CompanyName
	}

	If ($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date -Format G): Check Default Cover Page for $($WordCultureCode)"
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'
			{
				If ($CoverPage -eq "Sideline")
				{
					$CoverPage = "Línia lateral"
					$CPChanged = $True
				}
			}

			'da-'
			{
				If ($CoverPage -eq "Sideline")
				{
					$CoverPage = "Sidelinje"
					$CPChanged = $True
				}
			}

			'de-'
			{
				If ($CoverPage -eq "Sideline")
				{
					$CoverPage = "Randlinie"
					$CPChanged = $True
				}
			}

			'es-'
			{
				If ($CoverPage -eq "Sideline")
				{
					$CoverPage = "Línea lateral"
					$CPChanged = $True
				}
			}

			'fi-'
			{
				If ($CoverPage -eq "Sideline")
				{
					$CoverPage = "Sivussa"
					$CPChanged = $True
				}
			}

			'fr-'
			{
				If ($CoverPage -eq "Sideline")
				{
					If ($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
					{
						$CoverPage = "Lignes latérales"
						$CPChanged = $True
					}
					Else
					{
						$CoverPage = "Ligne latérale"
						$CPChanged = $True
					}
				}
			}

			'nb-'
			{
				If ($CoverPage -eq "Sideline")
				{
					$CoverPage = "Sidelinje"
					$CPChanged = $True
				}
			}

			'nl-'
			{
				If ($CoverPage -eq "Sideline")
				{
					$CoverPage = "Terzijde"
					$CPChanged = $True
				}
			}

			'pt-'
			{
				If ($CoverPage -eq "Sideline")
				{
					$CoverPage = "Linha Lateral"
					$CPChanged = $True
				}
			}

			'sv-'
			{
				If ($CoverPage -eq "Sideline")
				{
					$CoverPage = "Sidlinje"
					$CPChanged = $True
				}
			}

			'zh-'
			{
				If ($CoverPage -eq "Sideline")
				{
					$CoverPage = "边线型"
					$CPChanged = $True
				}
			}
		}

		If ($CPChanged)
		{
			Write-Verbose "$(Get-Date -Format G): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date -Format G): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If (!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date -Format G): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date -Format G): Culture code $($Script:WordCultureCode)"
		Write-Error "
		`n`n
	For $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.
		`n`n
	Script cannot Continue.
		`n`n
		"
		AbortScript
	}

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date -Format G): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013/2016
	$BuildingBlocksCollection = $Script:Word.Templates | Where-Object { $_.name -eq "Built-In Building Blocks.dotx" }

	Write-Verbose "$(Get-Date -Format G): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
		ForEach-Object {
			If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage)
			{
				$BuildingBlocks = $_
			}
		}        

	If ($Null -ne $BuildingBlocks)
	{
		$BuildingBlocksExist = $True

		Try
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If ($Null -ne $part)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If (!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date -Format G): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Host "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist." -ForegroundColor White
		Write-Host "This report will not have a Cover Page." -ForegroundColor White
	}

	Write-Verbose "$(Get-Date -Format G): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If ($Null -eq $Script:Doc)
	{
		Write-Verbose "$(Get-Date -Format G): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	An empty Word document could not be created. You may need to repair your Word installation.
		`n`n
	Script cannot Continue.
		`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If ($Null -eq $Script:Selection)
	{
		Write-Verbose "$(Get-Date -Format G): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	An unknown error happened selecting the entire Word document for default formatting options.
		`n`n
	Script cannot Continue.
		`n`n"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 =.50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date -Format G): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If ($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date -Format G): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range, $True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date -Format G): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If ($Null -eq $toc)
		{
			Write-Verbose "$(Get-Date -Format G): "
			Write-Host "Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved." -ForegroundColor White
			Write-Host "This report will not have a Table of Contents." -ForegroundColor White
		}
		Else
		{
			$toc.insert($Script:Selection.Range, $True) | Out-Null
		}
	}
	Else
	{
		Write-Host "Table of Contents are not installed." -ForegroundColor White
		Write-Host "Table of Contents are not installed so this report will not have a Table of Contents." -ForegroundColor White
	}

	#set the footer
	Write-Verbose "$(Get-Date -Format G): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date -Format G): Get the footer and format font"
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach ($footer in $footers)
	{
		If ($footer.exists)
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date -Format G): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date -Format G): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#updated 8-Jun-2017 with additional cover page fields
	#Update document properties
	If ($MSWORD -or $PDF)
	{
		If ($Script:CoverPagesExist)
		{
			Write-Verbose "$(Get-Date -Format G): Set Cover Page Properties"
			#8-Jun-2017 put these 4 items in alpha order
			Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value $UserName
			Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value $Script:CoName
			Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value $SubjectTitle
			Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value $Script:title

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where-Object { $_.NamespaceURI -match "coverPageProps$" }

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where-Object { $_.basename -eq "Abstract" }
			#set the text
			If ([String]::IsNullOrEmpty($Script:CoName))
			{
				[string]$abstract = $AbstractTitle
			}
			Else
			{
				[string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
			}
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object { $_.basename -eq "CompanyAddress" }
			#set the text
			[string]$abstract = $CompanyAddress
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object { $_.basename -eq "CompanyEmail" }
			#set the text
			[string]$abstract = $CompanyEmail
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object { $_.basename -eq "CompanyFax" }
			#set the text
			[string]$abstract = $CompanyFax
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object { $_.basename -eq "CompanyPhone" }
			#set the text
			[string]$abstract = $CompanyPhone
			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where-Object { $_.basename -eq "PublishDate" }
			#set the text
			[string]$abstract = (Get-Date -Format d).ToString()
			$ab.Text = $abstract

			Write-Verbose "$(Get-Date -Format G): Update the Table of Contents"
			#update the Table of Contents
			$Script:Doc.TablesOfContents.item(1).Update()
			$cp = $Null
			$ab = $Null
			$abstract = $Null
		}
	}
}
#endregion

#region registry functions
#http://stackoverflow.com/questions/5648931/test-If-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null   If it is missing
Function Get-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If ($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}

# Gets the specified registry value or $Null   If it is missing
Function Get-RegistryValue2
{
	[CmdletBinding()]
	Param([string]$path, [string]$name, [string]$ComputerName)
	If ($ComputerName -eq $env:computername)
	{
		$key = Get-Item -LiteralPath $path -EA 0
		If ($key)
		{
			Return $key.GetValue($name, $Null)
		}
		Else
		{
			Return $Null
		}
	}
	Else
	{
		#path needed here is different for remote registry access
		$path = $path.SubString(6)
		$path2 = $path.Replace('\', '\\')
		$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
		$RegKey = $Reg.OpenSubKey($path2)
		If ($RegKey)
		{
			$Results = $RegKey.GetValue($name)

			If ($Null -ne $Results)
			{
				Return $Results
			}
			Else
			{
				Return $Null
			}
		}
		Else
		{
			Return $Null
		}
	}
}
#endregion

#region word, text and html line output functions
Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexch on Twitter
#https://essential.exchange/blog
#for creating the formatted text report
{
 #created March 2011
 #updated March 2014
 # updated March 2019 to use StringBuilder (about 100 times more efficient than simple strings)
	Param
	(
		[Int]    $tabs = 0, 
		[String] $name = '', 
		[String] $value = '', 
		[String] $newline = [System.Environment]::NewLine, 
		[Switch] $nonewline
	)

	while ( $tabs -gt 0 )
	{
		$Null = $script:Output.Append( "`t" )
		$tabs--
	}

	If ( $nonewline )
	{
		$Null = $script:Output.Append( $name + $value )
	}
	Else
	{
		$Null = $script:Output.AppendLine( $name + $value )
	}
}

Function WriteWordLine
#Function created by Ryan Revord
{
 #@rsrevord on Twitter
 #Function created to make output to Word easy in this script
 #updated 27-Mar-2014 to include font name, font size, italics and bold options
	Param([int]$style = 0, 
		[int]$tabs = 0, 
		[string]$name = '', 
		[string]$value = '', 
		[string]$fontName = $Null,
		[int]$fontSize = 0,
		[bool]$italics = $False,
		[bool]$boldface = $False,
		[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 { $Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break }
		1 { $Script:Selection.Style = $Script:MyHash.Word_Heading1; Break }
		2 { $Script:Selection.Style = $Script:MyHash.Word_Heading2; Break }
		3 { $Script:Selection.Style = $Script:MyHash.Word_Heading3; Break }
		4 { $Script:Selection.Style = $Script:MyHash.Word_Heading4; Break }
		5 { $Script:Selection.Style = $Script:MyHash.Word_Heading5; Break }
		Default { $Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break }
	}
	
	#build # of tabs
	While ($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If (![String]::IsNullOrEmpty($fontName))
	{
		$Script:Selection.Font.name = $fontName
	} 

	If ($fontSize -ne 0)
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If ($italics -eq $True)
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If ($boldface -eq $True)
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If ($nonewline)
	{
		# Do nothing.
	}
	Else
	{
		$Script:Selection.TypeParagraph()
	}

	#put these two back
	If ($italics -eq $True)
	{
		$Script:Selection.Font.Italic = $False
	} 
 
	If ($boldface -eq $True)
	{
		$Script:Selection.Font.Bold = $False
	} 
}

#***********************************************************************************************************
# WriteHTMLLine
#***********************************************************************************************************

<#
.Synopsis
	Writes a line of output for HTML output
.DESCRIPTION
	This function formats an HTML line
.USAGE
	WriteHTMLLine <Style> <Tabs> <Name> <Value> <Font Name> <Font Size> <Options>

	0 for Font Size denotes using the default font size of 2 or 10 point

.EXAMPLE
	WriteHTMLLine 0 0 " "

	Writes a blank line with no style or tab stops, obviously none needed.

.EXAMPLE
	WriteHTMLLine 0 1 "This is a regular line of text indented 1 tab stops"

	Writes a line with 1 tab stop.

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in italics" "" $Null 0 $htmlitalics

	Writes a line omitting font and font size and setting the italics attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold" "" $Null 0 $htmlBold

	Writes a line omitting font and font size and setting the bold attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold italics" "" $Null 0 ($htmlBold -bor $htmlitalics)

	Writes a line omitting font and font size and setting both italics and bold options

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in 10 point" "" $Null 2  # 10 point font

	Writes a line using 10 point font

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 

	Writes a line using Courier New Font and 0 font point size (default = 2   If set to 0)

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of RED text indented 0 tab stops with the computer name as data in 10 point Courier New bold italics: " $env:computername "Courier New" 2 ($htmlBold -bor $htmlred -bor $htmlitalics)

	Writes a line using Courier New Font with first and second string values to be used, also uses 10 point font with bold, italics and red color options set.

.NOTES

	Font Size - Unlike word, there is a limited set of font sizes that can be used in HTML. They are:
		0 - default which actually gives it a 2 or 10 point.
		1 - 7.5 point font size
		2 - 10 point
		3 - 13.5 point
		4 - 15 point
		5 - 18 point
		6 - 24 point
		7 - 36 point
	Any number larger than 7 defaults to 7

	Style - Refers to the headers that are used with output and resemble the headers in word, 
	HTML supports headers h1-h6 and h1-h4 are more commonly used. Unlike word, H1 will not 
	give you a blue colored font, you will have to set that yourself.

	Colors and Bold/Italics Flags are:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack       
#>

# to suppress $crlf in HTML documents, replace this with '' (empty string)
# but this was added to make the HTML readable
$crlf = [System.Environment]::NewLine

Function WriteHTMLLine
#Function created by Ken Avram
#Function created to make output to HTML easy in this script
{
 #headings fixed 12-Oct-2016 by Webster
 #errors with $HTMLStyle fixed 7-Dec-2017 by Webster
 # re-implemented/re-based by Michael B. Smith
	Param
	(
		[Int]    $style = 0, 
		[Int]    $tabs = 0, 
		[String] $name = '', 
		[String] $value = '', 
		[String] $fontName = $null,
		[Int]    $fontSize = 1,
		[Int]    $options = $htmlblack
	)

	#FIXME - long story short, this function was wrong and had been wrong for a long time. 
	## The function generated invalid HTML, and ignored fontname and fontsize parameters. I fixed
	## those items, but that made the report unreadable, because all of the formatting had been based
	## on this function not working properly.

	## here is a typical H1 previously generated:
	## <h1>///&nbsp;&nbsp;Forest Information&nbsp;&nbsp;\\\<font face='Calibri' color='#000000' size='1'></h1></font>

	## fixing the function generated this (unreadably small):
	## <h1><font face='Calibri' color='#000000' size='1'>///&nbsp;&nbsp;Forest Information&nbsp;&nbsp;\\\</font></h1>

	## So I took all the fixes out. This routine now generates valid HTML, but the fontName, fontSize,
	## and options parameters are ignored; so the routine generates equivalent output as before. I took
	## the fixes out instead of fixing all the call sites, because there are 225 call sites!   If you are
	## willing to update all the call sites, you can easily re-instate the fixes. They have only been
	## commented out with '##' below.

	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 1024 )

	If ( [String]::IsNullOrEmpty( $name ) )
	{
		## $HTMLBody = '<p></p>'
		$null = $sb.Append( '<p></p>' )
	}
	Else
	{
		[Bool] $ital = $options -band $htmlitalics
		[Bool] $bold = $options -band $htmlBold
		If ( $ital ) { $null = $sb.Append( '<i>' ) }
		If ( $bold ) { $null = $sb.Append( '<b>' ) } 

		switch ( $style )
		{
			1 { $HTMLOpen = '<h1>'; $HTMLClose = '</h1>'; Break }
			2 { $HTMLOpen = '<h2>'; $HTMLClose = '</h2>'; Break }
			3 { $HTMLOpen = '<h3>'; $HTMLClose = '</h3>'; Break }
			4 { $HTMLOpen = '<h4>'; $HTMLClose = '</h4>'; Break }
			Default { $HTMLOpen = ''; $HTMLClose = ''; Break }
		}

		$null = $sb.Append( $HTMLOpen )

		$null = $sb.Append( ( '&nbsp;&nbsp;&nbsp;&nbsp;' * $tabs ) + $name + $value )

		If ( $HTMLClose -eq '' ) { $null = $sb.Append( '<br>' ) }
		Else { $null = $sb.Append( $HTMLClose ) }

		If ( $ital ) { $null = $sb.Append( '</i>' ) }
		If ( $bold ) { $null = $sb.Append( '</b>' ) } 

		If ( $HTMLClose -eq '' ) { $null = $sb.Append( '<br />' ) }
	}
	$null = $sb.AppendLine( '' )

	Out-File -FilePath $Script:HtmlFileName -Append -InputObject $sb.ToString() 4>$Null
}
#endregion

#region HTML table functions
#***********************************************************************************************************
# AddHTMLTable - Called from FormatHTMLTable function
# Created by Ken Avram
# modified by Jake Rutski
# re-implemented by Michael B. Smith. Also made the documentation match reality.
#***********************************************************************************************************
Function AddHTMLTable
{
	Param
	(
		[String]   $fontName = 'Calibri',
		[Int]      $fontSize = 2,
		[Int]      $colCount = 0,
		[Int]      $rowCount = 0,
		[Object[]] $rowInfo = $null,
		[Object[]] $fixedInfo = $null
	)

	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 8192 )

	If ( $rowInfo -and $rowInfo.Length -lt $rowCount )
	{
		$rowCount = $rowInfo.Length
	}

	for ( $rowCountIndex = 0; $rowCountIndex -lt $rowCount; $rowCountIndex++ )
	{
		$null = $sb.AppendLine( '<tr>' )
		## $htmlbody += '<tr>'
		## $htmlbody += $crlf make the HTML readable

		## each row of rowInfo is an array
		## each row consists of tuples: an item of text followed by an item of formatting data

		## reset
		$row = $rowInfo[ $rowCountIndex ]

		$subRow = $row
		If ( $subRow -is [Array] -and $subRow[ 0 ] -is [Array] )
		{
			$subRow = $subRow[ 0 ]
		}

		$subRowLength = $subRow.Length
		for ( $columnIndex = 0; $columnIndex -lt $colCount; $columnIndex += 2 )
		{
			$item = If ( $columnIndex -lt $subRowLength ) { $subRow[ $columnIndex ] }   Else { 0 }

			$text = If ( $item ) { $item.ToString() }   Else { '' }
			$format = If ( ( $columnIndex + 1 ) -lt $subRowLength ) { $subRow[ $columnIndex + 1 ] }   Else { 0 }
			## item, text, and format ALWAYS have values, even   If empty values
			$color = $Script:htmlColor[ $format -band 0xffffc ]
			[Bool] $bold = $format -band $htmlBold
			[Bool] $ital = $format -band $htmlitalics

			If ( $null -eq $fixedInfo -or $fixedInfo.Length -eq 0 )
			{
				$null = $sb.Append( "<td style=""background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>" )
			}
			Else
			{
				$null = $sb.Append( "<td style=""width:$( $fixedInfo[ $columnIndex / 2 ] ); background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>" )
			}

			If ( $bold ) { $null = $sb.Append( '<b>' ) }
			If ( $ital ) { $null = $sb.Append( '<i>' ) }

			If ( $text -eq ' ' -or $text.length -eq 0)
			{
				$null = $sb.Append( '&nbsp;&nbsp;&nbsp;' )
			}
			Else
			{
				for ($inx = 0; $inx -lt $text.length; $inx++ )
				{
					If ( $text[ $inx ] -eq ' ' )
					{
						$null = $sb.Append( '&nbsp;' )
					}
					Else
					{
						break
					}
				}
				$null = $sb.Append( $text )
			}

			If ( $bold ) { $null = $sb.Append( '</b>' ) }
			If ( $ital ) { $null = $sb.Append( '</i>' ) }

			$null = $sb.AppendLine( '</font></td>' )
		}

		$null = $sb.AppendLine( '</tr>' )
	}

	Out-File -FilePath $Script:HtmlFileName -Append -InputObject $sb.ToString() 4>$Null 
}

#***********************************************************************************************************
# FormatHTMLTable 
# Created by Ken Avram
# modified by Jake Rutski
# reworked by Michael B. Smith
#***********************************************************************************************************

<#
.Synopsis
	Format table for a HTML output document.
.DESCRIPTION
	This function formats a table for HTML from multiple arrays of strings.
.PARAMETER noBorder
	If set to $true, a table will be generated without a border (border = '0'). Otherwise the table will be generated
	with a border (border = '1').
.PARAMETER noHeadCols
	This parameter should be used when generating tables which do not have a separate array containing column headers
	(columnArray is not specified). Set this parameter equal to the number of columns in the table.
.PARAMETER rowArray
	This parameter contains the row data array for the table.
.PARAMETER columnArray
	This parameter contains column header data for the table.
.PARAMETER fixedWidth
	This parameter contains widths for columns in pixel format ("100px") to override auto column widths
	The variable should contain a width for each column you wish to override the auto-size setting
	For example: $fixedWidth = @("100px","110px","120px","130px","140px")
.PARAMETER tableHeader
	A string containing the header for the table (printed at the top of the table, left justified). The
	default is a blank string.
.PARAMETER tableWidth
	The width of the table in pixels, or 'auto'. The default is 'auto'.
.PARAMETER fontName
	The name of the font to use in the table. The default is 'Calibri'.
.PARAMETER fontSize
	The size of the font to use in the table. The default is 2. Note that this is the HTML size, not the pixel size.

.USAGE
	FormatHTMLTable <Table Header> <Table Width> <Font Name> <Font Size>

.EXAMPLE
	FormatHTMLTable "Table Heading" "auto" "Calibri" 3

	This example formats a table and writes it out into an html file. All of the parameters are optional
	defaults are used   If not supplied.

	for <Table format>, the default is auto which will autofit the text into the columns and adjust to the longest text in that column. You can also use percentage i.e. 25%
	which will take only 25% of the line and will auto word wrap the text to the next line in the column. Also, instead of using a percentage, you can use pixels i.e. 400px.

	FormatHTMLTable "Table Heading" "auto" -rowArray $rowData -columnArray $columnData

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, column header data from $columnData and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -noHeadCols 3

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, no header, and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -fixedWidth $fixedColumns

	This example creates an HTML table with a heading of 'Table Heading, no header, row data from $rowData, and fixed columns defined by $fixedColumns

.NOTES
	In order to use the formatted table it first has to be loaded with data. Examples below will show how to load the table:

	First, initialize the table array

	$rowdata = @()

	Then Load the array.   If you are using column headers then load those into the column headers array, otherwise the first line of the table goes into the column headers array
	and the second and subsequent lines go into the $rowdata table as shown below:

	$columnHeaders = @('Display Name',$htmlsb,'Status',$htmlsb,'Startup Type',$htmlsb)

	The first column is the actual name to display, the second are the attributes of the column i.e. color anded with bold or italics. For the anding, parens are required or it will
	not format correctly.

	This is following by adding rowdata as shown below. As more columns are added the columns will auto adjust to fit the size of the page.

	$rowdata = @()
	$columnHeaders = @("User Name",$htmlsb,$UserName,$htmlwhite)
	$rowdata += @(,("Save as PDF',$htmlsb,$PDF.ToString(),$htmlwhite))
	$rowdata += @(,("Save as TEXT',$htmlsb,$TEXT.ToString(),$htmlwhite))
	$rowdata += @(,("Save as WORD',$htmlsb,$MSWORD.ToString(),$htmlwhite))
	$rowdata += @(,("Save as HTML',$htmlsb,$HTML.ToString(),$htmlwhite))
	$rowdata += @(,("Add DateTime',$htmlsb,$AddDateTime.ToString(),$htmlwhite))
	$rowdata += @(,("Hardware Inventory',$htmlsb,$Hardware.ToString(),$htmlwhite))
	$rowdata += @(,("Computer Name',$htmlsb,$ComputerName,$htmlwhite))
	$rowdata += @(,("FileName',$htmlsb,$Script:FileName,$htmlwhite))
	$rowdata += @(,("OS Detected',$htmlsb,$Script:RunningOS,$htmlwhite))
	$rowdata += @(,("PSUICulture',$htmlsb,$PSCulture,$htmlwhite))
	$rowdata += @(,("PoSH version',$htmlsb,$Host.Version.ToString(),$htmlwhite))
	FormatHTMLTable "Example of Horizontal AutoFitContents HTML Table" -rowArray $rowdata

	The 'rowArray' paramater is mandatory to build the table, but it is not set as such in the function -   If nothing is passed, the table will be empty.

	Colors and Bold/Italics Flags are shown below:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack     

#>

Function FormatHTMLTable
{
	Param
	(
		[String]   $tableheader = '',
		[String]   $tablewidth = 'auto',
		[String]   $fontName = 'Calibri',
		[Int]      $fontSize = 2,
		[Switch]   $noBorder = $false,
		[Int]      $noHeadCols = 1,
		[Object[]] $rowArray = $null,
		[Object[]] $fixedWidth = $null,
		[Object[]] $columnArray = $null
	)

	## FIXME - the help text for this function is wacky wrong - MBS
	## FIXME - Use StringBuilder - MBS - this only builds the table header - benefit relatively small

	$HTMLBody = "<b><font face='" + $fontname + "' size='" + ($fontsize + 1) + "'>" + $tableheader + "</font></b>" + $crlf

	If ( $null -eq $columnArray -or $columnArray.Length -eq 0)
	{
		$NumCols = $noHeadCols + 1
	}  # means we have no column headers, just a table
	Else
	{
		$NumCols = $columnArray.Length
	}  # need to add one for the color attrib

	If ( $null -ne $rowArray )
	{
		$NumRows = $rowArray.length + 1
	}
	Else
	{
		$NumRows = 1
	}

	If ( $noBorder )
	{
		$HTMLBody += "<table border='0' width='" + $tablewidth + "'>"
	}
	Else
	{
		$HTMLBody += "<table border='1' width='" + $tablewidth + "'>"
	}
	$HTMLBody += $crlf

	If ( $columnArray -and $columnArray.Length -gt 0 )
	{
		$HTMLBody += '<tr>' + $crlf

		for ( $columnIndex = 0; $columnIndex -lt $NumCols; $columnindex += 2 )
		{
			$val = $columnArray[ $columnIndex + 1 ]
			$tmp = $Script:htmlColor[ $val -band 0xffffc ]
			[Bool] $bold = $val -band $htmlBold
			[Bool] $ital = $val -band $htmlitalics

			If ( $null -eq $fixedWidth -or $fixedWidth.Length -eq 0 )
			{
				$HTMLBody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			Else
			{
				$HTMLBody += "<td style=""width:$($fixedWidth[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			If ( $bold ) { $HTMLBody += '<b>' }
			If ( $ital ) { $HTMLBody += '<i>' }

			$array = $columnArray[ $columnIndex ]
			If ( $array )
			{
				If ( $array -eq ' ' -or $array.Length -eq 0 )
				{
					$HTMLBody += '&nbsp;&nbsp;&nbsp;'
				}
				Else
				{
					for ( $i = 0; $i -lt $array.Length; $i += 2 )
					{
						If ( $array[ $i ] -eq ' ' )
						{
							$HTMLBody += '&nbsp;'
						}
						Else
						{
							break
						}
					}
					$HTMLBody += $array
				}
			}
			Else
			{
				$HTMLBody += '&nbsp;&nbsp;&nbsp;'
			}
			
			If ( $bold ) { $HTMLBody += '</b>' }
			If ( $ital ) { $HTMLBody += '</i>' }
		}

		$HTMLBody += '</font></td>'
		$HTMLBody += $crlf
	}

	$HTMLBody += '</tr>' + $crlf

	Out-File -FilePath $Script:HtmlFileName -Append -InputObject $HTMLBody 4>$Null 
	$HTMLBody = ''

	If ( $rowArray )
	{

		AddHTMLTable -fontName $fontName -fontSize $fontSize `
			-colCount $numCols -rowCount $NumRows `
			-rowInfo $rowArray -fixedInfo $fixedWidth
		$rowArray = $null
		$HTMLBody = '</table>'
	}
	Else
	{
		$HTMLBody += '</table>'
	}

	Out-File -FilePath $Script:HtmlFileName -Append -InputObject $HTMLBody 4>$Null 
}
#endregion

#region other HTML functions
Function SetupHTML
{
	Write-Verbose "$(Get-Date -Format G): Setting up HTML"
	If (!$AddDateTime)
	{
		[string]$Script:HtmlFileName = "$($Script:pwdpath)\$($OutputFileName).html"
	}
	ElseIf ($AddDateTime)
	{
		[string]$Script:HtmlFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}

	$htmlhead = "<html><head><meta http-equiv='Content-Language' content='da'><title>" + $Script:Title + "</title></head><body>"
	Out-File -FilePath $Script:HtmlFileName -Force -InputObject $HTMLHead 4>$Null
}
#endregion

#region Iain's Word table functions

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is Returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress.   If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress.   If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>

Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True, ParameterSetName = 'Hashtable', Position = 0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True, ParameterSetName = 'CustomObject', Position = 0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		#   If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName = $True)] [AllowNull()] [string[]] $Columns = $Null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName = $True)] [AllowNull()] [string[]] $Headers = $Null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName = $True)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		[Switch] $NoInternalGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName = $True)] [int] $Format = 0
	)

	Begin
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check   If -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If (($Null -eq $Columns) -and ($Null -eq $Headers))
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $Null;
		}
		ElseIf (($Null -ne $Columns) -and ($Null -ne $Headers))
		{
			## Check   If number of specified -Columns matches number of specified -Headers
			If ($Columns.Length -ne $Headers.Length)
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end   ElseIf
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
		[System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName)
		{
			'CustomObject'
			{
				If ($Null -eq $Columns)
				{
					## Build the available columns from all availble PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach ($Property in ($CustomObject | Get-Member -MemberType NoteProperty))
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If (-not $List)
				{
					Write-Debug ("$(Get-Date -Format G): `t`tBuilding table headers");
					If ($Null -ne $Headers)
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date -Format G): `t`tBuilding table rows");
				ForEach ($Object in $CustomObject)
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach ($Column in $Columns)
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach
				Write-Debug ("$(Get-Date -Format G): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default
			{
				## Hashtable
				If ($Null -eq $Columns)
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If (-not $List)
				{
					Write-Debug ("$(Get-Date -Format G): `t`tBuilding table headers");
					If ($Null -ne $Headers)
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date -Format G): `t`tBuilding table rows");
				ForEach ($Hash in $Hashtable)
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach ($Column in $Columns)
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach

				Write-Debug ("$(Get-Date -Format G): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date -Format G): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If ($Format -ge 0)
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $True);
			$ConvertToTableArguments.Add("ApplyShading", $True);
			$ConvertToTableArguments.Add("ApplyFont", $True);
			$ConvertToTableArguments.Add("ApplyColor", $True);
			If (!$List)
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $True); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $True);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $True);
			$ConvertToTableArguments.Add("ApplyLastColumn", $True);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date -Format G): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable", # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$Null, # Binder
			$WordRange, # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)), ## Named argument values
			$Null, # Modifiers
			$Null, # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If ($Format -lt 0)
		{
			Write-Debug ("$(Get-Date -Format G): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If ($AutoFit -ne -1)
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If (!$List)
		{
			#the next line causes the heading row to flow across page breaks
			$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
		}

		If (!$NoGridLines)
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}
		If ($NoGridLines)
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleNone;
		}
		If ($NoInternalGridLines)
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@   { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells Returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells Returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) Returns a single Word COM cells object.
#>

Function SetWordCellFormat
{
	[CmdletBinding(DefaultParameterSetName = 'Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory = $True, ValueFromPipeline = $True, ParameterSetName = 'Collection', Position = 0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory = $True, ParameterSetName = 'Cell', Position = 0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory = $True, ParameterSetName = 'Hashtable', Position = 0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory = $True, ParameterSetName = 'Hashtable', Position = 1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $Null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $Null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] [int]$BackgroundColor = $Null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process
	{
		Switch ($PSCmdlet.ParameterSetName)
		{
			'Collection'
			{
				ForEach ($Cell in $Collection)
				{
					If ($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If ($Bold) { $Cell.Range.Font.Bold = $True; }
					If ($Italic) { $Cell.Range.Font.Italic = $True; }
					If ($Underline) { $Cell.Range.Font.Underline = 1; }
					If ($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If ($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If ($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If ($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end ForEach
			} # end Collection
			'Cell'
			{
				If ($Bold) { $Cell.Range.Font.Bold = $True; }
				If ($Italic) { $Cell.Range.Font.Italic = $True; }
				If ($Underline) { $Cell.Range.Font.Underline = 1; }
				If ($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
				If ($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
				If ($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If ($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If ($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable'
			{
				ForEach ($Coordinate in $Coordinates)
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If ($Bold) { $Cell.Range.Font.Bold = $True; }
					If ($Italic) { $Cell.Range.Font.Italic = $True; }
					If ($Underline) { $Cell.Range.Font.Underline = 1; }
					If ($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If ($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If ($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If ($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If ($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (If one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this function is called by the AddWordTable function   If an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>

Function SetWordTableAlternateRowColor
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory = $True, ValueFromPipeline = $True, Position = 0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory = $True, Position = 1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName = $True, Position = 2)] [ValidateSet('First', 'Second')] [string] $Seed = 'First'
	)

	Process
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If ($Seed.ToLower() -eq 'second')
		{ 
			$StartRowIndex = 2; 
		}
		Else
		{ 
			$StartRowIndex = 1; 
		}

		For ($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2)
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
#endregion

#region general script functions
Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date -Format G): Save and Close document and Shutdown Word"
	If ($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If ($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date -Format G): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:WordFileName, [ref]$SaveFormat)
		If ($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:PDFFileName, [ref]$saveFormat)
		}
	}
	ElseIf ($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
	{
		If ($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date -Format G): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:WordFileName, [ref]$wdFormatDocumentDefault)
		If ($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:PDFFileName, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date -Format G): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	Write-Verbose "$(Get-Date -Format G): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If (Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global 4>$Null
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	
	#is the winword Process still running? kill it

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId

	#Find out   If winword running in our session
	$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object { $_.SessionId -eq $SessionID }) | Select-Object -Property Id 
	If ( $wordprocess -and $wordprocess.Id -gt 0)
	{
		Write-Verbose "$(Get-Date -Format G): WinWord Process is still running. Attempting to stop WinWord Process # $($wordprocess.Id)"
		Stop-Process $wordprocess.Id -EA 0
	}
}

Function SetupText
{
	Write-Verbose "$(Get-Date -Format G): Setting up Text"
	[System.Text.StringBuilder] $Script:Output = New-Object System.Text.StringBuilder( 16384 )

	If (!$AddDateTime)
	{
		[string]$Script:TextFileName = "$($Script:pwdpath)\$($OutputFileName).txt"
	}
	ElseIf ($AddDateTime)
	{
		[string]$Script:TextFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}
}

Function SaveandCloseTextDocument
{
	Write-Verbose "$(Get-Date -Format G): Saving Text file"
	Line 0 ""
	Line 0 "Report Complete"
	Write-Output $script:Output.ToString() | Out-File $Script:TextFileName 4>$Null
	[System.Text.StringBuilder] $Script:Output = New-Object System.Text.StringBuilder( 16384 )
}

Function SaveandCloseHTMLDocument
{
	Write-Verbose "$(Get-Date -Format G): Saving HTML file"
	WriteHTMLLine 0 0 ""
	WriteHTMLLine 0 0 "Report Complete"
	Out-File -FilePath $Script:HTMLFileName -Append -InputObject "<p></p></body></html>" 4>$Null
}

Function SetFilenames
{
	Param([string]$OutputFileName)
	
	If ($MSWord -or $PDF)
	{
		CheckWordPreReq
		
		SetupWord
	}
	If ($Text)
	{
		SetupText
	}
	If ($HTML)
	{
		SetupHTML
	}
	ShowScriptOptions
}

Function OutputReportFooter
{
	<#
	Report Footer
		Report information:
			Created with: <Script Name> - Release Date: <Script Release Date>
			Script version: <Script Version>
			Started on <Date Time in Local Format>
			Elapsed time: nn days, nn hours, nn minutes, nn.nn seconds
			Ran from domain <Domain Name> by user <Username>
			Ran from the folder <Folder Name>

	Script Name and Script Release date are script-specific variables.
	Script version is a script variable.
	Start Date Time in Local Format is a script variable.
	Domain Name is $env:USERDNSDOMAIN.
	Username is $env:USERNAME.
	Folder Name is a script variable.
	#>

	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days,   {1} hours,   {2} minutes,   {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)

	If ($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Report Footer"
		WriteWordLine 2 0 "Report Information:"
		WriteWordLine 0 1 "Created with: $Script:ScriptName - Release Date: $Script:ReleaseDate"
		WriteWordLine 0 1 "Script version: $Script:MyVersion"
		WriteWordLine 0 1 "Started on $Script:StartTime"
		WriteWordLine 0 1 "Elapsed time: $Str"
		WriteWordLine 0 1 "Ran from domain $env:USERDNSDOMAIN by user $env:USERNAME"
		WriteWordLine 0 1 "Ran from the folder $Script:pwdpath"
	}
	If ($Text)
	{
		Line 0 "///  Report Footer  \\\"
		Line 1 "Report Information:"
		Line 2 "Created with: $Script:ScriptName - Release Date: $Script:ReleaseDate"
		Line 2 "Script version: $Script:MyVersion"
		Line 2 "Started on $Script:StartTime"
		Line 2 "Elapsed time: $Str"
		Line 2 "Ran from domain $env:USERDNSDOMAIN by user $env:USERNAME"
		Line 2 "Ran from the folder $Script:pwdpath"
	}
	If ($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Report Footer&nbsp;&nbsp;\\\"
		WriteHTMLLine 2 0 "Report Information:"
		WriteHTMLLine 0 1 "Created with: $Script:ScriptName - Release Date: $Script:ReleaseDate"
		WriteHTMLLine 0 1 "Script version: $Script:MyVersion"
		WriteHTMLLine 0 1 "Started on $Script:StartTime"
		WriteHTMLLine 0 1 "Elapsed time: $Str"
		WriteHTMLLine 0 1 "Ran from domain $env:USERDNSDOMAIN by user $env:USERNAME"
		WriteHTMLLine 0 1 "Ran from the folder $Script:pwdpath"
	}
}

Function ProcessDocumentOutput
{
	Param([string] $Condition)
	
	If ($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}
	If ($Text)
	{
		SaveandCloseTextDocument
	}
	If ($HTML)
	{
		SaveandCloseHTMLDocument
	}

	If ($Condition -eq "Regular")
	{
		$GotFile = $False

		If ($MSWord)
		{
			If (Test-Path "$($Script:WordFileName)")
			{
				Write-Verbose "$(Get-Date -Format G): $($Script:WordFileName) is ready for use"
				$GotFile = $True
			}
			Else
			{
				Write-Warning "$(Get-Date -Format G): Unable to save the output file, $($Script:WordFileName)"
				Write-Error "Unable to save the output file, $($Script:WordFileName)"
			}
		}
		If ($PDF)
		{
			If (Test-Path "$($Script:PDFFileName)")
			{
				Write-Verbose "$(Get-Date -Format G): $($Script:PDFFileName) is ready for use"
				$GotFile = $True
			}
			Else
			{
				Write-Warning "$(Get-Date -Format G): Unable to save the output file, $($Script:PDFFileName)"
				Write-Error "Unable to save the output file, $($Script:PDFFileName)"
			}
		}
		If ($Text)
		{
			If (Test-Path "$($Script:TextFileName)")
			{
				Write-Verbose "$(Get-Date -Format G): $($Script:TextFileName) is ready for use"
				$GotFile = $True
			}
			Else
			{
				Write-Warning "$(Get-Date -Format G): Unable to save the output file, $($Script:TextFileName)"
				Write-Error "Unable to save the output file, $($Script:TextFileName)"
			}
		}
		If ($HTML)
		{
			If (Test-Path "$($Script:HTMLFileName)")
			{
				Write-Verbose "$(Get-Date -Format G): $($Script:HTMLFileName) is ready for use"
				$GotFile = $True
			}
			Else
			{
				Write-Warning "$(Get-Date -Format G): Unable to save the output file, $($Script:HTMLFileName)"
				Write-Error "Unable to save the output file, $($Script:HTMLFileName)"
			}
		}
		
		#email output file   If requested
		If ($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
		{
			If ($MSWord)
			{
				$emailAttachment = $Script:WordFileName
				SendEmail $emailAttachment
			}
			If ($PDF)
			{
				$emailAttachment = $Script:PDFFileName
				SendEmail $emailAttachment
			}
			If ($Text)
			{
				$emailAttachment = $Script:TextFileName
				SendEmail $emailAttachment
			}
			If ($HTML)
			{
				$emailAttachment = $Script:HTMLFileName
				SendEmail $emailAttachment
			}
		}
	}
}

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): Add DateTime         : $($AddDateTime)"
	If ($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): Company Name         : $($Script:CoName)"
		Write-Verbose "$(Get-Date -Format G): Company Address      : $($CompanyAddress)"
		Write-Verbose "$(Get-Date -Format G): Company Email        : $($CompanyEmail)"
		Write-Verbose "$(Get-Date -Format G): Company Fax          : $($CompanyFax)"
		Write-Verbose "$(Get-Date -Format G): Company Phone        : $($CompanyPhone)"
		Write-Verbose "$(Get-Date -Format G): Cover Page           : $($CoverPage)"
	}
	Write-Verbose "$(Get-Date -Format G): Dev                  : $($Dev)"
	If ($Dev)
	{
		Write-Verbose "$(Get-Date -Format G): DevErrorFile         : $($Script:DevErrorFile)"
	}
	If ($MSWord)
	{
		Write-Verbose "$(Get-Date -Format G): Word FileName        : $($Script:WordFileName)"
	}
	If ($HTML)
	{
		Write-Verbose "$(Get-Date -Format G): HTML FileName        : $($Script:HtmlFileName)"
	} 
	If ($PDF)
	{
		Write-Verbose "$(Get-Date -Format G): PDF FileName         : $($Script:PDFFileName)"
	}
	If ($Text)
	{
		Write-Verbose "$(Get-Date -Format G): Text FileName        : $($Script:TextFileName)"
	}
	Write-Verbose "$(Get-Date -Format G): Folder               : $($Folder)"
	Write-Verbose "$(Get-Date -Format G): From                 : $($From)"
	Write-Verbose "$(Get-Date -Format G): Host or Pool         : $($ServerName)"
	Write-Verbose "$(Get-Date -Format G): Log                  : $($Log)"
	Write-Verbose "$(Get-Date -Format G): Report Footer        : $ReportFooter"
	Write-Verbose "$(Get-Date -Format G): Save As HTML         : $($HTML)"
	Write-Verbose "$(Get-Date -Format G): Save As PDF          : $($PDF)"
	Write-Verbose "$(Get-Date -Format G): Save As TEXT         : $($TEXT)"
	Write-Verbose "$(Get-Date -Format G): Save As WORD         : $($MSWORD)"
	Write-Verbose "$(Get-Date -Format G): ScriptInfo           : $($ScriptInfo)"
	Write-Verbose "$(Get-Date -Format G): Section              : $($Section)"
	Write-Verbose "$(Get-Date -Format G): Smtp Port            : $($SmtpPort)"
	Write-Verbose "$(Get-Date -Format G): Smtp Server          : $($SmtpServer)"
	Write-Verbose "$(Get-Date -Format G): Title                : $($Script:Title)"
	Write-Verbose "$(Get-Date -Format G): To                   : $($To)"
	Write-Verbose "$(Get-Date -Format G): Use SSL              : $($UseSSL)"
	Write-Verbose "$(Get-Date -Format G): User                 : $($Script:User)"
	If ($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): User Name            : $($UserName)"
	}
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): OS Detected          : $($Script:RunningOS)"
	Write-Verbose "$(Get-Date -Format G): PoSH version         : $($Host.Version)"
	Write-Verbose "$(Get-Date -Format G): PSCulture            : $($PSCulture)"
	Write-Verbose "$(Get-Date -Format G): PSUICulture          : $($PSUICulture)"
	Write-Verbose "$(Get-Date -Format G): XenServer Version    : $($Script:XSVersion)"
	If ($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): Word language        : $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date -Format G): Word version         : $($Script:WordProduct)"
	}
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): Script start         : $($Script:StartTime)"
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region email function
Function SendEmail
{
	Param([array]$Attachments)
	Write-Verbose "$(Get-Date -Format G): Prepare to email"

	$emailAttachment = $Attachments
	$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.

"@ 

	If ($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
	}

	$error.Clear()
	
	If ($From -Like "anonymous@*")
	{
		#https://serverfault.com/questions/543052/sending-unauthenticated-mail-through-ms-exchange-with-powershell-windows-server
		$anonUsername = "anonymous"
		$anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
		$anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername, $anonPassword)

		If ($UseSSL)
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
				-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
				-UseSsl -Credential $anonCredentials *>$Null 
		}
		Else
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
				-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
				-Credential $anonCredentials *>$Null 
		}
		
		If ($?)
		{
			Write-Verbose "$(Get-Date -Format G): Email successfully sent using anonymous credentials"
		}
		ElseIf (!$?)
		{
			$e = $error[0]

			Write-Verbose "$(Get-Date -Format G): Email was not sent:"
			Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
		}
	}
	Else
	{
		If ($UseSSL)
		{
			Write-Verbose "$(Get-Date -Format G): Trying to send email using current user's credentials with SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
				-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
				-UseSsl *>$Null
		}
		Else
		{
			Write-Verbose  "$(Get-Date -Format G): Trying to send email using current user's credentials without SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
				-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
		}

		If (!$?)
		{
			$e = $error[0]
			
			#error 5.7.57 is O365 and error 5.7.0 is gmail
			If ($null -ne $e.Exception -and $e.Exception.ToString().Contains("5.7"))
			{
				#The server response was: 5.7.xx SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
				Write-Verbose "$(Get-Date -Format G): Current user's credentials failed. Ask for usable credentials."

				If ($Dev)
				{
					Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
				}

				$error.Clear()

				$emailCredentials = Get-Credential -UserName $From -Message "Enter the password to send email"

				If ($UseSSL)
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
						-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
						-UseSsl -Credential $emailCredentials *>$Null 
				}
				Else
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
						-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
						-Credential $emailCredentials *>$Null 
				}

				If ($?)
				{
					Write-Verbose "$(Get-Date -Format G): Email successfully sent using new credentials"
				}
				ElseIf (!$?)
				{
					$e = $error[0]

					Write-Verbose "$(Get-Date -Format G): Email was not sent:"
					Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date -Format G): Email was not sent:"
				Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
			}
		}
	}
}
#endregion

#region script start function
Function ProcessScriptSetup
{
	$script:startTime = Get-Date

	#make sure XenServerPSModule module is loaded
	If (!(Check-LoadedModule "XenServerPSModule"))
	{
		Write-Error "
		`n`n
		The XenServerPSModule module could not be loaded.
		`n`n
		Are you running this script against a XenServer 8.2 host or pool?
		`n`n
		Please see the Prerequisites section in the ReadMe file (XS_InventoryReadMe.rtf).
		`n`n
		insert sharefile link to readme
		`n`n
		Script cannot continue.
		`n`n
		"
		Write-Verbose "$(Get-Date -Format G): "
		AbortScript
	}

	#If computer name is an IP address, get host name from DNS
	#http://blogs.technet.com/b/gary/archive/2009/08/29/resolve-ip-addresses-to-hostname-using-powershell.aspx
	#help from Michael B. Smith
	$ip = $Script:ServerName -as [System.Net.IpAddress]
	If ($ip)
	{
		try
		{ 
			$Result = [System.Net.Dns]::gethostentry($ip) 
		}
		catch
		{
			$Result = $null
		}
		
		If ($? -and $Null -ne $Result)
		{
			$Script:ServerName = $Result.HostName
			Write-Verbose "$(Get-Date -Format G): Server name has been changed from $ip to $Script:ServerName"
		}
		Else
		{
			Write-Warning "Unable to resolve $Script:ServerName to a hostname"
		}
	}
	Else
	{
		#server is online but for some reason $Script:ServerName cannot be converted to a System.Net.IpAddress
	}

	If (![String]::IsNullOrEmpty($Script:ServerName))
	{
		#get server name
		#first test to make sure the server is reachable
		Write-Verbose "$(Get-Date -Format G): Testing to see if $Script:ServerName is online and reachable"
		If (Test-Connection -ComputerName $Script:ServerName -Quiet -EA 0)
		{
			Write-Verbose "$(Get-Date -Format G): Server $Script:ServerName is online."
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Server $Script:ServerName is offline"
			$ErrorActionPreference = $SaveEAPreference
			Write-Error "
			`n`n
			Computer $Script:ServerName is either offline or is not a valid XenServer Host or Pool name.
			`n`n
			Script cannot continue.
			`n`n
			"
			AbortScript
		}
	}

	#attempt to connect to the XenServer Host or Pool
	
	#get XenServer host login credentials
	Write-Verbose "$(Get-Date -Format G): Get login credentials"
	$script:XSCredentials = Get-Credential -UserName $User -Message "Enter the XenServer login credentials" 
	
	#connect to XenServer host/pool
	Write-Verbose "$(Get-Date -Format G): Attempt to connect to XenServer $Script:ServerName"
	
	try
	{
		$Script:Session = Connect-XenServer -Server $Script:ServerName -Creds $XSCredentials -SetDefaultSession -NoWarnCertificates -PassThru 4>$Null
	}
	
	catch  [XenAPI.Failure] 
	{
		If ($_.Exception.ErrorDescription[0] -eq "HOST_IS_SLAVE")
		{
			$tmp = $Script:ServerName

			$Script:ServerName = $_.Exception.ErrorDescription[1]

			$ip = $Script:ServerName -as [System.Net.IpAddress]

			If ($ip)
			{
				try
				{ 
					$Result = [System.Net.Dns]::gethostentry($ip) 
				}
				
				catch
				{
					$Result = $null
				}
				
				If ($? -and $Null -ne $Result)
				{
					$Script:ServerName = $Result.HostName
					Write-Verbose "$(Get-Date -Format G): Server name has been changed from $ip to $Script:ServerName"
				}
				Else
				{
					Write-Warning "Unable to resolve $Script:ServerName to a hostname"
				}
			}
			Else
			{
				#server is online but for some reason $Script:ServerName cannot be converted to a System.Net.IpAddress
			}
			Write-Host "
		$tmp is a Slave. 
		Attempt to connect to the Pool Master $Script:ServerName
			" -ForegroundColor White
		}
	}

	$Script:Session = Connect-XenServer -Server $Script:ServerName -Creds $XSCredentials -SetDefaultSession -NoWarnCertificates -PassThru 4>$Null
	
	If ($? -and $Null -ne $Script:Session)
	{
		Write-Host "
		Successfully connected to the Pool Master $Script:ServerName
		" -ForegroundColor White
		#success
	}
	Else
	{
		#error
		Write-Error "Unable to connect to XenServer Pool Master $($Script:Server). Script cannot continue."
		Return $False
	}

	Write-Verbose "$(Get-Date -Format G): Retrieve XenServer hosts"
	$tmptext = "XenServer Hosts"
	$Script:XSHosts = Get-XenHost -SessionOpaqueRef $Script:Session.Opaque_Ref -EA 0
	If ($? -and $Null -ne $Script:XSHosts)
	{
		#success
		$Script:XSHosts = $Script:XSHosts | Sort-Object Name_Label
	}
	ElseIf ($? -and $Null -eq $Script:XSHosts)
	{
		#success but no data and should not proceed with no host data
		Write-Error "There are no $($TmpText).  Script cannot continue."
		Return $False
	}
	Else
	{
		#error
		Write-Error "Unable to retrieve $($TmpText).  Script cannot continue."
		Return $False
	}
	
	Write-Verbose "$(Get-Date -Format G): Retrieve Pool data"
	$tmptext = "XenServer Pool"
	$Script:XSPool = Get-XenPool -SessionOpaqueRef $Session.opaque_ref -EA 0
	If ($? -and $Null -ne $Script:XSPool)
	{
		#success
		$Script:PoolMasterInfo = ($Script:XSPool | Get-XenPoolProperty -XenProperty master -EA 0)
		$Script:OtherConfig = ($XSPool | Get-XenPoolProperty -XenProperty OtherConfig -EA 0)
	}
	ElseIf ($? -and $Null -eq $Script:XSPool)
	{
		#success but no data and should not proceed with no pool data
		Write-Error "There is no $($TmpText)"
		Return $False
	}
	Else
	{
		#error
		Write-Error "Unable to retrieve $($TmpText)"
		Return $False
	}

	Write-Verbose "$(Get-Date -Format G): Get XenServer Version"
	$Script:XSVersion = [version]$Script:PoolMasterInfo.software_version.product_version
	
	If ($Null -ne $Script:XSVersion)
	{
		#this script is only for XS 8.2
		If ($Script:XSVersion -ge [version]"8.2")
		{
			#we are good
		}
		Else
		{
			#wrong XS version
			Write-Host "You are running XenServer version $Script:XSVersion" -ForegroundColor White
			Write-Error "
	`n`n
	This script is designed for XenServer 8.2 and should not be run on $Script:XSVersion.
	`n`n
	Script cannot continue
	`n`n
		"
			AbortScript
		}
	}
	Else
	{
		Write-Error "
	`n`n
	This script is designed for XenServer 8.2 and your XenServer version could not be determined.
	`n`n
	Script cannot continue
	`n`n
		"
		AbortScript
	}

	Write-Verbose "$(Get-Date -Format G): Running XenServer version $($Script:XSVersion)"
	Write-Verbose "$(Get-Date -Format G):"
	
	Write-Verbose "$(Get-Date -Format G): Retrieve list of VM names"
	#get a list of VM names for VMs that are not templates or snapshot or control domain
	#also exclude hidden VMs http://support.citrix.com/proddocs/topic/xencenter-62/xs-xc-intro-hiddenobjects.html
	$tmptext = "Virtual Machines"
	$tmp = 'true'
	$strkey = 'HideFromXenCenter'
	$Script:VMNames = Get-XenVM -SessionOpaqueRef $Script:Session.Opaque_Ref -EA 0 | `
			Where-Object { !$_.is_a_template -and `
				!$_.is_a_snapshot -and `
				!$_.is_control_domain -and `
				!$_.other_config.TryGetValue($strkey, [ref]$tmp) } | `
			Select-Object name_label | `
			Sort-Object name_label	
	If ($? -and $Null -ne $Script:VMNames)
	{
		#success
		$Script:VMNames = $Script:VMNames | Sort-Object Name_Label
	}
	ElseIf ($? -and $Null -eq $Script:VMNames)
	{
		#success but no data 
		Write-Warning "There are no $($TmpText)"
	}
	Else
	{
		#error
		Write-Warning "Unable to retrieve $($TmpText)"
	}

	#support multiple section items
	If ($Section.Count -eq 1 -and $Section -eq "All")
	{
		[string]$Script:Title = "Citrix XenServer Inventory"
	}
	ElseIf ($Section.Count -eq 1)
	{
		Switch ($Section)
		{
			"Pool"	{ [string]$Script:Title = "Citrix XenServer Inventory (Pool Only)"; Break }
			"Host"	{ [string]$Script:Title = "Citrix XenServer Inventory (Hosts Only)"; Break }
			"VM"	{ [string]$Script:Title = "Citrix XenServer Inventory (VMs Only"; Break }
			Default	{ [string]$Script:Title = "Citrix XenServer Inventory (Missing a section title for $Section"; Break }
		}
	}
	ElseIf ($Section.Count -gt 1)
	{
		[string]$Script:Title = "Citrix XenServer Inventory ("
		Switch ($Section)
		{
			"Pool"	{ [string]$Script:Title += "Pool " }
			"Host"	{ [string]$Script:Title += "Hosts " }
			"VM"	{ [string]$Script:Title += "VMs " }
			Default	{ [string]$Script:Title += "Missing a section title for $Section" }
		}
		[string]$Script:Title = $Script:Title.Substring(0, $Script:Title.LastIndexOf(" ")) + ")"
	}
	Return $True
}
#endregion

#region script end
Function ProcessScriptEnd
{
	Write-Verbose "$(Get-Date -Format G): Script has completed"
	Write-Verbose "$(Get-Date -Format G): "

	#http://poshtips.com/measuring-elapsed-time-in-powershell/
	Write-Verbose "$(Get-Date -Format G): Script started: $($Script:StartTime)"
	Write-Verbose "$(Get-Date -Format G): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days,   {1} hours,   {2} minutes,   {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date -Format G): Elapsed time: $($Str)"

	If ($Dev)
	{
		If ($SmtpServer -eq "")
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}
		Else
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}
	}

	If ($ScriptInfo)
	{
		$SIFile = "$Script:pwdpath\XSInventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime         : $AddDateTime" 4>$Null
		If ($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Name         : $Script:CoName" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Address      : $CompanyAddress" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Email        : $CompanyEmail" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Fax          : $CompanyFax" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Phone        : $CompanyPhone" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Cover Page           : $CoverPage" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Dev                  : $Dev" 4>$Null
		If ($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile         : $Script:DevErrorFile" 4>$Null
		}
		If ($MSWord)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word FileName        : $Script:WordFileName" 4>$Null
		}
		If ($HTML)
		{
			Out-File -FilePath $SIFile -Append -InputObject "HTML FileName        : $Script:HtmlFileName" 4>$Null
		}
		If ($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "PDF Filename         : $Script:PDFFileName" 4>$Null
		}
		If ($Text)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Text FileName        : $Script:TextFileName" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder               : $Folder" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From                 : $From" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Host/Pool            : $Script:ServerName" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Log                  : $Log" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Report Footer        : $ReportFooter" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As HTML         : $HTML" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF          : $PDF" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As TEXT         : $TEXT" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD         : $MSWORD" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info          : $ScriptInfo" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Section              : $($Section)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port            : $SmtpPort" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server          : $SmtpServer" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title                : $Script:Title" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To                   : $To" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL              : $UseSSL" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "User                 : $Script:User" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "XenServer Version    : $($Script:XSVersion)"
		If ($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "User Name            : $UserName" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected          : $Script:RunningOS" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version         : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture            : $PSCulture" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture          : $PSUICulture" 4>$Null
		If ($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word language        : $Script:WordLanguageValue" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "Word version         : $Script:WordProduct" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start         : $Script:StartTime" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time         : $Str" 4>$Null
	}

	#stop transcript logging
	If ($Log -eq $True)
	{
		If ($Script:StartLog -eq $true)
		{
			try
			{
				Stop-Transcript | Out-Null
				Write-Verbose "$(Get-Date -Format G): $Script:LogPath is ready for use"
			}
			catch
			{
				Write-Verbose "$(Get-Date -Format G): Transcript/log stop failed"
			}
		}
	}
	
	#cleanup obj variables
	$Script:Output = $Null
}
#endregion

#region XS Specific Functions
Function Get-XSCustomFields 
{
	Param([hashtable] $OtherConfig)
	#Write-Verbose "$(Get-Date -Format G): `tProcessing Custom Fields" #write-verbose commented out by Webster
	$CustomFields = New-Object System.Collections.ArrayList
	ForEach ($Item in $OtherConfig.Keys)
	{
		If ($Item -like "*customfields*")
		{
			$value = $($OtherConfig.$item)
			if ($value -as [DateTime])
			{
				#Thanks to Michael B. SMith for the next two lines
				$datetime = $value -as [DateTime]
				$value = $datetime.ToLongDateString() + ' ' + $datetime.ToLongTimeString()
			}
			$obj1 = [PSCustomObject] @{
				Name  = $Item.Replace("XenCenter.CustomFields.", $null)
				Value = $value
			}
			$Null = $CustomFields.Add($obj1)
		}
	}
	Write-Output  $CustomFields
}

function Convert-SizeToString
{
	param (
		[Parameter(Mandatory = $true)]
		[Int64]$Size,

		[Int]$Decimal = 2
	)

	$tb = 1TB
	$gb = 1GB
	$mb = 1MB
	$kb = 1KB

	If ($size -ge $tb)
	{
		$result = "{0:N$Decimal} TB" -f ($size / $tb)
	}
	ElseIf ($size -ge $gb)
	{
		$result = "{0:N$Decimal} GB" -f ($size / $gb)
	}
	ElseIf ($size -ge $mb)
	{
		$result = "{0:N$Decimal} MB" -f ($size / $mb)
	}
	ElseIf ($size -ge $kb)
	{
		$result = "{0:N$Decimal} KB" -f ($size / $kb)
	}
	Else
	{
		$result = "{0} B" -f $size
	}

	return $result
}

#endregion Functions


#region pool
Function ProcessPool
{
	Write-Verbose "$(Get-Date -Format G): Processing the $($Script:XSPool.name_label) Pool"
	If ($Null -eq $Script:XSPool.name_label)
	{
		Return
	}
	Else
	{
		If ($MSWord -or $PDF)
		{
			$Selection.InsertNewPage()
			WriteWordLine 1 0 "$($Script:XSPool.name_label) Pool"
		}
		If ($Text)
		{
			Line 0 "$($Script:XSPool.name_label) Pool"
			Line 0 ""
		}
		If ($HTML)
		{
			WriteHTMLLine 1 0 "$($Script:XSPool.name_label) Pool"
		}
		
		OutputPoolGeneralOverview
		OutputPoolUpdates
		OutputPoolManagementInterfaces
		OutputPoolGeneral
		OutputPoolCustomFields
		OutputPoolEmailOptions
		OutputPoolPowerOn
		OutputPoolLivePatching
		OutputPoolNetworkOptions
		OutputPoolClustering
		OutputPoolMemory
		OutputPoolStorage
		OutputPoolNetworking
		OutputPoolGPU
		OutputPoolHA
		OutputPoolWLB
		OutputPoolUsers
	}
}

Function OutputPoolGeneralOverview
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool General Overview"
	#This is what you see when look at Pool General Properties, not a pool's Properties, General
	[array]$xtags = @()
	ForEach ($tag in $Script:XSPool.tags)
	{
		$xtags += $tag
	}
	If ($xtags.count -gt 0)
	{
		[array]$xtags = $xtags | Sort-Object
	}
	Else
	{
		[array]$xtags = @("<None>")
	}
	
	$NumSockets = (($xshosts).cpu_info.socket_count | Measure-Object -Sum).sum
	
	$PoolLicense = ""

	<#
		express
		premium-per-socket
		premium-per-user
		standard-per-socket
		desktop
		desktop-plus
		desktop-cloud
	#>
	Switch ($Script:PoolMasterInfo.edition)
	{
		"express" { $PoolLicense = "Express" }
		"premium-per-socket"	{ $PoolLicense = "Citrix Hypervisor Premium Per-Socket" }
		"premium-per-user" { $PoolLicense = "Citrix Hypervisor Premium Per-User" }
		"standard-per-socket"	{ $PoolLicense = "Citrix Hypervisor Standard Per-Socket" }
		"desktop" { $PoolLicense = "Citrix Virtual Apps and Desktops" }
		"desktop-cloud" { $PoolLicense = "Citrix Virtual Apps and Desktops Citrix Cloud" }
		"desktop-plus" { $PoolLicense = "XenApp/XenDesktop Platinum" }
		Default { $PoolLicense = "Unable to determine Pool License: $($Script:PoolMasterInfo.edition)" }
	}
	
	If ([String]::IsNullOrEmpty($($Script:XSPool.Other_Config["folder"])))
	{
		$folderName = "None"
	}
	Else
	{
		$folderName = $Script:XSPool.Other_Config["folder"]
	}


	If ($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "General Overview"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Pool name"; Value = $Script:XSPool.name_label; }
		$ScriptInformation += @{ Data = "Description"; Value = $Script:XSPool.name_description; }
		$ScriptInformation += @{ Data = "Tags"; Value = $($xtags -join ", ") }
		$ScriptInformation += @{ Data = "Folder"; Value = $folderName; }
		$ScriptInformation += @{ Data = "Pool License"; Value = $PoolLicense; }
		$ScriptInformation += @{ Data = "Number of Sockets"; Value = $NumSockets; }
		$ScriptInformation += @{ Data = "XenServer Version"; Value = $Script:PoolMasterInfo.software_version.product_version_text_short; }
		$ScriptInformation += @{ Data = "UUID"; Value = $Script:XSPool.uuid; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 1 "General Overview"
		Line 2 "Pool name`t`t: " $Script:XSPool.name_label
		Line 2 "Description`t`t: " $Script:XSPool.name_description
		Line 2 "Tags`t`t`t: " $($xtags -join ", ")
		Line 2 "Folder`t`t`t: " $folderName
		Line 2 "Pool License`t`t: " $PoolLicense
		Line 2 "Number of Sockets`t: " $NumSockets
		Line 2 "XenServer Version`t: " $Script:PoolMasterInfo.software_version.product_version_text_short
		Line 0 ""
	}
	If ($HTML)
	{
		#for HTML output, remove the < and > from <None> xtags and foldername if they are there
		$xtags = $xtags.Trim("<", ">")
		$folderName = $folderName.Trim("<", ">")
		WriteHTMLLine 2 0 "General Overview"
		$rowdata = @()
		$columnHeaders = @("Pool name", ($htmlsilver -bor $htmlbold), $Script:XSPool.name_label, $htmlwhite)
		$rowdata += @(, ('Description', ($htmlsilver -bor $htmlbold), $Script:XSPool.name_description, $htmlwhite))
		$rowdata += @(, ('Tags', ($htmlsilver -bor $htmlbold), "$($xtags -join ", ")", $htmlwhite))
		$rowdata += @(, ('Folder', ($htmlsilver -bor $htmlbold), $folderName, $htmlwhite))
		$rowdata += @(, ('Pool License', ($htmlsilver -bor $htmlbold), $PoolLicense, $htmlwhite))
		$rowdata += @(, ('Number of Sockets', ($htmlsilver -bor $htmlbold), $NumSockets, $htmlwhite))
		$rowdata += @(, ('XenServer Version', ($htmlsilver -bor $htmlbold), $Script:PoolMasterInfo.software_version.product_version_text_short, $htmlwhite))
		$rowdata += @(, ('UUID', ($htmlsilver -bor $htmlbold), $Script:XSPool.uuid, $htmlwhite))

		$msg = ""
		$columnWidths = @("150", "250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputPoolUpdates
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool Updates"
	$Updates = Get-XenPoolPatch -SessionOpaqueRef $Script:Session.Opaque_Ref -EA 0 4>$Null | Select-Object name_label, version | Sort-Object name_label

	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $WordTable = @();
		
		WriteWordLine 2 0 "Updates" 
		
		ForEach ($tmp in $Updates)
		{
			$WordTableRowHash = @{ 
				Update = "$($tmp.name_label) (version $($tmp.version))";
			}
			$WordTable += $WordTableRowHash;
		}
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $WordTable `
			-Columns Update `
			-Headers "Fully applied" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 1 "Updates"
		<#
		Line 2 "Fully applied`t: " "$($Updates[0].name_label) (version $($tmp.version))"
		$cnt = -1
		ForEach($tmp in $Updates)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 4 "  " "$($tmp.name_label) (version $($tmp.version))"
			}
		}
		#>
		Line 2 "Fully applied`t: " ""
		ForEach ($tmp in $Updates)
		{
			Line 3 "" "$($tmp.name_label) (version $($tmp.version))"
		}

		Line 0 ""
	}
	If ($HTML)
	{
		WriteHTMLLine 2 0 "Updates"
		$rowdata = @()

		ForEach ($tmp in $Updates)
		{
			$rowdata += @(, (
					"$($tmp.name_label) (version $($tmp.version))", $htmlwhite))
		}
		
		$columnHeaders = @(
			'Fully applied', ($htmlsilver -bor $htmlbold))

		$msg = ""
		$columnWidths = @("150")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
	}
}

Function OutputPoolManagementInterfaces
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool Management Interfaces"
	#Sample code from the XenServer team
	<#
		this is what XenCenter does:
		$servers = get-xenhost
		
		foreach ($server in $servers)   {
			Write-host "DNS hostname:" $server.hostname
			
			$mPifs = $server.PIFs | get-xenPIF | where   {$_.management}

			foreach ($pif in $mPifs)   {
				If ($pif.IP)   {
					Write-Host "Management interface on" $server.name_label ":" $pif.IP
				}
				ElseIf ($pif.ip_configuration_mode -eq [XenAPI.ip_configuration_mode]::DHCP)   {
					Write-Host "Management interface on" $server.name_label ":" "DHCP"
				}
				Else   {
					Write-Host "Management interface on" $server.name_label ":" "Unknown"
				}
			} 
		}
	#>

	<#ForEach ($Server in $Servers) 
	{
		Write-host "DNS hostname:" $Server.hostname
		
		$mPifs = $Server.PIFs | Get-XenPIF -EA 0 | Where-Object   {$_.management}

		ForEach ($pif in $mPifs) 
		{
			If ($pif.IP) 
			{
				Write-Host "Management interface on" $server.name_label ":" $pif.IP
			}
			ElseIf ($pif.ip_configuration_mode -eq [XenAPI.ip_configuration_mode]::DHCP) 
			{
				Write-Host "Management interface on" $server.name_label ":" "DHCP"
			}
			Else 
			{
				Write-Host "Management interface on" $server.name_label ":" "Unknown"
			}
		} 
	}#>

	If ($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Management Interfaces"
		$MITable = @()

		ForEach ($Server in $Script:XSHosts)
		{
			$MITable += @{
				Column1 = "DNS hostname on $($server.name_label):"
				Column2 = $Server.hostname
			}
			
			$mPifs = $Server.PIFs | Get-XenPIF -EA 0 | Where-Object { $_.management }

			ForEach ($pif in $mPifs)
			{
				If ($pif.IP)
				{
					$MITable += @{
						Column1 = "Management interface on $($server.name_label):"
						Column2 = "$($pif.IP)"
					}
				}
				ElseIf ($pif.ip_configuration_mode -eq [XenAPI.ip_configuration_mode]::DHCP)
				{
					$MITable += @{
						Column1 = "Management interface on $($server.name_label):"
						Column2 = "DHCP"
					}
				}
				Else
				{
					$MITable += @{
						Column1 = "Management interface on $($server.name_label):"
						Column2 = "Unknown"
					}
				}
			} 
		}

		If ($MITable.Count -gt 0)
		{
			$Table = AddWordTable -Hashtable $MITable `
				-Columns Column1, Column2 `
				-Headers "", "" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 250;
			$Table.Columns.Item(2).Width = 100;
			
			$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
	}
	If ($Text)
	{
		Line 2 "Management Interfaces"
		Line 2 "=============================================================="
		#		1234567890123456789012345678901234567890SS12345678901234567890
		#		Management interface on 'XenServer1       255.255.255.255
		ForEach ($Server in $Script:XSHosts)
		{
			Line 2 ( "{0,-40}    {1,-20}" -f "DNS hostname on $($server.name_label):", $Server.hostname)
			
			$mPifs = $Server.PIFs | Get-XenPIF -EA 0 | Where-Object { $_.management }

			ForEach ($pif in $mPifs)
			{
				If ($pif.IP)
				{
					Line 2 ( "{0,-40}    {1,-20}" -f "Management interface on $($server.name_label):", "$($pif.IP)")
				}
				ElseIf ($pif.ip_configuration_mode -eq [XenAPI.ip_configuration_mode]::DHCP)
				{
					Line 2 ( "{0,-40}    {1,-20}" -f "Management interface on $($server.name_label):", "DHCP")
				}
				Else
				{
					Line 2 ( "{0,-40}    {1,-20}" -f "Management interface on $($server.name_label):", "Unknown")
				}
			} 
		}
		Line 0 ""
	}
	If ($HTML)
	{
		WriteHTMLLine 2 0 "Management Interfaces"

		$rowdata = @()

		ForEach ($Server in $Script:XSHosts)
		{
			$rowdata += @(, (
					"DNS hostname on $($server.name_label):", $htmlwhite,
					$Server.hostname, $htmlwhite)
			)
			
			$mPifs = $Server.PIFs | Get-XenPIF -EA 0 | Where-Object { $_.management }

			ForEach ($pif in $mPifs)
			{
				If ($pif.IP)
				{
					$rowdata += @(, (
							"Management interface on $($server.name_label):", $htmlwhite,
							"$($pif.IP)", $htmlwhite)
					)
				}
				ElseIf ($pif.ip_configuration_mode -eq [XenAPI.ip_configuration_mode]::DHCP)
				{
					$rowdata += @(, (
							"Management interface on $($server.name_label):", $htmlwhite,
							"DHCP", $htmlwhite)
					)
				}
				Else
				{
					$rowdata += @(, (
							"Management interface on $($server.name_label):", $htmlwhite,
							"Unknown", $htmlwhite)
					)
				}
			} 
		}

		$columnHeaders = @(
			"", ($Script:htmlsb),
			"", ($Script:htmlsb)
		)

		$msg = ""
		$columnWidths = @("250", "100")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputPoolGeneral
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool General"
	[array]$xtags = @()
	ForEach ($tag in $Script:XSPool.tags)
	{
		$xtags += $tag
	}
	If ($xtags.count -gt 0)
	{
		[array]$xtags = $xtags | Sort-Object
	}
	Else
	{
		[array]$xtags = @("<None>")
	}
	
	If ([String]::IsNullOrEmpty($($Script:XSPool.Other_Config["folder"])))
	{
		$folderName = "None"
	}
	Else
	{
		$folderName = $Script:XSPool.Other_Config["folder"]
	}


	If ($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "General"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Name"; Value = $Script:XSPool.name_label; }
		$ScriptInformation += @{ Data = "Description"; Value = $Script:XSPool.name_description; }
		$ScriptInformation += @{ Data = "Folder"; Value = $folderName; }
		$ScriptInformation += @{ Data = "Tags"; Value = $($xtags -join ", ") }
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 1 "General"
		Line 2 "Name`t`t: " $Script:XSPool.name_label
		Line 2 "Description`t`t: " $Script:XSPool.name_description
		Line 2 "Folder`t`t`t: " $folderName
		Line 2 "Tags`t`t`t: " $($xtags -join ", ")
		Line 0 ""
	}
	If ($HTML)
	{
		#for HTML output, remove the < and > from <None> xtags and foldername if they are there
		$xtags = $xtags.Trim("<", ">")
		$folderName = $folderName.Trim("<", ">")
		WriteHTMLLine 2 0 "General"
		$rowdata = @()
		$columnHeaders = @("Name", ($htmlsilver -bor $htmlbold), $Script:XSPool.name_label, $htmlwhite)
		$rowdata += @(, ('Description', ($htmlsilver -bor $htmlbold), $Script:XSPool.name_description, $htmlwhite))
		$rowdata += @(, ('Folder', ($htmlsilver -bor $htmlbold), $folderName, $htmlwhite))
		$rowdata += @(, ('Tags', ($htmlsilver -bor $htmlbold), "$($xtags -join ", ")", $htmlwhite))

		$msg = ""
		$columnWidths = @("150", "250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputPoolCustomFields
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool Custom Fields"

	$CustomFields = Get-XSCustomFields $($Script:XSPool.other_config)

	If ($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Custom Fields"
	}
	If ($Text)
	{
		Line 1 "Custom Fields"
	}
	If ($HTML)
	{
		WriteHTMLLine 2 0 "Custom Fields"
	}
	
	If ([String]::IsNullOrEmpty($CustomFields) -or $CustomFields.Count -eq 0)
	{
		$PoolName = $Script:XSPool.Name_Label
		If ($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There are no Custom Fields for Pool $PoolName"
			WriteWordLine 0 0 ""
		}
		If ($Text)
		{
			Line 2 "There are no Custom Fields for Pool $PoolName"
			Line 0 ""
		}
		If ($HTML)
		{
			WriteHTMLLine 0 1 "There are no Custom Fields for Pool $PoolName"
			WriteHTMLLine 0 0 ""
		}
	}
	Else
	{
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
		}
		If ($Text)
		{
			#nothing
		}
		If ($HTML)
		{
			$rowdata = @()
		}

		[int]$cnt = -1
		ForEach ($Item in $CustomFields)
		{
			$cnt++
			If ($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = $($Item.Name); Value = $Item.Value; }
			}
			If ($Text)
			{
				Line 2 "$($Item.Name): " $Item.Value
			}
			If ($HTML)
			{
				If ($cnt -eq 0)
				{
					$columnHeaders = @($($Item.Name), ($htmlsilver -bor $htmlbold), $Item.Value, $htmlwhite)
				}
				Else
				{
					$rowdata += @(, ($($Item.Name), ($htmlsilver -bor $htmlbold), $Item.Value, $htmlwhite))
				}
			}
		}
		
		If ($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data, Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 250;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If ($Text)
		{
			Line 0 ""
		}
		If ($HTML)
		{
			$msg = ""
			$columnWidths = @("250", "250")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
}

Function OutputPoolEmailOptions
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool Email Options"

	$EmailOptions = New-Object System.Collections.ArrayList
	
	ForEach ($Item in $Script:OtherConfig.Keys)
	{
		If ($Item -like "*mail*")
		{
			$obj1 = [PSCustomObject] @{
				Name  = $Item
				Value = $($OtherConfig.$item)
			}
			$Null = $EmailOptions.Add($obj1)
		}
	}
	
	If ($EmailOptions.Count -gt 0)
	{
		$MailEnabled = $True
	}
	Else
	{
		$MailEnabled = $False
	}

	If ($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Email Options"
	}
	If ($Text)
	{
		Line 1 "Email Options"
	}
	If ($HTML)
	{
		WriteHTMLLine 2 0 "Email Options"
	}
	
	If ($MailEnabled)
	{
		#Name                           Value
		#----                           -----
		#ssmtp-mailhub                  smtp.office365.com:587
		#mail-destination               webster@carlwebster.com
		#mail-language                  en-US
		
		#en-US = English (United States)
		#zh-CN = Chinese (Simplified)
		#ja-JP = Japanese (Japan)

		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Delivery address"; Value = ""; }
		}
		If ($Text)
		{
			#nothing
			Line 2 "Delivery address"
		}
		If ($HTML)
		{
			$rowdata = @()
			$columnHeaders = @("Delivery address", ($htmlsilver -bor $htmlbold), "", $htmlwhite)
		}
		
		ForEach ($Item in $EmailOptions)
		{
			If ($Item.Name -eq "ssmtp-mailhub")
			{
				[array]$TmpArray = $Item.Value.Split(":")
				$DAText = "$($TmpArray[0])"
				$DAData = "Port: $($TmpArray[1])"
				If ($MSWord -or $PDF)
				{
					$ScriptInformation += @{ Data = "     SMTP server"; Value = "$($DAText)     $($DAData)"; }
				}
				If ($Text)
				{
					Line 3 "SMTP server`t: $($DAText)`t" $DAData
				}
				If ($HTML)
				{
					$rowdata += @(, ("     SMTP server", ($htmlsilver -bor $htmlbold), "$($DAText)     $($DAData)", $htmlwhite))
				}
			}
			ElseIf ($Item.Name -eq "mail-destination")
			{
				If ($MSWord -or $PDF)
				{
					$ScriptInformation += @{ Data = "     Email address"; Value = $Item.Value; }
				}
				If ($Text)
				{
					Line 3 "Email address`t: " $Item.Value
				}
				If ($HTML)
				{
					$rowdata += @(, ("     Email address", ($htmlsilver -bor $htmlbold), $Item.Value, $htmlwhite))
				}
			}
			ElseIf ($Item.Name -eq "mail-language")
			{
				Switch ($Item.Value)
				{
					"en-US" { $DAData = "English (United States)"; Break }
					"ja-JP" { $DAData = "Japanese (Japan)"; Break }
					"zh-CN" { $DAData = "Chinese (Simplified)"; Break }
					Default { $DAData = "Unable to determine the email language: $($Item.Value)"; Break }
				}

				If ($MSWord -or $PDF)
				{
					$ScriptInformation += @{ Data = "     Mail language"; Value = $DAData; }
				}
				If ($Text)
				{
					Line 3 "Mail language`t: " $DAData
				}
				If ($HTML)
				{
					$rowdata += @(, ("     Mail language", ($htmlsilver -bor $htmlbold), $DAData, $htmlwhite))
				}
			}
			Else
			{
				#oops we shouldn't be here
				If ($MSWord -or $PDF)
				{
					$ScriptInformation += @{ Data = "     $($Item.Name)"; Value = $Item.Value; }
				}
				If ($Text)
				{
					Line 3 "$($Item.Name)`t: " $Item.Value
				}
				If ($HTML)
				{
					$rowdata += @(, ("     $($Item.Name)", ($htmlsilver -bor $htmlbold), $Item.Value, $htmlwhite))
				}
			}
		}
	}
	Else
	{
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Send email alert notifications"; Value = "Disabled"; }
		}
		If ($Text)
		{
			Line 2 "Send email alert notifications: " "Disabled"
		}
		If ($HTML)
		{
			$rowdata = @()
			$columnHeaders = @("Send email alert notifications", ($htmlsilver -bor $htmlbold), "Disabled", $htmlwhite)
		}
	}
	
	If ($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 0 ""
	}
	If ($HTML)
	{
		$msg = ""
		$columnWidths = @("150", "250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputPoolPowerOn
{
	#cycle through each host and get the power_on_mode and power_on_config properties
	#If power_on_mode -eq "", then it is Disabled
	#DRAC is Dell Remote Access Controller (DRAC)
	#wake-on-lan is Wake-on-LAN (WoL)
	#Otherwise, power_on_mode is Custom power-on script /etc/xapi.d/plugins/<value of power_on_mode>
	
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool Power On"

	If ($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Power On"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
	}
	If ($Text)
	{
		Line 1 "Power On"
	}
	If ($HTML)
	{
		WriteHTMLLine 2 0 "Power On"
		$rowdata = @()
	}
	
	[int]$cnt = -1
	ForEach ($XSHost in $Script:XSHosts)
	{
		$cnt++
		If ($XSHost.power_on_mode -eq "")
		{
			#disabled
			If ($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "Server"; Value = "$($XSHost.Name_Label)"; }
				$ScriptInformation += @{ Data = "Power On mode"; Value = "Disabled"; }
				$ScriptInformation += @{ Data = ""; Value = ""; }
			}
			If ($Text)
			{
				Line 2 "Server`t`t: " "$($XSHost.Name_Label)"
				Line 2 "Power On mode`t: " "Disabled"
				Line 0 ""
			}
			If ($HTML)
			{
				If ($cnt -eq 0)
				{
					$columnHeaders = @("Server", ($htmlsilver -bor $htmlbold), "$($XSHost.Name_Label)", $htmlwhite)
					$rowdata += @(, ("    Power On mode", ($htmlsilver -bor $htmlbold), "Disabled", $htmlwhite))
					$rowdata += @(, ("", ($htmlsilver -bor $htmlbold), "", $htmlwhite))
				}
				Else
				{
					$rowdata += @(, ("Server", ($htmlsilver -bor $htmlbold), "$($XSHost.Name_Label)", $htmlwhite))
					$rowdata += @(, ("    Power On mode", ($htmlsilver -bor $htmlbold), "Disabled", $htmlwhite))
					$rowdata += @(, ("", ($htmlsilver -bor $htmlbold), "", $htmlwhite))
				}
			}
		}
		ElseIf ($XSHost.power_on_mode -eq "DRAC")
		{
			[array]$PowerKeys = $XSHost.power_on_config.Keys.Split() 
			[array]$PowerValues = $XSHost.power_on_config.Values.Split() 
			If ($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "Server"; Value = "$($XSHost.Name_Label)"; }
				$ScriptInformation += @{ Data = "Power On mode"; Value = "Dell Remote Access Controller (DRAC)"; }
				$ScriptInformation += @{ Data = "Configuration options"; Value = ""; }

				[int]$cnt2 = -1
				ForEach ($Item in $PowerKeys)
				{
					$cnt2++
					$Value = $PowerValues[$cnt2]
					
					If ($Item -like "*power_on_ip*")
					{
						$ScriptInformation += @{ Data = "     IP address"; Value = $Value; }
					}
					If ($Item -like "*power_on_user*")
					{
						$ScriptInformation += @{ Data = "     Username"; Value = $Value; }
					}
				}
				$ScriptInformation += @{ Data = ""; Value = ""; }
			}
			If ($Text)
			{
				Line 2 "Server`t`t: " "$($XSHost.Name_Label)"
				Line 2 "Power On mode`t: " "Dell Remote Access Controller (DRAC)"
				Line 2 "Configuration options"

				[int]$cnt2 = -1
				ForEach ($Item in $PowerKeys)
				{
					$cnt2++
					$Value = $PowerValues[$cnt2]
					
					If ($Item -like "*power_on_ip*")
					{
						Line 3 "IP address: " $Value
					}
					If ($Item -like "*power_on_user*")
					{
						Line 3 "Username  : " $Value
					}
				}
				Line 0 ""
			}
			If ($HTML)
			{
				If ($cnt -eq 0)
				{
					$columnHeaders = @("Server", ($htmlsilver -bor $htmlbold), "$($XSHost.Name_Label)", $htmlwhite)
				}
				Else
				{
					$rowdata += @(, ("Server", ($htmlsilver -bor $htmlbold), "$($XSHost.Name_Label)", $htmlwhite))
				}
				$rowdata += @(, ("Power On mode", ($htmlsilver -bor $htmlbold), "Dell Remote Access Controller (DRAC)", $htmlwhite))
				$rowdata += @(, ("Configuration options", ($htmlsilver -bor $htmlbold), "", $htmlwhite))

				[int]$cnt2 = -1
				ForEach ($Item in $PowerKeys)
				{
					$cnt2++
					$Value = $PowerValues[$cnt2]
					
					If ($Item -like "*power_on_ip*")
					{
						$rowdata += @(, ("     IP address: ", ($htmlsilver -bor $htmlbold), $Value, $htmlwhite))
					}
					If ($Item -like "*power_on_user*")
					{
						$rowdata += @(, ("     Username: ", ($htmlsilver -bor $htmlbold), $Value, $htmlwhite))
					}
				}
				$rowdata += @(, ("", ($htmlsilver -bor $htmlbold), "", $htmlwhite))
			}
		}
		ElseIf ($XSHost.power_on_mode -eq "wake-on-lan")
		{
			If ($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "Server"; Value = "$($XSHost.Name_Label)"; }
				$ScriptInformation += @{ Data = "Power On mode"; Value = "Wake-on-LAN (WoL)"; }
				$ScriptInformation += @{ Data = ""; Value = ""; }
			}
			If ($Text)
			{
				Line 2 "Server`t`t: " "$($XSHost.Name_Label)"
				Line 2 "Power On mode`t: " "Wake-on-LAN (WoL)"
				Line 0 ""
			}
			If ($HTML)
			{
				If ($cnt -eq 0)
				{
					$columnHeaders = @("Server", ($htmlsilver -bor $htmlbold), "$($XSHost.Name_Label)", $htmlwhite)
				}
				Else
				{
					$rowdata += @(, ("Server", ($htmlsilver -bor $htmlbold), "$($XSHost.Name_Label)", $htmlwhite))
				}
				$rowdata += @(, ("Power On mode", ($htmlsilver -bor $htmlbold), "Wake-on-LAN (WoL)", $htmlwhite))
				$rowdata += @(, ("", ($htmlsilver -bor $htmlbold), "", $htmlwhite))
			}
		}
		Else
		{
			#custom script
			[array]$PowerKeys = $XSHost.power_on_config.Keys.Split() 
			[array]$PowerValues = $XSHost.power_on_config.Values.Split() 
			If ($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "Server"; Value = "$($XSHost.Name_Label)"; }
				$ScriptInformation += @{ Data = "Power On mode"; Value = "Custom power-on script /etc/xapi.d/plugins/$($XSHost.power_on_mode)"; }
				$ScriptInformation += @{ Data = "Configuration options"; Value = ""; }

				[int]$cnt2 = -1
				ForEach ($Item in $PowerKeys)
				{
					$cnt2++
					$Value = $PowerValues[$cnt2]
					
					$ScriptInformation += @{ Data = "     Key: $Item"; Value = "Value: $Value"; }
				}
				$ScriptInformation += @{ Data = ""; Value = ""; }
			}
			If ($Text)
			{
				Line 2 "Server`t`t: " "$($XSHost.Name_Label)"
				Line 2 "Power On mode`t: " "Custom power-on script /etc/xapi.d/plugins/$($XSHost.power_on_mode)"
				Line 2 "Configuration options"

				[int]$cnt2 = -1
				ForEach ($Item in $PowerKeys)
				{
					$cnt2++
					$Value = $PowerValues[$cnt2]
					Line 3 "Key  : " $Item
					Line 3 "Value: " $Value
					Line 0 ""
				}
				Line 0 ""
			}
			If ($HTML)
			{
				If ($cnt -eq 0)
				{
					$columnHeaders = @("Server", ($htmlsilver -bor $htmlbold), "$($XSHost.Name_Label)", $htmlwhite)
				}
				Else
				{
					$rowdata += @(, ("Server", ($htmlsilver -bor $htmlbold), "$($XSHost.Name_Label)", $htmlwhite))
				}
				$rowdata += @(, ("Power On mode", ($htmlsilver -bor $htmlbold), "Custom power-on script /etc/xapi.d/plugins/$($XSHost.power_on_mode)", $htmlwhite))
				$rowdata += @(, ("Configuration options", ($htmlsilver -bor $htmlbold), "", $htmlwhite))

				[int]$cnt2 = -1
				ForEach ($Item in $PowerKeys)
				{
					$cnt2++
					$Value = $PowerValues[$cnt2]
				
					$rowdata += @(, ("     Key: $Item", ($htmlsilver -bor $htmlbold), "Value: $Value", $htmlwhite))
				}
				$rowdata += @(, ("", ($htmlsilver -bor $htmlbold), "", $htmlwhite))
			}
		}	
	}

	If ($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 300;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		#Nothing
	}
	If ($HTML)
	{
		$msg = ""
		$columnWidths = @("150", "300")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputPoolLivePatching
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool Live Patching"

	If ($Script:XSPool.live_patching_disabled)
	{
		$LivePatching = "Disabled"
	}
	Else
	{
		$LivePatching = "Enabled"
	}

	If ($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Live Patching"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Live patching"; Value = $LivePatching; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 1 "Live Patching"
		Line 2 "Live patching: " $LivePatching
		Line 0 ""
	}
	If ($HTML)
	{
		WriteHTMLLine 2 0 "Live Patching"
		$rowdata = @()
		$columnHeaders = @("Live patching", ($htmlsilver -bor $htmlbold), $LivePatching, $htmlwhite)

		$msg = ""
		$columnWidths = @("150", "250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputPoolNetworkOptions
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool Network Options"

	If ($Script:XSPool.igmp_snooping_enabled)
	{
		$IGMPsnooping = "Enabled"
	}
	Else
	{
		$IGMPsnooping = "Disabled"
	}

	If ($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Network Options"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "IGMP snooping"; Value = $IGMPsnooping; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 1 "Network Options"
		Line 2 "IGMP snooping: " $IGMPsnooping
		Line 0 ""
	}
	If ($HTML)
	{
		WriteHTMLLine 2 0 "Network Options"
		$rowdata = @()
		$columnHeaders = @("IGMP snooping", ($htmlsilver -bor $htmlbold), $IGMPsnooping, $htmlwhite)

		$msg = ""
		$columnWidths = @("150", "250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputPoolClustering
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool Clustering"
	
	$results = Get-XenCluster -SessionOpaqueRef $Script:Session.opaque_ref -EA 0

	If (!$?)
	{
		Write-Warning "
		`n
		Unable to retrieve Clustering for Pool $($Pool.name_label)`
		"
		If ($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Clustering for Pool $($Pool.name_label)"
		}
		If ($Text)
		{
			Line 0 "Unable to retrieve Clustering for Pool $($Pool.name_label)"
		}
		If ($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Clustering for Pool $($Pool.name_label)"
		}
	}
	ElseIf ($? -and $Null -eq $results)
	{
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			WriteWordLine 2 0 "Clustering"
			$ScriptInformation += @{ Data = "Clustering"; Value = "Disabled"; }
		}
		If ($Text)
		{
			Line 1 "Clustering"
			Line 2 "Clustering: " "Disabled"
		}
		If ($HTML)
		{
			$rowdata = @()
			WriteHTMLLine 2 0 "Clustering"
			$columnHeaders = @("Clustering", ($htmlsilver -bor $htmlbold), "Disabled", $htmlwhite)
		}
	}
	Else
	{
		#From Citrix
		#for the cluster network 
		#get-xencluster |  Get-XenClusterProperty -XenProperty Network
		#since the network is available as a cluster property getter,  I would have expected it 
		#to also appear in the list when calling get-xencluster.
		#this may be an API bug.
		
		$ClusterNetwork = ($results | Get-XenClusterProperty -XenProperty Network)
		
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			WriteWordLine 2 0 "Clustering"
			$ScriptInformation += @{ Data = "Clustering"; Value = "Enabled"; }
			$ScriptInformation += @{ Data = "Network"; Value = $ClusterNetwork; }
		}
		If ($Text)
		{
			Line 1 "Clustering"
			Line 2 "Clustering: " "Enabled"
			Line 2 "Network   : " $ClusterNetwork
		}
		If ($HTML)
		{
			$rowdata = @()
			WriteHTMLLine 2 0 "Clustering"
			$columnHeaders = @("Clustering", ($htmlsilver -bor $htmlbold), "Enabled", $htmlwhite)
			$rowdata += @(, ("Network", ($htmlsilver -bor $htmlbold), $ClusterNetwork, $htmlwhite))
		}
	}
	
	If ($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 0 ""
	}
	If ($HTML)
	{
		$msg = ""
		$columnWidths = @("150", "250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputPoolMemory
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool Memory"
	Write-Verbose "$(Get-Date -Format G): `t Pool Storage Gathering data"

	$XSPoolMemories = @()
	foreach ($XSHost in $Script:XSHosts)
	{
		$XSHostName = $XSHost.name_label
		$XSHostMetrics = $XSHost.metrics | Get-XenHostMetrics
		$memTotal = Convert-SizeToString -size $XSHostMetrics.memory_total -Decimal 1
		$memFree = Convert-SizeToString -size $XSHostMetrics.memory_free -Decimal 1
		$memoryText = "$memFree RAM available ($memTotal)"
		$hostAllRunningVMs = @( $XSHost.resident_VMs | Get-XenVM | Sort-Object -Property name_label)
		$hostRunningVMs = @($hostAllRunningVMs | Where-Object { $_.is_control_domain -eq $false -and $_.power_state -like "running" })
		$dom0VM = $hostAllRunningVMs | Where-Object { $_.is_control_domain -eq $true }
		$vmText = @()
		$vmMemoryUsed = [Int64]0
		ForEach ($vm in $hostRunningVMs)
		{
			$vmText += '{0}: using  {1}' -f $vm.name_label, $(Convert-SizeToString -size $vm.memory_target  -Decimal 1)
			$vmMemoryUsed = $vmMemoryUsed + $vm.memory_target
		}
	
		$memXSNum = $($dom0VM.memory_target + $XSHost.memory_overhead)
		$memXS = Convert-SizeToString -size $memXSNum -Decimal 1
		$memXSUsedNum = ($dom0VM.memory_target + $XSHost.memory_overhead + $vmMemoryUsed)
		$memXSAvailableNum = ($XSHostMetrics.memory_total - $memXSUsedNum)
		$memXSUsed = Convert-SizeToString -size $memXSUsedNum -Decimal 1
		$memXSUsedPct = '{0}%' -f [Math]::Round($memXSUsedNum / ($XSHostMetrics.memory_total / 100))
		$memXSAvailable = Convert-SizeToString -size $memXSAvailableNum -Decimal 1
		$cdMemory = Convert-SizeToString -size $dom0VM.memory_target -Decimal 1


		$XSPoolMemories += "" | Select-Object -Property `
		@{Name = 'XSHostName'; Expression = { $XSHostName } },
		@{Name = 'XSHostRef'; Expression = { $XSHost.opaque_ref } },
		@{Name = 'Server'; Expression = { $memoryText } },
		@{Name = 'VMs'; Expression = { $hostRunningVMs.Count } },
		@{Name = 'VMTexts'; Expression = { $vmText } },
		@{Name = 'XenServerMemory'; Expression = { $memXS } },
		@{Name = 'ControlDomainMemory'; Expression = { $cdMemory } },
		@{Name = 'AvailableMemory'; Expression = { $memXSAvailable } },
		@{Name = 'TotalMaxMemory'; Expression = { "$($memXSUsed) ($memXSUsedPct of total memory)" } }
	}
	$Script:XSPoolMemories = $XSPoolMemories | Sort-Object -Property XSHostName

	if ($NoPoolMemory -eq $false)
	{
		Write-Verbose "$(Get-Date -Format G): `t Pool Memory writing output"
		If ($MSWord -or $PDF)
		{
			$Selection.InsertNewPage()
			WriteWordLine 2 0 "Memory"
		}
		If ($Text)
		{
			Line 0 ""
			Line 1 "Memory"
		}
		If ($HTML)
		{
			WriteHTMLLine 2 0 "Memory"
		}

		ForEach ($XSHostMemory in $Script:XSPoolMemories)
		{
			If ($MSWord -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
			
			}
			If ($Text)
			{
			}
			If ($HTML)
			{
				$rowdata = @()
			}
			
			If ($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "Host"; Value = "$($XSHostMemory.XSHostName)"; }
				$ScriptInformation += @{ Data = "Server"; Value = "$($XSHostMemory.Server)"; }
				$ScriptInformation += @{ Data = "VMs"; Value = "$($XSHostMemory.VMs)"; }
				$XSHostMemory.VMTexts | ForEach-Object { $ScriptInformation += @{ Data = ""; Value = "$($_)"; } }
				$ScriptInformation += @{ Data = "Citrix Hypervisor"; Value = "$($XSHostMemory.XenServerMemory)"; }
				$ScriptInformation += @{ Data = "Control domain memory"; Value = "$($XSHostMemory.ControlDomainMemory)"; }
				$ScriptInformation += @{ Data = "Available memory"; Value = "$($XSHostMemory.AvailableMemory)"; }
				$ScriptInformation += @{ Data = "Total max memory"; Value = "$($XSHostMemory.TotalMaxMemory)"; }
			}
			If ($Text)
			{
				Line 3 "Host`t`t`t: " "$($XSHostMemory.XSHostName)"
				Line 3 "Server`t`t`t: " "$($XSHostMemory.Server)"
				Line 3 "VMs`t`t`t: " "$($XSHostMemory.VMs)"
				$XSHostMemory.VMTexts | ForEach-Object { Line 6 "  $($_)" }
				Line 3 "Citrix Hypervisor`t: " "$($XSHostMemory.XenServerMemory)"
				Line 3 "Control domain memory`t: " "$($XSHostMemory.ControlDomainMemory)"
				Line 3 "Available memory`t: " "$($XSHostMemory.AvailableMemory)"
				Line 3 "Total max memory`t: " "$($XSHostMemory.TotalMaxMemory)"
			}
			If ($HTML)
			{
				$columnHeaders = @("Host", ($htmlsilver -bor $htmlbold), "$($XSHostMemory.XSHostName)", ($htmlsilver -bor $htmlbold))
				$rowdata +=  @(, ("Server", ($htmlsilver -bor $htmlbold), "$($XSHostMemory.Server)", $htmlwhite))
				$rowdata += @(, ("VMs", ($htmlsilver -bor $htmlbold), "$($XSHostMemory.VMs)", $htmlwhite))
				$XSHostMemory.VMTexts | ForEach-Object { $rowdata += @(, ("", ($htmlsilver -bor $htmlbold), "$($_)", $htmlwhite)) }
				$rowdata += @(, ("Citrix Hypervisor", ($htmlsilver -bor $htmlbold), "$($XSHostMemory.XenServerMemory)", $htmlwhite))
				$rowdata += @(, ("Control domain memory", ($htmlsilver -bor $htmlbold), "$($XSHostMemory.ControlDomainMemory)", $htmlwhite))
				$rowdata += @(, ("Available memory", ($htmlsilver -bor $htmlbold), "$($XSHostMemory.AvailableMemory)", $htmlwhite))
				$rowdata += @(, ("Total max memory", ($htmlsilver -bor $htmlbold), "$($XSHostMemory.TotalMaxMemory)", $htmlwhite))
			}
		
			If ($MSWord -or $PDF)
			{
				$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data, Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;
			
				## IB - Set the header row format
				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
			
				$Table.Columns.Item(1).Width = 150;
				$Table.Columns.Item(2).Width = 200;
			
				$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)
			
				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If ($Text)
			{
				Line 0 ""
			}
			If ($HTML)
			{
				$msg = ""
				$columnWidths = @("150", "200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): `t Pool Memory skipped"
	}
}

Function OutputPoolStorage
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool Storage"
	
	Write-Verbose "$(Get-Date -Format G): `t Pool Storage Gathering data"
	$pbds = Get-XenPBD | Where-Object { $_.host.opaque_ref -in $Script:XSHosts.opaque_ref }

	$XSPoolStorages = @()
	ForEach ($item in $pbds)
	{
		$XSHost = $Script:XSHosts | Where-Object { $_.opaque_ref -like $item.host.opaque_ref }
		$XSHostName = $XSHost.name_label
		$sr = $item.SR | Get-XenSR -EA 0
		If ([String]::IsNullOrEmpty($($sr.name_description))) 
		{
			$description = '{0} on {1}' -f $sr.name_label, $XSHostName
		}
		Else
		{
			$description = $($sr.name_description)
		}
		If ($sr.shared -like $true)
		{
			$shared = "yes"
		}
		Else
		{
			$shared = "no"
		}
		$virtualAlloc = Convert-SizeToString -Size $sr.virtual_allocation -Decimal 1
		$size = Convert-SizeToString -Size $sr.physical_size -Decimal 1
		$used = Convert-SizeToString -Size $sr.physical_utilisation -Decimal 1
		If ($sr.physical_utilisation -le 0 -or $sr.physical_size -le 0)
		{
			$usage = '0% (0 B)'
		}
		Else
		{
			$usage = '{0}% ({1} used)' -f [math]::Round($($sr.physical_utilisation / ($sr.physical_size / 100))), $used
		}
		$XSPoolStorages += $sr | Select-Object -Property `
		@{Name = 'XSHostName'; Expression = { $XSHostName } },
		@{Name = 'XSHostRef'; Expression = { $XSHost.opaque_ref } },
		@{Name = 'Name'; Expression = { $_.name_label } },
		@{Name = 'Description'; Expression = { $description } },
		@{Name = 'Type'; Expression = { $_.type } },
		@{Name = 'Shared'; Expression = { $shared } },
		@{Name = 'Usage'; Expression = { $usage } },
		@{Name = 'Size'; Expression = { $size } },
		@{Name = 'VirtualAllocation'; Expression = { $virtualAlloc } }
	}
	$XSPoolStorages = @($XSPoolStorages | Sort-Object -Property XSHostName, Name)
	$storageCount = $XSPoolStorages.Count
	$Script:XSPoolStorages = $XSPoolStorages

	if ($NoPoolStorage -eq $false) 
	{
		Write-Verbose "$(Get-Date -Format G): `t Pool Storage writing output"
		If ($MSWord -or $PDF)
		{
			$Selection.InsertNewPage()
			WriteWordLine 2 0 "Storage"
		}
		If ($Text)
		{
			Line 0 ""
			Line 1 "Storage"
		}
		If ($HTML)
		{
			WriteHTMLLine 2 0 "Storage"
		}

		If ($storageCount -lt 1)
		{
			If ($MSWord -or $PDF)
			{
				WriteWordLine 0 1 "There is no storage configured for Host $XSHostName"
			}
			If ($Text)
			{
				Line 3 "There is no storage configured for Host $XSHostName"
				Line 0 ""
			}
			If ($HTML)
			{
				WriteHTMLLine 0 1 "There is no storage configured for Host $XSHostName"
			}
		}
		Else
		{
			If ($MSWord -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Number of storages"; Value = "$storageCount"; }
			}
			If ($Text)
			{
				Line 3 "Number of storages: " "$storageCount"
				Line 0 ""
			}
			If ($HTML)
			{
				$columnHeaders = @("Number of storages", ($htmlsilver -bor $htmlbold), "$storageCount", $htmlwhite)
				$rowdata = @()
			}

			ForEach ($Item in $XSPoolStorages)
			{
				If ($MSWord -or $PDF)
				{
					$ScriptInformation += @{ Data = "Name"; Value = $($item.Name); }
					$ScriptInformation += @{ Data = "     Description"; Value = $($item.Description); }
					$ScriptInformation += @{ Data = "     Type"; Value = $($item.Type); }
					$ScriptInformation += @{ Data = "     Shared"; Value = $($item.Shared); }
					$ScriptInformation += @{ Data = "     Usage"; Value = $($item.Usage); }
					$ScriptInformation += @{ Data = "     Size"; Value = $($item.Size); }
					$ScriptInformation += @{ Data = "     Virtual allocation"; Value = $($item.VirtualAllocation); }
				}
				If ($Text)
				{
					Line 3 "Name: " $($item.Name)
					Line 4 "Description`t`t: " $($item.Description)
					Line 4 "Type`t`t`t: " $($item.Type)
					Line 4 "Shared`t`t`t: " $($item.Shared)
					Line 4 "Usage`t`t`t: " $($item.Usage)
					Line 4 "Size`t`t`t: " $($item.Size)
					Line 4 "Virtual allocation`t: " $($item.VirtualAllocation)
					Line 0 ""
				}
				If ($HTML)
				{
					$rowdata += @(, ("Name", ($htmlsilver -bor $htmlbold), $($item.Name), ($htmlsilver -bor $htmlbold)))
					$rowdata += @(, ("     Description", ($htmlsilver -bor $htmlbold), $($item.Description), $htmlwhite))
					$rowdata += @(, ("     Type", ($htmlsilver -bor $htmlbold), $($item.Type), $htmlwhite))
					$rowdata += @(, ("     Shared", ($htmlsilver -bor $htmlbold), $($item.Shared), $htmlwhite))
					$rowdata += @(, ("     Usage", ($htmlsilver -bor $htmlbold), $($item.Usage), $htmlwhite))
					$rowdata += @(, ("     Size", ($htmlsilver -bor $htmlbold), $($item.Size), $htmlwhite))
					$rowdata += @(, ("     Virtual allocation", ($htmlsilver -bor $htmlbold), $($item.VirtualAllocation), $htmlwhite))
				}
			}
		
			If ($MSWord -or $PDF)
			{
				$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data, Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

				## IB - Set the header row format
				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 150;
				$Table.Columns.Item(2).Width = 350;

				$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If ($Text)
			{
				Line 0 ""
			}
			If ($HTML)
			{
				$msg = ""
				$columnWidths = @("150", "350")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): `t Pool Networking skipped"
	}
}

Function OutputPoolNetworking
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool Networking"
	Write-Verbose "$(Get-Date -Format G): `t Gathering Networking data"
	$networks = @(Get-XenNetwork -EA 0 | Where-Object { $_.other_config["is_host_internal_management_network"] -notlike $true })
	$XSNetworks = @()
	$nrNetworks = $networks.Count
	If ($nrNetworks -ge 1)
	{
		ForEach ($Item in $networks)
		{
			ForEach ($XSHost in $Script:XSHosts)
               {
				$pif = $Item.PIFs | Get-XenPIF -EA 0 | Where-Object { $XSHost.opaque_ref -in $_.host }
				if ([String]::IsNullOrEmpty($pif))
				{
					$nic = ""
					$vlan = ""
					$autoAssign = "No"
					$linkStatus = "<None>"
					$mac = "-"
				}
				else
				{
					$nic = $pif.device.Replace("eth", "NIC ")
					
					If ([String]::IsNullOrEmpty($($pif.VLAN)) -or ($pif.VLAN -lt 0))
					{
						$vlan = "-"
						$mac = $pif.MAC
					}
					Else
					{
						$vlan = "$($pif.VLAN)"
						$mac = "-"
					}
					if ($Item.other_config["automatic"] -like $true)
					{
						$autoAssign = "Yes"
					}
					else
					{
						$autoAssign = "No"
					}
					
					$pifMetrics = $pif.metrics | Get-XenPIFMetrics
					if ($pifMetrics.carrier -like $true)
					{
						$linkStatus = "Connected"
					}
					else
					{
						$linkStatus = "Disconnected"
					}
				}
				if (($XSHost.opaque_ref -eq $Script:XSPool.master.opaque_ref))
				{
					$hostIsPoolMaster = $true
				}
				else
				{
					$hostIsPoolMaster = $false
				}
				$XSNetworks += $Item | Select-Object -Property `
				@{Name = 'XSHostname'; Expression = { $XSHost.name_label } },
				@{Name = 'XSHostref'; Expression = { $XSHost.opaque_ref } },
				@{Name = 'XSHostPoolMaster'; Expression = { $hostIsPoolMaster } },
				@{Name = 'Name'; Expression = { $item.name_label.Replace("Pool-wide network associated with eth", "Network ") } },
				@{Name = 'Description'; Expression = { $_.name_description } },
				@{Name = 'NIC'; Expression = { $nic } },
				@{Name = 'VLAN'; Expression = { $vlan } },
				@{Name = 'Auto'; Expression = { $autoAssign } },
				@{Name = 'LinkStatus'; Expression = { $linkStatus } },
				@{Name = 'MAC'; Expression = { $mac } },
				@{Name = 'MTU'; Expression = { $item.MTU } },
				@{Name = 'SRIOV'; Expression = { "" } }
			}
		}
	}
	$XSNetworks = @($XSNetworks | Sort-Object -Property XSHostname, Name)
	$Script:XSPoolNetworks = $XSNetworks
	#Choose to use Pool Master data as original XenCenter pool data is more or less "random"
	if ($NoPoolNetworking -eq $false) 
	{
		Write-Verbose "$(Get-Date -Format G): `t Pool Networking writing output"
		$XSNetworks = $XSNetworks | Where-Object { $_.XSHostPoolMaster -eq $true }
		$nrNetworking = $XSNetworks.Count
		If ($MSWord -or $PDF)
		{
			$Selection.InsertNewPage()
			WriteWordLine 2 0 "Networking"
		}
		If ($Text)
		{
			Line 0 ""
			Line 1 "Networking"
		}
		If ($HTML)
		{
			WriteHTMLLine 2 0 "Networking"
		}

		If ($nrNetworking -lt 1)
		{
			If ($MSWord -or $PDF)
			{
				WriteWordLine 0 1 "There are no networks configured for Host $XSHostName"
			}
			If ($Text)
			{
				Line 3 "There are no Network networks configured for Host $XSHostName"
				Line 0 ""
			}
			If ($HTML)
			{
				WriteHTMLLine 0 1 "There are no networks configured for Host $XSHostName"
			}
		}
		Else
		{
			If ($MSWord -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Number of networks"; Value = "$nrNetworking"; }
			}
			If ($Text)
			{
				Line 3 "Number of networks: " "$nrNetworking"
				Line 0 ""
			}
			If ($HTML)
			{
				$columnHeaders = @("Number of networks", ($htmlsilver -bor $htmlbold), "$nrNetworking", $htmlwhite)
				$rowdata = @()
			}

			ForEach ($Item in $XSNetworks)
			{
				If ($MSWord -or $PDF)
				{
					$ScriptInformation += @{ Data = "Name"; Value = $($item.Name); }
					$ScriptInformation += @{ Data = "     Description"; Value = $($item.Description); }
					$ScriptInformation += @{ Data = "     NIC"; Value = $($item.NIC); }
					$ScriptInformation += @{ Data = "     VLAN"; Value = $($item.VLAN); }
					$ScriptInformation += @{ Data = "     Auto"; Value = $($item.Auto); }
					$ScriptInformation += @{ Data = "     Link Status"; Value = $($item.LinkStatus); }
					$ScriptInformation += @{ Data = "     MAC"; Value = $($item.MAC); }
					$ScriptInformation += @{ Data = "     MTU"; Value = $($item.MTU); }
					$ScriptInformation += @{ Data = "     SR-IOV"; Value = $($item.SRIOV); }
				}
				If ($Text)
				{
					Line 3 "Name: " $($item.Name)
					Line 4 "Description`t: " $($item.Description)
					Line 4 "NIC`t`t: " $($item.NIC)
					Line 4 "VLAN`t`t: " $($item.VLAN)
					Line 4 "Auto`t`t: " $($item.Auto)
					Line 4 "Link Status`t: " $($item.LinkStatus)
					Line 4 "MAC`t`t: " $($item.MAC)
					Line 4 "MTU`t`t: " $($item.MTU)
					Line 4 "SR-IOV`t`t: " $($item.SRIOV)
					Line 0 ""
				}
				If ($HTML)
				{
					$rowdata += @(, ("Name", ($htmlsilver -bor $htmlbold), $($item.Name), ($htmlsilver -bor $htmlbold)))
					$rowdata += @(, ("     Description", ($htmlsilver -bor $htmlbold), $($item.Description), $htmlwhite))
					$rowdata += @(, ("     NIC", ($htmlsilver -bor $htmlbold), $($item.NIC), $htmlwhite))
					$rowdata += @(, ("     VLAN", ($htmlsilver -bor $htmlbold), $($item.VLAN), $htmlwhite))
					$rowdata += @(, ("     Auto", ($htmlsilver -bor $htmlbold), $($item.Auto), $htmlwhite))
					$rowdata += @(, ("     Link Status", ($htmlsilver -bor $htmlbold), $($item.LinkStatus), $htmlwhite))
					$rowdata += @(, ("     MAC", ($htmlsilver -bor $htmlbold), $($item.MAC), $htmlwhite))
					$rowdata += @(, ("     MTU", ($htmlsilver -bor $htmlbold), $($item.MTU), $htmlwhite))
					$rowdata += @(, ("     SRIOV", ($htmlsilver -bor $htmlbold), $($item.SRIOV), $htmlwhite))
				}
			}
		
			If ($MSWord -or $PDF)
			{
				$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data, Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

				## IB - Set the header row format
				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 150;
				$Table.Columns.Item(2).Width = 175;

				$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If ($Text)
			{
				Line 0 ""
			}
			If ($HTML)
			{
				$msg = ""
				$columnWidths = @("150", "200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
		}
	} else {
		Write-Verbose "$(Get-Date -Format G): `t Pool Networking skipped"
	}
}

Function OutputPoolGPU
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool GPU"
	If ($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 2 0 "GPU"
	}
	If ($Text)
	{
		Line 0 ""
		Line 1 "GPU"
	}
	If ($HTML)
	{
		WriteHTMLLine 2 0 "GPU"
	}
	
}

Function OutputPoolHA
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool HA"
	If ($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 2 0 "HA"
	}
	If ($Text)
	{
		Line 0 ""
		Line 1 "HA"
	}
	If ($HTML)
	{
		WriteHTMLLine 2 0 "HA"
	}
	
	If ($Script:XSPool.ha_enabled)
	{
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Configuration"; Value = ""; }
			$ScriptInformation += @{ Data = "     Pool HA enabled"; Value = "Yes"; }
			$ScriptInformation += @{ Data = "     Configured failure capacity"; Value = $Script:XSPool.ha_host_failures_to_tolerate.ToString(); }
			$ScriptInformation += @{ Data = "     Current failure capacity"; Value = $Script:XSPool.ha_plan_exists_for.ToString(); }

			$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data, Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If ($Text)
		{
			Line 2 "Configuration"
			Line 3 "Pool HA enabled            : " "Yes"
			Line 3 "Configured failure capacity: " $Script:XSPool.ha_host_failures_to_tolerate.ToString()
			Line 3 "Current failure capacity   : " $Script:XSPool.ha_plan_exists_for.ToString()
			Line 0 ""
		}
		If ($HTML)
		{
			$rowdata = @()
			$columnHeaders = @("Configuration", ($htmlsilver -bor $htmlbold), "", $htmlwhite)
			$rowdata += @(, ("     Pool HA enabled", ($htmlsilver -bor $htmlbold), "Yes", $htmlwhite))
			$rowdata += @(, ("     Configured failure capacity", ($htmlsilver -bor $htmlbold), $Script:XSPool.ha_host_failures_to_tolerate.ToString(), $htmlwhite))
			$rowdata += @(, ("     Current failure capacity", ($htmlsilver -bor $htmlbold), $Script:XSPool.ha_plan_exists_for.ToString(), $htmlwhite))

			$msg = ""
			$columnWidths = @("150", "250")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
	Else
	{
		$HAPool = "`'$($Script:XSPool.name_label)`'"
		$HATxt = "HA is not currently enabled for pool"
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = $HATxt; Value = $HAPool; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data, Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 200;

			$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If ($Text)
		{
			Line 2 "$($HATxt): " $HAPool
			Line 0 ""
		}
		If ($HTML)
		{
			$rowdata = @()
			$columnHeaders = @($HATxt, ($htmlsilver -bor $htmlbold), $HAPool, $htmlwhite)

			$msg = ""
			$columnWidths = @("200", "200")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
}

Function OutputPoolWLB
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool WLB"
	If ($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 2 0 "WLB"
	}
	If ($Text)
	{
		Line 0 ""
		Line 1 "WLB"
	}
	If ($HTML)
	{
		WriteHTMLLine 2 0 "WLB"
	}
	
	If ($Script:XSPool.wlb_enabled)
	{
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Server Address"; Value = ""; }
			$ScriptInformation += @{ Data = "     Address"; Value = "$($Script:XSPool.wlb_url)"; }
			$ScriptInformation += @{ Data = "WLB Server Credentials"; Value = ""; }
			$ScriptInformation += @{ Data = "     Username"; Value = "$($Script:XSPool.wlb_username)"; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data, Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If ($Text)
		{
			Line 2 "Server Address"
			Line 3 "Address : " "$($Script:XSPool.wlb_url)"
			Line 2 "WLB Server Credentials"
			Line 3 "Username: " "$($Script:XSPool.wlb_username)"
			Line 0 ""
		}
		If ($HTML)
		{
			$rowdata = @()
			$columnHeaders = @("Server Address", ($htmlsilver -bor $htmlbold), "", $htmlwhite)
			$rowdata += @(, ("     Address", ($htmlsilver -bor $htmlbold), "$($Script:XSPool.wlb_url)", $htmlwhite))
			$rowdata += @(, ("WLB Server Credentials", ($htmlsilver -bor $htmlbold), "", $htmlwhite))
			$rowdata += @(, ("     Username", ($htmlsilver -bor $htmlbold), "$($Script:XSPool.wlb_username)", $htmlwhite))

			$msg = ""
			$columnWidths = @("150", "250")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
	Else
	{
		$WLBTxt = "Pool `($($Script:XSPool.name_label)`) is not currently connected to a Workload Balancing (WLB) server"
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = $WLBTxt; Value = ""; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data, Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 400;
			$Table.Columns.Item(2).Width = 20;

			$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If ($Text)
		{
			Line 2 $WLBTxt
			Line 0 ""
		}
		If ($HTML)
		{
			$rowdata = @()
			$columnHeaders = @($WLBTxt, ($htmlsilver -bor $htmlbold), "", $htmlwhite)

			$msg = ""
			$columnWidths = @("450", "10")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
}

Function OutputPoolUsers
{
	Write-Verbose "$(Get-Date -Format G): `tOutput Pool Users"
	If ($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 2 0 "Users"
	}
	If ($Text)
	{
		Line 0 ""
		Line 1 "Users"
	}
	If ($HTML)
	{
		WriteHTMLLine 2 0 "Users"
	}

	#domain info
	If ($Script:PoolMasterInfo.external_auth_type -eq "AD")
	{
		$DomainTxt = "Pool `'$($Script:XSPool.name_label)`' belongs to the domain"
		$DomainName = "`'$($Script:PoolMasterInfo.external_auth_service_name)`'"
	}
	Else
	{
		$DomainTxt = "AD is not currently configured for pool"
		$DomainName = "`'$($Script:XSPool.name_label)`'"
	}
	
	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Active Directory Users"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = $DomainTxt; Value = $DomainName; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 2 "Active Directory Users"
		Line 3 "$($DomainTxt): " $DomainName
		Line 0 ""
	}
	If ($HTML)
	{
		WriteHTMLLine 2 0 "Active Directory Users"
		$rowdata = @()
		$columnHeaders = @($DomainTxt, ($htmlsilver -bor $htmlbold), $DomainName, $htmlwhite)

		$msg = ""
		$columnWidths = @("225", "250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
	
	#user info
	#code by JB
	$subjects = Get-XenSubject -EA 0
	$users = @()
	ForEach ($item in $subjects)
	{
		$roles = $item.roles | Get-XenRole
		if ($item.IsGroup)
		{
			$type = "Group"
		}
		else
		{
			$type = "User"
		}
		$users += $item | Select-Object -Property `
		@{Name = 'Type'; Expression = { $type } },
		@{Name = 'Subject'; Expression = { $_.SubjectName } },
		@{Name = 'Name'; Expression = { $_.DisplayName } },
		@{Name = 'Roles'; Expression = { $roles.name_label -join ", " } },
		@{Name = 'AccountDisabled'; Expression = { $item.other_config["subject-account-disabled"] } },
		@{Name = 'AccountExpired'; Expression = { $item.other_config["subject-account-expired"] } },
		@{Name = 'AccountLocked'; Expression = { $item.other_config["subject-account-locked"] } },
		@{Name = 'PasswordExpired'; Expression = { $item.other_config["subject-password-expired"] } }
	}
	$users = @($users | Sort-Object -Property Type, Name)
	
	#role names
	<#
		Pool Admin     - pool-admin
		Pool Operator  - pool-operator
		VM Power Admin - vm-power-admin
		VM Admin       - vm-admin
		VM Operator    - vm-operator
		Read Only      - read-only
	#>
	
	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "User and Groups with Access"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Type"; Value = "User"; }
		$ScriptInformation += @{ Data = "Subject"; Value = "Local root account"; }
		$ScriptInformation += @{ Data = "Roles"; Value = "(Always granted access)"; }
		$ScriptInformation += @{ Data = ""; Value = ""; }
	}
	If ($Text)
	{
		Line 2 "User and Groups with Access"
		Line 3 "Type            : User"
		Line 3 "Subject         : Local root account"
		Line 3 "Roles           : (Always granted access)"
		Line 0 ""
	}
	If ($HTML)
	{
		WriteHTMLLine 2 0 "User and Groups with Access"
		$rowdata = @()
		$columnHeaders = @("Type", ($htmlsilver -bor $htmlbold), "User", $htmlwhite)
		$rowdata += @(, ("Subject", ($htmlsilver -bor $htmlbold), "Local root account", $htmlwhite))
		$rowdata += @(, ("Roles", ($htmlsilver -bor $htmlbold), "(Always granted access)", $htmlwhite))
		$rowdata += @(, ("", ($htmlsilver -bor $htmlbold), "", $htmlwhite))
	}
	
	ForEach ($User in $Users)
	{
		Switch ($User.Roles)
		{
			"pool-admin" { $UserRole = "Pool Admin"; Break }
			"pool-operator" { $UserRole = "Pool Operator"; Break }
			"vm-power-admin"	{ $UserRole = "VM Power Admin"; Break }
			"vm-admin" { $UserRole = "VM Admin"; Break }
			"vm-operator" { $UserRole = "VM Operator"; Break }
			"read-only" { $UserRole = "Read Only"; Break }
			Default { $UserRole = "Unable to determine the user role: $($User.Roles)"; Break }
		}
		
		If ($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Type"; Value = $User.Type; }
			$ScriptInformation += @{ Data = "Subject"; Value = $User.Subject; }
			$ScriptInformation += @{ Data = "Name"; Value = $User.Name; }
			$ScriptInformation += @{ Data = "Roles"; Value = $UserRole; }
			$ScriptInformation += @{ Data = "Account disabled"; Value = $User.AccountDisabled; }
			$ScriptInformation += @{ Data = "Account expired"; Value = $User.AccountExpired; }
			$ScriptInformation += @{ Data = "Account locked"; Value = $User.AccountLocked; }
			$ScriptInformation += @{ Data = "Password expired"; Value = $User.PasswordExpired; }
			$ScriptInformation += @{ Data = ""; Value = ""; }
		}
		If ($Text)
		{
			Line 3 "Type            : " $User.Type
			Line 3 "Subject         : " $User.Subject
			Line 3 "Name            : " $User.Name
			Line 3 "Roles           : " $UserRole
			Line 3 "Account disabled: " $User.AccountDisabled
			Line 3 "Account expired : " $User.AccountExpired
			Line 3 "Account locked  : " $User.AccountLocked
			Line 3 "Password expired: " $User.PasswordExpired
			Line 0 ""
		}
		If ($HTML)
		{
			$rowdata += @(, ("Type", ($htmlsilver -bor $htmlbold), $User.Type, $htmlwhite))
			$rowdata += @(, ("Subject", ($htmlsilver -bor $htmlbold), $User.Subject, $htmlwhite))
			$rowdata += @(, ("Names", ($htmlsilver -bor $htmlbold), $User.Name, $htmlwhite))
			$rowdata += @(, ("Roles", ($htmlsilver -bor $htmlbold), $UserRole, $htmlwhite))
			$rowdata += @(, ("Account disabled", ($htmlsilver -bor $htmlbold), $User.AccountDisabled, $htmlwhite))
			$rowdata += @(, ("Account expired", ($htmlsilver -bor $htmlbold), $User.AccountExpired, $htmlwhite))
			$rowdata += @(, ("Account locked", ($htmlsilver -bor $htmlbold), $User.AccountLocked, $htmlwhite))
			$rowdata += @(, ("Password expired", ($htmlsilver -bor $htmlbold), $User.PasswordExpired, $htmlwhite))
			$rowdata += @(, ("", ($htmlsilver -bor $htmlbold), "", $htmlwhite))
		}
	}
	
	If ($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 0 ""
	}
	If ($HTML)
	{
		$msg = ""
		$columnWidths = @("150", "250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}

}
#endregion

#region hosts
Function ProcessHosts
{
	Write-Verbose "$(Get-Date -Format G): Process Hosts"
	If ($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Hosts"
	}
	If ($Text)
	{
		Line 0 ""
		Line 0 "Hosts"
		Line 0 ""
	}
	If ($HTML)
	{
		WriteHTMLLine 1 0 "Hosts"
	}
	
	$HostFirst = $True
	ForEach ($XSHost in $Script:XSHosts)
	{
		Write-Verbose "$(Get-Date -Format G): `tOutput Host $($XSHost.name_label)"
		OutputHostGeneralOverview $XSHost $HostFirst
		OutputHostLicense $XSHost
		OutputHostVersion $XSHost
		OutputHostUpdates $XSHost
		OutputHostManagement $XSHost
		OutputHostMemoryOverview $XSHost
		OutputHostCPUs $XSHost
		OutputHostGeneral $XSHost
		OutputHostCustomFields $XSHost
		OutputHostAlerts $XSHost
		OutputHostMultipathing $XSHost
		OutputHostLogDestination $XSHost
		OutputHostPowerOn $XSHost
		OutputHostGPUProperties $XSHost
		OutputHostMemory $XSHost
		OutputHostStorage $XSHost
		OutputHostNetworking $XSHost
		OutputHostNICs $XSHost
		OutputHostGPU $XSHost
		$HostFirst = $False
	}
}

Function OutputHostGeneralOverview
{
	Param([object]$XSHost, [bool]$HostFirst)
	
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host General Overview"
	
	#is this host the pool master?
	If ($XSHost.opaque_ref -eq $Script:XSPool.master.opaque_ref)
	{
		$IAmThePoolMaster = "Yes"
	}
	Else
	{
		$IAmThePoolMaster = "No"
	}

	[array]$xtags = @()
	ForEach ($tag in $XSHost.tags)
	{
		$xtags += $tag
	}
	If ($xtags.count -gt 0)
	{
		[array]$xtags = $xtags | Sort-Object
	}
	Else
	{
		[array]$xtags = @("<None>")
	}
	
	$LogLocation = ""
	If ($XSHost.Logging.Count -eq 0)
	{
		$LogLocation = "Local"
	}
	Else
	{
		$LogLocation = "Local and Remote ($($XSHost.logging.syslog_destination))"
	}
	
	#Thanks to Michael B. Smith for the help in the following calculations
	[int64]$UnixTime = $XSHost.other_config.boot_time
	$ServerTime = $XSHost | Get-XenHostProperty -XenProperty servertime
	$Origin = Get-Date -Year 1970 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
	
	$ServerBootTime = $Origin.AddSeconds($UnixTime)
	$ServerUptime = $ServerTime - $ServerBootTime
	$ServerUptimeString = [string]::format("{0} days,   {1} hours,   {2} minutes",
		$ServerUptime.Days,
		$ServerUptime.Hours,
		$ServerUptime.Minutes)

	$AgentStartTime = $Origin.AddSeconds($UnixTime)
	$AgentUptime = $ServerTime - $AgentStartTime
	$AgentUptimeString = [string]::format("{0} days,   {1} hours,   {2} minutes",
		$AgentUptime.Days,
		$AgentUptime.Hours,
		$AgentUptime.Minutes)

	If ([String]::IsNullOrEmpty($($XSHost.Other_Config["folder"])))
	{
		$folderName = "None"
	}
	Else
	{
		$folderName = $XSHost.Other_Config["folder"]
	}

	If ($XSHost.name_description -eq "Default install")
	{
		$XSHostDescription = "Default install of Citrix Hypervisor"
	}
	ELse
	{
		$XSHostDescription = $XSHost.name_description
	}
	
	If ($MSWord -or $PDF)
	{
		If ($HostFirst -eq $False)
		{
			#Put the 2nd Host on, on a new page
			$Selection.InsertNewPage()
		}
		
		WriteWordLine 2 0 "Host: $($XSHost.name_label)"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Name"; Value = $XSHost.name_label; }
		$ScriptInformation += @{ Data = "Description"; Value = $XSHostDescription; }
		$ScriptInformation += @{ Data = "Tags"; Value = $($xtags -join ", "); }
		$ScriptInformation += @{ Data = "Folder"; Value = $folderName; }
		$ScriptInformation += @{ Data = "Pool master"; Value = $IAmThePoolMaster; }
		$ScriptInformation += @{ Data = "Enabled"; Value = $XSHost.enabled.ToString(); }
		$ScriptInformation += @{ Data = "iSCSI IQN"; Value = $XSHost.iscsi_iqn; }
		$ScriptInformation += @{ Data = "Log destination"; Value = $LogLocation; }
		$ScriptInformation += @{ Data = "Server uptime"; Value = $ServerUptimeString; }
		$ScriptInformation += @{ Data = "Toolstack uptime"; Value = $AgentUptimeString; }
		$ScriptInformation += @{ Data = "Domain"; Value = $XSHost.external_auth_service_name; }
		$ScriptInformation += @{ Data = "UUID"; Value = $XSHost.uuid; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 1 "Name: " "$($XSHost.name_label)"
		Line 2 "Description`t`t: " $XSHostDescription
		Line 2 "Tags`t`t`t: " "$($xtags -join ", ")"
		Line 2 "Folder`t`t`t: " $folderName
		Line 2 "Pool master`t`t: " $IAmThePoolMaster
		Line 2 "Enabled`t`t`t: " $XSHost.enabled.ToString()
		Line 2 "iSCSI IQN`t`t: " $XSHost.iscsi_iqn
		Line 2 "Log destination`t`t: " $LogLocation
		Line 2 "Server uptime`t`t: " $ServerUptimeString
		Line 2 "Toolstack uptime`t: " $AgentUptimeString
		Line 2 "Domain`t`t`t: " $XSHost.external_auth_service_name
		Line 2 "UUID`t`t`t: " $XSHost.uuid
		Line 0 ""
	}
	If ($HTML)
	{
		#for HTML output, remove the < and > from <None> xtags and foldername if they are there
		$xtags = $xtags.Trim("<", ">")
		$folderName = $folderName.Trim("<", ">")
		WriteHTMLLine 2 0 "Host: $($XSHost.name_label)"
		$rowdata = @()
		$columnHeaders = @("Name", ($htmlsilver -bor $htmlbold), $XSHost.name_label, $htmlwhite)
		$rowdata += @(, ('Description', ($htmlsilver -bor $htmlbold), $XSHostDescription, $htmlwhite))
		$rowdata += @(, ('Tags', ($htmlsilver -bor $htmlbold), "$($xtags -join ", ")", $htmlwhite))
		$rowdata += @(, ('Folder', ($htmlsilver -bor $htmlbold), "$folderName", $htmlwhite))
		$rowdata += @(, ('Pool master', ($htmlsilver -bor $htmlbold), $IAmThePoolMaster, $htmlwhite))
		$rowdata += @(, ("Enabled", ($htmlsilver -bor $htmlbold), $XSHost.enabled.ToString(), $htmlwhite))
		$rowdata += @(, ("iSCSI IQN", ($htmlsilver -bor $htmlbold), $XSHost.iscsi_iqn, $htmlwhite))
		$rowdata += @(, ("Log destination", ($htmlsilver -bor $htmlbold), $LogLocation, $htmlwhite))
		$rowdata += @(, ("Server uptime", ($htmlsilver -bor $htmlbold), $ServerUptimeString, $htmlwhite))
		$rowdata += @(, ("Toolstack uptime", ($htmlsilver -bor $htmlbold), $AgentUptimeString, $htmlwhite))
		$rowdata += @(, ("Domain", ($htmlsilver -bor $htmlbold), "$($XSHost.external_auth_service_name)", $htmlwhite))
		$rowdata += @(, ("UUID", ($htmlsilver -bor $htmlbold), $XSHost.uuid, $htmlwhite))

		$msg = ""
		$columnWidths = @("150", "250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLIne 0 0 ""
	}
}

Function OutputHostLicense
{
	Param([object]$XSHost)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host License Details"

	$licenseExpiryDate = [datetime]::parseexact($XSHost.license_params["expiry"], "yyyyMMddTHH:mm:ssZ", $([cultureinfo]::InvariantCulture))
	If (($XSHost.license_server["address"] -like "localhost") -or ([String]::IsNullOrEmpty($($XSHost.license_params["license_type"]))) -or ($licenseExpiryDate -lt (Get-Date).AddDays(30)))
	{
		$licenseStatus = "Unlicensed"
	}
	Else
	{
		$licenseStatus = "Licensed"
	} 

	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "License Details"
	}
	If ($Text)
	{
		Line 2 "License Details"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "License Details"
	}

	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
	}
	If ($Text)
	{
	}
	If ($HTML)
	{
		$rowdata = @()
	}

	If ($MSWord -or $PDF)
	{
		$ScriptInformation += @{ Data = "Status"; Value = $licenseStatus; }
		$ScriptInformation += @{ Data = "License"; Value = "$($XSHost.license_params["sku_marketing_name"])"; }
		$ScriptInformation += @{ Data = "Number of Sockets"; Value = $($XSHost.license_params["sockets"]); }
		$ScriptInformation += @{ Data = "License Server Address"; Value = "$($XSHost.license_server["address"])"; }
		$ScriptInformation += @{ Data = "License Server Port"; Value = "$($XSHost.license_server["port"])"; }
	}
	If ($Text)
	{
		Line 3 "Status`t`t`t: " "$licenseStatus"
		Line 3 "License`t`t`t: " "$($XSHost.license_params["sku_marketing_name"])"
		Line 3 "Number of Sockets`t: " $($XSHost.license_params["sockets"])
		Line 3 "License Server Address`t: " "$($XSHost.license_server["address"])"
		Line 3 "License Server Port`t: " "$($XSHost.license_server["port"])"
	}
	If ($HTML)
	{
		$columnHeaders = @("Status", ($htmlsilver -bor $htmlbold), $licenseStatus, $htmlwhite)
		$rowdata += @(, ("License", ($htmlsilver -bor $htmlbold), $($XSHost.license_params["sku_marketing_name"]), $htmlwhite))
		$rowdata += @(, ("Number of Sockets", ($htmlsilver -bor $htmlbold), "$($XSHost.license_params["sockets"])", $htmlwhite))
		$rowdata += @(, ("License Server Address", ($htmlsilver -bor $htmlbold), $($XSHost.license_server["address"]), $htmlwhite))
		$rowdata += @(, ("License Server Port", ($htmlsilver -bor $htmlbold), "$($XSHost.license_server["port"])", $htmlwhite))
	}
	If ($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 0 ""
	}
	If ($HTML)
	{
		$msg = ""
		$columnWidths = @("150", "250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}

}

Function OutputHostVersion
{
	Param([object]$XSHost)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host Version Details"

	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Version Details"
	}
	If ($Text)
	{
		Line 2 "Version Details"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Version Details"
	}

	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()

	}
	If ($Text)
	{
	}
	If ($HTML)
	{
		$rowdata = @()
	}

	If ($MSWord -or $PDF)
	{
		$ScriptInformation += @{ Data = "Build date"; Value = "$($XSHost.software_version["date"])"; }
		$ScriptInformation += @{ Data = "Version"; Value = "$($XSHost.software_version["product_version_text"])"; }
		$ScriptInformation += @{ Data = "DBV"; Value = "$($XSHost.software_version["dbv"])"; }
	}
	If ($Text)
	{
		Line 3 "Build date`t: " "$($XSHost.software_version["date"])"
		Line 3 "Version`t`t: " "$($XSHost.software_version["product_version_text"])"
		Line 3 "DBV`t`t: " "$($XSHost.software_version["dbv"])"
	}
	If ($HTML)
	{
		$columnHeaders = @("Build date", ($htmlsilver -bor $htmlbold), "$($XSHost.software_version["date"])", $htmlwhite)
		$rowdata += @(, ("Version", ($htmlsilver -bor $htmlbold), $($XSHost.software_version["product_version_text"]), $htmlwhite))
		$rowdata += @(, ("DBV", ($htmlsilver -bor $htmlbold), "$($XSHost.software_version["dbv"])", $htmlwhite))
	}
	If ($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 0 ""
	}
	If ($HTML)
	{
		$msg = ""
		$columnWidths = @("150", "150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}

}

Function OutputHostUpdates
{
	Param([object]$XSHost)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host Updates"
	#$Updates = Get-XenHostPatch -SessionOpaqueRef $Script:Session.Opaque_Ref -EA 0 4>$Null | Select-Object name_label, version | Sort-Object name_label
	#$Updates = Get-XenHostPatch | `
	#Where-Object {(Get-XenHost -ref $_.host.opaque_ref).hostname -eq $XSHost.hostname} | `
	#Select-Object name_label, version | `
	#Sort-Object name_label
	
	$HostUpdates = $XSHost.updates 
	$HostUpdates = $HostUpdates | Sort-Object opaque_ref
	
	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $WordTable = @();
		
		WriteWordLine 3 0 "Updates" 
		
		ForEach ($HostUpdate in $HostUpdates)
		{
			$Update = Get-XenPoolUpdate -Ref $HostUpdate.opaque_ref -EA 0 4>$Null
			$WordTableRowHash = @{ 
				Update = "$($Update.name_label) (version $($Update.version))";
			}
			$WordTable += $WordTableRowHash;
		}
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $WordTable `
			-Columns Update `
			-Headers "Applied" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 1 "Updates"
		Line 2 "Applied`t: " ""
		ForEach ($HostUpdate in $HostUpdates)
		{
			$Update = Get-XenPoolUpdate -Ref $HostUpdate.opaque_ref -EA 0 4>$Null
			Line 3 "" "$($Update.name_label) (version $($Update.version))"
		}

		Line 0 ""
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Updates"
		$rowdata = @()

		ForEach ($HostUpdate in $HostUpdates)
		{
			$Update = Get-XenPoolUpdate -Ref $HostUpdate.opaque_ref -EA 0 4>$Null
			$rowdata += @(, (
					"$($Update.name_label) (version $($Update.version))", $htmlwhite))
		}
		
		$columnHeaders = @(
			'Applied', ($htmlsilver -bor $htmlbold))

		$msg = ""
		$columnWidths = @("150")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
	}
}

Function OutputHostManagement
{
	Param([object]$XSHost)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host Management Interface"

	$mPifs = $XSHost.PIFs | Get-XenPIF -EA 0 | Where-Object { $_.management }
	$managementIPs = @()
	ForEach ($pif in $mPifs)
	{
		If ($pif.IP)
		{
			$managementIPs += "$($pif.IP)"
		}
		ElseIf ($pif.ip_configuration_mode -eq [XenAPI.ip_configuration_mode]::DHCP)
		{
			$managementIPs += "DHCP"
		}
		Else
		{
			$managementIPs += "Unknown"
		}
	} 

	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Management Interface"
	}
	If ($Text)
	{
		Line 2 "Management Interface"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Management Interface"
	}
	
	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
	
	}
	If ($Text)
	{
	}
	If ($HTML)
	{
		$rowdata = @()
	}
	
	If ($MSWord -or $PDF)
	{
		$ScriptInformation += @{ Data = "DNS hostname"; Value = "$($XSHost.hostname)"; }
		$ScriptInformation += @{ Data = "Management interface"; Value = "$($managementIPs -join ", ")"; }
	}
	If ($Text)
	{
		Line 3 "DNS hostname`t`t: " "$($XSHost.hostname)"
		Line 3 "Management interface`t: " "$($managementIPs -join ", ")"
	}
	If ($HTML)
	{
		$columnHeaders = @("DNS hostname", ($htmlsilver -bor $htmlbold), "$($XSHost.hostname)", $htmlwhite)
		$rowdata += @(, ("Management interface", ($htmlsilver -bor $htmlbold), $($managementIPs -join ", "), $htmlwhite))
	}
	If ($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;
	
		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
	
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 150;
	
		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)
	
		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 0 ""
	}
	If ($HTML)
	{
		$msg = ""
		$columnWidths = @("150", "150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
	
}

Function OutputHostMemoryOverview
{
	Param([object]$XSHost)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host Memory Overview"

	$XSHostMetrics = $XSHost.metrics | Get-XenHostMetrics
	$memTotal = Convert-SizeToString -size $XSHostMetrics.memory_total -Decimal 1
	$memFree = Convert-SizeToString -size $XSHostMetrics.memory_free -Decimal 1
	$memoryText = "$memFree RAM available ($memTotal)"
	$hostAllRunningVMs = @( $XSHost.resident_VMs | Get-XenVM | Sort-Object -Property name_label)
	$hostRunningVMs = @($hostAllRunningVMs | Where-Object { $_.is_control_domain -eq $false -and $_.power_state -like "running" })
	$dom0VM = $hostAllRunningVMs | Where-Object { $_.is_control_domain -eq $true }
	$vmText = @()
	$vmMemoryUsed = [Int64]0
	ForEach ($vm in $hostRunningVMs)
	{
		$vmText += '{0}: using  {1}' -f $vm.name_label, $(Convert-SizeToString -size $vm.memory_target  -Decimal 1)
		$vmMemoryUsed = $vmMemoryUsed + $vm.memory_target
	}

	$memXSNum = $($dom0VM.memory_target + $XSHost.memory_overhead)
	$memXS = Convert-SizeToString -size $memXSNum -Decimal 1
	$memXSUsedNum = ($dom0VM.memory_target + $XSHost.memory_overhead + $vmMemoryUsed)
	$memXSAvailableNum = ($XSHostMetrics.memory_total - $memXSUsedNum)
	$memXSUsed = Convert-SizeToString -size $memXSUsedNum -Decimal 1
	$memXSUsedPct = '{0}%' -f [Math]::Round($memXSUsedNum / ($XSHostMetrics.memory_total / 100))
	$memXSAvailable = Convert-SizeToString -size $memXSAvailableNum -Decimal 1
	$cdMemory = Convert-SizeToString -size $dom0VM.memory_target -Decimal 1


	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Memory"
	}
	If ($Text)
	{
		Line 2 "Memory"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Memory"
	}
	
	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
	
	}
	If ($Text)
	{
	}
	If ($HTML)
	{
		$rowdata = @()
	}
	
	If ($MSWord -or $PDF)
	{
		$ScriptInformation += @{ Data = "Server"; Value = "$($memoryText)"; }
		$ScriptInformation += @{ Data = "VMs"; Value = "$($hostRunningVMs.Count)"; }
		$vmText | ForEach-Object { $ScriptInformation += @{ Data = ""; Value = "$($_)"; } }
		$ScriptInformation += @{ Data = "Citrix Hypervisor"; Value = "$($memXS)"; }
		$ScriptInformation += @{ Data = "Control domain memory"; Value = "$($cdMemory)"; }
		$ScriptInformation += @{ Data = "Available memory"; Value = "$($memXSAvailable)"; }
		$ScriptInformation += @{ Data = "Total max memory"; Value = "$($memXSUsed) ($memXSUsedPct of total memory)"; }
	}
	If ($Text)
	{
		Line 3 "Server`t`t`t: " "$($memoryText)"
		Line 3 "VMs`t`t`t: " "$($hostRunningVMs.Count)"
		$vmText | ForEach-Object { Line 6 "  $($_)" }
		Line 3 "Citrix Hypervisor`t: " "$($memXS)"
		Line 3 "Control domain memory`t: " "$($cdMemory)"
		Line 3 "Available memory`t: " "$($memXSAvailable)"
		Line 3 "Total max memory`t: " "$($memXSUsed) ($memXSUsedPct of total memory)"
	}
	If ($HTML)
	{
		$columnHeaders = @("Server", ($htmlsilver -bor $htmlbold), "$($memoryText)", $htmlwhite)
		$rowdata += @(, ("VMs", ($htmlsilver -bor $htmlbold), "$($hostRunningVMs.Count)", $htmlwhite))
		$vmText | ForEach-Object { $rowdata += @(, ("", ($htmlsilver -bor $htmlbold), "$($_)", $htmlwhite)) }
		$rowdata += @(, ("Citrix Hypervisor", ($htmlsilver -bor $htmlbold), "$($memXS)", $htmlwhite))
		$rowdata += @(, ("Control domain memory", ($htmlsilver -bor $htmlbold), "$($cdMemory)", $htmlwhite))
		$rowdata += @(, ("Available memory", ($htmlsilver -bor $htmlbold), "$($memXSAvailable)", $htmlwhite))
		$rowdata += @(, ("Total max memory", ($htmlsilver -bor $htmlbold), "$($memXSUsed) ($memXSUsedPct of total memory)", $htmlwhite))
	}
	If ($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;
	
		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
	
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;
	
		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)
	
		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 0 ""
	}
	If ($HTML)
	{
		$msg = ""
		$columnWidths = @("150", "200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
		
}

Function OutputHostCPUs
{
	Param([object]$XSHost)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host CPUs"

	$hostCPUCount = '0 - {0}' -f $($XSHost.cpu_info["cpu_count"] - 1)
	$hostCPUInfo = @()
	$hostCPUInfo += 'Vendor:   {0}' -f $XSHost.cpu_info["vendor"]
	$hostCPUInfo += 'Model:   {0}' -f $XSHost.cpu_info["modelname"]
	$hostCPUInfo += 'Speed:   {0} MHz' -f [math]::Round($XSHost.cpu_info["speed"])

	$hostVendor = $XSHost.cpu_info["vendor"]
	$hostModel = $XSHost.cpu_info["modelname"]
	$hostSpeed = '{0} MHz' -f [math]::Round($XSHost.cpu_info["speed"])



	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "CPUs"
	}
	If ($Text)
	{
		Line 2 "CPUs"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "CPUs"
	}
	
	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
	}
	If ($Text)
	{
	}
	If ($HTML)
	{
		$rowdata = @()
	}
	
	If ($MSWord -or $PDF)
	{
		$ScriptInformation += @{ Data = "CPU"; Value = "$($hostCPUCount)"; }
		$ScriptInformation += @{ Data = "Vendor"; Value = "$($hostVendor)"; }
		$ScriptInformation += @{ Data = "Model"; Value = "$($hostModel)"; }
		$ScriptInformation += @{ Data = "Speed"; Value = "$($hostSpeed)"; }
	}
	If ($Text)
	{
		Line 3 "CPU`t: " "$($hostCPUCount)"
		Line 3 "Vendor`t: " "$($hostVendor)"
		Line 3 "Model`t: " "$($hostModel)"
		Line 3 "Speed`t: " "$($hostSpeed)"
	}
	If ($HTML)
	{
		$columnHeaders = @("CPU", ($htmlsilver -bor $htmlbold), "$($hostCPUCount)", $htmlwhite)
		$rowdata += @(, ("Vendor", ($htmlsilver -bor $htmlbold), $($hostVendor), $htmlwhite))
		$rowdata += @(, ("Model", ($htmlsilver -bor $htmlbold), $($hostModel), $htmlwhite))
		$rowdata += @(, ("Speed", ($htmlsilver -bor $htmlbold), $($hostSpeed), $htmlwhite))
	}
	If ($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;
	
		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
	
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;
	
		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)
	
		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 0 ""
	}
	If ($HTML)
	{
		$msg = ""
		$columnWidths = @("150", "300")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
		
}

Function OutputHostGeneral
{
	Param([object]$XSHost)
	
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host General"
	
	[array]$xtags = @()
	ForEach ($tag in $XSHost.tags)
	{
		$xtags += $tag
	}
	If ($xtags.count -gt 0)
	{
		[array]$xtags = $xtags | Sort-Object
	}
	Else
	{
		[array]$xtags = @("<None>")
	}
	
	If ([String]::IsNullOrEmpty($($XSHost.Other_Config["folder"])))
	{
		$folderName = "None"
	}
	Else
	{
		$folderName = $XSHost.Other_Config["folder"]
	}

	If ($XSHost.name_description -eq "Default install")
	{
		$XSHostDescription = "Default install of Citrix Hypervisor"
	}
	ELse
	{
		$XSHostDescription = $XSHost.name_description
	}
	
	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Name"; Value = $XSHost.name_label; }
		$ScriptInformation += @{ Data = "Description"; Value = $XSHostDescription; }
		$ScriptInformation += @{ Data = "Folder"; Value = $folderName; }
		$ScriptInformation += @{ Data = "Tags"; Value = $($xtags -join ", "); }
		$ScriptInformation += @{ Data = "iSCSI IQN"; Value = $XSHost.iscsi_iqn; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 2 "Description`t`t: " $XSHostDescription
		Line 2 "Folder`t`t`t: " $folderName
		Line 2 "Tags`t`t`t: " "$($xtags -join ", ")"
		Line 2 "iSCSI IQN`t`t: " $XSHost.iscsi_iqn
		Line 0 ""
	}
	If ($HTML)
	{
		#for HTML output, remove the < and > from <None> xtags and foldername if they are there
		$xtags = $xtags.Trim("<", ">")
		$folderName = $folderName.Trim("<", ">")
		$rowdata = @()
		$columnHeaders = @("Name", ($htmlsilver -bor $htmlbold), $XSHost.name_label, $htmlwhite)
		$rowdata += @(, ('Description', ($htmlsilver -bor $htmlbold), $XSHostDescription, $htmlwhite))
		$rowdata += @(, ('Folder', ($htmlsilver -bor $htmlbold), "$folderName", $htmlwhite))
		$rowdata += @(, ('Tags', ($htmlsilver -bor $htmlbold), "$($xtags -join ", ")", $htmlwhite))
		$rowdata += @(, ("iSCSI IQN", ($htmlsilver -bor $htmlbold), $XSHost.iscsi_iqn, $htmlwhite))

		$msg = ""
		$columnWidths = @("150", "250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLIne 0 0 ""
	}
}

Function OutputHostCustomFields
{
	Param([object] $XSHost)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host Custom Fields"

	$CustomFields = Get-XSCustomFields $($XSHost.other_config)
	
	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Custom Fields"
	}
	If ($Text)
	{
		Line 2 "Custom Fields"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Custom Fields"
	}
	
	If ([String]::IsNullOrEmpty($CustomFields) -or $CustomFields.Count -eq 0)
	{
		$HostName = $XSHost.Name_Label

		If ($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There are no Custom Fields for Host $HostName"
		}
		If ($Text)
		{
			Line 3 "There are no Custom Fields for Host $HostName"
			Line 0 ""
		}
		If ($HTML)
		{
			WriteHTMLLine 0 1 "There are no Custom Fields for Host $HostName"
		}
	}
	Else
	{
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
		}
		If ($Text)
		{
			#nothing
		}
		If ($HTML)
		{
			$rowdata = @()
		}

		[int]$cnt = -1
		ForEach ($Item in $CustomFields)
		{
			$cnt++
			If ($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = $($Item.Name); Value = $Item.Value; }
			}
			If ($Text)
			{
				Line 3 "$($Item.Name): " $Item.Value
			}
			If ($HTML)
			{
				If ($cnt -eq 0)
				{
					$columnHeaders = @($($Item.Name), ($htmlsilver -bor $htmlbold), $Item.Value, $htmlwhite)
				}
				Else
				{
					$rowdata += @(, ($($Item.Name), ($htmlsilver -bor $htmlbold), $Item.Value, $htmlwhite))
				}
			}
		}
		
		If ($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data, Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 250;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If ($Text)
		{
			Line 0 ""
		}
		If ($HTML)
		{
			$msg = ""
			$columnWidths = @("250", "250")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
}

Function OutputHostAlerts
{
	Param([object]$XSHost)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host Alerts"
	
	<#
		On the host's Properties, Alerts, there are five items:
		
		Alert repeat interval: nn minutes - This is found in the host's properties for each key/value pair
		
		Generate CPU usage alerts - These are found in the host's properties. If exists, enabled else disabled
			When CPU usage exceeds nn %
			For longer than nn minutes
			
		Generate network usage alerts - These are found in the host's properties. If exists, enabled else disabled
			When network usage exceeds nnn KB/s
			For longer than nn minutes
			
		Generate memory usage alerts - These are found in the host's properties. If exists, enabled else disabled
			When free memory falls below nnnn MB
			For longer than nn minutes
		
		Generate control domain memory usage alerts - These are found in Dom 0's VM properties. If exists, enabled else disabled
			When control domain memory usage exceeds nn %
			For longer than nn minutes

		To get Dom 0's properties:
		$Session = Connect-XenServer -server 192.168.1.82 -SetDefaultSession -NoWarnCertificates -PassThru
		$dom0conf = Get-XenVm | `
		Where-Object {$_.is_control_domain -and $_.domid -eq 0 -and $_.name_label -like "*$($XSHost.hostname)*"} | `
		Get-XenVmProperty -XenProperty OtherConfig
		[xml]$XML = $dom0conf.perfmon
		ForEach($Alert in $XML.config.variable)
		{
			"Alert Name: $($Alert.Name.Value)"	#mem_usage
			"Trigger level: $($Alert.alarm_trigger_level.Value)"
			"Trigger period: $($Alert.alarm_trigger_period.Value)"
			"Inhibit period: $($Alert.alarm_auto_inhibit_period.value)"
		}

		To get the host's alert properties:
		$OtherConfig = ($XSHost | Get-XenHostProperty -XenProperty OtherConfig -EA 0)
		[xml]$XML = $OtherConfig.perfmon
			
		ForEach($Alert in $XML.config.variable)
		{
			"Host: $($XSHost.name_label)"
			"Alert Name: $($Alert.Name.Value)"
			"Trigger level: $($Alert.alarm_trigger_level.Value)"
			"Trigger period: $($Alert.alarm_trigger_period.Value)"
			"Inhibit period: $($Alert.alarm_auto_inhibit_period.value)"
		}

		Alert Name: mem_usage Trigger level: 0.9 Trigger period: 120 Inhibit period: 3900
		Alert Name: mem_usage Trigger level: 0.95 Trigger period: 240 Inhibit period: 1800
		Host: XenServer2 (82) - Alert Name: cpu_usage Trigger level: 0.5 Trigger period: 60 Inhibit period: 1800
		Host: XenServer2 (82) - Alert Name: network_usage Trigger level: 102400 Trigger period: 120 Inhibit period: 1800
		Host: XenServer2 (82) - Alert Name: memory_free_kib Trigger level: 1024000 Trigger period: 180 Inhibit period: 1800

		cpu_usage 0.5 60 60 #cpu_usage is Generate CPU usage alerts, .5 is When CPU usage exceeds: 50%, 60 is For longer than 1 minutes
		network_usage 102400 120 120 #network_usage is Generate network usage alerts, 102400 is When network usage exceeds 100 KB/s, 120 is For longer than 2 minutes
		memory_free_kib 1024000 180 180 #memory_free_kib is Generate memory usage alerts, 1024000 is When free memory falls below 1000 MB, 180 is For longer than 3 minutes
	#>

	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Alerts"
	}
	If ($Text)
	{
		Line 2 "Alerts"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Alerts"
	}

	$GenerateDom0MemUsageAlerts = "Not selected"
	[double]$WhenDom0MemUsageExceeds = 0
	[int32]$WhenDom0ForLongerThan = 0
	
	$dom0conf = Get-XenVM | `
			Where-Object { $_.is_control_domain -and $_.domid -eq 0 -and $_.name_label -like "*$($XSHost.hostname)*" } | `
			Get-XenVMProperty -XenProperty OtherConfig
	[xml]$XML = $dom0conf.perfmon

	ForEach ($Alert in $XML.config.variable)
	{
		If ($Alert.Name.Value -eq "mem_usage")
		{
			$GenerateDom0MemUsageAlerts = "Selected"
			[double]$tmp = $Alert.alarm_trigger_level.Value
			$WhenDom0MemUsageExceeds = $tmp * 100
			$WhenDom0ForLongerThan = $Alert.alarm_trigger_period.Value / 60
		}
	}
	
	[int32]$AlertRepeatInterval = 0

	$GenerateHostCPUUsageAlerts = "Not selected"
	[double]$WhenHostCPUUsageExceeds = 0
	[int32]$WhenHostCPUForLongerThan = 0

	$GenerateHostNetworkUsageAlerts = "Not selected"
	[int32]$WhenHostNetworkUsageExceeds = 0
	[int32]$WhenHostNetworkForLongerThan = 0

	$GenerateHostMemoryUsageAlerts = "Not selected"
	[int32]$WhenHostMemUsageExceeds = 0
	[int32]$WhenHostMemForLongerThan = 0

	$OtherConfig = ($XSHost | Get-XenHostProperty -XenProperty OtherConfig -EA 0)
	
	If ($OtherConfig.ContainsKey("perfmon"))
	{
		[xml]$XML = $OtherConfig.perfmon
			
		ForEach ($Alert in $XML.config.variable)
		{
			If ($Alert.Name.Value -eq "cpu_usage")
			{
				$GenerateHostCPUUsageAlerts = "Selected"
				[double]$tmp = $Alert.alarm_trigger_level.Value
				$WhenHostCPUUsageExceeds = $tmp * 100
				$WhenHostCPUForLongerThan = $Alert.alarm_trigger_period.Value / 60
				$AlertRepeatInterval = $Alert.alarm_auto_inhibit_period.Value / 60
			}
			ElseIf ($Alert.Name.Value -eq "network_usage")
			{
				$GenerateHostNetworkUsageAlerts = "Selected"
				$WhenHostNetworkUsageExceeds = $Alert.alarm_trigger_level.Value / 1024
				$WhenHostNetworkForLongerThan = $Alert.alarm_trigger_period.Value / 60
				$AlertRepeatInterval = $Alert.alarm_auto_inhibit_period.Value / 60
			}
			ElseIf ($Alert.Name.Value -eq "memory_free_kib")
			{
				$GenerateHostMemoryUsageAlerts = "Selected"
				$WhenHostMemUsageExceeds = $Alert.alarm_trigger_level.Value / 1024
				$WhenHostMemForLongerThan = $Alert.alarm_trigger_period.Value / 60
				$AlertRepeatInterval = $Alert.alarm_auto_inhibit_period.Value / 60
			}
		}
	}

	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		If ($GenerateHostCPUUsageAlerts -eq "Selected" -or
			$GenerateHostNetworkUsageAlerts -eq "Selected" -or
			$GenerateHostMemoryUsageAlerts -eq "Selected" -or
			$GenerateDom0MemUsageAlerts -eq "Selected")
		{
			$ScriptInformation += @{ Data = "Alert repeat interval"; Value = "$($AlertRepeatInterval) minutes"; }
		}
		Else
		{
			$ScriptInformation += @{ Data = "Alert repeat interval"; Value = "Not Set"; }
		}
		$ScriptInformation += @{ Data = "Generate CPU usage alerts"; Value = $GenerateHostCPUUsageAlerts; }
		If ($GenerateHostCPUUsageAlerts -eq "Selected")
		{
			$ScriptInformation += @{ Data = "     When CPU usage exceeds"; Value = "$($WhenHostCPUUsageExceeds) %"; }
			$ScriptInformation += @{ Data = "     For longer than"; Value = "$($WhenHostCPUForLongerThan) minutes"; }
		}
		$ScriptInformation += @{ Data = "Generate network usage alerts"; Value = $GenerateHostNetworkUsageAlerts; }
		If ($GenerateHostNetworkUsageAlerts -eq "Selected")
		{
			$ScriptInformation += @{ Data = "     When network usage exceeds"; Value = "$($WhenHostNetworkUsageExceeds) KB/s"; }
			$ScriptInformation += @{ Data = "     For longer than"; Value = "$($WhenHostNetworkForLongerThan) minutes"; }
		}
		$ScriptInformation += @{ Data = "Generate memory usage alerts"; Value = $GenerateHostMemoryUsageAlerts; }
		If ($GenerateHostMemoryUsageAlerts -eq "Selected")
		{
			$ScriptInformation += @{ Data = "     When memory usage exceeds"; Value = "$($WhenHostMemUsageExceeds) MB"; }
			$ScriptInformation += @{ Data = "     For longer than"; Value = "$($WhenHostMemForLongerThan) minutes"; }
		}
		$ScriptInformation += @{ Data = "Generate control domain memory usage alerts"; Value = $GenerateDom0MemUsageAlerts; }
		If ($GenerateDom0MemUsageAlerts -eq "Selected")
		{
			$ScriptInformation += @{ Data = "    When control domain memory usage exceeds "; Value = "$($WhenDom0MemUsageExceeds) %"; }
			$ScriptInformation += @{ Data = "     For longer than"; Value = "$($WhenDom0ForLongerThan) minutes"; }
		}

		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 250;
		$Table.Columns.Item(2).Width = 100;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		If ($GenerateHostCPUUsageAlerts -eq "Selected" -or
			$GenerateHostNetworkUsageAlerts -eq "Selected" -or
			$GenerateHostMemoryUsageAlerts -eq "Selected" -or
			$GenerateDom0MemUsageAlerts -eq "Selected")
		{
			Line 3 "Alert repeat interval`t`t`t`t: $($AlertRepeatInterval) minutes"
		}
		Else
		{
			Line 3 "Alert repeat interval`t`t`t`t: Not Set"
		}
		Line 3 "Generate CPU usage alerts`t`t`t: " $GenerateHostCPUUsageAlerts
		If ($GenerateHostCPUUsageAlerts -eq "Selected")
		{
			Line 4 "When CPU usage exceeds    : " "$($WhenHostCPUUsageExceeds) %"
			Line 4 "For longer than           : " "$($WhenHostCPUForLongerThan) minutes"
		}
		Line 3 "Generate network usage alerts`t`t`t: " $GenerateHostNetworkUsageAlerts
		If ($GenerateHostNetworkUsageAlerts -eq "Selected")
		{
			Line 4 "When network usage exceeds: " "$($WhenHostNetworkUsageExceeds) KB/s"
			Line 4 "For longer than           : " "$($WhenHostNetworkForLongerThan) minutes"
		}
		Line 3 "Generate memory usage alerts`t`t`t: " $GenerateHostMemoryUsageAlerts
		If ($GenerateHostMemoryUsageAlerts -eq "Selected")
		{
			Line 4 "When memory usage exceeds : " "$($WhenHostMemUsageExceeds) MB"
			Line 4 "For longer than           : " "$($WhenHostMemForLongerThan) minutes"
		}
		Line 3 "Generate control domain memory usage alerts`t: " $GenerateDom0MemUsageAlerts
		If ($GenerateDom0MemUsageAlerts -eq "Selected")
		{
			Line 4 "When control domain memory usage exceeds: " "$($WhenDom0MemUsageExceeds) %"
			Line 4 "For longer than                         : " "$($WhenDom0ForLongerThan) minutes"
		}
		Line 0 ""
	}
	If ($HTML)
	{
		$rowdata = @()
		If ($GenerateHostCPUUsageAlerts -eq "Selected" -or
			$GenerateHostNetworkUsageAlerts -eq "Selected" -or
			$GenerateHostMemoryUsageAlerts -eq "Selected" -or
			$GenerateDom0MemUsageAlerts -eq "Selected")
		{
			$columnHeaders = @("Alert repeat interval", ($htmlsilver -bor $htmlbold), "$($AlertRepeatInterval) minutes", $htmlwhite)
		}
		Else
		{
			$columnHeaders = @("Alert repeat interval", ($htmlsilver -bor $htmlbold), "Not set", $htmlwhite)
		}
		$rowdata += @(, ("Generate CPU usage alerts", ($htmlsilver -bor $htmlbold), $GenerateHostCPUUsageAlerts, $htmlwhite))
		If ($GenerateHostCPUUsageAlerts -eq "Selected")
		{
			$rowdata += @(, ("     When CPU usage exceeds", ($htmlsilver -bor $htmlbold), "$($WhenHostCPUUsageExceeds) %", $htmlwhite))
			$rowdata += @(, ("     For longer than", ($htmlsilver -bor $htmlbold), "$($WhenHostCPUForLongerThan) minutes", $htmlwhite))
		}
		$rowdata += @(, ("Generate network usage alerts", ($htmlsilver -bor $htmlbold), $GenerateHostNetworkUsageAlerts , $htmlwhite))
		If ($GenerateHostNetworkUsageAlerts -eq "Selected")
		{
			$rowdata += @(, ("     When network usage exceeds", ($htmlsilver -bor $htmlbold), "$($WhenHostNetworkUsageExceeds) KB/s", $htmlwhite))
			$rowdata += @(, ("     For longer than", ($htmlsilver -bor $htmlbold), "$($WhenHostNetworkForLongerThan) minutes", $htmlwhite))
		}
		$rowdata += @(, ("Generate memory usage alerts", ($htmlsilver -bor $htmlbold), $GenerateHostMemoryUsageAlerts, $htmlwhite))
		If ($GenerateHostMemoryUsageAlerts -eq "Selected")
		{
			$rowdata += @(, ("     When memory usage exceeds", ($htmlsilver -bor $htmlbold), "$($WhenHostMemUsageExceeds) MB", $htmlwhite))
			$rowdata += @(, ("     For longer than", ($htmlsilver -bor $htmlbold), "$($WhenHostMemForLongerThan) minutes", $htmlwhite))
		}
		$rowdata += @(, ("Generate control domain memory usage alerts", ($htmlsilver -bor $htmlbold), $GenerateDom0MemUsageAlerts, $htmlwhite))
		If ($GenerateDom0MemUsageAlerts -eq "Selected")
		{
			$rowdata += @(, ("     When control domain memory usage exceeds", ($htmlsilver -bor $htmlbold), "$($WhenDom0MemUsageExceeds) %", $htmlwhite))
			$rowdata += @(, ("     For longer than", ($htmlsilver -bor $htmlbold), "$($WhenDom0ForLongerThan) minutes", $htmlwhite))
		}

		$msg = ""
		$columnWidths = @("275", "100")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLIne 0 0 ""
	}
}

Function OutputHostMultipathing
{
	Param([object]$XSHost)
}

Function OutputHostLogDestination
{
	Param([object]$XSHost)
	
	$LogLocation = ""
	If ($XSHost.Logging.Count -eq 0)
	{
		$LogLocation = "Local"
		$StoreLogs = "Not selected"
		$LogServer = ""
	}
	Else
	{
		$LogLocation = "Local and Remote ($($XSHost.logging.syslog_destination))"
		$StoreLogs = "Selected"
		$LogServer = "$($XSHost.logging.syslog_destination)"
	}
	
	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Log Destination"
	}
	If ($Text)
	{
		Line 2 "Log Destination"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Log Destination"
	}
	
	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Log destination"; Value = $LogLocation; }
		$ScriptInformation += @{ Data = "Also store the system logs on a remote server"; Value = $StoreLogs; }
		$ScriptInformation += @{ Data = "Server"; Value = $LogServer; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 250;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 3 "Log destination`t`t`t`t`t: " $LogLocation
		Line 3 "Also store the system logs on a remote server`t: " $StoreLogs
		Line 3 "Server`t`t`t`t`t`t: " $LogServer
		Line 0 ""
	}
	If ($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Log destination", ($htmlsilver -bor $htmlbold), $LogLocation, $htmlwhite)
		$rowdata += @(, ("Also store the system logs on a remote server", ($htmlsilver -bor $htmlbold), $StoreLogs, $htmlwhite))
		$rowdata += @(, ("Server", ($htmlsilver -bor $htmlbold), $LogServer, $htmlwhite))

		$msg = ""
		$columnWidths = @("275", "250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLIne 0 0 ""
	}
}

Function OutputHostPowerOn
{
	Param([object]$XSHost)

	#for the host, get the power_on_mode and power_on_config properties
	#If power_on_mode -eq "", then it is Disabled
	#DRAC is Dell Remote Access Controller (DRAC)
	#wake-on-lan is Wake-on-LAN (WoL)
	#Otherwise, power_on_mode is Custom power-on script /etc/xapi.d/plugins/<value of power_on_mode>
	
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host Power On"

	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Power On"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
	}
	If ($Text)
	{
		Line 2 "Power On"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Power On"
		$rowdata = @()
	}
	
	[int]$cnt = -1
	$cnt++
	If ($XSHost.power_on_mode -eq "")
	{
		#disabled
		If ($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Server"; Value = "$($XSHost.Name_Label)"; }
			$ScriptInformation += @{ Data = "     Power On mode"; Value = "Disabled"; }
		}
		If ($Text)
		{
			Line 3 "Server`t`t: " "$($XSHost.Name_Label)"
			Line 3 "Power On mode`t: " "Disabled"
		}
		If ($HTML)
		{
			If ($cnt -eq 0)
			{
				$cnt++
				$columnHeaders = @("Server", ($htmlsilver -bor $htmlbold), "$($XSHost.Name_Label)", $htmlwhite)
				$rowdata += @(, ("    Power On mode", ($htmlsilver -bor $htmlbold), "Disabled", $htmlwhite))
			}
			Else
			{
				$rowdata += @(, ("Server", ($htmlsilver -bor $htmlbold), "$($XSHost.Name_Label)", $htmlwhite))
				$rowdata += @(, ("    Power On mode", ($htmlsilver -bor $htmlbold), "Disabled", $htmlwhite))
			}
		}
	}
	ElseIf ($XSHost.power_on_mode -eq "DRAC")
	{
		[array]$PowerKeys = $XSHost.power_on_config.Keys.Split() 
		[array]$PowerValues = $XSHost.power_on_config.Values.Split() 
		If ($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Server"; Value = "$($XSHost.Name_Label)"; }
			$ScriptInformation += @{ Data = "Power On mode"; Value = "Dell Remote Access Controller (DRAC)"; }
			$ScriptInformation += @{ Data = "Configuration options"; Value = ""; }

			[int]$cnt2 = -1
			ForEach ($Item in $PowerKeys)
			{
				$cnt2++
				$Value = $PowerValues[$cnt2]
				
				If ($Item -like "*power_on_ip*")
				{
					$ScriptInformation += @{ Data = "     IP address"; Value = $Value; }
				}
				If ($Item -like "*power_on_user*")
				{
					$ScriptInformation += @{ Data = "     Username"; Value = $Value; }
				}
			}
		}
		If ($Text)
		{
			Line 3 "Server`t`t: " "$($XSHost.Name_Label)"
			Line 3 "Power On mode`t: " "Dell Remote Access Controller (DRAC)"
			Line 3 "Configuration options"

			[int]$cnt2 = -1
			ForEach ($Item in $PowerKeys)
			{
				$cnt2++
				$Value = $PowerValues[$cnt2]
				
				If ($Item -like "*power_on_ip*")
				{
					Line 4 "IP address: " $Value
				}
				If ($Item -like "*power_on_user*")
				{
					Line 4 "Username  : " $Value
				}
			}
		}
		If ($HTML)
		{
			If ($cnt -eq 0)
			{
				$cnt++
				$columnHeaders = @("Server", ($htmlsilver -bor $htmlbold), "$($XSHost.Name_Label)", $htmlwhite)
			}
			Else
			{
				$rowdata += @(, ("Server", ($htmlsilver -bor $htmlbold), "$($XSHost.Name_Label)", $htmlwhite))
			}
			$rowdata += @(, ("Power On mode", ($htmlsilver -bor $htmlbold), "Dell Remote Access Controller (DRAC)", $htmlwhite))
			$rowdata += @(, ("Configuration options", ($htmlsilver -bor $htmlbold), "", $htmlwhite))

			[int]$cnt2 = -1
			ForEach ($Item in $PowerKeys)
			{
				$cnt2++
				$Value = $PowerValues[$cnt2]
				
				If ($Item -like "*power_on_ip*")
				{
					$rowdata += @(, ("     IP address: ", ($htmlsilver -bor $htmlbold), $Value, $htmlwhite))
				}
				If ($Item -like "*power_on_user*")
				{
					$rowdata += @(, ("     Username: ", ($htmlsilver -bor $htmlbold), $Value, $htmlwhite))
				}
			}
		}
	}
	ElseIf ($XSHost.power_on_mode -eq "wake-on-lan")
	{
		If ($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Server"; Value = "$($XSHost.Name_Label)"; }
			$ScriptInformation += @{ Data = "Power On mode"; Value = "Wake-on-LAN (WoL)"; }
		}
		If ($Text)
		{
			Line 3 "Server`t`t: " "$($XSHost.Name_Label)"
			Line 3 "Power On mode`t: " "Wake-on-LAN (WoL)"
		}
		If ($HTML)
		{
			If ($cnt -eq 0)
			{
				$cnt++
				$columnHeaders = @("Server", ($htmlsilver -bor $htmlbold), "$($XSHost.Name_Label)", $htmlwhite)
			}
			Else
			{
				$rowdata += @(, ("Server", ($htmlsilver -bor $htmlbold), "$($XSHost.Name_Label)", $htmlwhite))
			}
			$rowdata += @(, ("Power On mode", ($htmlsilver -bor $htmlbold), "Wake-on-LAN (WoL)", $htmlwhite))
		}
	}
	Else
	{
		#custom script
		[array]$PowerKeys = $XSHost.power_on_config.Keys.Split() 
		[array]$PowerValues = $XSHost.power_on_config.Values.Split() 
		If ($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Server"; Value = "$($XSHost.Name_Label)"; }
			$ScriptInformation += @{ Data = "Power On mode"; Value = "Custom power-on script /etc/xapi.d/plugins/$($XSHost.power_on_mode)"; }
			$ScriptInformation += @{ Data = "Configuration options"; Value = ""; }

			[int]$cnt2 = -1
			ForEach ($Item in $PowerKeys)
			{
				$cnt2++
				$Value = $PowerValues[$cnt2]
				
				$ScriptInformation += @{ Data = "     Key: $Item"; Value = "Value: $Value"; }
			}
		}
		If ($Text)
		{
			Line 3 "Server`t`t: " "$($XSHost.Name_Label)"
			Line 3 "Power On mode`t: " "Custom power-on script /etc/xapi.d/plugins/$($XSHost.power_on_mode)"
			Line 3 "Configuration options"

			[int]$cnt2 = -1
			ForEach ($Item in $PowerKeys)
			{
				$cnt2++
				$Value = $PowerValues[$cnt2]
				Line 4 "Key  : " $Item
				Line 4 "Value: " $Value
				Line 0 ""
			}
		}
		If ($HTML)
		{
			If ($cnt -eq 0)
			{
				$cnt++
				$columnHeaders = @("Server", ($htmlsilver -bor $htmlbold), "$($XSHost.Name_Label)", $htmlwhite)
			}
			Else
			{
				$rowdata += @(, ("Server", ($htmlsilver -bor $htmlbold), "$($XSHost.Name_Label)", $htmlwhite))
			}
			$rowdata += @(, ("Power On mode", ($htmlsilver -bor $htmlbold), "Custom power-on script /etc/xapi.d/plugins/$($XSHost.power_on_mode)", $htmlwhite))
			$rowdata += @(, ("Configuration options", ($htmlsilver -bor $htmlbold), "", $htmlwhite))

			[int]$cnt2 = -1
			ForEach ($Item in $PowerKeys)
			{
				$cnt2++
				$Value = $PowerValues[$cnt2]
			
				$rowdata += @(, ("     Key: $Item", ($htmlsilver -bor $htmlbold), "Value: $Value", $htmlwhite))
			}
		}
	}	

	If ($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 300;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 0 ""
	}
	If ($HTML)
	{
		$msg = ""
		$columnWidths = @("150", "300")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputHostGPUProperties
{
	Param([object]$XSHost)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host GPU"
	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "GPU"
	}
	If ($Text)
	{
		Line 2 "GPU"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "GPU"
	}

	<#
		From the XenServer team:
		
		XC finds the host pgpu for which is_system_display_device=true
		then checks this pgpu's dom0_access
		
		the server is using integrated gpu if both Host.display and the above dom0_access 
		have value enabled or disable_on_reboot (note that despite the same name the values are of different type)
		
		the server will use integrated gpu on next reboot if both Host.display and the above dom0_access have value enabled or enable_on_reboot.
	#>
	
	$HostGPU = $XSHost.PGPUs | Get-XenPGPU
	
	<#
		This code is from the XenServer team
	#>
	If(
		( ($XSHost.display -eq [XenAPI.host_display]::enabled) -or ($XSHost.display -eq [XenAPI.host_display]::disable_on_reboot) ) -and
		( ($HostGPU.dom0_access -eq [XenAPI.pgpu_dom0_access]::enabled) -or ($HostGPU.dom0_access -eq [XenAPI.pgpu_dom0_access]::disable_on_reboot) )
	  )
	{
		$HostGPUTxt = "This server is currently using the integrated GPU"
	}
	Else
	{
		$HostGPUTxt = "This server is currently not using the integrated GPU"
	}

	If(
		( ($XSHost.display -eq [XenAPI.host_display]::enabled) -or ($XSHost.display -eq [XenAPI.host_display]::enable_on_reboot) ) -and
		( ($HostGPU.dom0_access -eq [XenAPI.pgpu_dom0_access]::enabled) -or ($HostGPU.dom0_access -eq [XenAPI.pgpu_dom0_access]::enable_on_reboot) )
	  )
	{
		$HostGPUTxt = "This server will use the integrated GPU on next reboot"
	}
	Else
	{
		$HostGPUTxt = "This server will not use the integrated GPU on next reboot"
	}
	
	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = $HostGPUTxt; Value = ""; }
	}
	If ($Text)
	{
		Line 3 $HostGPUTxt
	}
	If ($HTML)
	{
		$columnHeaders = @($HostGPUTxt, ($htmlsilver -bor $htmlbold), "", $htmlwhite)
		$rowdata = @()
	}

	If ($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 250;
		$Table.Columns.Item(2).Width = 20;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 0 ""
	}
	If ($HTML)
	{
		$msg = ""
		$columnWidths = @("250", "10")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputHostMemory
{
	Param([object]$XSHost)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host Memory"

	$XSHostMemory = @($Script:XSPoolMemories | Where-Object { $_.XSHostRef -like $XSHost.opaque_ref } )

	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Memory"
	}
	If ($Text)
	{
		Line 2 "Memory"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Memory"
	}
	
	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
	
	}
	If ($Text)
	{
	}
	If ($HTML)
	{
		$rowdata = @()
	}
	
	If ($MSWord -or $PDF)
	{
		$ScriptInformation += @{ Data = "Server"; Value = "$($XSHostMemory.Server)"; }
		$ScriptInformation += @{ Data = "VMs"; Value = "$($XSHostMemory.VMs)"; }
		$XSHostMemory.VMTexts | ForEach-Object { $ScriptInformation += @{ Data = ""; Value = "$($_)"; } }
		$ScriptInformation += @{ Data = "Citrix Hypervisor"; Value = "$($XSHostMemory.XenServerMemory)"; }
		$ScriptInformation += @{ Data = "Control domain memory"; Value = "$($XSHostMemory.ControlDomainMemory)"; }
		$ScriptInformation += @{ Data = "Available memory"; Value = "$($XSHostMemory.AvailableMemory)"; }
		$ScriptInformation += @{ Data = "Total max memory"; Value = "$($XSHostMemory.TotalMaxMemory)"; }
	}
	If ($Text)
	{
		Line 3 "Server`t`t`t: " "$($XSHostMemory.Server)"
		Line 3 "VMs`t`t`t: " "$($XSHostMemory.VMs)"
		$XSHostMemory.VMTexts | ForEach-Object { Line 6 "  $($_)" }
		Line 3 "Citrix Hypervisor`t: " "$($XSHostMemory.XenServerMemory)"
		Line 3 "Control domain memory`t: " "$($XSHostMemory.ControlDomainMemory)"
		Line 3 "Available memory`t: " "$($XSHostMemory.AvailableMemory)"
		Line 3 "Total max memory`t: " "$($XSHostMemory.TotalMaxMemory)"
	}
	If ($HTML)
	{
		$columnHeaders = @("Server", ($htmlsilver -bor $htmlbold), "$($XSHostMemory.Server)", $htmlwhite)
		$rowdata += @(, ("VMs", ($htmlsilver -bor $htmlbold), "$($XSHostMemory.VMs)", $htmlwhite))
		$XSHostMemory.VMTexts | ForEach-Object { $rowdata += @(, ("", ($htmlsilver -bor $htmlbold), "$($_)", $htmlwhite)) }
		$rowdata += @(, ("Citrix Hypervisor", ($htmlsilver -bor $htmlbold), "$($XSHostMemory.XenServerMemory)", $htmlwhite))
		$rowdata += @(, ("Control domain memory", ($htmlsilver -bor $htmlbold), "$($XSHostMemory.ControlDomainMemory)", $htmlwhite))
		$rowdata += @(, ("Available memory", ($htmlsilver -bor $htmlbold), "$($XSHostMemory.AvailableMemory)", $htmlwhite))
		$rowdata += @(, ("Total max memory", ($htmlsilver -bor $htmlbold), "$($XSHostMemory.TotalMaxMemory)", $htmlwhite))
	}

	If ($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;
	
		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
	
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;
	
		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)
	
		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 0 ""
	}
	If ($HTML)
	{
		$msg = ""
		$columnWidths = @("150", "200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
		
}

Function OutputHostStorage
{
	Param([object]$XSHost)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host Storage"

	$XSHostStorages = @($Script:XSPoolStorages | Where-Object { $_.XSHostRef -like $XSHost.opaque_ref } )
	
	$storageCount = $XSHostStorages.Count

	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Storage"
	}
	If ($Text)
	{
		Line 2 "Storage"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Storage"
	}

	If ($storageCount -lt 1)
	{
		If ($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There is no storage configured for Host $XSHostName"
		}
		If ($Text)
		{
			Line 3 "There is no storage configured for Host $XSHostName"
			Line 0 ""
		}
		If ($HTML)
		{
			WriteHTMLLine 0 1 "There is no storage configured for Host $XSHostName"
		}
	}
	Else
	{
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Number of storages"; Value = "$storageCount"; }
		}
		If ($Text)
		{
			Line 3 "Number of storages: " "$storageCount"
			Line 0 ""
		}
		If ($HTML)
		{
			$columnHeaders = @("Number of storages", ($htmlsilver -bor $htmlbold), "$storageCount", $htmlwhite)
			$rowdata = @()
		}

		ForEach ($Item in $XSHostStorages)
		{
			If ($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "Name"; Value = $($item.Name); }
				$ScriptInformation += @{ Data = "     Description"; Value = $($item.Description); }
				$ScriptInformation += @{ Data = "     Type"; Value = $($item.Type); }
				$ScriptInformation += @{ Data = "     Shared"; Value = $($item.Shared); }
				$ScriptInformation += @{ Data = "     Usage"; Value = $($item.Usage); }
				$ScriptInformation += @{ Data = "     Size"; Value = $($item.Size); }
				$ScriptInformation += @{ Data = "     Virtual allocation"; Value = $($item.VirtualAllocation); }
			}
			If ($Text)
			{
				Line 3 "Name: " $($item.Name)
				Line 4 "Description`t`t: " $($item.Description)
				Line 4 "Type`t`t`t: " $($item.Type)
				Line 4 "Shared`t`t`t: " $($item.Shared)
				Line 4 "Usage`t`t`t: " $($item.Usage)
				Line 4 "Size`t`t`t: " $($item.Size)
				Line 4 "Virtual allocation`t: " $($item.VirtualAllocation)
				Line 0 ""
			}
			If ($HTML)
			{
				$rowdata += @(, ("Name", ($htmlsilver -bor $htmlbold), $($item.Name), ($htmlsilver -bor $htmlbold)))
				$rowdata += @(, ("     Description", ($htmlsilver -bor $htmlbold), $($item.Description), $htmlwhite))
				$rowdata += @(, ("     Type", ($htmlsilver -bor $htmlbold), $($item.Type), $htmlwhite))
				$rowdata += @(, ("     Shared", ($htmlsilver -bor $htmlbold), $($item.Shared), $htmlwhite))
				$rowdata += @(, ("     Usage", ($htmlsilver -bor $htmlbold), $($item.Usage), $htmlwhite))
				$rowdata += @(, ("     Size", ($htmlsilver -bor $htmlbold), $($item.Size), $htmlwhite))
				$rowdata += @(, ("     Virtual allocation", ($htmlsilver -bor $htmlbold), $($item.VirtualAllocation), $htmlwhite))
			}
		}
		
		If ($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data, Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 350;

			$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If ($Text)
		{
			Line 0 ""
		}
		If ($HTML)
		{
			$msg = ""
			$columnWidths = @("150", "350")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}


}

Function OutputHostNetworking
{
	Param([object]$XSHost)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host Networking"

	$XSHostNetworks = @($Script:XSPoolNetworks | Where-Object { $_.XSHostRef -like $XSHost.opaque_ref } )
	$nrNetworking = $XSHostNetworks.Count
	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Networking"
	}
	If ($Text)
	{
		Line 2 "Networking"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Networking"
	}

	If ($nrNetworking -lt 1)
	{
		If ($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There are no networks configured for Host $XSHostName"
		}
		If ($Text)
		{
			Line 3 "There are no Network networks configured for Host $XSHostName"
			Line 0 ""
		}
		If ($HTML)
		{
			WriteHTMLLine 0 1 "There are no networks configured for Host $XSHostName"
		}
	}
	Else
	{
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Number of networks"; Value = "$nrNetworking"; }
		}
		If ($Text)
		{
			Line 3 "Number of networks: " "$nrNetworking"
			Line 0 ""
		}
		If ($HTML)
		{
			$columnHeaders = @("Number of networks", ($htmlsilver -bor $htmlbold), "$nrNetworking", $htmlwhite)
			$rowdata = @()
		}

		ForEach ($Item in $XSHostNetworks)
		{
			If ($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "Name"; Value = $($item.Name); }
				$ScriptInformation += @{ Data = "     Description"; Value = $($item.Description); }
				$ScriptInformation += @{ Data = "     NIC"; Value = $($item.NIC); }
				$ScriptInformation += @{ Data = "     VLAN"; Value = $($item.VLAN); }
				$ScriptInformation += @{ Data = "     Auto"; Value = $($item.Auto); }
				$ScriptInformation += @{ Data = "     Link Status"; Value = $($item.LinkStatus); }
				$ScriptInformation += @{ Data = "     MAC"; Value = $($item.MAC); }
				$ScriptInformation += @{ Data = "     MTU"; Value = $($item.MTU); }
				$ScriptInformation += @{ Data = "     SR-IOV"; Value = $($item.SRIOV); }
			}
			If ($Text)
			{
				Line 3 "Name: " $($item.Name)
				Line 4 "Description`t: " $($item.Description)
				Line 4 "NIC`t`t: " $($item.NIC)
				Line 4 "VLAN`t`t: " $($item.VLAN)
				Line 4 "Auto`t`t: " $($item.Auto)
				Line 4 "Link Status`t: " $($item.LinkStatus)
				Line 4 "MAC`t`t: " $($item.MAC)
				Line 4 "MTU`t`t: " $($item.MTU)
				Line 4 "SR-IOV`t`t: " $($item.SRIOV)
				Line 0 ""
			}
			If ($HTML)
			{
				$rowdata += @(, ("Name", ($htmlsilver -bor $htmlbold), $($item.Name), ($htmlsilver -bor $htmlbold)))
				$rowdata += @(, ("     Description", ($htmlsilver -bor $htmlbold), $($item.Description), $htmlwhite))
				$rowdata += @(, ("     NIC", ($htmlsilver -bor $htmlbold), $($item.NIC), $htmlwhite))
				$rowdata += @(, ("     VLAN", ($htmlsilver -bor $htmlbold), $($item.VLAN), $htmlwhite))
				$rowdata += @(, ("     Auto", ($htmlsilver -bor $htmlbold), $($item.Auto), $htmlwhite))
				$rowdata += @(, ("     Link Status", ($htmlsilver -bor $htmlbold), $($item.LinkStatus), $htmlwhite))
				$rowdata += @(, ("     MAC", ($htmlsilver -bor $htmlbold), $($item.MAC), $htmlwhite))
				$rowdata += @(, ("     MTU", ($htmlsilver -bor $htmlbold), $($item.MTU), $htmlwhite))
				$rowdata += @(, ("     SR-IOV", ($htmlsilver -bor $htmlbold), $($item.SRIOV), $htmlwhite))
			}
		}
		
		If ($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data, Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 175;

			$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If ($Text)
		{
			Line 0 ""
		}
		If ($HTML)
		{
			$msg = ""
			$columnWidths = @("150", "200")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
}

Function OutputHostNICs
{
	Param([object]$XSHost)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host NICs"

	$XSNICSs = @($XSHost.PIFs | Get-XenPIF | Where-Object { $_.physical -like $true -or -Not [String]::IsNullOrEmpty($_.bond_master_of) } | Sort-Object -Property device)
	$XSHostName = $XSHost.Name_Label
	$nrNICSs = $XSNICSs.Count
	$nics = @()
	If ($nrNICSs -ge 1)
	{
		ForEach ($Item in $XSNICSs)
		{
			$pifMetrics = $item.metrics | Get-XenPIFMetrics
			<#
			if (-Not [String]::IsNullOrEmpty($Item.bond_master_of))   {
				$nicBond = $Item.bond_master_of | Get-XenBond
				$nicSlaves = $XSNICSs | Where-Object   { $_.opaque_ref -in $nicBond.slaves}
			}
			#>
			if ($pifMetrics.carrier -like $true)
			{
				$linkStatus = "Connected"
				If ([String]::IsNullOrEmpty($($pifMetrics.duplex)) -or $pifMetrics.duplex -like $false)
				{
					$nicDuplex = "Full"
				}
				Else
				{
					$nicDuplex = "Half"
				}
				If ($pifMetrics.speed -gt 0)
				{
					$nicSpeed = '{0} Mbit/s' -f $pifMetrics.speed
				}
				Else
				{
					$nicSpeed = "-"
				}
			}
			else
			{
				$linkStatus = "Disconnected"
				$nicDuplex = "-"
				$nicSpeed = "-"
			}
			If ("fcoe" -in $item.capabilities)
			{
				$fcoeCapable = "Yes"
			}
			Else
			{
				$fcoeCapable = "No"
			}

			If ("sriov" -in $item.capabilities)
			{
				$sriovCapable = "Yes"
			}
			Else
			{
				$sriovCapable = "No"
			}

			$nics += $Item | Select-Object -Property `
			@{Name = 'Name'; Expression = { $_.device.Replace("eth", "NIC ") } },
			@{Name = 'DeviceID'; Expression = { $_.device } },
			MAC,
			@{Name = 'LinkStatus'; Expression = { $linkStatus } },
			@{Name = 'Speed'; Expression = { $nicSpeed } },
			@{Name = 'Duplex'; Expression = { $nicDuplex } },
			@{Name = 'Vendor'; Expression = { "$($pifMetrics.vendor_name)" } },
			@{Name = 'Device'; Expression = { "$($pifMetrics.device_name)" } },
			@{Name = 'PCIBusPath'; Expression = { "$($pifMetrics.pci_bus_path)" } },
			@{Name = 'FCoE'; Expression = { $fcoeCapable; } },
			@{Name = 'SRIOV'; Expression = { $sriovCapable; } }
		}
	}
	$nics = $nics | Sort-Object -Property DeviceID
	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "NICs"
	}
	If ($Text)
	{
		Line 2 "NICs"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "NICs"
	}

	If ($nrNICSs -lt 1)
	{

		If ($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There are no NICs configured for Host $XSHostName"
		}
		If ($Text)
		{
			Line 3 "There are no Network NICs configured for Host $XSHostName"
			Line 0 ""
		}
		If ($HTML)
		{
			WriteHTMLLine 0 1 "There are no NICs configured for Host $XSHostName"
		}
	}
	Else
	{
		
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Number of NICs"; Value = "$nrNICSs"; }
		}
		If ($Text)
		{
			Line 3 "Number of NICs: " "$nrNICSs"
			Line 0 ""
		}
		If ($HTML)
		{
			$columnHeaders = @("Number of NICs", ($htmlsilver -bor $htmlbold), "$nrNICSs", $htmlwhite)
			$rowdata = @()
		}

		ForEach ($Item in $nics)
		{
			If ($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "NIC"; Value = $($Item.Name); }
				$ScriptInformation += @{ Data = "     MAC"; Value = $($item.MAC); }
				$ScriptInformation += @{ Data = "     Link Status"; Value = $($item.LinkStatus); }
				$ScriptInformation += @{ Data = "     Speed"; Value = $($item.Speed); }
				$ScriptInformation += @{ Data = "     Duplex"; Value = $($item.Duplex); }
				$ScriptInformation += @{ Data = "     Vendor"; Value = $($item.Vendor); }
				$ScriptInformation += @{ Data = "     Device"; Value = $($item.Device); }
				$ScriptInformation += @{ Data = "     PCI Bus Path"; Value = $($item.PCIBusPath); }
				$ScriptInformation += @{ Data = "     FCoE Capable"; Value = $($item.FCoE); }
				$ScriptInformation += @{ Data = "     SR-IOV Capable"; Value = $($item.SRIOV); }
			}
			If ($Text)
			{
				Line 3 "NIC: " $($Item.Name)
				Line 4 "MAC`t`t: " $($item.MAC)
				Line 4 "Link Status`t: " $($item.LinkStatus)
				Line 4 "Speed`t`t: " $($item.Speed)
				Line 4 "Duplex`t`t: " $($item.Duplex)
				Line 4 "Vendor`t`t: " $($item.Vendor)
				Line 4 "Device`t`t: " $($item.Device)
				Line 4 "PCI Bus Path`t: " $($item.PCIBusPath)
				Line 4 "FCoE Capable`t: " $($item.FCoE)
				Line 4 "SR-IOV Capable`t: " $($item.SRIOV)
				Line 0 ""
			}
			If ($HTML)
			{
				$rowdata += @(, ("NIC", ($htmlsilver -bor $htmlbold), $($Item.Name), ($htmlsilver -bor $htmlbold)))
				$rowdata += @(, ("     MAC", ($htmlsilver -bor $htmlbold), $($item.MAC), $htmlwhite))
				$rowdata += @(, ("     Link Status", ($htmlsilver -bor $htmlbold), $($item.LinkStatus), $htmlwhite))
				$rowdata += @(, ("     Speed", ($htmlsilver -bor $htmlbold), $($item.Speed), $htmlwhite))
				$rowdata += @(, ("     Duplex", ($htmlsilver -bor $htmlbold), $($item.Duplex), $htmlwhite))
				$rowdata += @(, ("     Vendor", ($htmlsilver -bor $htmlbold), $($item.Vendor), $htmlwhite))
				$rowdata += @(, ("     Device", ($htmlsilver -bor $htmlbold), $($item.Device), $htmlwhite))
				$rowdata += @(, ("     PCI Bus Path", ($htmlsilver -bor $htmlbold), $($item.PCIBusPath), $htmlwhite))
				$rowdata += @(, ("     FCoE Capable", ($htmlsilver -bor $htmlbold), $($item.FCoE), $htmlwhite))
				$rowdata += @(, ("     SR-IOV Capable", ($htmlsilver -bor $htmlbold), $($item.SRIOV), $htmlwhite))
			}
		}
		
		If ($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data, Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If ($Text)
		{
			Line 0 ""
		}
		If ($HTML)
		{
			$msg = ""
			$columnWidths = @("150", "250")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}

}

Function OutputHostGPU
{
	Param([object]$XSHost)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput Host GPU"
	$pGPUs = @($XSHost.PGPUs | Get-XenPGPU)
	$XSHostName = $XSHost.Name_Label
	$nrGPUs = $pGPUs.Count
	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "GPU"
	}
	If ($Text)
	{
		Line 2 "GPU"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "GPU"
	}

	If ($nrGPUs -lt 1)
	{

		If ($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There are no GPU's configured for Host $XSHostName"
		}
		If ($Text)
		{
			Line 3 "There are no GPU's configured for Host $XSHostName"
			Line 0 ""
		}
		If ($HTML)
		{
			WriteHTMLLine 0 1 "There are no GPU's configured for Host $XSHostName"
		}
	}
	Else
	{
		
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Number of GPU's Installed"; Value = "$nrGPUs"; }
		}
		If ($Text)
		{
			Line 3 "Number of GPU's Installed: " "$nrGPUs"
		}
		If ($HTML)
		{
			$columnHeaders = @("Number of GPU's Installed", ($htmlsilver -bor $htmlbold), "$nrGPUs", $htmlwhite)
			$rowdata = @()
		}

		ForEach ($Item in $pGPUs)
		{
			$gpuGroup = $item.GPU_group | Get-XenGPUGroup
			$allocation = "$(($gpuGroup.allocation_algorithm).ToString().Replace("depth_first","Maximum density").Replace("breadth_first","Maximum performance")) ($($gpuGroup.allocation_algorithm))"
			$gpuTypes = $gpuGroup.supported_VGPU_types | Get-XenVGPUType | Sort-Object -Property framebuffer_size, model_name
			$gpuTypesText = ""
			ForEach ($type in $gpuTypes)
			{
				$gpuTypesLine = "$($type.model_name.ToString().Replace("NVIDIA",$null).Trim()) / Framebuffer:$(Convert-SizeToString -size $type.framebuffer_size -Decimal 1) "
				If ($type.opaque_ref -in $Item.enabled_VGPU_types.opaque_ref)
				{
					$gpuTypesLine += "/ Enabled"
				}
				Else
				{
					$gpuTypesLine += "/ Disabled"
				}
				$gpuTypesText += "$gpuTypesLine`r`n"
			}
			If ([string]::IsNullOrEmpty($gpuTypesText))
			{
				$gpuTypesText = "none"
			}
			If ([String]::IsNullOrEmpty($($Item.is_system_display_device)))
			{
				$primaryAdapter = "False"
			}
			Else
			{
				$primaryAdapter = "$($Item.is_system_display_device)"
			}
			
			If ($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = ""; Value = ""; }
				$ScriptInformation += @{ Data = "Name"; Value = $($gpuGroup.name_label); }
				$ScriptInformation += @{ Data = "vGPU allocation"; Value = $($allocation); }
				$ScriptInformation += @{ Data = "Primary host display adapter"; Value = $($primaryAdapter); }
				$ScriptInformation += @{ Data = "vGPU Pofiles"; Value = $($gpuTypesText); }

			}
			If ($Text)
			{
				Line 3 "" ""
				Line 3 "Name`t`t`t`t: " $($gpuGroup.name_label)
				Line 3 "vGPU allocation`t`t`t: " $($allocation)
				Line 3 "Primary host display adapter`t: " $($primaryAdapter)
				Line 3 "vGPU profiles`t`t`t: " $($gpuTypesText)
				Line 0 ""
			}
			If ($HTML)
			{
				$rowdata += @(, ("", ($htmlsilver -bor $htmlbold), "", ($htmlsilver -bor $htmlbold)))
				$rowdata += @(, ("Name", ($htmlsilver -bor $htmlbold), $($gpuGroup.name_label), $htmlwhite))
				$rowdata += @(, ("vGPU allocation", ($htmlsilver -bor $htmlbold), $($allocation), $htmlwhite))
				$rowdata += @(, ("Primary host display adapter", ($htmlsilver -bor $htmlbold), $($primaryAdapter), $htmlwhite))
				$rowdata += @(, ("vGPU profiles", ($htmlsilver -bor $htmlbold), $($gpuTypesText.Replace("`r`n", "<br>")), $htmlwhite))
			}
		}
		
		If ($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data, Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 350;

			$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If ($Text)
		{
			Line 0 ""
		}
		If ($HTML)
		{
			$msg = ""
			$columnWidths = @("150", "350")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
}
#endregion

#region VMs
Function ProcessVMs
{
	Write-Verbose "$(Get-Date -Format G): Process Virtual Machines"
	If ($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Virtual Machines"
	}
	If ($Text)
	{
		Line 0 ""
		Line 0 "Virtual Machines"
		Line 0 ""
	}
	If ($HTML)
	{
		WriteHTMLLine 1 0 "Virtual Machines"
	}
	
	$VMFirst = $True
	ForEach ($VMName in $Script:VMNames)
	{
		Write-Verbose "$(Get-Date -Format G): `tOutput VM $($VMName.name_label)"
		$VM = Get-XenVM -Name $VMName.name_label
		$VMMetrics = $VM.guest_metrics | Get-XenVMGuestMetrics
		if ([String]::IsNullOrEmpty($VMMetrics) -or [String]::IsNullOrEmpty($($VMMetrics.os_version)) -or $VMMetrics.os_version.Count -lt 1 -or [String]::IsNullOrEmpty($($VMMetrics.os_version.name)))
		{
			$VMOSName = "Unknown"
		}
		else 
		{
			$VMOSName = $VMMetrics.os_version.name
		}

		$VMHostData = $VM.resident_on | Get-XenHost
		if ([String]::IsNullOrEmpty($VMHostData) -or [String]::IsNullOrEmpty($($VMHostData.name_label)))
		{
			If ($VM.power_state -ne "Running")
			{
				$VMHost = "VM not running"
			}
			Else
			{
				$VMHost = "N/A"
			}
			
		}
		else 
		{
			$VMHost = $VMHostData.name_label
		}

		<#If (!$?)
		{
			If ($VM.power_state -ne "Running")
			{
				$VMHost = "VM not running"
			}
			Else
			{
				$VMHost = "N/A"
			}
		}#>
		OutputVM $VM $VMOSName $VMHost $VMFirst
		OutputVMCustomFields $VM
		OutputVMCPU $VM
		OutputVMBootOptions $VM
		OutputVMStartOptions $VM
		OutputVMAlerts $VM
		OutputVMHomeServer $VM
		OutputVMGPU $VM
		OutputVMAdvancedOptions $VM
		OutputVMStorage $VM $VMHost
		OutputVMNIC $VM
		OutputVMSnapshots $VM
		$VMFirst = $False
	}
}

Function OutputVM
{
	Param([object]$VM, [string]$VMOSName, [string]$VMHost, [bool]$VMFirst)
	
	Write-Verbose "$(Get-Date -Format G): `t`tOutput VM General"
	If ($VMOSName -ne "Unknown")
	{
		#remove the pipe symbol from the $VMOSName variable
		$pos = -1
		$pos = $VMOSName.IndexOf('|')
		If ($pos -gt -1)
		{
			$VMOSName = $VMOSName.SubString(0, $pos)
		}
	}

	If ($VM.memory_dynamic_max -ne $VM.memory_dynamic_min)
	{
		$xenVMDynamicMemory = $true
	}
	Else
	{
		$xenVMDynamicMemory = $false
	}

	$xenVmMem = Convert-SizeToString -size $VM.memory_static_max -Decimal 1
	$xenVmMemMax = Convert-SizeToString -size $VM.memory_dynamic_max -Decimal 1
	$xenVmMemMin = Convert-SizeToString -size $VM.memory_dynamic_min -Decimal 1
	$bootorder = $vm.HVM_boot_params["order"].Replace("c", "Disk;").Replace("d", "DVD-drive;").Replace("n", "Network;").TrimEnd(";").Split(";")
	for ($c = 1 ; $c -le $bootorder.count ; $c++)
	{
		$bootorder[$c - 1] = "[$c] $($bootorder[$c-1])"
	}

	If ($null -eq $($vm.platform["cores-per-socket"]))
	{
		$vCPUcoreText = "1 core"
	}
	ElseIf ($vm.platform["cores-per-socket"] -gt 1)
	{
		$vCPUcoreText = "$($vm.platform["cores-per-socket"]) cores"
	}
	Else
	{
		$vCPUcoreText = "$($vm.platform["cores-per-socket"]) core"
	}

	try
	{
		$sockets = $([int]$VM.VCPUs_max / [int]$vm.platform["cores-per-socket"])
		If ($sockets -gt 1)
		{
			$vCPUText = "$($VM.VCPUs_max) ($sockets sockets with $vCPUcoreText each)"
		}
		Else
		{
			$vCPUText = "$($VM.VCPUs_max) ($sockets socket with $vCPUcoreText)"
		}
	}
	catch
	{
		$vCPUText = "$($VM.VCPUs_max) ($vCPUcoreText per socket)"
	}

	If ([String]::IsNullOrEmpty($($VM.HVM_boot_params["secureboot"])))
	{
		$xenVMSecureBoot = "False"
	}
	Else
	{
		$xenVMSecureBoot = $VM.HVM_boot_params["secureboot"]
	}

	If ($MSWord -or $PDF)
	{
		If ($VMFirst -eq $False)
		{
			#Put the 2nd VM on, on a new page
			$Selection.InsertNewPage()
		}
		
		WriteWordLine 2 0 "VM: $($VM.name_label)"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "VM name"; Value = $($VM.name_label); }
		$ScriptInformation += @{ Data = "Xen host name"; Value = $VMHost; }
		$ScriptInformation += @{ Data = "VM Operating System"; Value = $VMOSName; }
		$ScriptInformation += @{ Data = "Number of vCPUs"; Value = $VCPUText; }

		If ($xenVMDynamicMemory)
		{
			$ScriptInformation += @{ Data = "Dynamic Memory"; Value = "True (DEPRECATED!)"; }
			$ScriptInformation += @{ Data = "Minimum Memory"; Value = $xenVmMemMin; }
			$ScriptInformation += @{ Data = "Maximum Memory"; Value = $xenVmMemMax; }
		}
		Else
		{
			$ScriptInformation += @{ Data = "Memory"; Value = $xenVmMem; }
		}
		$ScriptInformation += @{ Data = "Boot order"; Value = $($bootorder -join ", "); }
		$ScriptInformation += @{ Data = "Boot mode"; Value = $VM.HVM_boot_params["firmware"].ToUpper(); }
		$ScriptInformation += @{ Data = "Secure boot"; Value = $xenVMSecureBoot; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 1 "VM name: " $($VM.name_label)
		Line 2 "Xen host name`t`t: " $VMHost
		Line 2 "VM Operating System`t: " $VMOSName
		Line 2 "Number of vCPUs`t`t: " $VCPUText

		If ($xenVMDynamicMemory)
		{
			Line 2 "Dynamic Memory`t`t: " "True (DEPRECATED!)"
			Line 2 "Minimum Memory`t`t: " $xenVmMemMin
			Line 2 "Maximum Memory`t`t: " $xenVmMemMax
		}
		Else
		{
			Line 2 "Memory`t`t`t: " $xenVmMem
		}
		Line 2 "Boot order`t`t: " $($bootorder -join ", ")
		Line 2 "Boot mode`t`t: " $VM.HVM_boot_params["firmware"].ToUpper()
		Line 2 "Secure boot`t`t: " $xenVMSecureBoot

		Line 0 ""
	}
	If ($HTML)
	{
		WriteHTMLLine 2 0 "VM: $($VM.name_label)"
		$rowdata = @()
		$columnHeaders = @("VM name", ($htmlsilver -bor $htmlbold), $($VM.name_label), $htmlwhite)
		$rowdata += @(, ('Xen host name', ($htmlsilver -bor $htmlbold), $VMHost, $htmlwhite))
		$rowdata += @(, ('VM Operating System', ($htmlsilver -bor $htmlbold), $VMOSName, $htmlwhite))
		$rowdata += @(, ('Number of vCPUs', ($htmlsilver -bor $htmlbold), $VCPUText, $htmlwhite))

		If ($xenVMDynamicMemory)
		{
			$rowdata += @(, ('Dynamic Memory', ($htmlsilver -bor $htmlbold), "True (DEPRECATED!)", $htmlwhite))
			$rowdata += @(, ('Minimum Memory', ($htmlsilver -bor $htmlbold), $xenVmMemMin, $htmlwhite))
			$rowdata += @(, ('Maximum Memory', ($htmlsilver -bor $htmlbold), $xenVmMemMax, $htmlwhite))
		}
		Else
		{
			$rowdata += @(, ('Memory', ($htmlsilver -bor $htmlbold), $xenVmMem, $htmlwhite))
		}
		$rowdata += @(, ('Boot order', ($htmlsilver -bor $htmlbold), $($bootorder -join ", "), $htmlwhite))
		$rowdata += @(, ('Boot mode', ($htmlsilver -bor $htmlbold), $VM.HVM_boot_params["firmware"].ToUpper(), $htmlwhite))
		$rowdata += @(, ('Secure boot', ($htmlsilver -bor $htmlbold), $xenVMSecureBoot, $htmlwhite))

		$msg = ""
		$columnWidths = @("150", "250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputVMCustomFields
{
	Param([object] $VM)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput VM Custom Fields"

	$CustomFields = $CustomFields = Get-XSCustomFields $($vm.other_config)
	
	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Custom Fields"
	}
	If ($Text)
	{
		Line 2 "Custom Fields"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Custom Fields"
	}
	
	If ([String]::IsNullOrEmpty($CustomFields) -or $CustomFields.Count -eq 0)
	{
		$VMName = $VM.Name_Label

		If ($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There are no Custom Fields for VM $VMName"
		}
		If ($Text)
		{
			Line 3 "There are no Custom Fields for VM $VMName"
			Line 0 ""
		}
		If ($HTML)
		{
			WriteHTMLLine 0 1 "There are no Custom Fields for VM $VMName"
		}
	}
	Else
	{
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
		}
		If ($Text)
		{
			#nothing
		}
		If ($HTML)
		{
			$rowdata = @()
		}

		[int]$cnt = -1
		ForEach ($Item in $CustomFields)
		{
			$cnt++
			If ($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = $($Item.Name); Value = $Item.Value; }
			}
			If ($Text)
			{
				Line 3 "$($Item.Name): " $Item.Value
			}
			If ($HTML)
			{
				If ($cnt -eq 0)
				{
					$columnHeaders = @($($Item.Name), ($htmlsilver -bor $htmlbold), $Item.Value, $htmlwhite)
				}
				Else
				{
					$rowdata += @(, ($($Item.Name), ($htmlsilver -bor $htmlbold), $Item.Value, $htmlwhite))
				}
			}
		}
		
		If ($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data, Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 250;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If ($Text)
		{
			Line 0 ""
		}
		If ($HTML)
		{
			$msg = ""
			$columnWidths = @("250", "250")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
}

Function OutputVMCPU
{
	Param([object] $VM)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput VM CPU"

	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "CPU"
	}
	If ($Text)
	{
		Line 2 "CPU"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "CPU"
	}
	#from the XS team
	#the topology is in the VM's platform property, key cores-per-socket
	#VCPUs_at_startup	8
	#VCPUs_max		8

	If ($null -eq $($vm.platform["cores-per-socket"]))
	{
		$vCPUcoreText = "1 core"
	}
	ElseIf ($vm.platform["cores-per-socket"] -gt 1)
	{
		$vCPUcoreText = "$($vm.platform["cores-per-socket"]) cores"
	}
	Else
	{
		$vCPUcoreText = "$($vm.platform["cores-per-socket"]) core"
	}

	try
	{
		$sockets = $([int]$VM.VCPUs_max / [int]$vm.platform["cores-per-socket"])
		If ($sockets -gt 1)
		{
			$vCPUText = "$sockets sockets with $vCPUcoreText each"
		}
		Else
		{
			$vCPUText = "$sockets socket with $vCPUcoreText"
		}
	}
	catch
	{
		$vCPUText = "$($VM.VCPUs_max) ($vCPUcoreText per socket)"
	}
	
	#VCPUs_params		{[weight, 256]}
	#1 = lowest - first slider tick
	#4 = second slider tick
	#16 = third slider tick
	#64 = fourth slider tick
	#256 = normal - fifth slider tick
	#1024 = sixth slider tick
	#4096 = seventh slider tick
	#16384 = eighth slider tick
	#65535 = highest - ninth slider tick
	
	If ($vm.VCPUs_params["weight"] -gt 1)
	{
		$tmp = $vm.VCPUs_params["weight"]
		
		If ($tmp -eq 1)
		{
			$vCPUPriority = "Lowest"
		}
		ElseIf ($tmp -eq 4)
		{
			$vCPUPriority = "4 (Second tick on the slider)"
		}
		ElseIf ($tmp -eq 16)
		{
			$vCPUPriority = "16 (third tick on the slider)"
		}
		ElseIf ($tmp -eq 64)
		{
			$vCPUPriority = "64 (fourth tick on the slider)"
		}
		ElseIf ($tmp -eq 256)
		{
			$vCPUPriority = "Normal"
		}
		ElseIf ($tmp -eq 1024)
		{
			$vCPUPriority = "1024 (sixth tick on the slider)"
		}
		ElseIf ($tmp -eq 4096)
		{
			$vCPUPriority = "4096 (seventh tick on the slider)"
		}
		ElseIf ($tmp -eq 16384)
		{
			$vCPUPriority = "16384 (eigth tick on the slider)"
		}
		ElseIf ($tmp -eq 65535)
		{
			$vCPUPriority = "Highest"
		}
		Else
		{
			$vCPUPriority = $tmp.ToString()
		}
	}
	Else
	{
		#for some VMs, the default value is not saved until the user makes a manual change
		$vCPUPriority = "Normal"
	}

	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Number of vCPUs"; Value = "$($VM.VCPUs_max)"; }
		$ScriptInformation += @{ Data = "Topology"; Value = $vCPUText; }
		$ScriptInformation += @{ Data = "vCPU priority for this virtual machine"; Value = $vCPUPriority; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 3 "Number of vCPUs                       : " "$($VM.VCPUs_max)"
		Line 3 "Topology                              : " $vCPUText
		Line 3 "vCPU priority for this virtual machine: " $vCPUPriority
		Line 0 ""
	}
	If ($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Number of vCPUs", ($htmlsilver -bor $htmlbold), "$($VM.VCPUs_max)", $htmlwhite)
		$rowdata += @(, ("Topology", ($htmlsilver -bor $htmlbold), $vCPUText, $htmlwhite))
		$rowdata += @(, ("vCPU priority for this virtual machine", ($htmlsilver -bor $htmlbold), $vCPUPriority, $htmlwhite))

		$msg = ""
		$columnWidths = @("200", "150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLIne 0 0 ""
	}
}

Function OutputVMBootOptions
{
	Param([object] $VM)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput VM Boot Options"

	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Boot Options"
	}
	If ($Text)
	{
		Line 2 "Boot Options"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Boot Options"
	}

	$bootorder = $vm.HVM_boot_params["order"].Replace("c", "Disk;").Replace("d", "DVD-drive;").Replace("n", "Network;").TrimEnd(";").Split(";")
	for ($c = 1 ; $c -le $bootorder.count ; $c++)
	{
		$bootorder[$c - 1] = "[$c] $($bootorder[$c-1])"
	}

	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Boot order"; Value = $($bootorder -join ", "); }

		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 3 "Boot order: " $($bootorder -join ", ")
		Line 0 ""
	}
	If ($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Boot order", ($htmlsilver -bor $htmlbold), $($bootorder -join ", "), $htmlwhite)

		$msg = ""
		$columnWidths = @("150", "250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
	
}

Function OutputVMStartOptions
{
	Param([object] $VM)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput VM Start Options"

	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Start Options"
	}
	If ($Text)
	{
		Line 2 "Start Options"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Start Options"
	}
	
	$StartOrder = $VM.order.ToString()
	$StartNextVMAfter = $VM.start_delay
	
	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Start order"; Value = $StartOrder; }
		$ScriptInformation += @{ Data = "Attempt to start next VM after"; Value = "$($StartNextVMAfter) seconds"; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 175;
		$Table.Columns.Item(2).Width = 100;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 3 "Start order                   : " $StartOrder
		Line 3 "Attempt to start next VM after: " "$($StartNextVMAfter) seconds"
		Line 0 ""
	}
	If ($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Start order", ($htmlsilver -bor $htmlbold), $StartOrder, $htmlwhite)
		$rowdata += @(, ("Attempt to start next VM after", ($htmlsilver -bor $htmlbold), "$($StartNextVMAfter) seconds", $htmlwhite))

		$msg = ""
		$columnWidths = @("200", "100")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLIne 0 0 ""
	}
}

Function OutputVMAlerts
{
	Param([object] $VM)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput VM Alerts"

	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Alerts"
	}
	If ($Text)
	{
		Line 2 "Alerts"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Alerts"
	}
	
	[int32]$AlertRepeatInterval = 0

	$GenerateVMCPUUsageAlerts = "Not selected"
	[double]$WhenVMCPUUsageExceeds = 0
	[int32]$WhenVMCPUForLongerThan = 0

	$GenerateVMNetworkUsageAlerts = "Not selected"
	[int32]$WhenVMNetworkUsageExceeds = 0
	[int32]$WhenVMNetworkForLongerThan = 0

	$GenerateVMDiskUsageAlerts = "Not selected"
	[int32]$WhenVMMemUsageExceeds = 0
	[int32]$WhenVMMemForLongerThan = 0

	$OtherConfig = ($VM | Get-XenVMProperty -XenProperty OtherConfig -EA 0)

	If ($OtherConfig.ContainsKey("perfmon"))
	{
		[xml]$XML = $OtherConfig.perfmon
			
		ForEach ($Alert in $XML.config.variable)
		{
			If ($Alert.Name.Value -eq "cpu_usage")
			{
				$GenerateVMCPUUsageAlerts = "Selected"
				[double]$tmp = $Alert.alarm_trigger_level.Value
				$WhenVMCPUUsageExceeds = $tmp * 100
				$WhenVMCPUForLongerThan = $Alert.alarm_trigger_period.Value / 60
				$AlertRepeatInterval = $Alert.alarm_auto_inhibit_period.Value / 60
			}
			ElseIf ($Alert.Name.Value -eq "network_usage")
			{
				$GenerateVMNetworkUsageAlerts = "Selected"
				$WhenVMNetworkUsageExceeds = $Alert.alarm_trigger_level.Value / 1024
				$WhenVMNetworkForLongerThan = $Alert.alarm_trigger_period.Value / 60
				$AlertRepeatInterval = $Alert.alarm_auto_inhibit_period.Value / 60
			}
			ElseIf ($Alert.Name.Value -eq "disk_usage")
			{
				$GenerateVMDiskUsageAlerts = "Selected"
				$WhenVMMemUsageExceeds = $Alert.alarm_trigger_level.Value / 1024
				$WhenVMMemForLongerThan = $Alert.alarm_trigger_period.Value / 60
				$AlertRepeatInterval = $Alert.alarm_auto_inhibit_period.Value / 60
			}
		}
	}

	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		If ($GenerateVMCPUUsageAlerts -eq "Selected" -or
			$GenerateVMNetworkUsageAlerts -eq "Selected" -or
			$GenerateVMDiskUsageAlerts -eq "Selected")
		{
			$ScriptInformation += @{ Data = "Alert repeat interval"; Value = "$($AlertRepeatInterval) minutes"; }
		}
		Else
		{
			$ScriptInformation += @{ Data = "Alert repeat interval"; Value = "Not Set"; }
		}
		$ScriptInformation += @{ Data = "Generate CPU usage alerts"; Value = $GenerateVMCPUUsageAlerts; }
		If ($GenerateVMCPUUsageAlerts -eq "Selected")
		{
			$ScriptInformation += @{ Data = "     When CPU usage exceeds"; Value = "$($WhenVMCPUUsageExceeds) %"; }
			$ScriptInformation += @{ Data = "     For longer than"; Value = "$($WhenVMCPUForLongerThan) minutes"; }
		}
		$ScriptInformation += @{ Data = "Generate network usage alerts"; Value = $GenerateVMNetworkUsageAlerts; }
		If ($GenerateVMNetworkUsageAlerts -eq "Selected")
		{
			$ScriptInformation += @{ Data = "     When network usage exceeds"; Value = "$($WhenVMNetworkUsageExceeds) KB/s"; }
			$ScriptInformation += @{ Data = "     For longer than"; Value = "$($WhenVMNetworkForLongerThan) minutes"; }
		}
		$ScriptInformation += @{ Data = "Generate Disk usage alerts"; Value = $GenerateVMDiskUsageAlerts; }
		If ($GenerateVMDiskUsageAlerts -eq "Selected")
		{
			$ScriptInformation += @{ Data = "     When Disk usage exceeds"; Value = "$($WhenVMMemUsageExceeds) KB/s"; }
			$ScriptInformation += @{ Data = "     For longer than"; Value = "$($WhenVMMemForLongerThan) minutes"; }
		}

		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 175;
		$Table.Columns.Item(2).Width = 100;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		If ($GenerateVMCPUUsageAlerts -eq "Selected" -or
			$GenerateVMNetworkUsageAlerts -eq "Selected" -or
			$GenerateVMDiskUsageAlerts -eq "Selected")
		{
			Line 3 "Alert repeat interval             : $($AlertRepeatInterval) minutes"
		}
		Else
		{
			Line 3 "Alert repeat interval             : Not Set"
		}
		Line 3 "Generate CPU usage alerts         : " $GenerateVMCPUUsageAlerts
		If ($GenerateVMCPUUsageAlerts -eq "Selected")
		{
			Line 4 "When CPU usage exceeds    : " "$($WhenVMCPUUsageExceeds) %"
			Line 4 "For longer than           : " "$($WhenVMCPUForLongerThan) minutes"
		}
		Line 3 "Generate network usage alerts     : " $GenerateVMNetworkUsageAlerts
		If ($GenerateVMNetworkUsageAlerts -eq "Selected")
		{
			Line 4 "When network usage exceeds: " "$($WhenVMNetworkUsageExceeds) KB/s"
			Line 4 "For longer than           : " "$($WhenVMNetworkForLongerThan) minutes"
		}
		Line 3 "Generate Disk usage alerts        : " $GenerateVMDiskUsageAlerts
		If ($GenerateVMDiskUsageAlerts -eq "Selected")
		{
			Line 4 "When Disk usage exceeds   : " "$($WhenVMMemUsageExceeds) KB/s"
			Line 4 "For longer than           : " "$($WhenVMMemForLongerThan) minutes"
		}
		Line 0 ""
	}
	If ($HTML)
	{
		$rowdata = @()
		If ($GenerateVMCPUUsageAlerts -eq "Selected" -or
			$GenerateVMNetworkUsageAlerts -eq "Selected" -or
			$GenerateVMDiskUsageAlerts -eq "Selected")
		{
			$columnHeaders = @("Alert repeat interval", ($htmlsilver -bor $htmlbold), "$($AlertRepeatInterval) minutes", $htmlwhite)
		}
		Else
		{
			$columnHeaders = @("Alert repeat interval", ($htmlsilver -bor $htmlbold), "Not set", $htmlwhite)
		}
		$rowdata += @(, ("Generate CPU usage alerts", ($htmlsilver -bor $htmlbold), $GenerateVMCPUUsageAlerts, $htmlwhite))
		If ($GenerateVMCPUUsageAlerts -eq "Selected")
		{
			$rowdata += @(, ("     When CPU usage exceeds", ($htmlsilver -bor $htmlbold), "$($WhenVMCPUUsageExceeds) %", $htmlwhite))
			$rowdata += @(, ("     For longer than", ($htmlsilver -bor $htmlbold), "$($WhenVMCPUForLongerThan) minutes", $htmlwhite))
		}
		$rowdata += @(, ("Generate network usage alerts", ($htmlsilver -bor $htmlbold), $GenerateVMNetworkUsageAlerts , $htmlwhite))
		If ($GenerateVMNetworkUsageAlerts -eq "Selected")
		{
			$rowdata += @(, ("     When network usage exceeds", ($htmlsilver -bor $htmlbold), "$($WhenVMNetworkUsageExceeds) KB/s", $htmlwhite))
			$rowdata += @(, ("     For longer than", ($htmlsilver -bor $htmlbold), "$($WhenVMNetworkForLongerThan) minutes", $htmlwhite))
		}
		$rowdata += @(, ("Generate Disk usage alerts", ($htmlsilver -bor $htmlbold), $GenerateVMDiskUsageAlerts, $htmlwhite))
		If ($GenerateVMDiskUsageAlerts -eq "Selected")
		{
			$rowdata += @(, ("     When Disk usage exceeds", ($htmlsilver -bor $htmlbold), "$($WhenVMMemUsageExceeds) KB/s", $htmlwhite))
			$rowdata += @(, ("     For longer than", ($htmlsilver -bor $htmlbold), "$($WhenVMMemForLongerThan) minutes", $htmlwhite))
		}

		$msg = ""
		$columnWidths = @("200", "100")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLIne 0 0 ""
	}
}

Function OutputVMHomeServer
{
	Param([object] $VM)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput VM Home Server"

	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Home Server"
	}
	If ($Text)
	{
		Line 2 "Home Server"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Home Server"
	}
	
	If ($VM.affinity.opaque_ref -eq "OpaqueRef:NULL")
	{
		#there is no home server
		$HomeServerText = "Don't assign this VM a home server"
		$HomeServer = ""
		
		If ($VM.power_state -ne "Running")
		{
			$HomeServer = "VM's power state is $($vm.power_state). Unable to determine the running host."
		}
		Else
		{
			#find host currently on
			$HomeServerRef = $VM.resident_on.opaque_ref
			
			$results = Get-XenHost -Ref $HomeServerRef -EA 0
			
			If (!$?)
			{
				#unable to retrieve the running host
				$HomeServer = "Unable to retrieve the running host"
			}
			ElseIf ($Null -eq $results)
			{
				$HomeServer = "Unable to determine the running host"
			}
			Else
			{
				#we have the home server
				$HomeServer = "Currently running on host $($results.name_label)"
			}
		}
	}
	Else
	{
		$HomeServerRef = $VM.affinity.opaque_ref
		
		$results = Get-XenHost -Ref $HomeServerRef -EA 0
		
		If (!$?)
		{
			#unable to retrieve the home server
			$HomeServerText = "Unable to retrieve Home Server"
			$HomeServer = ""
		}
		Else
		{
			#we have the home server
			$HomeServerText = "Place the VM on this server"
			$HomeServer = $results.name_label
		}
	}

	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = $HomeServerText; Value = "$($HomeServer)"; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 175;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 3 "$($HomeServerText): " "$($HomeServer)"
		Line 0 ""
	}
	If ($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("$HomeServerText", ($htmlsilver -bor $htmlbold), "$($HomeServer)", $htmlwhite)

		$msg = ""
		$columnWidths = @("200", "250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLIne 0 0 ""
	}
}

Function OutputVMGPU
{
	Param([object] $VM)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput VM GPU"

	$xenVMGPUs = @($vm.VGPUs | Get-XenVGPU | ForEach-Object { Get-XenVGPUType -Ref $_.type })
	$VMName = $VM.Name_Label
	$nrGPUs = $xenVMGPUs.Count
	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "GPU"
	}
	If ($Text)
	{
		Line 2 "GPU"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "GPU"
	}

	If ($nrGPUs -lt 1)
	{

		If ($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There are no GPU's configured for VM $VMName"
		}
		If ($Text)
		{
			Line 3 "There are no GPU's configured for VM $VMName"
			Line 0 ""
		}
		If ($HTML)
		{
			WriteHTMLLine 0 1 "There are no GPU's configured for VM $VMName"
		}
	}
	Else
	{
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Number of GPU's"; Value = "$nrGPUs"; }
		}
		If ($Text)
		{
			Line 3 "Number of GPU's`t`t: " "$nrGPUs"
		}
		If ($HTML)
		{
			$columnHeaders = @("Number of GPU's", ($htmlsilver -bor $htmlbold), "$nrGPUs", $htmlwhite)
			$rowdata = @()
		}

		$gpuCount = 0
		ForEach ($Item in $xenVMGPUs)
		{
			$xenGPUFrameBufferSize = $(Convert-SizeToString -size $Item.framebuffer_size -Decimal 1)
			
			If ($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "Vendor name"; Value = $Item.vendor_name; }
				$ScriptInformation += @{ Data = "Model name"; Value = $Item.model_name; }
				$ScriptInformation += @{ Data = "Framebuffer size"; Value = $xenGPUFrameBufferSize; }
			}
			If ($Text)
			{
				Line 3 "Vendor name`t`t: " $Item.vendor_name
				Line 3 "Model name`t`t: " $Item.model_name
				Line 3 "Framebuffer size`t: " $xenGPUFrameBufferSize
			}
			If ($HTML)
			{
				$rowdata += @(, ("Vendor name", ($htmlsilver -bor $htmlbold), $Item.vendor_name, $htmlwhite))
				$rowdata += @(, ("Model name", ($htmlsilver -bor $htmlbold), $Item.model_name, $htmlwhite))
				$rowdata += @(, ("Framebuffer size", ($htmlsilver -bor $htmlbold), $xenGPUFrameBufferSize, $htmlwhite))
			}
			$gpuCount++
		}
		
		If ($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data, Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If ($Text)
		{
			Line 0 ""
		}
		If ($HTML)
		{
			$msg = ""
			$columnWidths = @("150", "250")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
}

Function OutputVMAdvancedOptions
{
	Param([object] $VM)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput VM Advanced Options"

	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Advanced Options"
	}
	If ($Text)
	{
		Line 2 "Advanced Options"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Advanced Options"
	}

	#from the XS team
	#the optimisation on the Advanced tab refers to the the VM's HVM_shadow_multiplier
	#XC uses 1 for general optimisation and 4 for citrix virtual apps
	
	$ShadowValue = $VM.HVM_shadow_multiplier

	If ($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
	}
	If ($Text)
	{
		#nothing
	}
	If ($HTML)
	{
		$rowdata = @()
	}
	
	If ($ShadowValue -eq 1)
	{
		If ($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Optimize for general use"; Value = ""; }
			$ScriptInformation += @{ Data = "     Shadow memory multiplier"; Value = $ShadowValue.ToString(); }
		}
		If ($Text)
		{
			Line 3 "Optimize for general use"
			Line 4 "Shadow memory multiplier: " $ShadowValue.ToString()
		}
		If ($HTML)
		{
			$columnHeaders = @("Optimize for general use", ($htmlsilver -bor $htmlbold), "", $htmlwhite)
			$rowdata += @(, ("     Shadow memory multiplier", ($htmlsilver -bor $htmlbold), $ShadowValue.ToString(), $htmlwhite))
		}
	}
	ElseIf ($ShadowValue -eq 4)
	{
		If ($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Optimize for Citrix Virtual Apps"; Value = ""; }
			$ScriptInformation += @{ Data = "     Shadow memory multiplier"; Value = $ShadowValue.ToString(); }
		}
		If ($Text)
		{
			Line 3 "Optimize for Citrix Virtual Apps"
			Line 4 "Shadow memory multiplier: " $ShadowValue.ToString()
		}
		If ($HTML)
		{
			$columnHeaders = @("Optimize for Citrix Virtual Apps", ($htmlsilver -bor $htmlbold), "", $htmlwhite)
			$rowdata += @(, ("     Shadow memory multiplier", ($htmlsilver -bor $htmlbold), $ShadowValue.ToString(), $htmlwhite))
		}
	}
	Else
	{
		If ($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Optimize manually (advanced use only)"; Value = ""; }
			$ScriptInformation += @{ Data = "     Shadow memory multiplier"; Value = $ShadowValue.ToString(); }
		}
		If ($Text)
		{
			Line 3 "Optimize manually (advanced use only)"
			Line 4 "Shadow memory multiplier: " $ShadowValue.ToString()
		}
		If ($HTML)
		{
			$columnHeaders = @("Optimize manually (advanced use only)", ($htmlsilver -bor $htmlbold), "", $htmlwhite)
			$rowdata += @(, ("     Shadow memory multiplier", ($htmlsilver -bor $htmlbold), $ShadowValue.ToString(), $htmlwhite))
		}
	}

	If ($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data, Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 100;

		$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If ($Text)
	{
		Line 0 ""
	}
	If ($HTML)
	{
		$msg = ""
		$columnWidths = @("200", "100")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputVMStorage
{
	Param([object]$VM, [string]$VMHost)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput VM Storage"
	$VMName = $VM.Name_Label
	$vbds = $Vm.VBDs | Get-XenVBD -EA 0
	
	$storages = @()
	ForEach ($item in $($vbds | Where-Object { $_.type -ne "CD" } | Sort-Object -Property userdevice))
	{
		$vdi = $item.VDI | Get-XenVDI -EA 0 | Where-Object { $_.is_a_snapshot -like $false }
		$sr = $vdi.SR | Get-XenSR -EA 0
		If ($vdi.read_only -like $true)
		{
			$readonly = "Yes"
		}
		Else
		{
			$readonly = "No"
		}
		If ($item.currently_attached -like $true)
		{
			$active = "Yes"
		}
		Else
		{
			$active = "No"
		}
		If ([String]::IsNullOrEmpty($($item.device)))
		{
			$device = "<unknown>"
		}
		Else
		{
			$device = '/dev/{0}' -f $item.device
		}
		If ([String]::IsNullOrEmpty($($item.qos_algorithm_params["class"]))) 
		{
			$priority = "0 (Lowest)"
		}
		ElseIf ($item.qos_algorithm_params["class"] -eq 7) 
		{
			$priority = "7 (Highest)"
		}
		Else 
		{
			$priority = "$($item.qos_algorithm_params["class"])"
		}
		$srText = '{0} on {1}' -f $sr.name_label, $VMHost
		$storages += $item | Select-Object -Property `
		@{Name = 'Position'; Expression = { $_.userdevice } },
		@{Name = 'Name'; Expression = { $vdi.name_label } },
		@{Name = 'Description'; Expression = { "$($vdi.name_description)" } },
		@{Name = 'SR'; Expression = { $srText } },
		@{Name = 'Size'; Expression = { $(Convert-SizeToString -Size $vdi.virtual_size -Decimal 0) } },
		@{Name = 'ReadOnly'; Expression = { $readonly } },
		@{Name = 'Priority'; Expression = { $priority } },
		@{Name = 'Active'; Expression = { $active } },
		@{Name = 'DevicePath'; Expression = { $device } }
	}
	$storages = @($storages | Sort-Object -Property Position, Name)
	$storageCount = $storages.Count

	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Storage"
	}
	If ($Text)
	{
		Line 2 "Storage"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Storage"
	}

	If ($storageCount -lt 1)
	{
		If ($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There is no storage configured for VM $VMName"
		}
		If ($Text)
		{
			Line 3 "There is no storage configured for VM $VMName"
			Line 0 ""
		}
		If ($HTML)
		{
			WriteHTMLLine 0 1 "There is no storage configured for VM $VMName"
		}
	}
	Else
	{
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Number of storages"; Value = "$storageCount"; }
		}
		If ($Text)
		{
			Line 3 "Number of storages: " "$storageCount"
		}
		If ($HTML)
		{
			$columnHeaders = @("Number of storages", ($htmlsilver -bor $htmlbold), "$storageCount", $htmlwhite)
			$rowdata = @()
		}

		ForEach ($Item in $storages)
		{
			If ($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "Position"; Value = $($item.Position); }
				$ScriptInformation += @{ Data = "     Name"; Value = $($item.Name); }
				$ScriptInformation += @{ Data = "     Description"; Value = $($item.Description); }
				$ScriptInformation += @{ Data = "     SR"; Value = $($item.SR); }
				$ScriptInformation += @{ Data = "     Size"; Value = $($item.Size); }
				$ScriptInformation += @{ Data = "     Read Only"; Value = $($item.ReadOnly); }
				$ScriptInformation += @{ Data = "     Priority"; Value = $($item.Priority); }
				$ScriptInformation += @{ Data = "     Active"; Value = $($item.Active); }
				$ScriptInformation += @{ Data = "     Device Path"; Value = $($item.DevicePath); }
			}
			If ($Text)
			{
				Line 3 "Position`t`t`t: " $($item.Position)
				Line 4 "Name`t`t`t: " $($item.Name)
				Line 4 "Description`t`t: " $($item.Description)
				Line 4 "SR`t`t`t: " $($item.SR)
				Line 4 "Size`t`t`t: " $($item.Size)
				Line 4 "Read Only`t`t: " $($item.ReadOnly)
				Line 4 "Priority`t`t: " $($item.Priority)
				Line 4 "Active`t`t`t: " $($item.Active)
				Line 4 "Device Path`t`t: " $($item.DevicePath)
				Line 0 ""
			}
			If ($HTML)
			{
				$rowdata += @(, ("Position", ($htmlsilver -bor $htmlbold), $($item.Position), ($htmlsilver -bor $htmlbold)))
				$rowdata += @(, ("     Name", ($htmlsilver -bor $htmlbold), $($item.Name), $htmlwhite))
				$rowdata += @(, ("     Description", ($htmlsilver -bor $htmlbold), $($item.Description), $htmlwhite))
				$rowdata += @(, ("     SR", ($htmlsilver -bor $htmlbold), $($item.SR), $htmlwhite))
				$rowdata += @(, ("     Size", ($htmlsilver -bor $htmlbold), $($item.Size), $htmlwhite))
				$rowdata += @(, ("     Read Only", ($htmlsilver -bor $htmlbold), $($item.ReadOnly), $htmlwhite))
				$rowdata += @(, ("     Priority", ($htmlsilver -bor $htmlbold), $($item.Priority), $htmlwhite))
				$rowdata += @(, ("     Active", ($htmlsilver -bor $htmlbold), $($item.Active), $htmlwhite))
				$rowdata += @(, ("     Device Path", ($htmlsilver -bor $htmlbold), $($item.DevicePath), $htmlwhite))
			}
		}
		
		If ($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data, Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 350;

			$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If ($Text)
		{
			Line 0 ""
		}
		If ($HTML)
		{
			$msg = ""
			$columnWidths = @("150", "350")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}


}

Function OutputVMNIC
{
	Param([object] $VM)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput VM Network"

	$xenVMVIFs = @($vm.VIFs | Get-XenVIF -EA 0)
	$nrVIFs = $xenVMVIFs.Count
	$networks = @()
	ForEach ($Item in $xenVMVIFs)
	{
		$limit = "QoS disabled"
		if ($item.qos_algorithm_type -like "ratelimit")
		{
			$limit = 'QoS limit of {0} kbps' -f $item.qos_algorithm_params["kbps"]
		}
		If ($item.currently_attached -like $true)
		{
			$active = "Yes"
		}
		Else
		{
			$active = "No"
		}
		$networks += $item | Select-Object -Property `
		@{Name = 'Device'; Expression = { "$($Item.device)" } },
		@{Name = 'MAC'; Expression = { "$($Item.MAC)" } },
		@{Name = 'MACautogenerated'; Expression = { "$($Item.MAC_autogenerated)" } },
		@{Name = 'limit'; Expression = { $limit } },
		@{Name = 'IPAddress'; Expression = { "$(($Item.ipv4_addresses + $Item.ipv6_addresses) -join ", ")" } },
		@{Name = 'Active'; Expression = { $active } }
	}	
	$networks = $networks | Sort-Object -Property Device
	$VMName = $VM.Name_Label

	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Networking"
	}
	If ($Text)
	{
		Line 2 "Networking"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Networking"
	}

	If ($nrVIFs -lt 1)
	{

		If ($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There are no Virtual Network Interfaces configured for VM $VMName"
		}
		If ($Text)
		{
			Line 3 "There are no Virtual Network Interfaces configured for VM $VMName"
			Line 0 ""
		}
		If ($HTML)
		{
			WriteHTMLLine 0 1 "There are no Virtual Network Interfaces configured for VM $VMName"
		}
	}
	Else
	{
		If ($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Number of NICs"; Value = "$($nrVIFs)"; }
		}
		If ($Text)
		{
			Line 3 "Number of NICs`t`t: " "$nrVIFs"
		}
		If ($HTML)
		{
			$columnHeaders = @("Number of NICs", ($htmlsilver -bor $htmlbold), "$nrVIFs", $htmlwhite)
			$rowdata = @()
		}
		ForEach ($Item in $networks)
		{
			If ($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "Device"; Value = "$($Item.device)"; }
				$ScriptInformation += @{ Data = "  MAC address"; Value = "$($Item.MAC)"; }
				$ScriptInformation += @{ Data = "  MAC autogenerated"; Value = "$($Item.MACautogenerated)"; }
				$ScriptInformation += @{ Data = "  Limit"; Value = "$($Item.limit)"; }
				$ScriptInformation += @{ Data = "  IP Address"; Value = "$($Item.IPAddress)"; }
				$ScriptInformation += @{ Data = "  Active"; Value = "$($Item.Active)"; }
			}
			If ($Text)
			{
				Line 3 "Device`t`t`t: " "$($Item.device)"
				Line 3 "  MAC address`t`t: " "$($Item.MAC)"
				Line 3 "  MAC autogenerated`t: " "$($Item.MACautogenerated)"
				Line 3 "  Limit`t`t`t: " "$($Item.limit)"
				Line 3 "  IP Address`t`t: " "$($Item.IPAddress)"
				Line 3 "  Active`t`t: " "$($Item.Active)"
			}
			If ($HTML)
			{
				$rowdata += @(, ("Device", ($htmlsilver -bor $htmlbold), "$($Item.device)", $htmlwhite))
				$rowdata += @(, ("  MAC address", ($htmlsilver -bor $htmlbold), "$($Item.MAC)", $htmlwhite))
				$rowdata += @(, ("  MAC autogenerated", ($htmlsilver -bor $htmlbold), "$($Item.MACautogenerated)", $htmlwhite))
				$rowdata += @(, ("  Limit", ($htmlsilver -bor $htmlbold), "$($Item.limit)", $htmlwhite))
				$rowdata += @(, ("  IP Address", ($htmlsilver -bor $htmlbold), "$($Item.IPAddress)", $htmlwhite))
				$rowdata += @(, ("  Active", ($htmlsilver -bor $htmlbold), "$($Item.Active)", $htmlwhite))
			}
		}
		
		If ($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data, Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If ($Text)
		{
			Line 0 ""
		}
		If ($HTML)
		{
			$msg = ""
			$columnWidths = @("150", "250")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
}

Function OutputVMSnapshots
{
	Param([object] $VM)
	Write-Verbose "$(Get-Date -Format G): `t`tOutput VM Snapshots"

	$xenVMs = $vm.VBDs | Get-XenVBD | Where-Object { $_.type -like "Disk" } | Select-Object -ExpandProperty VDI | Get-XenVDI | Select-Object -ExpandProperty snapshots | Get-XenVDI | Select-Object -ExpandProperty VBDs | Get-XenVBD | Select-Object -ExpandProperty VM | Get-XenVM | Sort-Object snapshot_time, uuid -Unique

	$snapshots = @()
	ForEach ($item in $xenVMs)
	{
		if ($vm.opaque_ref -eq $item.parent.opaque_ref)
		{
			$parent = $($vm.name_label)
		}
		else
		{
			$parentVM = $xenVMs | Where-Object { $_.opaque_ref -eq $item.parent.opaque_ref } | Select-Object -ExpandProperty  name_label
			if ($null -ne $parentVM)
			{
				$parent = $parentVM
			}
			else
			{
				$parent = "<self>"
			}
		}
		if ($item.children.Count -eq 0)
		{
			$children = ""
		}
		else
		{
			$childVMs = @($xenVMs | Where-Object { $_.opaque_ref -in $item.children.opaque_ref } | Select-Object -ExpandProperty name_label)
			if ($childVMs.Count -eq 0)
			{
				$children = ""
			}
			else
			{
				$children = $childVMs -join ", "
			}
		}
		if ($item.power_state -like "Halted")
		{
			$type = "Disks Only"
		}
		Else
		{
			$type = "Disks and Memory"
		}
		$snapshotDateTime = $($item.snapshot_time -as [datetime]).ToLocalTime()
		$snapshotDateTimeValue = '{0} {1}' -f $snapshotDateTime.ToLongDateString(), $snapshotDateTime.ToLongTimeString()
		$snapshots += $item | Select-Object -Property `
		@{Name = 'Name'; Expression = { $_.name_label } },
		@{Name = 'Description'; Expression = { $_.name_description } },
		@{Name = 'Parent'; Expression = { $parent } },
		@{Name = 'Children'; Expression = { $children } },
		@{Name = 'Type'; Expression = { $type } },
		@{Name = 'SnapshotTime'; Expression = { $snapshotDateTimeValue } },
		@{Name = 'Tags'; Expression = { $item.tags -join ", " } },
		@{Name = 'Folder'; Expression = { $item.other_config["folder"] } },
		@{Name = 'CustomFields'; Expression = { Get-XSCustomFields $item.other_config } }
	}
		
	$nrSnapshots = $snapshots.Count
	$VMName = $VM.Name_Label

	If ($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Snapshots"
	}
	If ($Text)
	{
		Line 2 "Snapshots"
	}
	If ($HTML)
	{
		WriteHTMLLine 3 0 "Snapshots"
	}

	If ($nrSnapshots -lt 1)
	{

		If ($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There are no snapshots configured for VM $VMName"
		}
		If ($Text)
		{
			Line 3 "There are no snapshots configured for VM $VMName"
			Line 0 ""
		}
		If ($HTML)
		{
			WriteHTMLLine 0 1 "There are no snapshots configured for VM $VMName"
		}
	}
	Else
	{
		If ($MSWord -or $PDF)
		{
		}
		If ($Text)
		{
			
		}
		If ($HTML)
		{
		}

		ForEach ($Item in $snapshots)
		{
			If ($MSWord -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Name"; Value = "$($Item.Name)"; }
				$ScriptInformation += @{ Data = "Description"; Value = "$($Item.Description)"; }
				$ScriptInformation += @{ Data = "Parent"; Value = "$($Item.Parent)"; }
				$ScriptInformation += @{ Data = "Children"; Value = "$($Item.Children)"; }
				$ScriptInformation += @{ Data = "Type"; Value = "$($Item.Type)"; }
				$ScriptInformation += @{ Data = "Snapshot time"; Value = "$($Item.SnapshotTime)"; }
				$ScriptInformation += @{ Data = "Tags"; Value = "$($Item.Tags)"; }
				$ScriptInformation += @{ Data = "Folder"; Value = "$($Item.Folder)"; }
				$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data, Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

				## IB - Set the header row format
				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 100;
				$Table.Columns.Item(2).Width = 200;

				$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If ($Text)
			{
				Line 3 "Name`t`t: " "$($Item.Name)"
				Line 3 "Description`t: " "$($Item.Description)"
				Line 3 "Parent`t`t: " "$($Item.Parent)"
				Line 3 "Children`t`t: " "$($Item.Children)"
				Line 3 "Type`t`t: " "$($Item.Type)"
				Line 3 "Snapshot time`t: " "$($Item.SnapshotTime)"
				Line 3 "Tags`t`t: " "$($Item.Tags)"
				Line 3 "Folder`t`t: " "$($Item.Folder)"
				Line 0 ""
			}
			If ($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Name", ($htmlsilver -bor $htmlbold), "$($Item.Name)", ($htmlsilver -bor $htmlbold))
				$rowdata += @(, ("Description", ($htmlsilver -bor $htmlbold), "$($Item.Description)", $htmlwhite))
				$rowdata += @(, ("Parent", ($htmlsilver -bor $htmlbold), "$($Item.Parent)", $htmlwhite))
				$rowdata += @(, ("Children", ($htmlsilver -bor $htmlbold), "$($Item.Children)", $htmlwhite))
				$rowdata += @(, ("Type", ($htmlsilver -bor $htmlbold), "$($Item.Type)", $htmlwhite))
				$rowdata += @(, ("SnapshotTime", ($htmlsilver -bor $htmlbold), "$($Item.SnapshotTime)", $htmlwhite))
				$rowdata += @(, ("Tags", ($htmlsilver -bor $htmlbold), "$($Item.Tags)", $htmlwhite))
				$rowdata += @(, ("Folder", ($htmlsilver -bor $htmlbold), "$($Item.Folder)", $htmlwhite))
				$msg = ""
				$columnWidths = @("100", "250")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
			If (($item.CustomFields | Measure-Object | Select-Object -ExpandProperty Count) -lt 1)
			{
				If ($MSWord -or $PDF)
				{
					WriteWordLine 0 1 "There are no Custom Fields specified for $($Item.Name)"
					WriteWordLine 0 0 ""
				}
				If ($Text)
				{
					Line 3 "There are no Custom Fields specified for $($Item.Name)"
					Line 0 ""
				}
				If ($HTML)
				{
					WriteHTMLLine 0 1 "There are no Custom Fields specified for $($Item.Name)"
				}
	
			}
			Else
			{
				If ($MSWord -or $PDF)
				{
					WriteWordLine 0 1 "Custom Fields for snapshot $($Item.Name)"
					WriteWordLine 0 0 ""
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					$ScriptInformation += @{ Data = "Name"; Value = "Value"; }
				}
				If ($Text)
				{
					Line 3 "Custom Fields for snapshot $($Item.Name)"
					Line 0 ""
					Line 3 "Name`t`t: " "Value"
				}
				If ($HTML)
				{
					WriteHTMLLine 0 1 "Custom Fields for snapshot $($Item.Name)"
					$columnHeaders = @("Name", ($htmlsilver -bor $htmlbold), "Value", ($htmlsilver -bor $htmlbold))
					$rowdata = @()
				}

				foreach ($customfield in $item.CustomFields)
				{
					If ($MSWord -or $PDF)
					{
	
						$ScriptInformation += @{ Data = "$($customfield.Name)"; Value = "$($customfield.Value)"; }
					}
					If ($Text)
					{
						Line 3 "$($customfield.Name)`t: " "$($customfield.Value)"
					}
					If ($HTML)
					{
						$rowdata += @(, ("$($customfield.Name)", ($htmlsilver -bor $htmlbold), "$($customfield.Value)", $htmlwhite))
					}
				}

				If ($MSWord -or $PDF)
				{
					$Table = AddWordTable -Hashtable $ScriptInformation `
						-Columns Data, Value `
						-List `
						-Format $wdTableGrid `
						-AutoFit $wdAutoFitFixed;
	
					## IB - Set the header row format
					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
	
					$Table.Columns.Item(1).Width = 150;
					$Table.Columns.Item(2).Width = 250;
	
					$Table.Rows.SetLeftIndent($Indent0TabStops, $wdAdjustProportional)
	
					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If ($Text)
				{
					Line 0 ""
				}
				If ($HTML)
				{
					$msg = ""
					$columnWidths = @("200", "300")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
					WriteHTMLLine 0 0 ""
				}
	
			}
		}
	}
}

#endregion

#region script core
#Script begins

$Results = ProcessScriptSetup
If (!$Results)
{
	Exit
}

If ($Null -eq $Script:XSPool.name_label)
{
	SetFileNames "$($Script:XSHosts[0].hostname)"
}
Else
{
	SetFileNames "$($Script:XSPool.name_label)"
}

If ($Null -eq $Script:XSPool.name_label)
{
	[string]$Script:Title = "Inventory Report for the XenServer Host: $($Script:XSHosts[0].hostname)"
}
Else
{
	[string]$Script:Title = "Inventory Report for the XenServer Pool: $($Script:XSPool.name_label)"
}

Write-Verbose "$(Get-Date -Format G): Start writing report data"

If (("Pool" -in $Section) -or ("All" -in $Section))
{
	ProcessPool
}
If (("Host" -in $Section) -or ("All" -in $Section))
{
	ProcessHosts
}
If (("VM" -in $Section) -or ("All" -in $Section))
{
	ProcessVMs
}
#endregion

#region finish script
Write-Verbose "$(Get-Date -Format G): Finishing up document"
#end of document processing

If (($MSWORD -or $PDF) -and ($Script:CoverPagesExist))
{
	$AbstractTitle = "XenServer Inventory Report"
	$SubjectTitle = "XenServer Inventory Report"
	UpdateDocumentProperties $AbstractTitle $SubjectTitle
}

If ($ReportFooter)
{
	OutputReportFooter
}

ProcessDocumentOutput "Regular"

#disconnect from Pool Master
Write-Host "
	Disconnect from Pool Master $Script:ServerName
	" -ForegroundColor White

Disconnect-XenServer -Session $Script:Session 4>$Null

ProcessScriptEnd

#endregion