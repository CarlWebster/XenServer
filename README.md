# XenServer

What? Another XenServer Documentation script attempt? I thought you were retiring on July 1st?

In a "vigorous" discussion with the keeper of the Citrix SDK docs on GitHub, I ended the conversation with, "Don't even get me started on the horrible, terrible, not-good XenServer PowerShell implementation!!!!!!".

That led to a long conversation involving the XenServer engineering and dev teams. They told me they had improved XS PoSH, and it was no longer "a pathetic piece of crap". So I spent a couple of hours yesterday bringing the old 2015 XS6.2 doc script attempt up to the current doc script standards and making it work with the XS 8.2.3 SDK. I told them that even though I am retiring from EUC community service work on July 1st, I have waited and wanted a XenServer doc script for years that I would work on it when I could.

I created this GitHub repo for this script attempt.

If you still use or want to test this script, have at it. I am busy with work and Lions Clubs stuff, but I will do what I can when I can.

If you want to test it, you must do whatever you do on GitHub to receive notifications when I do updates. I started the initial version at .001, as I think it will take a few updates with my schedule to make significant progress.

SYNOPSIS

    Creates an inventory of a XenServer 8.2 CU1 Pool.
    
SYNTAX

    C:\PSScripts\XS_Inventory.ps1 -ServerName <String> [-User <String>] [-HTML] [-Text] 
    [-Folder <String>] [-Section <String[]>] [-AddDateTime] [-Dev] [-Log] [-ScriptInfo] 
    [-ReportFooter] [-SmtpPort <Int32>] [-SmtpServer <String>] [-From <String>] [-To 
    <String>] [-UseSSL] [<CommonParameters>]

    C:\PSScripts\XS_Inventory.ps1 -ServerName <String> [-User <String>] [-HTML] [-Text] 
    [-Folder <String>] [-Section <String[]>] [-AddDateTime] [-Dev] [-Log] [-ScriptInfo] 
    [-ReportFooter] [-MSWord] [-PDF] [-CompanyAddress <String>] [-CompanyEmail <String>] 
    [-CompanyFax <String>] [-CompanyName <String>] [-CompanyPhone <String>] [-CoverPage 
    <String>] [-UserName <String>] [-SmtpPort <Int32>] [-SmtpServer <String>] [-From 
    <String>] [-To <String>] [-UseSSL] [<CommonParameters>]


DESCRIPTION

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


PARAMETERS

    -ServerName <String>
        Specifies which XenServer Pool to use to run the script against.

        You can enter the ServerName as the NetBIOS name, FQDN, or IP Address.

        If entered as an IP address, the script attempts o determine and use the actual
        pool or poolmaster name.
	
        ServerName should be the Pool Master. If you use a Slave host, the script attempts 
        to determine the Pool Master and then makes a connection attempt to the Pool Master. 
        If successful, the script continues. If not successful, the script ends.

        Required?                    true
        Position?                    named
        Default value
        Accept pipeline input?       false
        Accept wildcard characters? false

    -User <String>
        Username to use for the connection to the XenServer Host or Pool.

        Required?                    false
        Position?                    named
        Default value
        Accept pipeline input?       false
        Accept wildcard characters? false

    -HTML [<SwitchParameter>]
        Creates an HTML file with an .html extension.

        HTML is the default report format.

        This parameter is set to True if no other output format is selected.

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters? false

    -Text [<SwitchParameter>]
        Creates a formatted text file with a .txt extension.
        Text formatting is based on the default tab spacing of 8 by Microsoft Notepad.
        
        This parameter is disabled by default.

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters? false

    -Folder <String>
        Specifies the optional output folder to save the output report.

        Required?                    false
        Position?                    named
        Default value
        Accept pipeline input?       false
        Accept wildcard characters? false

    -Section <String[]>
        Processes one or more sections of the report.
        Valid options are:
                Pool
                Host
                VM (Virtual Machines)
                All

        This parameter defaults to All sections.

        A comma separates multiple sections. -Section host, pool

        Required?                    false
        Position?                    named
        Default value                All
        Accept pipeline input?       false
        Accept wildcard characters? false

    -NoPoolMemory [<SwitchParameter>]
        Excludes Pool Memory information from the output document.
	
        This Switch is useful in large XenServer pools, where there may be many hosts.
	
        This parameter is disabled by default.
        This parameter has an alias of NPM.

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters? false

    -NoPoolStorage [<SwitchParameter>]
        Excludes Pool Storage information from the output document.
	
        This Switch is useful in large XenServer pools, where there may be many storage repositories and hosts.
	
        This parameter is disabled by default.
        This parameter has an alias of NPS.

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters? false

    -NoPoolNetworking [<SwitchParameter>]
        Excludes Pool Networking information from the output document.
	
        This Switch is useful in large XenServer pools, where there may be many hosts.
	
        This parameter is disabled by default.
        This parameter has an alias of NPN.

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters? false

    -AddDateTime [<SwitchParameter>]
        Adds a date timestamp to the end of the file name.

        The timestamp is in the format of yyyy-MM-dd_HHmm.
        June 1, 2024 at 6PM is 2024-06-01_1800.

        The output filename is ReportName_2024-06-01_1800.<ext>.

        This parameter is disabled by default.
        This parameter has an alias of ADT.

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters? false

    -Dev [<SwitchParameter>]
        Clears errors at the beginning of the script.
        Outputs all errors to a text file at the end of the script.

        This is used when the script developer requests more troubleshooting data.
        The text file is placed in the same folder from where the script is run.

        This parameter is disabled by default.

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters? false

    -Log [<SwitchParameter>]
        Generates a log file for troubleshooting.

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters? false

    -ScriptInfo [<SwitchParameter>]
        Outputs information about the script to a text file.
        The text file is placed in the same folder from where the script is run.

        This parameter is disabled by default.
        This parameter has an alias of SI.

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters? false

    -ReportFooter [<SwitchParameter>]
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

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters? false

    -MSWord [<SwitchParameter>]
        SaveAs DOCX file

        Microsoft Word is no longer the default report format.
        This parameter is disabled by default.

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters? false

    -PDF [<SwitchParameter>]
        SaveAs PDF file instead of DOCX file.

        The PDF file is roughly 5X to 10X larger than the DOCX file.

        This parameter requires Microsoft Word to be installed.
        This parameter uses Word's SaveAs PDF capability.

        This parameter is disabled by default.

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters? false

    -CompanyAddress <String>
        Company Address to use for the Cover Page if the Cover Page has the Address field.

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

        Required?                    false
        Position?                    named
        Default value
        Accept pipeline input?       false
        Accept wildcard characters? false

    -CompanyEmail <String>
        Company Email to use for the Cover Page if the Cover Page has the Email field.

        The following Cover Pages have an Email field:
                Facet (Word 2013/2016)

        This parameter is only valid with the MSWORD and PDF output parameters.
        This parameter has an alias of CE.

        Required?                    false
        Position?                    named
        Default value
        Accept pipeline input?       false
        Accept wildcard characters? false

    -CompanyFax <String>
        Company Fax to use for the Cover Page if the Cover Page has the Fax field.

        The following Cover Pages have a Fax field:
                Contrast (Word 2010)
                Exposure (Word 2010)

        This parameter is only valid with the MSWORD and PDF output parameters.
        This parameter has an alias of CF.

        Required?                    false
        Position?                    named
        Default value
        Accept pipeline input?       false
        Accept wildcard characters? false

    -CompanyName <String>
        Company Name to use for the Cover Page.
        The default value is contained in
        HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
        HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated
        on the computer running the script.

        This parameter is only valid with the MSWORD and PDF output parameters.
        This parameter has an alias of CN.

        Required?                    false
        Position?                    named
        Default value
        Accept pipeline input?       false
        Accept wildcard characters? false

    -CompanyPhone <String>
        Company Phone to use for the Cover Page if the Cover Page has the Phone field.

        The following Cover Pages have a Phone field:
                Contrast (Word 2010)
                Exposure (Word 2010)

        This parameter is only valid with the MSWORD and PDF output parameters.
        This parameter has an alias of CPh.

        Required?                    false
        Position?                    named
        Default value
        Accept pipeline input?       false
        Accept wildcard characters? false

    -CoverPage <String>
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
                Exposure (Word 2010. Works if you like looking sideways)
                Facet (Word 2013/2016. Works)
                Filigree (Word 2013/2016. Works)
                Grid (Word 2010/2013/2016. Works in 2010)
                Integral (Word 2013/2016. Works)
                Ion (Dark) (Word 2013/2016. Top date doesn't fit; box needs to be
                manually resized or font changed to 8 point)
                Ion (Light) (Word 2013/2016. Top date doesn't fit; box needs to be
                manually resized or font changed to 8 point)
                Mod (Word 2010. Works)
                Motion (Word 2010/2013/2016. Works if the top date is manually changed to
                36 point)
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

        Required?                    false
        Position?                    named
        Default value                Sideline
        Accept pipeline input?       false
        Accept wildcard characters? false

    -UserName <String>
        Username to use for the Cover Page and Footer.
        The default value is contained in $env:username
        This parameter has an alias of UN.
        This parameter is only valid with the MSWORD and PDF output parameters.

        Required?                    false
        Position?                    named
        Default value                $env:username
        Accept pipeline input?       false
        Accept wildcard characters? false

    -SmtpPort <Int32>
        Specifies the SMTP port for the SmtpServer.
        The default is 25.

        Required?                    false
        Position?                    named
        Default value                25
        Accept pipeline input?       false
        Accept wildcard characters? false

    -SmtpServer <String>
        Specifies the optional email server to send the output report(s).

        If From or To are used, this is a required parameter.

        Required?                    false
        Position?                    named
        Default value
        Accept pipeline input?       false
        Accept wildcard characters? false

    -From <String>
        Specifies the username for the From email address.

        If SmtpServer or To are used, this is a required parameter.

        Required?                    false
        Position?                    named
        Default value
        Accept pipeline input?       false
        Accept wildcard characters? false

    -To <String>
        Specifies the username for the To email address.

        If SmtpServer or From are used, this is a required parameter.

        Required?                    false
        Position?                    named
        Default value
        Accept pipeline input?       false
        Accept wildcard characters? false

    -UseSSL [<SwitchParameter>]
        Specifies whether to use SSL for the SmtpServer.
        The default is False.

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters? false

    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (https:/go.microsoft.com/fwlink/?LinkID=113216).

INPUTS

    None. You cannot pipe objects to this script.


OUTPUTS

    No objects are output from this script. This script creates a Word, PDF, HTML, or plain
    text document.


NOTES

        NAME: XS_Inventory.ps1
        VERSION: 0.018
        AUTHOR: Carl Webster and John Billekens along with help from Michael B. Smith, Guy Leech and the XenServer team
        LASTEDIT: July 26, 2023

EXAMPLES

    -------------------------- EXAMPLE 1 --------------------------

    PS C:\PSScript >.\XS_Inventory.ps1

    Outputs, by default, to HTML.
    Prompts for the XenServer Host or Pool and login credentials.




    -------------------------- EXAMPLE 2 --------------------------

    PS C:\>PS C:\PSScript .\XS_Inventory.ps1 -MSWord -CompanyName "Carl Webster
    Consulting" -CoverPage "Mod" -UserName "Carl Webster" -ServerName XS01

    Uses:
        Carl Webster Consulting for the Company Name.
        Mod for the Cover Page format.
        Carl Webster for the User Name.
        XenServer host named XS01 for the ServerName.

    Outputs to Microsoft Word.
    Prompts for the XenServer Host login credentials.




    -------------------------- EXAMPLE 3 --------------------------

    PS C:\>PS C:\PSScript .\XS_Inventory.ps1 -PDF -CN "Carl Webster Consulting" -CP
    "Mod" -UN "Carl Webster"

    Uses:
        Carl Webster Consulting for the Company Name (alias CN).
        Mod for the Cover Page format (alias CP).
        Carl Webster for the User Name (alias UN).

    Outputs to PDF.
    Prompts for the XenServer Host or Pool and login credentials.




    -------------------------- EXAMPLE 4 --------------------------

    PS C:\>PS C:\PSScript .\XS_Inventory.ps1 -CompanyName "Sherlock Holmes
    Consulting" -CoverPage Exposure -UserName "Dr. Watson" -CompanyAddress "221B Baker
    Street, London, England" -CompanyFax "+44 1753 276600" -CompanyPhone "+44 1753 276200"
    -MSWord

    Uses:
        Sherlock Holmes Consulting for the Company Name.
        Exposure for the Cover Page format.
        Dr. Watson for the User Name.
        221B Baker Street, London, England for the Company Address.
        +44 1753 276600 for the Company Fax.
        +44 1753 276200 for the Company Phone.

    Outputs to Microsoft Word.
    Prompts for the XenServer Host or Pool and login credentials.




    -------------------------- EXAMPLE 5 --------------------------

    PS C:\>PS C:\PSScript .\XS_Inventory.ps1 -CompanyName "Sherlock Holmes
    Consulting" -CoverPage Facet -UserName "Dr. Watson" -CompanyEmail
    SuperSleuth@SherlockHolmes.com
    -PDF

    Uses:
        Sherlock Holmes Consulting for the Company Name.
        Facet for the Cover Page format.
        Dr. Watson for the User Name.
        SuperSleuth@SherlockHolmes.com for the Company Email.

    Outputs to PDF.
    Prompts for the XenServer Host or Pool and login credentials.




    -------------------------- EXAMPLE 6 --------------------------

    PS C:\PSScript >.\XS_Inventory.ps1 -Dev -ScriptInfo -Log

    Creates an HTML report.

    Creates a text file named XSInventoryScriptErrors_yyyyMMddTHHmmssffff.txt that
    contains up to the last 250 errors reported by the script.

    Creates a text file named XSInventoryScriptInfo_yyyy-MM-dd_HHmm.txt that
    contains all the script parameters and other basic information.

    Creates a text file for transcript logging named
    XSDocScriptTranscript_yyyyMMddTHHmmssffff.txt.
    Prompts for the XenServer Host or Pool and login credentials.




    -------------------------- EXAMPLE 7 --------------------------

    PS C:\PSScript >.\XS_Inventory.ps1 -Section Pool

    Creates an HTML report that contains only Pool information.
    Processes only the Pool section of the report.
    Prompts for the XenServer Host or Pool and login credentials.




    -------------------------- EXAMPLE 8 --------------------------

    PS C:\PSScript >.\XS_Inventory.ps1 -Section Pool, Host

    Creates an HTML report.

    The report includes only the Pool and Host sections.
    Prompts for the XenServer Host or Pool and login credentials.




    -------------------------- EXAMPLE 9 --------------------------

    PS C:\PSScript >.\XS_Inventory.ps1 -SmtpServer mail.domain.tld -From
    XSAdmin@domain.tld -To ITGroup@domain.tld -Text

    The script uses the email server mail.domain.tld, sending from XSAdmin@domain.tld
    and sending to ITGroup@domain.tld.

    The script uses the default SMTP port 25 and does not use SSL.

    If the current user's credentials are not valid to send an email, the script prompts
    the user to enter valid credentials.

    Outputs to a text file.
    Prompts for the XenServer Host or Pool and login credentials.




    -------------------------- EXAMPLE 10 --------------------------

    PS C:\PSScript >.\XS_Inventory.ps1 -SmtpServer mailrelay.domain.tld -From
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

    To send an email using a Gmail or g-suite account, you may have to turn ON the "Less
    secure app access" option on your account.
    ***GMAIL/G SUITE SMTP RELAY***

    The script generates an anonymous, secure password for the anonymous@domain.tld
    account.

    Outputs, by default, to HTML.
    Prompts for the XenServer Host or Pool and login credentials.




    -------------------------- EXAMPLE 11 --------------------------

    PS C:\PSScript >.\XS_Inventory.ps1 -SmtpServer
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
    Prompts for the XenServer Host or Pool and login credentials.




    -------------------------- EXAMPLE 12 --------------------------

    PS C:\PSScript >.\XS_Inventory.ps1 -SmtpServer smtp.office365.com -SmtpPort 587
    -UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com

    The script uses the email server smtp.office365.com on port 587 using SSL, sending from
    webster@carlwebster.com and sending to ITGroup@carlwebster.com.

    If the current user's credentials are not valid to send an email, the script prompts
    the user to enter valid credentials.

    Outputs, by default, to HTML.
    Prompts for the XenServer Host or Pool and login credentials.




    -------------------------- EXAMPLE 13 --------------------------

    PS C:\PSScript >.\XS_Inventory.ps1 -SmtpServer smtp.gmail.com -SmtpPort 587
    -UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com

    *** NOTE ***
    To send an email using a Gmail or g-suite account, you may have to turn ON the "Less
    secure app access" option on your account.
    *** NOTE ***

    The script uses the email server smtp.gmail.com on port 587 using SSL, sending from
    webster@gmail.com and sending to ITGroup@carlwebster.com.

    If the current user's credentials are not valid to send an email, the script prompts
    the user to enter valid credentials.

    Outputs, by default, to HTML.
    Prompts for the XenServer Host or Pool and login credentials.


RELATED LINKS
