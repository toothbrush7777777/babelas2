# Introduction #

This guide gives you a exhaustive step-by-step procedure for installing and configuring **BabelAS2** version 0.03.


# Installation #

This installation procedure have been tested with the following operation systems:
  * Windows XP with IIS V5.1
  * Windows 2003 Server with IIS V6.0

**Extract BabelAS2.v0.03.zip:**
  1. Create your "`<`Install`>`" folder, e.g. "C:\Program Files\BabelAS2"
  1. Extract "BabelAS2.v0.03.zip" with path names in your "`<`Install`>`" folder.

**Install CAPICOM:**
  1. Copy "`<`Install`>`\Setup\capicom.dll" to "C:\WINDOWS\system32"
  1. Click on Start/Programs/Accessories/Command Prompt
  1. Type "C:"
  1. Type "cd C:\WINDOWS\system32"
  1. Type "regsvr32.exe capicom.dll"
  1. Click on the "OK" button.
  1. Type "exit"

**_Remark:_** You can download the "capicom.dll" yourself directly from Microsoft at http://www.microsoft.com/downloads/details.aspx?FamilyID=860EE43A-A843-462F-ABB5-FF88EA5896F6&displaylang=en.

**Install the test certificate(s) and private key(s):**
  1. Double-click on "`<`Install`>`\Setup\BabelAS2.msc"
  1. Expand the "Certificates (Local Computer)" node
  1. Right-click on "Personal" and select "All Tasks/Import..."
  1. Click on the "Next" button.
  1. Click on the "Browse..." button
  1. Select "Personal Information Exchange (`*`.pfx;`*`.p12)" in the "Files of type:" drop-down list
  1. Select "`<`Install`>`\Cert\BabelAS2`_`Test`_`Client.pfx"
  1. Click on the "Open" button
  1. Click on the "Next" button
  1. Type "test" in the "Password:" edit box
  1. Check "Mark this key as exportable. This will allow you to back up or transport your keys at a later time."
  1. Click on the "Next" button.
  1. Click on the "Next" button.
  1. Click on the "Finish" button.
  1. Click on the "OK" button.
  1. Repeat from step **3**, with selecting "`<`Install`>`\Cert\BabelAS2`_`Test`_`Server.pfx" at step **7**.

**_Remark:_** If you only send messages, you can import "`<`Install`>`\Cert\BabelAS2`_`Test`_`Server.**cer**" instead. And, if you only want to receive, you can import "`<`Install`>`\Cert\BabelAS2`_`Test`_`Client.**cer**".

**Send a test message to the _BabelAS2 Test Server_:**
  1. Double-click on "`<`Install`>`\Send.vbs"
  1. Click on the "OK" button
  1. Go to http://babelas2.babelabout.net/Test/ to check if your test message has arrived
  1. Report any problem (or success) to [mailto:babelabout@gmail.com](mailto:babelabout@gmail.com) ;-)

**Create the BabelAS2 user:**
  1. Double-click on "`<`Install`>`\BabelAS2.msc" (if needed)
  1. Expand the "Local User and Groups (Local)" node
  1. Right-click on "Users" and select "New User..."
  1. Type "BabelAS2" in the "User name:" edit box
  1. Type "BabelAS2" in the "Full name:" edit box
  1. Type "BabelAS2" in the "Description:" edit box
  1. Type "`<`Password`>`" in the "Password:" edit box
  1. Type "`<`Password`>`" in the "Confirm password:" edit box
  1. Uncheck "User must change password at next logon"
  1. Check "User cannot change password"
  1. Check "Password never expires"
  1. Leave unchecked "Account is disabled"
  1. Click on the "Create" button.
  1. Click on the "Close" button.
  1. Double-click on "BabelAS2" on the right pane
  1. Click on the "Member Of" tab
  1. Select the "Users" item
  1. Click on the "Remove" button
  1. Click on the "OK" button.

In order to receive messages, we recommend that you create a dedicated web site on a Windows 2003 server, but you can also just create a virtual directory on a Windows XP machine, see below.

**Create the babelas2.domain.net web site (Windows 2003 only):**
  1. Expand the "Internet Information Services" node
  1. Expand the "`<`Computer`>` (local computer)" node
  1. Right-click on "Web Sites" and select "New/Web Site..." (not available on Windows XP)
  1. Click on the "Next" button
  1. Type "babelas2.domain.net" in the "Description:" edit box
  1. Click on the "Next" button
  1. Type "babelas2.domain.net" in the "Host header for this Web site (Default: None):" edit box
  1. Click on the "Next" button
  1. Click on the "Browse..." button
  1. Select "`<`Install`>`\Receive"
  1. Click on the "OK" button
  1. Check "Allow anonymous access to this Web site
  1. Click on the "Next" button
  1. Check "Read"
  1. Check "Run scripts (such as ASP)"
  1. Uncheck "Execute (such as ISAPI applications or CGI)"
  1. Uncheck "Write"
  1. Uncheck "Browse"
  1. Click on the "Next" button
  1. Click on the "Finish" button
  1. Right-click on "babelas2.domain.net" and select "Properties"
  1. Click on the "Home Directory" tab
  1. Click on the "Configuration..." button
  1. Click on the "Options" tab
  1. Check "Enable session state"
  1. Check "Enable buffering"
  1. Check "Enable parent paths"
  1. Click on the "OK" button
  1. Click on the "Documents" tab
  1. Click on the "Add..." button
  1. Type "Receive.asp" in the "Default content page:" edit box
  1. Click on the "OK" button
  1. Click on the "Directory Security" tab
  1. Click on the "Edit..." button of the "Authentication and access control" section
  1. Check "Enable anonymous access"
  1. Click on the "Browse..." button
  1. Type "BabelAS2" in the "Enter the object names to select (examples):" text box
  1. Click on the "Check Names" button
  1. Click on the "OK" button
  1. Type "`<`Password`>`" in the "Password:" edit box
  1. Uncheck "Integrated Windows authentication"
  1. Uncheck "Digest authentication for Windows domain servers"
  1. Uncheck "Basic authentication (password is sent in clear text)
  1. Uncheck ".NET Passport authentication"
  1. Click on the "OK" button
  1. Type "`<`Password`>`" in the "Please re-enter the password to confirm:" edit box
  1. Click on the "OK" button
  1. Click on the "OK" button
  1. Click on Start/Programs/Accessories/Windows Explorer
  1. Select "`<`Install`>`\Data"
  1. Right-click on "Data" and select "Properties"
  1. Click on the "Security" tab
  1. Click on the "Add..." button
  1. Type "BabelAS2" in the "Enter the object names to select (examples):" text box
  1. Click on the "Check Names" button
  1. Click on the "OK" button
  1. Check "Full Control" in the "Allow" column
  1. Click on the "OK" button
  1. Click on Start/Programs/Accessories/Command Prompt
  1. Type "C:"
  1. Type "cd `<`Install`>`\Setup"
  1. Type "winhttpcertcfg -g -c LOCAL\_MACHINE\My -s "BabelAS2 Test Server" -a BabelAS2"
  1. Type "exit".

**_Remark:_** You can download the Windows HTTP Services Certificate Configuration Tool (WinHttpCertCfg.exe) yourself directly from Microsoft at http://www.microsoft.com/downloads/details.aspx?familyid=c42e27ac-3409-40e9-8667-c748e422833f&displaylang=en.

**Create the BabelAS2 virtual directory (Windows XP):**
  1. Expand the "Internet Information Services" node
  1. Expand the "`<`Computer`>` (local computer)" node
  1. Expand the "Web Sites" node
  1. Right-click on "Default Web Sites" and select "New/Virtual Directory..."
  1. Click on the "Next" button
  1. Type "BabelAS2" in the "Alias:" edit box
  1. Click on the "Next" button
  1. Click on the "Browse..." button
  1. Select "`<`Install`>`\Receive"
  1. Click on the "OK" button
  1. Click on the "Next" button
  1. Check "Read"
  1. Check "Run scripts (such as ASP)"
  1. Uncheck "Execute (such as ISAPI applications or CGI)"
  1. Uncheck "Write"
  1. Uncheck "Browse"
  1. Click on the "Next" button
  1. Click on the "Finish" button
  1. Expand the "Default Web Sites" node
  1. Right-click on "BabelAS2" and select "Properties"
  1. Click on the "Configuration..." button
  1. Click on the "Options" tab
  1. Check "Enable session state"
  1. Check "Enable buffering"
  1. Check "Enable parent paths"
  1. Click on the "OK" button
  1. Click on the "Directory Security" tab
  1. Click on the "Edit..." button of the "Anonymous access and authentication control" section
  1. Check "Anonymous access"
  1. Click on the "Browse..." button
  1. Type "BabelAS2" in the "Enter the object names to select (examples):" text box
  1. Click on the "Check Names" button
  1. Click on the "OK" button
  1. Uncheck "Allow IIS to control password"
  1. Type "`<`Password`>`" in the "Password:" edit box
  1. Uncheck "Digest authentication for Windows domain servers"
  1. Uncheck "Basic authentication (password is sent in clear text)
  1. Uncheck "Integrated Windows authentication"
  1. Click on the "OK" button
  1. Type "`<`Password`>`" in the "Please re-enter the password to confirm:" edit box
  1. Click on the "OK" button
  1. Click on the "OK" button
  1. Click on Start/Programs/Accessories/Windows Explorer
  1. Select "`<`Install`>`\Data"
  1. Right-click on "Data" and select "Properties"
  1. Click on the "Security" tab
  1. Click on the "Add..." button
  1. Type "BabelAS2" in the "Enter the object names to select (examples):" text box
  1. Click on the "Check Names" button
  1. Click on the "OK" button
  1. Check "Full Control" in the "Allow" column
  1. Click on the "OK" button
  1. Click on Start/Programs/Accessories/Command Prompt
  1. Type "C:"
  1. Type "cd `<`Install`>`\Setup"
  1. Type "winhttpcertcfg -g -c LOCAL\_MACHINE\My -s "BabelAS2 Test Server" -a BabelAS2"
  1. Type "exit".

**Send a test message to your BabelAS2 server:**
  1. Click on Start/Programs/Accessories/Windows Explorer
  1. Select "`<`Install`>`\Send.vbs"
  1. Right-click on "Send.vbs" and select "Edit"
  1. Go to the line 21 and replace "http://babelas2.babelabout.net/" by "http://babelas2.domain.net/" (Windows 2003) or "http://localhost/BabelAS2/Receive.asp" (Windows XP)
  1. Click on "File/Save"
  1. Click on "File/Exit"
  1. Double-click on "`<`Install`>`\Send.vbs"
  1. Click on the "OK" button
  1. Go to "`<`Install`>`\Data to check if your test message has arrived
  1. Report any problem (or success) to [mailto:babelabout@gmail.com](mailto:babelabout@gmail.com) ;-)