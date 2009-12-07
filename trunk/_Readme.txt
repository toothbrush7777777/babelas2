
How to install and configure BabelAS2 version 0.03, babelabout@gmail.com, 2009-12-07

= Introduction =

This guide gives you a exhaustive step-by-step procedure for installing and configuring BabelAS2.

= Installation =

This installation procedure have been tested with the following operation systems:
  - Windows XP with IIS V5.1
  - Windows 2003 Server with IIS V6.0

Extract BabelAS2.v0.03.zip:

 1. Create your "<Install>" folder, e.g. "C:\Program Files\BabelAS2"
 2. Extract "BabelAS2.v0.03.zip" with path names in your "<Install>" folder.

Install CAPICOM:

  1. Copy "<Install>\Setup\capicom.dll" to "C:\WINDOWS\system32"
  2. Click on Start/Programs/Accessories/Command Prompt
  3. Type "C:"
  4. Type "cd C:\WINDOWS\system32"
  5. Type "regsvr32.exe capicom.dll"
  6. Click on the "OK" button.
  7. Type "exit"

[Remark: You can download the "capicom.dll" yourself directly from Microsoft at http://www.microsoft.com/downloads/details.aspx?FamilyID=860EE43A-A843-462F-ABB5-FF88EA5896F6&displaylang=en]

Install the test certificate(s) and private key(s):

  1. Double-click on "<Install>\Setup\BabelAS2.msc"
  2. Expand the "Certificates (Local Computer)" node
  3. Right-click on "Personal" and select "All Tasks/Import..."
  4. Click on the "Next" button.
  5. Click on the "Browse..." button
  6. Select "Personal Information Exchange (*.pfx;*.p12)" in the "Files of type:" drop-down list
  7. Select "<Install>\Cert\BabelAS2_Test_Client.pfx"
  8. Click on the "Open" button
  9. Click on the "Next" button
 10. Type "test" in the "Password:" edit box
 11. Check "Mark this key as exportable. This will allow you to back up or transport your keys at a later time."
 12. Click on the "Next" button.
 13. Click on the "Next" button.
 14. Click on the "Finish" button.
 15. Click on the "OK" button.
 16. Repeat from step 3, with selecting "<Install>\Cert\BabelAS2_Test_Server.pfx" at step 7.

[Remark: If you only send messages, you can import "<Install>\Cert\BabelAS2_Test_Server.cer" instead. And, if you only want to receive, you can import "<Install>\Cert\BabelAS2_Test_Client.cer"]

Send a test message to the BabelAS2 Test Server:

  1. Double-click on "<Install>\Send.vbs"
  2. Click on the "OK" button
  3. Go to http://babelas2.babelabout.net/Test/ to check if your test message has arrived
  4. Report any problem (or success) to mailto:babelabout@gmail.com ;-)

Create the BabelAS2 user:

  1. Double-click on "<Install>\Setup\BabelAS2.msc" (if needed)
  2. Expand the "Local User and Groups (Local)" node
  3. Right-click on "Users" and select "New User..."
  4. Type "BabelAS2" in the "User name:" edit box
  5. Type "BabelAS2" in the "Full name:" edit box
  6. Type "BabelAS2" in the "Description:" edit box
  7. Type "<Password>" in the "Password:" edit box
  8. Type "<Password>" in the "Confirm password:" edit box
  9. Uncheck "User must change password at next logon"
 10. Check "User cannot change password"
 11. Check "Password never expires"
 12. Leave unchecked "Account is disabled"
 13. Click on the "Create" button.
 14. Click on the "Close" button.
 15. Double-click on "BabelAS2" on the right pane
 16. Click on the "Member Of" tab
 17. Select the "Users" item
 18. Click on the "Remove" button
 19. Click on the "OK" button.

In order to receive messages, we recommend that you create a dedicated web site on a Windows 2003 server, but you can also just create a virtual directory on a Windows XP machine, see below.

Create the babelas2.domain.net web site (Windows 2003 only):

  1. Expand the "Internet Information Services" node
  2. Expand the "<Computer> (local computer)" node
  3. Right-click on "Web Sites" and select "New/Web Site..." (not available on Windows XP)
  4. Click on the "Next" button
  5. Type "babelas2.domain.net" in the "Description:" edit box
  6. Click on the "Next" button
  7. Type "babelas2.domain.net" in the "Host header for this Web site (Default: None):" edit box
  8. Click on the "Next" button
  9. Click on the "Browse..." button
 10. Select "<Install>\Receive"
 11. Click on the "OK" button
 12. Check "Allow anonymous access to this Web site
 13. Click on the "Next" button
 14. Check "Read"
 15. Check "Run scripts (such as ASP)"
 16. Uncheck "Execute (such as ISAPI applications or CGI)"
 17. Uncheck "Write"
 18. Uncheck "Browse"
 19. Click on the "Next" button
 20. Click on the "Finish" button
 21. Right-click on "babelas2.domain.net" and select "Properties"
 22. Click on the "Home Directory" tab
 23. Click on the "Configuration..." button
 24. Click on the "Options" tab
 25. Check "Enable session state"
 26. Check "Enable buffering"
 27. Check "Enable parent paths"
 28. Click on the "OK" button
 29. Click on the "Documents" tab
 30. Click on the "Add..." button
 31. Type "Receive.asp" in the "Default content page:" edit box
 32. Click on the "OK" button
 33. Click on the "Directory Security" tab
 34. Click on the "Edit..." button of the "Authentication and access control" section
 35. Check "Enable anonymous access"
 36. Click on the "Browse..." button
 37. Type "BabelAS2" in the "Enter the object names to select (examples):" text box
 38. Click on the "Check Names" button
 39. Click on the "OK" button
 40. Type "<Password>" in the "Password:" edit box
 41. Uncheck "Integrated Windows authentication"
 42. Uncheck "Digest authentication for Windows domain servers"
 43. Uncheck "Basic authentication (password is sent in clear text)
 44. Uncheck ".NET Passport authentication"
 45. Click on the "OK" button
 46. Type "<Password>" in the "Please re-enter the password to confirm:" edit box
 47. Click on the "OK" button
 48. Click on the "OK" button
 49. Click on Start/Programs/Accessories/Windows Explorer
 50. Select "<Install>\Data"
 51. Right-click on "Data" and select "Properties"
 52. Click on the "Security" tab
 53. Click on the "Add..." button
 54. Type "BabelAS2" in the "Enter the object names to select (examples):" text box
 55. Click on the "Check Names" button
 56. Click on the "OK" button
 57. Check "Full Control" in the "Allow" column
 58. Click on the "OK" button
 59. Click on Start/Programs/Accessories/Command Prompt
 60. Type "C:"
 61. Type "cd <Install>\Setup"
 62. Type "winhttpcertcfg -g -c LOCAL_MACHINE\My -s "BabelAS2 Test Server" -a BabelAS2"
 63. Type "exit".

Create the BabelAS2 virtual directory (Windows XP):

  1. Expand the "Internet Information Services" node
  2. Expand the "<Computer> (local computer)" node
  3. Expand the "Web Sites" node
  4. Right-click on "Default Web Sites" and select "New/Virtual Directory..."
  5. Click on the "Next" button
  6. Type "BabelAS2" in the "Alias:" edit box
  7. Click on the "Next" button
  8. Click on the "Browse..." button
  9. Select "<Install>\Receive"
 10. Click on the "OK" button
 11. Click on the "Next" button
 12. Check "Read"
 13. Check "Run scripts (such as ASP)"
 14. Uncheck "Execute (such as ISAPI applications or CGI)"
 15. Uncheck "Write"
 16. Uncheck "Browse"
 17. Click on the "Next" button
 18. Click on the "Finish" button
 19. Expand the "Default Web Sites" node
 20. Right-click on "BabelAS2" and select "Properties"
 21. Click on the "Configuration..." button
 22. Click on the "Options" tab
 23. Check "Enable session state"
 24. Check "Enable buffering"
 25. Check "Enable parent paths"
 26. Click on the "OK" button
 27. Click on the "Directory Security" tab
 28. Click on the "Edit..." button of the "Anonymous access and authentication control" section
 29. Check "Anonymous access"
 30. Click on the "Browse..." button
 31. Type "BabelAS2" in the "Enter the object names to select (examples):" text box
 32. Click on the "Check Names" button
 33. Click on the "OK" button
 34. Uncheck "Allow IIS to control password"
 35. Type "<Password>" in the "Password:" edit box
 36. Uncheck "Digest authentication for Windows domain servers"
 37. Uncheck "Basic authentication (password is sent in clear text)
 38. Uncheck "Integrated Windows authentication"
 39. Click on the "OK" button
 40. Type "<Password>" in the "Please re-enter the password to confirm:" edit box
 41. Click on the "OK" button
 42. Click on the "OK" button
 43. Click on Start/Programs/Accessories/Windows Explorer
 44. Select "<Install>\Data"
 45. Right-click on "Data" and select "Properties"
 46. Click on the "Security" tab
 47. Click on the "Add..." button
 48. Type "BabelAS2" in the "Enter the object names to select (examples):" text box
 49. Click on the "Check Names" button
 50. Click on the "OK" button
 51. Check "Full Control" in the "Allow" column
 52. Click on the "OK" button
 53. Click on Start/Programs/Accessories/Command Prompt
 54. Type "C:"
 55. Type "cd <Install>\Setup"
 56. Type "winhttpcertcfg -g -c LOCAL_MACHINE\My -s "BabelAS2 Test Server" -a BabelAS2"
 57. Type "exit".

Send a test message to your BabelAS2 server:

  1. Click on Start/Programs/Accessories/Windows Explorer
  2. Select "<Install>\Send.vbs"
  3. Right-click on "Send.vbs" and select "Edit"
  4. Go to the line 21 and replace "http://babelas2.babelabout.net/" by "babelas2.domain.net" (Windows 2003) or "http://localhost/BabelAS2/Receive.asp" (Windows XP)
  5. Click on "File/Save"
  6. Click on "File/Exit"
  7. Double-click on "<Install>\Send.vbs"
  8. Click on the "OK" button
  9. Go to "<Install>\BabelAS2\Data to check if your test message has arrived
 10. Report any problem (or success) to mailto:babelabout@gmail.com ;-)
