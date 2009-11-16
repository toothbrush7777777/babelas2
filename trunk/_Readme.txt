Installation Procedure of BabelAS2
----------------------------------

Patrice, 2009-11-16

1. Unzip BabelAS2.v0.mm.zip in a <folder>.

How to install CAPICOM.

2. Copy "<folder>\BabelAS2\tags\v0.mm\Setup\capicom.dll" to "C:\WINDOWS\system32".
3. Execute "regsvr32.exe capicom.dll" from "C:\WINDOWS\system32".

How to install BabelAS2 to send messages to the BabelAS2 test server.

[Install the certificate and private key of the sender (BabelAS2 Test Client)]
4. Double-click on "BabelAS2\tags\v0.mm\Cert\BabelAS2_Test_Client.pfx".
5. Click "Next".
6. Click "Next".
7. Type the password "test", leave the two check boxes and click "Next".
8. Select the radio button "Place all certificates in the following store", click the "Browse..." button, select the "Personal" certificate store, click "OK", and click "Next".
9. Click "Finish".
[Install the certificate of the receiver (BabelAS2 Test Server)]
10. Right-click on "<folder>\BabelAS2\tags\v0.mm\Cert\BabelAS2_Test_Server.cer" and select "Install Certificate".
11. Click "Next".
12. Select the radio button "Place all certificates in the following store", click the "Browse..." button, select the "Personal" certificate store, click "OK", and click "Next".
13. Click "Finish".
[Send the test EDIFACT message to the BabelAS2 test server]
14. Double-click on "<folder>\BabelAS2\tags\v0.mm\Send.vbs"

Please send any question/comment/feedback to babelabout@gmail.com

How to install BabelAS2 to send messages to the mendelson AS2 test server.

[Install the certificate and private key of the sender (mycompanyAS2)]
4. Double-click on "BabelAS2\tags\v0.mm\Cert\key1.pfx" (downloaded from "http://www.mendelson.de/en/mecas2/key1.pfx").
5. Click "Next".
6. Click "Next".
7. Type the password "test", leave the two check boxes and click "Next".
8. Select the radio button "Place all certificates in the following store", click the "Browse..." button, select the "Personal" certificate store, click "OK", and click "Next".
9. Click "Finish".
[Install the certificate of the receiver (mendelsontestAS2)]
10. Right-click on "<folder>\BabelAS2\tags\v0.mm\Cert\key2.cer" (downloaded form "http://www.mendelson.de/en/mecas2/key2.cer") and select "Install Certificate".
11. Click "Next".
12. Select the radio button "Place all certificates in the following store", click the "Browse..." button, select the "Personal" certificate store, click "OK", and click "Next".
13. Click "Finish".
[Send the test EDIFACT message to the mendelson AS2 test server]
14. Edit the "Send.vbs" to comment the lines 4 to 11, and un-comment the line 13 to 20.
15. Double-click on "<folder>\BabelAS2\tags\v0.mm\Send.vbs"

Please send any question/comment/feedback to babelabout@gmail.com

[/]