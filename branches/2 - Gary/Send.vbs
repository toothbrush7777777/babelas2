'Subversion Info: $Id$
Option Explicit

'This is the configuration to send a message to the mendelson AS2 test server:
'Dim strMDN: strMDN = SendAS2( _
'  "mycompanyAS2", _
'  "3d a0 27 42 4d 92 6d 04 bb 74 66 1d 48 3e 61 6a 46 2a 05 b7", _
'  "hello.txt", _
'  "plain/text", _
'  "http://as2.mendelson-e-c.com:8080/as2/HttpReceiver", _
'  "mendelsontestAS2", _
'  "6d 9a 2c 79 02 0b f1 6b 20 78 e4 a3 be df 93 dd 2a ad b7 40")
'WScript.Echo strMDN

'This is the configuration to send a message to the BabelAS2 test server:
Dim strMDN: strMDN = SendAS2( _
  """BabelAS2 Test Client""", _
  "a5 bc 87 4a b5 96 9d c4 11 d1 5a 93 ac 49 cf 74 1a 12 29 97", _
  "hello.txt", _
  "plain/text", _
  "http://localhost/gary/Receive.asp", _
  """BabelAS2 Test Server""", _
  "67 8f a8 49 b4 7c 7c 94 8e b0 8b ab 0b e8 be fc 65 68 ab 33")
WScript.Echo strMDN

Public Function SendAS2( _
  strMyAS2Id, _
  strMyCertThumbprint, _
  strFileName, _
  strContentType, _
  strPartnerURL, _
  strPartnerAS2Id, _
  strPartnerCertThumbprint)

  Dim strGUID: strGUID = UTIL_CreateGUID()
  Dim xhttp: Set xhttp = CreateObject("MSXML2.ServerXMLHTTP")
  xhttp.open "POST", strPartnerURL, False
  xhttp.setRequestHeader "Connection", "close"
  xhttp.setRequestHeader "Message-Id", "<" & strGUID & "@BabelAS2>"
  xhttp.setRequestHeader "Date", CStr(Now)
  xhttp.setRequestHeader "From", "BabelAS2"
  xhttp.setRequestHeader "Subject", "AS2 Communication with BabelAS2"
  xhttp.setRequestHeader "Mime-Version", "1.0"
  xhttp.setRequestHeader "AS2-Version", "1.1"
  xhttp.setRequestHeader "AS2-From", strMyAS2Id
  xhttp.setRequestHeader "AS2-To", strPartnerAS2Id
  xhttp.setRequestHeader "Content-Type", "application/pkcs7-mime; smime-type=enveloped-data; name=""smime.p7m"""
  xhttp.setRequestHeader "Content-Transfer-Encoding", "base64"
  xhttp.setRequestHeader "Content-Disposition", "attachment; filename=""smime.p7m"""
  xhttp.setRequestHeader "Disposition-notification-To", "babelabout@gmail.com"
  'xhttp.setRequestHeader "Disposition-notification-options", "signed-receipt-protocol=optional,pkcs7-signature;signed-receipt-micalg=optional,sha1"
  xhttp.setRequestHeader "User-Agent", "BabelAS2 - http://code.google.com/p/babelas2/"
  Dim stm: Set stm = CreateObject("ADODB.Stream")
  stm.Open
  stm.Type = 1 'This line is mandatory, otherwise you get the "FF FE" :-/
  stm.LoadFromFile strFileName
  Dim bstrMIME: bstrMIME = UTIL_ASCIIStringToBinaryString( _
    "Content-Type: " & strContentType & "; name=""" & strFileName & """" & vbCrLf & _
    "Content-Disposition: attachment; filename=""" & strFileName & """" & vbCrLf & _
    vbCrLf) & _
    UTIL_ByteArrayToBinaryString(stm.Read)
  stm.Close
  bstrMIME = UTIL_ASCIIStringToBinaryString( _
    "MIME-Version: 1.0" & vbCrLf & _
    "Content-Type: multipart/signed; protocol=""application/pkcs7-signature""; micalg=sha1; boundary=""Part""" & vbCrLf & _
    vbCrLf & _
    "--Part" & vbCrLf) & _
    bstrMIME & _
    UTIL_ASCIIStringToBinaryString(vbCrLf & _
    "--Part" & vbCrLf & _
    "Content-Type: application/pkcs7-signature; name=""smime.p7s""" & vbCrLf & _
    "Content-Disposition: attachment; filename=""smime.p7s""" & vbCrLf & _
    "Content-Transfer-Encoding: base64" & vbCrLf & _
    vbCrLf) & _
    UTIL_ASCIIStringToBinaryString(CRYPTO_Sign(bstrMIME, strMyCertThumbprint)) & _
    UTIL_ASCIIStringToBinaryString("--Part--")
  Dim strBase64: strBase64 = CRYPTO_Encrypt(bstrMIME, strPartnerCertThumbprint)
  stm.Open
  stm.Type = 2
  stm.Charset = "ascii"
  stm.WriteText strBase64
  stm.Position = 0 'This line is mandatory!
  xhttp.send stm
  stm.Close
  WScript.Echo CStr(xhttp.status) & " " & xhttp.statusText
  SendAS2 = xhttp.responseText
  Dim fsoMDN: Set fsoMDN = CreateObject("Scripting.FileSystemObject")
  Dim stmMDN: Set stmMDN = fsoMDN.CreateTextFile("MDN." & strGUID & ".txt", False)
  stmMDN.Write xhttp.responseText
  stmMDN.Close
  Set stmMDN = Nothing
  Set fsoMDN = Nothing

  Set stm = Nothing
  Set xhttp = Nothing
End Function

Function UTIL_CreateGUID()
  Dim tl: Set tl = CreateObject("Scriptlet.TypeLib")
  UTIL_CreateGUID = Mid(tl.Guid, 2, 36)
  Set tl = Nothing
End Function

Function UTIL_ByteArrayToBinaryString(aobContent)
  Dim oUtils: Set oUtils = CreateObject("CAPICOM.Utilities")
  UTIL_ByteArrayToBinaryString = oUtils.ByteArrayToBinaryString(aobContent)
  Set oUtils = Nothing
End Function

Function UTIL_ASCIIStringToBinaryString(str)
  Dim stm: Set stm = CreateObject("ADODB.Stream")
  stm.Open
  stm.Type = 2
  stm.Charset = "iso-8859-1"
  stm.WriteText str
  stm.Position = 0
  stm.Type = 1
  Dim oUtils: Set oUtils = CreateObject("CAPICOM.Utilities")
  UTIL_ASCIIStringToBinaryString = oUtils.ByteArrayToBinaryString(stm.Read)
  stm.Close
  Set oUtils = Nothing
  Set stm = Nothing
End Function

Private Function CRYPTO_GetCertificate(strCertThumbprint) 'As CAPICOM.Certificate
  Const CAPICOM_LOCAL_MACHINE_STORE = 1 'Now we use the Local Computer Personal store!
  Const CAPICOM_STORE_OPEN_READ_ONLY = 0
  strCertThumbprint = Replace(strCertThumbprint, " " , "")
  strCertThumbprint = UCase(strCertThumbprint)
  Dim cer: Set cer = Nothing
  Dim st: Set st = CreateObject("CAPICOM.Store")
  st.Open CAPICOM_LOCAL_MACHINE_STORE, "My", CAPICOM_STORE_OPEN_READ_ONLY
  For Each cer In st.Certificates
    If (StrComp(cer.Thumbprint, strCertThumbprint, vbTextCompare) = 0) Then
      Set CRYPTO_GetCertificate = cer
      Exit For
    End If
  Next
  Set st = Nothing
End Function

Function CRYPTO_Sign(bstrContent, strCertThumbprint) 'As (base64) String
  Const CAPICOM_ENCODE_BASE64 = 0
  Dim oSignedData: Set oSignedData = CreateObject("CAPICOM.SignedData")
  Dim oSigner: Set oSigner = CreateObject("CAPICOM.Signer")
  oSignedData.Content = bstrContent
  oSigner.Certificate = CRYPTO_GetCertificate(strCertThumbprint)
  CRYPTO_Sign = oSignedData.Sign(oSigner, True, CAPICOM_ENCODE_BASE64)
  Set oSigner = Nothing
  Set oSignedData = Nothing
End Function

Function CRYPTO_Encrypt(bstrContent, strCertThumbprint) 'As (base64) String
  Const CAPICOM_ENCODE_BASE64 = 0
  Dim oEnvelopedData: Set oEnvelopedData = CreateObject("CAPICOM.EnvelopedData")
  oEnvelopedData.Algorithm.Name = 3 'CAPICOM_ENCRYPTION_ALGORITHM_3DES
  oEnvelopedData.Content = bstrContent
  oEnvelopedData.Recipients.Add CRYPTO_GetCertificate(strCertThumbprint)
  CRYPTO_Encrypt = oEnvelopedData.Encrypt(CAPICOM_ENCODE_BASE64)
  Set oEnvelopedData = Nothing
End Function
