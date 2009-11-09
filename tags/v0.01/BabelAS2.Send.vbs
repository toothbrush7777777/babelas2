'Subversion Info: $Id$
Option Explicit

Dim strMDN: strMDN = SendAS2( _
  "mycompanyAS2", _
  "3d a0 27 42 4d 92 6d 04 bb 74 66 1d 48 3e 61 6a 46 2a 05 b7", _
  "EDIFACT.instance.txt", _
  "application/EDIFACT", _
  "http://as2.mendelson-e-c.com:8080/as2/HttpReceiver", _
  "mendelsontestAS2", _
  "6d 9a 2c 79 02 0b f1 6b 20 78 e4 a3 be df 93 dd 2a ad b7 40")
WScript.Echo strMDN

Public Function SendAS2( _
  strMyAS2Id, _
  strMyCertThumbprint, _
  strFileName, _
  strContentType, _
  strPartnerURL, _
  strPartnerAS2Id, _
  strPartnerCertThumbprint)

  Dim xhttp: Set xhttp = CreateObject("MSXML2.ServerXMLHTTP")
  xhttp.open "POST", strPartnerURL, False
  xhttp.setRequestHeader "Connection", "close"
  xhttp.setRequestHeader "Message-Id", "<" & CreateGUID() & "@BA>"
  xhttp.setRequestHeader "Date", CStr(Now)
  xhttp.setRequestHeader "From", "BabelAS2"
  xhttp.setRequestHeader "Subject", "AS2 Communication"
  xhttp.setRequestHeader "Mime-Version", "1.0"
  xhttp.setRequestHeader "AS2-Version", "1.1"
  xhttp.setRequestHeader "AS2-From", strMyAS2Id
  xhttp.setRequestHeader "AS2-To", strPartnerAS2Id
  xhttp.setRequestHeader "Content-Type", "application/pkcs7-mime; smime-type=enveloped-data; name=""smime.p7m"""
  xhttp.setRequestHeader "Content-Transfer-Encoding", "base64"
  xhttp.setRequestHeader "Content-Disposition", "attachment; filename=""smime.p7m"""
  xhttp.setRequestHeader "Disposition-notification-To", "dummy@email.com"
  xhttp.setRequestHeader "Disposition-notification-options", "signed-receipt-protocol=optional,pkcs7-signature;signed-receipt-micalg=optional,sha1"
  Dim stm: Set stm = CreateObject("ADODB.Stream")
  stm.Open
  stm.Type = 1 'This line is mandatory, otherwise you get the "FF FE" :-/
  stm.LoadFromFile strFileName
  stm.Type = 2
  stm.Charset = "ascii"
  Dim str: str = _
    "Content-Type: " & strContentType & "; name=""" & strFileName & """" & vbCrLf & _
    "Content-Transfer-Encoding: binary" & vbCrLf & _
    "Content-Disposition: attachment; filename=""" & strFileName & """" & vbCrLf & _
    vbCrLf & _
    stm.ReadText()
  str = _
    "MIME-Version: 1.0" & vbCrLf & _
    "Content-Type: multipart/signed; protocol=""application/pkcs7-signature""; micalg=sha1; boundary=""Part""" & vbCrLf & _
    vbCrLf & _
    "--Part" & vbCrLf & _
    str & vbCrLf & _
    "--Part" & vbCrLf & _
    "Content-Type: application/pkcs7-signature; name=""smime.p7s""" & vbCrLf & _
    "Content-Disposition: attachment; filename=""smime.p7s""" & vbCrLf & _
    "Content-Transfer-Encoding: base64" & vbCrLf & _
    vbCrLf & _
    CRYPTO_Sign(str, strMyCertThumbprint) & _
    "--Part--"
  str = CRYPTO_Envelop(str, strPartnerCertThumbprint)
  stm.Close
  stm.Open
  stm.Type = 2
  stm.Charset = "ascii"
  stm.WriteText str
  stm.Position = 0 ' 'This line is mandatory!
  xhttp.send stm
  SendAS2 = xhttp.responseText
  stm.Close
  Set stm = Nothing
  Set xhttp = Nothing
End Function

Const CAPICOM_CURRENT_USER_STORE = 2
Const CAPICOM_STORE_OPEN_READ_ONLY = 0
Const CAPICOM_ENCODE_BASE64 = 0

Private Function CRYPTO_GetCertificate(strCertThumbprint) 'As CAPICOM.Certificate
  strCertThumbprint = Replace(strCertThumbprint, " " , "")
  strCertThumbprint = UCase(strCertThumbprint)
  Dim cer: Set cer = Nothing
  Dim st: Set st = CreateObject("CAPICOM.Store")
  st.Open CAPICOM_CURRENT_USER_STORE, "My", CAPICOM_STORE_OPEN_READ_ONLY
  For Each cer In st.Certificates
    If (StrComp(cer.Thumbprint, strCertThumbprint, vbTextCompare) = 0) Then
      Set CRYPTO_GetCertificate = cer
      Exit For
    End If
  Next
  Set st = Nothing
End Function

Private Function CRYPTO_Sign(strData, strCertThumbprint) 'As String
  Dim oSignedData: Set oSignedData = CreateObject("CAPICOM.SignedData")
  Dim oSigner: Set oSigner = CreateObject("CAPICOM.Signer")
  Dim stm: Set stm = CreateObject("ADODB.Stream")
  stm.Open
  stm.Type = 2
  stm.Charset = "ascii"
  stm.WriteText strData
  stm.Position = 0
  stm.Type = 1
  oSignedData.Content = stm.Read
  oSigner.Certificate = CRYPTO_GetCertificate(strCertThumbprint)
  CRYPTO_Sign = oSignedData.Sign(oSigner, True, CAPICOM_ENCODE_BASE64)
  stm.Close
  Set stm = Nothing
  Set oSigner = Nothing
  Set oSignedData = Nothing
End Function

Private Function CRYPTO_Envelop(strData, strCertThumbprint) 'As String
  Dim oEnvelopedData: Set oEnvelopedData = CreateObject("CAPICOM.EnvelopedData")
  Dim stm: Set stm = CreateObject("ADODB.Stream")
  stm.Open
  stm.Type = 2
  stm.Charset = "ascii"
  stm.WriteText strData
  stm.Position = 0
  stm.Type = 1
  oEnvelopedData.Algorithm.Name = 3 'CAPICOM_ENCRYPTION_ALGORITHM_3DES
  oEnvelopedData.Content = stm.Read
  oEnvelopedData.Recipients.Add CRYPTO_GetCertificate(strCertThumbprint)
  CRYPTO_Envelop = oEnvelopedData.Encrypt(CAPICOM_ENCODE_BASE64)
  stm.Close
  Set stm = Nothing
  Set oEnvelopedData = Nothing
End Function

Function CreateGUID()
  Dim oUtil: Set oUtil = CreateObject("Event.Util")
  Dim strGUID: strGUID = oUtil.GetNewGUID
  CreateGUID = Mid(strGUID, 2, Len(strGUID)-2)
  Set oUtil = Nothing
End Function