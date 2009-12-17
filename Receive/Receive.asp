<% Option Explicit

Dim strGUID: strGUID = CreateGUID()
Trace strGUID, "strGUID = """ & strGUID & """."

Dim strContentType: strContentType = Request.ServerVariables("HTTP_Content_Type")
Dim strAS2From: strAS2From = Request.ServerVariables("HTTP_AS2_From")
Dim strAS2To: strAS2To = Request.ServerVariables("HTTP_AS2_To")
Dim strMessageId: strMessageId = Request.ServerVariables("HTTP_Message_Id")
Dim strContentTransferEncoding: strContentTransferEncoding = Request.ServerVariables("HTTP_Content_Transfer_Encoding")
Dim strDispositionNotificationTo: strDispositionNotificationTo = Request.ServerVariables("HTTP_Disposition_Notification_To")
Dim strDispositionNotificationOptions: strDispositionNotificationOptions = Request.ServerVariables("HTTP_Disposition_Notification_Options")

' Dump the HTTP Headers:
Dim fsoHeaders: Set fsoHeaders = CreateObject("Scripting.FileSystemObject")
Dim stmHeaders: Set stmHeaders = fsoHeaders.CreateTextFile(Server.MapPath("..\Data\" & strGUID & ".headers.txt"), False)
stmHeaders.Write CStr(Request.ServerVariables("ALL_RAW"))
stmHeaders.Close
Set stmHeaders = Nothing
Set fsoHeaders = Nothing

'Get the HTTP payload:
If Request.TotalBytes = 0 Then
  Response.Write "ERROR: The HTTP payload is empty :-("
  Response.Status = "406 Not Acceptable"
  Response.End
End If
Dim stm: Set stm = CreateObject("ADODB.Stream")
stm.Open
stm.Type = 1 'adTypeBinary
stm.Write Request.BinaryRead(Request.TotalBytes)
stm.SaveToFile Server.MapPath("..\Data\" & strGUID & ".encrypted.txt")
Dim strData: strData = ""
If LCase(strContentTransferEncoding) = "base64" Then
  strData = CRYPTO_Decrypt(stm, "us-ascii")
Else
  strData = CRYPTO_Decrypt(stm, "unicode")
End If
Trace strGUID, "strPayloadPart = [start from the next line, and we add an additional CRLF at the end]" & vbCrLf & strData & vbCrLf

'Serialise the AS2 payload:
stm.Open
stm.Type = 2 'adTypeText
stm.Charset = "us-ascii"
stm.WriteText strData
stm.SaveToFile Server.MapPath("..\Data\" & strGUID & ".decrypted.txt")
stm.Close

'Verify the signature:
Dim nPos: nPos = 0
Trace strGUID, "nPos = " & CStr(nPos)
Dim strBoundary: strBoundary = Extract(1, strData, "boundary=""", """", nPos)
Trace strGUID, "strBoundary = """ & strBoundary & """."
Trace strGUID, "nPos = " & CStr(nPos)
Dim strPayloadPart: strPayloadPart = Extract(1, strData, _
  vbCrLf & "--" & strBoundary & vbCrLf, _
  vbCrLf & "--" & strBoundary & vbCrLf, nPos)
Trace strGUID, "strPayloadPart = [start from the next line, and we add an additional CRLF at the end]" & vbCrLf & strPayloadPart & vbCrLf
Trace strGUID, "nPos = " & CStr(nPos)
'WARNING: The "Signature Part" must be after "Payload Part", see usage of "nPos" ;-)
Dim strSignaturePart: strSignaturePart = Extract(nPos, strData, _
  vbCrLf & "--" & strBoundary & vbCrLf, _
  vbCrLf & "--" & strBoundary & "--", nPos)
Trace strGUID, "strSignaturePart = [start from the next line, and we add an additional CRLF at the end]" & vbCrLf & strSignaturePart & vbCrLf
Trace strGUID, "nPos = " & CStr(nPos)
Dim strPayLoad: strPayLoad = Extract(1, strPayloadPart, vbCrLf & vbCrLf, "", nPos)
Trace strGUID, "strPayLoad = [start from the next line, and we add an additional CRLF at the end]" & vbCrLf & strPayLoad & vbCrLf
Trace strGUID, "nPos = " & CStr(nPos)
Dim strSignature: strSignature = Extract(1, strSignaturePart, vbCrLf & vbCrLf, "", nPos)
Trace strGUID, "strSignature = [start from the next line, and we add an additional CRLF at the end]" & vbCrLf & strSignature & vbCrLf
Trace strGUID, "nPos = " & CStr(nPos)
If CRYPTO_Verify(strPayloadPart, strSignature) Then
  Trace strGUID, "The signature is verified."
Else
  Trace strGUID, "ERROR: There is a problem with the signature!"
  Response.Status = "406 Not Acceptable"
  Response.End
End If

'Serialise the AS2 payload:
stm.Open
stm.Type = 2 'adTypeText
stm.Charset = "us-ascii"
stm.WriteText strPayLoad
stm.SaveToFile Server.MapPath("..\Data\" & strGUID & ".payload.txt")
stm.Close

Dim strMyCertThumbprint: strMyCertThumbprint = ""
If strAS2To = """BabelAS2 Test Server""" Then
  strMyCertThumbprint = "67 8f a8 49 b4 7c 7c 94 8e b0 8b ab 0b e8 be fc 65 68 ab 33"
Else
  Response.Status = "406 Not Acceptable"
  Response.End
End If

'Create the MDN:
Dim strPartToBeSigned: strPartToBeSigned = _
  "Content-Type: multipart/report; report-type=disposition-notification; " & _
  "boundary=""MDNboundary""" & vbCrLf & _
  vbCrLf & _
  vbCrLf & _
  "--MDNboundary" & vbCrLf & _
  "Content-Type: text/plain; charset=""us-ascii""" & vbCrLf & _
  vbCrLf & _
  "This is an MDN." & vbCrLf & _
  vbCrLf & _
  "--MDNboundary" & vbCrLf & _
  "Content-Type: message/disposition-notification" & vbCrLf & _
  vbCrLf & _
  "Original-Recipient: rfc822;" & strAS2To & vbCrLf & _
  "Final-Recipient: rfc822;" & strAS2From & vbCrLf & _
  "Original-Message-ID: " & strMessageId & vbCrLf & _
  "Disposition: automatic-action/MDN-sent-automatically; processed" & vbCrLf & _
  "Received-Content-MIC: " & CRYPTO_SHA1(strPayloadPart) & ", sha1" & vbCrLf & _
  vbCrLf & _
  "--MDNboundary--" & vbCrLf
strBoundary = "GLOBALMDNboundary=="
Dim strMDN: strMDN = _
  "--GLOBALMDNboundary==" & vbCrLf & _
  strPartToBeSigned & _
  vbCrLf & _
  "--GLOBALMDNboundary==" & vbCrLf & _
  "Content-Type: application/pkcs7-signature; name=""smime.p7s""" & vbCrLf & _
  "Content-Disposition: attachment; filename=""smime.p7s""" & vbCrLf & _
  "Content-Transfer-Encoding: base64" & vbCrLf & _
  vbCrLf & _
  CRYPTO_Sign(strPartToBeSigned, strMyCertThumbprint) & _
  "--GLOBALMDNboundary==--" & _
  vbCrLf

Response.AddHeader "AS2-From", strAS2To
Response.AddHeader "AS2-To", strAS2From
Response.AddHeader "AS2-Version", "1.1"
Response.AddHeader "Subject", "Message Disposition Notification"
Response.AddHeader "Content-Description", "MIME Message"
Response.AddHeader "Mime-Version", "1.0"
Response.AddHeader "Message-Id", "<" & strGUID & "@BabelAS2>"
Response.ContentType = "multipart/signed; micalg=sha1; protocol=""application/pkcs7-signature""; " & _
    "boundary=""GLOBALMDNboundary=="""
Response.Write strMDN
Response.Status = "200 OK"
Response.End

Private Function Extract(nStartPos, str, strStartSearchString, strStopSearchString, ByRef nStopPos) 'As String
  Dim nLength: nLength = 0
  nStartPos = InStr(nStartPos, str, strStartSearchString)
  If nStartPos > 0 Then
    nStartPos = nStartPos + Len(strStartSearchString)
    If Len(strStopSearchString) <> 0 Then
      nLength = InStr(nStartPos, str, strStopSearchString) - nStartPos
      Extract = Mid(str, nStartPos, nLength)
    Else
      Extract = Mid(str, nStartPos)
    End If
    nStopPos = nStartPos + nLength
  End If
End Function

Private Function CRYPTO_Decrypt(stmData, strCharset) 'As String
  Dim oEnvelopedData: Set oEnvelopedData = CreateObject("CAPICOM.EnvelopedData")
  oEnvelopedData.Algorithm.Name = 3 'CAPICOM_ENCRYPTION_ALGORITHM_3DES
  stmData.Position = 0
  stmData.Type = 2 'adTypeText
  stmData.Charset = strCharset
  Dim strData: strData = stmData.ReadText
  oEnvelopedData.Decrypt strData
  stmData.Close
  stmData.Open
  stmData.Type = 2 'adTypeText
  stmData.Charset = "unicode"
  stmData.WriteText oEnvelopedData.Content
  Set oEnvelopedData = Nothing
  stmData.Position = 0
  stmData.Charset = "us-ascii"
  stmData.Position = 2
  CRYPTO_Decrypt = stmData.ReadText
  stmData.Close
End Function

Private Function CRYPTO_SHA1(strData) 'As String
  Const CAPICOM_HASH_ALGORITHM_SHA1 = 0

  Dim hash: Set hash = CreateObject("CAPICOM.HashedData")
  Dim util: Set util = CreateObject("CAPICOM.Utilities")
  Dim stm: Set stm = CreateObject("ADODB.Stream")

  stm.Open
  stm.Type = 2 'adTypeText
  stm.Charset = "us-ascii"
  stm.WriteText strData
  stm.Position = 0
  stm.Type = 1 'adTypeBinary

  hash.Algorithm = CAPICOM_HASH_ALGORITHM_SHA1
  hash.Hash stm.Read
  CRYPTO_SHA1 = util.Base64Encode(util.HexToBinary(hash.Value))
  CRYPTO_SHA1 = Left(CRYPTO_SHA1, Len(CRYPTO_SHA1)-Len(vbCrLf))

  stm.Close
  Set stm = Nothing
  Set util = Nothing
  Set hash = Nothing
End Function

Private Function CRYPTO_GetCertificate(strCertThumbprint) 'As CAPICOM.Certificate
  Const CAPICOM_LOCAL_MACHINE_STORE = 1
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

Private Function CRYPTO_Sign(strData, strCertThumbprint) 'As String
  Const CAPICOM_ENCODE_BASE64 = 0
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

Private Function CRYPTO_Verify(strData, strSignature) ' As Boolean
  Const CAPICOM_VERIFY_SIGNATURE_ONLY = 0
  Dim stm: Set stm = CreateObject("ADODB.Stream")
  stm.Open
  stm.Type = 2 'adTypeText
  stm.Charset = "us-ascii"
  stm.WriteText strData
  stm.Position = 0
  stm.Type = 1' adTypeBinary
  Dim oSignedData: Set oSignedData = CreateObject("CAPICOM.SignedData")
  oSignedData.Content = stm.Read
  On Error Resume Next
  oSignedData.Verify _
    strSignature, _
    True, _
    CAPICOM_VERIFY_SIGNATURE_ONLY
  If Err.number = 0 Then 
    CRYPTO_Verify = True
  Else
    CRYPTO_Verify = False
  End If
  On Error GoTo 0
  stm.Close
  Set stm = Nothing
  Set oSignedData = Nothing
End Function

Function CreateGUID()
  Dim tl: Set tl = CreateObject("Scriptlet.TypeLib")
  CreateGUID = Mid(tl.Guid, 2, 36)
  Set tl = Nothing
End Function

'------------------------------------------------------------------------------
Sub Trace(strGUID, strMessage)
  Dim dtNow: dtNow = Now
  Dim strResult: strResult = CStr(Year(dtNow))
  Dim str: str = CStr(Month(dtNow))
  If Len(str) = 1 Then str = "0" & str End If
  strResult = strResult & "-" & str
  str = CStr(Day(dtNow))
  If Len(str) = 1 Then str = "0" & str End If
  strResult = strResult & "-" & str
  str = CStr(Hour(dtNow))
  If Len(str) = 1 Then str = "0" & str End If
  strResult = strResult & "T" & str
  str = CStr(Minute(dtNow))
  If Len(str) = 1 Then str = "0" & str End If
  strResult = strResult & ":" & str
  str = CStr(Second(dtNow))
  If Len(str) = 1 Then str = "0" & str End If
  strResult = strResult & ":" & str

  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim stm: Set stm = fso.OpenTextFile(Server.MapPath("..\Log\" & strGUID & ".BabelAS2.log"), 8, True)
  stm.Write strResult & "|" & strMessage & vbCrLf
  stm.Close
  Set stm = Nothing
  Set fso = Nothing
End Sub
%>
