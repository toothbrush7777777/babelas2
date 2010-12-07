<%
Option Explicit

Dim strGUID: strGUID = UTIL_CreateGUID()
UTIL_Trace strGUID, "strGUID = """ & strGUID & """."

Dim strContentType: strContentType = Request.ServerVariables("HTTP_Content_Type")
Dim strAS2From: strAS2From = Request.ServerVariables("HTTP_AS2_From")
Dim strAS2To: strAS2To = Request.ServerVariables("HTTP_AS2_To")
Dim strMessageId: strMessageId = Request.ServerVariables("HTTP_Message_Id")
Dim strContentTransferEncoding: strContentTransferEncoding = Request.ServerVariables("HTTP_Content_Transfer_Encoding")
Dim strDispositionNotificationTo: strDispositionNotificationTo = Request.ServerVariables("HTTP_Disposition_Notification_To")
Dim strDispositionNotificationOptions: strDispositionNotificationOptions = Request.ServerVariables("HTTP_Disposition_Notification_Options")

' Dump the HTTP Headers:
Dim fsoHeaders: Set fsoHeaders = CreateObject("Scripting.FileSystemObject")
Dim stmHeaders: Set stmHeaders = fsoHeaders.CreateTextFile(Server.MapPath( _
  "..\Data\" & Request.ServerVariables("REMOTE_ADDR") & "_" & strGUID & ".1.headers.txt"), False)
stmHeaders.Write CStr(Request.ServerVariables("ALL_RAW"))
stmHeaders.Close
Set stmHeaders = Nothing
Set fsoHeaders = Nothing

Dim bstrEncrypted:         bstrEncrypted = ""
Dim strEncryptedBase64:    strEncryptedBase64 = ""
Dim bstrDecrypted:         bstrDecrypted = ""
Dim bstrPayloadPart:       bstrPayloadPart = ""
Dim bstrPayload:           bstrPayload = ""
Dim bstrSignaturePart:     bstrSignaturePart = ""
Dim bstrSignatureEncoding: bstrSignatureEncoding = ""
Dim bstrSignature:         bstrSignature = ""
Dim strSignatureBase64:    strSignatureBase64 = ""

'1. Get Encrypted Data as HTTP payload:
If Request.TotalBytes = 0 Then
  Response.Write "ERROR: The HTTP payload is empty :-("
  Response.Status = "406 Not Acceptable"
  Response.End
End If
Dim util: Set util = CreateObject("CAPICOM.Utilities")
bstrEncrypted = util.ByteArrayToBinaryString(Request.BinaryRead(Request.TotalBytes))
Set util = Nothing
UTIL_SaveBinaryStringToFile bstrEncrypted, Server.MapPath( _
  "..\Data\" & Request.ServerVariables("REMOTE_ADDR") & "_" & strGUID & ".2.encrypted.txt")

'2. Decrypt:
If UCase(strContentTransferEncoding) = "BASE64" Then
  strEncryptedBase64 = U(bstrEncrypted)
  bstrDecrypted = CRYPTO_Decrypt(strEncryptedBase64)
Else
  bstrDecrypted = CRYPTO_Decrypt(bstrEncrypted)
End If
UTIL_SaveBinaryStringToFile bstrDecrypted, Server.MapPath( _
  "..\Data\" & Request.ServerVariables("REMOTE_ADDR") & "_" & strGUID & ".3.decrypted.txt")

'3. Extract the Payload Part, the Signature Part and the Signature:
Dim nPos: nPos = 0
Dim bstrBoundary: bstrBoundary = UTIL_ExtractB(1, bstrDecrypted, B("boundary="""), B(""""), nPos)
UTIL_Trace strGUID, "U(bstrBoundary) = """ & U(bstrBoundary) & """."
UTIL_Trace strGUID, "nPos = " & CStr(nPos)
bstrPayloadPart = UTIL_ExtractB(1, bstrDecrypted, _
  B(vbCrLf & "--") & bstrBoundary & B(vbCrLf), _
  B(vbCrLf & "--") & bstrBoundary & B(vbCrLf), nPos)
UTIL_Trace strGUID, "U(bstrPayloadPart) = [start from the next line, and we add an additional CRLF at the end]" & vbCrLf & U(bstrPayloadPart) & vbCrLf
UTIL_Trace strGUID, "nPos = " & CStr(nPos)
'WARNING: The "Signature Part" must be after "Payload Part", see usage of "nPos" ;-)
bstrSignaturePart = UTIL_ExtractB(nPos, bstrDecrypted, _
  B(vbCrLf & "--") & bstrBoundary & B(vbCrLf), _
  B(vbCrLf & "--") & bstrBoundary & B("--"), nPos)
UTIL_Trace strGUID, "U(bstrSignaturePart) = [start from the next line, and we add an additional CRLF at the end]" & vbCrLf & U(bstrSignaturePart) & vbCrLf
UTIL_Trace strGUID, "nPos = " & CStr(nPos)
bstrSignature = UTIL_ExtractB(1, bstrSignaturePart, B(vbCrLf & vbCrLf), "", nPos)
UTIL_Trace strGUID, "U(bstrSignature) = [start from the next line, and we add an additional CRLF at the end]" & vbCrLf & U(bstrSignature) & vbCrLf
UTIL_Trace strGUID, "nPos = " & CStr(nPos)
bstrSignatureEncoding = UTIL_ExtractB(1, bstrSignaturePart, B("Content-Transfer-Encoding: "), B(vbCrLf), nPos)
UTIL_Trace strGUID, "U(bstrSignatureEncoding) = " & U(bstrSignatureEncoding)

'4. Verify the Payload Part against the Signature:
Dim bVerified: bVerified = False
If UCase(U(bstrSignatureEncoding)) = "BASE64" Then
  strSignatureBase64 = U(bstrSignature)
  bVerified = CRYPTO_Verify(bstrPayloadPart, strSignatureBase64)
Else
  bVerified = CRYPTO_Verify(bstrPayloadPart, bstrSignature)
End If
If bVerified Then
  UTIL_Trace strGUID, "The signature is verified."
Else
  UTIL_Trace strGUID, "ERROR: There is a problem with the signature!"
  Response.Status = "406 Not Acceptable"
  Response.End
End If

'5. Extract the Payload:
bstrPayLoad = UTIL_ExtractB(1, bstrPayloadPart, B(vbCrLf & vbCrLf), "", nPos)
UTIL_SaveBinaryStringToFile bstrPayLoad, Server.MapPath( _
  "..\Data\" & Request.ServerVariables("REMOTE_ADDR") & "_" & strGUID & ".4.payload.txt")

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
  "Received-Content-MIC: " & CRYPTO_SHA1(bstrPayloadPart) & ", sha1" & vbCrLf & _
  vbCrLf & _
  "--MDNboundary--" & vbCrLf
Dim strBoundary: strBoundary = "GLOBALMDNboundary=="
Dim strMDN: strMDN = _
  "--GLOBALMDNboundary==" & vbCrLf & _
  strPartToBeSigned & _
  vbCrLf & _
  "--GLOBALMDNboundary==" & vbCrLf & _
  "Content-Type: application/pkcs7-signature; name=""smime.p7s""" & vbCrLf & _
  "Content-Disposition: attachment; filename=""smime.p7s""" & vbCrLf & _
  "Content-Transfer-Encoding: base64" & vbCrLf & _
  vbCrLf & _
  CRYPTO_Sign(B(strPartToBeSigned), strMyCertThumbprint) & _
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

'<UTIL>
Function UTIL_CreateGUID()
  Dim tl: Set tl = CreateObject("Scriptlet.TypeLib")
  UTIL_CreateGUID = Mid(tl.Guid, 2, 36)
  Set tl = Nothing
End Function

Sub UTIL_Trace(strGUID, strLine)
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
  Dim stm: Set stm = fso.OpenTextFile(Server.MapPath( _
    "..\Log\" & Request.ServerVariables("REMOTE_ADDR") & "_" & strGUID & ".BabelAS2.log"), 8, True)
  stm.Write strResult & "|" & strLine & vbCrLf
  stm.Close
  Set stm = Nothing
  Set fso = Nothing
End Sub

Function UTIL_ByteArrayToBinaryString(aobContent) 'As (binary-packed) String
  Dim oUtils: Set oUtils = CreateObject("CAPICOM.Utilities")
  UTIL_ByteArrayToBinaryString = oUtils.ByteArrayToBinaryString(aobContent)
  Set oUtils = Nothing
End Function

Sub UTIL_SaveBinaryStringToFile(bstrContent, strFileName)
  Dim oUtils: Set oUtils = CreateObject("CAPICOM.Utilities")
  Dim stm: Set stm = CreateObject("ADODB.Stream")
  stm.Open
  stm.Type = 1
  stm.Write oUtils.BinaryStringToByteArray(bstrContent)
  stm.SaveToFile strFileName
  stm.Close
  Set stm = Nothing
End Sub

Function U(bstr) ' As (unicode) String
  'Convert a binary-packed String to an unicode String:
  Dim oUtils: Set oUtils = CreateObject("CAPICOM.Utilities")
  Dim stm: Set stm = CreateObject("ADODB.Stream")
  stm.Open
  stm.Type = 1
  stm.Write oUtils.BinaryStringToByteArray(bstr)
  stm.Position = 0
  stm.Type = 2
  stm.Charset = "ascii"
  U = stm.ReadText
  stm.Close
  Set stm = Nothing
  Set oUtils = Nothing
End Function

Function B(str) 'As (binary-packed) String
  'Convert an unicode String into a binary-packed String:
  Dim stm: Set stm = CreateObject("ADODB.Stream")
  stm.Open
  stm.Type = 2
  stm.Charset = "ascii"
  stm.WriteText str
  stm.Position = 0
  stm.Type = 1
  Dim oUtils: Set oUtils = CreateObject("CAPICOM.Utilities")
  B = oUtils.ByteArrayToBinaryString(stm.Read)
  stm.Close
  Set oUtils = Nothing
  Set stm = Nothing
End Function

Function UTIL_ExtractB(nStartPos, bstr, bstrStartSearchString, bstrStopSearchString, ByRef nStopPos) 'As (binary-packed) String
  Dim nLength: nLength = 0
  nStartPos = InStrB(nStartPos, bstr, bstrStartSearchString)
  If nStartPos > 0 Then
    nStartPos = nStartPos + LenB(bstrStartSearchString)
    If LenB(bstrStopSearchString) <> 0 Then
      nLength = InStrB(nStartPos, bstr, bstrStopSearchString) - nStartPos
      UTIL_ExtractB = MidB(bstr, nStartPos, nLength)
    Else
      UTIL_ExtractB = MidB(bstr, nStartPos)
    End If
    nStopPos = nStartPos + nLength
  End If
End Function
'</UTIL>

'<CRYPTOGRAPHY>
Function CRYPTO_SHA1(bstrContent) 'As (base64 in unicode) String
  Const CAPICOM_HASH_ALGORITHM_SHA1 = 0
  Dim hash: Set hash = CreateObject("CAPICOM.HashedData")
  Dim util: Set util = CreateObject("CAPICOM.Utilities")
  hash.Algorithm = CAPICOM_HASH_ALGORITHM_SHA1
  hash.Hash bstrContent
  CRYPTO_SHA1 = util.Base64Encode(util.HexToBinary(hash.Value))
  CRYPTO_SHA1 = Left(CRYPTO_SHA1, Len(CRYPTO_SHA1)-Len(vbCrLf))
  Set util = Nothing
  Set hash = Nothing
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

Function CRYPTO_Decrypt(bstrCipher_or_strCipherBase64) 'As (binary-packed) String
  Dim oEnvelopedData: Set oEnvelopedData = CreateObject("CAPICOM.EnvelopedData")
  oEnvelopedData.Algorithm.Name = 3 'CAPICOM_ENCRYPTION_ALGORITHM_3DES
  oEnvelopedData.Decrypt bstrCipher_or_strCipherBase64
  CRYPTO_Decrypt = oEnvelopedData.Content
  Set oEnvelopedData = Nothing
End Function

Function CRYPTO_Sign(bstrContent, strCertThumbprint) 'As (base64 in unicode) String
  Const CAPICOM_ENCODE_BASE64 = 0 'We use Base64 ;-)
  Dim oSignedData: Set oSignedData = CreateObject("CAPICOM.SignedData")
  Dim oSigner: Set oSigner = CreateObject("CAPICOM.Signer")
  oSignedData.Content = bstrContent
  oSigner.Certificate = CRYPTO_GetCertificate(strCertThumbprint)
  CRYPTO_Sign = oSignedData.Sign(oSigner, True, CAPICOM_ENCODE_BASE64)
  Set oSigner = Nothing
  Set oSignedData = Nothing
End Function

Private Function CRYPTO_Verify(bstrSignedData, bstrSignature_or_strSignatureBase64) 'As Boolean
  Const CAPICOM_VERIFY_SIGNATURE_ONLY = 0
  Dim oSignedData: Set oSignedData = CreateObject("CAPICOM.SignedData")
  oSignedData.Content = bstrSignedData
  On Error Resume Next
  oSignedData.Verify _
    bstrSignature_or_strSignatureBase64, _
    True, _
    CAPICOM_VERIFY_SIGNATURE_ONLY
  If Err.number = 0 Then 
    CRYPTO_Verify = True
  Else
    CRYPTO_Verify = False
  End If
  On Error GoTo 0
  Set oSignedData = Nothing
End Function
'</CRYPTOGRAPHY>%>
