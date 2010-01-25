<% Option Explicit

Const URL = "http://localhost/BabelAS2/Receive.asp"

Dim xhttp: Set xhttp = CreateObject("MSXML2.ServerXMLHTTP.4.0")
xhttp.open "POST", URL, False

Dim strLine: strLine = ""
Dim strHeader: strHeader = ""
Dim strValue: strValue = ""

'----- HTTP Request Headers
For Each strLine In Split(Request.ServerVariables("ALL_RAW"), vbCrLf, -1, 1)
  If Len(strLine) <> 0 Then
    strHeader = Left(strLine, InStr(strLine, ": ")-1)
    strValue = Mid(strLine, InStr(strLine, ": ")+2)
    xhttp.setRequestHeader strHeader, strValue
  End If
Next
xhttp.setRequestHeader "Reverse-Proxy-Source-IP-Address", Request.ServerVariables("REMOTE_ADDR")
'----- HTTP Request Payload
If Request.TotalBytes <> 0 Then
  xhttp.send Request.BinaryRead(Request.TotalBytes)
Else
  xhttp.send
End If
'----- HTTP Response Headers
For Each strLine In Split(xhttp.getAllResponseHeaders, vbCrLf, -1, 1)
  If Len(strLine) <> 0 Then
    strHeader = Left(strLine, InStr(strLine, ": ")-1)
    strValue = Mid(strLine, InStr(strLine, ": ")+2)
    Response.AddHeader strHeader, strValue
  Else
  End If
Next
'----- HTTP Response Payload
Response.BinaryWrite xhttp.responseBody
'----- HTPP Response Status
Response.Status = xhttp.status & " " & xhttp.statusText

Set xhttp = Nothing
%>
