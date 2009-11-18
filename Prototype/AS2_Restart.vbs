'AS2_Restart.vbs, babelabout@gmail.com, 2009-11-18
'Implementation based on http://www.ietf.org/id/draft-harding-as2-restart-00.txt
Option Explicit

Dim strURL: strURL = "http://localhost/AS2_Restart/AS2_Restart.asp"
Dim strFolderName: strFolderName = "."
Dim strFileName: strFileName = "Message.txt"
Dim strEtag: strEtag = ""
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim nStartPos: nStartPos = 0

If fso.FileExists(strFileName) Then
  strEtag = CreateGUID()
  fso.MoveFile strFileName, strFileName & ".Etag." & strEtag
Else
  Dim fil: Set fil = Nothing
  Dim fld: Set fld = fso.GetFolder(strFolderName)
  For Each fil in fld.Files
    If Left(fil.Name, InStrRev(fil.Name, ".Etag.")) = strFileName & "." Then
      strEtag = Mid(fil.Name, InStrRev(fil.Name, ".Etag.")+Len(".Etag."))
      Exit For
    End If
  Next
  Set fld = Nothing
  If Len(strEtag) = 0 Then
    WScript.Echo "ERROR #1: :-("
    WScript.Quit 1
  Else
    Dim xhttpHEAD: Set xhttpHEAD = CreateObject("MSXML2.ServerXMLHTTP")
    xhttpHEAD.open "HEAD", strURL, False
    xhttpHEAD.setRequestHeader "Etag", strEtag
    xhttpHEAD.send
    If xhttpHEAD.status = 200 Then
      nStartPos = CLng(xhttpHEAD.getResponseHeader("Content-Length"))
      'Indicates the number of bytes of data already received from a previous send.
      Set xhttpHEAD = Nothing
    Else
      Set xhttpHEAD = Nothing
      WScript.Echo "ERROR #2: :-("
      WScript.Quit 2
    End If
  End If
End If

Dim stm: Set stm = CreateObject("ADODB.Stream")
stm.Open
stm.Type = 1 'adTypeBinary
stm.LoadFromFile strFileName & ".Etag." & strEtag
stm.Position = nStartPos

Dim xhttp: Set xhttp = CreateObject("MSXML2.ServerXMLHTTP")
xhttp.open "POST", strURL, False
xhttp.setRequestHeader "Etag", strEtag
If stm.Position > 0 Then
  xhttp.setRequestHeader "Content-Range", stm.Position & "-" & CStr(stm.Size-1) & "/" & CStr(stm.Size)
End If
xhttp.send stm

WScript.Echo _
  xhttp.status & " - " & xhttp.statusText & vbCrLf & _
  vbCrLf & _
  xhttp.responseText

stm.Close

Set xhttp = Nothing
Set stm = Nothing

fso.MoveFile strFileName & ".Etag." & strEtag, strFileName '& ".Sent", commented for testing purpose ;-)

Function CreateGUID()
  Dim tl: Set tl = CreateObject("Scriptlet.TypeLib")
  CreateGUID = Mid(tl.Guid, 2, 36)
  Set tl = Nothing
End Function