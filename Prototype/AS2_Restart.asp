<%
'AS2_Restart.asp, babelabout@gmail.com, 2009-11-18
'Implementation based on http://www.ietf.org/id/draft-harding-as2-restart-00.txt
Option Explicit

Dim strPrivatePath: strPrivatePath = "./"

' Dump the HTTP Headers:
Dim fsoHeaders: Set fsoHeaders = CreateObject("Scripting.FileSystemObject")
Dim stmHeaders
Set stmHeaders = fsoHeaders.CreateTextFile(Server.MapPath(strPrivatePath & CStr(Session.SessionID) & ".hdr.txt"), False)
stmHeaders.Write CStr(Request.ServerVariables("ALL_RAW"))
stmHeaders.Close
Set stmHeaders = Nothing
Set fsoHeaders = Nothing

Dim strEtag: strEtag = Request.ServerVariables("HTTP_Etag")

Select Case UCase(Request.ServerVariables("REQUEST_METHOD"))
  Case "POST"
    If Request.TotalBytes <> 0 Then
      Dim nTotalSize: nTotalSize = Request.TotalBytes
      Dim nChunckSize: nChunckSize = 1*1024 'Check of 1KB for testing
      If nChunckSize > nTotalSize Then nChunckSize = nTotalSize End If

      Dim nBytesLeft: nBytesLeft = nTotalSize 'Number of bytes left to read is initially all bytes

      Dim stmData: Set stmData = Server.CreateObject("ADODB.Stream")
      Dim fsoF: Set fsoF = CreateObject("Scripting.FileSystemObject")
      Dim stmF: Set stmF = fsoF.OpenTextFile(Server.MapPath(strEtag & ".txt"), 8, True)

      Do While nBytesLeft > 0
        If nBytesLeft < nChunckSize Then nChunckSize = nBytesLeft End If
        stmData.Open
        stmData.Type = 1 'adTypeBinary
        stmData.Write Request.BinaryRead(nChunckSize)
        nBytesLeft = nBytesLeft - nChunckSize
        stmData.Position = 0
        stmData.Type = 2 'adTypeText
        stmData.Charset = "us-ascii"
        stmF.Write stmData.ReadText
        stmData.Close
      Loop

      stmF.Close
      Set stmF = Nothing
      Set fsoF = Nothing
      Set stmData = Nothing
      Response.Status = "200 OK"
      Response.End
    Else
      Response.Write "No HTTP Data!"
      Response.Status = "406"
      Response.End
    End If
  Case "GET"

  Case "HEAD"
    Dim fsoT: Set fsoT = CreateObject("Scripting.FileSystemObject")
    Dim filT: Set filT = fsoT.GetFile(Server.MapPath(strEtag & ".txt"))
    Response.AddHeader "Content-Length", CStr(filT.Size)
    Set filT = Nothing
    Set fsoT = Nothing
    Response.Status = "200 OK"
    Response.End
  Case Else

End Select
%>