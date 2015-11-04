<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%'Response.Redirect "TempDown.htm"%>
<!--#INCLUDE FILE="include/dbConnection.asp"-->
<!--#include file="include/Functions.inc"-->
<%
'## Query parameters:
'## view = archive: show list of annexes in the right hand frame
'##        nonexisting or other value: show comments in the right hand frame
'## filename: when blank or nonexisting - return to ruledocs.asp
'##           otherwise show the file in the left hand frame
'## RPId: pass RPId (rule proposal no) to the right hand frame(s) - for annexes / comments

DIM fileName, iRPId, fso, fFile, strSrcPath, strDestFolder
On Error Resume Next
 ' On Error Goto 0       ' DEVELOPMENT & DEBUG
IdentifyUser
' strUserID = Request.Cookies("UID") ' Taken care of by IdentifyUser

If Request("fileName") = "" Then
  Response.redirect("ruleDocs.asp")
Else
  fileName = Server.HTMLEncode(Replace(Replace(Replace(Request("fileName"), "..", ""), "/", ""), "\", ""))  ' Stop short cuts to inappropriate file system levels
End If
iRPId = CLng(Request("RPId"))   ' strRPNo = Server.HTMLEncode(Request("RPNo"))

set fso = Server.CreateObject("Scripting.FileSystemObject")
strSrcPath = strDocRepository & fileName
strDestFolder = strDocBuffer & strUserID & "\"
If Not fso.FolderExists(strDestFolder) Then
  fso.CreateFolder(strDestFolder)
End If
If Err.Number <> 0 Then Call SeriousError
If fso.FileExists(strSrcPath) Then
  ' Copy the file to buffer
   ' fso.CopyFile strSrcPath, strDestFolder, True   ' Error prone: appears to copy files with different endings: "RulesContents_xxx.pdf" is copied when spec is "RulesContents.pdf" / BGJ 2007-03-06
   Set fFile = fso.GetFile(strSrcPath)
   fFile.Copy(strDestFolder)
End If
If Err.Number <> 0 Then Call SeriousError
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%=strAppTitle%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</head>
<frameset framespacing="2" frameborder="0" cols="*,470">
  <frame name="contents" target="main" src="<%= Replace(Server.URLEncode("docbuf/" & strUserID & "/" & fileName), "+", "%20")%>" scrolling="auto">
<% If Trim(Request.QueryString("view")) = "archive" Then %>
  <frame name="commentFrame" src="AnnexList.asp?RPId=<%=iRPId %>" scrolling="auto">
<% ElseIf LCase(strUserID) = "dnv_browser" Then %>
    <frame name="commentFrame" src="comments.asp?RPId=<%= iRPId %>" scrolling="auto">
<% Else %>
  <frameset rows="*,365">
    <frame name="commentFrame" src="comments.asp?RPId=<%= iRPId %>"
    scrolling="auto">
    <frame name="makecommentFrame" src="addComment.asp?RPId=<%= iRPId %>"
    scrolling="auto" noresize>
  </frameset>
<% End If %>
  <noframes>
  <body>
  <p>This page uses frames, but your browser doesn't support them.</p>
  </body>
  </noframes>
</frameset>
</html>

<%
If IsObject(fso) Then Set fso = Nothing
%>