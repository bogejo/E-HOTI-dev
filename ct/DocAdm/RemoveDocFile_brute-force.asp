<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<% ' https://rules.dnvgl.com/rulehearing/ct/DocAdm/RemoveDocFile_brute-force.asp?filename= %>

<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/popup.asp"-->
<% 
' On Error Resume Next
On Error Goto 0         ' DEVELOPMENT & DEBUG

IdentifyUser
If Not bIsAdm Then Call SeriousError

DIM fso, strFile, strFileSpec
Dim strConsequences, strResult
strFile = Trim(LCase(Request.QueryString("filename")))
strConsequences = "The operation will remove the file"


If Request.QueryString("ConfirmDelete") <> "true" Then 
  Response.Write "<br><br><br>"
    call popup("Remove Hearing File?", _ 
               "Delete this Hearing file?", _
               strFile, _
               Request.ServerVariables("SCRIPT_NAME") & "?filename=" & strFile & "&ConfirmDelete=true", _
               strConsequences, _
               "center","350",False) %>
<!-- font face="arial" size="0" -->
<html>

<head>
<title><%=strAppTitle%> - Remove Hearing File (brute force)</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
</head>

<body bgcolor="#FFFFFF">

<p class="text">Go to <a href="../AdminMenu.asp">Admin menu</a><!--/font--></p>
<%
  Response.End
End If
%>

<%
'## Delete the file
strFileSpec = strDocRepository & strFile
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Err.Number <> 0 Then CleanUpAndQuit
 Response.Write "File=" & strFileSpec & "<br>"           ' DEVELOPMENT & DEBUG

If Not fso.FileExists(strFileSpec) Then
   strResult = strFileSpec & " was not found."
Else
   fso.DeleteFile strFileSpec
   If Err.Number <> 0 Then CleanUpAndQuit
   strResult = """" & strFileSpec & """ was removed from the Hearing document repository"
End If

%>

<html>

<head>
<title><%=strAppTitle%> - Remove Hearing File (brute force)</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
</head>

<body bgcolor="#FFFFFF">
<p><font face="arial" size="+1"><b><%=strAppTitle%></b></font></p>
<p><font face="arial"><%= Server.HTMLEncode(strResult) %></font></p>

<hr style="max-width: 1000px;">
<p class="text" style="font-size: 90%">
[<a href="../AdminMenu.asp">To Admin menu</a>]</p>
<%
  '***********************
  '**  Close connection.  **
  '***********************
  Set fso = Nothing
%>
</body>
</html>

<%
Sub CleanUpAndQuit ()
  Set fso = Nothing
  Call SeriousError
End Sub
%>