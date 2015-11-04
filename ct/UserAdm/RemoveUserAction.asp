<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
'## Query parameters (all QueryString)
'##   ConfirmDelete: "true" deletes the UserID, otherwise user is prompted to confirm deletion
'##   UID:           the UserID to be deleted
'## Adapted to indexing tables for HearingBodies / 2005-07-21
%>
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/popup.asp"-->
<%
On Error Resume Next
 On Error Goto 0         ' DEVELOPMENT & DEBUG

IdentifyUser
If Not bIsAdm Then Call SeriousError
%>
<%  
DIM dbConnect, strSQL, rsData, strResult

SET dbConnect = Server.CreateObject("ADODB.Connection")
If Err.Number <> 0 Then Call SeriousError
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError

If Request.QueryString("ConfirmDelete") <> "true" Then 
  strResult = ""
  Set rsData = dbConnect.Execute("exec UsersSelProc " & dbText(Request.QueryString("UID")))
  If Err.Number <> 0 Then CleanUpAndQuit
  If Not rsData.EOF Then
    Response.Write "<br><br><br>"
      call popup("Remove UserID", _ 
                 "Remove this UserID?", _
                 CleanUIDinput(Request.QueryString("UID")), _
                 Request.ServerVariables("SCRIPT_NAME") & "?UID=" & Server.URLencode(Request.QueryString("UID")) & "&ConfirmDelete=true", _
                 "", _
                 "center","300",True)
  Else
    strResult = "<p class='text'><b><font color='red'>Failure:</font></b> Cannot find User ID <b>" & Server.HTMLencode(Request.QueryString("UID")) & "</b> in the database.</b></p>"
  End If 
 %>
<html>

<head>
<title><%=strAppTitle%> - Remove Member</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
</head>

<body bgcolor="#FFFFFF">
<%=strResult%>
<p class="text">To <a href="../AdminMenu.asp">administrator menu</a></p>
<%
  Set rsData = Nothing
  dbConnect.Close
  SET dbConnect = Nothing
  Response.End
End If
%>

<%
' Request.QueryString("ConfirmDelete") = "true", so delete the userID, subject to restrictions
If (InStr(Request.QueryString("UID"), "Member_") = 1) OR (LCase(Request.QueryString("UID")) = "dnvgl_browser") Then
  strResult = "<b><font color='red'>Failure:</font></b> Cannot remove 'collective' User ID <b>" & Server.HTMLencode(Request.QueryString("UID")) & "</b>"
Else
  '## Delete User ID
  dbConnect.Execute "UsersDelProc " & dbText(Request.QueryString("UID"))
  If Err.Number <> 0 Then CleanUpAndQuit
  Set rsData = dbConnect.Execute("exec UsersSelProc " & dbText(Request.QueryString("UID")))
  If Err.Number <> 0 Then CleanUpAndQuit
  If rsData.EOF Then
    strResult = "<b>Removed User ID " & Server.HTMLencode(Request.QueryString("UID")) & "</b>"
  Else
    strResult = "<b><font color='red'>Failure:</font></b> Could not remove User ID <b>" & Server.HTMLencode(Request.QueryString("UID")) & "</b><br>" & _
                "There may be comments from the member in the database.<br>Cannot remove a member who has commented on a document."
  End If
End If
%>

<html>

<head>
<title><%=strAppTitle%> - Remove Member</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
</head>

<body bgcolor="#FFFFFF">
<p><font face="arial" size="+1"><b><%=strAppTitle%></b></font></p>
<p><font face="arial"><%=strResult%></font></p>

<p class="text">
[<a href="UpdateOrRemoveUser.asp">Update or remove another member</a>]&nbsp;&nbsp;&nbsp;&nbsp;
[<a href="../AdminMenu.asp">To Admin menu</a>]</p>
<%
  '***********************
  '**  Close connection.  **
  '***********************
  dbConnect.Close
  SET dbConnect = Nothing
%>
</body>
</html>

<%
Sub CleanUpAndQuit ()
  Set rsData = Nothing
  If IsObject(dbConnect) Then dbConnect.Close
  SET dbConnect = Nothing
  Call SeriousError
End Sub
%>