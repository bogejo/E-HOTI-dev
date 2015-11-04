<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
'## Query parameters (all QueryString)
'##   ConfirmDelete: "true" deletes the Hearing Body, otherwise user is prompted to confirm deletion
'##   HBID:           the ID of the Hearing Body to be deleted
'## Adapted to indexing tables for HearingBodies / 2005-07-21
%>
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/popup.asp"-->
<%
On Error Resume Next
' On Error Goto 0         ' DEVELOPMENT & DEBUG

IdentifyUser
If Not bIsAdm Then Call SeriousError
%>
<%  
DIM dbConnect, strSQL, rsData, strResult, strAuthorisedReferer, strReferer, strNameHB

'## Check that the call comes from the authorised UpdateOrRemoveHearingBody page
'strAuthorisedReferer = Request.ServerVariables("SCRIPT_NAME")
'strAuthorisedReferer = Request.ServerVariables("HTTP_HOST") & Left(strAuthorisedReferer, InStrRev(strAuthorisedReferer, "/")) & "UpdateOrRemoveHearingBody.asp"
'strReferer = Request.ServerVariables("HTTP_REFERER")
'If strReferer = "" Then Call SeriousError  ' Wasn't called from the proper screen

'strReferer = Right(strReferer, Len(strReferer) - InStr(strReferer, "//")-1)
'If StrComp(LCase(strReferer), LCase(strAuthorisedReferer)) <> 0 Then Call SeriousError  ' Wasn't called from the proper screen 



SET dbConnect = Server.CreateObject("ADODB.Connection")
If Err.Number <> 0 Then Call SeriousError
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError

Set rsData = dbConnect.Execute("select * from HearingBodies where ID = " & Request.QueryString("HBID"))
If Err.Number <> 0 Then CleanUpAndQuit
If Not rsData.EOF Then
  strNameHB = Trim(rsData("NameHB"))
Else
  strNameHB = ""
End If

If Request.QueryString("ConfirmDelete") <> "true" Then 
  strResult = ""
  If strNameHB <> "" Then
    Response.Write "<br><br><br>"
      call popup("Remove Hearing Body", _ 
                 "Remove this Hearing Body?", _
                 strNameHB, _
                 Request.ServerVariables("SCRIPT_NAME") & "?HBID=" & Server.URLencode(Request.QueryString("HBID")) & "&ConfirmDelete=true", _
                 "", _
                 "center","300",False)
  Else
    strResult = "<p class='text'><b><font color='red'>Failure:</font></b> Cannot find Hearing Body ID <b>" & Server.HTMLencode(Request.QueryString("HBID")) & "</b> in the database.</b></p>"
  End If 
 %>
<html>

<head>
<title><%=strAppTitle%> - Remove Hearing Body</title>
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
' Request.QueryString("ConfirmDelete") = "true", so delete the Hearing Body, subject to restrictions
'## Delete Hearing Body
dbConnect.Execute "HearingBodiesDelProc " & dbText(Request.QueryString("HBID"))
If Err.Number <> 0 Then CleanUpAndQuit
Set rsData = dbConnect.Execute("select * from HearingBodies where ID = " & Request.QueryString("HBID"))
If Err.Number <> 0 Then CleanUpAndQuit
If rsData.EOF Then
  strResult = "<b>Removed Hearing Body: '" & Server.HTMLencode(strNameHB) & "'</b>"
Else
  strResult = "<b><font color='red'>Failure:</font></b> Could not remove Hearing Body <b>'" & Server.HTMLencode(strNameHB) & "'</b><br>" & _
              "'sysadm', 'docadm' and 'DNV GL Employees' cannot be deleted."
End If
%>

<html>

<head>
<title><%=strAppTitle%> - Remove Hearing Body</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
</head>

<body bgcolor="#FFFFFF">
<p><font face="arial" size="+1"><b><%=strAppTitle%></b></font></p>
<p><font face="arial"><%=strResult%></font></p>

<p class="text">
[<a href="UpdateOrRemoveHearingBody.asp">Update or remove another Hearing Body</a>]&nbsp;&nbsp;&nbsp;&nbsp;
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