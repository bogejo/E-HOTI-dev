<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" runat="server" src="../include/md5_PAJ.js"></SCRIPT>
<%
'## Query parameters, all FORM:
'##  txtHearingBody:  Name of the Hearing Body
'##  txtHBAbbrev:     The Hearing Body's abbreviation (optional)
%>
<%
On Error Resume Next
' On Error Goto 0     ' DEVELOPMENT & DEBUG
' Global variables
Dim strAuthorisedReferer, strReferer, strHBToAdd, strHBabbrToAdd, strHBmoderatorsEmailToAdd, bIsModeratedToAdd, strModeratorName
Dim strSQLHearingBody, strSQL, rsSQL, dbConnect, strHBname, strHBabbrev, strIsModerated, strModeratorID, strIsAdminGroup

'## Check that the call comes from the authorised AddUser page
strAuthorisedReferer = Request.ServerVariables("SCRIPT_NAME")
strAuthorisedReferer = Request.ServerVariables("HTTP_HOST") & Left(strAuthorisedReferer, InStrRev(strAuthorisedReferer, "/")) & "AddHearingBody.asp"
strReferer = Request.ServerVariables("HTTP_REFERER")
If strReferer = "" Then Call SeriousError  ' Wasn't called from the proper screen

strReferer = Right(strReferer, Len(strReferer) - InStr(strReferer, "//")-1)
If StrComp(LCase(strReferer), LCase(strAuthorisedReferer)) <> 0 Then Call SeriousError  ' Wasn't called from the proper screen 

IdentifyUser
If Not bIsAdm Then Call SeriousError

'## Get the posted fields, and verify required fields.
Call CheckRequiredValue(Request.Form("txtHearingBody"),"Name of the Hearing Body")
strHBToAdd = Request.Form("txtHearingBody")
strHBabbrToAdd = Request.Form("txtHBAbbrev")
strHBmoderatorsEmailToAdd = Trim(Request.Form("txtModeratorsEmail"))
If strHBmoderatorsEmailToAdd <> "" Then
  bIsModeratedToAdd = 1
  strHBmoderatorsEmailToAdd = Replace(strHBmoderatorsEmailToAdd, ",", ";")
  If Right(strHBmoderatorsEmailToAdd, 1) <> ";" Then strHBmoderatorsEmailToAdd = strHBmoderatorsEmailToAdd & ";"  ' ;-terminated list
Else
  bIsModeratedToAdd = 0
End If

SET dbConnect = Server.CreateObject("ADODB.Connection")
If Err.Number <> 0 Then Call SeriousError
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError

'## Check whether the Hearing Body already exists.
strSQLHearingBody = "select * from HearingBodies where NameHB = " & dbText(strHBToAdd)
' Response.Write "strSQLHearingBody=" & strSQLHearingBody & "<br>"     ' DEVELOPMENT & DEBUG
Set rsSQL = dbConnect.Execute(strSQLHearingBody)
If Err.Number <> 0 Then
  Set rsSQL = Nothing
  Call SeriousError
End If

If Not rsSQL.EOF Then UpdateHBinsteadQ "HBexists"  ' UID already present. Ask whether user wishes to update it.

'## All required fields are supplied, now store this in the database:
' Response.Write "strHBToAdd=" & strHBToAdd & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strHBabbrToAdd=" & strHBabbrToAdd & ";<br>"   ' DEVELOPMENT & DEBUG

strSQL = "exec dbo.HearingBodiesInsProc " & _
      dbText(Trim(strHBabbrToAdd)) & ", " & _
      dbText(Trim(strHBToAdd)) & ", " & _
      bIsModeratedToAdd & ", " & _
      SQLAddValueOrNull(strHBmoderatorsEmailToAdd, False) & ", " & _
      "NULL, " & _
      "0" 

 ' Response.Write "strSQL=" & strSQL & "<br>"   '  DEVELOPMENT & DEBUG
 ' Response.End         ' DEVELOPMENT & DEBUG
dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError

Set rsSQL = dbConnect.Execute(strSQLHearingBody)
strHBname = rsSQL("NameHB")
If rsSQL("Abbrev") <> "" Then
  strHBabbrev = rsSQL("Abbrev")
Else
  strHBabbrev = ""
End If
If rsSQL("IsModerated") Then
  strIsModerated = "Yes"
Else
  strIsModerated = "No"
End If
'If rsSQL("Moderator_ID") <> "" Then
'  strModeratorID = rsSQL("Moderator_ID")
'Else
'  strModeratorID = ""
'End If
strModeratorName = ""
If rsSQL("Moderators_Email") <> "" Then strModeratorName = Trim(rsSQL("Moderators_Email"))

If rsSQL("IsAdministratorGroup")  = 1 Then
  strIsAdminGroup = "Yes"
Else
  strIsAdminGroup = "No"
End If
%>
<html>

<head>
<title><%=strAppTitle%> - Hearing Body Added</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<!-- link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css" -->
</head>
<body bgcolor="#FFFFFF"  style="max-width: 1000px">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">This Hearing Body has been added:</font></strong></td>
  </tr>
</table>
<table border="0" cellspacing="4">
  <tr>
    <td align="right" style="font-family: Arial; font-size: 10pt">Hearing Body:</td>
    <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strHBname)%></td>
  </tr>
  <tr>
    <td align="right" style="font-family: Arial; font-size: 10pt">Abbreviation:</td>
    <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strHBabbrev)%></td>
  </tr>
  <tr>
    <td align="right" style="font-family: Arial; font-size: 10pt">Moderated:</td>
    <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strIsModerated)%></td>
  </tr>
  <tr>
    <td align="right" style="font-family: Arial; font-size: 10pt">Moderator(s):</td>
    <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strModeratorName)%>
  </tr>
  <tr>
    <td align="right" style="font-family: Arial; font-size: 10pt">Administrator Group:</td>
    <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strIsAdminGroup)%></td>
  </tr>
</table>


<p>
<font face="Arial" size="-1">
[<a href="AddHearingBody.asp" target="_self">Add another Hearing Body</a>]&nbsp;&nbsp;&nbsp;&nbsp;
[<a href="../ruleDocs.asp" target="_self">Go to document list</a>]
</font>
</p>


</body>
</html>

<%
' Release objects from memory

'***********************
'**  Close connection.  **
'***********************
Set rsSQL = Nothing
If IsObject(dbConnect) Then dbConnect.Close
SET dbConnect = Nothing
%>


<%
'************************************
Sub UpdateHBinsteadQ (strReason)
' The Hearing Body is already present in the databae. Ask whether user wishes to update the existing Hearing Body instead.

Dim strReasonExplained, strQuery
strReasonExplained = "No explanation"

If strReason = "HBexists" Then
  strReasonExplained = "Hearing Body <b>" & Server.HTMLencode(rsSQL("NameHB")) & "</b> is already registered in the database.<br>" & _
    "Do you wish to update the existing Hearing Body <b>" & Server.HTMLencode(rsSQL("NameHB")) &  "</b>?"
End If
%>
  <html>
  <head>
  <title><%=strAppTitle%> - Hearing Body already present</title>
  <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
  <link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
  </head>
  <body bgcolor="#FFFFFF" style="max-width: 1000px">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">Already present - cannot add</font></strong></td>
  </tr>
</table>

  <p>
    <%=strReasonExplained%>
  </p>
<% If Not rsSQL.EOF Then %>
  <p>
    <a href="UpdateHearingBody.asp?UID=<%=Server.URLencode(rsSQL("NameHB"))%>">Update <%=Server.HTMLencode(rsSQL("NameHB"))%></a>
  </p>
<% End If %>
  <p>
    <a href="AddHearingBody.asp">Back to Add Hearing Body</a>
  </p>
  </body>
  </html>
<%
  Set rsSQL = Nothing
  dbConnect.Close
  Set dbConnect = Nothing
  Response.End
End Sub

%>
