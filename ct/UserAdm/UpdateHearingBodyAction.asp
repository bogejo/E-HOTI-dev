<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
'## Query parameters (all Form):
'##   txtHBID
'##   txtHearingBody
'##   txtHBAbbrev
%>
<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<%
On Error Resume Next
' On Error Goto 0     ' DEVELOPMENT & DEBUG

' Global variables
Dim dbConnect, strAuthorisedReferer, strReferer, strHBIDtoUpdate, strHBNameToUpdate, strHBabbrevToUpdate, strHBmoderatorsEmailToUpdate
Dim bIsModeratedToUpdate, strSQL, rsSQL, strHBabbrev, strIsModerated, strModeratorName, strIsAdminGroup, rsModerator

'## Check that the call comes from the authorised AddUser page
strAuthorisedReferer = Request.ServerVariables("SCRIPT_NAME")
strAuthorisedReferer = Request.ServerVariables("HTTP_HOST") & Left(strAuthorisedReferer, InStrRev(strAuthorisedReferer, "/")) & "UpdateHearingBody.asp"
strReferer = Request.ServerVariables("HTTP_REFERER")
strReferer = Left(strReferer, InStrRev(strReferer, "?")-1)  ' Chop any URL query part
If strReferer = "" Then Call SeriousError  ' Wasn't called from the proper screen
strReferer = Right(strReferer, Len(strReferer) - InStr(strReferer, "//")-1)
If StrComp(LCase(strReferer), LCase(strAuthorisedReferer)) <> 0 Then Call SeriousError  ' Wasn't called from the proper screen 

IdentifyUser
If Not bIsAdm Then Call SeriousError

'## Get the posted fields, and verify required fields.
Call CheckRequiredValue(Request.Form("txtHBID"),"Hearing Body's ID")
strHBIDtoUpdate = Request.Form("txtHBID")
Call CheckRequiredValue(Request.Form("txtHearingBody"), "Hearing Body's name")
strHBNameToUpdate = Request.Form("txtHearingBody")
strHBabbrevToUpdate = Request.Form("txtHBAbbrev")
strHBmoderatorsEmailToUpdate = Trim(Request.Form("txtModeratorsEmail"))

SET dbConnect = Server.CreateObject("ADODB.Connection")
If Err.Number <> 0 Then Call SeriousError
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError

'## Verify that the HBID already exists.
strSQL = "select * from HearingBodies where ID = " & CLng(strHBIDtoUpdate)
' Response.Write "strSQL=" & strSQL & "<br>"     ' DEVELOPMENT & DEBUG
Set rsSQL = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then
  Set rsSQL = Nothing
  Call SeriousError
End If

If rsSQL.EOF Then CannotUpdateHB strHBIDtoUpdate, "noSuchID", "" , ""

'## Other HB with the same name or abbreviation?
strSQL = "select * from HearingBodies HB where HB.ID <> " & CLng(strHBIDtoUpdate) & " AND (HB.NameHB = " & dbText(strHBNameToUpdate) & " or (HB.Abbrev Is Not Null AND HB.Abbrev <> '' AND HB.Abbrev = " & dbText(strHBabbrevToUpdate) & "))"
' Response.Write "strSQL=" & strSQL & "<br>"     ' DEVELOPMENT & DEBUG
Set rsSQL = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then
  Set rsSQL = Nothing
  Call SeriousError
End If

If Not rsSQL.EOF Then CannotUpdateHB strHBIDtoUpdate, "Duplicate", strHBNameToUpdate, strHBabbrevToUpdate

'## All required fields are supplied, now store this in the database:
' Response.Write "strHBIDtoUpdate=" & strHBIDtoUpdate & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strHBNameToUpdate=" & strHBNameToUpdate & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strHBabbrevToUpdate=" & strHBabbrevToUpdate & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strHBmoderatorsEmailToUpdate=" & strHBmoderatorsEmailToUpdate & "<br>"        ' DEVELOPMENT & DEBUG
If strHBmoderatorsEmailToUpdate <> "" Then 
  bIsModeratedToUpdate = 1
  strHBmoderatorsEmailToUpdate = Replace(strHBmoderatorsEmailToUpdate, ",", ";")
  If Right(strHBmoderatorsEmailToUpdate, 1) <> ";" Then strHBmoderatorsEmailToUpdate = strHBmoderatorsEmailToUpdate & ";"  ' ;-terminated list
Else
  bIsModeratedToUpdate = 0
End If

strSQL = "exec dbo.HearingBodiesUpdProc '" & CLng(strHBIDtoUpdate) & "', " & dbText(strHBNameToUpdate) & ", "
      strSQL = strSQL & SQLAddValueOrNull(strHBabbrevToUpdate , False) & ", " & bIsModeratedToUpdate & ", " & _
               SQLAddValueOrNull(strHBmoderatorsEmailToUpdate, False)

' Response.Write "strSQL=" & strSQL & "<br>"   '  DEVELOPMENT & DEBUG
'Set rsSQL = Nothing   ' DEVELOPMENT & DEBUG
'dbConnect.Close   ' DEVELOPMENT & DEBUG
'Set dbConnect = Nothing   ' DEVELOPMENT & DEBUG
'Response.End   ' DEVELOPMENT & DEBUG

dbConnect.Execute(strSQL)
' Response.Write "dbConnect.Errors.Count=" & dbConnect.Errors.Count & "<br>"        ' DEVELOPMENT & DEBUG
If dbConnect.Errors.Count > 0 Then CannotUpdateHB strHBIDtoUpdate, "UpdateFails", "", ""

Set rsSQL = dbConnect.Execute("select * from HearingBodies where ID = " & CLng(strHBIDtoUpdate))
If Err.Number <> 0 Then
  Set rsSQL = Nothing
  Call SeriousError
End If
If IsNull(rsSQL("Abbrev")) Then
  strHBabbrev = ""
Else
  strHBabbrev = rsSQL("Abbrev")
End If

If rsSQL("IsModerated") Then
  strIsModerated = "Yes"
Else
  strIsModerated = "No"
End If
strModeratorName = ""
'If rsSQL("Moderator_ID") <> "" Then
'  Set rsModerator = dbConnect.Execute("select NameUser From Users Where UserID = " & dbText(rsSQL("Moderator_ID")))
'  If Err.Number <> 0 Then Call SeriousError
'  If Not rsModerator.EOF Then strModeratorName = rsModerator("NameUser") 
'  Set rsModerator = Nothing
'End If
If rsSQL("Moderators_Email") <> "" Then strModeratorName = Trim(rsSQL("Moderators_Email"))
If rsSQL("IsAdministratorGroup")  = 1 Then
  strIsAdminGroup = "Yes"
Else
  strIsAdminGroup = "No"
End If
%>
<html>

<head>
<title><%=strAppTitle%> - Hearing Body Updated</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<!-- link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css" -->
</head>
<body bgcolor="#FFFFFF" style="max-width: 1000px">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">Hearing Body has been updated:</font></strong></td>
  </tr>
</table>
<table border="0" cellspacing="4">
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Hearing Body:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(rsSQL("NameHB"))%></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Abbreviation:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strHBabbrev)%></td>
    </tr>
<!--
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Moderated:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strIsModerated)%>
    </tr>
-->
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Moderator(s):</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strModeratorName)%></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Admin group:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strIsAdminGroup)%></td>
    </tr>
  </table>


<P>
<font face="Arial" size="-1">
[<a href="UpdateOrRemoveHearingBody.asp" target="_self">Update or remove another Hearing Body</a>]&nbsp;&nbsp;&nbsp;&nbsp;
[<a href="../ruleDocs.asp" target="_self">Go to document list</a>]
</font>
</P>


</body>
</html>

<%
' Release objects from memory

'***********************
'**  Close connection.  **
'***********************
Set rsSQL = Nothing
Set rsModerator = Nothing
If IsObject(dbConnect) Then dbConnect.Close
SET dbConnect = Nothing
%>

<%
'************************************
Sub CannotUpdateHB (strHBID, strSymptom, strHBName, strHBabbrev)
' The Hearing Body ID was not present in the databae.

Dim strCaption, strReasonExplained

Select Case LCase(strSymptom)
  Case "duplicate"
    strCaption = "Conflicting Hearing Body particulars"
    strReasonExplained = "Another Hearing Body already exists with the same name and/or abbreviation:<br><b>" & Server.HTMLEncode(strHBName) & "</b>"
    If strHBabbrev <> "" Then strReasonExplained = strReasonExplained & " <b>(" & Server.HTMLEncode(strHBabbrev) & ")</b>"

   Case "nosuchid"
    strCaption = "Hearing Body ID not found"
    strReasonExplained = "Hearing Body ID <b>" & Server.HTMLencode(strHBID) & "</b> is not registered in the database.<br>" & _
      "Please go back and verify your input"

  Case "updatefails"
    strCaption = "Hearing Body"
    strReasonExplained = "The database refuses to update this Hearing Body<br>" & _
      "'sysadm', 'docadm' and 'DNV GL Employees' cannot be modified."

  Case Else
    strCaption = "Problem with Hearing Body particulars"
    strReasonExplained = "There is a problem with the attempted Hearing Body particulars:<br><b>" & Server.HTMLEncode(strHBName) & "</b>"
    If strHBabbrev <> "" Then strReasonExplained = strReasonExplained & " <b>(" & Server.HTMLEncode(strHBabbrev) & ")</b>"
End Select

%>
  <html>
  <head>
  <title><%=strAppTitle%> - <%=strCaption%></title>
  <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
  <link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
  </head>
  <body bgcolor="#FFFFFF" style="max-width: 1000px">
  <!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF"><%=strCaption%> - cannot update</font></strong></td>
  </tr>
</table>

  <p>
    <%=strReasonExplained%>
  </p>
  <p>
    <a href="UpdateHearingBody.asp?HBID=<%=CLng(strHBIDtoUpdate)%>">Back to Update Hearing Body</a>
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
