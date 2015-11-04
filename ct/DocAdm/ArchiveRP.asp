<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
'## This script archives or un-archives hearing Document(s)
'## Query parameters:
'##   RPId:     The RP to archive/un-archive

'##   action:   "archive" = archive (default); "unarchive" = re-activate i.e, move from archive to active

'##   RPset:   "active" = archive (default); "archived" = re-activate i.e, move from archive to active
'##   Confirm:  when "true", executes the delete operation; otherwise displays a "confirm delete" dialog
'## Adapted to indexing tables instead of ;-terminated list for HearingBodies / 2005-07-21

%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
<head>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">

<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<!--#INCLUDE FILE="../include/popup.asp"-->

<%
DIM i, j, x, iNoOfRPs, dbConnect, iRPId, strRPNo, strDocTitle, strFileName, rsSQL, strSQL
Dim strConfirmCaption, strConfirmHeading, strConfirmPrompt, strPageTitleConfirm, strPageTitleExecuted, strCaptionExcecuted
Dim strRPidSQL, arrStrRPid, strCheckboxGroup
arrStrRPid = Array("dummy")

On Error Resume Next
' On Error Goto 0     ' DEVELOPMENT & DEBUG

IdentifyUser
If Not bIsAdm Then Call SeriousError


' DEVELOPMENT & DEBUG - BLOCK
'      Response.Write("<br><br>Request.Form.key(x) = Request.Form.item(x)<br>" )
'          for x = 1 to Request.Form.count() 
'              Response.Write(Request.Form.key(x) & " = ") 
'              Response.Write(Request.Form.item(x) & "<br>") 
'          next 
'      Response.Write("<br>")
' DEVELOPMENT & DEBUG - END OF BLOCK

'## Check if the user typed in the required fields.

strCheckboxGroup = Request.Form("checkboxGroup")
Call CheckRequiredValue(Request.Form(strCheckboxGroup)(1), "Proposal number")
' Call CheckRequiredValue(Request.Querystring("RPId"), "Rule Proposal number")

iRPId = CLng(Request.Querystring("RPId"))   ' strRPid = Request.Querystring("RPid")

strRPidSQL = " RPid IN ("
' Response.Write "Request.Form(strCheckboxGroup)=" & Request.Form(strCheckboxGroup) & "<br>"        ' DEVELOPMENT & DEBUG
arrStrRPid = Split(Request.Form(strCheckboxGroup), ", ")
For i = LBound(arrStrRPid) To UBound(arrStrRPid)
  arrStrRPid(i) = CLng(arrStrRPid(i))
  ' Response.Write "arrStrRPid(" & i & ")=" & arrStrRPid(i) & "<br>"        ' DEVELOPMENT & DEBUG
  strRPidSQL = strRPidSQL & arrStrRPid(i) & ","
Next

strRPidSQL = Left(strRPidSQL, Len(strRPidSQL) - 1) & ")"  ' Chop trailing ",", add closing bracket

'## Find the RP
strSQL = "SELECT RPId, RPNo, Title, FileName FROM RuleProps WHERE " & strRPidSQL & " ORDER BY RPNo"
' Response.Write "strSQL=" & strSQL & "<br>"        ' DEVELOPMENT & DEBUG
' Response.End          ' DEVELOPMENT & DEBUG
SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError

Set rsSQL = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then
  Call SeriousError
End If

If rsSQL.EOF Then 
  Resonse.Write "No such Hearing document(s) found"
  If IsObject(dbConnect) Then dbConnect.Close
  Response.End
End If

iNoOfRPs = 0
While Not rsSQL.EOF
  iNoOfRPs = iNoOfRPs + 1
  strRPNo = strRPNo & rsSQL("RPNo") & ", "
  strDocTitle = strDocTitle & rsSQL("Title") & ",<br>"
  strFileName = strFileName & rsSQL("FileName") & ", "
  rsSQL.MoveNext
Wend

strDocTitle = ""
strRPNo = Left(strRPNo, Len(strRPNo) - Len(", "))
strFileName = Left(strFileName, Len(strFileName) - Len(", "))


'## Archive or Un-archive?
Select Case LCase(Request.Form("RPset"))
  Case "archived"
    strPageTitleConfirm = "Un-archive " & iNoOfRPs & " Hearing Document(s)"
    strConfirmCaption = "Un-archive " & iNoOfRPs & " Hearing Document(s)?"
    strConfirmHeading = "Re-activate " & strRPNo & "?"
    strConfirmPrompt = "Click OK to put the Document(s) back on the active hearings list"
    strSQL = "UPDATE RuleProps SET DateArchived = NULL WHERE " & strRPidSQL
    strPageTitleExecuted =  iNoOfRPs &  "Proposal(s) Un-archived"
    strCaptionExcecuted = "Un-archived Proposal(s):"
  Case Else
    strPageTitleConfirm = "Archive " & iNoOfRPs & " Hearing Document(s)"
    strConfirmCaption = "Archive " & iNoOfRPs & " Hearing Document(s)?"
    strConfirmHeading = "Archive " & strRPNo & "?"
    strConfirmPrompt = "Click OK to put the Document(s) in the hearings archive"
    strSQL = "UPDATE RuleProps SET DateArchived = GETDATE() WHERE " & strRPidSQL
    strPageTitleExecuted =  iNoOfRPs & " Proposal(s) Archived"
    strCaptionExcecuted = "Archived Proposal(s):"
End Select

If LCase(Request.Form("Confirm")) = "true" Then

'## Archiving is flagged by a non NULL value for "DateArchived"
dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError
%>

<title><%=strAppTitle%> - <%=strPageTitleExecuted%></title>
</head>
<body  style="max-width: 1000px" bgcolor="#FFFFFF">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF"><%=strCaptionExcecuted%></font></strong></td>
  </tr>
</table>
<table border="0" cellspacing="4">
  <tr>
    <td align="right" style="font-family: Arial; font-size: 10pt">Proposal(s):</td>
    <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strRPNo)%></td>
  </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Document title(s):</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strDocTitle)%></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Document file(s):</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strFileName)%></td>
    </tr>
  </table>

<hr style="max-width: 1000px">
<p class="text">
[<a href="UpdateOrRemoveRuleDocs.asp" target="_self">Update another hearing document</a>]&nbsp;&nbsp;&nbsp;&nbsp;
[<a href="../ruleDocs.asp" target="_self">Go to document list</a>]
</p>

<% Else %>
<% ' "confirm" <> True; ask for confirmation %>
<title><%=strAppTitle%> - <%=strPageTitleConfirm%></title>
</head>

<body bgcolor="#FFFFFF">
<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post">
<%
For x = 1 To Request.Form.count() 
  Response.Write("<input type='Hidden' name='" & Request.Form.key(x) & "' value='" & Request.Form.item(x) &"'>")
Next 
%>  

<input type="Hidden" name="Confirm" value="true">
<%
  Response.Write "<br><br><br>"
    call popup(strConfirmCaption, _ 
               strConfirmHeading, _
               strDocTitle, _
               Request.ServerVariables("SCRIPT_NAME") & "RPset=" & Server.URLEncode(Request.Form("RPset")) & "&RPId=" & iRPId & "&Confirm=true", _
               strConfirmPrompt, _
               "center","350",False) %>
</form>

<p class="text">Go to <a href="../AdminMenu.asp">Admin menu</a></p>

<% End If %>

</body>
</html>

<%
'***********************
'**  Close connection.  **
'***********************
  Set rsSQL = Nothing
  If IsObject(dbConnect) Then dbConnect.Close
  SET dbConnect = Nothing
%>
