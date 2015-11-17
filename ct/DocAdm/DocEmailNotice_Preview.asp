<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
'## Previews email with notice about one or more Hearing Documents
'## to the users targeted by membership
'## Called by DocEmailNotice_Prepare.asp
'## Query parameters:
'##   chkRPId<NN>:  check boxes, <NN> is RPId
'##   txtSubject:   the email's "subject"
'##   txtEmailBody: the message text
'##   txtChosenAudiences: A ;-terminated list of chosen Hearing Bodies, e.g., "3;11;13;". 'All" -> empty list - "".
'## History:
'##   version 1.0 2007-08-27 / Bo Johanson
%>

<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->

<%
DIM dbConnect, strRPNo, strSQL_RP, rsRP, cdo, strEmailBody, iSpaces, iTab, rsHB, strSalutation
Dim strEmailSubject, iRPNo, strWarning, strRPrec, iRP, strSQLFROM, strSQLWHEREcommon
Dim strSQL_Addresses, strChosenHBs, rsRecipients, strRecipient, strAllRecipients, strSelectedHBs

On Error Resume Next
' On Error Goto 0     ' DEVELOPMENT & DEBUG

IdentifyUser
If Not bIsAdm Then Call SeriousError
Response.Clear
Response.Buffer = False  ' Allows output of "huge" lists

strWarning = ""

If Request.Form("chkRPId").Count = 0 Then
  Response.Write "No Rule Proposal was selected. Please go back and check at least one."
  Response. End
End If

SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError

' Get recipients and their email addresses
Set rsRecipients = Server.CreateObject("ADODB.Recordset")
strChosenHBs = Replace(Trim(Request.Form("txtChosenAudiences")), ";", ",")
' Response.Write "strChosenHBs=" & strChosenHBs & "<br>"        ' DEVELOPMENT & DEBUG

strSQLFROM = "FROM HearingBodies INNER JOIN " &_
                     "HB_Membership ON HearingBodies.ID = HB_Membership.HearingBodyID INNER JOIN " &_
                     "Users ON HB_Membership.UserID = Users.UserID "


If strChosenHBs = "" Then    ' "All readers in all committees" was selected   
   strSQL_Addresses = "SELECT Moderators_Email As eMailAddress From HearingBodies WHERE IsModerated = 1"
   rsRecipients.Open strSQL_Addresses, dbConnect, 3         ' Cursor adOpenStatic = 3; facilitates RecordCount and MovePrevious/Next
   If Err.Number <> 0 Then Call SeriousError
   While Not rsRecipients.EOF
      strAllRecipients = strAllRecipients & rsRecipients("eMailAddress") & "; "
      rsRecipients.MoveNext
   Wend
   rsRecipients.Close

   strSQLWHEREcommon = "(Users.eMailAddress Is Not NULL) AND (HearingBodies.IsAdministratorGroup = 0) "
   strSQL_Addresses = _
     "SELECT     Users.eMailAddress AS eMailAddress " &_
      strSQLFROM &_
     "WHERE     (Users.eMailAddress <> '=ID') AND " & strSQLWHEREcommon &_
     "UNION " &_
     "SELECT    Users.UserID AS eMailAddress " &_
      strSQLFROM &_
     "WHERE     (Users.eMailAddress = '=ID') AND " & strSQLWHEREcommon &_
     "ORDER BY eMailAddress"
   ' Response.Write "strSQL_Addresses=" & strSQL_Addresses & "<br>"        ' DEVELOPMENT & DEBUG
   rsRecipients.Open strSQL_Addresses, dbConnect, 3         ' Cursor adOpenStatic = 3; facilitates RecordCount and MovePrevious/Next
   If Err.Number <> 0 Then Call SeriousError
   While Not rsRecipients.EOF
      strAllRecipients = strAllRecipients & rsRecipients("eMailAddress") & "; "
      rsRecipients.MoveNext
   Wend
   strSalutation = "all members in all committees."  ' used in email body
   strSelectedHBs = "<br>All hearing bodies"             ' used to list hearingbodies in preview screen under email.

Else  ' A subset of Hearing Bodies was selected
   strChosenHBs = Left(strChosenHBs, Len(strChosenHBs)-1)    ' Chop the trailing ','
   strSQL_Addresses = "SELECT Moderators_Email As eMailAddress From HearingBodies WHERE IsModerated = 1 AND ID IN (" & strChosenHBs & ")"
   ' Response.Write "strSQL_Addresses=" & strSQL_Addresses & "<br>"        ' DEVELOPMENT & DEBUG
   rsRecipients.Open strSQL_Addresses, dbConnect, 3         ' Cursor adOpenStatic = 3; facilitates RecordCount and MovePrevious/Next
   If Err.Number <> 0 Then Call SeriousError
   While Not rsRecipients.EOF
      strAllRecipients = strAllRecipients & rsRecipients("eMailAddress") & "; "
      rsRecipients.MoveNext
   Wend
   rsRecipients.Close

       ' strSQLWHEREcommon = "(HearingBodies.IsModerated <> 1) AND (HearingBodies.ID IN (" & strChosenHBs & ")) "   Drop the 'IsModerated criterion. It would omit all members of "moderated" hearing bodies. /2015-0-05 /BGJ
   strSQLWHEREcommon = "(HearingBodies.ID IN (" & strChosenHBs & ")) "
   strSQL_Addresses = _
     "SELECT     Users.eMailAddress AS eMailAddress " &_
      strSQLFROM &_
     "WHERE     (Users.eMailAddress <> '=ID') AND " & strSQLWHEREcommon &_
     "UNION " &_
     "SELECT    Users.UserID AS eMailAddress " &_
      strSQLFROM &_
     "WHERE     (Users.eMailAddress = '=ID') AND " & strSQLWHEREcommon &_
     "ORDER BY eMailAddress"
   ' Response.Write "strSQL_Addresses=" & strSQL_Addresses & "<br>"        ' DEVELOPMENT & DEBUG
   rsRecipients.Open strSQL_Addresses, dbConnect, 3         ' Cursor adOpenStatic = 3; facilitates RecordCount and MovePrevious/Next
   If Err.Number <> 0 Then Call SeriousError
   While Not rsRecipients.EOF
      ' Check for invalid email addresses (RegExp), skip them and add to a separate list to be "highlighted" at the end / BGJ 2015-08-25
      strAllRecipients = strAllRecipients & rsRecipients("eMailAddress") & "; "
      rsRecipients.MoveNext
   Wend
   Set rsHB = dbConnect.Execute("SELECT NameHB from HearingBodies WHERE ID IN (" & strChosenHBs & ") ORDER BY NameHB")
   While Not rsHB.EOF
      strSalutation = strSalutation & Chr(10) & rsHB("NameHB")	  
      rsHB.MoveNext
   Wend   
   strSelectedHBs = Replace(strSalutation, Chr(10), "<br>")
   strSalutation = "members of:" & strSalutation
End If
rsRecipients.Close

strAllRecipients = Replace(strAllRecipients, ";;", ";")      ' 'moderators' can be a ;-teminated 'sublist', which would be included with an extra ; added
' Response.Write "strAllRecipients=" & strAllRecipients & "<br>"        ' DEVELOPMENT & DEBUG

'## Check if documents exist.
strSQL_RP = "SELECT * FROM RuleProps WHERE RPId IN ("

For Each iRPNo In Request.Form("chkRPId")
   strSQL_RP = strSQL_RP & CInt(iRPNo) & ","
Next

strSQL_RP = Left(strSQL_RP, Len(strSQL_RP)-1)  & ") Order By RPNo"   ' Chop the trailing ',' and close the IN bracket
' Response.Write "strSQL_RP=" & strSQL_RP & "<br>"        ' DEVELOPMENT & DEBUG
' Response.End          ' DEVELOPMENT & DEBUG


Set rsRP = Server.CreateObject("ADODB.Recordset")
rsRP.Open strSQL_RP, dbConnect, 3         ' Cursor adOpenStatic = 3; facilitates RecordCount and MovePrevious/Next
If Err.Number <> 0 Then Call SeriousError

If Request.Form("chkRPId").Count <> rsRP.RecordCount Then
  strWarning = "Not all selected Rule Proposals were found. Please review the email text."
End If

strEmailBody = Trim(Request.Form("txtEmailBody"))
strEmailBody = Replace(strEmailBody, "<<automatically suggested>>", strSalutation)
' Response.Write "strEmailBody=" & strEmailBody & "<br>"        ' DEVELOPMENT & DEBUG

strEmailSubject = Trim(Request.Form("txtSubject"))

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%=strAppTitle%> - Review Email</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<style type="text/css">
<!--
table.RPlist td {padding-right: 10px; padding-left: 10px}  
-->
</style>

<style type="text/css">
A.MenuLink {
  color: blue
  }
</style>
<script type="text/javascript" language="javascript">
function SubmitEmail(frmForm) {
  frmForm.submit();
  return true;
  }
</script>

</head>
<body style="max-width: 1000px;" topmargin="0" leftmargin="0" bgcolor="#ffffff">
<!--#include file="../include/topright.asp"-->


<form method="post" name="email" action="DocEmailNoticeAction.asp">
<table cellspacing="10" align="center">
  <tbody>
    <tr>
     <td align="center" colspan="2" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF" size="+1"><strong>Preview and edit e-mail</font></strong></td>
   </tr>
<% If strWarning <> "" Then %>
   <tr>
      <td style="padding-top: 10px; padding-bottom: 10px;"><font face="Arial"><strong><%=strWarning%></strong></font></td>
   </tr>
<% End If %>
    <tr>
      <td valign="top"><font face="Arial" size="-1"><strong>From:</strong></font>
      </td>
      <td valign="top"><input name="txtFrom" type="text" id="txtFrom" value="Rules@dnvgl.com" size="143">
      </td>
   </tr>
    <tr>
      <td>&nbsp;</td>
      <td valign="bottom"><font face="Arial" size="-1"><strong>; (semicolon) to separate multiple e-mail addresses: Bo.Johanson@dnvgl.com;Anne.Haukeland@dnvgl.com</strong></font></td>
   </tr>
    <tr>
      <td valign="top"><font face="Arial" size="-1"><strong>To:</strong></font>
      </td>
      <td valign="top"><textarea name="txtTo" cols="108" rows="5">Rules@dnvgl.com</textarea>  <!-- We need a copy ourselves. The email is not saved in Exchange Server -->
      </td>
   </tr>
    <tr>
      <td valign="top"><font face="Arial" size="-1"><strong>Cc:</strong></font>
      </td>
      <td valign="top"><textarea name="txtCc" cols="108" rows="1"></textarea>
      </td>
   </tr>
    <tr>
      <td valign="top"><font face="Arial" size="-1"><strong>Bcc:</strong></font>
      </td>
      <td valign="top"><textarea name="txtBcc" cols="108" rows="1"><%=strAllRecipients%></textarea>
      </td>
   </tr>
    <tr>
      <td valign="top"><font face="Arial" size="-1"><strong>Subject:</strong></font>
      </td>
      <td valign="top"><input name="txtSubject" type="text" id="txtSubject" value="<%=strEmailSubject%>" size="143">
      </td>
   </tr>
    <tr>
      <td valign="top"><font face="Arial" size="-1"><strong>Body:</strong></font>
      </td>
      <td valign="top"><textarea name="txtEmailBody" cols="108" rows="20"><%=strEmailBody%>
Documents:

Due date
<%
rsRP.MoveFirst
iRP = 1
While Not rsRP.EOF 
  ' This one makes too long lines:
  strRPrec = rsRP("RPNo") & "   " & rsRP("Title") & "   " & FormatDateTimeISO(rsRP("AddedDate"), false) & "   " & FormatDateTimeISO(rsRP("DueDate"), false) & Chr(13) & Chr(10)
  ' This one is a bit lengthy, with "added" and "due date":
  strRPrec = iRP & ".  " & rsRP("RPNo")  & Chr(13) & Chr(10) &_
             "    " & rsRP("Title") & Chr(13) & Chr(10) &_
             "    Added:    " & FormatDateTimeISO(rsRP("AddedDate"), false) & Chr(13) & Chr(10) &_
             "    Due Date: " & FormatDateTimeISO(rsRP("DueDate"), false) & Chr(13) & Chr(10)  & Chr(13) & Chr(10)
  ' This one is more like the 'templates' authored by Anne:
  iSpaces = Math.Max(3, 25 - Len(rsRP("RPNo")))
  ' strRPrec = rsRP("RPNo") & Space(iSpaces) & rsRP("Title") & Chr(13) & Chr(10)
  strRPrec = FormatDateTimeISO(rsRP("DueDate"), false) & "  " & rsRP("RPNo") & " "
  For iTab = 1 To Int((19 - Len(rsRP("RPNo")))/6)  ' Add TABs
     strRPrec = strRPrec & Chr(9)
  Next 
  strRPrec = strRPrec & Chr(9) & rsRP("Title") & Chr(13) & Chr(10)
  Response.Write strRPrec
%>
<% 
  rsRP.MoveNext
  iRP = iRP + 1
Wend

%>
</textarea>
      </td>
   </tr>
<tr><td colspan="2" align="center">
       <input name="btnBackToSelect" type="button" value="<< Back to Prepare" onClick="window.history.back()">
       &nbsp;&nbsp;&nbsp;
       <input name="btnSend" type="button" value="Send email >>" onClick="SubmitEmail(this.form)">
   </td></tr>
   
  </tbody>
  </table>  
  <input type="hidden" name="txtSelectedHBs" value="<%=strSelectedHBs%>">
</form>
<b>Selected Hearing bodies:</b>
<%=strSelectedHBs%>

</body>
</html>

<%
'***********************
'**  Close connection.  **
'***********************
CloseAndDiscardObject(rsRP)
CloseAndDiscardObject(rsHB)
Set cdo = Nothing
CloseAndDiscardObject(dbConnect)
%>
