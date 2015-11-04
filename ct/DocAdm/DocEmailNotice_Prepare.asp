<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
'## Prepare an email with notice about one or more Hearing Documents
'## History:
'##   version 1.0 2007-08-27 / Bo Johanson
%>
<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<%
 On Error Resume Next
' On Error GoTo 0        ' DEVELOPMENT & DEBUG

IdentifyUser
If Not bIsAdm Then Call SeriousError

Response.Clear
Response.Buffer = False  ' Allows output of "huge" lists

DIM dbConnect, rsActiveProps, strSQL, fFirstPass, colorTR
Dim iCurrentRPId, strAddedDate, strDueDate, strRestrictedTo, strRestrictedToTemp
Dim iPosLFirstSlash, strSite, strLoginURL

' Response.Write "Request.ServerVariables(""SCRIPT_NAME"")=" & Request.ServerVariables("SCRIPT_NAME") & "<br>"        ' DEVELOPMENT & DEBUG
iPosLFirstSlash = InStr(2, Request.ServerVariables("SCRIPT_NAME"), "/")
strSite = Left(Request.ServerVariables("SCRIPT_NAME"), iPosLFirstSlash)
strLoginURL = "https://" & Request.ServerVariables("SERVER_NAME") & strSite
' Response.Write "strLoginURL=" & strLoginURL & "<br>"        ' DEVELOPMENT & DEBUG
' Response.End        ' DEVELOPMENT & DEBUG


SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError


'## Get the active RPs
strSQL = "SELECT RuleProps.RPId, RuleProps.RPNo, RuleProps.Title, RuleProps.AddedDate, RuleProps.DueDate, RuleProps.FileName, HearingBodies.Abbrev AS Restricted_To_Abbrev, HearingBodies.NameHB AS Restricted_To_NameHB  " & _
"FROM HearingBodies INNER JOIN RestrictedRuleProps ON HearingBodies.ID = RestrictedRuleProps.HearingBodyID RIGHT OUTER JOIN RuleProps ON RestrictedRuleProps.RPId = RuleProps.RPId " & _
"WHERE     (RuleProps.DateArchived IS NULL) " & _
"ORDER BY RPNo"

' Response.Write "strSQL=" & strSQL & "<br>"        ' DEVELOPMENT & DEBUG
' Response.End         ' DEVELOPMENT & DEBUG

Set rsActiveProps = Server.CreateObject("ADODB.Recordset")
rsActiveProps.Open strSQL, dbConnect, 3         ' Cursor adOpenStatic = 3; facilitates RecordCount and MovePrevious/Next
If Err.Number <> 0 Then Call SeriousError
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head><title><%=strAppTitle%> - Prepare E-mail Notice</title>

<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<!--#include file="../include/SelectAudiencesSetup.inc"-->
<!--#include file="../include/SelectAudiencesExecute.inc"-->

<style type="text/css">
  table.RPlist td {padding-right: 10px; padding-left: 10px;}
  A.MenuLink {
    color: blue;
    }
</style>

<script type="text/javascript" language="javascript">
function initialize() {
  nodeAudience = document.getElementById("idAudienceDummyNode");    // Must assign (a dummy) input node to 'nodeAudience'
  inpChosenAudiences = document.forms.Email_Spec.txtChosenAudiences;
  }

function SubmitForPreview(frmForm) {
  fnUpdateChecked(self, frmForm);
  frmForm.submit();
  return true;
  }

</script>


</head>
<body style="max-width: 1000px;" topmargin="0" leftmargin="0" bgcolor="#ffffff" onLoad="initialize()";>
<!--#include file="../include/topright.asp"-->

<table cellspacing="10" width="100%">
  <tbody>
    <tr>
    <td align="center"><font face="Arial" size="+2"><strong>Prepare E-mail notice</strong></font></td>
  </tr>
  </tbody>
</table>

<%
If rsActiveProps.EOF Then %>
  <p>There are no active hearing documents</p>
  <%
  Response.End
End If

iCurrentRPId  = -1
rsActiveProps.MoveFirst
%>

<form name="Email_Spec" method="post" action="DocEmailNotice_Preview.asp">
  <input type="Hidden" name="txtAudienceDummyNode" id="idAudienceDummyNode"><!-- referred (indirectly) by fnUpdateChecked() -->

<table cellspacing="10" width="100%">
  <tbody>
    <tr>
    <td colspan="6" bgcolor="#12b1ee"><font face="Arial" color="#FFFFFF" size="+1"><strong>1. Select documents for e-mail</strong></font></td>
  </tr>
<tr>
   <td>
     <table class="RPlist" style="font-family: Arial; font-size: 10pt;" border="0" width="100%">
        <thead>
          <tr> 
            <td align="center" bgcolor="#12b1ee">&nbsp;</td>
            <td align="center" bgcolor="#12b1ee"><b><font color="#ffffff" size="-1">RP 
              No.</font></b></td>
            <td bgcolor="#12b1ee"><font color="#ffffff" size="-1"><b>Rule 
              Proposal Title</b></font></td>
            <td align="center" bgcolor="#12b1ee"><font color="#ffffff" size="-1"><b>Added</b></font></td>
            <td align="center" bgcolor="#12b1ee"><font color="#ffffff" size="-1"><b>Due 
              Date</b></font></td>
            <td align="center" bgcolor="#12b1ee" nowrap="nowrap"><font color="#ffffff" size="-1"><b>Restricted 
              to</b></font> </td>
          </tr>
        </thead>
        <tbody>

<%
colorTR = ""

Do While Not rsActiveProps.EOF
  iCurrentRPId = rsActiveProps("RPId")
  If colorTR = "#FFFFFF" Then
    colorTR = "#ffe1e2"
  Else
    colorTR = "#FFFFFF"
  End If 

  strAddedDate = FormatDateTimeISO(rsActiveProps("AddedDate"), False)
  strDueDate = FormatDateTimeISO(rsActiveProps("DueDate"), False)
%>
          <tr> 
<!--            <td bgcolor="<%=colorTR%>" nowrap="nowrap"><input type="checkbox" name="chkRPId<%=iCurrentRPId%>" value="<%=iCurrentRPId%>"></td>  -->
            <td bgcolor="<%=colorTR%>" nowrap="nowrap"><input type="checkbox" name="chkRPId" value="<%=iCurrentRPId%>"></td>
            <td bgcolor="<%=colorTR%>" nowrap="nowrap"><a href="../review.asp?fileName=<%= rsActiveProps("FileName") %>&amp;RPId=<%= rsActiveProps("RPId") %>" target="Review"><%=Server.HTMLEncode(rsActiveProps("RPNo"))%></a></td>
            <td bgcolor="<%=colorTR%>"><%=Server.HTMLEncode(rsActiveProps("Title"))%></td>
            <td bgcolor="<%=colorTR%>" align="center" nowrap="nowrap"><%=strAddedDate%></td>
            <td bgcolor="<%=colorTR%>" align="center" nowrap="nowrap"><%=strDueDate%></td>
<%
  If Not IsNull(rsActiveProps("Restricted_To_NameHB")) Then
    If IsNull(rsActiveProps("Restricted_To_Abbrev")) Then 
      strRestrictedToTemp = rsActiveProps("Restricted_To_NameHB")
    Else
      strRestrictedToTemp = rsActiveProps("Restricted_To_Abbrev")
    End If
    strRestrictedTo = Trim(strRestrictedToTemp)
    rsActiveProps.MoveNext  ' Bundle all "Restricted_To" for this RPId
    Do While Not rsActiveProps.EOF
      If rsActiveProps("RPId") <> iCurrentRPId Then Exit Do
      If IsNull(rsActiveProps("Restricted_To_Abbrev")) Then 
        strRestrictedToTemp = rsActiveProps("Restricted_To_NameHB")
      Else
        strRestrictedToTemp = rsActiveProps("Restricted_To_Abbrev")
      End If
      strRestrictedTo = strRestrictedTo & "; " & Trim(strRestrictedToTemp)
      rsActiveProps.MoveNext
    Loop
    rsActiveProps.MovePrevious  ' Back up, since the loop has overshot
  Else
    strRestrictedTo = ""
  End If
%>
                <td bgcolor="<%=colorTR%>"><%=Server.HTMLEncode(strRestrictedTo)%></td>
              </tr>

<% 
  rsActiveProps.MoveNext
Loop 
%>
            </tbody>
          </table>

    </td>
  </tr>
</tbody>
</table>

<% ' 2007-07-13: Works up to here, exept haven't tested when there are no active hearings %>

<p></p>
<table cellspacing="10" width="100%">
  <tbody>
<tr><td bgcolor="#12b1ee"><font face="Arial" color="#FFFFFF" size="+1"><strong>2. Message:</strong></font>
</td></tr>
<tr>
   <td><font face="arial" size="-1"><strong>Subject:</strong></font>&nbsp;&nbsp;&nbsp;
      <input name="txtSubject" type="text" size="100" value="DNV GL Proposals on External Hearing"></td></tr>
<tr><td>
<textarea name="txtEmailBody" cols="100" rows="10">Dear hearing participant,
As a valued contributor you are invited to comment on the hearing for the proposal listed below.

The proposals listing shows the closing date(s) for the hearing.

The proposals may be reviewed and commented on through our "Hearing on the Internet" site.
Please use the following link: <%=strLoginURL%>

Please sign on with your userID (e-mail address as used in this email) and password.
If this is the first time you use DNV GL's "Hearing on the Internet", the initial password is:
DnV140RS

The initial password must be changed after first log on.


User Information:
Open the site <%=strLoginURL%> with your userID (e-mail address as used in this email) and password. 
Recipients of this e-mail are already registered as users of this hearing site - any additional
reviewers must be registered by DNV GL's unit 'Rules and Standards Publishing House' (rules@dnvgl.com).

Choose the document you are to review from the document list.

The window opens with the document on the left-hand side and the comments field
on the right hand side.  
It is possible to view the document in its own window - see link top right.
In the comments box it is important to enter the exact reference location to each individual
comment and then click the submit button. It is also possible to select "general" here.
If you have reviewed the document but have no comments, please click on the
"Reviewed no comment" button.



Best regards,
for DNV GL AS
_______________________________ 
Sille Grjotheim
Rules and Standards Publishing House, MCCNO831
mailto:rules@dnvgl.com

</textarea>
</td></tr>
<!--
<tr><td style="text-align:center;"><input name="btnSubmit" type="button" value="Preview" onClick="window.location='Email_Preview_proto.htm'">
</td></tr>
-->
  </tbody>
</table>
<p></p>
<table cellspacing="10">
  <tbody>
    <tr>
    <td align="center" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF" size="+1"><strong>3. Select recipients</font></strong></td>
  </tr>
</table>
<p></p>
  <input type="hidden" name="txtChosenAudiences">
  <table border="0" cellpadding="2" cellspacing="0" style="margin-left:10pt; font-family: Arial; font-size: 10pt;">
    <tbody><tr><td><input onclick="exclusiveAud(this)" name="chkAudExclusiveAll" value="All" checked="checked" type="checkbox"></td><td>All readers in all committees</td></tr>
    <tr>
      <td colspan="2"><b>Members of these hearing bodies:</b></td></tr>
<% ListAudiences %>

  </tbody></table>
<p align="center"><br><input name="btnSubmit" type="button" value="Preview message" onClick="SubmitForPreview(this.form)"></p>

</form>

</body></html>

<%
SelectAudiencesCleanUp
If IsObject(rsActiveProps) Then CloseAndDiscardObject(rsActiveProps)
%>