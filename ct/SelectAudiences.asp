<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<% Response.Buffer = True 'Buffers the content so Response.Redirect will work %>
<!--#include file="include/Functions.inc"-->
<!--#INCLUDE FILE="include/dbConnection.asp"-->
<%
On Error Resume Next
' On Error GoTo 0        ' DEVELOPMENT & DEBUG

DIM dbConnect, strSqlHearingsBodies, rsHearingsBodiesByName, iAudNo, strHB, strHBabbr
Dim strReferer, strHeading

strReferer = Request.ServerVariables("HTTP_REFERER")

If InStr(LCase(strReferer), "/docadm/") > 0 Then
   strHeading = "Restrict document access to these audiences:"
Else
   strHeading = "The comment may be read by members of these audiences:"
End If
' Check whether Signed in - Pending

SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError
strSqlHearingsBodies = "SELECT ID, NameHB, Abbrev FROM HearingBodies WHERE IsAdministratorGroup = 0 ORDER BY NameHB"
SET rsHearingsBodiesByName = dbConnect.Execute(strSqlHearingsBodies)
If Err.Number <> 0 Then Call SeriousError
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<HEAD>
<META NAME="author" CONTENT="Bo Johanson">
<TITLE><%=strAppTitle%> - Select Audiences</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link REL="STYLESHEET" HREF="include/main.css" TYPE="text/css">
<STYLE TYPE="text/css">
<!--
A:visited
{
    COLOR: blue;
}
-->
</STYLE>
<!--#include file="include/SelectAudiencesExecute.inc"-->
<%
strChkExclusiveGroup = "chkAudExclusive"
strPreChosen = Request("PreChosen")
If strPreChosen = "" Then strPreChosen = "All"
%>
</HEAD>

<body bgcolor="#FFFFE0" onLoad="javascript:window.focus()">
<P class="title2">
<%=strHeading%>
<P>
<P class="subtitle">
Check boxes as required<br>
</P>
<form name="frmAudience">
  <table BORDER=0 CELLSPACING=0 CELLPADDING=2>
    <tr><td><input TYPE='Checkbox' onClick='exclusiveAud(this)' NAME='<%=strChkExclusiveGroup%>All' VALUE='All' <% If bIsInList(strPreChosen, "All", ";", true, false) Then Response.Write " checked" %>></td><td><%=strAudienceAll%></td></tr>
    <tr><td><input TYPE='Checkbox' onClick='exclusiveAud(this)' NAME='<%=strChkExclusiveGroup%>DNV GL' VALUE='DNV' <% If bIsInList(strPreChosen, "DNV", ";", true, false) Then Response.Write " checked" %>></td><td><%=strAudienceDNV%></td></tr>
    <tr>
      <td colspan="2"><b>Members of these hearing bodies:</b></td></tr>
<% 
Do While Not rsHearingsBodiesByName.EOF
  iAudNo = rsHearingsBodiesByName("ID")
  strHB = Trim(rsHearingsBodiesByName("NameHB"))
  strHBabbr = Trim(rsHearingsBodiesByName("Abbrev"))
  If strHBabbr <> "" Then strHB = strHB & "&nbsp;(" & strHBabbr & ")"
%>    <tr><td><input TYPE='Checkbox' onClick='excludeExclusiveAuds(this)' NAME='chkAudno<%=iAudNo%>' VALUE='<%=iAudNo%>' <% If bIsInList(strPreChosen, iAudNo, ";", true, false) Then Response.Write " checked" %>></td><td><%=strHB%></td></tr>
<%
  rsHearingsBodiesByName.MoveNext
Loop %>
  </table>
  <table>
    <tr>
      <td>
        <INPUT TYPE="Button" NAME="btnUpdateChecked" VALUE="Save and close" onClick="fnUpdateChecked(window.opener, this.form)"></td><td><INPUT TYPE="Button" NAME="btnCancel" VALUE="Cancel" onClick="window.close();">
      </td>
    </tr>
  </table>
</form>
</BODY>
</HTML>

<% 
SelectAudiencesCleanUp
%>
