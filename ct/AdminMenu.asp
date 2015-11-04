<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="include/Functions.inc"-->
<!--#INCLUDE FILE="include/dbConnection.asp"-->
<%
On Error Resume Next
IdentifyUser
If Not bIsAdm Then Call SeriousError
%>

<html>
<head>
<title><%=strAppTitle%> - Administration</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<style type="text/css">
A.MenuLink {
  color: blue
  }
</style>

</head>

<body bgcolor="#FFFFFF" style="max-width: 1000px">
<!--#include file="include/topright.asp"-->

<table border="0" width="600">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">Administrator
    menu</font></strong></td>
  </tr>
</table>

<ul style="line-height: 1.5em"><font face="Arial" size="-1">
  <li><a class="MenuLink" href="DocAdm/AddRuleDoc.asp">Add document</a> to the Rule Hearing database</li>
  <li><a class="MenuLink" href="DocAdm/UpdateOrRemoveRuleDocs.asp">Manage hearing documents</a></li>
  <li><a class="MenuLink" href="DocAdm/DocEmailNotice_Prepare.asp">Email document availability</a></li>
  <li><a class="MenuLink" href="UserAdm/AddUser.asp">Add member</a></li>
  <li><a class="MenuLink" href="UserAdm/UpdateOrRemoveUser.asp">Update or remove member</a></li>
  <li><a class="MenuLink" href="UserAdm/ListUsers.asp?type=excel">List of members (Excel)</a>&nbsp;&nbsp;-&nbsp;&nbsp;<a class="MenuLink" href="UserAdm/ListUsers.asp">List of members (web browser)</a></li>
  <li><a class="MenuLink" href="UserAdm/AddHearingBody.asp">Add Hearing Body</a></li>
  <li><a class="MenuLink" href="UserAdm/UpdateOrRemoveHearingBody.asp">Update or remove Hearing Body</a></li>
  <li><a class="MenuLink" href="ListORScomments.asp">Extract ORS comments (Notepad "Save as / Encoding UTF-8" -> Excel)</a></li>
</font></ul>

<p><font face="Arial" size="-2">Go to <a class="MenuLink" href="ruleDocs.asp">document list</a></font></p>

</body>
</html>
