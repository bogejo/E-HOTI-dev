<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<%
'## Query parameters:
'##  RPId - the parent PR (internal ID)
'##  RPNo - the parent PR (public reference)
'## 
'## Based on AddRuleDoc.asp

On Error Resume Next
' On Error GoTo 0     ' DEVELOPMENT & DEBUG

IdentifyUser
If Not bIsAdm Then Call SeriousError
%>

<%
Dim iYearNow, iYr
iYearNow = Year(Date)

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%=strAppTitle%> - Add Minutes / Annex</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</head>
<script language="JavaScript">
<!-- Hide me
  
/*************************************************************/
/* CheckLengthOfInput is called when user leaves the field.  */
/* Checks if the length of the filed is not exceeded.        */
/*************************************************************/
function CheckLengthOfInput(InputObject, MaxLength)
{
  var InputString = InputObject.value;
  if(InputString.length > MaxLength)
  {
    alert("The value you entered was too long. Max length is " + MaxLength + ".");
    InputObject.select();
    InputObject.focus();
  }
}

//End -->
</script>
<body style="max-width: 1000px" bgcolor="#FFFFFF" onLoad="document.RPAnnexdata.AnnexTitle.focus();document.RPAnnexdata.AnnexTitle.select()">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">Add Minutes / Annex to Rule Proposal <%=Server.HTMLEncode(Request.QueryString("RPNo"))%></font></strong></td>
  </tr>
</table>

<form method="POST" name="RPAnnexdata" action="addRPAnnexAction.asp" encType="multipart/form-data">
  <input type="hidden" name="RPId" value="<%=CLng(Request.QueryString("RPId"))%>">
  <table border="0" cellspacing="4">
    <tr>
      <td style="font-family: Arial; font-size: 10pt;" align="right">Annex title:</td>
      <td style="font-family: Arial; font-size: 10pt;"><input name="AnnexTitle" size="85" onblur="CheckLengthOfInput(this, 100)" type="text" value="1. Minutes of Design Review yyyy-mm-dd"></td>
    </tr>
    <tr>
      <td style="font-family: Arial; font-size: 10pt;" align="right">Annex 
      file:</td>
      <td style="font-family: Arial; font-size: 10pt"><INPUT type="File" name="File1" size="70"></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>
        <input type="submit" value="Add to database">
      </td>
    </tr>
  </table>
</form>
</body>
</html>
