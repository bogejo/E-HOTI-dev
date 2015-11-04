<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->

<%
'## Based on UpdateRuleDoc.asp
'## Query parameters:
'##   type = "annex"
'##   ID   - the identifying (index) database field value


IdentifyUser
If Not bIsAdm Then Call SeriousError
%>

<%
Dim dbConnect, iYearNow, iID, strTitle, strFilename
Dim strRPNo, strSQL, rsAnnex

iID = CLng(Request.QueryString("ID"))
Set dbConnect = Server.CreateObject("ADODB.Connection")
On Error Resume Next
' On Error Goto 0   ' DEVELOPMENT & DEBUG
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError
strSQL = "SELECT RP.RPNo As RPNo, A.AnnexTitle As Title, A.FileName As FileName FROM RPAnnex A JOIN RuleProps RP on A.RPId = RP.RPId WHERE ID = " & iID
Set rsAnnex = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError
If rsAnnex.EOF Then
  Set rsAnnex = Nothing
  dbConnect.Close
  Call NoSuchAnnex(iID)
Else
  strRPNo = rsAnnex("RPNo")
  strTitle = rsAnnex("Title")
  If IsNull(strTitle) Then strTitle = ""
  strFilename = rsAnnex("FileName")
  If IsNull(strFilename) Then strFilename = ""
End If

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%=strAppTitle%> - Update RP Annex</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</head>
<script language="JavaScript">
<!-- Hide me
  
/*************************************************************/
/* CheckLengthOfInput is Called when user leaves the field.  */
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
<body style="max-width: 1000px" bgcolor="#FFFFFF" onLoad="document.DocData.Title.focus();document.DocData.Title.select()">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">Update Minutes / Annex to Rule Proposal <%=Server.HTMLEncode(strRPNo)%></font></strong></td>
  </tr>
</table>

<form method="POST" name="DocData" action="UpdateRPAnnexAction.asp" encType="multipart/form-data">
  <input type="hidden" name="ID" value="<%=Server.HTMLEncode(iID)%>">
  <input type="hidden" name="RPNo" value="<%=Server.HTMLEncode(strRPNo)%>">
  <table border="0" cellspacing="4">
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Annex title:</td>
      <td style="font-family: Arial; font-size: 10pt"><input type="text"
      name="Title" size="85" onblur="CheckLengthOfInput(this, 100)" value="<%=Server.HTMLencode(strTitle)%>"></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt;">Annex file:</td>
      <td style="font-family: Arial; font-size: 10pt"><b><%=Server.HTMLencode(strFilename)%></b></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Optional new Annex file:</td>
      <td style="font-family: Arial; font-size: 10pt"><INPUT type="File" name="File1" size="70"></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>
        <input type="submit" value="Update database">
      </td>
    </tr>
  </table>
</form>

</body>
</html>

<%
If IsObject(dbConnect) Then
  dbConnect.Close
  Set dbConnect = Nothing
End If

Sub NoSuchAnnex(iID)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%=strAppTitle%> - Update RP Annex</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</head>
<body style="max-width: 1000px" bgcolor="#FFFFFF">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">Update an RP Annex</font></strong></td>
  </tr>
</table>
<P><font face="arial">Could not find Annex <%=iID%></font></P>
<%
Response.End
End Sub

%>