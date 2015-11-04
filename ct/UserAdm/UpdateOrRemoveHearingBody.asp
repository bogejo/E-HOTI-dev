<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
On Error Resume Next
' On Error Goto 0   ' DEVELOPMENT & DEBUG
%>
<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<%
IdentifyUser
If Not bIsAdm Then
  Call SeriousError
End If

Response.Clear
Response.Buffer = False  ' Allows output of "huge" lists

%>

<%
DIM dbConnect, rsHBs, strFormStyle, strSubTitle, strSQL
strSubTitle = "Update or remove Hearing Body"
' strFormStyle = "border-style: solid; border-color:#D0D0D0;"
strFormStyle = ""
SET dbConnect = Server.CreateObject("ADODB.Connection")
If Err.Number <> 0 Then Call SeriousError
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError
'## Get list of Hearing Bodies
strSQL = "SELECT ID, Abbrev, NameHB, IsModerated, Moderator_ID, IsAdministratorGroup FROM dbo.HearingBodies ORDER BY NameHB"
' Response.Write"strSQL=" & strSQL        ' DEVELOPMENT & DEBUG
Set rsHBs = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<head>
<title><%=strAppTitle%> - <%=strSubTitle%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</head>

<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" style="max-width: 1000px">
<!--#include file="../include/topright.asp"-->

<form name="frmUpdateOrRemoveUsers">
<table style="<%=strFormStyle%>" cellpadding="2">
  <tr>
  <td align="center" style="font-family: Arial; font-size: medium; font-weight: bold"><%=strSubTitle%></td>
  </tr>
    <td>
      <table border="0" style="font-family: Arial; font-size: 10pt">
      <tr>
        <td bgcolor="#12b1ee" align="center"><font color="#FFFFFF" size="-1"><b>Remove</b></font></td>
        <td bgcolor="#12b1ee" align="center"><font color="#FFFFFF" size="-1"><b>Update</b></font></td>
        <td bgcolor="#12b1ee" align="center"><font color="#FFFFFF" size="-1"><b>Hearing Body</b></font></td>
        <td bgcolor="#12b1ee" align="center"><font color="#FFFFFF" size="-1"><b>Abbreviation</b></font></td>
      </tr>
<%
Dim arrBgColors(2), iCnt, bgcolor
arrBgColors(0) = "#ffe1e2"
arrBgColors(1) = "#FFFFFF"
iCnt = 0
While Not rsHBs.EOF
'     Response.Write rsHBs("NameHB") & "<br>"   ' DEVELOPMENT & DEBUG
  iCnt = iCnt + 1
  bgcolor = arrBgColors(iCnt MOD 2)
%>
      <tr>
        <td valign="top" bgcolor="<%=bgcolor%>" align="center"><font size="-1"><a href="RemoveHearingBodyAction.asp?HBID=<%=Server.URLEncode(rsHBs("ID"))%>" target="_top">remove</a></font></td>
        <td valign="top" bgcolor="<%=bgcolor%>" align="center"><font size="-1"><a href="UpdateHearingBody.asp?HBID=<%=Server.URLEncode(rsHBs("ID"))%>">update</a></font></td>
        <td valign="top" bgcolor="<%=bgcolor%>"><font size="-1"><%=Server.HTMLencode(Trim(rsHBs("NameHB")))%></font></td>
        <td valign="top" bgcolor="<%=bgcolor%>"><font size="-1"><%=Server.HTMLencode(Trim(rsHBs("Abbrev")))%></font></td>
      </tr>
<%
  rsHBs.MoveNext
Wend
%>
    </table>
    </td>
  </tr>
</table>
</form>

<%
   '***********************
   '**   Close connection.   **
   '***********************
   rsHBs.Close
   SET rsHBs= Nothing
   dbConnect.Close
   SET dbConnect = Nothing
%>
</body>
</HTML>
