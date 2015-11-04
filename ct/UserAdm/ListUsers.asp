<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
'## ListUsers.asp
'## Query parameters:
'##   type: 'excel', when requesting to open in MS Excel
%>
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
%>

<%
DIM dbConnect, rsUsers, strFormStyle, strSubTitle, strSQL, strTableRowStyle, strTableHeadStyle

Response.Clear
Response.Buffer = False  ' Allows output of this now "huge" list of users, more than 7.200 members

If LCase(Request.QueryString("type")) = "excel" Then
  Response.ContentType = "application/vnd.ms-excel"
  strTableHeadStyle = "style='font-size: x-small; color: #FFFFFF; font-weight: bold;'"
  strTableRowStyle = "style='font-size: x-small;'"
Else
  strTableHeadStyle = "style='font-size: small; color: #FFFFFF; font-weight: bold;'"
  strTableRowStyle = "style='font-size: small;'"
End If

strSubTitle = "List of members"
' strFormStyle = "border-style: solid; border-color:#D0D0D0;"
strFormStyle = ""
SET dbConnect = Server.CreateObject("ADODB.Connection")
If Err.Number <> 0 Then Call SeriousError
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError
'## Get list of users
strSQL = "SELECT U.UserID, U.NameUser as Name, " &_
            "case " &_
            "  when U.eMailAddress is NULL then '' " &_
            "  when U.eMailAddress = '=ID' then U.UserID " &_
            "  else U.eMailAddress " &_
            "  end " &_
            "AS eMailAddress, " &_
            "case " &_
            "  when U.Organisation is NULL then '' " &_
            "  else U.Organisation " &_
            "  end " &_
            "as Organisation, " &_
            "case " &_
            "  when HBM.UserID Is NULL Then '' " &_
            "  else HB.NameHB " &_
            "  end " &_
            "as [Hearing Body]," &_
            "case " &_
            "  when HB.Abbrev is NULL then '' " &_
            "  else HB.Abbrev " &_
            "  end " &_
            "AS [HB Abbrev] " &_
            "FROM dbo.Users U LEFT OUTER JOIN " &_
            "     dbo.HB_Membership HBM ON U.UserID = HBM.UserID LEFT OUTER JOIN " &_
            "     dbo.HearingBodies HB ON HB.ID = HBM.HearingBodyID " &_
            "ORDER BY HB.NameHB, U.UserID"
' Response.Write"strSQL=" & strSQL & "<br>"       ' DEVELOPMENT & DEBUG
' Response.End          ' DEVELOPMENT & DEBUG

Dim NoOfRecords
Set rsUsers = dbConnect.Execute(strSQL, NoOfRecords)
' Response.Write "NoOfRecords=" & NoOfRecords & "<br>"        ' DEVELOPMENT & DEBUG
' Response.End          ' DEVELOPMENT & DEBUG

If Err.Number <> 0 Then Call SeriousError
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<head>
<title><%=strAppTitle%> - <%=strSubTitle%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link rel="icon" type="image/png" href="/rulehearing/favicon.ico" />
</head>

<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" style="max-width: 1000px">
<% If LCase(Request.QueryString("type")) = "excel" Then %>
  <P><strong><font size="+1"><%=strAppTitle%></font></strong></P>
<% Else %>
  <!--#include file="../include/topright.asp"-->
  <p style="font-family: Arial; font-size: small;">[<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?type=excel">Open with Excel</a>]</p>
<% End If %>

<table style="<%=strFormStyle%>" cellpadding="2">
  <tr>
  <td align="center" style="font-family: Arial; font-size: medium; font-weight: bold"><%=strSubTitle%></td>
  </tr>
    <td>
      <table border="0" style="font-family: Arial; font-size: 10pt">
      <tr>
        <td bgcolor="#12b1ee" align="center" <%=strTableHeadStyle%>>UserID</td>
        <td bgcolor="#12b1ee" align="center" <%=strTableHeadStyle%>>Name</td>
        <td bgcolor="#12b1ee" align="center" <%=strTableHeadStyle%>>E-mail</td>
        <td bgcolor="#12b1ee" align="center" <%=strTableHeadStyle%>>Organisation</td>
        <td bgcolor="#12b1ee" align="center" nowrap <%=strTableHeadStyle%>>Hearing Body</td>
        <td bgcolor="#12b1ee" align="center" nowrap <%=strTableHeadStyle%>>HB Abbrev</td>
      </tr>
<%
Dim arrBgColors(2), iCnt, bgcolor
arrBgColors(0) = "#ffe1e2"
arrBgColors(1) = "#FFFFFF"
iCnt = 0

While Not rsUsers.EOF
'     Response.Write rsUsers("UserID") & "<br>"   ' DEVELOPMENT & DEBUG
  iCnt = iCnt + 1
  bgcolor = arrBgColors(iCnt MOD 2)
  ' IdentifyMember(Trim(rsUsers("UserID")))
  ' If "" <> Trim(rsUsers("Organisation")) Then strMemberOrg = Trim(rsUsers("Organisation"))
  ' If "" <> Trim(rsUsers("eMailAddress")) Then strMemberEmailAddr = Trim(rsUsers("eMailAddress"))
  ' If strMemberEmailAddr = "=ID" Then strMemberEmailAddr = rsUsers("UserID")
%>
      <tr>
        <td valign="top" bgcolor="<%=bgcolor%>" <%=strTableRowStyle%>><%=Server.HTMLencode(Trim(rsUsers("UserID")))%></td>
        <td valign="top" bgcolor="<%=bgcolor%>" <%=strTableRowStyle%>><%=Server.HTMLencode(Trim(rsUsers("Name")))%></td>
        <td valign="top" bgcolor="<%=bgcolor%>" <%=strTableRowStyle%>><%=Server.HTMLencode(Trim(rsUsers("eMailAddress")))%></td>
        <td valign="top" bgcolor="<%=bgcolor%>" <%=strTableRowStyle%>><%=Server.HTMLencode(Trim(rsUsers("Organisation")))%></td>
        <td valign="top" bgcolor="<%=bgcolor%>" <%=strTableRowStyle%>><%=Server.HTMLencode(Trim(rsUsers("Hearing Body")))%></td>
        <td valign="top" bgcolor="<%=bgcolor%>" <%=strTableRowStyle%>><%=Server.HTMLencode(Trim(rsUsers("HB Abbrev")))%></td>
      </tr>
<%
  rsUsers.MoveNext
Wend
%>
    </table>
    </td>
  </tr>
</table>

<%
   '***********************
   '**   Close connection.   **
   '***********************
   rsUsers.Close
   SET rsUsers= Nothing
   dbConnect.Close
   SET dbConnect = Nothing
%>
</body>
</HTML>
