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
%>

<%
DIM dbConnect, rsUsers, strFormStyle, strSubTitle, strSQL, strAbbrev

Response.Clear
Response.Buffer = False  ' Allows output of this now "huge" list of users, more than 7.200 members

strSubTitle = "Update or remove member"
' strFormStyle = "border-style: solid; border-color:#D0D0D0;"
strFormStyle = ""
SET dbConnect = Server.CreateObject("ADODB.Connection")
If Err.Number <> 0 Then Call SeriousError
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError
'## Get list of users
strSQL = "SELECT Users.UserID, Users.NameUser, Users.Organisation, Users.eMailAddress, HearingBodies.ID As HBID, HearingBodies.NameHB, HearingBodies.Abbrev " &_
         "FROM Users " &_
         "LEFT OUTER JOIN HB_Membership ON Users.UserID = HB_Membership.UserID " &_
         "LEFT OUTER JOIN HearingBodies ON HB_Membership.HearingBodyID = HearingBodies.ID " &_
         "ORDER BY Users.UserID, HearingBodies.NameHB"

' Response.Write"strSQL=" & strSQL        ' DEVELOPMENT & DEBUG

Set rsUsers = Server.CreateObject("ADODB.Recordset")
rsUsers.Open strSQL, dbConnect, 3         ' Cursor adOpenStatic = 3; facilitates RecordCount and MovePrevious/Next
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
        <td bgcolor="#12b1ee" align="center"><font color="#FFFFFF" size="-1"><b>UserID &dArr;</b></font></td>
        <td bgcolor="#12b1ee" align="center"><font color="#FFFFFF" size="-1"><b>Name</b></font></td>
        <td bgcolor="#12b1ee" align="center"><font color="#FFFFFF" size="-1"><b>Organisation</b></font></td>
        <td bgcolor="#12b1ee" align="center"><font color="#FFFFFF" size="-1"><b>E-mail</b></font></td>
        <td bgcolor="#12b1ee" align="center" nowrap><font color="#FFFFFF" size="-1"><b>Hearing Bodies</b></font></td>
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

  strUserID = Trim(rsUsers("UserID"))
  strMemberName = Trim(rsUsers("NameUser"))
  strMemberOrg = Trim(rsUsers("Organisation"))
  If IsNull(strMemberOrg) Then strMemberOrg = ""
  strMemberEmailAddr = Trim(rsUsers("eMailAddress"))
  If IsNull(strMemberEmailAddr) OR strMemberEmailAddr = "=ID" Then strMemberEmailAddr = strUserID

  strMembersHearingBodies = ""
  Do While Not rsUsers.EOF
    If strUserID <> Trim(rsUsers("UserID")) Then Exit Do
    strMembersHearingBodies = strMembersHearingBodies & Trim(rsUsers("NameHB"))
    strAbbrev = Trim(rsUsers("Abbrev"))
    If strAbbrev <> "" Then strMembersHearingBodies = strMembersHearingBodies & " (" & strAbbrev & ")"
    strMembersHearingBodies = strMembersHearingBodies  & ", "
    rsUsers.MoveNext
  Loop
  rsUsers.MovePrevious  ' Back up; have moved into next userID
  If strMembersHearingBodies <> "" Then strMembersHearingBodies = Left(strMembersHearingBodies, Len(strMembersHearingBodies) - 2)  ' chop traling "' "

%>
      <tr>
        <td valign="top" bgcolor="<%=bgcolor%>" align="center"><font size="-1"><a href="RemoveUserAction.asp?UID=<%=Server.URLEncode(strUserID)%>" target="_top">remove</a></font></td>
        <td valign="top" bgcolor="<%=bgcolor%>" align="center"><font size="-1"><a href="UpdateUser.asp?UID=<%=Server.URLEncode(strUserID)%>">update</a></font></td>
        <td valign="top" bgcolor="<%=bgcolor%>"><font size="-1"><%=Server.HTMLencode(strUserID)%></font></td>
        <td valign="top" bgcolor="<%=bgcolor%>"><font size="-1"><%=Server.HTMLencode(strMemberName)%></font></td>
        <td valign="top" bgcolor="<%=bgcolor%>"><font size="-1"><%=Server.HTMLencode(strMemberOrg)%></font></td>
        <td valign="top" bgcolor="<%=bgcolor%>"><font size="-1"><%=Server.HTMLencode(strMemberEmailAddr)%></font></td>
        <td valign="top" bgcolor="<%=bgcolor%>"><font size="-1"><%=Server.HTMLencode(strMembersHearingBodies)%></font></td>
      </tr>
<%
  rsUsers.MoveNext
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
   rsUsers.Close
   SET rsUsers= Nothing
   dbConnect.Close
   SET dbConnect = Nothing
%>
</body>
</HTML>
