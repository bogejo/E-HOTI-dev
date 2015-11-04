<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
'## Query parameters (all Form):
'##   selSelectedHearingBodies
'##   txtEmailAddr
'##   txtOrg
'##   txtPwd
'##   txtUserID
'##   txtUserName
'## Adapted to indexing tables for HearingBodies / 2005-07-21
%>
<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" runat="server" src="../include/md5_PAJ.js"></SCRIPT>
<%
On Error Resume Next
' On Error Goto 0     ' DEVELOPMENT & DEBUG

' Global variables
Dim strAuthorisedReferer, strReferer, strUIDToUpdate, strNameToUpdate, strPwdToUpdate, strEmailAddrToUpdate, txtOrgToUpdate, strSQL
Dim rsSQL, strHBmembership, strSQLHearingBodies, strAddedHearingBodies, rsHB, strPWDunchangedMatch, strDefaultEmailStringMatch, rxRegExp
Dim strPwdToStore, dbConnect, i, strAbbrev

strPWDunchangedMatch = "\<no change\>"
strDefaultEmailStringMatch = "\[=UserID\]|=ID"


'## Check that the call comes from the authorised AddUser page
strAuthorisedReferer = Request.ServerVariables("SCRIPT_NAME")
strAuthorisedReferer = Request.ServerVariables("HTTP_HOST") & Left(strAuthorisedReferer, InStrRev(strAuthorisedReferer, "/")) & "UpdateUser.asp"
strReferer = Request.ServerVariables("HTTP_REFERER")
strReferer = Left(strReferer, InStrRev(strReferer, "?")-1)  ' Chop any URL query part
If strReferer = "" Then Call SeriousError  ' Wasn't called from the proper screen
strReferer = Right(strReferer, Len(strReferer) - InStr(strReferer, "//")-1)
If StrComp(LCase(strReferer), LCase(strAuthorisedReferer)) <> 0 Then Call SeriousError  ' Wasn't called from the proper screen 

IdentifyUser
If Not bIsAdm Then Call SeriousError

'## Get the posted fields, and verify required fields.
Call CheckRequiredValue(Request.Form("txtUserID"),"Member's UserID")
strUIDToUpdate = Request.Form("txtUserID")
Call CheckRequiredValue(Request.Form("txtUserName"),"Member's name")
strNameToUpdate = Request.Form("txtUserName")
strPwdToUpdate = Request.Form("txtPwd")
If strPwdToUpdate <> "" Then
  Set rxRegExp = New RegExp
  rxRegExp.Pattern = strPWDunchangedMatch
  rxRegExp.IgnoreCase = True
  If rxRegExp.Test(strPwdToUpdate) Then
    strPwdToStore  = ""  ' A default value was input - <no change>
    strPwdToUpdate = "<unchanged>"
  Else
    strPwdToStore = hex_md5(strPwdToUpdate)   ' Password shall be changed. Encrypt it with MD5
  End If
  Set rxRegExp = Nothing
End If

strEmailAddrToUpdate = Request.Form("txtEmailAddr")
If strEmailAddrToUpdate <> "" Then
  strEmailAddrToUpdate = Request.Form("txtEmailAddr")
  Set rxRegExp = New RegExp
  rxRegExp.Pattern = strDefaultEmailStringMatch
  rxRegExp.IgnoreCase = True
  If rxRegExp.Test(strEmailAddrToUpdate) Then
    strEmailAddrToUpdate= ""  ' A default value was input - [=UserID] or =ID
  ElseIf LCase(strEmailAddrToUpdate) = LCase(strUIDToUpdate) Then
    strEmailAddrToUpdate  = "=ID"
  End If
  Set rxRegExp = Nothing
End If

Call CheckRequiredValue(Request.Form("selSelectedHearingBodies"),"Hearing body membership")
txtOrgToUpdate = Request.Form("txtOrg")

'## Hearing bodies: build HearingBodiesMembership comma separated list  - "7,4" (July 2005 - indexing tables),
strHBmembership = ""
For i = 1 To Request.form("selSelectedHearingBodies").Count
  strHBmembership = strHBmembership & Request.form("selSelectedHearingBodies")(i) & ","
Next
If strHBmembership <> "" Then strHBmembership = Left(strHBmembership, Len(strHBmembership)-1)  ' chop the trailing , (comma)
If strHBmembership = "default" Then strHBmembership = ""   ' Membership is not mandatory / BGJ 2005-09-01


SET dbConnect = Server.CreateObject("ADODB.Connection")
If Err.Number <> 0 Then Call SeriousError
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError

'## Verify that the UserID already exists.
strSQL = "exec UsersSelProc " & dbText(strUIDToUpdate)
' Response.Write "strSQL=" & strSQL & "<br>"     ' DEVELOPMENT & DEBUG
Set rsSQL = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then
  Set rsSQL = Nothing
  Call SeriousError
End If

If rsSQL.EOF Then UserNotFound strUIDToUpdate

'## All required fields are supplied, now store this in the database:
' Response.Write "strUIDToUpdate=" & strUIDToUpdate & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strNameToUpdate=" & strNameToUpdate & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strPwdToUpdate=" & strPwdToUpdate & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strEmailAddrToUpdate=" & strEmailAddrToUpdate & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strHBmembership=" & strHBmembership & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "txtOrgToUpdate=" & txtOrgToUpdate & ";<br>"   ' DEVELOPMENT & DEBUG

strSQL = "exec dbo.UsersUpdProc " & _
      dbText(strUIDToUpdate) & ", "
      strSQL = strSQL & SQLAddValueOrNull(strHBmembership, True)
      strSQL = strSQL & SQLAddValueOrNull(strPwdToStore, True)
      strSQL = strSQL & SQLAddValueOrNull(strNameToUpdate, True)
      strSQL = strSQL & SQLAddValueOrNull(strEmailAddrToUpdate, True)
      strSQL = strSQL & SQLAddValueOrNull(txtOrgToUpdate, False)

' Response.Write "strSQL=" & strSQL & "<br>"   '  DEVELOPMENT & DEBUG
'Set rsSQL = Nothing   ' DEVELOPMENT & DEBUG
'dbConnect.Close   ' DEVELOPMENT & DEBUG
'Set dbConnect = Nothing   ' DEVELOPMENT & DEBUG
'Response.End   ' DEVELOPMENT & DEBUG

dbConnect.Execute(strSQL)
If Err.Number <> 0 Then
  Set rsSQL = Nothing
  Call SeriousError
End If

strSQL = "exec UsersSelProc " & dbText(strUIDToUpdate)
Set rsSQL = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then
  Set rsSQL = Nothing
  Call SeriousError
End If
If IsNull(rsSQL("Organisation")) Then
  txtOrgToUpdate = ""
Else
  txtOrgToUpdate = rsSQL("Organisation")
End If


' Retrive the name(s) and abbreviation(s) for the Hearing bodies where UserID is member  ' July 2005, using indexing table HB_Membership
strSQLHearingBodies = _
        "SELECT NameHB, Abbrev FROM HearingBodies WHERE ID IN " &_
           "(SELECT HearingBodyID FROM HB_Membership WHERE UserID = " & dbText(strUIDToUpdate) & ")"
' Response.Write "strSQLHearingBodies=" & strSQLHearingBodies & "<br>"        ' DEVELOPMENT & DEBUG
Set rsHB = dbConnect.Execute(strSQLHearingBodies)
If Err.Number <> 0 Then Call SeriousError

strAddedHearingBodies = ""
While Not rsHB.EOF
  strAddedHearingBodies = strAddedHearingBodies & Trim(rsHB("NameHB"))
  strAbbrev = Trim(rsHB("Abbrev"))
  If strAbbrev <> "" Then strAddedHearingBodies = strAddedHearingBodies & " (" & strAbbrev & ")"
  strAddedHearingBodies = strAddedHearingBodies  & ", "
  rsHB.MoveNext
Wend
If strAddedHearingBodies <> "" Then strAddedHearingBodies = Left(strAddedHearingBodies, Len(strAddedHearingBodies) - 2)  ' chop traling ", "

%>
<html>

<head>
<title><%=strAppTitle%> - Member Updated</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<!-- link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css" -->
</head>
<body bgcolor="#FFFFFF" style="max-width: 1000px">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">Member credentials has been updated:</font></strong></td>
  </tr>
</table>
<table border="0" cellspacing="4">
  <tr>
    <td align="right" style="font-family: Arial; font-size: 10pt">UserID:</td>
    <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strUIDToUpdate)%></td>
  </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Name:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(rsSQL("NameUser"))%></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Password:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strPwdToUpdate)%></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">E-mail:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(rsSQL("eMailAddress"))%>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Hearing Bodies:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strAddedHearingBodies)%></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Organisation:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(txtOrgToUpdate)%></td>
    </tr>
  </table>


<P>
<font face="Arial" size="-1">
[<a href="UpdateOrRemoveUser.asp" target="_self">Update or remove another member</a>]&nbsp;&nbsp;&nbsp;&nbsp;
[<a href="../ruleDocs.asp" target="_self">Go to document list</a>]
</font>
</P>


</body>
</html>

<%
' Release objects from memory

'***********************
'**  Close connection.  **
'***********************
Set rsSQL = Nothing
Set rsHB = Nothing
If IsObject(dbConnect) Then dbConnect.Close
SET dbConnect = Nothing
%>

<%
'************************************
Sub UserNotFound (strUID)
' The UserID was not present in the databae.

Dim strReasonExplained, strQuery
strReasonExplained = "UserID <b>" & Server.HTMLencode(strUID) & "</b> is not registered in the database.<br>" & _
    "Please go back and verify your input"
%>
  <html>
  <head>
  <title><%=strAppTitle%> - UserID not found</title>
  <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
  <link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
  </head>
  <body bgcolor="#FFFFFF" style="max-width: 1000px">
  <!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">Member not found - cannot update</font></strong></td>
  </tr>
</table>

  <p>
    <%=strReasonExplained%>
  </p>
  <p>
    <a href="UpdateUser.asp">Back to Update Member</a>
  </p>
  </body>
  </html>
<%
  Set rsSQL = Nothing
  dbConnect.Close
  Set dbConnect = Nothing
  Response.End
End Sub

%>
