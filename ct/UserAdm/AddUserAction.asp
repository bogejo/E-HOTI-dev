<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" runat="server" src="../include/md5_PAJ.js"></SCRIPT>
<%
'## Query parameters, all FORM:
'##  selSelectedHearingBodies:  multiple values of HearingBody IDs - parses like selSelectedHearingBodies=8&selSelectedHearingBodies=10
'##  txtEmailAddr:              
'##  txtOrg:                    
'##  txtPwd:                    
'##  txtUserID:                 
'##  txtUserName:               
%>
<%
On Error Resume Next
' On Error Goto 0     ' DEVELOPMENT & DEBUG
' Global variables
Dim strAuthorisedReferer, strReferer, strUIDToAdd, strNameToAdd, strPwdToAdd, strEmailAddrToAdd, txtOrgToAdd, strSQL
Dim rsSQL, strHBmembership, strSQLHearingBodies, strAddedHearingBodies, rsHB, dbConnect, strAbbrev, i

'## Check that the call comes from the authorised AddUser page
strAuthorisedReferer = Request.ServerVariables("SCRIPT_NAME")
strAuthorisedReferer = Request.ServerVariables("HTTP_HOST") & Left(strAuthorisedReferer, InStrRev(strAuthorisedReferer, "/")) & "AddUser.asp"
strReferer = Request.ServerVariables("HTTP_REFERER")
If strReferer = "" Then Call SeriousError  ' Wasn't called from the proper screen

strReferer = Right(strReferer, Len(strReferer) - InStr(strReferer, "//")-1)
If StrComp(LCase(strReferer), LCase(strAuthorisedReferer)) <> 0 Then Call SeriousError  ' Wasn't called from the proper screen 

IdentifyUser
If Not bIsAdm Then Call SeriousError

'## Get the posted fields, and verify required fields.
Call CheckRequiredValue(Request.Form("txtUserID"),"Member's UserID")
strUIDToAdd = RTrim(LTrim(Request.Form("txtUserID")))
Call CheckRequiredValue(Request.Form("txtUserName"),"Member's name")
strNameToAdd = RTrim(LTrim(Request.Form("txtUserName")))
Call CheckRequiredValue(Request.Form("txtPwd"),"Password for member")
strPwdToAdd = RTrim(LTrim(Request.Form("txtPwd")))

strEmailAddrToAdd = RTrim(LTrim(Request.Form("txtEmailAddr")))
' If strEmailAddrToAdd <> "" Then strEmailAddrToAdd = ReplaceQuote(Request.Form("txtEmailAddr"))
Call CheckRequiredValue(Request.Form("selSelectedHearingBodies"),"Hearing body membership")
txtOrgToAdd = RTrim(LTrim(Request.Form("txtOrg")))
' If txtOrgToAdd <> "" Then txtOrgToAdd = ReplaceQuote(txtOrgToAdd)

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

'## Check whether the UserID already exists.
strSQL = "exec UsersSelProc " & dbText(strUIDToAdd)
' Response.Write "strSQL=" & strSQL & "<br>"     ' DEVELOPMENT & DEBUG
Set rsSQL = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then
  Set rsSQL = Nothing
  Call SeriousError
End If

If Not rsSQL.EOF Then UpdateUIDinsteadQ "UIDexists"  ' UID already present. Ask whether user wishes to update it.

'## All required fields are supplied, now store this in the database:
' Response.Write "strUIDToAdd=" & strUIDToAdd & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strNameToAdd=" & strNameToAdd & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strPwdToAdd=" & strPwdToAdd & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strEmailAddrToAdd=" & strEmailAddrToAdd & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strHBmembership=" & strHBmembership & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "txtOrgToAdd=" & txtOrgToAdd & ";<br>"   ' DEVELOPMENT & DEBUG

' Encrypt password with MD5
strSQL = "exec dbo.UsersInsProc " & _
      dbText(strNameToAdd) & ", " & _
      dbText(txtOrgToAdd) & ", " & _
      dbText(strUIDToAdd) & ", " & _
      dbText(strHBmembership) & ", " & _
      dbText(hex_md5(strPwdToAdd))

' Response.Write "strSQL=" & strSQL & "<br>"   '  DEVELOPMENT & DEBUG
' Response.Write "hex_md5(strPwdToAdd)=" & hex_md5(strPwdToAdd) & "<br>"   '  DEVELOPMENT & DEBUG
' Response.End         ' DEVELOPMENT & DEBUG
dbConnect.Execute(strSQL)
' Response.Write "Err.Number=" & Err.Number & "<br>"        ' DEVELOPMENT & DEBUG
' Response.Write "Err.Source=" & Err.Source & "<br>"        ' DEVELOPMENT & DEBUG
If Err.Number <> 0 Then Call SeriousError


' Retrive the name(s) and abbreviation(s) for the Hearing bodies where UserID is member  ' July 2005, using indexing table HB_Membership
strSQLHearingBodies = _
        "SELECT NameHB, Abbrev FROM HearingBodies WHERE ID IN " &_
           "(SELECT HearingBodyID FROM HB_Membership WHERE UserID = " & dbText(strUIDToAdd) & ")"
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

' Release upload object from memory

'***********************
'**  Close connection.  **
'***********************
Set rsSQL = Nothing
Set rsHB = Nothing
If IsObject(dbConnect) Then dbConnect.Close
SET dbConnect = Nothing

%>
<html>

<head>
<title><%=strAppTitle%> - Member Added</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<!-- link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css" -->
</head>
<body bgcolor="#FFFFFF"  style="max-width: 1000px">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">This Member has been added:</font></strong></td>
  </tr>
</table>
<table border="0" cellspacing="4">
  <tr>
    <td align="right" style="font-family: Arial; font-size: 10pt">UserID:</td>
    <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strUIDToAdd)%></td>
  </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Name:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strNameToAdd)%></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Password:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strPwdToAdd)%></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">E-mail:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strEmailAddrToAdd)%>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Hearing Bodies:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strAddedHearingBodies)%></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Organisation:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(txtOrgToAdd)%></td>
    </tr>
  </table>


<p>
<font face="Arial" size="-1">
[<a href="AddUser.asp" target="_self">Add another member</a>]&nbsp;&nbsp;&nbsp;&nbsp;
[<a href="../ruleDocs.asp" target="_self">Go to document list</a>]
</font>
</p>


</body>
</html>


<%
'************************************
Sub UpdateUIDinsteadQ (strReason)
' The UserID is already present in the databae. Ask whether user wishes to update the existing UserID instead.

Dim strReasonExplained, strQuery
strReasonExplained = "No explanation"

If strReason = "UIDexists" Then
  strReasonExplained = "UserID <b>" & Server.HTMLencode(rsSQL("UserID")) & "</b> is already registered in the database.<br>" & _
    "Do you wish to update the existing UserID <b>" & Server.HTMLencode(rsSQL("UserID")) &  "</b>?"
End If
%>
  <html>
  <head>
  <title><%=strAppTitle%> - UserID already present</title>
  <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
  <link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
  </head>
  <body bgcolor="#FFFFFF" style="max-width: 1000px">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">Already present - cannot add</font></strong></td>
  </tr>
</table>

  <p>
    <%=strReasonExplained%>
  </p>
<% If NOt rsSQL.EOF Then %>
  <p>
    <a href="UpdateUser.asp?UID=<%=Server.URLencode(rsSQL("UserID"))%>">Update <%=Server.HTMLencode(rsSQL("UserID"))%></a>
  </p>
<% End If %>
  <p>
    <a href="AddUser.asp">Back to Add Member</a>
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
