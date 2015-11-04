<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit
'## This script adds an RP Annex to the database and file storage. It uploads
'## the physical file to the file storage area, and adds information in the database about the file 
'## and the person who performed the addition.
'## Based on AddRuleDocAction.asp
'## Query (form) parameters:
'##   RPId - The parent RP (internal ID)
'##   RPNo - The parent RP (public reference)
'##   AnnexTitle - the Annex's Title
'##   File - file system path to the Annex file, where the file is uploaded from
%>

<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<!--#INCLUDE FILE="clsUpload.asp"-->

<%
DIM oUpload, strFileName, dbConnect, iRPId, strRPNo, strTitle, strInvalidChars, oUploadedFile
Dim fso, rsRP, rsExistingAnnex, rsExistingFile, strSourcePath, strDestPath, strSQL

On Error Resume Next
' On Error Goto 0           ' DEVELOPMENT & DEBUG

IdentifyUser
If Not bIsAdm Then Call SeriousError

Set oUpload = New clsUpload
Set fso = CreateObject("Scripting.FileSystemObject")
If Err.Number <> 0 Then Call SeriousError


Set oUploadedFile = oUpload.Fields("File1")
' Grab the file name
strFileName = oUploadedFile.FileName
strSourcePath = oUploadedFile.FileDir
strDestPath = strDocRepository & strFileName
' Response.Write "strFileName=" & strFileName & ";<br>"   '  DEVELOPMENT & DEBUG
' Response.Write "strDestPath=" & strDestPath & ";<br>"   '  DEVELOPMENT & DEBUG

strInvalidChars =  strInvalidFnameChars(strFileName)
If strInvalidChars <> "" Then %>
  <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
  <html>

  <head>
  <title><%=strAppTitle%> - File name error</title>
  <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
  <link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
  </head>
  <body  style="max-width: 1000px" bgcolor="#FFFFFF">
  <p>
  The file name '<%=strFileName%>' contains the following characters that cannot be accepted:<br>
  <%=Server.HTMLencode(strInvalidChars)%><br>
  Please rename the file.
  </p>
  <input type="button" value="Back" onClick='history.back()'>
  </body>
  </html>
  <%
  Set oUploadedFile = Nothing
  Set oUpload = Nothing
  Response.End
End If
%>

<%  '## Check if the user typed in the required fields.
' Response.Write "oUpload(""RPId"").Value=" & oUpload("RPId").Value & ";<br>"  ' DEVELOPMENT & DEBUG
Call CheckRequiredValue(oUpload("RPId").Value,"Rule Proposal's system internal ID")
iRPId = CLng(oUpload("RPId").Value)   ' strRPNo = oUpload("RPNo").Value
Call CheckRequiredValue(oUpload("AnnexTitle").Value,"Annex Title")
strTitle = Trim(ReplaceQuote(oUpload("AnnexTitle").Value))
'  Call CheckRequiredValue(oUpload("AddedBy").Value,"Your name")  ' Omitted - the logged on user's name is used

SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError

'## Find RP's public reference
strSQL = "SELECT RPNo from RuleProps WHERE RPId = " & iRPId
Set rsRP = dbConnect.Execute(strSQL)
If (Not rsRP.EOF) Then strRPNo = rsRP("RPNo")
Set rsRP = Nothing

'## Check if an Annex with this Title already exists for the RP.
strSQL = "SELECT * FROM RPAnnex WHERE RPId = " & iRPId & " AND AnnexTitle = " & dbText(strTitle)
' Response.Write "strSQL=" & strSQL & "<br>"        ' DEVELOPMENT & DEBUG
Set rsExistingAnnex = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then
  Set oUploadedFile = Nothing
  Set oUpload = Nothing
  Call SeriousError
End If

If (Not rsExistingAnnex.EOF) Then UpdateInsteadQ "AnnexExists", rsExistingAnnex("ID")  ' Annex Title is already present for this RP. Ask whether user wishes to update it.

' Check whether a file with same name exists on destination.
If fso.FileExists(strDestPath) Then UpdateInsteadQ "FileExists", ""  ' Ask whether user wishes to update an existing RP

' Save the binary data to the file system
oUploadedFile.SaveAs strDestPath
If Not fso.FileExists(strDestPath) Then
  Response.Write "<p><b>Error!</b> File " & strFileName & " failed to upload.</p>"
  If IsObject(dbConnect) Then dbConnect.Close
  Set dbConnect = Nothing
  Set fso = Nothing
  Response.End
End If

'## All required fields are supplied, now store this in the database:
' Response.Write "iRPId=" & iRPId & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strTitle=" & strTitle & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strFileName=" & strFileName & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "oUpload(""AddedBy"").Value=" & oUpload("AddedBy").Value & ";<br>"   ' DEVELOPMENT & DEBUG

strSQL = "exec dbo.RPAnnexInsProc " & _
      iRPId & ", " & _
      dbText(strTitle) & ", " & _
      dbText(strFileName) & ", " & _
      dbText(strUserID)
          
' Response.Write "strSQL=" & strSQL & ";<br>"     ' DEVELOPMENT & DEBUG
' Response.End     ' DEVELOPMENT & DEBUG
dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError

' Release upload object from memory

'***********************
'**  Close connection.  **
'***********************
dbConnect.Close
SET dbConnect = Nothing
Set fso = Nothing

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%=strAppTitle%> - RP Annex Posted</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
</head>
<body  style="max-width: 1000px" bgcolor="#FFFFFF">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">This Annex has been added to <%=Server.HTMLencode(strRPNo)%>:</font></strong></td>
  </tr>
</table>
<table border="0" cellspacing="4">
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Annex title:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strTitle)%></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">File:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strFileName)%></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Added by:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strUserName)%></td>
    </tr>
  </table>

<hr max-width: 1000px">
<P class="text">
[<a href="UpdateRuleDoc.asp?RPId=<%= iRPId %>" target="_self">Continue updating "<%=Server.HTMLencode(strRPNo)%>"</a>]&nbsp;&nbsp;&nbsp;&nbsp;
[<a href="../ruleDocs.asp" target="_self">Go to document list</a>]
</P>


</body>
</html>

<%
Set oUploadedFile = Nothing
Set oUpload = Nothing

'************************************
Sub UpdateInsteadQ (strReason, iID)
' The RP Annex, or the uploaded file, is already present. Ask whether user wishes to update the existing Annex instead.

Dim strReasonExplained, strQuery, strReasonBrief

strReasonExplained = "No explanation"

If strReason = "AnnexExists" Then
  strReasonBrief = "Annex already present"
  strReasonExplained = "Rule Proposal " & Server.HTMLencode(strRPNo) & " already has this annex: <b>" & Server.HTMLEncode(strTitle) & "</b>.<br><br>" & _
    "Do you wish to update the existing annex?<br>" & _
    "<a href='UpdateRPannex.asp?type=annex&ID=" & iID & "'>Yes, update the existing annex <b>" & Server.HTMLencode(strTitle) & "</b></a><br>" &_
    "<a href='javascript:history.back(-1)'>No,&nbsp;&nbsp;&nbsp;go back and change annex data</a>"
End If
If strReason = "FileExists" Then
  strReasonBrief = "File already present"
  strReasonExplained = "The file <b>" & strFileName & "</b> is already present among the hearing documents"
  strQuery = "SELECT * FROM " &_
             "  (SELECT RPId As RPId, RPNo As RPNo, Title As Title, FileName As FileName " &_
             "  FROM dbo.RuleProps " &_
             "  UNION " &_
             "  SELECT A.RPId As RPId, RP.RPNo As RPNo, '[Annex] ' + A.AnnexTitle As Title, A.FileName As FileName " &_
             "  FROM dbo.RPAnnex A Join dbo.RuleProps RP On RP.RPId = A.RPId) DERIVEDTBL " &_
             "WHERE (FileName = '" & strFileName & "')"
  Set rsExistingFile = dbConnect.Execute(strQuery)
  If Err.Number <> 0 Then
    Set oUploadedFile = Nothing
    Set oUpload = Nothing
    Call SeriousError
  End If
  If Not rsExistingFile.EOF Then
    strReasonExplained = strReasonExplained & ", associated with:<br><b>" & _
      Server.HTMLencode(rsExistingFile("RPNo")) & " - " & Server.HTMLencode(rsExistingFile("Title")) & "</b><br><br>" & _
    "Do you wish to upload another file instead?<br>" & _
    "<a href='javascript:history.back(-1)'>Yes, go back - upload another file</a><br>" &_
    "<a href='UpdateRuleDoc.asp?RPId=" & Server.HTMLencode(rsExistingFile("RPId")) & "'>No,&nbsp;&nbsp;&nbsp;update the existing Rule Proposal <b>" & Server.HTMLencode(rsExistingFile("RPNo")) & "</b></a>"
  End If
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%=strAppTitle%> - <%=strReasonBrief%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
</head>
<body  style="max-width: 1000px" bgcolor="#FFFFFF">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font style="font-size: larger" face="Arial" color="#FFFFFF">Cannot add - <%=strReasonBrief%></font></strong></td>
  </tr>
</table>

<p>
  <%=strReasonExplained%>
</p>
</body>
</html>
<%
  Set oUploadedFile = Nothing
  Set oUpload = Nothing
  Set rsExistingFile = Nothing
  Set rsExistingAnnex = Nothing
  If IsObject(dbConnect) Then dbConnect.Close
  Set dbConnect = Nothing
  Response.End
End Sub

'******************
'## CheckRequiredValue verifies that a required value from the user input form has
'## been supplied.

Function CheckRequiredValue(value, fieldName)
'  Response.Write "value=" & value & "<br>"   ' DEVELOPMENT & DEBUG
  If value = "" Then %>
    <h3><font face="arial">Required field: <%= fieldName %></font></h3>
    Please click the Back button to go back and complete the form.
<%  Set oUploadedFile = Nothing
    Set oUpload = Nothing
    Response.End
  End If
End Function

%>
