<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
'## Adds a rule hearing document to the database and uploads
'## the physical file to the RP Documents Repository
'## Called by AddRuleDoc.asp
'## Query parameters - all FORM (passed via a clsUpload object instead of the standard Request object):
'##   RPNo:      the hearing document's "Proposal Number"
'##   Title:     the RP's Title
'##   File1:     the uploaded file
'##   Due year:  yyyy
'##   Due month: mm
'##   DueDay:    dd
'##   txtChosenAudiences: ;-terminated list of HearingBodyIDs - "3;13;"
'## Adapted to indexing tables instead of ;-terminated list for HearingBodies / 2005-07-19
'##
%>

<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<!--#INCLUDE FILE="clsUpload.asp"-->

<%
DIM oUpload, strFileName, dbConnect, strRPNo, strTitle, strInvalidChars, oUploadedFile
Dim fso, txtChosenAudiences, rsSQL, strDestPath, strSQL, strSourcePath, iCommentsRefused, txtNoCommentsMsg

On Error Resume Next
' On Error Goto 0     ' DEVELOPMENT & DEBUG

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
' Response.Write "oUpload(""RPNo"").Value=" & oUpload("RPNo").Value & ";<br>"  ' DEVELOPMENT & DEBUG
Call CheckRequiredValueLocal(oUpload("RPNo").Value,"Rule Proposal number")
' strRPNo = ReplaceQuote(oUpload("RPNo").Value)  ' Overkill!
strRPNo = oUpload("RPNo").Value
Call CheckRequiredValueLocal(oUpload("Title").Value,"Document Title")
strTitle = oUpload("Title").Value
'  Call CheckRequiredValueLocal(oUpload("AddedBy").Value,"Your name")  ' Omitted - the logged on user's name is used
txtChosenAudiences = oUpload("txtChosenAudiences").Value
txtChosenAudiences = strSCTL2CSL(txtChosenAudiences)   ' If it's a ;-terminated list, convert it to Comma Separated list 
If oUpload("chkRefuseComments").Value = "on" Then
  iCommentsRefused = 1
  txtNoCommentsMsg = oUpload("txtNoCommentsMsg").Value
Else
  iCommentsRefused = 0
End If

SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError

'## Check if a document with this document number already exists.
strSQL = "SELECT * FROM RuleProps WHERE RPNo = " & dbText(strRPNo)
Set rsSQL = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then
  Set oUploadedFile = Nothing
  Set oUpload = Nothing
  Call SeriousError
End If

If Not rsSQL.EOF Then UpdateInsteadQ "RPexists"  ' RP already present. Ask whether user wishes to update it.

' Check whether a file with same name exists on destination.
If fso.FileExists(strDestPath) Then UpdateInsteadQ "FileExists"  ' Ask whether user wishes to update an existing RP

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
' Response.Write "strRPNo=" & strRPNo & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strTitle=" & strTitle & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strFileName=" & strFileName & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "oUpload(""DueYear"").Value=" & oUpload("DueYear").Value & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "oUpload(""DueMonth"").Value=" & oUpload("DueMonth").Value & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "oUpload(""DueDay"").Value=" & oUpload("DueDay").Value & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "oUpload(""AddedBy"").Value=" & oUpload("AddedBy").Value & ";<br>"   ' DEVELOPMENT & DEBUG

strSQL = "exec dbo.RulePropsInsProc " & _
      dbText(strRPNo) & ", " & _
      dbText(strTitle) & ", " & _
      dbText(strFileName) & ", " & _
      dbText(strUserID) & ", " & _
      dbText(oUpload("DueYear").Value & "-" & oUpload("DueMonth").Value & "-" & oUpload("DueDay").Value) & ", " & _
      dbText(iCommentsRefused) & ", " & _
      dbText(txtNoCommentsMsg)

If txtChosenAudiences <> "" Then strSQL = strSQL & ", " & dbText(txtChosenAudiences)
'  Response.Write "strSQL=" & strSQL & "<br>"     ' DEVELOPMENT & DEBUG
'  Response.End     ' DEVELOPMENT & DEBUG
dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError


' Release upload object from memory

'***********************
'**  Close connection.  **
'***********************
dbConnect.Close
SET dbConnect = Nothing

%>
<html>

<head>
<title><%=strAppTitle%> - Rule Proposal Posted</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
</head>
<body  style="max-width: 1000px" bgcolor="#FFFFFF">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">This Rule Proposal has been added:</font></strong></td>
  </tr>
</table>
<table border="0" cellspacing="4">
  <tr>
    <td align="right" style="font-family: Arial; font-size: 10pt">Rule Proposal:</td>
    <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strRPNo)%></td>
  </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Document title:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strTitle)%></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Document file:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strFileName)%></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Due Date:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(oUpload("DueYear").Value & "-" & oUpload("DueMonth").Value & "-" & oUpload("DueDay").Value)%></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Disable comments:</td>
      <td style="font-family: Arial; font-size: 10pt"><%If iCommentsRefused = 1 Then%>Yes<%Else%>No<%End If%></td>
    </tr>
<%If iCommentsRefused = 1 Then%>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">"No comments" message:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(txtNoCommentsMsg)%></td>
    </tr>
<%End If%>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Added by:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strUserName)%></td>
    </tr>
  </table>


<P class="text">
[<a href="AddRuleDoc.asp" target="_self">Add another hearing document</a>]&nbsp;&nbsp;&nbsp;&nbsp;
[<a href="../ruleDocs.asp" target="_self">Go to document list</a>]
</P>


</body>
</html>

<%
Set oUploadedFile = Nothing
Set oUpload = Nothing

'************************************
Sub UpdateInsteadQ (strReason)
' The RP, or the uploaded file, is already present. Ask whether user wishes to update the existing RP instead.

Dim strReasonExplained, strQuery
strReasonExplained = "No explanation"

If strReason = "RPexists" Then
  strReasonExplained = "Rule Proposal <b>" & Server.HTMLencode(rsSQL("RPNo")) & "</b> is already registered in the database.<br>" & _
    "Do you wish to update the existing Rule Proposal <b>" & Server.HTMLencode(rsSQL("RPNo") & " - " & rsSQL("Title")) & "</b>?"    
End If
If strReason = "FileExists" Then
  strReasonExplained = "The file <b>" & strFileName & "</b> is already present among the hearing documents"
  strQuery = "SELECT * FROM " &_
             "  (SELECT RPId As RPId, RPNo AS RPNo, Title AS Title, FileName AS FileName " &_
             "  FROM dbo.RuleProps " &_
             "  UNION " &_
             "  SELECT A.RPId As RPId, RP.RPNo AS RPNo, '[Annex] ' + A.AnnexTitle AS Title, A.FileName AS FileName " &_
             "  FROM dbo.RPAnnex A Join RuleProps RP on A.RPId = RP.RPId) DERIVEDTBL " &_
             "WHERE (FileName = '" & strFileName & "')"
  Set rsSQL = dbConnect.Execute(strQuery)
  If Err.Number <> 0 Then
    Set oUploadedFile = Nothing
    Set oUpload = Nothing
    Call SeriousError
  End If
  If Not rsSQL.EOF Then
    strReasonExplained = strReasonExplained & ", associated with Rule Proposal:<br><b>" & _
      Server.HTMLencode(rsSQL("RPNo") & " - " & rsSQL("Title")) & "</b><br><br>" & _
      "Do you wish to update the existing Rule Proposal <b>" & Server.HTMLencode(rsSQL("RPNo")) & "</b>?"
  End If
End If
%>
  <html>
  <head>
  <title><%=strAppTitle%> - Rule Proposal already present</title>
  <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
  <link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
  </head>
  <body  style="max-width: 1000px" bgcolor="#FFFFFF">
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
    <a href="UpdateRuleDoc.asp?RPId=<%=rsSQL("RPId")%>">Update <%=Server.HTMLencode(rsSQL("RPNo"))%></a>
  </p>
<% End If %>
  <p>
    <a href="javascript:history.back()">Back to Add Document</a>
  </p>
  </body>
  </html>
<%
  Set oUploadedFile = Nothing
  Set oUpload = Nothing
  If IsObject(dbConnect) Then dbConnect.Close
  Set dbConnect = Nothing
  Response.End
End Sub

'## Function to check if a required value from the user input form has
'## been typed in. This could be done by JavaScript in the browser, but requires
'## more work from the programmers point of view, and often encounters compability
'## problems between browsers. The method used here requires more resources from
'## the server, but this is not expected to pose an real problems, since the site
'## is not expected to receive any heavy usage.
Function CheckRequiredValueLocal(value, fieldName)
'  Response.Write "value=" & value & "<br>"   ' DEVELOPMENT & DEBUG
  If value = "" Then %>
    <h3><font face="arial">Required field: <%= fieldName %></font></h3>
    Please click your back button, and complete the form.
<%  Set oUploadedFile = Nothing
    Set oUpload = Nothing
    Response.End
  End If
End Function

%>
