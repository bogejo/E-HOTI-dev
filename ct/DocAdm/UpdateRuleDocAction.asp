<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
'## Adds a rule hearing document to the database. It uploads
'## the physical file in the database, and adds information about the file 
'## and the person who performed the addition.
'## Query parameters - all FORM (passed via a clsUpload object instead of the standard Request object):
'##   RPId:      the hearing document's ID (integer)
'##   Title:     the RP's Title
'##   File1:     the uploaded file
'##   DueYear:   yyyy
'##   DueMonth:  mm
'##   DueDay:    dd
'##   txtChosenAudiences: ;-terminated list of HearingBodyIDs - "3;13;"
%>

<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<!--#INCLUDE FILE="clsUpload.asp"-->

<%
Dim oUpload, strOldFileName, strNewFileName, dbConnect, rsRP, iRPId, strRPNo, strTitle, strInvalidChars, oUploadedFile, strConfirmedFilename
Dim strDRDate
Dim fso, strSQL, strOldFilePath, txtChosenAudiences, strDestPath, strSourcePath, iCommentsRefused, txtNoCommentsMsg

On Error Resume Next
' On Error Goto 0     ' DEVELOPMENT & DEBUG
IdentifyUser
If Not bIsAdm Then Call SeriousError

strOldFileName = ""
strNewFileName = ""
Set oUpload = New clsUpload
Set fso = CreateObject("Scripting.FileSystemObject")
If Err.Number <> 0 Then Call SeriousError

'## Check if the required fields have been supplied
 ' Response.Write "oUpload(""RPId"").Value=" & oUpload("RPId").Value & ";<br>"  ' DEVELOPMENT & DEBUG
 ' Response.End  ' DEVELOPMENT & DEBUG
Call CheckRequiredValueLocal(oUpload("RPId").Value,"Document number")
iRPId = CLng(oUpload("RPId").Value)   ' strRPNo = ReplaceQuote(oUpload("RPNo").Value)
Call CheckRequiredValueLocal(oUpload("Title").Value,"Document Title")
strTitle = oUpload("Title").Value
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
strSQL = "SELECT RPId, RPNo, FileName FROM RuleProps WHERE RPId = " & iRPId
Set rsRP = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then
  Set rsRP = Nothing
  Set oUpload = Nothing
  Call SeriousError
End If

If rsRP.EOF Then %>
   <html>

   <head>
   <title><%=strAppTitle%> - Rule Proposal not found</title>
   <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
   <link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
   </head>
   <body style="max-width: 1000px" bgcolor="#FFFFFF">
   <p>
   Could not find a corresponding record in the database.<br>
   </p>
   <input type="button" value="Back" onClick='history.back()'>
   </body>
   </html>
   <%
   Set rsRP = Nothing
   Set oUpload = Nothing
   dbConnect.Close
   Seet dbConnect = Nothing
   Response.End
End If

' The RP exists in the database
strRPNo = rsRP("RPNo")
strOldFileName = rsRP("FileName")
strNewFileName = strOldFileName
strOldFilePath = strDocRepository & strOldFileName

' Find out if a new file is to replace the existing.
Set oUploadedFile = oUpload.Fields("File1")
If oUploadedFile.DataLength > 0 Then   ' A file was uploaded / User has specified a file to overwrite the existing
  ' Grab the file name
  strNewFileName = oUploadedFile.FileName
  strSourcePath = oUploadedFile.FileDir
  strDestPath = strDocRepository & strNewFileName
  ' Response.Write "strNewFileName=" & strNewFileName & ";<br>"   '  DEVELOPMENT & DEBUG
  ' Response.Write "oUploadedFile.FilePath=" & oUploadedFile.FilePath & ";<br>"   '  DEVELOPMENT & DEBUG
  ' Response.Write "oUploadedFile.DataLength=" & oUploadedFile.DataLength & ";<br>"   '  DEVELOPMENT & DEBUG
  ' Response.Write "strDestPath=" & strDestPath & ";<br>"   '  DEVELOPMENT & DEBUG
  ' Response.End                                            '  DEVELOPMENT & DEBUG

  strInvalidChars =  strInvalidFnameChars(strNewFileName)
  If strInvalidChars <> "" Then %>
    <html>

    <head>
    <title><%=strAppTitle%> - File name error</title>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
    <link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
    </head>
    <body style="max-width: 1000px" bgcolor="#FFFFFF">
    <p>
    The file name '<%=strNewFileName%>' contains the following characters that cannot be accepted:<br>
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

  ' Check whether a file with same name exists, associated with another RPId.
  If fso.FileExists(strDestPath) Then UpdateInsteadQ "FileExists"  ' Ask whether user wishes to update an existing RP

  If fso.FileExists(strOldFilePath) Then fso.DeleteFile strOldFilePath, true  ' The old file shall be discarded whether the same file name is reused or not
  ' Save the binary data to the file system
  oUploadedFile.SaveAs strDestPath
End If ' oUploadedFile.DataLength > 0 

'## All required fields are supplied, now store this in the database:
' Response.Write "iRPId=" & iRPId & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strTitle=" & strTitle & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strNewFileName=" & strNewFileName & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "oUpload(""DueYear"").Value=" & oUpload("DueYear").Value & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "oUpload(""DueMonth"").Value=" & oUpload("DueMonth").Value & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "oUpload(""DueDay"").Value=" & oUpload("DueDay").Value & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "oUpload(""DRYear"").Value=" & oUpload("DRYear").Value & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "oUpload(""DRMonth"").Value=" & oUpload("DRMonth").Value & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "oUpload(""DRDay"").Value=" & oUpload("DRDay").Value & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "oUpload(""AddedBy"").Value=" & oUpload("AddedBy").Value & ";<br>"   ' DEVELOPMENT & DEBUG

If oUpload("DRYear").Value = "" Or oUpload("DRMonth").Value = "" Or oUpload("DRDay").Value = "" Then
  strDRDate = ""
Else
  strDRDate = oUpload("DRYear").Value & "-" & oUpload("DRMonth").Value & "-" & oUpload("DRDay").Value
End If

strSQL = "EXEC RulePropsUpdate " & _
            iRPId & ", " & _
            dbText(strTitle) & ", " & _
            dbText(oUpload("DueYear").Value & "-" & oUpload("DueMonth").Value & "-" & oUpload("DueDay").Value) & ", " & _
            dbText(strDRDate) & ", " & _
            dbText(strNewFileName) & ", " & _
            dbText(iCommentsRefused) & ", " & _
            dbText(txtNoCommentsMsg)

If txtChosenAudiences <> "" Then strSQL = strSQL & ", " & dbText(txtChosenAudiences)
' Response.Write "strSQL=" & strSQL & ";<br>"     ' DEVELOPMENT & DEBUG
' Response.End     ' DEVELOPMENT & DEBUG
dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError

%>
<html>

<head>
<title><%=strAppTitle%> - Rule Proposal Updated</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
</head>
<body style="max-width: 1000px" bgcolor="#FFFFFF">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">This Rule Proposal has been updated:</font></strong></td>
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
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strNewFileName)%></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Due Date:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(oUpload("DueYear").Value & "-" & oUpload("DueMonth").Value & "-" & oUpload("DueDay").Value)%>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Design Review Date:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(oUpload("DRYear").Value & "-" & oUpload("DRMonth").Value & "-" & oUpload("DRDay").Value)%>
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
  </table>

<hr style="max-width: 1000px">
<P class="text">
[<a href="UpdateOrRemoveRuleDocs.asp">Update or remove another Hearing document</a>]&nbsp;&nbsp;&nbsp;&nbsp;
[<a href="../ruleDocs.asp" target="_self">Go to document list</a>]
</P>


</body>
</html>

<%
Set oUploadedFile = Nothing
Set oUpload = Nothing
Set fso = Nothing
Set rsRP = Nothing
'***********************
'**  Close connection.  **
'***********************
dbConnect.Close
SET dbConnect = Nothing

'************************************
Sub UpdateInsteadQ (strReason)
  ' The RP, or the uploaded file, is already present. Ask whether user wishes to update the existing RP instead.
  
  Dim strReasonExplained, strQuery, rsSQL, strMenu
  strReasonExplained = "No explanation"
  
  If strReason = "FileExists" Then
    ' Dont' touch the existing file if it's associated with an annex
    strQuery = "SELECT A.RPId As RPId, RP.RPNo As RPNo, A.AnnexTitle As Title, A.FileName FROM RPAnnex A JOIN RuleProps RP on A.RPId = RP.RPId WHERE (A.FileName = '" & strNewFileName & "')"
    Set rsSQL = dbConnect.Execute(strQuery)
    If Err.Number <> 0 Then
      Set oUploadedFile = Nothing
      Set oUpload = Nothing
      Call SeriousError
    End If
    If Not rsSQL.EOF Then ' It's an annex
        strReasonExplained = _
           "The file <b>" & strNewFileName & "</b> is already present among the hearing documents, " &_
           "being an annex to Rule Proposal:<br><b>" & _
           Server.HTMLencode(rsSQL("RPNo") & " - " & rsSQL("Title")) & "</b><br><br>" & _
           "Do you wish to update Rule Proposal - <b>" & Server.HTMLencode(rsSQL("RPNo")) & "</b>?" 
        strMenu = _
           "<ul style='list-style: disc;'>" &_
           "<li>" &_
           "Yes, update <a href='UpdateRuleDoc.asp?RPId=" & rsSQL("RPId") & "'>" & Server.HTMLencode(rsSQL("RPNo"))& "</a>" &_
           "</li>" & _
           "</ul>"
        If IsObject(rsSQL) Then
         rsSQL.Close
         Set rsSQL = Nothing
        End If
        ShowConflict strReasonExplained, strMenu
    End If
    
    ' The file is not an annex. Check whether it's associated with another RP
    strQuery = "SELECT RPId, RPNo, Title, FileName FROM RuleProps WHERE (FileName = '" & strNewFileName & "') "
    Set rsSQL = dbConnect.Execute(strQuery)
    If Err.Number <> 0 Then
      Set oUploadedFile = Nothing
      Set oUpload = Nothing
      Call SeriousError
    End If
    If rsSQL.EOF Then   ' The file is not registered in the RP database, so just go on and overwrite it
      rsSQL.Close
      Set rsSQL = Nothing
      Exit Sub
    Else
      If iRPId = rsSQL("RPId") Then
        rsSQL.Close
        Set rsSQL = Nothing
        Exit Sub  ' The file is to overwrite an existing with the same name, associated with the same RP. So just go ahead.
      Else
        strReasonExplained = "The file <b>" & strNewFileName & "</b> is already present among the hearing documents"
        strReasonExplained = strReasonExplained & ", associated with Rule Proposal:<br><b>" & _
          Server.HTMLencode(rsSQL("RPNo") & " - " & rsSQL("Title")) & "</b><br><br>" & _
          "Do you wish to update the other Rule Proposal - <b>" & Server.HTMLencode(rsSQL("RPNo")) & "</b>?"
        strMenu = _   
           "<ul style='list-style: disc;'>" &_
           "<li>" &_
           "Yes, update <a href='UpdateRuleDoc.asp?RPId=" & rsSQL("RPId") & "'>" & Server.HTMLencode(rsSQL("RPNo"))& "</a>" &_
           "</li>" & _
           "</ul>"
        If IsObject(rsSQL) Then
         rsSQL.Close
         Set rsSQL = Nothing
        End If
        ShowConflict strReasonExplained, strMenu
      End If
    End If
  End If
End Sub

Sub ShowConflict(strReasonExplained, strMenu)
  %>
  <html>
  <head>
  <title><%=strAppTitle%> - Document already present (<%=strLogonUser%>)</title>
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
  <%=strMenu%>
  <ul style="list-style: disc;">
  <li>
    Go back to update <a href="javascript:history.back()"><%=strRPNo%></a>
  </li>
  </ul>
  </body>
  </html>
  <%
  CleanUpAndQuit
End Sub

Sub CleanUpAndQuit
  Set oUploadedFile = Nothing
  Set oUpload = Nothing
  If IsObject(dbConnect) Then dbConnect.Close
  Set dbConnect = Nothing
  Response.End
End Sub


'## CheckRequiredValueLocal checks if a required value from the user input form has
'## been typed in. This could be done by JavaScript in the browser, but requires
'## more work from the programmers point of view, and often encounters compability
'## problems between browsers. The method used here requires more resources from
'## the server, but this is not expected to pose an real problems, since the site
'## is not expected to receive any heavy usage.
Function CheckRequiredValueLocal(value, fieldName)
'  Response.Write "value=" & value & "<br>"   ' DEVELOPMENT & DEBUG
  If value = "" Then %>
    <h3><font face="Arial,arial">Required field: <%= fieldName %></font></h3>
    Please click the "Back" button, and complete the form.
<%  Set oUploadedFile = Nothing
    Set oUpload = Nothing
    Response.End
  End If
End Function 
%>