<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit
'## Based on UpdateRuleDocAction.asp
'## Query parameters via the clsUpload classes
'## Query parameters:
'##   ID   - the identifying (index) database field value
'##   Title - the Annex title
'##   File1 - the optional file to replace existing
%>

<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<!--#INCLUDE FILE="clsUpload.asp"-->

<%
DIM oUpload, strUploadedFileName, strPreviousFileName, dbConnect, rsAnnex, iID, iRPId, strRPNo, strTitle, strInvalidChars, oUploadedFile, strConfirmedFilename
Dim fso, strSQL, strOldFilePath, txtChosenAudiences, rsOverlappingAnnex, strAnnexFileName
Dim strDestPath, strSourcePath

 On Error Resume Next
'   On Error Goto 0     ' DEVELOPMENT & DEBUG
IdentifyUser
If Not bIsAdm Then Call SeriousError

strUploadedFileName = ""
Set oUpload = New clsUpload
Set fso = CreateObject("Scripting.FileSystemObject")
If Err.Number <> 0 Then Call SeriousError

'## Check if the required fields have been supplied
Call CheckRequiredValue(oUpload("ID").Value,"Annex ID no")
iID = CLng(Replace(oUpload("ID").Value, "'", ""))
Call CheckRequiredValue(oUpload("Title").Value,"Annex Title")
strTitle = Trim(oUpload("Title").Value)

SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError

'## Get data for the Annex
strSQL = "SELECT A.RPId, A.AnnexTitle, A.FileName, RP.RPNo FROM RPAnnex A Join RuleProps RP on A.RPId = RP.RPId WHERE A.ID = " & iID
' Response.Write "strSQL=" & strSQL & "<br>"     '   DEVELOPMENT & DEBUG
Set rsAnnex = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then
  Set rsAnnex = Nothing
  Set oUpload = Nothing
  Call SeriousError
End If

If rsAnnex.EOF Then %>
   <html>

   <head>
   <title><%=strAppTitle%> - RP Annex not found</title>
   <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
   <link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
   </head>
   <body style="max-width: 1000px" bgcolor="#FFFFFF">
   <p>
   Could not find '<%=iID%>' in the database.<br>
   </p>
   <input type="button" value="Back" onClick='history.back()'>
   </body>
   </html>
   <%
   Set rsAnnex = Nothing
   Set oUpload = Nothing
   dbConnect.Close
   Seet dbConnect = Nothing
   Response.End
End If

' The Annex exists in the database
strRPNo = rsAnnex("RPNo")
iRPId = CLng(rsAnnex("RPId"))
strPreviousFileName = rsAnnex("FileName")
strAnnexFileName = strPreviousFileName
strOldFilePath = strDocRepository & strPreviousFileName

'## Check if another Annex with this Title already exists for the RP.
strSQL = "SELECT A.RPId, A.AnnexTitle, RP.RPNo  FROM RPAnnex A Join RuleProps RP on A.RPId = RP.RPId WHERE A.RPId = " & iRPId & " AND A.AnnexTitle = " & dbText(strTitle) & " AND A.ID <> " & iID
'Response.Write "strSQL=" & strSQL & "<br>"        ' DEVELOPMENT & DEBUG
Set rsOverlappingAnnex = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then
  Set oUploadedFile = Nothing
  Set oUpload = Nothing
  Call SeriousError
End If
If (Not rsOverlappingAnnex.EOF) Then OverlapError "Annex title """ & strTitle & """ already exists", dbText(rsOverlappingAnnex("RPNo"))  ' Annex Title is already present for this RP.

' Find out if a new file is to replace the existing.
Set oUploadedFile = oUpload.Fields("File1")
If oUploadedFile.DataLength > 0 Then   ' A file was uploaded / User has specified a file to overwrite the existing
  ' Grab the file name
  strUploadedFileName = oUploadedFile.FileName
  strSourcePath = oUploadedFile.FileDir
  strDestPath = strDocRepository & strUploadedFileName
  ' Response.Write "strUploadedFileName=" & strUploadedFileName & ";<br>"   '  DEVELOPMENT & DEBUG
  ' Response.Write "oUploadedFile.FilePath=" & oUploadedFile.FilePath & ";<br>"   '  DEVELOPMENT & DEBUG
  ' Response.Write "oUploadedFile.DataLength=" & oUploadedFile.DataLength & ";<br>"   '  DEVELOPMENT & DEBUG
  ' Response.Write "strDestPath=" & strDestPath & ";<br>"   '  DEVELOPMENT & DEBUG
  ' Response.End                                            '  DEVELOPMENT & DEBUG

  strInvalidChars =  strInvalidFnameChars(strUploadedFileName)
  If strInvalidChars <> "" Then %>
    <html>

    <head>
    <title><%=strAppTitle%> - File name error</title>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
    <link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
    </head>
    <body style="max-width: 1000px" bgcolor="#FFFFFF">
    <p>
    The file name '<%=strUploadedFileName%>' contains the following characters that cannot be accepted:<br>
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


'*** If the uploaded file is already present, prevent overwriting if the file isn't associated with this Annex ID
  If fso.FileExists(strDestPath) AND strUploadedFileName <> strPreviousFileName Then 
    OverlapError "File is associated with another RP or Annex", strUploadedFileName
  End If

  If fso.FileExists(strOldFilePath) Then fso.DeleteFile strOldFilePath, true  ' The old file shall be discarded whether the same file name is reused or not
  ' Save the binary data to the file system
  oUploadedFile.SaveAs strDestPath
  strAnnexFileName = strUploadedFileName
End If ' oUploadedFile.DataLength > 0 

'## All required fields are supplied, now store this in the database:
' Response.Write "iID=" & iID & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strTitle=" & strTitle & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "strUploadedFileName=" & strUploadedFileName & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "oUpload(""DueYear"").Value=" & oUpload("DueYear").Value & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "oUpload(""DueMonth"").Value=" & oUpload("DueMonth").Value & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "oUpload(""DueDay"").Value=" & oUpload("DueDay").Value & ";<br>"   ' DEVELOPMENT & DEBUG
' Response.Write "oUpload(""AddedBy"").Value=" & oUpload("AddedBy").Value & ";<br>"   ' DEVELOPMENT & DEBUG

strSQL = "EXEC RPAnnexUpdate " & _
            iID & ", " & _
            dbText(strTitle) & ", " & _
            dbText(strAnnexFileName) & ", " & _
            dbText(strUserID)

' Response.Write "strSQL=" & strSQL & ";<br>"     ' DEVELOPMENT & DEBUG
' Response.End     ' DEVELOPMENT & DEBUG
dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>

<head>
<title><%=strAppTitle%> - RP Annex Updated</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
</head>
<body style="max-width: 1000px" bgcolor="#FFFFFF">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">This Annex has been updated:</font></strong></td>
  </tr>
</table>
<table border="0" cellspacing="4">
  <tr>
    <td align="right" style="font-family: Arial; font-size: 10pt">Rule Proposal:</td>
    <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strRPNo)%></td>
  </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Annex title:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strTitle)%></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Annex file:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=Server.HTMLencode(strAnnexFileName)%></td>
    </tr>
  </table>

<p class="text">
[<a href="UpdateRuleDoc.asp?RPId=<%= iRPId %>">Continue updating <%=Server.HTMLencode(strRPNo)%></a>]&nbsp;&nbsp;&nbsp;&nbsp;
[<a href="../ruleDocs.asp" target="_self">Go to document list.</a>]
</p>


</body>
</html>

<%
Call CleanUp


'**************** FUNCTIONS AND SUBROUTINES *****************

'## CheckRequiredValue checks if a required value from the user input form has
'## been typed in. This could be done by JavaScript in the browser, but requires
'## more work from the programmers point of view, and often encounters compability
'## problems between browsers. The method used here requires more resources from
'## the server, but this is not expected to pose an real problems, since the site
'## is not expected to receive any heavy usage.
Function CheckRequiredValue(value, fieldName)
'  Response.Write "value=" & value & "<br>"   ' DEVELOPMENT & DEBUG
  If value = "" Then %>
    <h3><font face="Arial,arial">Required field: <%= fieldName %></font></h3>
    Please click the "Back" button, and complete the form.
<%  Set oUploadedFile = Nothing
    Set oUpload = Nothing
    Response.End
  End If
End Function 

Sub OverlapError(strWhat, iID)
  Response.Write "<p>Overlapping: " & strWhat & " - " & iID & "</p>"        ' Primitive Stub
  Response.Write "<p><a href='javascript:history.back()'>Go back</a></p>"   ' Primitive Stub
  Call CleanUp
  Response.End
End Sub

Sub CleanUp
  Set oUploadedFile = Nothing
  Set oUpload = Nothing
  Set fso = Nothing
  Set rsAnnex = Nothing
  '***********************
  '**  Close connection.  **
  '***********************
  dbConnect.Close
  SET dbConnect = Nothing
End Sub
%>