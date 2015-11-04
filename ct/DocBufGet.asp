<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%'Response.Redirect "TempDown.htm"%>
<!--#INCLUDE FILE="include/dbConnection.asp"-->
<!--#include file="include/Functions.inc"-->
<%
'## Return a file from the document repository via the user's document buffer
'## Can be used in master documents to link to annexes, Ref. CSR Bulk & Tankers hearings September 2007
'## Usage: https://rules.dnvgl.com/rulehearing/ct/DocBufGet.asp?fileName=CSR_Bulk_RCP1_TB(2nd_hearing).pdf
'## Query parameters:
'## filename: the file to be returned

DIM dbConnect, fileName, fso, fFile, strSrcPath, strDestFolder
Dim iPosLastSlash, strScriptFolder, strDocURL

On Error Resume Next
'  On Error Goto 0       ' DEVELOPMENT & DEBUG
IdentifyUser

If Request("fileName") = "" Then
  Response.End
Else
  fileName = Server.HTMLEncode(Replace(Replace(Replace(Request("fileName"), "..", ""), "/", ""), "\", ""))  ' Stop short cuts to inappropriate file system levels
End If
iPosLastSlash = InStrRev(Request.ServerVariables("SCRIPT_NAME"), "/")
strScriptFolder = Left(Request.ServerVariables("SCRIPT_NAME"), iPosLastSlash)
strDocURL = strProtocol & "://" & Request.ServerVariables("SERVER_NAME") & strScriptFolder & Replace("docbuf/" & strUserID & "/" & fileName, "+", "%20")
' Response.Write "strDocURL=" & strDocURL & "<br>"        ' DEVELOPMENT & DEBUG
' Response.End        ' DEVELOPMENT & DEBUG


set fso = Server.CreateObject("Scripting.FileSystemObject")
strSrcPath = strDocRepository & fileName
strDestFolder = strDocBuffer & strUserID & "\"
If Not fso.FolderExists(strDestFolder) Then
  fso.CreateFolder(strDestFolder)
End If
If Err.Number <> 0 Then Call SeriousError
If fso.FileExists(strSrcPath) Then
  ' Copy the file to buffer
   ' fso.CopyFile strSrcPath, strDestFolder, True   ' Error prone: appears to copy files with different endings: "RulesContents_xxx.pdf" is copied when spec is "RulesContents.pdf" / BGJ 2007-03-06
   Set fFile = fso.GetFile(strSrcPath)
   fFile.Copy(strDestFolder)
End If
If Err.Number <> 0 Then Call SeriousError
If IsObject(fso) Then Set fso = Nothing

Response.Redirect strDocURL
%>
