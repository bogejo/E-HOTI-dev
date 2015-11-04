<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/popup.asp"-->
<html>

<head>
<title><%=strAppTitle%> - Remove Hearing Document</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
</head>

<body bgcolor="#FFFFFF">

<% 
'## Query parameters:
'##   type = "annex" | "RP" (default)
'##   ID   - For annex: the identifying (index) database field value
'##   strCheckboxGroup - For RP: the list of RPs
'##   ConfirmDelete - when "true", executes the delete operation; otherwise displays a "confirm delete" dialog

On Error Resume Next
' On Error Goto 0         ' DEVELOPMENT & DEBUG

IdentifyUser
If Not bIsAdm Then Call SeriousError

DIM dbConnect, strSQL, rsData, fso, strFile, strDocumentName, rxRegExp, strDocNameToken, arrDocNameToken, i, strType
Dim iID, arrStrRPid, strRPidSQL, strCheckboxGroup
Dim strConsequences, strResult, strRPNo, rsAppendixFile, x, iNoDocsToRemove
arrStrRPid = Array()
arrDocNameToken = Array()

strType = LCase(Request.QueryString("type"))
If strType = "" Then strType = "rp"

' DEVELOPMENT & DEBUG - BLOCK
'      Response.Write("<br><br>Request.Form.key(x) = Request.Form.item(x)<br>" )
'          for x = 1 to Request.Form.count() 
'              Response.Write(Request.Form.key(x) & " = ") 
'              Response.Write(Request.Form.item(x) & "<br>") 
'          next 
'      Response.Write("<br>")
' DEVELOPMENT & DEBUG - END OF BLOCK


 Select Case strType
  Case "rp"
    ' strRPNo = Request.Form(strCheckboxGroup)
    strCheckboxGroup = Request.Form("checkboxGroup")
    Call CheckRequiredValue(Request.Form(strCheckboxGroup)(1), "Proposal number")      '## Check if the user typed in the required fields.
    strRPidSQL = " RPid IN ("
    ' Response.Write "Request.Form(strCheckboxGroup)=" & Request.Form(strCheckboxGroup) & "<br>"        ' DEVELOPMENT & DEBUG
    arrStrRPid = Split(Request.Form(strCheckboxGroup), ", ")
    For i = LBound(arrStrRPid) To UBound(arrStrRPid)
      arrStrRPid(i) = CLng(arrStrRPid(i))
      ' Response.Write "arrStrRPid(" & i & ")=" & arrStrRPid(i) & "<br>"        ' DEVELOPMENT & DEBUG
      strRPidSQL = strRPidSQL & arrStrRPid(i) & ","
    Next
    strRPidSQL = Left(strRPidSQL, Len(strRPidSQL) - 1) & ")"  ' Chop trailing ",", add closing bracket
    strSQL = "SELECT RPId, RPNo, Title, FileName FROM RuleProps WHERE " & strRPidSQL & " ORDER BY RPNo"
    strDocNameToken = "%%RPNo%%"  ' %%<database field>%%
    strConsequences = "The operation will remove the Hearing document(s) with all comments and annexes"

  Case "annex"
    iID = CLng(Request.QueryString("ID"))
    strSQL = "SELECT RP.RPId As RPId, RP.RPNo As RPNo, A.AnnexTitle As Title, A.FileName FROM RPAnnex A join RuleProps RP on RP.RPId = A.RPId WHERE A.ID = " & iID
    strDocNameToken = "%%RPNo%% [Annex] %%Title%%"  ' %%<database field>%%
    strConsequences = "The operation will remove the RP annex"
  Case Else
    Call SeriousError
End Select

'## Substitute values for any databse field references in strDocNameToken
Set rxRegExp = New RegExp
With rxRegExp
  .Global = true ' replace ALL matching substrings
  .IgnoreCase = True
  .Pattern = "%%([^%]+)%%"
End With

' Response.Write "1. strDocumentName=" & strDocumentName & "<br>"   'DEVELOPMENT & DEBUG
' Response.Write "strSQL=" & strSQL & "<br>"        ' DEVELOPMENT & DEBUG

SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError

Set rsData = Server.CreateObject("ADODB.Recordset")
rsData.Open strSQL, dbConnect, 3         ' Cursor adOpenStatic = 3; facilitates RecordCount and MovePrevious/Next

If Err.Number <> 0 OR rsData.EOF Then CleanUpAndQuit
iNoDocsToRemove = rsData.RecordCount

While Not rsData.EOF
  arrDocNameToken = Split(strDocNameToken)
  For i = 0 To UBound(arrDocNameToken)
    ' Response.Write "1. arrDocNameToken(" & i & ")=" & arrDocNameToken(i) & "<br>"   ' DEVELOPMENT & DEBUG
    If rxRegExp.Test(arrDocNameToken(i)) Then 
      arrDocNameToken(i) = rxRegExp.Replace(arrDocNameToken(i), "rsData(""$1"")")
      ' Response.Write "2. arrDocNameToken(" & i & ")=" & arrDocNameToken(i) & "<br>"   ' DEVELOPMENT & DEBUG
      arrDocNameToken(i) = Eval(arrDocNameToken(i))
    End If
  Next
  strDocumentName = strDocumentName & Join(arrDocNameToken, " ") & ", "
  rsData.MoveNext
Wend

strDocumentName = Left(strDocumentName, Len(strDocumentName) - Len(", "))  ' Chop trailing "' "
' Response.Write "2. strDocumentName=" & strDocumentName & "<br>"   'DEVELOPMENT & DEBUG

Set rxRegExp = Nothing

If Request.QueryString("ConfirmDelete") <> "true" Then 
%>
<form action="<%=Request.ServerVariables("SCRIPT_NAME") & "?type=" & Server.URLEncode(strType) & "&ID=" & Server.URLencode(Request.QueryString("ID")) & "&ConfirmDelete=true"%>" method="post">
<%
For x = 1 To Request.Form.count() 
  Response.Write("<input type='Hidden' name='" & Request.Form.key(x) & "' value='" & Request.Form.item(x) &"'>")
Next 
%>  

<%
  Response.Write "<br><br><br>"
    call popup("Remove " & iNoDocsToRemove & " Hearing Document(s)?", _ 
               "Delete " & iNoDocsToRemove & " Hearing document(s)?", _
               strDocumentName, _
               Request.ServerVariables("SCRIPT_NAME") & "?type=" & Server.URLEncode(strType) & "&ID=" & iID & "&ConfirmDelete=true", _
               strConsequences, _
               "center","350",False) %>

<p class="text">Go to <a href="../AdminMenu.asp">Admin menu</a></p>
<%
  Set rsData = Nothing
  Set fso = Nothing
  dbConnect.Close
  SET dbConnect = Nothing
  Response.End
End If
%>

<%
'## Delete the file(s)
rsData.MoveFirst
While Not rsData.EOF
  strFile = rsData("FileName")
  iID = rsData("RPId")
  Set fso = Server.CreateObject("Scripting.FileSystemObject")
  If Err.Number <> 0 Then CleanUpAndQuit
   ' Response.Write "File=" & strDocRepository & strFile & "<br>"           ' DEVELOPMENT & DEBUG
   fso.DeleteFile strDocRepository & strFile
  If Err.Number <> 0 Then CleanUpAndQuit

  '## Remove all of the RP's appendix files
  If strType = "rp" Then
    ' Response.Write "SQL=" & "SELECT FileName FROM RPAnnex WHERE RPId = " & iID & "<br>"  ' DEVELOPMENT & DEBUG
    Set rsAppendixFile = dbConnect.Execute("SELECT FileName FROM RPAnnex WHERE RPId = " & iID)
    While Not rsAppendixFile.EOF
      ' Response.Write "File=" & strDocRepository & rsAppendixFile("FileName") & "<br>"   ' DEVELOPMENT & DEBUG
      fso.DeleteFile strDocRepository & rsAppendixFile("FileName")
      rsAppendixFile.MoveNext
    Wend
  End If


  '## Delete document: either RP with comments and annexes, or annex
   Select Case strType
    Case "rp"
      dbConnect.Execute("RulePropsAndCommentsDelProc " & iID)
      strResult = """" & strDocumentName & """ with comments and annexes were removed from the Hearing database"
      strRPNo = rsData("RPNo")
    Case "annex"
      dbConnect.Execute("RPAnnexDelProc " & CLng(Request.QueryString("ID")))
      strResult = """" & strDocumentName & """ was removed from the Hearing database"
      strRPNo = rsData("RPNo")
    Case Else
  End Select

  If Err.Number <> 0 Then CleanUpAndQuit
  rsData.MoveNext
Wend

%>

<p><font face="arial" size="+1"><b><%=strAppTitle%></b></font></p>
<p><font face="arial"><b><%=iNoDocsToRemove%> Hearing document(s) removed</b></font></p>

<p><font face="arial"><%= Server.HTMLEncode(strResult) %></font></p>

<hr style="max-width: 1000px;">
<p class="text" style="font-size: 90%">
<% If strType = "annex" Then %>
[<a href="UpdateRuleDoc.asp?RPId=<%=iID%>">Continue updating "<%= Server.HTMLEncode(strRPNo) %>"</a>]&nbsp;&nbsp;&nbsp;&nbsp;
<% End If %>
[<a href="UpdateOrRemoveRuleDocs.asp">Update or remove another Hearing Document</a>]&nbsp;&nbsp;&nbsp;&nbsp;
[<a href="../AdminMenu.asp">To Admin menu</a>]</p>
<%
  '***********************
  '**  Close connection.  **
  '***********************
  Set rsData = Nothing
  Set fso = Nothing
  dbConnect.Close
  SET dbConnect = Nothing
%>
</body>
</html>

<%
Sub CleanUpAndQuit ()
  Set rsData = Nothing
  Set fso = Nothing
  dbConnect.Close
  SET dbConnect = Nothing
  Call SeriousError
End Sub
%>