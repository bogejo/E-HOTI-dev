<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
'## This script adds a Comment to a Rule Hearing document.
'## The comment is stored in a database, and is assosiated with
'## the document by it's document number. 
'## Query parameters:
'##   RPId               - the integer Id of the hearing doc being commented (QueryString)
'##   Comment            - the comment text - "blah blah blah blah"  (Form)
'##   txtChosenAudiences - a ;-separated list of Hearing Bodies for restricted access - "7;3;13;5;" (Form)
'##   txtMember          - the commenter; default value generated by 'strCommentBy()' - "Bo Johanson, Det Norske Veritas" (Form)
'##   txtPlace           - (optional) The referred location in the hearing document - "Page 5, paragraph B103"(Form)
'## Adapted to indexing tables instead of ;-separated list for HearingBodies / 2005-07-19

%>
<!--#INCLUDE FILE="include/dbConnection.asp"-->
<!--#INCLUDE FILE="include/Functions.inc"-->
<%
DIM dbConnect
On Error Resume Next
' On Error Goto 0       ' DEVELOPMENT & DEBUG
SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr ' DISABLE during pre-development
If Err.Number <> 0 Then   ' Attempt to hide from attempts to reverse engineer by SQL injection. Output generated up to here will be sent to browser
  Call SeriousError
End If

'## CheckRequiredValueLocal checks if a required value from the user input form has
'## been typed in. This could be done by JavaScript in the browser, but requires
'## more work from the programmers point of view, and often encounters compability
'## problems between browsers. The method used here requires more resources from
'## the server, but this is not expected to pose an real problems, since the site
'## is not expected to receive any heavy usage.

Function CheckRequiredValueLocal(value, fieldName)
   If value = "" Then %>
     <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
     <HTML>
     <head>
     <title>Missing required information in hearing comment</title>
     <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
     </head>
     <body bgcolor="#F2F2F2">
     <h3><font face="arial">Required field: <%= fieldName %></font></h3>
     <p>Please click your back button, and complete the form.</p>
     </body>
     </HTML>
<% Response.End
   End If
End Function 

Dim iRPId, strCommentByParticulars, CommentToPlace, Comment, strSQL, txtChosenAudiences, rxRegExp
IdentifyUser
iRPId = CLng(Request("RPId"))   ' RPNo = Request.QueryString("RPNo")
strCommentByParticulars = Trim(Request.Form("txtMember"))
If strCommentByParticulars = strCommentBy() Then
  strCommentByParticulars = "NULL"
Else
  SetCookie "UsrParticulars", strCommentByParticulars, strCookieFolder, 0
  strCommentByParticulars = dbText(strCommentByParticulars)
End If
txtChosenAudiences = Request.Form("txtChosenAudiences")
txtChosenAudiences = strSCTL2CSL(txtChosenAudiences)   ' If it's a ;-terminated list, convert it to Comma Separated list 
If txtChosenAudiences <> "" Then txtChosenAudiences = dbText(txtChosenAudiences)

CommentToPlace = ""
If Request.Form("txtPlace") <> "" Then CommentToPlace = dbText(Request.Form("txtPlace"))
Comment = dbText(Request.Form("Comment"))
'## Filter out character entities - &#nnnn;
Set rxRegExp = New RegExp
With rxRegExp
  .Global = true ' replace ALL matching substrings
  .IgnoreCase = True
  .Pattern = "&#\d+;"
End With
Comment = rxRegExp.Replace(Comment, "")
CommentToPlace = rxRegExp.Replace(CommentToPlace, "")

' Normalizing commenter's various ways to express "place": "Sect, Section", etc -> "Sec."  /BGJ 2012-02-16
Set rxRegExp = New RegExp
rxRegExp.Global = True
rxRegExp.IgnoreCase = True
rxRegExp.Pattern = "(sec)((tion)|t| )( |\.)*"
' Response.Write "CommentToPlace=" & CommentToPlace & "<br>"        ' DEVELOPMENT & DEBUG
CommentToPlace = rxRegExp.Replace(CommentToPlace, "$1.")
' Response.Write "CommentToPlace=" & CommentToPlace & "<br>"        ' DEVELOPMENT & DEBUG
' Response.End        ' DEVELOPMENT & DEBUG

Set rxRegExp = Nothing

'## Check if the user typed in the required fields.
Call CheckRequiredValueLocal(Request.Form("txtMember"), "Your name or user ID")
Call CheckRequiredValueLocal(Comment, "Your comment") %> <%   '## All required fields are supplied, now store this in the database:
strSQL = "exec dbo.CommentsInsProc " & iRPId & ", " & dbText(strGetUserID()) & ", " & strCommentByParticulars & ", " & CommentToPlace & ", " & Comment
' If txtChosenAudiences <> "" Then strSQL = strSQL & ", " & txtChosenAudiences   ' Restricted comments disabled 2010-04-20. See "addComment.asp", form field "txtChosenAudiences"
' Response.Write "strSQL : " & strSQL  ' DEVELOPMENT & DEBUG
' Response.End  ' DEVELOPMENT & DEBUG
dbConnect.Execute(strSQL) 
If Err.Number <> 0 Then   ' Attempt to hide from attempts to reverse engineer by SQL injection. Output generated up to here will be sent to browser
  Call SeriousError
End If
dbConnect.Close
Set dbConnect = Nothing %> 

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<head>
<title><%=strAppTitle%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript">
parent.commentFrame.location.href=parent.commentFrame.location.href; // Refresh the Comments window
location.href = "AddComment.asp?RPid=<%= iRPId %>";  // Restore the AddComment form
</SCRIPT>

</head>
<body bgcolor="#F2F2F2" onLoad="parent.commentFrame.location.href=parent.commentFrame.location.href">
<h3><font face="arial" size="-1">Your comment to the hearing document has been added.</font></h3>
<font face="arial" size="-1">
<p>To see the comment in the comment list, right-click your mouse in the comment list
frame, and select refresh.<br>
<br>
</font><a href="AddComment.asp?RPId=<%= iRPId %>" target="_self">Add another</a> comment.<br>
Go to <a href="ruleDocs.asp" target="_top">document list.</a> </p>
</body>
</HTML>