<%@ LANGUAGE="VBSCRIPT" %>
<% 
Option Explicit
DIM dbConnect, iRPId, strOutputCommentBy, strCommenterParticulars, rs
Dim strSQL, rsRP, CommentRS, strComment, color, fFirstPass, rxRegExp, ReviewerRS
Dim strHBrestrict, iID, iDNVemployee, bIsDNVemployee, DesignReviewDate
Dim strFileName
%>
<%
'## This script lists all comments attached to a particular 
'## rule hearing document. Output: comments, authors and 
'## date when comments were made.
'**   Adapted to indexing tables instead of ;-separated lists for Hearing Bodies / 2005-07-18
%>
<!--#include file="include/Functions.inc"-->
<!--#INCLUDE FILE="include/dbConnection.asp"-->


<%
On Error Resume Next
' On Error Goto 0       ' DEVELOPMENT & DEBUG
IdentifyUser
SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError
iRPId = CLng(Request("RPId"))

strSQL = "select RPNo from RuleProps where RPId = " & iRPId
Set rsRP = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError
If Not rsRP.EOF Then
  strFileName = "Comments-Responses_" & rsRP("RPNo") & ".doc"

  Response.Clear
  Response.AddHeader "Content-Disposition", "attachment; filename=" & strFileName
  Response.Charset = "ISO-8859-1"  ' Enforcing character set helps preventing cross-site scripting - But will it work when ContentType = "application/msword"?
  Response.ContentType = "application/msword"
End If
%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<head>
<title><%=strAppTitle%> - Comments-Response table for <%=iRPId%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<meta name="Microsoft Border" content="none">
<style type="text/css">
<!--
p {
  font-family: Calibri, Arial;
  font-size: 10pt;
  }

table {
  font-family: Calibri, Arial;
  font-size: 10pt;
  }
td {
  font-family: Calibri, Arial;
  font-size: 10pt;
  }
 -->
</style>

</head>

<body bgcolor="#FFFFFF">
<%  If rsRP.EOF Then %>
<p><font face="arial" size=2>If a document is selected, this window will list all comments related to that document.</font></p> 
<%    Response.End
  End If %>
<%
  '## Find the comments that this user may read
' July 2005, using index tables 'RestictedComments' and 'HB_Membership'
  '## Is logged on user a DNV GL employee, who may read all comments?
  bIsDNVemployee = bIsUserDNVEmployee(strUserID, dbConnect)

  '## Retrieve all comments to this document that the logged on user may read
'  strSQL = _ 
'     "SELECT C.CommentID, C.RPId, C.CommentByUserID, C.CommentByUser_Particulars, C.CommentDate, C.Comment, (SELECT TOP 1 RC.ID FROM RestrictedComments RC WHERE RC.CommentID = C.CommentID) AS RestrictedID " &_
'     "FROM Comments C " &_
'     "WHERE RPId = " & iRPId

  '## Retrieve all comments to this document that the logged on user may read
  strSQL = _ 
     "SELECT C.CommentID, C.RPId, C.CommentByUserID, C.CommentByUser_Particulars, C.CommentDate, C.CommentToPageNo, C.CommentToPlace, C.Comment, " &_
     "  (SELECT TOP 1 RC.ID FROM RestrictedComments RC WHERE RC.CommentID = C.CommentID) AS RestrictedID, RuleProps.RPNo, RuleProps.Title, RuleProps.DesignReviewDate " &_
     "FROM Comments C INNER JOIN RuleProps ON C.RPId = RuleProps.RPId " &_
     " WHERE C.RPId = " & iRPId


  If Not (bIsAdm OR bIsDNVemployee) Then   ' admins and DNV GL Employees may read all comments
    strSQL = strSQL &_ 
             " AND (" & _
                  "   (C.CommentByUserID = '" & strUserID & "' AND C.CommentByUserID NOT Like 'Member_%') OR " &_
                  "   C.CommentID NOT IN ( " &_
                  "     select CommentID from RestrictedComments " &_
                  "     ) OR " &_
                  "   C.CommentID IN ( " &_
                  "     SELECT DISTINCT RC2.CommentID " &_
                  "     FROM RestrictedComments RC2 JOIN HB_Membership HBM ON RC2.HearingBodyID = HBM.HearingBodyID " &_
                  "     WHERE HBM.UserID = '" & strUserID & "'" &_
                  "    ) " &_
                  ") "
  End If
  ' strSQL = strSQL & " ORDER BY C.CommentDate DESC"
  strSQL = strSQL & " ORDER BY C.CommentToPageNo, C.CommentDate"

 '  Response.Write strSQL & "<br>"   ' DEVELOPMENT & DEBUG
  Set CommentRS = dbConnect.Execute(strSQL)
  If Err.Number <> 0 Then Call SeriousError

  '## Go through a loop to list all comments, but only if there are any.
  If Not CommentRS.EOF Then
  DesignReviewDate = FormatDateTimeISO(CommentRS("DesignReviewDate"), False)
%>

<p style="text-align:center; font-weight:bold; font-size: 12pt;">
Comments - Proposed Responses<br>
Design Review Meeting <%=DesignReviewDate%></p>
<p style="font-style:italic; font-size: 10pt; margin-top: 0pt; margin-bottom: 0pt;">Participants: </p>
<p style="font-style:italic; font-size: 10pt; margin-top: 6pt; margin-bottom: 0pt">Comments were received through external hearing:</p>
<%
  strSQL = _ 
     "SELECT DISTINCT C.CommentByUserID, C.CommentByUser_Particulars " &_
     "FROM Comments C " &_
     "WHERE C.RPId = " & iRPId

  Set ReviewerRS = dbConnect.Execute(strSQL)
%>
<table border="0" style="margin-top: 0pt;">
<tbody>
  <tr>
    <td style="font-style:italic; font-weight:bold;">Company</td>
    <td style="font-style:italic; font-weight:bold;">Representatives</td>
    <td style="font-style:italic; font-weight:bold;">email address</td>
  </tr>

<%
  While Not ReviewerRS.EOF
  strCommenterParticulars = Trim(ReviewerRS("CommentByUser_Particulars"))
  If strCommenterParticulars <> "" then
     strMemberOrg = ""
     strOutputCommentBy = strCommenterParticulars

       If InStr(strCommenterParticulars, "@") > 0 Then
         strMemberEmailAddr = strCommenterParticulars
       Else
         strMemberEmailAddr = ""
       End If

  Else 
    IdentifyMember(ReviewerRS("CommentByUserID"))' Fills strMemberName, strMemberOrg, strMembersHearingBodies, strMemberEmailAddr
    strOutputCommentBy = strMemberName
    If bIsCollectiveUserID(ReviewerRS("CommentByUserID")) Then strMemberEmailAddr = ""
  End If %>
  <tr>
    <td><%=strMemberOrg%></td>
    <td><%=strOutputCommentBy%></td>
    <td><%=strMemberEmailAddr%></td>
  </tr>
<%
  ReviewerRS.MoveNext
  Wend
%>
  <tr>
    <td></td>
    <td></td>
  </tr>
</table>

<p style="font-size: 1em; font-weight:bold;">Proposal No.
<%= CommentRS("RPNo") %><br>
<%= CommentRS("Title") %></p>

<table border="0">
<tbody>
<tr>
<th valign="top" valign="top">Comm. No.</th><th valign="top">Page no.</th><th valign="top">Reference</th><th valign="top">Commentator</th><th valign="top">Need update</th><th valign="top">Comment</th><th valign="top">Proposed response</th><th valign="top">Status follow up action</th>
</tr>
<%
  color = "#FFFFFF"
  fFirstPass = True
  Do
    If Not fFirstPass Then
      CommentRS.MoveNext
    Else
      fFirstPass = False
    End If
    If CommentRS.EOF Then Exit Do
    If color = "#FFFFFF" Then
      color = "#ffe1e2"
    Else
      color = "#FFFFFF"
    End If %>
  <tr>

<%
  strCommenterParticulars = Trim(CommentRS("CommentByUser_Particulars"))
  If strCommenterParticulars <> "" then
     strOutputCommentBy = strCommenterParticulars
  Else 
    IdentifyMember(CommentRS("CommentByUserID"))' Fills strMemberName, strMemberOrg, strMembersHearingBodies
    strOutputCommentBy = strMemberName
    If strMemberOrg <> "" Then strOutputCommentBy = strOutputCommentBy & ", " & strMemberOrg
    ' strOutputCommentBy = strOutputCommentBy & ", " & strMembersHearingBodies  ' Add Committee (s)
  End If

'  Response.Write "strCommenterParticulars=" & strCommenterParticulars & ";<br>"   ' DEVELOPMENT & DEBUG
'  Response.Write FormatDateTimeISO(CommentRS("CommentDate"), False) & ": "
'  Response.Write Server.HTMLencode(strOutputCommentBy) & "</b><br>"
'  If Not IsNull(CommentRS("RestrictedID")) Then
'    Response.Write "<b><i><font size='-2' color='red'>RESTRICTED</font></i></b><br>"
'  End If
  
  '## Print the comment, but filter out character entity codes - &#nnnn;
  '## Also, replace carriage return with <br>
  strComment = CommentRS("Comment")
  Set rxRegExp = New RegExp
  With rxRegExp
    .Global = true ' replace ALL matching substrings
    .IgnoreCase = True
    .Pattern = "&#\d+;"
  End With
  strComment = rxRegExp.Replace(strComment, "")
  Set rxRegExp = Nothing
'  Response.Write Replace(Replace(Server.HTMLencode(strComment),chr(13),"<br>"), chr(11), "") %>
    <td valign="top"></td>
    <td valign="top"><%=CommentRS("CommentToPageNo")%></td>
    <td valign="top"><%=CommentRS("CommentToPlace")%></td>
    <td valign="top"><%=Server.HTMLencode(strOutputCommentBy)%></td>
    <td valign="top" align="center">Yes/No</td>
    <td valign="top"><%
      If Not IsNull(CommentRS("RestrictedID")) Then
        Response.Write "<span style='font:bold italic; color: red;'>RESTRICTED<br></span>"
      End If
  
'## Print the comment, but filter out character entity codes - &#nnnn;
'## Also, replace carriage return with <br>
    Response.Write Replace(Replace(Replace(Server.HTMLEncode(strComment),chr(13),"<br>"), chr(10), ""), chr(11), "") %></td>
    <td valign="top"></td>
    <td valign="top"></td>

  </tr>
<%  Loop %>
</tbody>
</table>

<p></p>

<%  '## no comments exists for this document
  Else %><p><font face="arial" size="2"><br>
<br>
No comments</font></p>

<%  End If

  dbConnect.Execute("exec dbo.MarkCommentsAsReadByUser " & iRPId & ", '" & strUserID & "'")  ' Mark comments as read by this reviewer
  If Err.Number <> 0 Then Call SeriousError
%>

<%
  '***********************
  '**  Close connection.  **
  '***********************
  Set CommentRS = Nothing
  Set ReviewerRS = Nothing
  dbConnect.Close
  SET dbConnect = Nothing
%>
  
</body>
</HTML>
