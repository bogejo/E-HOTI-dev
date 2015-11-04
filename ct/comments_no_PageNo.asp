<%@ LANGUAGE="VBSCRIPT" %>
<% 
Option Explicit
DIM dbConnect, iRPId, strOutputCommentBy, strCommenterParticulars, rs
Dim strSQL, rsRP, CommentRS, strComment, color, fFirstPass, rxRegExp
Dim strHBrestrict, iID, iDNVemployee, bIsDNVemployee
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
'  On Error Goto 0       ' DEVELOPMENT & DEBUG
IdentifyUser
SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError
iRPId = CLng(Request("RPId"))
' Response.End

Response.Charset = "ISO-8859-1"  ' Enforcing character set helps preventing cross-site scripting
%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<head>
<title><%=strAppTitle%> - Comments</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<meta name="Microsoft Border" content="none">
<style type="text/css">
<!--
@media print
  {
  .ScreenOnly {display: none}
  }
 -->
</style>

<script type="text/javascript" language="javascript">
<!--
function undockDOC() {
  var thisURL=window.location.href;
  window.open(parent.contents.location.href, 'HearingDoc');
  window.location.href=thisURL;
  }
//-->
</script>
</head>

<body bgcolor="#F2F2F2" style="margin-top: 2pt;">
<P>
<!-- <A HREF="http://www.dnvgl.com/" TARGET="_top"><IMG SRC="images/dnv_logo.gif" BORDER="0" ALIGN="bottom" ALT="DNV GL home"></A> -->
<A HREF="http://www.dnvgl.com/" TARGET="_top"><IMG SRC="images/DNVGL_logo_small.png" BORDER="0" ALIGN="bottom" ALT="DNV GL home"></A>
&nbsp;&nbsp;&nbsp;
<FONT face=arial><%=strAppTitle%></FONT>
<span class="ScreenOnly" style="position: relative; left: 250px; top: -50px;">
  <INPUT TYPE="Button" NAME="btnPrint" value="Print comments" onClick="javascript:print()">
</span>
</P>
<div class="ScreenOnly"><FONT face="arial" size="-2"><A href="ruleDocs.asp" target=_top>List of documents for hearing</A>
  <span style="position: absolute; right: 30px;">
    <a href="javascript:undockDOC()">Open document in own window</a>
  </span>
  </FONT>
</div>

<%
  strSQL = "select RPNo from RuleProps where RPId = " & iRPId
  Set rsRP = dbConnect.Execute(strSQL)
  If Err.Number <> 0 Then Call SeriousError

  If rsRP.EOF Then %>
<p><font face="arial" size=2>If a document is selected, this window will list all comments related to that document.</font></p> 
<%    Response.End
  End If %>
<p>
<strong><font face="arial" size=3>Comments to <font color="blue"><%= Server.HTMLEncode(rsRP("RPNo")) %></font>:</font></strong><br>
</p>
<%
  '## Find the comments that this user may read
' July 2005, using index tables 'RestictedComments' and 'HB_Membership'
  '## Is logged on user a DNV GL employee, who may read all comments?
  bIsDNVemployee = bIsUserDNVEmployee(strUserID, dbConnect)

  '## Retrieve all comments to this document that the logged on user may read
  strSQL = _ 
     "SELECT C.CommentID, C.RPId, C.CommentByUserID, C.CommentByUser_Particulars, C.CommentDate, C.CommentToPlace, C.Comment, (SELECT TOP 1 RC.ID FROM RestrictedComments RC WHERE RC.CommentID = C.CommentID) AS RestrictedID " &_
     "FROM Comments C " &_
     "WHERE RPId = " & iRPId

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
  strSQL = strSQL & " ORDER BY C.CommentDate DESC"

  ' Response.Write strSQL & "<br>"   ' DEVELOPMENT & DEBUG
  Set CommentRS = dbConnect.Execute(strSQL)
  If Err.Number <> 0 Then Call SeriousError

  '## Go through a loop to list all comments, but only if there are any.
  If Not CommentRS.EOF Then %>

<table border="0" width="100%" style="font-family: Arial; font-size: 10pt">
<%  color = "#FFFFFF"
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
    <td bgcolor="<%= color %>" height="56"><b>
<%
  strCommenterParticulars = Trim(CommentRS("CommentByUser_Particulars"))
  If strCommenterParticulars <> "" then
     strOutputCommentBy = strCommenterParticulars
  Else 
    IdentifyMember(CommentRS("CommentByUserID"))' Fills strMemberName, strMemberOrg, strMembersHearingBodies
    strOutputCommentBy = strMemberName
    If strMemberOrg <> "" Then strOutputCommentBy = strOutputCommentBy & ", " & strMemberOrg
    strOutputCommentBy = strOutputCommentBy & ", " & strMembersHearingBodies  ' Add Committee (s)
  End If
'  Response.Write "strCommenterParticulars=" & strCommenterParticulars & ";<br>"   ' DEVELOPMENT & DEBUG
  Response.Write FormatDateTimeISO(CommentRS("CommentDate"), False) & ": "
  Response.Write Server.HTMLencode(strOutputCommentBy) & "</b><br>"
  If Not IsNull(CommentRS("RestrictedID")) Then
    Response.Write "<b><i><font size='-2' color='red'>RESTRICTED</font></i></b><br>"
  End If
  
  '## Print the comment, but filter out character entity codes - &#nnnn;
  '## Also, replace carriage return with <br>

  strComment = ""
  If CommentRS("CommentToPlace") <> "" Then
    ' strComment = Replace(Replace(Replace(Server.HTMLEncode(CommentRS("CommentToPlace")),chr(13),"<br>"), chr(10), ""), chr(11), "") & ":" & chr(13) & chr(10)
    strComment = CommentRS("CommentToPlace") & ":" & chr(13) & chr(10)
  End If

  strComment = strComment & CommentRS("Comment")
  Set rxRegExp = New RegExp
  With rxRegExp
    .Global = true ' replace ALL matching substrings
    .IgnoreCase = True
    .Pattern = "&#\d+;"
  End With
  strComment = rxRegExp.Replace(strComment, "")
  Set rxRegExp = Nothing
  Response.Write Replace(Replace(Server.HTMLencode(strComment),chr(13),"<br>"), chr(11), "") %></td>
  </tr>
<%  Loop %>
</table>
<div class="ScreenOnly"><br><FONT face="arial" size="-2">To the <A href="ruleDocs.asp" target=_top>list of documents for hearing</A>
  <span style="position: absolute; right: 30px;">
    <a href="javascript:undockDOC()">Open document in own window</a>
  </span>
  </FONT>
</div>
<%  '## no comments exists for this document
  Else %><p><font face="arial" size="2"><br>
<br>
No comments</font></p>
<%  End If

  ' Response.Write "exec dbo.MarkCommentsAsReadByUser " & iRPId & ", '" & strUserID & "'"  ' DEVELOPMENT & DEBUG
  dbConnect.Execute("exec dbo.MarkCommentsAsReadByUser " & iRPId & ", '" & strUserID & "'")  ' Mark comments as read by this reviewer
  If Err.Number <> 0 Then Call SeriousError
%>

<%
  '***********************
  '**  Close connection.  **
  '***********************
  Set CommentRS = Nothing
  dbConnect.Close
  SET dbConnect = Nothing
%>
  
</body>
</HTML>
