<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
'## This script lists all the rule hearing documents available in the database. 
'## An index is created (table), containing RPId with a link for viewing and 
'## commenting on the document.
'## Query parameters:
'##   set: "archive" - view archived documents
'##        otherwise - view documents on active hearing
'## Adapted to using indexing tables instead of ;-separated lists for HearingBodies restrictions / 2005-07-18
%>
<%
 On Error Resume Next
' On Error Goto 0     ' DEVELOPMENT & DEBUG	
%>

<%'Response.Redirect "TempDown.htm"%>
<!--#include file="include/Functions.inc"-->
<!--#INCLUDE FILE="include/dbConnection.asp"-->
<%
Dim strParamOrderRPby, strThisPageURL, strOrderRPby
IdentifyUser
strParamOrderRPby = LCase(nStr(Request.QueryString("sortBy")))
' strOrderRPby
Select Case strParamOrderRPby
  Case "designreviewdate"
    strOrderRPby = "RuleProps.DesignReviewDate, RuleProps.RPNo"
  Case "duedate"
    strOrderRPby = "RuleProps.DueDate, RuleProps.RPNo"
  Case "title"
    strOrderRPby = "RuleProps.Title, RuleProps.RPNo"
  Case Else
    strOrderRPby = "RuleProps.RPNo"
End Select
%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<head>
<title><%=strAppTitle%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<style type="text/css">
<!--
table.RPlist td {padding-right: 10px; padding-left: 10px}  
a.noformat {
  text-decoration:none;
  color: #FFFFFF;
  }
-->
</style>

</head>

<body style="max-width: 1000px" bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
<!--#include file="include/topright.asp"-->
<%
DIM dbConnect, Sql, iRPId, strRPNo, fFirstPass, color, strInHB, bIsMemberOfTargetHB, bIsDNVemployee
SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError
bIsDNVemployee = bIsUserDNVEmployee(strUserID, dbConnect)
%>

<% If (LCase(Request.QueryString("set")) <> "archive") Or Not (bIsAdm Or bIsDNVemployee) Then  ' Show the active hearing documents %>
<%
   strThisPageURL = Request.ServerVariables("SCRIPT_NAME")&"?set=active"
%>
<div style="margin-left: 12px">
<% If bIsAdm Or bIsDNVemployee Then ' Only DNV GL employees to see the Archive; Nov. 2010 %>
<font face="arial" size="-1">[<a style="color: blue" href="<%=Request.ServerVariables("SCRIPT_NAME")&"?set=archive"%>">Go to archive</a>] (visible to DNV GL employees only)</font>
<% End If %>
</div>

<table cellspacing="10" width="100%">
  <tr>
    <td><table class="RPlist" border="0" width="100%" style="font-family: Arial; font-size: 10pt">
      <tr>
        <td bgcolor="#12b1ee" align="center"><font color="#FFFFFF" size="-1"><a href="<%=strThisPageURL%>&sortBy=RPNo" class="noformat"><b>Prop. No.<%=strFlagSorted("RPNo")%></b></a></font></td>
        <td bgcolor="#12b1ee" align="center"><font color="#FFFFFF" size="-1"><b>Status</b></font></td>
        <td bgcolor="#12b1ee"><font color="#FFFFFF" size="-1"><a href="<%=strThisPageURL%>&sortBy=Title" class="noformat"><b>Rule Proposal Title<%=strFlagSorted("Title")%></b></a></font></td>
        <td bgcolor="#12b1ee" align="center"><font color="#FFFFFF" size="-1"><a href="<%=strThisPageURL%>&sortBy=DueDate" class="noformat"><b>Due Date<%=strFlagSorted("DueDate")%></b></a></font></td>
        <td bgcolor="#12b1ee" nowrap align="center"><font color="#FFFFFF" size="-1"><b>Comments (kB)</b></font>
        <%If True or bIsAdm Then   ' Let all users see how much is new %>  
          <td bgcolor="#12b1ee" nowrap align="center"><font color="#FFFFFF" size="-1"><b>New (kB)</b></font>
        <%End If%>
      </tr>

<%   '## Get a recordset containing all the rule hearing documents.
   Dim strHBrestrict, iID, iDaysToExpiry, rsRuleProp, sqlWhere

'## July 2005 - uses indexing tables instead of ;-separated lists

   If bIsAdm or bIsDNVemployee Then
     sqlWhere = ""   ' admins and DNV GL Employees may view all hearing documents
   Else
     sqlWhere = _
         " AND  ((RPId NOT IN " &_
         "         (SELECT RPId " &_
         "          FROM RestrictedRuleProps)) OR " &_
         "       (RPId IN " &_
         "         (SELECT RPId " &_
         "          FROM RestrictedRuleProps " &_
         "          WHERE HearingBodyID IN " &_
         "            (SELECT HBM.HearingBodyID " &_
         "             FROM  HB_Membership HBM " &_
         "             WHERE UserID = '" & strUserID & "')))) "
   End If
   Sql = _
         "SELECT RPId, RPNo, Title, DueDate, AddedDate, FileName " &_
         "FROM   RuleProps " &_
         "WHERE DateArchived Is Null " & sqlWhere &_
         " ORDER BY RPNo"

   ' Response.Write "SQL=" & Sql & "<br>"                   ' DEVELOPMENT & DEBUG
   Set rsRuleProp = dbConnect.Execute(Sql)
   If Err.Number <> 0 Then Call SeriousError

   If Not rsRuleProp.EOF Then
     fFirstPass = True
     '## Now loop through all hearing documents that the user may view
     Do
       If Not fFirstPass Then
          rsRuleProp.MoveNext
       Else
          color = ""
          fFirstPass = False
       End If
       
       If rsRuleProp.EOF Then Exit Do
       
       If color = "#FFFFFF" Then
          color = "#ffe1e2"
       Else
          color = "#FFFFFF"
       End If
     
       DIM AddedDate, DueDate, DueDateList, rsNewComments, SQLGetNewNoComments, strNewFlag, strDaysToExpire, strPlural
       AddedDate = rsRuleProp("AddedDate")
       DueDate = rsRuleProp("DueDate")
       strRPNo = rsRuleProp("RPNo")
       iRPId = rsRuleProp("RPId")
       DueDateList = FormatDateTimeISO(rsRuleProp("DueDate"), False)
       iDaysToExpiry = Datediff("d",Now(),DueDate)
       strDaysToExpire = ""
       If Abs(iDaysToExpiry) > 1 Then strPlural = "s" Else strPlural = "" 
       If iDaysToExpiry < 0 Then
         strDaysToExpire = "Expired " & Abs(iDaysToExpiry) & " day" & strPlural & " ago"
       ElseIf iDaysToExpiry = 0 Then
         strDaysToExpire = "Expires today"
       ElseIf iDaysToExpiry < 6 Then
         strDaysToExpire = "Expires in " & iDaysToExpiry & " day" & strPlural
       End If
       If strDaysToExpire <> "" Then strDaysToExpire = "<font color=red>" & strDaysToExpire & "</font>"
       strNewFlag = ""          
       If Request.Cookies("LastVisit") <> "" Then
         If DateDiff("d", CDate(AddedDate), CDate(Request.Cookies("LastVisit")), vbMonday, vbFirstFourDays) < 5 Then strNewFlag = "<font color=red>New </font>"
       End If

       '## Find the comments that this user may read
       '## July 2005, using indexing tables instead of ;-separated lists
       If (bIsAdm OR bIsDNVemployee) Then   ' admins and DNV GL Employees may read all comments
         strHBrestrict = ""
       Else
         strHBrestrict = _ 
          " AND (" & _
               "   CommentID NOT IN ( " &_
               "     select CommentID from RestrictedComments " &_
               "     ) OR " &_
               "   CommentID IN ( " &_
               "     SELECT DISTINCT RC.CommentID " &_
               "     FROM RestrictedComments RC JOIN HB_Membership HBM ON RC.HearingBodyID = HBM.HearingBodyID " &_
               "     WHERE HBM.UserID = '" & strUserID & "'" &_
               "    ) " &_
               ") "
       End If


       '## Count number of new comments.
       SQLGetNewNoComments = _ 
         "SELECT RPId, COUNT(RPId) 'NoNewComments', dbo.TotalCommentsLengthKB(RPId, '" & strUserID & "') as 'TotalNewCommentLengthKB'" & _
         " FROM Comments" & _
         " WHERE RPId = "  & iRPId & _ 
           strHBrestrict & _
         " AND CommentID Not In " & _
                              "(SELECT * FROM dbo.tblCommentsReadBy('" & strUserID & "'))" & _ 
         " GROUP BY RPId"
       ' Response.Write "<p>SQLGetNewNoComments : " & SQLGetNewNoComments  & "</p>" ' DEVELOPMENT & DEBUG

       Set rsNewComments = dbConnect.Execute(SQLGetNewNoComments)
       If Err.Number <> 0 Then Call SeriousError

    
       Dim iNoNewComments, strNewCommentsKB, rsComments, SQLGetNoComments
       iNoNewComments = ""
       strNewCommentsKB = ""
       If Not rsNewComments.EOF Then       
          iNoNewComments = rsNewComments("NoNewComments")
          If iNoNewComments > 0 Then strNewCommentsKB = "&nbsp;(" & rsNewComments("TotalNewCommentLengthKB") & ")"
       End if

       rsNewComments.Close
       SET rsNewComments = Nothing
       
       '## Count total number of comments.
       SQLGetNoComments = _ 
         "SELECT RPId, COUNT(RPId) As 'NoComments', dbo.TotalCommentsLengthKB(RPId, '" & strUserID & "') as 'TotalCommentLengthKB'" & _
         " FROM Comments" & _
         " WHERE RPId = " & iRPId & _
           strHBrestrict & _
         " GROUP BY RPId "
       ' Response.Write"SQLGetNoComments : " & SQLGetNoComments & "<br>"        ' DEVELOPMENT & DEBUG
       Set rsComments = dbConnect.Execute(SQLGetNoComments)
       If Err.Number <> 0 Then Call SeriousError
       
       Dim iNoComments, strCommLengthKB
       iNoComments = ""
       strCommLengthKB = ""
       If Not rsComments.EOF Then       
          iNoComments = rsComments("NoComments")
          If iNoComments > 0 Then strCommLengthKB = "&nbsp;(" & rsComments("TotalCommentLengthKB") & ")"
       End if

  %>
       <tr>
         <td bgcolor="<%= color %>" nowrap><a href="review.asp?fileName=<%= rsRuleProp("FileName") %>&amp;RPId=<%= rsRuleProp("RPId") %>" target="_top"><%= Server.HTMLEncode(rsRuleProp("RPNo")) %></a></td>
         <td bgcolor="<%= color %>" nowrap align="center"><%=strNewFlag%><%=strDaysToExpire%></td>
         <td bgcolor="<%= color %>"><%= Trim(rsRuleProp("Title")) %></td>
         <td align=center bgcolor="<%= color %>" nowrap><%=DueDateList%></td>
         <td align=center bgcolor="<%= color %>" nowrap><%=iNoComments%><%=strCommLengthKB%></td>
         <%If True or bIsAdm Then ' Let all reviewers see how much is new %>
             <td align=center bgcolor="<%= color %>"><%=iNoNewComments%><%=strNewCommentsKB%></td>
         <%End If%>
       </tr>
  <% Loop
     rsComments.Close
     SET rsComments = Nothing %>
    </table>
<%   Else %>
    <font face="arial" color="#000000" size="-1"><p>Found no hearing documents</p></font>
<%   End If %>
    </td>
  </tr>
</table>

<% Else ' Show the archive  - because LCase(Request.QueryString("set")) = "archive" %>
<div style="margin-left: 12px">
<font face="arial" size="-1"><img src="images/GREENARROW.gif" border="0"> Go to <a style="color: blue" href="<%=Request.ServerVariables("SCRIPT_NAME")&"?set=active"%>">active hearings</a></font>
</div>
<!--
Link to official archives
Proposals including Minutes of Meetings (up to 2010): http://one.dnv.com/rulesecretariat/linkup/
Proposal archive including MoMs (from 2011): http://groups.dnv.com/sites/Rules_and_Standards/Lists/Proposal%20List1/Proposal%20archive%20including%20MoMs.aspx
-->
<div style="margin-top: 1em; margin-left: 12px">
<font face="arial" size="-1"><img src="images/GREENARROW.gif" border="0"> <a style="color: blue" target="_blank" href="http://one.dnv.com/rulesecretariat/linkup/">Official archive up to 2010 of Proposals including Minutes of Meetings</a></font><br>
<font face="arial" size="-1"><img src="images/GREENARROW.gif" border="0"> <a style="color: blue" target="_blank" href="http://groups.dnv.com/sites/Rules_and_Standards/Lists/Proposal%20List1/Proposal%20archive%20including%20MoMs.aspx">Official archive from 2011 of Proposals including MoMs</a></font>
</div>
<%
'***********************
'** Archived RPs
'***********************
DIM rsArchivedRP

strThisPageURL = Request.ServerVariables("SCRIPT_NAME")&"?set=archive"

'## July 2005 - uses indexing tables instead of ;-separated lists

If bIsAdm or bIsDNVemployee Then
  sqlWhere = ""   ' admins and DNV GL Employees may view all hearing documents
Else
  sqlWhere = _
      "AND ((RPId NOT IN " &_
      "         (SELECT RPId " &_
      "          FROM RestrictedRuleProps)) OR " &_
      "       (RPId IN " &_
      "         (SELECT RPId " &_
      "          FROM RestrictedRuleProps " &_
      "          WHERE HearingBodyID IN " &_
      "            (SELECT HBM.HearingBodyID " &_
      "             FROM  HB_Membership HBM " &_
      "             WHERE UserID = '" & strUserID & "')))) "
End If

Sql = _
      "SELECT RPId, RPNo, Title, FileName, DateArchived, DesignReviewDate " &_
      "FROM RuleProps " &_
      "WHERE DateArchived Is Not Null " & sqlWhere &_
      "ORDER BY " & strOrderRPby    'RPNo ASC"

' Response.Write "Sql=" & Sql                    ' DEVELOPMENT & DEBUG
Set rsArchivedRP = dbConnect.Execute(Sql)
If Err.Number <> 0 Then Call SeriousError

'## Generate the index only if there are documents stored in the database.
If Not rsArchivedRP.EOF Then %>

<table cellspacing="10">
  <tbody>
<tr>
<td colspan="3"><font face="Arial" size="+1" color="#333333"><strong>Archive (visible to DNV GL employees only)</strong></font>
</td>
</tr>
    <tr>
    <td><table class="RPlist" style="font-family: Arial; font-size: 10pt;" border="0" width="100%">
      <tbody><tr>
        <td bgcolor="#12b1ee" align="center"><font color="#ffffff" size="-1"><a href="<%=strThisPageURL%>&sortBy=RPNo" class="noformat"><b>Prop. No.<%=strFlagSorted("RPNo")%></b></a></font></td>
        <td bgcolor="#12b1ee"><font color="#ffffff" size="-1"><a href="<%=strThisPageURL%>&sortBy=Title" class="noformat"><b>Rule Proposal Title<%=strFlagSorted("Title")%></b></a></font></td>
<!--        <td align="center" bgcolor="#12b1ee"><font color="#ffffff" size="-1"><b>Minutes</b></font></td>  -->
        <td align="center" bgcolor="#12b1ee"><font color="#ffffff" size="-1"><a href="<%=strThisPageURL%>&sortBy=DesignReviewDate" class="noformat"><b>Design review<%=strFlagSorted("DesignReviewDate")%></b></a></font></td>
        <td align="center" bgcolor="#12b1ee"><font color="#ffffff" size="-1"><b>No. of Annexes</b></font></td>
      </tr>

<%
   fFirstPass = True
   '## Now loop through all the archived RPs
   Do
     If Not fFirstPass Then
        rsArchivedRP.MoveNext
     Else
        color = ""
        fFirstPass = False
     End If
     
     If rsArchivedRP.EOF Then Exit Do
     If color = "#FFFFFF" Then
        color = "#ffe1e2"
     Else
        color = "#FFFFFF"
     End If
  
     DIM sqlAnnexes, rsAnnexes, iAnnexes
     strRPNo = Trim(rsArchivedRP("RPNo"))
     iRPId = rsArchivedRP("RPId")

     '## Get number of annexes and dates for archiving
     sqlAnnexes = _ 
       "SELECT COUNT(RPId) As 'NoAnnexes'" & _
       " FROM RPAnnex" & _
       " WHERE RPId = "  & iRPId
     ' Response.Write "<p>sqlAnnexes : " & sqlAnnexes  & "</p>" ' DEVELOPMENT & DEBUG

     Set rsAnnexes = dbConnect.Execute(sqlAnnexes)
     If Err.Number <> 0 Then Call SeriousError

     iAnnexes = ""
     If Not rsAnnexes.EOF Then iAnnexes = rsAnnexes("NoAnnexes")
     rsAnnexes.Close
     SET rsAnnexes = Nothing
     
 %>
     <tr>
       <td bgcolor="<%= color %>" nowrap><a href="review.asp?view=archive&amp;fileName=<%=rsArchivedRP("FileName") %>&amp;RPId=<%=rsArchivedRP("RPId") %>" target="_blank"><%= Server.HTMLEncode(rsArchivedRP("RPNo")) %></a></td>
       <td bgcolor="<%= color %>"><%= Server.HTMLEncode(Trim(rsArchivedRP("Title"))) %></td>
<!--       <td align=center bgcolor="<%= color %>"><%If iAnnexes > 0 Then%>Published<%End If%></td>  -->
       <td align=center bgcolor="<%= color %>" nowrap><%=Server.HTMLEncode(FormatDateTimeISO(rsArchivedRP("DesignReviewDate"), False))%></td>
       <td align=center bgcolor="<%= color %>"><%=iAnnexes%></td>
     </tr>
  <% Loop %>
    </table>
<% Else %> 
    <font face="arial" color="#000000" size="-1"><p>Found no archived documents</p></font>
<% End If 
   rsArchivedRP.Close
   SET rsArchivedRP = Nothing
'***********************
'** End of Archived RPs
'***********************
End If       ' Show active vs. archive

'***********************
'**   Close connection.   **
'***********************
dbConnect.Close
SET dbConnect = Nothing
%> </td>
  </tr>
</table>
</body>
</HTML>
