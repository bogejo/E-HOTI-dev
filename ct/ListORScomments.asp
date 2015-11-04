<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
'## ListORScomments.asp
'## Query parameters:
'##   none
%>
<%
On Error Resume Next
 On Error Goto 0   ' DEVELOPMENT & DEBUG
%>
<!--#include file="include/Functions.inc"-->
<!--#INCLUDE FILE="include/dbConnection.asp"-->
<%
IdentifyUser
If Not bIsAdm Then
  Call SeriousError
End If
%>

<%
DIM dbConnect, rsUsers, strFormStyle, strSubTitle, strSQL, strTableRowStyle, strTableHeadStyle

Response.Clear
Response.Buffer = False  ' Allows output of this now "huge" list of users, more than 7.200 members
' Response.ContentType = "text/plain"
Response.ContentType = "application/vnd.ms-excel;charset=UTF-8"
Response.AddHeader "Content-Disposition","attachment; filename=" & "E-HOTI_ORS_Comments.xls"

Dim dbCmd, sResponseStream, strResponse, sOutStream

SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
SET dbCmd = Server.CreateObject("ADODB.Command")
SET sResponseStream = Server.CreateObject("ADODB.Stream")
sResponseStream.Charset = "windows-1252"
sResponseStream.Open
If Err.Number <> 0 Then Call SeriousError


If Err.Number <> 0 Then Call SeriousError
'## Get comments on XML format
strSQL = "select C.CommentID as [CommentID_E-HOTI], RP.RPNo as [Proposal_no.], RP.Title as [Document_title], C.CommentToPlace as [E-HOTI_Reference], " &_
         "concat('', C.CommentToPageNo) as [Page_no.], convert(xml, U.NameUser) as [Commentator_name], U.Organisation as [Commentator_organisation], " &_
         "C.CommentByUserID as [Commentator_email], convert(ntext, C.Comment) as [Comment], " &_
"case when LTrim(RTrim(concat('', C.Comment))) Like 'Reviewed, no comment.%' then 'True' else '' end As [No_action_needed], " &_
         "RP.DueDate as [Due_date] " &_
         "from RuleProps RP join Comments C on RP.RPId = C.RPId join Users U on C.CommentByUserID = U.UserID where RP.RPNo like 'ORS%' order by C.CommentID " &_
         "FOR XML PATH('CommentRecord'), Type, ELEMENTS XSINIL"

' Response.Write"strSQL=" & strSQL & "<br>"       ' DEVELOPMENT & DEBUG
' Response.End          ' DEVELOPMENT & DEBUG

dbCmd.ActiveConnection = dbConnect
' Response.Write "dbCmd.Properties.Count=" & dbCmd.Properties.Count & "<br>"        ' DEVELOPMENT & DEBUG
' Dim i
' For i = 0 To dbCmd.Properties.Count - 1
'   Response.Write i & ": Name: " & dbCmd.Properties(i).Name & ", Type: " & dbCmd.Properties(i).Type & ", Value: " & dbCmd.Properties(i).Value & "<br>"
' Next 
' Response.Write "dbCmd.Properties(""Output stream"")=" & dbCmd.Properties("Output stream") & "<br>"        ' DEVELOPMENT & DEBUG
dbCmd.Properties("Output stream") = sResponseStream
dbCmd.Properties("xml root") = "Comments"
' Response.Write "dbCmd.Properties(""Output stream"").Name=" & dbCmd.Properties("Output stream").Name & "<br>"        ' DEVELOPMENT & DEBUG
' Response.Write "dbCmd.Properties(""Output stream"").Attributes=" & dbCmd.Properties("Output stream").Attributes & "<br>"        ' DEVELOPMENT & DEBUG
dbCmd.CommandText = strSQL
dbCmd.CommandType = 1 ' adCmdText = 1
' Response.Write "dbCmd.CommandText=" & dbCmd.CommandText & "<br>"        ' DEVELOPMENT & DEBUG
' Response.Write "sResponseStream.Size=" & sResponseStream.Size & "<br>"        ' DEVELOPMENT & DEBUG
' Response.Write "sResponseStream.Mode=" & sResponseStream.Mode & "<br>"        ' DEVELOPMENT & DEBUG
dbCmd.Execute ,, 1024  ' adExecuteStream = 1024
' Response.Write "sResponseStream.Size=" & sResponseStream.Size & "<br>"        ' DEVELOPMENT & DEBUG
' Response.Write "sResponseStream.Mode=" & sResponseStream.Mode & "<br>"        ' DEVELOPMENT & DEBUG
sResponseStream.Position = 0
' strResponse = sResponseStream.ReadText

' This manouvering with two streams, response and output, is an attempt to convert problematic characters from the database's character set to UTF-8.
' Unfortunately, it does not have the expected effect. "Illegal" characters do not get converted. The problematic items include:
' - HTML character entities - &amp; The "&" fails in XML
' - 'æøå'
' The approach was inspired by http://www.sqlservercentral.com/Forums/Topic1431690-146-1.aspx

Set sOutStream = Server.CreateObject("ADODB.Stream")
sOutStream.Open
sOutStream.Charset = "utf-8"
sOutStream.WriteText sResponseStream.ReadText
sOutStream.Position = 0
' strResponse = sOutStream.ReadText  
' Response.Write strResponse
Response.Write sOutStream.ReadText

   '***********************
   '**   Close connection.   **
   '***********************
   Set dbCmd = Nothing
   sResponseStream.Close
   Set sResponseStream = Nothing
   sOutStream.Close
   Set sOutStream = Nothing
   dbConnect.Close
   SET dbConnect = Nothing
   Response.End
%>

