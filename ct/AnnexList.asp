<%@ LANGUAGE="VBSCRIPT" %>
<% 
Option Explicit
DIM dbConnect, iRPId, strOutputCommentBy, strCommenterParticulars, rs
Dim strSQL, rsRPAnnex, strAnnexTitle
Dim strHBrestrict, iID
Dim fso, strSrcPath, strFileName, strDestFolder, fFirstPass
%>
<%  '## This script lists all annexes to a particular archived rule hearing document.
    '## Based on comments.asp
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
  iRPId = CLng(Request("RPId"))    ' strRPNo = Request("RPNo")
  Response.Charset = "ISO-8859-1"  ' Enforcing character set helps preventing cross-site scripting

%>
<HTML>
<head>
<title><%=strAppTitle%> - Annexes</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
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

<body bgcolor="#F2F2F2">
<P>
<!-- <A HREF="http://www.dnvgl.com/" TARGET="_top"><IMG SRC="images/dnv_logo.gif" BORDER="0" ALIGN="bottom" ALT="DNV GL home"></A> -->
<A HREF="http://www.dnvgl.com/" TARGET="_top"><IMG SRC="images/DNVGL_logo_small.png" BORDER="0" ALIGN="bottom" ALT="DNV GL home"></A>
&nbsp;&nbsp;&nbsp;
<FONT face=arial><%=strAppTitle%></FONT>
</P>
<div class="ScreenOnly"><FONT face="arial" size="-2"><A href="ruleDocs.asp?set=archive" target=_top>To the archive list</A>
  <span style="position: absolute; right: 30px;">
    <a href="javascript:undockDOC()">Open document in own window</a>
  </span>
  </FONT>
</div>


<%  '## RPId must be given (the document is identified by this number)
'## Find the annexes to this RP
strSQL = _ 
  "SELECT A.FileName, A.RPId, RP.RPNo As RPNo, A.AnnexTitle" & _
  " FROM RuleProps RP Left Join RPAnnex A on RP.RPId = A.RPId" & _
  " WHERE RP.RPId = " & iRPId & _
  " ORDER BY AnnexTitle ASC"
' Response.Write strSQL & "<br>"   ' DEVELOPMENT & DEBUG
Set rsRPAnnex = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError
%>

<%  If FALSE AND rsRPAnnex.EOF Then %>
<p><font face="arial" size=2>If a document is selected, this window will list all comments related to that document.</font></p> 
<%    Response.End
  End If %>

<p>
<strong><font face="arial" size=3>Minutes and Annexes to <font color="blue"><%= Server.HTMLEncode(Trim(rsRPAnnex("RPNo"))) %></font>:</font></strong><br>
</p>
<%

' An RP with no annex will have a "pseudo response" - a record with NULL except for the RPNo retrieved from RuleProps in the SQL Join above
While IsNull(rsRPAnnex("RPId"))
  rsRPAnnex.MoveNext
Wend

'## Copy annex files to user's doc. buffer; List the annexes
If Not rsRPAnnex.EOF Then
  set fso = Server.CreateObject("Scripting.FileSystemObject")
%>
<table bgcolor="#FFFFFF" border="0" width="100%" style="font-family: Arial; font-size: 10pt">
<%
  fFirstPass = True
  Do
    If Not fFirstPass Then
      rsRPAnnex.MoveNext
    Else
      fFirstPass = False
    End If
    If rsRPAnnex.EOF Then Exit Do
    strFileName = Trim(rsRPAnnex("FileName"))
    strSrcPath = strDocRepository & strFileName
    strDestFolder = strDocBuffer & strUserID & "\"
    If Not fso.FolderExists(strDestFolder) Then
      fso.CreateFolder(strDestFolder)
    End If
    If fso.FileExists(strSrcPath) Then
      ' Copy the file to buffer
       fso.CopyFile strSrcPath, strDestFolder, True
    End If
    strAnnexTitle = Trim(rsRPAnnex("AnnexTitle"))
%>
  <tr>
    <td>
      <a href="<%=Replace(Server.URLEncode("docbuf/" & strUserID & "/" & strFileName), "+", "%20")%>" target="_blank"><%=Server.HTMLencode(strAnnexTitle)%></a><br>
    </td>
  </tr>  <%
  Loop %>
</table>
<%
'## no annexes exists for this document
Else %>
<p><font face="arial" size="-1"><br>
<br>
No annexes</font></p> 
<%
End If
%>
<p><font face="arial" size="-1">[<a href="comments.asp?RPId=<%=iRPId%>">View comments from hearing</a>]</p>

<%
'***********************
'**  Close connection.  **
'***********************
Set rsRPAnnex = Nothing
dbConnect.Close
SET dbConnect = Nothing
%>

</body>
</HTML>

