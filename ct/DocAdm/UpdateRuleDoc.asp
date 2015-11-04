<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<%
'## Query parameters
'##   RPId:  the hearing document's ID (integer)
'## Adapted to indexing tables instead of ;-terminated list for HearingBodies / 2005-07-21
%>
<%
On Error Resume Next
' On Error Goto 0   ' DEVELOPMENT & DEBUG
IdentifyUser
If Not bIsAdm Then Call SeriousError
%>

<%
Dim dbConnect, iYearNow, iYr, iMth, iDay
Dim iYearPrevDue, iMonthPrevDue, iDayPrevDue
Dim iYearPrevDR, iMonthPrevDR, iDayPrevDR
Dim iRPId, strRPNo, strRPtitle, strDueDate, strDRDate, strRPfilename
Dim strMth, strDay, strToHearingBodies, rsRP, strSQL, strRPFileLastModified, rsRestrToHB
Dim bCommentsRefused, strMsgWhenCommentsRefused
iYearNow = Year(Date)

iRPId = CLng(Request.QueryString("RPId"))   ' strRPNo = Request.QueryString("RPNo")
Set dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError
strSQL = "SELECT RPId, RPNo, Title, DueDate, AddedDate, FileName, DesignReviewDate, CommentsRefused, MsgWhenCommentsRefused FROM RuleProps Where RPId = " & iRPId
Set rsRP = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError
If rsRP.EOF Then
  Set rsRP = Nothing
  dbConnect.Close
  Call NoSuchRP(iRPId)
Else
  strRPNo = rsRP("RPNo")
  strRPtitle = rsRP("Title")
  If IsNull(strRPtitle) Then strRPtitle = ""
  strDueDate = rsRP("DueDate")
  If IsNull(strDueDate) Then strDueDate = ""
  strRPfilename = rsRP("FileName")
  strDRDate = rsRP("DesignReviewDate")
  If IsNull(strDRDate) Then strDRDate = ""
  If IsNull(strRPfilename) Then strRPfilename = ""
  bCommentsRefused = rsRP("CommentsRefused")
  strMsgWhenCommentsRefused = rsRP("MsgWhenCommentsRefused")
  If IsNull(strMsgWhenCommentsRefused) Then strMsgWhenCommentsRefused = strCommentsRefuseDefaultMessage

  strSQL = "SELECT RRP.HearingBodyID FROM RestrictedRuleProps RRP JOIN HearingBodies HB ON RRP.HearingBodyID = HB.ID WHERE RPId = " & iRPId & " ORDER BY HB.NameHB"
  ' Response.Write "strSQL=" & strSQL & "<br>"        ' DEVELOPMENT & DEBUG
  ' Response.End                             ' DEVELOPMENT & DEBUG
  Set rsRestrToHB = dbConnect.Execute(strSQL)
  If Err.Number <> 0 Then Call SeriousError
  strToHearingBodies = ""   ' Build a ;-terminated list of HB-IDs where restricted
  While Not rsRestrToHB.EOF 
    strToHearingBodies = strToHearingBodies & rsRestrToHB("HearingBodyID") & ";"
    rsRestrToHB.MoveNext
  Wend
  strRPFileLastModified = CopyToDocBuf(strRPfilename)
End If

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%=strAppTitle%> - Update Document</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<!--#include file="../include/SelectAudiencesSetup.inc"-->
</head>
<script language="JavaScript">
<!-- Hide me
  
function initialize() {
  inpChosenAudiences  = document.RPdata.txtChosenAudiences;  // Global var, declared in include/SelectAudiencesSetup.inc
  nodeAudience = document.getElementById('lblAudience').firstChild;  // Global var, declared in include/SelectAudiencesSetup.inc
  nodeAudience.nodeValue = arrAudienceNames["All"];  // initial value; works in IE6 and Mozilla 1.7
  fnUpdateChosenAudiences("<%=strToHearingBodies%>");
  }

/*************************************************************/
/* CheckLengthOfInput is Called when user leaves the field.  */
/* Checks if the length of the filed is not exceeded.        */
/*************************************************************/
function CheckLengthOfInput(InputObject, MaxLength)
{
  var InputString = InputObject.value;
  if(InputString.length > MaxLength)
  {
    alert("The value you entered was too long. Max length is " + MaxLength + ".");
    InputObject.select();
    InputObject.focus();
  }
}

function NoCommentsToggle(chk)
{
  if (chk.checked) {
    chk.form.txtNoCommentsMsg.style.color="black";
    }
  else {
    chk.form.txtNoCommentsMsg.style.color="lightgrey";
    }      
}

function NoCommentsFocus(txtBox) {
  if (!txtBox.form.chkRefuseComments.checked) {
    txtBox.blur();   
    }
  }

//End -->
</script>
<body style="max-width: 1000px" bgcolor="#FFFFFF" onLoad="initialize();document.RPdata.Title.focus()">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">Update a Document</font></strong></td>
  </tr>
</table>

<form method="POST" name="RPdata" action="UpdateRuleDocAction.asp" encType="multipart/form-data">
  <table border="0" cellspacing="4">
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Rule Proposal:</td>
      <td style="font-family: Arial; font-size: 10pt"><b><%=Server.HTMLencode(strRPNo)%></b>
      <input type="hidden" name="RPId" value="<%= iRPId %>"></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Document title:</td>
      <td style="font-family: Arial; font-size: 10pt"><input type="text"
      name="Title" size="85" value="<%=Server.HTMLencode(strRPtitle)%>"></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt;">Document file:
      </td>
      <td style="font-family: Arial; font-size: 10pt"><a href="<%=Replace(Server.URLEncode("../docbuf/" & strUserID & "/" & strRPfilename), "+", "%20")%>" target="DocWindow"><%=Server.HTMLencode(strRPfilename)%></a>
      <!-- input type="hidden" name="txtFileName" value="<%=Server.HTMLencode(strRPfilename)%>" -->
      </td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt" nowrap>Optional new Document file:</td>
      <td style="font-family: Arial; font-size: 10pt"><INPUT type="File" name="File1" size="70"></td>
    </tr>
<%
 iYearPrevDue = Year(strDueDate)
 iMonthPrevDue = Month(strDueDate)
 iDayPrevDue = Day(strDueDate)
 %>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Due Date:</td>
      <td style="font-family: Arial; font-size: 10pt"> <select name="DueYear" size="1">
<% For iYr = iYearNow To iYearNow + 5 %>
        <option value="<%=iYr%>" <%If iYr = iYearPrevDue Then%> selected<%End If%>><%=iYr%></option>
<% Next %>
      </select>(Year)&nbsp; <select name="DueMonth" size="1">
<% For iMth = 1 To 12
     strMth = Right("0" & iMth, 2) %>
        <option value="<%=strMth%>" <%If iMth = iMonthPrevDue Then%> selected<%End If%>><%=strMth%></option>
<% Next %>
      </select>(Month)&nbsp; <select name="DueDay"
      size="1">
<% For iDay = 1 To 31
     strDay = Right("0" & iDay, 2) %>
        <option value="<%=strDay%>" <%If iDay = iDayPrevDue Then%> selected<%End If%>><%=strDay%></option>
<% Next %>
      </select>(Day)</td>
    </tr>

<% ' Date for Design Review %>
<%
 If strDRDate = "" Then
   iYearPrevDR = ""
   iMonthPrevDR = ""
   iDayPrevDR = ""
 Else
   iYearPrevDR = Year(strDRDate)
   iMonthPrevDR = Month(strDRDate)
   iDayPrevDR = Day(strDRDate)
 End If
%>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Date for Design Review:</td>
      <td style="font-family: Arial; font-size: 10pt"> <select name="DRYear" size="1">
        <option value="" <%If iYr = iYearPrevDR Then%> selected<%End If%>></option>
<% For iYr = iYearNow To iYearNow + 5 %>
        <option value="<%=iYr%>" <%If iYr = iYearPrevDR Then%> selected<%End If%>><%=iYr%></option>
<% Next %>
      </select>(Year)&nbsp; <select name="DRMonth" size="1">
        <option value="" <%If iMth = iMonthPrevDR Then%> selected<%End If%>></option>
<% For iMth = 1 To 12
     strMth = Right("0" & iMth, 2) %>
        <option value="<%=strMth%>" <%If iMth = iMonthPrevDR Then%> selected<%End If%>><%=strMth%></option>
<% Next %>
      </select>(Month)&nbsp; <select name="DRDay"
      size="1">
        <option value="" <%If iDay = iDayPrevDR Then%> selected<%End If%>></option>
<% For iDay = 1 To 31
     strDay = Right("0" & iDay, 2) %>
        <option value="<%=strDay%>" <%If iDay = iDayPrevDR Then%> selected<%End If%>><%=strDay%></option>
<% Next %>
      </select>(Day)</td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Disable comments:</td>
      <td style="font-family: Arial; font-size: 10pt"><input type="checkbox"
      name="chkRefuseComments" size="85" onClick="NoCommentsToggle(this);" <%If bCommentsRefused Then%>checked<%End If%>></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">"No comments" message:</td>
      <td style="font-family: Arial; font-size: 10pt"><input type="text"
      name="txtNoCommentsMsg" size="85" style="color:<%If bCommentsRefused Then%>black<%Else%>lightgrey<%End If%>" onFocus="NoCommentsFocus(this);" value="<%=Server.HTMLencode(strMsgWhenCommentsRefused)%>"></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>
        <input type="hidden" name="txtChosenAudiences" value="">
        <input name="btnAudience" type="button" value="Change >" name="btnChangeAudience" Title="Click to review and change audience" onClick="openAudiences(event);">&nbsp;
        <font face="arial" size="-1"><bbb><span id="lblAudience" style="background-color: #FFFFFF; color: #000000;">&nbsp;</span></bbb>&nbsp;may view this document</font>
      </td>
    </tr>
<!--
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Added by:</td>
      <td style="font-family: Arial; font-size: 10pt"><%=strUserName%></td>
    </tr>
-->
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>
        <input type="submit" value="Update database">
      </td>
    </tr>
  </table>
</form>

<% '*** Minutes / Annexes *** %>
<p>&nbsp;</p>
<table border="0" width="700">
  <tbody><tr>
    <td width="100%"><strong><font color="#333333" face="Arial">Minutes / Annexes</font></strong><hr color="#12b1ee"</td>
  </tr>
</tbody></table>
<input type="Button" name="AddAnnex" value="Add Minutes / Annex" onClick="window.location.href='AddRPannex.asp?RPId=<%= iRPId %>&amp;RPNo=<%= Server.HTMLEncode(strRPNo) %>'">

<%
Dim rsRPAnnex, color, fFirstPass, strFileName, strAnnexFileLastModified
strSQL = "SELECT ID, RPId, AnnexTitle, FileName" & _
  " FROM RPAnnex" & _
  " WHERE RPId = " & iRPId & _
  " ORDER BY AnnexTitle ASC"
' Response.Write "<br>strSQL=" & strSQL & "<br>"    ' DEVELOPMENT & DEBUG
Err.Clear
Set dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
Set rsRPAnnex = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError
' Response.Write "Err.Number=" & Err.Number & "<br>"              ' DEVELOPMENT & DEBUG
' Response.Write "Err.Description=" & Err.Description & "<br>"    ' DEVELOPMENT & DEBUG
' Response.Write "rsRPAnnex.EOF=" & rsRPAnnex.EOF & "<br>"        ' DEVELOPMENT & DEBUG
If Not rsRPAnnex.EOF Then
  color = "#FFFFFF"
  fFirstPass = True
%>
<table style="font-family: Arial; font-size: 10pt; max-width: 700px" border="0">
  <tbody>
<%
  Do
    If Not fFirstPass Then
      rsRPAnnex.MoveNext
    Else
      color = ""
      fFirstPass = False
    End If
    If rsRPAnnex.EOF Then Exit Do
    If color = "#FFFFFF" Then
      color = "#ffe1e2"
    Else
      color = "#FFFFFF"
    End If
    strFileName = Trim(rsRPAnnex("FileName"))
    strAnnexFileLastModified = CopyToDocBuf(strFileName)

'    strSrcPath = strDocRepository & strFileName
'    strDestFolder = strDocBuffer & strUserID & "\"
'    If Not fso.FolderExists(strDestFolder) Then
'      fso.CreateFolder(strDestFolder)
'    End If
'    strFileLastModified = ""
'    If fso.FileExists(strSrcPath) Then
'      strFileLastModified = fso.GetFile(strSrcPath).DateLastModified
'      ' Copy the file to buffer
'       fso.CopyFile strSrcPath, strDestFolder, True
'    End If

%>
    <tr>
      <td style="padding-left: 5px; padding-right: 5px;" bgcolor="<%= color %>" nowrap valign="top" align="center"><font face="arial" size="-1"><a href="RemoveDocAction.asp?type=annex&ID=<%=Server.URLEncode(rsRPAnnex("ID"))%>" target="_top">remove</a></font></td>
      <td style="padding-left: 5px; padding-right: 5px;" bgcolor="<%= color %>" nowrap valign="top" align="center"><font face="arial" size="-1"><a href="UpdateRPannex.asp?type=annex&ID=<%=Server.URLEncode(rsRPAnnex("ID"))%>" target="_top">update</a></font></td>
      <td style="padding-left: 10px; padding-right: 20px;" bgcolor="<%= color %>"><a href="<%=Replace(Server.URLEncode("../docbuf/" & strUserID & "/" & strFileName), "+", "%20")%>" target="DocWindow"><%=Server.HTMLencode(Trim(rsRPAnnex("AnnexTitle")))%></a></td>
      <td style="padding-left: 5px; padding-right: 20px;" align="left" bgcolor="<%= color %>" nowrap><%=strFileName%></td>
      <td style="padding-left: 5px; padding-right: 20px;" bgcolor="<%= color %>" nowrap valign="top" align="left"><font face="arial" size="-1"><%=Server.HTMLEncode(FormatDateTimeISO(strAnnexFileLastModified, True))%></font></td>
    </tr>
<%
  Loop
%>
  </tbody>
</table>
<%
End If

Set rsRPAnnex = Nothing
%>


</body>
</html>

<%
If IsObject(dbConnect) Then
  dbConnect.Close
  Set dbConnect = Nothing
End If

Sub NoSuchRP(iRPId)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%=strAppTitle%> - Update Document</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</head>
<body style="max-width: 1000px" bgcolor="#FFFFFF">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">Update a Document</font></strong></td>
  </tr>
</table>
<P><font face="arial">Could not find a corresponding Rule Proposal</font></P>
<%
Response.End
End Sub

Function CopyToDocBuf(strFileName)
  Dim strSrcPath, strFileLastModified, strDestFolder, fso
  Set fso = Server.CreateObject("Scripting.FileSystemObject")
  strSrcPath = strDocRepository & strFileName
  strDestFolder = strDocBuffer & strUserID & "\"
  If Not fso.FolderExists(strDestFolder) Then
    fso.CreateFolder(strDestFolder)
  End If
  strFileLastModified = ""
  If fso.FileExists(strSrcPath) Then
    strFileLastModified = fso.GetFile(strSrcPath).DateLastModified
    ' Copy the file to buffer
     fso.CopyFile strSrcPath, strDestFolder, True
  End If
  Set fso = Nothing
  CopyToDocBuf = strFileLastModified
End Function

%>