<%@ LANGUAGE="VBSCRIPT" %>
<%
Option Explicit
DIM strToRecipient, strName, strAudience, strDefaultUserName, strDocRefPrompt, strNC, strRefReviewedPrompt
Dim dbConnect, iRPId
%>
<!--#include file="include/Functions.inc"-->
<!--#INCLUDE FILE="include/dbConnection.asp"-->
<%
' On Error Goto 0       ' DEVELOPMENT & DEBUG
 On Error Resume Next
IdentifyUser
If NOT mapUID2Member Then
Response.End
End If

' Is commenting allowed?
iRPId = CLng(Request("RPId"))
DIM strSQL, rsRP
SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError
strSQL = "SELECT CommentsRefused, MsgWhenCommentsRefused FROM RuleProps WHERE RPId = " & iRPId
Set rsRP = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError

If Not rsRP.EOF Then
  If rsRP("CommentsRefused") Then
  ' Commenting is refused. Build "No comments" frame
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<head><title><%=strAppTitle%> - Comments not allowed</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</head>
<body bgcolor="#F2F2F2">
<p align="center"><font face="arial" size="-1"><B>
<%
    If rsRP("MsgWhenCommentsRefused") <> "" Then
      Response.Write Server.HTMLencode(rsRP("MsgWhenCommentsRefused"))
    Else
      Response.Write strCommentsRefuseDefaultMessage
    End If
%>
</B></font>
</p>
</body>
</HTML>
<%
    Response.End
  End If
End If

' Commenting is allowed. Build normal comments frame
strDefaultUserName = "<e-mail address or name>"
' strDocRefPrompt = "<Document reference>"
strDocRefPrompt = "<Please specify Part, Chapter, Section, number, ..., as applicable>"   ' Changed 2015-01-21, Marval case 2149862
strNC = "Reviewed, no comment."
strRefReviewedPrompt = "<Insert reference to those parts reviewed>"
strAudience = "All readers"
If bIsCollectiveUserID(strUserID) Then
  If Request.Cookies("UsrParticulars") = "" Then
    strUserName = strDefaultUserName  ' "Collective UserIDs start with "Member_"
  Else
    strUserName = Request.Cookies("UsrParticulars")
  End If
End If
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<head><title><%=strAppTitle%> - Add a comment</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<!--#include file="include/SelectAudiencesSetup.inc"-->
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript">
var iWinX, iWinY, iAudiencesX, iAudiencesY, strAudiencesProps;

function initialize() {
  inpChosenAudiences  = document.frmAddComment.txtChosenAudiences;  // Global var, declared in include/SelectAudiencesSetup.inc
  if (document.getElementById('lblAudience') != null) {
    nodeAudience = document.getElementById('lblAudience').firstChild;  // Global var, declared in include/SelectAudiencesSetup.inc
    nodeAudience.nodeValue = arrAudienceNames["All"];  // initial value; works in IE6 and Mozilla 1.7
    }
  }

function clearPlacePrompt(txField) {
  if ('<%=strDocRefPrompt%>' == txField.value) {
    txField.value = "";
    }
  }

function restorePlacePrompt(txField) {
  if (txField.value == "") {
    txField.value = '<%=strDocRefPrompt%>';
    }
  }

function toggleNoComments(chkNC, taCmnt) {
  var strNC = "<%=strNC%>";
  var strNCrefPrompt = "<%=strRefReviewedPrompt%>";
  if (chkNC.checked) {
    if (taCmnt.value == "") {
      taCmnt.value = strNC;
      taCmnt.form.txtPlace.value = strNCrefPrompt;
      }
    else {
      if (confirm('Erase your current comment?')) {
        taCmnt.value = strNC;
        taCmnt.form.txtPlace.value = strNCrefPrompt;
        }
      else {
        chkNC.checked = false;
        }
      }
    }
  else if (strNC == taCmnt.value) {
    taCmnt.value = "";
    taCmnt.form.txtPlace.value = '<%=strDocRefPrompt%>';
    }
  }

/*
/* Form validation:
/*  - Place reference: Do you wish to specify place reference OK/Cancel? OK > back to form; Cancel -> submit, strip strDocRefPrompt
/*  - Comment from: other than default for "collective"?
/*  - Comment: any contents?
/*  - let recieveing ASP script handle SQL injection and XSS attempts
*/
function validate(frmForm) {
  var bIsValid = false;
//  bIsValid = '<%=strDefaultUserName%>' != frmForm.txtMember.value;
  if ('<%=strDefaultUserName%>' == frmForm.txtMember.value) {
    alert('Please fill in your name and (optionally) the organisation you represent.');
    frmForm.txtMember.focus();
    frmForm.txtMember.select();
    return false;
    }
//  bIsValid = bIsValid && ("" != frmForm.Comment.value);
  bIsValid = ("" != frmForm.Comment.value);
  if (bIsValid) {
    if ('<%=strDocRefPrompt%>' == frmForm.txtPlace.value || '<%=strRefReviewedPrompt%>' == frmForm.txtPlace.value) {
      if (confirm('Would you please fill in the Document reference?')) {
        frmForm.txtPlace.focus();
        return false;
        }
      }
    }
  if (bIsValid) {
    if ('<%=strDocRefPrompt%>' == frmForm.txtPlace.value || '<%=strRefReviewedPrompt%>' == frmForm.txtPlace.value) {
      frmForm.txtPlace.value = "";  // Rather blank than non-informational default prompts
      }
    frmForm.submit();   
    }
  else {
    alert("Please fill in the comment form");
    return false;
    }
  }

</SCRIPT>
</head>


<% 'Define e-mail address for "To" recipient.
strToRecipient = "Document.controller@dnvgl.com"
'strToRecipient = "Bo.Johanson@dnvgl.com"  ' Debug
%>

<body bgcolor="#F2F2F2" onLoad="initialize();">

<%
   If iRPId = "" Then %>
      <br><font face="arial" size="-1">If a document is selected, this frame will provide you with an input form for adding comments.</font>
      </body>
      </html>
<%      Response.End
   End If 
%>

<form name="frmAddComment" method="POST" action="addCommentAction.asp?RPId=<%= iRPId %>">
  <font face="arial" size="-1"><B>Comment from:</B></font>
  <br>
  <input type="text" name="txtMember" value="<%=strCommentBy()%>" size="25" Title="Your e-mail address, or name, company, organisation, etc.">
  <div style="position: relative; left: 12em; top: -5ex; " >
    <font face="arial" size="-1"><B>Committee:</B></font><br>
    <font face="arial" size="-1"  Title="The DNV GL hearing body you're a member of"><%=Server.HTMLencode(strHearingBodies)%></font>
  </div>
<div style="margin-top: -4ex;" >
  <font face="arial" size="-1"><B>Comment:</B></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="Checkbox" NAME="chkNoComment" onClick="toggleNoComments(this, this.form.Comment)" Title="Check when you agree entirely with the proposal">&nbsp;<font face="arial" size="-1">Reviewed, no comment<br></font>
  <font face="arial" size="-1">
    <script lanugage="javascript" type="text/javascript">document.write('<input type="text" name="txtPlace" value="<%=strDocRefPrompt%>" style="width: 440px;font-family: Arial; font-size: 10pt" Title="Please refer to the place in the hearing document" onFocus="clearPlacePrompt(this)" onBlur="restorePlacePrompt(this)">')</script>
  </font><br>
  <font face="arial" size="-1"><textarea rows="12" name="Comment" style="width: 440px;font-family: Arial; font-size: 10pt" Title="Type your comment here"></textarea></font>
<%
' Disable "Restricted" comments. BGJ 2010-04-20. See "AddCommentAction.asp", statement "If txtChosenAudiences <> "" Then strSQL = ..."
' <br>
'  <input type="hidden" name="txtChosenAudiences" value="">
'  <input style="font-size: small;" name="btnAudience" type="button" value="Change >" name="btnChangeVisibility" Title="Click to review ' and change audience" onClick="openAudiences(event);">&nbsp;
'  <font face="arial" size="-1"><bbb><span id="lblAudience" style="background-color: #FFFFFF; color: #000000;">&nbsp;</span></bbb>&nbsp;may view this comment</font>
%>
  <br>
  <input style="float: left; position: relative; top: 1ex;" type="button" style="font-size: small" value="Submit comment" name="btnSubmit" onClick="validate(this.form)" Title="Submit your comment to DNV GL">
  <span style="position: relative;left: 1ex; "><p style="display: inline; font-family: Arial; font-size: 70%;">If you have any comments you wish to make confidentially, please contact your Customer Service Manager</p></span>
</div>
</form>
</body>
</HTML>

<% SelectAudiencesCleanUp %>