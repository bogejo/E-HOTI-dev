<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<%
Dim dbConnect   ' Used in SelectAudiencesSetup.inc
On Error Resume Next
' On Error GoTo 0        ' DEVELOPMENT & DEBUG

IdentifyUser
If Not bIsAdm Then Call SeriousError
%>

<%
Dim iYearNow, iYr, DateDueDefault, iYearDueDefault, iMonthDueDefault, iDayDueDefault
iYearNow = Year(Date)
DateDueDefault = DateAdd("m", 1, Date)
iYearDueDefault = Year(DateDueDefault)
iMonthDueDefault = Month(DateDueDefault)
iDayDueDefault = Day(DateDueDefault)

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%=strAppTitle%> - Add Rule Proposal Document</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<!--#include file="../include/SelectAudiencesSetup.inc"-->
</head>
<script language="JavaScript">
<!-- Hide me
  
function initialize() {
  inpChosenAudiences  = document.RPdata.txtChosenAudiences;  // Global var, declared in include/SelectAudiencesSetup.inc
  nodeAudience = document.getElementById('lblAudience').firstChild;  // Global var, declared in include/SelectAudiencesSetup.inc
  nodeAudience.nodeValue = arrAudienceNames["All"];  // initial value; works in IE6 and Mozilla 1.7
  }

/*************************************************************/
/* CheckLengthOfInput is called when user leaves the field.  */
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
<body style="max-width: 1000px" bgcolor="#FFFFFF" onLoad="initialize();document.RPdata.RPNo.focus()">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">Add a Rule Proposal Document</font></strong></td>
  </tr>
</table>

<form method="POST" name="RPdata" action="addRuleDocAction.asp" encType="multipart/form-data">
  <table border="0" cellspacing="4">
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Rule Proposal
      no:</td>
      <td style="font-family: Arial; font-size: 10pt"><input type="text"
      name="RPNo" size="85" onBlur="CheckLengthOfInput(this, 25)"></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Rule Proposal doc. title:</td>
      <td style="font-family: Arial; font-size: 10pt"><input type="text" name="Title" size="85"></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Rule Proposal
      file:</td>
      <td style="font-family: Arial; font-size: 10pt"><INPUT type="File" name="File1" size="70"></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Due Date:</td>
      <td style="font-family: Arial; font-size: 10pt"> <select name="DueYear" size="1">
<% For iYr = iYearNow To iYearNow + 5 %>
        <option value="<%=iYr%>"<%If iYr = iYearDueDefault Then%> selected<%End If%>><%=iYr%></option>
<% Next %>
      </select>(Year)&nbsp; <select name="DueMonth" size="1">
        <option value="01"<%If iMonthDueDefault = 1 Then %> selected<%End If%>>01</option>
        <option value="02"<%If iMonthDueDefault = 2 Then %> selected<%End If%>>02</option>
        <option value="03"<%If iMonthDueDefault = 3 Then %> selected<%End If%>>03</option>
        <option value="04"<%If iMonthDueDefault = 4 Then %> selected<%End If%>>04</option>
        <option value="05"<%If iMonthDueDefault = 5 Then %> selected<%End If%>>05</option>
        <option value="06"<%If iMonthDueDefault = 6 Then %> selected<%End If%>>06</option>
        <option value="07"<%If iMonthDueDefault = 7 Then %> selected<%End If%>>07</option>
        <option value="08"<%If iMonthDueDefault = 8 Then %> selected<%End If%>>08</option>
        <option value="09"<%If iMonthDueDefault = 9 Then %> selected<%End If%>>09</option>
        <option value="10"<%If iMonthDueDefault = 10 Then %> selected<%End If%>>10</option>
        <option value="11"<%If iMonthDueDefault = 11 Then %> selected<%End If%>>11</option>
        <option value="12"<%If iMonthDueDefault = 12 Then %> selected<%End If%>>12</option>
      </select>(Month)&nbsp; <select name="DueDay"
      size="1">
        <option value="01"<%If iDayDueDefault = 1 Then %> selected<%End If%>>01</option>
        <option value="02"<%If iDayDueDefault = 2 Then %> selected<%End If%>>02</option>
        <option value="03"<%If iDayDueDefault = 3 Then %> selected<%End If%>>03</option>
        <option value="04"<%If iDayDueDefault = 4 Then %> selected<%End If%>>04</option>
        <option value="05"<%If iDayDueDefault = 5 Then %> selected<%End If%>>05</option>
        <option value="06"<%If iDayDueDefault = 6 Then %> selected<%End If%>>06</option>
        <option value="07"<%If iDayDueDefault = 7 Then %> selected<%End If%>>07</option>
        <option value="08"<%If iDayDueDefault = 8 Then %> selected<%End If%>>08</option>
        <option value="09"<%If iDayDueDefault = 9 Then %> selected<%End If%>>09</option>
        <option value="10"<%If iDayDueDefault = 10 Then %> selected<%End If%>>10</option>
        <option value="11"<%If iDayDueDefault = 11 Then %> selected<%End If%>>11</option>
        <option value="12"<%If iDayDueDefault = 12 Then %> selected<%End If%>>12</option>
        <option value="13"<%If iDayDueDefault = 13 Then %> selected<%End If%>>13</option>
        <option value="14"<%If iDayDueDefault = 14 Then %> selected<%End If%>>14</option>
        <option value="15"<%If iDayDueDefault = 15 Then %> selected<%End If%>>15</option>
        <option value="16"<%If iDayDueDefault = 16 Then %> selected<%End If%>>16</option>
        <option value="17"<%If iDayDueDefault = 17 Then %> selected<%End If%>>17</option>
        <option value="18"<%If iDayDueDefault = 18 Then %> selected<%End If%>>18</option>
        <option value="19"<%If iDayDueDefault = 19 Then %> selected<%End If%>>19</option>
        <option value="20"<%If iDayDueDefault = 20 Then %> selected<%End If%>>20</option>
        <option value="21"<%If iDayDueDefault = 21 Then %> selected<%End If%>>21</option>
        <option value="22"<%If iDayDueDefault = 22 Then %> selected<%End If%>>22</option>
        <option value="23"<%If iDayDueDefault = 23 Then %> selected<%End If%>>23</option>
        <option value="24"<%If iDayDueDefault = 24 Then %> selected<%End If%>>24</option>
        <option value="25"<%If iDayDueDefault = 25 Then %> selected<%End If%>>25</option>
        <option value="26"<%If iDayDueDefault = 26 Then %> selected<%End If%>>26</option>
        <option value="27"<%If iDayDueDefault = 27 Then %> selected<%End If%>>27</option>
        <option value="28"<%If iDayDueDefault = 28 Then %> selected<%End If%>>28</option>
        <option value="29"<%If iDayDueDefault = 29 Then %> selected<%End If%>>29</option>
        <option value="30"<%If iDayDueDefault = 30 Then %> selected<%End If%>>30</option>
        <option value="31"<%If iDayDueDefault = 31 Then %> selected<%End If%>>31</option>
      </select>(Day)</td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">Disable comments:</td>
      <td style="font-family: Arial; font-size: 10pt"><input type="checkbox" name="chkRefuseComments" onClick="NoCommentsToggle(this);"></td>
    </tr>
    <tr>
      <td align="right" style="font-family: Arial; font-size: 10pt">"No comments" message:</td>
      <td style="font-family: Arial; font-size: 10pt"><input type="text" name="txtNoCommentsMsg" size="85" style="color:lightgrey" onFocus="NoCommentsFocus(this);" value="<%=strCommentsRefuseDefaultMessage%>"></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>
        <input type="hidden" name="txtChosenAudiences" value="">
        <input style="font-size: small;" name="btnAudience" type="button" value="Change >" name="btnChangeVisibility" Title="Click to review and change audience" onClick="openAudiences(event);">&nbsp;
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
        <input type="submit" value="Add to database">
      </td>
    </tr>
  </table>
</form>
</body>
</html>
