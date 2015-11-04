<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<% ' Cookie based credentials from http://www.codefixer.com/codesnippets/cookieLogin.asp %>
<%
Response.Buffer = True 'Buffers the content so our Response.Redirect will work
%>
<!--#include file="include/Functions.inc"-->
<!--#INCLUDE FILE="include/dbConnection.asp"-->
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" runat="server" src="include/md5_PAJ.js"></SCRIPT>
<%
Dim bRenewPwd, iPwdDaysValid
On Error Resume Next
' On Error Goto 0   '  DEVELOPMENT & DEBUG
iPwdDaysValid = 3600   ' Max. no. of days allowed before password must be changed
bRenewPwd = False
If Trim(Request("pwdRen")) = "true" then bRenewPwd = true
If Request("chkChgPwd") = "1" Then bRenewPwd = true
%>

<html>
<head>
<title><%=strAppTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/JavaScript">
<!--
// Blow away any surrounding frames
if (top.document.URL != document.URL) {
	top.location.href = document.URL;
	}

function FDK_Validate(FormName, stopOnFailure, AutoSubmit, ErrorHeader)
{
 var theFormName = FormName;
 var theElementName = "";
 if (theFormName.indexOf(".")>=0)  
 {
   theElementName = theFormName.substring(theFormName.indexOf(".")+1)
   theFormName = theFormName.substring(0,theFormName.indexOf("."))
 }
 var ValidationCheck = eval("document."+theFormName+".ValidateForm")
 if (ValidationCheck)  
 {
  var theNameArray = eval(theFormName+"NameArray")
  var theValidationArray = eval(theFormName+"ValidationArray")
  var theFocusArray = eval(theFormName+"FocusArray")
  var ErrorMsg = "";
  var FocusSet = false;
  var i;
  var msg;
    
 
        // Go through the Validate Array that may or may not exist
        // and call the Validate function for all elements that have one.
  if (String(theNameArray)!="undefined")
  {
   for (i = 0; i < theNameArray.length; i ++)
   {
    msg="";
    if (theNameArray[i].name == theElementName || theElementName == "")
    {
      msg = eval(theValidationArray[i]);
    }
    if (msg != "")
    {
     ErrorMsg += "\n"+msg;                   
     if (stopOnFailure == "1") 
     {
       if (theFocusArray[i] && !FocusSet)  
      {
       FocusSet=true;
       theNameArray[i].focus();
       theNameArray[i].select();
      }
      alert(ErrorHeader+ErrorMsg);
      document.MM_returnValue = false; 
      break;
     }
     else  
     {
      if (theFocusArray[i] && !FocusSet)  
      {
       FocusSet=true;
       theNameArray[i].focus();
      }
     }
    }
   }
  }
  if (ErrorMsg!="" && stopOnFailure != "1") 
  {
   alert(ErrorHeader+ErrorMsg);
  }
  document.MM_returnValue = (ErrorMsg==""); 
  if (document.MM_returnValue && AutoSubmit)  
  {
   eval("document."+FormName+".submit()")
  }
 }
}

function FDK_StripChars(theFilter,theString)
{
	var strOut,i,curChar

	strOut = ""
	for (i=0;i < theString.length; i++)
	{		
		curChar = theString.charAt(i)
		if (theFilter.indexOf(curChar) < 0)	// if it's not in the filter, send it thru
			strOut += curChar		
	}	
	return strOut
}

function FDK_AddToValidateArray(FormName,FormElement,Validation,SetFocus)
{
    var TheRoot=eval("document."+FormName);
 
    if (!TheRoot.ValidateForm) 
    {
        TheRoot.ValidateForm = true;
        eval(FormName+"NameArray = new Array()")
        eval(FormName+"ValidationArray = new Array()")
        eval(FormName+"FocusArray = new Array()")
    }
    var ArrayIndex = eval(FormName+"NameArray.length");
    eval(FormName+"NameArray[ArrayIndex] = FormElement");
    eval(FormName+"ValidationArray[ArrayIndex] = Validation");
    eval(FormName+"FocusArray[ArrayIndex] = SetFocus");
 
}

function FDK_ValidateNonBlank(FormElement,ErrorMsg, iMinLength)
{
  var msg = ErrorMsg;
  var val = FormElement.value;  

  if ((FDK_StripChars(" \n\t\r",val).length != 0) && val.length >= iMinLength)
  {
     msg="";
  }

  return msg;
}

function FDK_AddNonBlankValidation(FormName,FormElementName,SetFocus,ErrorMsg, iMinLength)  {
  ErrorMsg = "'" + eval(ErrorMsg) + " (Min. " + iMinLength + " characters.)" + "'"
  var ValString = "FDK_ValidateNonBlank("+FormElementName+","+ErrorMsg+","+iMinLength+")"
  FDK_AddToValidateArray(FormName,eval(FormElementName),ValString,SetFocus)
}

function FDK_ValidateAlphaNum(FormElement,Required,ErrorMsg)
{
	var msg = "";
	var i, m, s, firstNonWhite
	var theString = FormElement.value;
 	var msgInvalid = ErrorMsg;

	if (FDK_StripChars(" ",theString).length == 0)	     {
		if (!Required)       {
          return "";		
        }
		else       {
          return msgInvalid;
        }
    }
	//Strip spaces off of the sides of the string
 	theString = FDK_Trim(theString);

    for (var n=0; n<theString.length; n++)     {
      theChar = theString.substring(n,n+1);
      if (!FDK_AllInRange("0","9",theChar) && !FDK_AllInRange("A","Z",theChar.toUpperCase()) && !(theChar == " "))     {
        return msgInvalid;
      }
    }

    return "";
}

function FDK_Trim(theString)
{
 var i,firstNonWhite

 if (FDK_StripChars(" \n\r\t",theString).length == 0 ) return ""

	i = -1
	while (1)
	{
		i++
		if (theString.charAt(i) != " ")
			break	
	}
	firstNonWhite = i
	//Count the spaces at the end
	i = theString.length
	while (1)
	{
		i--
		if (theString.charAt(i) != " ")
			break	
	}	

	return theString.substring(firstNonWhite,i + 1)

}

function FDK_AllInRange(x,y,theString)
{
	var i, curChar
	
	for (i=0; i < theString.length; i++)
	{
		curChar = theString.charAt(i)
		if (curChar < x || curChar > y) //the char is not in range
			return false
	}
	return true
}

function FDK_AddAlphaNumericValidation(FormName,FormElementName,Required,SetFocus,ErrorMsg)  {
  var ValString = "FDK_ValidateAlphaNum("+FormElementName+","+Required+","+ErrorMsg+")"
  FDK_AddToValidateArray(FormName,eval(FormElementName),ValString,SetFocus)
}

function requestNewPassword(txtUID) {
  window.open('SendNewPassword.asp?UID=' + txtUID,'PwdRequest','resizable=no,toolbar=no,status=no,width=550,height=500');
  }


//-->
</script>
</head>
<body style="max-width: 1000px" bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"
 onLoad="FDK_AddNonBlankValidation('credentials','document.credentials.txtUID',true,'\'Please type your user ID\'',6);FDK_AddNonBlankValidation('credentials','document.credentials.txtPwd',true,'\'Please type your password\'',8);<%If bRenewPwd Then%>FDK_AddNonBlankValidation('credentials','document.credentials.txtNewPwd',true,'\'Please type your new password\'',8);FDK_AddNonBlankValidation('credentials','document.credentials.txtRetypeNewPwd',true,'\'Please re-type your new password\'',8);<%End If%>">
<%
SetCookie "LastVisit", "", strCookieFolder, 0
'if form has not been filled in then display it otherwise check the details submitted
If Request.Form <> "" Then
  If Request.form("chkRememberMe") = "1" Then
    SetCookie "UID", CleanUIDinput(Request.Form("txtUID")), strCookieFolder, Now() + iDaysToRememberUser
    SetCookie "UserName", "Placeholder", strCookieFolder, Now() + iDaysToRememberUser
    SetCookie "Password", CleanUIDinput(Request.Form("txtPwd")), strCookieFolder, Now() + iDaysToRememberUser
    SetCookie "RememberMe", "1", strCookieFolder, Now() + iDaysToRememberUser
  Else
    SetCookie "RememberMe", "", strCookieFolder, 0
    SetCookie "UID", "", strCookieFolder, 0
    SetCookie "UserName", "", strCookieFolder, 0
    SetCookie "Password", "", strCookieFolder, 0
  End If
  '=== call checklogin subroutine
  CheckLoginForm
Else
  '=== call showlogin subroutine
  ShowLoginForm
End If

'=== begin subroutine ShowLoginForm
Sub ShowLoginForm
Dim strSplashText
If bRenewPwd Then
  strSplashText = "<b><font face='Arial, Helvetica, sans-serif'>Please change your password</font></b><br>" & _
                "<font size='-1' face='Arial, Helvetica, sans-serif'>Your password has expired.<br>" & _
                "Type your old password and your new password, minimum 8 characters.<br>"
Else
  strSplashText = "<b><font face='Arial, Helvetica, sans-serif'>" & strAppTitle & ":</font></b><br>" & _
                "<font size='-1' face='Arial, Helvetica, sans-serif'>The review centre for " & _
                "DNV GL Rules and other governing documents.<br>" & _
                "Members of DNV GL's Technical Committees and other hearing bodies<br>" & _
                "may sign in to read and comment on Rule Proposal documents.</font>"
End If
%>

<!--#include file="include/topright.asp"-->
<%
' Response.Write "strVirtPathSegments(Request.ServerVariables(""SCRIPT_NAME""), 2)=" & strVirtPathSegments(Request.ServerVariables("SCRIPT_NAME"), 2) & "<br>"        ' DEVELOPMENT & DEBUG
%>
<form action="<%=Request.ServerVariables("SCRIPT_NAME")%><%If bRenewPwd Then%>?pwdRen=true<%End If%>" method="post" name="credentials" onSubmit="FDK_Validate(this.name,true,false,'Please sign in\n');return document.MM_returnValue">
  <table width="100%" border="0" cellspacing="0" cellpadding="1" bgcolor="#ffffff">
    <tr valign="top"> 
      <td width="10">&nbsp;</td>
      <td> <table width="100%" border="0" cellspacing="0" cellpadding="4" bgcolor="#ffffff">
          <tr bgcolor="#12b1ee"> 
            <td valign="top"><font color="#ffffff" face="Arial, Helvetica, sans-serif"><b>Sign In</b></font></td>
            <td width=="5">&nbsp;</td>
            <td valign="top"><font color="#ffffff"><b>&nbsp;</b></font></td>
          </tr>
          <tr valign="top"> 
            <td width="200"><b><font size="-1" face="Arial, Helvetica, sans-serif">User ID</font></b><br> 
              <input name="txtUID" type="text" size="30" value="<%= CleanUIDinput(Request.Cookies("UID")) %>">
              <font face="Arial, Helvetica, sans-serif"><b><font size="-1"><%If bRenewPwd Then%>Old p<%Else%>P<%End If%>assword</font></b><br>
                <input name="txtPwd" type="password" size="30" value="<%= CleanUIDinput(Request.Cookies("Password")) %>"
                    onChange="FDK_AddAlphaNumericValidation('credentials','document.credentials.txtPwd',true,true,'\'Only letters and numbers are allowed in passwords.\'')">
                </font>
            </td>
            <td width=="5">&nbsp;</td>
            <td><%=strSplashText%>
            </td>
          </tr>
<%
If bRenewPwd Then
%>
          <tr>
            <td colspan="3">
              <font face="Arial, Helvetica, sans-serif"><b><font size="-1">New password</font></b><br>
                <input name="txtNewPwd" type="password" onChange="FDK_AddAlphaNumericValidation('credentials','document.credentials.txtNewPwd',true,true,'\'Only letters and numbers are allowed in passwords.\'')" size="30">
                </font>
            </td>
          </tr>
          <tr>
            <td colspan="3">
              <font face="Arial, Helvetica, sans-serif"><b><font size="-1">Retype new password</font></b><br>
                <input name="txtRetypeNewPwd" type="password" onChange="FDK_AddAlphaNumericValidation('credentials','document.credentials.txtRetypeNewPwd',true,true,'\'Only letters and numbers are allowed in passwords.\'')" size="30">
                </font>
            </td>
          </tr>
<%
End If
%>
          <tr>
            <td colspan="3"><font face="Arial, Helvetica, sans-serif" size="-1">
              <input type="submit" name="Submit" value="Sign In">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Forgot password?&nbsp;&nbsp;<input type="Button" name="btnRequestPassword" value="Request new password" onClick="requestNewPassword(this.form.txtUID.value)">
              </font>
            </td>
          </tr>
          <tr>
            <td colspan="3">
            <font size="-1" face="Arial, Helvetica, sans-serif"><input value="1" type="checkbox" name="chkChgPwd">&nbsp;Change password</font>
            </td>
          </tr>
          <tr>
            <td colspan="3">
            <font size="-1" face="Arial, Helvetica, sans-serif"><input value="1" type="checkbox" name="chkRememberMe"
              <% If CleanUIDinput(Request.Cookies("RememberMe")) = "1" Then
                Response.Write "CHECKED"
              Else
                Response.Write ""
              End If %>>&nbsp;Remember me</font>
            </td>
          </tr>
        </table>
      </td>
      <td width="10">&nbsp;</td>
    </tr>
  </table>
</form>
<%
  '=== end ShowLoginForm subroutine
End Sub
%>

<%
'===begin subroutine CheckLoginForm
Sub CheckLoginForm
  Dim txtUID, txtPwd, strStoredPwd, txtNewPwd, txtRetypeNewPwd, strNameUser
  txtUID = CleanUIDinput(Request.Form("txtUID"))
  txtPwd = CleanUIDinput(Request.Form("txtPwd"))
  'simple/basic protection against SQL injection use of the apostrophe
  If txtUID = "" or txtPwd = "" Then
    SetCookie "RememberMe", "", strCookieFolder, 0
    SetCookie "UID", "", strCookieFolder, 0
    SetCookie "UserName", "", strCookieFolder, 0
    SetCookie "Password", "", strCookieFolder, 0
    response.redirect Request.ServerVariables("SCRIPT_NAME")
  Else
    DIM dbConnect, strSQL, rsUserRec, strLastVisit, strpwdChangedDate
    SET dbConnect = Server.CreateObject("ADODB.Connection")

    dbConnect.Open strConnStr 
    If Err.Number <> 0 Then Call SeriousError    
    strSQL = "SELECT NameUser, pwd, pwdChangedDate, LastVisit from Users where UserID='" & txtUID & "'"
    ' response.write "strSQL=" & strSQL &"; "  ' DEVELOPMENT & DEBUG
    ' Response.End  ' DEVELOPMENT & DEBUG
    SET rsUserRec = dbconnect.Execute(strSQL)
    If Err.Number <> 0 Then Call SeriousError    
    If NOT rsUserRec.EOF Then
      strStoredPwd = Trim(rsUserRec("pwd"))
      strNameUser = Trim(rsUserRec("NameUser"))
      strLastVisit = Trim(rsUserRec("LastVisit"))
      If IsNull(strLastVisit) Then strLastVisit = ""
    End If
    rsUserRec.Close
    SET rsUserRec = Nothing
    'check to see if the form details filled in match 'username' and 'password' above
    ' Response.Write "txtPwd=" & txtPwd & "<br>"        ' DEVELOPMENT & DEBUG
    ' Response.Write "hex_md5(txtPwd)=" & hex_md5(txtPwd) & "<br>"        ' DEVELOPMENT & DEBUG
    ' Response.Write "strStoredPwd=" & strStoredPwd & "<br>"        ' DEVELOPMENT & DEBUG
    ' Response.End          ' DEVELOPMENT & DEBUG

    If hex_md5(txtPwd) = strStoredPwd Then  ' MD5 pwd encryption
      'if the correct login details are filled in then store credentials in cookies
      'and proceed

      bRenewPwd = bRenewPwd And Not bIsCollectiveUserID(txtUID)  ' Don't allow Collective userIDs to change password
      ' Password change
      If bRenewPwd Then
        txtNewPwd = CleanUIDinput(Request.Form("txtNewPwd"))
        txtRetypeNewPwd = CleanUIDinput(Request.Form("txtRetypeNewPwd"))
        ' Response.Write "txtNewPwd=" & txtNewPwd & ";<br>"  ' DEVELOPMENT & DEBUG
        ' Response.Write "txtRetypeNewPwd=" & txtRetypeNewPwd & ";<br>"  ' DEVELOPMENT & DEBUG
        ' Response.End  ' DEVELOPMENT & DEBUG
        If (txtNewPwd = "") Or (Len(txtNewPwd) < 8) Or (txtNewPwd <> txtRetypeNewPwd) Or (txtNewPwd = txtPwd) Then
          ' Incorrect passsword renewal. Take user back to logon screen
          dbConnect.Close
          Response.Redirect Request.ServerVariables("SCRIPT_NAME") & "?pwdRen=true"
        End If
        txtPwd = txtNewPwd   ' So that the password cookie get the new password
        strSQL = "UPDATE Users SET pwd = '" & hex_md5(txtPwd) & "', pwdChangedDate = GETDATE() WHERE UserID='" & txtUID & "'"
      '  Response.Write "strSQL=" & strSQL & ";<br>"    '  DEVELOPMENT & DEBUG
      '  Response.End    '  DEVELOPMENT & DEBUG
        SET rsUserRec = dbconnect.Execute(strSQL)
        If Err.Number <> 0 Then Call SeriousError

      End If

      SetCookie "UID", txtUID, strCookieFolder, 0
      SetCookie "Password", txtPwd, strCookieFolder, 0
      SetCookie "LastVisit", strLastVisit, strCookieFolder, 0
      If IsNull(strNameUser) Then 
        SetCookie "UserName", "", strCookieFolder, 0
      Else
        SetCookie "UserName", strNameUser, strCookieFolder, 0
      End If
      Session("UID") = txtUID  ' Supposed to enable Session_OnEnd to delete user's buffered document copies
      ' Update LastVisit
      strSQL = "UPDATE Users SET LastVisit=GETDATE() WHERE UserID='" & txtUID & "'"
      SET rsUserRec = dbconnect.Execute(strSQL)
      If Err.Number <> 0 Then Call SeriousError

      strSQL = "SELECT pwdChangedDate from Users where UserID='" & txtUID & "'"
      SET rsUserRec = dbconnect.Execute(strSQL)
      If Err.Number <> 0 Then Call SeriousError
      strpwdChangedDate = Trim(rsUserRec("pwdChangedDate"))
      If IsNull(strpwdChangedDate) Then strpwdChangedDate = ""
      If Not bIsCollectiveUserID(txtUID) Then   ' Don't allow nor force Collective userIDs to change password
        If strpwdChangedDate = "" Then Response.Redirect Request.ServerVariables("SCRIPT_NAME") & "?pwdRen=true"  ' Forces user to change password first time
        If DateDiff("d", CDate(strpwdChangedDate), Now(), vbMonday, vbFirstFourDays) > iPwdDaysValid Then
          Response.Redirect Request.ServerVariables("SCRIPT_NAME") & "?pwdRen=true"  ' Forces user to change password periodically
        End If
      End If

      Response.Redirect "ruleDocs.asp" 'direct to on successful login
    Else
      'if the correct details aren't filled in then show the subroutine showloginform again
      'and the statement below
      ShowLoginForm
      response.write "<div align='left'><font size='-1' color='#FF0000' face='Arial, Helvetica, sans-serif'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Sign in failure</b></font></div>"
    End If
    dbConnect.Close
    SET dbConnect = Nothing
  End If
End Sub
'=== end subroutine CheckLoginForm
%>

</body>
</html>
