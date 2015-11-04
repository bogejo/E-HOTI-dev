<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
'## Query parameters (all QueryString):
'##   UID:    UserID
'## Adapted to indexing tables for HearingBodies / 2005-07-21
%>
<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<%
On Error Resume Next
' On Error Goto 0   ' DEVELOPMENT & DEBUG
IdentifyUser
If Not bIsAdm Then
  Call SeriousError
End If
%>

<%
DIM dbConnect, strUID, strSQL, rsUser, rsHearingBodies, strPWDunchanged, strDefaultEmailString, iCnt
Dim strHBname, strEmailAddr, strHBMembership, rsHBMembership

strPWDunchanged = "<no change>"
strDefaultEmailString = "[=UserID]"

strUID = Trim(Request.QueryString("UID"))
If strUID = "" Then Call InvalidUID(strUID, -10)

SET dbConnect = Server.CreateObject("ADODB.Connection")
If Err.Number <> 0 Then Call SeriousError
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError

'## Look up userID
' strSQL = "exec UsersSelProc " & dbText(strUID)
strSQL = "Select UserID, NameUser, Organisation, eMailAddress From Users Where UserID = " & dbText(strUID)
' Response.Write "strSQL=" & strSQL & "<br>"    ' DEVELOPMENT & DEBUG
Set rsUser = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError
If rsUser.EOF Then Call InvalidUID(strUID, -20)
strEmailAddr = Trim(rsUser("eMailAddress"))
If strEmailAddr = "=ID" Then strEmailAddr = strDefaultEmailString

' Build a ;-delimited list of HearingBodies where UserID is member
strSQL = "Select HearingBodyID From HB_Membership Where UserID=" & dbText(strUID)
' Response.Write "strSQL=" & strSQL & "<br>"        ' DEVELOPMENT & DEBUG
' Response.End         ' DEVELOPMENT & DEBUG
Set rsHBMembership = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError
strHBMembership = ""
While Not rsHBMembership.EOF
  strHBMembership = strHBMembership & rsHBMembership("HearingBodyID") & ";"
  rsHBMembership.MoveNext
Wend
Set rsHBMembership = Nothing

'## Get list of all Hearing Bodies
strSQL = "SELECT ID, Abbrev, NameHB FROM HearingBodies ORDER BY NameHB"         
Set rsHearingBodies = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<head>
<title><%=strAppTitle%> - Update Member</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript">
<!--
var txtDefaultselSelectedHearingBodies = "*No Hearing Body";

function SortD(box)  {
  var temp_opts = new Array();
  var temp_text = new Object();
  var temp_value = new Object();
  for(var i=0; i<box.options.length; i++)  {
    temp_opts[i] = box.options[i];
    }
  for(var x=0; x<temp_opts.length-1; x++)  {
    for(var y=(x+1); y<temp_opts.length; y++)  {
      if(temp_opts[x].text > temp_opts[y].text)  {
        temp_text = temp_opts[x].text;
        temp_value = temp_opts[x].value;
        temp_opts[x].text = temp_opts[y].text;
        temp_opts[x].value = temp_opts[y].value;
        temp_opts[y].text = temp_text;
        temp_opts[y].value = temp_value;
        }
      }
    }
  for(var i=0; i<box.options.length; i++)  {
    box.options[i].value = temp_opts[i].value;
    box.options[i].text = temp_opts[i].text;
    }
  }

function fnPreselectHB(strItemNos) {
  //strItemNos - ';'-terminated list of preselected items, e.g., "3;4;"
  var selSelectedList = document.frmUpdateUser.selSelectedHearingBodies;
  var selAvailableList = document.frmUpdateUser.selAvailableHearingBodies;

  // Find out which HBs that should be preselected, and mark them as selected in the "available HBs" list
  for (var i = 0; i < selAvailableList.options.length; i++) {
    var rxListMatch = new RegExp("(^|;)" + selAvailableList.options[i].value + ";");
    selAvailableList.options[i].selected = rxListMatch.exec(strItemNos) == null ? false : true;
    }
  // Move the thus preselected HBs to the "selected HBs" list
  fnAddSelected(selSelectedList, selAvailableList);
  }

function fnAddSelected(selSelectedList, selAvailableList) {
  var txtDefaultText = eval('txtDefault' + selSelectedList.name);
  var selOpt;  
  for (var i = 0; i < selAvailableList.options.length; i++) {
    if (selAvailableList.options[i].selected) {
      selOpt = selAvailableList.options[i];
      if (selSelectedList.options.length == 1 ) 
//        if (selSelectedList.options[0].text == txtDefaultText ) {
        if (selSelectedList.options[0].value == "default" ) {
          selSelectedList.options[0] = null;
          }
      var no = new Option(selOpt.text,selOpt.value);
      selSelectedList.options[selSelectedList.options.length] = no;
      }
    }
  for (var i = selAvailableList.options.length; i > 0; i = i - 1 ) {          
    if (selAvailableList.options[i-1].selected)
      selAvailableList.options[i-1] = null;
    }
  SortD(selSelectedList);
  }

function fnDelSelected(selSelectedList, selAvailableList) {
  var txtDefaultText = eval('txtDefault' + selSelectedList.name);
  var selOpt;  
  for (var i = 0; i < selSelectedList.options.length; i++) {
    if (selSelectedList.options[i].selected) {
//      if (selSelectedList.options[i].text == txtDefaultText) {
      if (selSelectedList.options[i].value == "default") {
        selSelectedList.options[i].value = "default";
        selSelectedList.options[i].selected = 0;
        }
      else {
        selOpt = selSelectedList.options[i];
        var no = new Option(selOpt.text,selOpt.value);
        selAvailableList.options[selAvailableList.options.length] = no;
        }
    }
  }
  for (var i = selSelectedList.options.length; i > 0; i = i - 1 ) {
    if (selSelectedList.options[i-1].selected)
      selSelectedList.options[i-1] = null;
    }
  if (selSelectedList.options.length == 0 ) {
    selOpt = new Option(txtDefaultText,"default");
    selSelectedList.options[0] = selOpt;
    selSelectedList.options[0].selected = 0;
    }
  SortD(selAvailableList);
}

function fnSelectAll(selSelect) {
  //alert("Trace: fnSelectAll on select name=" + selSelect.name + ";");
  for (var iOption = 0; iOption < selSelect.options.length; iOption++) {
    //alert("Trace: fnSelectAll in " + selSelect.name + ": " + selSelect.options[iOption].text);
    selSelect.options[iOption].selected = true;
    }
  }

function fnSubmit(frmForm) {
  fnSelectAll(frmForm.selSelectedHearingBodies);
//  FDK_Validate('frmUpdateUser',true,false,'Missing or invalid form data\n\n');   // DEVELOPMENT & DEBUG - Don't autosubmit
  FDK_Validate('frmUpdateUser',true,true,'Missing or invalid form data\n\n');  // 3rd param is "AutoSubmit"
  return document.MM_returnValue;   // DEVELOPMENT & DEBUG
  }

function initValidateArray() {
  FDK_AddAlphaNumericValidation('frmUpdateUser','document.frmUpdateUser.txtUserName',true,true,'\'Please enter the user\\\'s name\'');
  FDK_AddToValidateArray('frmUpdateUser',eval('document.frmUpdateUser.txtPwd'),"ValidatePWD(document.frmUpdateUser.txtPwd,true,'Please enter the assigned password, min. 8 characters')",true);
  FDK_AddEmailValidation('frmUpdateUser','document.frmUpdateUser.txtEmailAddr',true,true,'\'Please enter a valid e-mail address.\\n(A valid e-mail address has an \\\'@\\\' and a \\\'.\\\')\\nOr leave blank to copy UserID as e-mail address.\'');
  AddHBselectionValidation('frmUpdateUser','document.frmUpdateUser.selSelectedHearingBodies',true,'\'Hearing Bodies is a required field. Please make a selection.\'');
  FDK_AddAlphaNumericValidation('frmUpdateUser','document.frmUpdateUser.txtOrg',false,true,'\'Please enter letters and numbers only. Special characters are not allowed.\'');
  }

function FDK_Validate(FormName, stopOnFailure, AutoSubmit, ErrorHeader)
{
 var theFormName = FormName;
 var theElementName = "";
 if (theFormName.indexOf(".")>=0)  
 {
   theElementName = theFormName.substring(theFormName.indexOf(".")+1);
   theFormName = theFormName.substring(0,theFormName.indexOf("."));
 }
 var ValidationCheck = eval("document."+theFormName+".ValidateForm");
 if (ValidationCheck)  
 {
  var theNameArray = eval(theFormName+"NameArray");
  var theValidationArray = eval(theFormName+"ValidationArray");
  var theFocusArray = eval(theFormName+"FocusArray");
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

function FDK_reformat(s)
{
    var arg;
    var sPos = 0;
    var resultString = "";

    for (var i = 1; i < FDK_reformat.arguments.length; i++) {
       arg = FDK_reformat.arguments[i];
       if (i % 2 == 1) 
           resultString += arg;
       else 
       {
           resultString += s.substring(sPos, sPos + arg);
           sPos += arg;
       }
    }
    return resultString;
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

function FDK_ValidateEmail(FormElement,Required,ErrorMsg)
{
   var strDBdefaultEmail = '=ID';   // The database's default value when userID is email address
   var msg = "";
   var msgInvalid = ErrorMsg;
   // Set to '=ID' for the default '[=UserID]'
   if  (FormElement.value == "<%=strDefaultEmailString%>") {
     FormElement.value = strDBdefaultEmail;
     }
   var val = FormElement.value;
   // Return OK for the default values - blank or '[=UserID]'
   if ('' == val || val == strDBdefaultEmail) {
     return msg;
     }

   var theLen = FDK_StripChars(" ",val).length
   if (theLen == 0)       {
     if (!Required) return "";
     else return msgInvalid;
   }

   if (val.indexOf("@",0) < 0 || val.indexOf(".")<0) 
   {
      msg = msgInvalid;
   }
   return msg;
}

function FDK_AddEmailValidation(FormName,FormElementName,Required,SetFocus,ErrorMsg)  {
  var ValString = "FDK_ValidateEmail("+FormElementName+","+Required+","+ErrorMsg+")"
  FDK_AddToValidateArray(FormName,eval(FormElementName),ValString,SetFocus)
}

//ValidatePWD(document.frmUpdateUser.txtPwd,true,'Please enter the assigned password, min. 8 characters')
function ValidatePWD(FormElement,Required,ErrorMsg) {
  var msg = "";
  var theString = FormElement.value;
  if ('<%=strPWDunchanged%>' == theString) {  // No change to the password
    return msg; 
    }
  var msgInvalid = ErrorMsg;
  if (theString.length < 8) {
    return msgInvalid;
    }
  return FDK_ValidateAlphaNum(FormElement,Required,ErrorMsg);  
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
      if (!FDK_AllInRange("0","9",theChar) && !FDK_AllInRange("A","Z",theChar.toUpperCase()) && !(theChar == " ") && !OtherValidChars(theChar.toUpperCase()))     {
        return msgInvalid;
      }
    }

    return "";
}

function OtherValidChars(cChar) {
  return (cChar == "'" || cChar == '.' || cChar == '@' || cChar == '-' || cChar == '_' || cChar == 'Ü' || cChar == 'Æ' || cChar == 'Ä' || cChar == 'Ø' || cChar == 'Å' || cChar == 'Ä' || cChar == 'Ö');
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

function FDK_AddAlphaNumericValidation(FormName,FormElementName,Required,SetFocus,ErrorMsg)  {
  var ValString = "FDK_ValidateAlphaNum("+FormElementName+","+Required+","+ErrorMsg+")"
  FDK_AddToValidateArray(FormName,eval(FormElementName),ValString,SetFocus)
}

function ValidateHBselection(FormElement,ErrorMsg)
{
  msg = "";
  fnSelectAll(FormElement);
/*  Allow default = no hearing body  /BGJ 2005-09-01
  if (FormElement.value == 'default') {
    msg = ErrorMsg;
    }
*/
  return msg;
  }

function AddHBselectionValidation(FormName,FormElementName,SetFocus,ErrorMsg)  {
  var ValString = "ValidateHBselection("+FormElementName+","+ErrorMsg+")"
  FDK_AddToValidateArray(FormName,eval(FormElementName),ValString,SetFocus)
}
//-->
</SCRIPT>
</head>

<body style="max-width: 1000px" bgcolor="#FFFFFF" topmargin="0" leftmargin="0" onLoad="initValidateArray();fnPreselectHB('<%=Trim(strHBMembership)%>')">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">Update Member</font></strong></td>
  </tr>
</table>
<form action="UpdateUserAction.asp" method="post" name="frmUpdateUser">
      <INPUT TYPE="Hidden" NAME="txtUserID" VALUE="<%=Trim(rsUser("UserID"))%>">
      <table border="0" style="font-family: Arial; font-size: 10pt">
      <tr>
        <td align="right"><font size="-1"><b>UserID</b></font></td><td><%=Server.HTMLEncode(Trim(rsUser("UserID")))%></td>
      </tr>
      <tr>
        <td align="right"><font color="#FF0000" size="-1"><b>Name</b></font></td><td><input name="txtUserName" type="text" size="40" value="<%=Server.HTMLEncode(Trim(rsUser("NameUser")))%>"></td>
      </tr>
      <tr>
        <td align="right"><font color="#FF0000" size="-1"><b>Password</b></font></td><td><input name="txtPwd" type="text" size="40" value="<%=strPWDunchanged%>"></td></tr>
      <tr>
        <td align="right"><font color="#FF0000" size="-1"><b>E-mail</b></font></td><td><input name="txtEmailAddr" type="text" size="40" value="<%=Server.HTMLEncode(strEmailAddr)%>"></td>
      </tr>

    <tr>
      <td valign="bottom" align="right" height="40">
        <font size="-1"><B>Hearing&nbsp;Bodies</B></font>
      </td>
      <td align="center" valign="bottom">
        <B>Selected</B>
      </td>
      <td valign="bottom">&nbsp;</td>
      <td align="center" valign="bottom"><B>Available</B></td>
    </tr>
    <tr>
      <td valign="top" align="right">&nbsp;
        
      </td>
      <td valign="top">
        <select MULTIPLE SIZE="5" NAME="selSelectedHearingBodies">
          <script language="JavaScript">
            document.write('<option value="default">' + txtDefaultselSelectedHearingBodies + '</option>\n');
          </script>
        </select>
      </td>

      <td valign="top">
        <INPUT TYPE="Button" NAME="btnAddHearingBodies" onClick="fnAddSelected(this.form.selSelectedHearingBodies, this.form.selAvailableHearingBodies)" VALUE="<&nbsp;Add&nbsp;&nbsp;&nbsp;"><br>
        <INPUT TYPE="Button" NAME="btnDelHearingBodies" onClick="fnDelSelected(this.form.selSelectedHearingBodies, this.form.selAvailableHearingBodies)" VALUE="&nbsp;&nbsp;&nbsp;&nbsp;Del&nbsp;>">
      </td>
      <td valign="top">
          <select MULTIPLE size="5" name="selAvailableHearingBodies">
<% 
While Not rsHearingBodies.EOF
  strHBname = Trim(rsHearingBodies("NameHB"))
  If Trim(rsHearingBodies("Abbrev")) <> "" Then strHBname = strHBname + " (" + Trim(rsHearingBodies("Abbrev")) + ")"
%>
            <option value="<%=Server.HTMLEncode(Trim(rsHearingBodies("ID")))%>"><%=Server.HTMLEncode(strHBname)%></option>
<%
rsHearingBodies.MoveNext
Wend
%>
          </select>
      </td>
    </tr>
     <tr>
       <td align="right"><font color="#000000" size="-1"><b>Organisation</b></font></td><td><input name="txtOrg" type="text" size="40" value="<%=Trim(rsUser("Organisation"))%>"></td>
     </tr>
     <tr><td colspan="2" align="center"><input name="btnUpdateUser" type="button" value="Update Member" onClick="fnSubmit(this.form);"></td></tr>
    </table>
</form>
</body>
</HTML>
<%
   '***********************
   '**   Close connection.   **
   '***********************
   Set rsUser = Nothing
   Set rsHearingBodies = Nothing
   If IsObject(dbConnect) Then dbConnect.Close
   SET dbConnect = Nothing
   Response.End
%>
<%
Sub InvalidUID (strUID, iCnt)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<HEAD>
  <title><%=strAppTitle%> - Update Member</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<body style="max-width: 1000px" bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
<!--#include file="../include/topright.asp"-->
<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">Update Member - Error</font></strong></td>
  </tr>
</table>
<P><font face="arial">Could not find User ID <%=Server.HTMLencode(strUID)%></font></P>
<P>iCnt=<%=iCnt%></P>
</body>
</HTML>
<%
Set rsUser = Nothing
If IsObject(dbConnect) Then dbConnect.Close
Set dbConnect = Nothing
Response.End
End Sub
%>