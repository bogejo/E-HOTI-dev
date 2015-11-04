<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<%
'## Query parameters: none
%>
<%
On Error Resume Next
' On Error Goto 0   ' DEVELOPMENT & DEBUG
IdentifyUser
If Not bIsAdm Then
  Call SeriousError
End If
%>

<%
DIM dbConnect, strFormStyle
Dim strSQL, rsHearingBodies, strHBname

strFormStyle = "border-style: solid; border-color:#D0D0D0;"
SET dbConnect = Server.CreateObject("ADODB.Connection")
If Err.Number <> 0 Then Call SeriousError
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError

'## Get list of Hearing Bodies
strSQL = "SELECT ID, Abbrev, NameHB FROM HearingBodies ORDER BY NameHB"         
Set rsHearingBodies = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<head>
<title><%=strAppTitle%> - Add Hearing Body</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript">
<!--
function fnSubmit(frmForm) {
//  FDK_Validate('frmAddHB',true,false,'Missing or invalid form data\n\n');   // DEVELOPMENT & DEBUG - Don't autosubmit
  FDK_Validate('frmAddHB',true,true,'Missing or invalid form data\n\n');  // 3rd param is "AutoSubmit"
  return document.MM_returnValue;   // DEVELOPMENT & DEBUG
  }

function initValidateArray() {
  FDK_AddAlphaNumericValidation('frmAddHB','document.frmAddHB.txtHearingBody',true,true,'\'Please enter the name of the Hearing Body\'');
  FDK_AddAlphaNumericValidation('frmAddHB','document.frmAddHB.txtHBAbbrev',false,true,'"Please enter the Hearing Body\'s abbreviation"');
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

//-->
</SCRIPT>
</head>

<body style="max-width: 1000px" bgcolor="#FFFFFF" topmargin="0" leftmargin="0" onLoad="initValidateArray()">
<!--#include file="../include/topright.asp"-->

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">Add Hearing Body</font></strong></td>
  </tr>
</table>
<form action="AddHearingBodyAction.asp" method="post" name="frmAddHB">
    <table border="0" style="font-family: Arial; font-size: 10pt">
    <tr>
      <td align="right"><font color="#FF0000" size="-1"><b>Hearing Body</b></font></td><td><input name="txtHearingBody" type="text" size="60"></td>
    </tr>
    <tr>
      <td align="right"><font size="-1"><b>Abbreviation</b></font></td><td><input name="txtHBAbbrev" type="text" size="60"></td>
    </tr>
    <tr>
      <td align="right"><font size="-1"><b>Moderator Email address(es)</b></font></td>
      <td><input name="txtModeratorsEmail" type="text" size="60">&nbsp;&nbsp;Separate with &#59;</td>
    </tr>
     <tr><td colspan="2" align="center"><input name="btnAddHB" type="button" value="Add Hearing Body" onClick="fnSubmit(this.form);"></td></tr>
    </table>
</form>
<p>&nbsp;</p>

<table border="0" style="font-family: Arial; font-size: 10pt">
  <tr>
    <td>
      <font size="-1"><B>Existing&nbsp;Hearing&nbsp;Bodies</B></font>
    </td>
  </tr>

<% 
While Not rsHearingBodies.EOF
  strHBname = Trim(rsHearingBodies("NameHB"))
  If Trim(rsHearingBodies("Abbrev")) <> "" Then strHBname = strHBname + " (" + Trim(rsHearingBodies("Abbrev")) + ")"
%>
  <tr>
    <td><%=strHBname%></td>
  </tr>
<%
rsHearingBodies.MoveNext
Wend
%>
      </td>
    </tr>
</table>
<%
   '***********************
   '**   Close connection.   **
   '***********************

   Set rsHearingBodies = Nothing
   dbConnect.Close
   SET dbConnect = Nothing
%>
</body>
</HTML>
