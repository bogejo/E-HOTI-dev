<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="include/Functions.inc"-->
<!--#INCLUDE FILE="include/dbConnection.asp"-->
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" runat="server" src="include/md5_PAJ.js"></SCRIPT>
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" runat="server">
function PwdGenerator(minLength, maxLength) {
  var strPwd = "";
  var randomchar = "";
  var numberofdigits = Math.floor((Math.random() * (maxLength - minLength + 1)) + minLength);
  for (var count=1; count<=numberofdigits; count++) {
    var chargroup = Math.floor((Math.random() * 3) + 1);
    if (chargroup==1) {
      randomchar = Math.floor((Math.random() * 26) + 65);
    }
    if (chargroup==2) {
      randomchar = Math.floor((Math.random() * 10) + 48);
    }
    if (chargroup==3) {
      randomchar = Math.floor((Math.random() * 26) + 97);
    }
    strPwd+=String.fromCharCode(randomchar);
  }
  return strPwd;
}
</SCRIPT>

<%
On Error Resume Next
' On Error Goto 0     ' DEVELOPMENT & DEBUG

Dim dbconnect, strPwd, strPwdEncrypted, strSQL, rsUserRec
Dim objEmail, oMailConfig

SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
strUserID = CleanUIDinput(Request.QueryString("UID"))
strPwd = PwdGenerator(8, 9)
strPwdEncrypted = hex_md5(strPwd)

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<head>
<title><%=strAppTitle%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link REL="STYLESHEET" HREF="include/main.css" TYPE="text/css">
</head>
<body>

<%

strSQL = "SELECT UserID from Users Where UserID = " & dbText(strUserID)
  ' Response.Write "strSQL=" & strSQL & "<br>"    '  DEVELOPMENT & DEBUG
  ' Response.End    '  DEVELOPMENT & DEBUG
SET rsUserRec = dbConnect.Execute(strSQL)
If rsUserRec.EOF Or bIsCollectiveUserID(strUserID) Or Err.Number <> 0 Then
  %>
  <p class="title">DNV GL Hearing on the Internet</p>
  <p class="text">Sorry, failed to generate new password.<br>
     Please review your Sign-in user ID on the log on screen.<br>
     Close this pop-up window and make sure your user ID is filled in on the logon screen.<br>
     Then click "Request new password".<br>
     <img src="images/Request_new_password.png"><br><br>
     Contact <a href="mailto:rules@dnvgl.com?Subject=DNV GL Hearing on the Internet&body=Password request">DNV GL 'Rules and Standards'</a>
  <p class="text"><a href="" onClick="window.close();">Close window</a></p>
  </body>
  </html>
  <%
  If IsObject(rsUserRec) Then Set rsUserRec = Nothing
  If IsObject(dbConnect) Then Set dbConnect = Nothing
  Response.End
End If

strSQL = "UPDATE Users SET pwd = '" & strPwdEncrypted & "', pwdChangedDate = GETDATE() WHERE UserID='" & strUserID & "'"
'  Response.Write "strSQL=" & strSQL & "<br>"    '  DEVELOPMENT & DEBUG
'  Response.End    '  DEVELOPMENT & DEBUG
SET rsUserRec = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then
  %>
  <p class="title">DNV GL Hearing on the Internet</p>
  <p class="text">Sorry, failed to generate new password.<br>
     Please contact <a href="mailto:rules@dnvgl.com?Subject=DNV GL Hearing on the Internet&body=Password request">DNV GL 'Rules and Standards'</a>
  <%
  Response.End
Else

Set oMailConfig = Server.CreateObject ("CDO.Configuration")
With oMailConfig
    .Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mailrelay.verit.dnv.com"
    .Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    .Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
    .Fields.Update
end With
If Err.Number <> 0 Then Call SeriousError

Set objEmail=CreateObject("CDO.Message")
Set objEmail.Configuration = oMailConfig
objEmail.Subject = strAppTitle
objEmail.From = "rules@dnvgl.com"
objEmail.To = strUserID
objEmail.TextBody = strAppTitle & ":" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Your new password is " & strPwd & _ 
                    Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                    "Please mind upper/lower case."
objEmail.Send
set objEmail=nothing
set oMailConfig = Nothing
%>

<p class="title">DNV GL Hearing on the Internet</p>
<p class="text">A new password has been sent to <%=Server.HTMLEncode(strUserID)%></p>

<%
End If
%>
<p class="text"><a href="" onClick="window.close();">Close window</a></p>
</body>
</html>
<%
If IsObject(rsUserRec) Then Set rsUserRec = Nothing
If IsObject(dbConnect) Then Set dbConnect = Nothing
%>