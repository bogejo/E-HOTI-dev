<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
'## Sends email with notice about one or more Hearing Documents
'## to the users targeted by membership
'## Called by DocEmailNotice_Preview.asp
'## Query parameters:
'##   txtFrom
'##   txtTo
'##   txtCc
'##   txtBcc
'##   txtSubject:   the email's "subject"
'##   txtEmailBody: the message text
'## History:
'##   version 1.0 2007-08-27 / Bo Johanson
%>

<!--#include file="../include/Functions.inc"-->
<!--#INCLUDE FILE="../include/dbConnection.asp"-->

<%
DIM oMailConfig, cdo, strEmailBody, strEmailFrom, strEmailTo, strEmailCc, strEmailBcc
Dim strEmailSubject, iRPNo, strWarning


On Error Resume Next
' On Error Goto 0     ' DEVELOPMENT & DEBUG

IdentifyUser
If Not bIsAdm Then Call SeriousError

%>
<%
strWarning = ""

strEmailFrom = Trim(Request.Form("txtFrom"))
strEmailTo   = Trim(Request.Form("txtTo"))
strEmailCc   = Trim(Request.Form("txtCc"))
strEmailBcc  = Trim(Request.Form("txtBcc"))

strEmailBody = Trim(Request.Form("txtEmailBody"))
'strEmailBody = Replace(strEmailBody, Chr(10), "<br>")
'strEmailBody = Replace(strEmailBody, Chr(13), "")
'strEmailBody = Replace(strEmailBody, " ", "&nbsp;")
'strEmailBody = "<font face='Courier New'>" & strEmailBody & "</font>"

' Response.Write "strEmailBody=" & strEmailBody & "<br>"        ' DEVELOPMENT & DEBUG
' Response.Write "strEmailBcc=" & strEmailBcc & "<br>"        ' DEVELOPMENT & DEBUG
' Response.End          ' DEVELOPMENT & DEBUG

strEmailSubject = Trim(Request.Form("txtSubject"))

Set oMailConfig = Server.CreateObject ("CDO.Configuration")
With oMailConfig
    .Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mailrelay.verit.dnv.com"
    .Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    .Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
    .Fields.Update
end With
If Err.Number <> 0 Then Call SeriousError

Set cdo = CreateObject("CDO.Message")
If Err.Number <> 0 Then Call SeriousError

Set cdo.Configuration = oMailConfig
strEmailBody = strEmailBody & Chr(10) & Chr (13) ' &_
                '  "Sent from " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME")         ' DEVELOPMENT & DEBUG
With cdo
'  .From    = "bo.johanson@dnvgl.com"        ' DEVELOPMENT & DEBUG
'  .To      = "bo.johanson@dnvgl.com"          ' DEVELOPMENT & DEBUG
  .From     = strEmailFrom
  .To       = strEmailTo
  .Cc       = strEmailCc
  .Bcc      = strEmailBcc & "; bo.johanson@dnvgl.com"
  .Subject  = strEmailSubject
'  .HTMLbody = strEmailBody
  .TextBody = strEmailBody
  .send
End With

If False Then        ' DEVELOPMENT & DEBUG
   With cdo
     Response.Write ".From=" & .From & "<br>"               ' DEVELOPMENT & DEBUG
     Response.Write ".To=" & .To & "<br>"                   ' DEVELOPMENT & DEBUG
     Response.Write ".Cc=" & .Cc & "<br>"                   ' DEVELOPMENT & DEBUG
     Response.Write ".Bcc=" & .Bcc & "<br>"                 ' DEVELOPMENT & DEBUG
     Response.Write ".Subject=" & .Subject & "<br>"         ' DEVELOPMENT & DEBUG
     Response.Write ".HTMLbody=" & .HTMLbody & "<br>"       ' DEVELOPMENT & DEBUG  
   End With
End If

'Response.Write "Err.Number=" & Err.Number & "<br>"        ' DEVELOPMENT & DEBUG

%>
<html>

<head>
<title><%=strAppTitle%> - Email sent</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link REL="STYLESHEET" HREF="../include/main.css" TYPE="text/css">
</head>
<body  style="max-width: 1000px" bgcolor="#FFFFFF">
<!--#include file="../include/topright.asp"-->

<%
' Response.Write strEmailBody         ' DEVELOPMENT & DEBUG
' Response.End        ' DEVELOPMENT & DEBUG
%>
<% If strWarning <> "" Then %>
<p><%=strWarning%></p>
<% End If %>

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#669933"><strong><font face="Arial" color="#FFFFFF">This email has been sent:</font></strong></td>
  </tr>
</table>
<table border="0" cellspacing="4">
  <tr>
    <td valign="top" align="right" style="font-family: Arial; font-size: 10pt">From:</td>
    <td valign="top" style="font-family: Arial; font-size: 10pt"><%=cdo.From%></td>
  </tr>
    <tr>
      <td valign="top" align="right" style="font-family: Arial; font-size: 10pt">To:</td>
      <td valign="top" style="font-family: Arial; font-size: 10pt"><%=cdo.To%></td>
    </tr>
    <tr>
      <td valign="top" align="right" style="font-family: Arial; font-size: 10pt">cc:</td>
      <td valign="top" style="font-family: Arial; font-size: 10pt"><%=cdo.Cc%></td>
    </tr>
    <tr>
      <td valign="top" align="right" style="font-family: Arial; font-size: 10pt">bcc:</td>
      <td valign="top" style="font-family: Arial; font-size: 10pt"><%=cdo.Bcc%></td>
    </tr>
    <tr>
      <td valign="top" align="right" style="font-family: Arial; font-size: 10pt">Subject:</td>
      <td valign="top" style="font-family: Arial; font-size: 10pt"><%=cdo.Subject%></td>
    </tr>
    <tr>
      <td valign="top" align="right" style="font-family: Arial; font-size: 10pt">Body:</td>
      <td valign="top" style="font-family: Courier New; font-size: 9pt"><pre><%=cdo.TextBody%></pre></td>
    </tr>
  </table>

</body>
</html>

<%
'***********************
'**  Close connection.  **
'***********************
Set cdo = Nothing
set oMailConfig = Nothing
%>

<%

'## Function to check if a required value from the user input form has
'## been typed in. This could be done by JavaScript in the browser, but requires
'## more work from the programmers point of view, and often encounters compability
'## problems between browsers. The method used here requires more resources from
'## the server, but this is not expected to pose an real problems, since the site
'## is not expected to receive any heavy usage.
Function CheckRequiredValue(value, fieldName)
'  Response.Write "value=" & value & "<br>"   ' DEVELOPMENT & DEBUG
  If value = "" Then %>
    <h3><font face="arial">Required field: <%= fieldName %></font></h3>
    Please click your back button, and complete the form.
<%  Set oUploadedFile = Nothing
    Set oUpload = Nothing
    Response.End
  End If
End Function

%>
