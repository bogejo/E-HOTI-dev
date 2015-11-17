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
Dim arrEmailBcc, rcp, i, iFirstInBatch, iBatchSize, iNoOfBccRecipients, iUpperInBatch
Dim strRecipientIndex
Dim strSelectedHBs

On Error Resume Next
' On Error Goto 0     ' DEVELOPMENT & DEBUG

iBatchSize = 180   ' The number of recipients in each sent email. Upper limit in VerIT is 500. CDO may have another upper limit, apparently around 200 (2015-05-07)
IdentifyUser
If Not bIsAdm Then Call SeriousError
Response.Clear
Response.Buffer = False  ' Allows output of "huge" lists

%>
<%
strWarning = ""

strEmailFrom = Trim(Request.Form("txtFrom"))
strEmailTo   = Trim(Request.Form("txtTo"))
strEmailCc   = Trim(Request.Form("txtCc"))
strEmailBcc  = Trim(Request.Form("txtBcc"))

arrEmailBcc = split(strEmailBcc, "; ")

strEmailBody = Trim(Request.Form("txtEmailBody"))
strSelectedHBs = Server.HTMLEncode(Trim(Request.Form("txtSelectedHBs")))
strSelectedHBs = Replace(strSelectedHBs, "&lt;br&gt;", "<br>")

'strEmailBody = Replace(strEmailBody, Chr(10), "<br>")
'strEmailBody = Replace(strEmailBody, Chr(13), "")
'strEmailBody = Replace(strEmailBody, " ", "&nbsp;")
'strEmailBody = "<font face='Courier New'>" & strEmailBody & "</font>"

' Response.Write "strEmailBody=" & strEmailBody & "<br>"        ' DEVELOPMENT & DEBUG
' Response.Write "strEmailBcc=" & strEmailBcc & "<br>"        ' DEVELOPMENT & DEBUG
'  Response.Write "LBound(arrEmailBcc)=" & LBound(arrEmailBcc) & "<br>"        ' DEVELOPMENT & DEBUG
'  Response.Write "UBound(arrEmailBcc)=" & UBound(arrEmailBcc) & "<br>"        ' DEVELOPMENT & DEBUG
'  For Each rcp In arrEmailBcc        ' DEVELOPMENT & DEBUG
'    Response.Write rcp & "<br>"        ' DEVELOPMENT & DEBUG
'  Next        ' DEVELOPMENT & DEBUG
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

iNoOfBccRecipients = UBound(arrEmailBcc) + 1
For iFirstInBatch = LBound(arrEmailBcc) To UBound(arrEmailBcc) Step iBatchSize

   Set cdo = CreateObject("CDO.Message")
   If Err.Number <> 0 Then Call SeriousError

   Set cdo.Configuration = oMailConfig
   strEmailBody = strEmailBody & Chr(10) & Chr (13) ' &_
                   '  "Sent from " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME")         ' DEVELOPMENT & DEBUG

   strEmailBcc = ""

   If UBound(arrEmailBcc) < iFirstInBatch + iBatchSize - 1 Then
     iUpperInBatch = UBound(arrEmailBcc)
   Else
     iUpperInBatch = iFirstInBatch + iBatchSize - 1
   End If

   If iUpperInBatch = UBound(arrEmailBcc) Then
'      strEmailBcc = strEmailBcc & "; Sille.Grjotheim@dnvgl.com; Helen.Moller@dnvgl.com;"  ' By request DNV-TSK-2290039
   End If

   For i = iFirstInBatch to iUpperInBatch
     ' Response.Write "arrEmailBcc(" & i & ")=" & arrEmailBcc(i) & "<br>"        ' DEVELOPMENT & DEBUG
     strEmailBcc = strEmailBcc & arrEmailBcc(i) & "; "
   Next

   strRecipientIndex = Chr(10) & Chr(13) & Chr(10) & Chr(13) & iFirstInBatch + 1 & "-" & iUpperInBatch + 1 & "(" & iNoOfBccRecipients & ")"

   strEmailBcc = Left(strEmailBcc, Len(strEmailBcc) - 2)   ' Strip the trailing "; "
   ' Response.Write "strEmailBcc=" & strEmailBcc & "<br>"        ' DEVELOPMENT & DEBUG
   ' Set cdo = Nothing        ' DEVELOPMENT & DEBUG
   ' set oMailConfig = Nothing        ' DEVELOPMENT & DEBUG
   ' Response.End          ' DEVELOPMENT & DEBUG

   With cdo
   '  .From    = "bo.johanson@dnvgl.com"        ' DEVELOPMENT & DEBUG
   '  .To      = "bo.johanson@dnvgl.com"          ' DEVELOPMENT & DEBUG
     .From     = strEmailFrom
     .To       = strEmailTo
     .Cc       = strEmailCc
     .Bcc      = strEmailBcc & "; bo.johanson@dnvgl.com"
     .Subject  = strEmailSubject
   '  .HTMLbody = strEmailBody
     .TextBody = strEmailBody & strRecipientIndex
                                                      ' .send    ' Moved further down
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

   On Error Resume Next
   ' On Error Goto 0     ' DEVELOPMENT & DEBUG
   cdo.send             ' HERE'S THE ACTION, SEND

   If Err.Number <> 0 Then
     Response.Write "Failed to send email<br>"
     Response.Write "Err.Number=" & Err.Number & "<br>"        ' DEVELOPMENT & DEBUG
     Response.Write "Err.description=" & Err.description & "<br>"        ' DEVELOPMENT & DEBUG
     Err.Clear
     Response.End        ' DEVELOPMENT & DEBUG
   End If

   ' Response.Write strEmailBody         ' DEVELOPMENT & DEBUG
   ' Response.End        ' DEVELOPMENT & DEBUG
   %>
   <% If strWarning <> "" Then %>
   <p><%=strWarning%></p>
   <% End If %>

<table border="0" width="700">
  <tr>
    <td width="100%" bgcolor="#12b1ee"><strong><font face="Arial" color="#FFFFFF">This email has been sent to recipients no. <%=iFirstInBatch + 1%> to <%=iUpperInBatch + 1%> (of <%=iNoOfBccRecipients%>):</font></strong></td>
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
  <p>&nbsp;</p>

 <%
   '***********************
   '**  Close connection.  **
   '***********************
   Set cdo = Nothing

Next  ' End of "For iFirstInBatch = LBound(arrEmailBcc) To UBound(arrEmailBcc) Step iBatchSize"

set oMailConfig = Nothing
%>
<b>Selected Hearing bodies:</b>
<%=strSelectedHBs%>
</body>
</html>

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
