<%@ LANGUAGE="VBSCRIPT" %>
<%
Option Explicit
DIM strProtocol, iPosLastSlash, strVirtRealStart, strRedirectTo

' Redirect to authenticated/protected subfolder ct
strProtocol = "https"  ' on production server
'strProtocol = "http"   ' on develoment server
iPosLastSlash = InStrRev(Request.ServerVariables("SCRIPT_NAME"), "/")
strVirtRealStart = Left(Request.ServerVariables("SCRIPT_NAME"), iPosLastSlash) & "ct/"
strRedirectTo = strProtocol & "://" & Request.ServerVariables("SERVER_NAME") & strVirtRealStart
Response.Redirect(strRedirectTo)  ' Using Respone.Redirect here instead of Server.Transfer. The latter only accepts a relative path, and can't be used to specify protocol
%>

<% Debugging code:
If False Then
  Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">"
  Response.Write "<HTML>"
  Response.Write "<head>"
  Response.Write "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=ISO-8859-1'>"
  Response.Write "</head>"
  Response.Write "<p>"
  Response.Write "SCRIPT_NAME=" & Request.ServerVariables("SCRIPT_NAME") & "<br>"
  Response.Write "LOGON_USER=" & Request.ServerVariables("LOGON_USER") & "<br>"
  Response.Write "HTTP_HOST=" & Request.ServerVariables("HTTP_HOST") & "<br>"
  Response.Write "SERVER_NAME=" & Request.ServerVariables("SERVER_NAME") & "<br>"
  Response.Write "PATH_INFO=" & Request.ServerVariables("PATH_INFO") & "<br>"
  Response.Write "PATH_TRANSLATED=" & Request.ServerVariables("PATH_TRANSLATED") & "<br>"
  Response.Write "URL=" & Request.ServerVariables("URL") & "<br>"
  Response.Write "SERVER_PROTOCOL=" & Request.ServerVariables("SERVER_PROTOCOL") & "<br>"
  Response.Write "strRedirectTo=" & strRedirectTo & "<br>"
  Response.Write "</p>"
  Response.Write "</HTML>"
End If
%>