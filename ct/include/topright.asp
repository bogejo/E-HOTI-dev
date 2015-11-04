<%
Dim strWeb
strWeb = strVirtPathSegments(Request.ServerVariables("SCRIPT_NAME"),1)
%>
<style type="text/css">
A.MenuLink {
  color: blue
  }
</style>
<!-- Top right page header -->
<table width="100%" cellspacing="0" cellpadding="0">
  <tr>
<!--
   <td width="60">
      <a href="http://www.dnvgl.com/" target="_top"><img src="<%=strWeb%>ct/images/DNV_logo.gif" align="right" vspace="10" border="0" hspace="0" alt="DNV GL Home"></a>
   </td>
-->
   <td width="270">
      <a href="http://www.dnvgl.com/" target="_top"><img src="<%=strWeb%>ct/images/DNVGL_logo_large.png" align="left" vspace="10" border="0" hspace="0" alt="DNV GL Home"></a>
   </td>
   <td valign="top" align="right">
      <font face="arial" size="-1"><br>
      <strong><font size="+2"><%=strAppTitle%></font></strong>&nbsp;&nbsp<br>
      Published by&nbsp;<a class="MenuLink" href="mailto:rules@dnvgl.com?Subject=<%=strAppTitle%>&body=Comments to <%=strAppTitle%>">DNV GL 'Rules and Standards'</a>&nbsp;&nbsp;</font>
      <%If bIsAdm Then%>
         <br>
         <font face="arial" size="-1">
         <img border="0" src="<%=strWeb%>ct/images/greenArrow.gif"><a class="MenuLink" href="<%=strWeb%>ct/AdminMenu.asp" target="_top">Admin menu</a>&nbsp;&nbsp;
         <img border="0" src="<%=strWeb%>ct/images/greenArrow.gif"><a class="MenuLink" href="<%=strWeb%>ct/ruledocs.asp" target="_top">Document List</a>&nbsp;&nbsp;
         </font>
      <%End If%>
   </td>
  </tr>
</table>
