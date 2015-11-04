<% Function popUp(barText, title, details, link, text, align, width, KeepForm)
      Dim item
        If KeepForm Then %>
		<form method="post" action="<%= link %>"> 
	<%	For Each item In Request.Form %>
			<input type="hidden" value="<%= Replace(Request(item),chr(34),"&#34;") %>" name="<%= item %>">
	<%	Next %>
<%	End If %>
<!-- <center><p> -->
<div style="margin-left: 50px;">
<table BORDER="1" CELLSPACING="0" COLS="1" WIDTH="<%= width %>" BGCOLOR="#f0f0f0">
  <tr>
  <td valign="top">
  <center>
    <table BORDER="0" WIDTH="<%= width %>" cellpadding="4">
      <tr>
        <td align="<%= align %>" colSpan="2" vAlign="top" bgcolor="navy">
        <font face="Arial" color="#FFFFFF" size="3"><b><%= Server.HTMLencode(barText) %>&nbsp;</b></font>
      <tr>
        <td valign="top" ALIGN="<%= align %>" COLSPAN="2"><br>
          <font FACE="Arial"><b><%= Server.HTMLencode(title) %></font></b><br><br>
          <font FACE="Arial" COLOR="#660000"><b><%= Server.HTMLencode(details) %></font></b>
        </td>
      </tr>
        <td valign="top" ALIGN="<%= align %>" COLSPAN="2">
          <br><font FACE="Arial" size="2"><%= Server.HTMLencode(text) %>&nbsp;</font>
        </td>
      </tr>
      </tr>
        <td valign="top" ALIGN="center" COLSPAN="2">
        <%  If False And KeepForm Then ' 2011-09-22: Why this clause? Disabled it, by anding False %>
            <input type="submit" name="submitButton" value="OK">
        <%  Else %>
          <% ' =link%>        <% ' DEVELOPMENT & DEBUG %>
          <br>
          <b><font FACE="Arial" size="2">
          <input type="Submit" name="btnOK" value="OK" onClick="location.href='<%= link %>';">
          &nbsp;
          <input type="Button" name="btnCancel" value="Cancel" onClick="history.go(-1)">
          <br>&nbsp;</font></b>
        <%  End If %>
        
        </td>
      </tr>
    </table></center>
  </td>
  </tr>
</table>
</div>
<!-- </center>&nbsp;  -->
<%  If KeepForm Then %>
    </form> 
<%  End If %>
<%  End Function %>

<%'  Call PopUp("This is the title", "Bar title", "login.asp", "This is the text", "left", "300") %>
