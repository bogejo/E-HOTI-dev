<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%  '## This script lists all the rule hearing documents available in the database. 
  '## An index is created (table), containing RPId with a link for viewing and 
  '## commenting on the document. %>
<!--#INCLUDE FILE="../include/dbConnection.asp"-->
<!--#include file="../include/Functions.inc"-->
<%
On Error Resume Next
' On Error Goto 0           ' DEVELOPMENT & DEBUG
IdentifyUser
If Not bIsAdm Then Call SeriousError

Response.Clear
Response.Buffer = False  ' Allows output of "huge" lists

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%=strAppTitle%> - Update, archive or remove document</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<style type="text/css">
  a.noformat {
    text-decoration:none;
    color: #FFFFFF;
    }
</style>
<script type="text/javascript" language="javascript" src="../include/jquery-1.6.1.min.js"></script>

<script>

var ActiveCheckClass = ".chkActive";
var ArchivedCheckClass = ".chkArchived";

var bArrAnyCheckedState = new Array();
bArrAnyCheckedState[ActiveCheckClass] = false;
bArrAnyCheckedState[ArchivedCheckClass] = false;

var bArrAllCheckedState = new Array();
bArrAllCheckedState[ActiveCheckClass] = false;
bArrAllCheckedState[ArchivedCheckClass] = false;

$(document).ready(function(){

  $("#chkCheckAllActive").attr('checked', false);
  $("#chkCheckAllArchived").attr('checked', false);

  // Configure the checkboxes in the headings for Active and Archived, respectively.
  // Checking/unchecking a heading checkbox will check/uncheck all subordinate checkboxes for individual documents
  $("#chkCheckAllActive").change(function() {
    checkAll(this, ActiveCheckClass);
    });
  $("#chkCheckAllArchived").change(function() {
    checkAll(this, ArchivedCheckClass);
    });

  // Configure checkboxes for individual documents, under resp. heading - Acrive, Archived
  // When all subordinate checkboxes are checked/unchecked, the heading checkbox will become checked/unchecked
  $(ActiveCheckClass).change(function() {
    toggleHeadingCheck(this, ActiveCheckClass, "#chkCheckAllActive");
    });
  $(ArchivedCheckClass).change(function() {
    toggleHeadingCheck(this, ArchivedCheckClass, "#chkCheckAllArchived");
    });

  });  // End of $(document).ready(function())


// Implements mass check/uncheck of subordinate checkboxes, controlled by a heading checkbox
function checkAll(headingCheck, checkboxGroup) {
  if (headingCheck.checked) {
    $(checkboxGroup).attr('checked', true);
    }
  else  {
    $(checkboxGroup).attr('checked', false);
    }  
  }

// Implements toggling of a heading checkbox according to status of subordinate individual checkboxes
function toggleHeadingCheck(checkbox, checkboxGroup, checkboxHeading) {
  // alert('bArrAllCheckedState[' + checkboxGroup + ']='+ bArrAllCheckedState[checkboxGroup]);           // DEVELOPMENT & DEBUG
  bArrAllCheckedState[checkboxGroup] = true;
  $(checkboxGroup).each( function() {
    bArrAllCheckedState[checkboxGroup] = bArrAllCheckedState[checkboxGroup] && this.checked;
  });
  // alert('bArrAllCheckedState[' + checkboxGroup + ']='+ bArrAllCheckedState[checkboxGroup]);           // DEVELOPMENT & DEBUG
  if (bArrAllCheckedState[checkboxGroup]) {
    $(checkboxHeading).attr('checked', true);
    }
  else {
    $(checkboxHeading).attr('checked', false);
    }
  }

function checkboxAction(button, checkboxGroup) {
  // alert(button.value + ": " + (anyChecked(checkboxGroup)? "At least one check" : "No check"));        // DEVELOPMENT & DEBUG
  if (anyChecked(checkboxGroup)) {
    switch (button.name) {
      case "btnArchiveSelectedActive": 
      case "btnUnArchiveSelectedArchived": 
        // alert(button.name + " was clicked");        // DEVELOPMENT & DEBUG
        button.form.action = "ArchiveRP.asp";
        break;
      case "btnRemoveSelectedActive": 
      case "btnRemoveSelectedArchived": 
        button.form.action = "RemoveDocAction.asp?type=rp";
        break;
      default:
        return false;
      }
    button.form.submit();
    }
  else {
    return false;
    }
  }

function anyChecked(checkboxGroup) {
  bArrAnyCheckedState[checkboxGroup] = false;
  $(checkboxGroup).each( function() {
    bArrAnyCheckedState[checkboxGroup] = bArrAnyCheckedState[checkboxGroup] || this.checked;
    });
  return(bArrAnyCheckedState[checkboxGroup]);
  }

</script>


</head>

<body bgcolor="#FFFFFF" style="max-width: 1000px">
<!--#include file="../include/topright.asp"-->

<div style="margin-left: 12px">
<font face="arial" size="-1"><a href="#archived">Go to archive</a></font>
</div>

<%  
DIM dbConnect, RuleRS, strSQL, fFirstPass, color, strParamOrderRPby, strOrderRPby, strThisPageURL

strThisPageURL = Request.ServerVariables("SCRIPT_NAME")
strParamOrderRPby = LCase(nStr(Request.QueryString("sortBy")))
' strOrderRPby
Select Case strParamOrderRPby
  Case "designreviewdate"
    strOrderRPby = "RuleProps.DesignReviewDate, RuleProps.RPNo"
  Case "duedate"
    strOrderRPby = "RuleProps.DueDate, RuleProps.RPNo"
  Case "title"
    strOrderRPby = "RuleProps.Title, RuleProps.RPNo"
  Case Else
    strOrderRPby = "RuleProps.RPNo"
End Select

SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError


'## Get the active RPs
strSQL = "SELECT RPId, RPNo, Title, DueDate, FileName FROM RuleProps Where DateArchived Is Null ORDER BY " & strOrderRPby
Set RuleRS = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError

'## Generate the index only if there are documents stored in the database.
If Not RuleRS.EOF Then
%>

<form name="Active" action="" method="post">
  <input type="Hidden" name="checkboxGroup" value="chkActive">
  <input type="Hidden" name="RPset" value="active">
  <table border="0" cellpadding="2">
  <tr>
    <td colspan="6"><font face="Arial" size="+1"><strong>Manage hearing document</strong></font></td>
  </tr>
  <tr>
    <td bgcolor="#12b1ee" nowrap align="center"><input type="Checkbox" id="chkCheckAllActive" name="chkCheckAllActive"></td>
<!--    <td bgcolor="#12b1ee" nowrap align="center"><font face="arial" color="#FFFFFF" size="-1"><b>Remove</b></font></td> -->
    <td bgcolor="#12b1ee" nowrap align="center"><font face="arial" color="#FFFFFF" size="-1"><b>Update</b></font></td>
    <td bgcolor="#12b1ee" nowrap><font color="#ffffff" face="arial" size="-1"><b>C-R table</b></font></td>
    <td bgcolor="#12b1ee" nowrap><a href="<%=strThisPageURL%>?sortBy=RPNo" class="noformat"><font face="arial" color="#FFFFFF" size="-1"><b>Prop. No.<%=strFlagSorted("RPNo")%></b></font></a></td>
    <td bgcolor="#12b1ee"><a href="<%=strThisPageURL%>?sortBy=Title" class="noformat"><font face="arial" color="#FFFFFF" size="-1"><b>Document Title<%=strFlagSorted("Title")%></b></font></a></td>
    <td bgcolor="#12b1ee" nowrap><a href="<%=strThisPageURL%>?sortBy=DueDate" class="noformat"><font face="arial" color="#FFFFFF" size="-1"><b>Due Date<%=strFlagSorted("DueDate")%></b></font></a></td>
  </tr>
<%  fFirstPass = True

  '## Now loop through all the documents available
  Do
    If Not fFirstPass Then
      RuleRS.MoveNext
    Else
      color = ""
      fFirstPass = False
    End If
    If RuleRS.EOF Then Exit Do
    If color = "#FFFFFF" Then
      color = "#ffe1e2"
    Else
      color = "#FFFFFF"
    End If 
    
    Dim DueDate
    DueDate = FormatDateTimeISO(RuleRs("DueDate"), False)
    %>
  <tr>
    <td bgcolor="<%= color %>" nowrap valign="top" align="center"><input type="Checkbox" class="chkActive" name="chkActive" value='<%=RuleRS("RPId")%>'></td>
<!--    <td bgcolor="<%= color %>" nowrap valign="top" align="center"><font face="arial" size="-1"><a href="RemoveDocAction.asp?type=RP&ID=<%= RuleRS("RPId")%>" target="_top">remove</a></font></td> -->
    <td bgcolor="<%= color %>" nowrap valign="top" align="center"><font face="arial" size="-1"><a href="UpdateRuleDoc.asp?RPId=<%= RuleRS("RPId")%>" target="_top">update</a></font></td>
    <td bgcolor="<%= color %>" nowrap valign="top" align="center"><font face="arial" size="-1"><a href="../CommentsResponseTable.asp?RPId=<%= RuleRS("RPId")%>" target="_top">C-R table</a></font></td>
    <td bgcolor="<%= color %>" nowrap valign="top"><font face="arial" size="-1"><a href="../review.asp?fileName=<%= RuleRS("FileName")%>&amp;RPId=<%= RuleRS("RPId")%>" target="Review"><%= Server.HTMLencode(RuleRS("RPNo"))%></a></font></td>
    <td bgcolor="<%= color %>" valign="top"><font face="arial" size="-1"><%= Server.HTMLencode(RuleRS("Title"))%></font></td>
    <td bgcolor="<%= color %>" nowrap valign="top"><font face="arial" size="-1"><%=DueDate%></font></td>
  </tr>
<%  Loop %>
</table>
<input type="Button" name="btnArchiveSelectedActive" value="Archive selected" onClick="checkboxAction(this, '.chkActive');">
<input type="Button" name="btnRemoveSelectedActive" value="Remove selected" onClick="checkboxAction(this, '.chkActive');">
</form>

<%  Else %>
<font face="arial" color="#000000" size="-1">

<p>No active hearing documents</p></font>
<%  End If
Set RuleRS = Nothing
%> 

<p><a name="archived"></a>&nbsp;</p>

<%
'****************************************
'** Archived RPs
'****************************************
Dim rsArchivedRP
strSQL = "SELECT RPId, RPNo, Title, FileName, DateArchived, DesignReviewDate FROM RuleProps Where DateArchived Is Not Null ORDER BY " & strOrderRPby
Set rsArchivedRP = dbConnect.Execute(strSQL)
If Err.Number <> 0 Then Call SeriousError

'## Generate the list only if there are matching RPs
If Not rsArchivedRP.EOF Then %>

<a name="Archived"></a>
<form name="Archived" action="" method="post">
  <input type="Hidden" name="checkboxGroup" value="chkArchived">
  <input type="Hidden" name="RPset" value="archived">
  <table border="0" cellpadding="2">
  <tr>
    <td colspan="8"><font face="Arial" size="+1" color="#333333"><strong>Manage Archive</strong></font></td>
  </tr>
  <tr style="font-family: Arial;">
    <td align="center" bgcolor="#12b1ee"><input type="Checkbox" id="chkCheckAllArchived" name="chkCheckAllArchived"></td>
<!--    <td align="center" bgcolor="#12b1ee"><font color="#ffffff" size="-1"><b>Remove</b></font></td> -->
    <td align="center" bgcolor="#12b1ee"><font color="#ffffff" size="-1"><b>Update</b></font></td>
    <td align="center" bgcolor="#12b1ee"><font color="#ffffff" size="-1"><b>C-R table</b></font></td>
    <td bgcolor="#12b1ee"><a href="<%=strThisPageURL%>?sortBy=RPNo#Archived" class="noformat"><font color="#ffffff" size="-1"><b>Prop. No.<%=strFlagSorted("RPNo")%></b></a></td>
    <td bgcolor="#12b1ee"><a href="<%=strThisPageURL%>?sortBy=Title#Archived" class="noformat"><font color="#ffffff" size="-1"><b>Document Title<%=strFlagSorted("Title")%></b></font></a></td>
    <td align="center" bgcolor="#12b1ee"><a href="<%=strThisPageURL%>?sortBy=DesignReviewDate#Archived" class="noformat"><font color="#ffffff" size="-1"><b>Design review<%=strFlagSorted("DesignReviewDate")%></b></font></a></td>
    <td align="center" bgcolor="#12b1ee"><font color="#ffffff" size="-1"><b>Annexes</b></font></td>
    <td align="center" bgcolor="#12b1ee"><font color="#ffffff" size="-1"><b>Archived</b></font></td>
  </tr>

<% 
  fFirstPass = True

  '## Now loop through all the documents available
  Do
    If Not fFirstPass Then
      rsArchivedRP.MoveNext
    Else
      color = ""
      fFirstPass = False
    End If
    If rsArchivedRP.EOF Then Exit Do
    If color = "#FFFFFF" Then
      color = "#ffe1e2"
    Else
      color = "#FFFFFF"
    End If

    DIM iRPId, sqlAnnexes, rsAnnexes, iAnnexes
    iRPId = rsArchivedRP("RPId")    ' RPNo = Trim(rsArchivedRP("RPNo"))

    '## Get number of annexes and dates for archiving
    sqlAnnexes = _ 
      "SELECT COUNT(RPId) As 'NoAnnexes'" & _
      " FROM RPAnnex" & _
      " WHERE RPId = "  & iRPId
    ' Response.Write "<p>sqlAnnexes : " & sqlAnnexes  & "</p>" ' DEVELOPMENT & DEBUG

    Set rsAnnexes = dbConnect.Execute(sqlAnnexes)
    If Err.Number <> 0 Then Call SeriousError

    iAnnexes = ""
    If Not rsAnnexes.EOF Then iAnnexes = rsAnnexes("NoAnnexes")
    rsAnnexes.Close
    SET rsAnnexes = Nothing
       
%>    
  <tr style="font-family: Arial;">
    <td bgcolor="<%= color %>" nowrap valign="top" align="center"><font size="-1"><input type="Checkbox" class="chkArchived" name="chkArchived" value='<%=rsArchivedRP("RPId")%>'></td>
<!--    <td bgcolor="<%= color %>" nowrap valign="top" align="center"><font size="-1"><a href="RemoveDocAction.asp?type=RP&ID=<%= rsArchivedRP("RPId") %>" target="_top">remove</a></font></td> -->
    <td bgcolor="<%= color %>" nowrap valign="top" align="center"><font size="-1"><a href="UpdateRuleDoc.asp?RPId=<%= rsArchivedRP("RPId") %>" target="_top">update</a></font></td>
    <td bgcolor="<%= color %>" nowrap valign="top" align="center"><font size="-1"><a href="../CommentsResponseTable.asp?RPId=<%= rsArchivedRP("RPId") %>" target="_top">C-R table</a></font></td>
    <td bgcolor="<%= color %>" nowrap valign="top"><font size="-1"><a href="../review.asp?view=archive&amp;fileName=<%= rsArchivedRP("FileName") %>&amp;RPId=<%= rsArchivedRP("RPId") %>" target="Review"><%= Server.HTMLencode(rsArchivedRP("RPNo"))%></a></font></td>
    <td bgcolor="<%= color %>" valign="top"><font size="-1"><%= Server.HTMLencode(rsArchivedRP("Title"))%></font></td>
    <td bgcolor="<%= color %>" nowrap align="center" valign="top"><font size="-1"><%=Server.HTMLEncode(FormatDateTimeISO(rsArchivedRP("DesignReviewDate"), False))%></font></td>
    <td align="center" valign="top" bgcolor="<%= color%>"><font size="-1"><%=iAnnexes%></font></td>
    <td bgcolor="<%= color %>" nowrap align="center" valign="top"><font size="-1"><%=Server.HTMLEncode(FormatDateTimeISO(rsArchivedRP("DateArchived"), False))%></font></td>
  </tr>
<%  Loop %>
</table>
<input type="Button" name="btnUnArchiveSelectedArchived" value="Un-archive selected" onClick="checkboxAction(this, '.chkArchived');">
<input type="Button" name="btnRemoveSelectedArchived" value="Remove selected" onClick="checkboxAction(this, '.chkArchived');">
</form>
<%
  End If
  Set rsArchivedRP = Nothing

%>




<%

'***********************
'**  Close connection.  **
'***********************
dbConnect.Close
SET dbConnect = Nothing

%>
</body>
</html>