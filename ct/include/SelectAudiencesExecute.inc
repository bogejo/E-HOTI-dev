<%
Dim strPreChosen, strChkExclusiveGroup

strChkExclusiveGroup = "chkAudExclusive"
strPreChosen = Request("PreChosen")
If strPreChosen = "" Then strPreChosen = "All"
%>

<SCRIPT TYPE="text/javascript" LANGUAGE="javascript">
function exclusiveAud(chkItem) {
  if (chkItem.checked) {
    uncheckAllOtherAuds(chkItem);
    }
  fnDefaultToAll(chkItem.form);  // If no items are checked, check "All"
  }

function uncheckAllOtherAuds(chkItem) {
  var frmForm = chkItem.form;
  var formElement;
  for (var iFormItem = 0; iFormItem < frmForm.elements.length; iFormItem++) {
    formElement = frmForm.elements[iFormItem];
    if (formElement.type == "checkbox" && formElement.name != chkItem.name) {
      formElement.checked = false;
      }
    }
  }

function excludeExclusiveAuds(chkItem) {
  var frmForm = chkItem.form;
  if (chkItem.checked) {
    var formElement;
    var iFormItem;
    for (iFormItem = 0; iFormItem < frmForm.elements.length; iFormItem++) {
      formElement = frmForm.elements[iFormItem];
      if (formElement.type == "checkbox" && formElement.name.indexOf("<%=strChkExclusiveGroup%>") == 0) {
        formElement.checked = false;
        }
      }
    }
  else {  // If no items are checked, check "All"
    fnDefaultToAll(frmForm);
    }
  }

function fnDefaultToAll(frmForm) {  // If no items are checked, check "All"
  var formElement;
  var iFormItem;
  var bIsChecked = false;
  for (iFormItem = 0; iFormItem < frmForm.elements.length; iFormItem++) {
    formElement = frmForm.elements[iFormItem];
    if (formElement.type == "checkbox" && formElement.checked) {
        bIsChecked = true;
        break;
       }
    }
  if (!bIsChecked) {
    frmForm.<%=strChkExclusiveGroup%>All.checked = true;
    }  
  }

function fnUpdateChecked(wndWindow, frmForm) {
  var strTargetAuds = "";     // A ;-terminated list of Hearing Body IDs, e.g, "All;" or "2;3;4;"
  for (var iElem = 0; iElem < frmForm.elements.length; iElem++) {
    var formElement = frmForm.elements[iElem];
    if (formElement.type == "checkbox" && formElement.name.indexOf('chkAud') > -1 && formElement.checked) {
       strTargetAuds += formElement.value + ';';
       }
    }
  wndWindow.fnUpdateChosenAudiences(strTargetAuds);
  if (wndWindow != self) {  window.close(); } // self
  }


</SCRIPT>

<% 
Sub SelectAudiencesCleanUp()
   If IsObject(rsHearingsBodiesByName) Then
      rsHearingsBodiesByName.Close
      Set rsHearingsBodiesByName = Nothing
   End If
   If IsObject(dbConnect) Then
      dbConnect.Close
      Set dbConnect = Nothing 
   End If
End Sub

Sub ListAudiences()
   Dim iAudNo, strHB, strHBabbr
   rsHearingsBodiesByName.MoveFirst
   Do While Not rsHearingsBodiesByName.EOF
     iAudNo = rsHearingsBodiesByName("ID")
     strHB = Trim(rsHearingsBodiesByName("NameHB"))
     strHBabbr = Trim(rsHearingsBodiesByName("Abbrev"))
     If strHBabbr <> "" Then strHB = strHB & "&nbsp;(" & strHBabbr & ")"
   %>    <tr><td><input TYPE='Checkbox' onClick='excludeExclusiveAuds(this)' NAME='chkAudno<%=iAudNo%>' VALUE='<%=iAudNo%>' <% If bIsInList(strPreChosen, iAudNo, ";", true, false) Then Response.Write " checked" %>></td><td><%=strHB%></td></tr>
   <%
     rsHearingsBodiesByName.MoveNext
   Loop
End Sub
%>
