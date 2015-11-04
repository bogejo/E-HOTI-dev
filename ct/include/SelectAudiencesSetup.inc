<%
Dim strSqlHearingsBodies, rsHearingsBodiesByName, strHB
SET dbConnect = Server.CreateObject("ADODB.Connection")
dbConnect.Open strConnStr
If Err.Number <> 0 Then Call SeriousError
strSqlHearingsBodies = "SELECT ID, NameHB, Abbrev FROM HearingBodies WHERE IsAdministratorGroup = 0 ORDER BY NameHB ASC"
SET rsHearingsBodiesByName = Server.CreateObject("ADODB.Recordset")
rsHearingsBodiesByName.Open strSqlHearingsBodies, dbConnect, 3         ' Cursor adOpenStatic = 3; facilitates RecordCount and MovePrevious/Next
rsHearingsBodiesByName.MoveFirst

If Err.Number <> 0 Then Call SeriousError
%>

<SCRIPT TYPE="text/javascript" LANGUAGE="javascript">
// Global variables
var nodeAudience;  // A (visible) document node to be updated. Declared here, but must be given a value in the module that INCLUDES fnUpdateChosenAudiences
var inpChosenAudiences;  // A (hidden) INPUT field to hold the list of chosen audiences. Declared here, but must be given a value in the module that INCLUDES fnUpdateChosenAudiences 
var arrChosenAudiences = new Array();  // Initially without values; gets value by call to fnUpdateChosenAudiences
var arrAudienceNames = new Array();  // A list of the names for defined Hearing Bodies, initialized "onLoad" from the DB

// Initialize some global variables
arrAudienceNames["All"] =  "<%=strAudienceAll%>";
arrAudienceNames["DNV"] =  "<%=strAudienceDNV%>";
<% While Not rsHearingsBodiesByName.EOF
  strHB = rsHearingsBodiesByName("Abbrev")
  If IsNull(strHB) Then
    strHB = ""
  Else
    strHB = " (" & Trim(strHB) & ")"
  End If
  strHB = Trim(rsHearingsBodiesByName("NameHB")) & strHB
%>arrAudienceNames[<%=rsHearingsBodiesByName("ID")%>] = "<%=strHB%>";
<%
  rsHearingsBodiesByName.MoveNext
Wend
%>


function fnUpdateChosenAudiences(strChosenAuds) {
  var strChosenAudiences = "";  // A ;-terminated list of chosen Hearing Bodies, e.g., "3;11;13;". 'All" -> empty list - "".
  var strAudiences = "";  // A ,-separated text string with the full names of Hearing Bodies
  if (strChosenAuds == "2;") {  // Hardcoded: 'docadm'
    strChosenAuds = "DNV";
    }
  if (strChosenAuds == "") {
    strAudiences = arrAudienceNames["All"];  // Default value, when strChosenAuds is empty
    }
  else {
    arrChosenAudiences = strChosenAuds.split(';');
    if (arrChosenAudiences[arrChosenAudiences.length-1] == "") {
      arrChosenAudiences.pop();
      }
    for (var i=0; i < arrChosenAudiences.length; i++) {
      var strThisChosenAud = arrChosenAudiences[i];
      strAudiences += arrAudienceNames[strThisChosenAud] + ", ";
      if (strThisChosenAud == "All") {
        strChosenAudiences = "";
        }
      else if (strThisChosenAud == "DNV") {
        strChosenAudiences = "2;";        // ID = 2 for "Document Administrators"
        }
      else {
        strChosenAudiences += strThisChosenAud + ";";
        }
      }
    strAudiences= strAudiences.slice(0, strAudiences.length - 2);  // Chop the trailing ", "
    }
// nodeAudience is globally declared above, but must be given a value in the module that INCLUDES fnUpdateChosenAudiences
  nodeAudience.nodeValue = strAudiences;
  inpChosenAudiences.value = strChosenAudiences;
  // alert('inpChosenAudiences.value="' + inpChosenAudiences.value + '"');   // DEBUG & DEVELOPMENT
  }

function openAudiences(theEvent) {
  // locate the file SelectAudiences.asp, it's at /<web>/ct/
  var strRootPathToPage = window.location.pathname;
  strRootPathToPage = strRootPathToPage.match(/^\/[^\/]+\/[^\/]+\//);  // Dig out the two first segments of the path, yielding /<web>/ct/
  strRootPathToPage += 'SelectAudiences.asp';  // Path to the file, from web root
  var iWinWidth = 500;
  var iWinHeight = 500;
  var strChosen = arrChosenAudiences.join(';');
  if (strChosen) {strChosen = "?PreChosen=" + strChosen}
  var eventClick = theEvent ? theEvent : window.event;
  iWinX = eventClick.screenX;
  iWinY = eventClick.screenY;
  iAudiencesX = iWinX - 90;
  iAudiencesY = iWinY - iWinHeight;
  var strWH = 'width=' + iWinWidth + ',innerWidth=' + iWinWidth + ',height=' + iWinHeight + ',innerHeight=' + iWinHeight;
  strAudiencesProps = 'left=' + iAudiencesX + ',top=' + iAudiencesY + ',menubar=no,status=no,titlebar=no,toolbar=no,scrollbars=yes,' + strWH + ',resizable=yes';
  window.open(strRootPathToPage + strChosen,'wndSelectAudiences',strAudiencesProps);
  }

</SCRIPT>
