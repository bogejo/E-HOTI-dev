<%
'*****************************************************************
'**   Global variables                                          **
'*****************************************************************

Dim bIsAdm                      ' True if the logged on user is member of a system administration group
Dim strAppTitle                 ' The application's title - "DNV Hearing on the Internet" (2004)
Dim strAudienceAll
Dim strAudienceDNV
Dim strCommentsRefuseDefaultMessage
Dim strDocBuffer
DIM strDocFolder
Dim strDocRepository            ' file system path to the RP Documents Repository
Dim strHearingBodies
Dim strLogonUser
Dim iDaysToRememberUser
Dim strMember
Dim strMemberEmailAddr
Dim strMemberGroup
Dim strMemberName
Dim strMemberOrg
Dim strMembersHearingBodies       ' Comma separated list of the HearingBody Names where a user or commenter is member: 
                                  ' "Nordic Committee for Safety at Sea (NSK)", "z_Document Administrators (docadm)"
                                  '   - still used in July 2005 rev (index tables)
Dim strProtocol                   ' "http" on develoment server; "https" on production server
Dim strRootFolder
Dim strCookieFolder               ' The folder where cookies' paths should be set, where SSL encryption applies
Dim strUserDomain
Dim strUserID                     ' The signed on user's RuleHearing ID
DIM strUserName                   ' The signed on user's name
Dim strUsrOrg                     ' The Organisation that the signed on user represents

'***************************************************************************************************
'*** Server dependent variables
'***************************************************************************************************
' strDocRepository: Specify file system path to the RP Documents Repository
strDocRepository = "d:\RuleHearingDocs\"  ' Production on rules.dnvgl.com
' strDocRepository = "d:\Inetpub\wwwroot\Site01\RuleHearing\ct\ruledocs\"  ' Development on osl059
strProtocol = "https"  ' on production server
' strProtocol = "http"   ' on development server
'***************************************************************************************************


strAppTitle = "DNV GL Hearing on the Internet"
strCommentsRefuseDefaultMessage = "This document is for information only. No comments required."
strRootFolder = Request("APPL_PHYSICAL_PATH")
strDocBuffer = strRootFolder & "ct\docbuf\"
strCookieFolder = strVirtPathSegments(Request.ServerVariables("SCRIPT_NAME"), 2)   '/hearings/ct/
strAudienceAll = "All readers in all committees"
strAudienceDNV = "DNV GL 'Rules and Standards' only"
iDaysToRememberUser = 190        ' Let user's web browser remember logon credentials for 1/2+ year - the typical interval between rule revisions
%>