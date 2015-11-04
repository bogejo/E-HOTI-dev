<!--#include file="globals.inc"-->
<%
'********************************************************************
'**   File:      Functions.inc                                     **
'**   Project:   Rule Hearing                                      **
'**   Changed:                                                     **
'**   Created:   2002.03.04 (BVS, CGEY for DNV)                    **
'**   Revision:  July 2005, Bo Johanson                            **
'**              Adapted to indexing tables for HB Membership      **
'**   Descr.:    Global functions                                  **
'**   Comments:  To use these functions, include                   **
'**              the following in ASP file:                        **
'**              <!--Includes file containing functions.-->        **
'**              <!--#include file="Include/Functions.inc"-->      **
'********************************************************************
'**
'** CONTENTS:
'** Function bHasQuote
'** Function bIsCollectiveUserID
'** Function bIsInList
'** Function bIsUserDNVEmployee
'** Function CleanUIDinput
'** Function dbText
'** Function FormatDateTimeISO
'** Function getUserName
'** Sub      IdentifyMember
'** Sub      IdentifyUser
'** Function mapUID2Member
'** Function ReplaceDoubleQuotes
'** Function ReplaceQuote
'** Sub      SeriousError
'** Sub      SetCookie
'** Function strCommentBy
'** Function strGetUserID
'** Function strInvalidFnameChars
'** Function strSCTL2CSL
'** function strURLunescape
'** Function Today
'** Function strVirtPathSegments
'** Function SQLAddValueOrNull
'** 

'*****************************************************************
'** Handle SQL errors and other serious page errors             **
'*****************************************************************
Sub SeriousError()
  Dim strLogin
  If IsObject(dbConnect) Then
    dbConnect.Close
    Set dbConnect = Nothing
  End If
  strLogin = Request.ServerVariables("SCRIPT_NAME")
  strLogin = Left(strLogin, InStr(2, strLogin, "/"))
  Response.Redirect strLogin  ' Throw user back to login page
  Response.End
End Sub


'*****************************************************************
'** Clean user input to prevent SQL injection and Cross-site scripting  **
'*****************************************************************
Function CleanUIDinput(strString)
  Dim regEx, strCleanUID
  Set regEx = New RegExp
  regEx.Global = true ' replace ALL matching substrings
  regEx.Pattern = "[^0-9a-zA-Z @_\-\.]"  ' find anything that isn't a '@', '-', '_', a numerical character, a lowercase or uppercase alphabetic character
  strCleanUID = regEx.Replace(strString, "")   ' replace our matches with nothing (''), as the matching characters will be the ones we don't want in our username
  Set regEx = Nothing
  CleanUIDinput = strCleanUID
End Function


'*****************************************************************
'** dbText: format text input before storing in DB.   **
'*****************************************************************
'## If a string is empty, make it "null". Also, if a string 
'## contains a single quote ('), convert it to "&#39;". The database (SQL Server)
'## sees a single quote as end of string, so the text must not contain
'## any single quotes at the time of submission to the database.
'********* BGJ: Extend the function to disarm other SQL injection and XSS techniques
Function dbText(text)
   Dim strTmp
'   Response.Write "dbText: text=" & text & ";<br>"   '   DEVELOPMENT & DEBUG
   strTmp = CStr(text)
   If strTmp = "" Then
      dbText = "null"
      Exit Function
   End If
   dbText = chr(39) & ReplaceQuote(strTmp) & chr(39)
End Function 


'*****************************************************************
'**   Returns todays date on ISO format yyyy-mm-dd              **
'*****************************************************************
Function Today()
   Dim dDate    
   dDate = Now
   Today = Year(dDate) & "-" & Right("0" & Month(dDate), 2) & "-" & Right("0" & Day(dDate), 2)
End Function 

'*****************************************************************
'** Retrieving the NT user-signature and return it              **
'*****************************************************************
Function strGetUserID()
   Dim strUserID
   strUserID = Request.Cookies("UID")
'   strGetUserID = strURLunescape(strUserID)
   strGetUserID = Trim(strUserID)
End Function


'*****************************************************************
' ** getUserName()
'*****************************************************************
Function getUserName()
  getUserName = strURLunescape(Trim(Request.Cookies("UserName")))
End Function

'*****************************************************************
'** User name and organisation  **
'*****************************************************************
Function strCommentBy()
  Dim strTmp
  strTmp = strUserName
  If strUsrOrg <> "" Then strTmp = strTmp & ", " & strUsrOrg
  'strCommentBy = Server.HTMLencode(strTmp)
  strCommentBy = strTmp
End Function

'*****************************************************************
'**   Re-formats a date from txt-format into ISO  yyyy-mm-dd    **
'*****************************************************************
Function FormatDateTimeISO(txtDate, bShowTime)
Dim myYear, myMonth, myDay, myMinute, mySecond, strResult
   If IsDate(txtDate) Then
      myYear = DatePart("YYYY", txtDate)
      myMonth = DatePart("m", txtDate)
      myDay = DatePart("d", txtDate)
      If Len(myMonth) = 1 Then
         myMonth = "0" & myMonth
      End If
      If Len(myDay) = 1 Then
         myDay = "0" & myDay
      End If
      strResult = myYear & "-" & myMonth & "-" & myDay
      If bShowTime Then strResult = strResult & " " & DatePart("h", txtDate) & ":" & DatePart("n", txtDate)
      FormatDateTimeISO = strResult
   End If
End Function


'*******************************************************************
'**   Detect ' - character in a string.                           **
'**   To prevent SQL injection                                    **
'*******************************************************************
Function bHasQuote(TextArea)
  If InStr(1, TextArea, "'", 1)>0 OR InStr(1, TextArea, "&#39;", 1)>0 Then
    bHasQuote = True
  Else
    bHasQuote = False
  End If
End Function

'*******************************************************************
'**   Removes the ' - character from a string. Replaces with ''.  **
'**   To prevent SQL injection                                    **
'*******************************************************************
Function ReplaceQuote(TextArea)
   Dim TempText
'   TempText = TextArea
   TempText = Replace(TextArea, "'", "''")
   TempText = Replace(TempText, "&#39;", "''")
   ReplaceQuote = TempText
End Function

'*********************************************************************
'**   Removes the " - character from a string. Replaces with ''''.  **
'*********************************************************************
Function ReplaceDoubleQuotes(TextArea)
   DIM TempText
'   TempText = TextArea
   TempText = Replace(TextArea, chr(34), "''''")
   ReplaceDoubleQuotes = TempText
End Function



'*********************************************************************
'**   Maps logged on userID to member name                          **
'*********************************************************************
' Should look up in user database. For now, just hardcoded demo values

Function mapUID2Member()
  If strGetUserID() = "" Then
    mapUID2Member = false
  Else
    strMember = getUserName()
    strMemberGroup = "The Committee"
    mapUID2Member = true
  End If
End Function


Sub IdentifyUser()
   '**************************
   '**   Look up user name **
   '**************************
   DIM rs, dbConnect, strAbbrev, recHBadmIDs, recHBadmID
   SET dbConnect = Server.CreateObject("ADODB.Connection")
   dbConnect.Open strConnStr
   If Err.Number <> 0 Then Call SeriousError
   DIM strSQLstatement, rsGetUserInfo
   ' Session("Admin")="NO"
   ' Session("Read/Write")="NO"
   strUserID = ReplaceDoubleQuotes(ReplaceQuote(strGetUserID()))
   If strUserID = "" Then Call SeriousError  ' Take user to logon screen upon attempts to bypass the logon procedure
   strUserName = Server.HTMLencode(getUserName())
   strSQLstatement = "SELECT Organisation FROM Users WHERE UserID = " & chr(39) & strUserID & chr(39)
   SET rs = dbconnect.Execute(strSQLstatement)
   IF Err.Number = 0 and Not rs.EOF THEN  ' Take user to logon screen on DB errors and upon attempts to bypass the logon procedure
      strUsrOrg = Trim(rs("Organisation"))
      If IsNull(strUsrOrg) Then strUsrOrg = ""
      ' strMembershipIDlist_User = rs("HearingBodiesMembership")
      ' strMembershipIDlist_User = Replace(strMembershipIDlist_User, ";", ",")  ' substitute "," for ";"
      ' strMembershipIDlist_User = Left(strMembershipIDlist_User, Len(strMembershipIDlist_User)-1)  ' strip trailing ","
      ' arrMembershipIDlist_User = split(strMembershipIDlist_User, ",")
      ' strSQLstatement = "SELECT ID, NameHB, Abbrev FROM HearingBodies WHERE ID IN (" & strMembershipIDlist_User & ") Order by NameHB"
      '## Using indexing tables instead of the ;-separated list for HB membership
      strSQLstatement = _
                        "SELECT NameHB, Abbrev " &_
                        "FROM   dbo.HearingBodies " &_
                        "WHERE  (ID IN " &_
                        "         (SELECT HearingBodyID " &_
                        "          FROM  HB_Membership " &_
                        "          WHERE      UserID = '" & strUserID & "')) " &_
                        "ORDER BY NameHB"


      SET rs = dbconnect.Execute(strSQLstatement)
      strHearingBodies = ""
      While Not rs.EOF
        strHearingBodies = strHearingBodies & Trim(rs("NameHB"))
        strAbbrev = Trim(rs("Abbrev"))
        If strAbbrev <> "" Then strHearingBodies = strHearingBodies & " (" & strAbbrev & ")"
        strHearingBodies = strHearingBodies  & ", "
        rs.MoveNext
      Wend
      strHearingBodies = Left(strHearingBodies, Len(strHearingBodies) - 2)  ' chop traling "' "
   ELSE
      Call SeriousError
   END IF

' Check whether user is sysadm or docadm
   bIsAdm = False
' July 2005, with index table HB_Membership
   strSQLstatement = _
      "Select ID from HB_Membership HBM " &_
      "Where HBM.UserID = '" & strUserID  & "' AND " &_
      "HBM.HearingBodyID IN ( " &_
      "Select ID As HBId " &_
      "From HearingBodies " &_
      "Where Abbrev IN ('sysadm', 'docadm')) "
   SET rs = dbconnect.Execute(strSQLstatement)
   If Err.Number <> 0 Then Call SeriousError
   bIsAdm = Not rs.EOF
   rs.Close
   SET rs = Nothing
   If IsObject(dbConnect) Then dbConnect.Close
   SET dbConnect = Nothing
End Sub

'*********************************************************************
'**   IdentifyMember
'**   Maps a userID to member name, organisation, hearing bodies membership, etc.
'**   Sets global varables:
'**      strMemberName
'**      strMemberOrg
'**      strMembersHearingBodies  - still used in July 2005 rev (index tables)
'**      strMemberEmailAddr
'**   Adapted to indexing tables instead of ;-separated lists for Hearing Bodies / 2005-07-18
'*********************************************************************
Sub IdentifyMember(strUserID)
  Dim strSQL, strMembershipList, strAbbrev, rs
  strUserID = ReplaceDoubleQuotes(ReplaceQuote(strUserID))
' strSQL = "SELECT NameUser, Organisation, HearingBodiesMembership, eMailAddress FROM Users WHERE UserID = " & chr(39) & strUserID & chr(39) ' ;-separated list for HB-membership
  strSQL = "SELECT NameUser, Organisation, eMailAddress FROM Users WHERE UserID = " & chr(39) & strUserID & chr(39)
'  Response.Write "strSQL=" & strSQL & ";<br>"    '  DEVELOPMENT & DEBUG
  SET rs = dbconnect.Execute(strSQL)
  If Err.Number <> 0 Then Call SeriousError
  If Not rs.EOF Then
    strMemberName = Trim(rs("NameUser"))
    strMemberOrg = Trim(rs("Organisation"))
    If IsNull(strMemberOrg) Then strMemberOrg = ""
    strMemberEmailAddr = Trim(rs("eMailAddress"))
    If IsNull(strMemberEmailAddr) OR strMemberEmailAddr = "=ID" Then strMemberEmailAddr = strUserID
'    strMembershipList = Trim(rs("HearingBodiesMembership"))
'    strMembershipList = Replace(strMembershipList, ";", ",")  ' substitute "," for ";"
'    strMembershipList = Left(strMembershipList, Len(strMembershipList)-1)  ' strip trailing ","
'    strSQL = "SELECT ID, NameHB, Abbrev FROM HearingBodies WHERE ID IN (" & strMembershipList & ") Order by NameHB"   ' Pre July 2005, using User/HearingBodiesMembership column

    ' July 2005: Using the HB_Membership index
    strSQL = "SELECT ID, NameHB, Abbrev FROM HearingBodies " &_
             "WHERE (ID IN " &_
             " (SELECT HearingBodyID " &_
             "  FROM HB_Membership " &_
             "  WHERE UserID = '" & strUserID & "')) " &_
             "ORDER BY NameHB"

'    Response.Write "strSQL=" & strSQL & ";<br>"    '  DEVELOPMENT & DEBUG
    SET rs = dbconnect.Execute(strSQL)
    If Err.Number <> 0 Then Call SeriousError
    strMembersHearingBodies = ""
    While Not rs.EOF
      strMembersHearingBodies = strMembersHearingBodies & Trim(rs("NameHB"))
      strAbbrev = Trim(rs("Abbrev"))
      If strAbbrev <> "" Then strMembersHearingBodies = strMembersHearingBodies & " (" & strAbbrev & ")"
      strMembersHearingBodies = strMembersHearingBodies  & ", "
      rs.MoveNext
    Wend
    If strMembersHearingBodies <> "" Then strMembersHearingBodies = Left(strMembersHearingBodies, Len(strMembersHearingBodies) - 2)  ' chop traling "' "
  End If
End Sub

Function bIsInList(strList, strItem, strDelimiter, bIgnoreCase, bGlobal)
   Dim regEx, strPatrn
   strPatrn = "^(" & strItem & ")$|^(" & strItem & ")" & strDelimiter & "|" & strDelimiter & "(" & strItem & ")" & strDelimiter & "|" & strDelimiter & "(" & strItem & ")$"
   ' response.write "<DEBUG>strPatrn=" & strPatrn & "</DEBUG>"   ' DEVELOPMENT & DEBUG
   Set regEx = New RegExp
   regEx.Pattern = strPatrn
   regEx.IgnoreCase = bIgnoreCase
   regEx.Global = bGlobal
   bIsInList = regEx.Test(strList)   ' Execute search.
End Function

Function bIsCollectiveUserID(strUserID)
  bIsCollectiveUserID = InStr(LCase(strUserID), "member_") = 1   ' "Collective" UserIDs start with "Member_"
  bIsCollectiveUserID = bIsCollectiveUserID OR (LCase(strUserID) = "dnvgl_browser")
End Function


Function bIsUserDNVEmployee(strUserID, ByRef dbConnect)
  Dim strSQL, rs, bIsDNVemployee
  strSQL = "SELECT ID FROM HB_Membership " &_
           "WHERE (UserID = '" & strUserID & "') AND " &_
           " (HearingBodyID IN  " &_
           "   (SELECT  ID " &_
           "     FROM   HearingBodies " &_
           "     WHERE  ID = 15)) "   ' DNV GL Employees
  Set rs = dbConnect.Execute(strSQL)
  If Err.Number <> 0 Then Call SeriousError
  If rs.EOF Then bIsDNVemployee = False Else bIsDNVemployee = True
  Set rs = Nothing
  bIsUserDNVEmployee = bIsDNVemployee
End Function

Function strSCTL2CSL(strSCTlist)
'## Convert ;-terminated list to comma separated list: "7;3;13;5;" -> "7,3,13,5"
  If strSCTlist <> "" And InStr(strSCTlist,";") >0 Then
    ' Convert ;-terminated list to Comma Separated list
    strSCTlist = Replace(strSCTlist,";",",")
    If Right(strSCTlist,1) = "," Then strSCTlist = Left(strSCTlist,Len(strSCTlist)-1)
  End If
  strSCTL2CSL = strSCTlist
End Function

Function strInvalidFnameChars(strFname)
  Dim regEx, bFlag, colMatches, oMatch, strIllegalChars
  Set regEx = New RegExp
  regEx.Global = true ' replace ALL matching substrings
  regEx.Pattern = "[^0-9a-zA-Z_\-\.\ \(\)]"  ' find anything that isn't a numerical character, a lowercase or uppercase alphabetic character, _ - . () or SPACE
'  bFlag = regEx.Test(strFname)
  Set colMatches = regEx.Execute(strFname)
  strIllegalChars = ""
  For Each oMatch in colMatches
    strIllegalChars = strIllegalChars & oMatch.Value & ", "
  Next
  If strIllegalChars <> "" Then strIllegalChars = Left(strIllegalChars, Len(strIllegalChars)-2)  ' strip trailing ", "
  Set colMatches  = Nothing
  Set regEx = Nothing
  strInvalidFnameChars = strIllegalChars
End Function

Function strVirtPathSegments(strVirtPath, iLevels)
  ' Returns the iLevels leftmost portions of the provided virtual path, i.e., the 'web' part, 
  ' e.g., '/rulehearing/'    from strVirtPathSegments('/rulehearing/ct/ruleDocs.asp', 1),
  '   and '/rulehearing/ct/' from strVirtPathSegments('/rulehearing/ct/ruleDocs.asp', 2)       
  Dim arrPathParts, strPathSegs
  arrPathParts = Split(strVirtPath, "/")
  ReDim Preserve arrPathParts(iLevels)  ' Keep just the two leftmost segments
  strPathSegs = Join(arrPathParts, "/") & "/"
  strVirtPathSegments = strPathSegs
End Function

Sub SetCookie(strName, strValue, strPath, iExpires)
  Response.Cookies(strName) = strValue
  If strPath <> "" Then Response.Cookies(strName).Path = strPath
  If iExpires <> 0 Then Response.Cookies(strName).expires = iExpires
End Sub

Function SQLAddValueOrNull(strValue, bAddTail)
  Dim strRet
  strRet = ""
  If strValue = "" Then
    strRet = "NULL"
  Else
    strRet = dbText(strValue)
  End If
  If bAddTail Then strRet = strRet & ", "
  SQLAddValueOrNull = strRet
End Function

Sub CloseAndDiscardObject(objObject)
  If IsObject(objObject) Then
    objObject.Close
    Set objObject = Nothing
  End If
End Sub

Function strFlagSorted(strField)
' strParamOrderRPby
' &dArr; &nabla;
  strField = LCase(strField)
  If strParamOrderRPby = "" And strField = "rpno" Then
    strFlagSorted = "  &nabla;"
  Else
    If strField = strParamOrderRPby Then
      strFlagSorted = "  &nabla;"
    Else
      strFlagSorted = ""
    End If
  End If
End Function


'************************************
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
<%  Response.End
  End If
End Function

%>

<%
'-------------------------------------------------------------------
' Functions for ASP / VBS to sanitize the client input.
' By Patrice Neff namics ag, 2002
'
' Protect against cross-site scripting and SQL
'
' HELP: What functions shall I use?
' - If you have some data that you want to display
' use cHtml. It escapes any HTML input so
' against cross-site scripting.
' - If you're using the data for database queries,
' functions:
' - nInt, nDbl, nLng
' This functions check if the input data
' convert it. If the data is invalid,
' - nStr
' Replaces the single quotes in strings.
' not possible.
' Placed in the GPL license.
'-------------------------------------------------------------------
Function cHtml(value)
  s = str
  s = Replace(s, "&", "&amp;")
  s = Replace(s, ">", "&qt;")
  s = Replace(s, "<", "&lt;")
  s = Replace(s, """", "&quot;")
  cHtml = s
End Function

Function nInt(value)
  If IsNumeric(value) Then
    nInt = CInt(value)
  Else
    nInt = 0
  End If
End Function

Function nDbl(value)
  If IsNumeric(value) Then
    nDbl = CDbl(value)
  Else
    nDbl = 0.0
  End If
End Function

Function nLng(value)
  If IsNumeric(value) Then
    nLng = CLng(value)
  Else
    nLng = 0
  End If
End Function

Function nStr(value)
  Dim s: s = value
  s = Replace(s, "'", "''")
  s = Replace(s, "‘", "''")
  nStr = s
End Function

' Constants for limitStr
Const alphabet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const alphnum = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
Function limitStr(value, limitset)
  Dim s: s = ""
  Dim i
  For i = 1 To Len(value)
    If InStr(1, limitset, Mid(value, i, 1))>0 Then
      s = s & Mid(value, i, 1)
    End If
  Next
  limitStr = s
End Function

' Currently does nothing. Provided for the checker tool and might be
' required in the future
Function mailVal(value)
  mailVal = value
End Function

' Required for cases where the string doesn't have to be sanitized
' e.g.: If Request.QueryString<>"" Then
Function ignVal(value)
  ignVal = value
End Function
%>

<SCRIPT runat="server" TYPE="text/javascript" LANGUAGE="javascript" >
function strURLunescape(strString) {
  return unescape(strString)
  }
</SCRIPT>
