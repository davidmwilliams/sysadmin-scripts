' USERLOGONS.VBS
' Reports on who didn't log in today, for payroll or other use, based on domain controller login events
' David M. Williams
' 24-May-2006

' Search for TODO to find items you need to change for your environment

' INPUTS: 
' Date to check for logon events (optional – defaults to <TODAY>)
' =================================================================================

Option Explicit
dim gTargetDate, gDictDCs, gDictUsers, oKey, gLog

' MAIN ROUTINE
' ======================================================================================================
ParseCommandLine()

WScript.Echo "Start script: " & Date() & " " & Time()

GetDCs()
GetUsers()
IgnoreContractors()

gLog = ""
For each oKey in gDictDCs.Keys
  EventLogQuery (oKey)
Next

FormatNoShows()
SendEmail()

WScript.Echo "End script: " & Date() & " " & Time()
WScript.Quit(1)

' ------------------------------------------------------------------------------------------------------
' Get the list of AD domain controllers
Sub GetDCs()
  dim objDSE, strDN, objConnection, objCommand, objRecordset, strQuery
  Set objDSE = GetObject("LDAP://rootDSE")
  strDN = "OU=Domain Controllers," & objDSE.Get("defaultNamingContext")

  set objConnection = CreateObject("ADODB.Connection")
  set objCommand = CreateObject("ADODB.Command")
  set objRecordset = CreateObject("ADODB.Recordset")

  objConnection.Provider = "ADsDSOObject"
  objConnection.Open ("Active Directory Provider")
  objCommand.ActiveConnection = objConnection

  strQuery = "SELECT name from 'LDAP:// " & strDN & "' where objectCategory='computer'"
  objCommand.CommandText = strQuery
  set objRecordset = objCommand.Execute

  set gDictDCs = CreateObject("Scripting.Dictionary")

  while not objRecordset.EOF
    gDictDCs.Add cstr (objRecordSet("Name")), ""
    objRecordset.MoveNext
  wend
End Sub

' ------------------------------------------------------------------------------------------------------
' Get the list of AD users
Sub GetUsers()
  dim objDSE, strDN, objConnection, objCommand, objRecordset, strQuery, objUser
  Set objDSE = GetObject("LDAP://rootDSE")
  strDN = "OU=TESA Group," & objDSE.Get("defaultNamingContext")

  set objConnection = CreateObject("ADODB.Connection")
  set objCommand = CreateObject("ADODB.Command")
  set objRecordset = CreateObject("ADODB.Recordset")

  objConnection.Provider = "ADsDSOObject"
  objConnection.Open ("Active Directory Provider")
  objCommand.ActiveConnection = objConnection

  strQuery = "SELECT distinguishedname from 'LDAP:// " & strDN & "' where objectCategory='person'"
  objCommand.CommandText = strQuery
  set objRecordset = objCommand.Execute

  set gDictUsers = CreateObject("Scripting.Dictionary")

  while not objRecordset.EOF
    Set objUser = GetObject ("LDAP://" & objRecordSet("distinguishedname"))
' Ignore contacts that aren't valid AD accounts
    if objUser.SAMAccountName <> "" then
      gDictUsers.Add "tesa\" & lcase (objUser.sAMAccountName), "false"
    end if
    objRecordset.MoveNext
  wend
End Sub

' ------------------------------------------------------------------------------------------------------
' Ignore people who aren't on the payroll or who can't log in to our servers
Sub IgnoreContractors()
' eg  Ignore ("boardroom")
End Sub

Sub Ignore(strPerson)
' TODO: Replace DOMAIN with your AD domainname below
  gDictUsers.Item ("DOMAIN\" & strPerson) = "true"
End Sub

' ------------------------------------------------------------------------------------------------------
' Query the event log for a specified server
Sub EventLogQuery(strLogonServer)
  dim objWMIService, EventItems, EventItem

  WScript.Echo "Checking server " & strLogonServer & " at " & Time()

  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate,(Security)}!\\" & _
      strLogonServer & "\root\cimv2")

  Set EventItems = objWMIService.ExecQuery("SELECT * FROM Win32_NTLogEvent WHERE Logfile = 'Security' AND EventCode = '538'")

  For Each EventItem in EventItems
    If (WMIDateStrToDate(EventItem.TimeGenerated) = CDate(gTargetDate)) then
      gDictUsers.Item (lcase(EventItem.User)) = "true"
    End If

    If (WMIDateStrToDate(EventItem.TimeGenerated) < CDate(gTargetDate)) then
      Exit For
    End If
  Next
End Sub

' ------------------------------------------------------------------------------------------------------
' Tidy the output
Sub FormatNoShows()
  gLog = gLog & chr(13) & "These people did not log on, on " & gTargetDate & ":" & chr(13) & chr(13)

  dim oKey
  For each oKey in gDictUsers.Keys
    If gDictUsers(oKey) = "false" then
      gLog = gLog & GetFullName (Mid (cstr(oKey), 6)) & chr(13)
    End If
  Next
End Sub

' ------------------------------------------------------------------------------------------------------
' Send e-mail
Sub SendEmail()
  dim objEmail

  Set objEmail = CreateObject("CDO.Message")

' TODO: Adjust these lines
  objEmail.From = "from@company.com"
  objEmail.To = "to@company.com"
  objEmail.Subject = "Absentee report for " & cstr(gTargetDate)
  objEmail.Textbody = gLog
  objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
  objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "SMTPSERVERIPADDRESS"
  objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

  objEmail.Configuration.Fields.Update
  objEmail.Send
End Sub

' ------------------------------------------------------------------------------------------------------
' Look up the staff member's real name
Function GetFullName(strName)
  dim objDSE, strDN, objConnection, objCommand, objRecordset, strQuery, objUser
  Set objDSE = GetObject("LDAP://rootDSE")
' TODO: Replace TOPLEVELOUGROUP below with wherever you store your users
  strDN = "OU=TOPLEVELOUGROUP," & objDSE.Get("defaultNamingContext")

  set objConnection = CreateObject("ADODB.Connection")
  set objCommand = CreateObject("ADODB.Command")
  set objRecordset = CreateObject("ADODB.Recordset")

  objConnection.Provider = "ADsDSOObject"
  objConnection.Open ("Active Directory Provider")
  objCommand.ActiveConnection = objConnection

  strQuery = "SELECT name from 'LDAP:// " & strDN & "' where objectCategory='person' and sAMAccountName='" & strName & "'"
  objCommand.CommandText = strQuery
  set objRecordset = objCommand.Execute

  GetFullName = cstr (objRecordset("name"))
End Function

' ------------------------------------------------------------------------------------------------------
' Date conversion
Function WMIDateStrToDate(dtmDate)
  WMIDateStrToDate = CDate(Mid(dtmDate, 7, 2) & "/" & Mid(dtmDate, 5, 2) & "/" & Left(dtmDate, 4))
End Function

' ------------------------------------------------------------------------------------------------------
' Get the command line arguments
Sub ParseCommandLine()
  Dim vArgs
  set vArgs = WScript.Arguments

  if (vArgs.Count = 1) then
    if vArgs(0) = "?" then
      DisplayUsageAndQuit()
    end if
  end if

  ' check arguments and set control variables 
  if vArgs.Count > 0 then
    gTargetDate = vArgs(0)
  else
    gTargetDate = Date()
  end if
End Sub

' ------------------------------------------------------------------------------------------------------
' Command line help
Sub DisplayUsageAndQuit()
  WScript.Echo ""
  WScript.Echo "Usage:"
  WScript.Echo "cscript " & WScript.ScriptName & " [optional <date>]"
  WScript.Quit(0)
End Sub
