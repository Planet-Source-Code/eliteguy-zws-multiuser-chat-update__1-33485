Attribute VB_Name = "mod_WsServer"
'Copyright 2002 bryan phillips
'Known bugs in kick system, (very minor) (build 820)

Public Enum ResponseTypes
    connecting = 1
    disconnecting = 2
    message = 3
    serverstop = 4
    kick = 5
    notice = 6
    userlist = 7
    commandRequest = 8
    Auth = 9
    pm = 10
    joins = 11
    quits = 12
    pmclose = 13
    declinepm = 14
    sPage = 15
End Enum

Public Enum CommandTypes
    server_stats = 1
End Enum

Public Type nServer
    runtimeseconds As Byte
    runtimeminutes As Byte
    runtimehours As Integer
End Type

Public Type wsServer
    currentDataPort As Long
    listenport As Long
    showWlcmMsg As Boolean
    showByeMsg As Boolean
    wlcmMsg As String
    byeMsg As String
    runtime As String
    stat_usersT As Long
    stat_users As Long
    stat_bytesT As Long
    stat_bytes As Long
    stat_msgsT As Long
    stat_msgs As Long
    StartTime As String
    lastShutDown As String
    rest_canMultiUser As Boolean
    rest_canRespondCmd As Boolean
    rest_isLogin As Boolean
    rest_showGroup As Boolean
    build As String
End Type

Public Type uLogin
    auth_user As String
    auth_pass As String
    enabled As Boolean
    uid As Long
    u_type As Byte
    maxLogins As Integer
    logged As Integer
    Operator As Boolean
End Type

Public Enum inputResponse
    i_Cancel = 0
    i_ok = 1
End Enum

Public Enum UserRestV
    multiplelogins = 1
End Enum

Public Server As wsServer
Public mServer As nServer
Public log() As String
Public users() As String
Public Const Sep As String = vbCrLf & "." & vbCrLf
Public Const lSep As String = vbCrLf & "\«/" & vbCrLf
Public Const noSuchuser As Integer = 32000
Public Const noSuchUserName As String = ""
Public Const UserListSep As String = ""
Public Const UserListSep1 As String = "{«}"
Public Const ParseLoginAuth As String = " & vbcrlf"
Public Const OpChar As String = "@"
Public Const f_Width As Long = 5715
Public Const f_Height As Long = 4965
Public inpGotResponse As Boolean
Public inpResp As inputResponse
Public inpRtext As String
Public Logins() As uLogin
Public LoginFilename As String

Public Sub initWsServer()
    With Server
        .currentDataPort = 4601
        .listenport = GetSetting("ZWSserver", "Server", "Port", 4000)
    End With
End Sub

Public Sub StartServer()
    With frm_main.ws_listen
        .LocalPort = Server.listenport
        If .State <> 0 Then: .Close
        .Listen
        
    End With
    ReDim log(0)
    ReDim users(0)
    users(0) = Chr(6)
    act "[" & DateTime.Now & "] " & "server started. on port " & Server.listenport
    frm_main.mnuSS.enabled = False
    frm_main.mnuSSD.enabled = True
End Sub

Public Sub StopServer()
    For i = 0 To frm_main.ws_data.UBound
        With frm_main.ws_data(i)
            If .State = sckConnected Then
                Send i, 5, " Server is closing..."
                DoEvents
            End If
            .Close
        End With
    Next
    frm_main.ws_listen.Close
    act "[" & DateTime.Now & "] " & "server shutdown"
    Server.lastShutDown = DateTime.Now
    
    frm_main.mnuSS.enabled = True
    frm_main.mnuSSD.enabled = False
End Sub
Public Sub Send(ByVal sIndex As Long, cCode As ResponseTypes, sData As String)
frm_main.ws_data(sIndex).SendData cCode & Sep & sData
DoEvents
End Sub

Public Function OpenWsk() As Long
Dim ind As Long
For i = 0 To (frm_main.ws_data.UBound)
    With frm_main.ws_data(i)
        Select Case .State
            Case sckConnected
            Case sckClosed
                OpenWsk = i
                Exit Function
            Case sckConnecting
            Case Else
                'we can close these
                .Close
                OpenWsk = i
                Exit Function
            End Select
    End With
Next
    'if all the winsock controls are not available, open a new one
    ind = (frm_main.ws_data.UBound) + 1 'increment the control number
    Load frm_main.ws_data(ind) 'load the new winsock control
    With Server
        If .currentDataPort < 65499 Then
            .currentDataPort = .currentDataPort + 1
        Else
            .currentDataPort = 4600
        End If
            frm_main.ws_data(ind).LocalPort = .currentDataPort
            'set a random port
    End With
    OpenWsk = ind
End Function

Public Sub LogError(strDescription As String)
    log(UBound(log)) = strDescription
    ReDim Preserve log((UBound(log)) + 1)
End Sub

Public Sub ClearLog()
    ReDim log(0)
End Sub

Public Function GetcCode(sData As String) As ResponseTypes
On Error Resume Next
    GetcCode = Parse(sData, 0)
End Function

Public Function Parse(strString As String, ByVal lItemNum As Long) As String
    Dim arrItems() As String
    arrItems = Split(strString, Sep, , vbTextCompare)
    If lItemNum <= UBound(arrItems) Then
        Parse = arrItems(lItemNum)
    End If
    Erase arrItems
End Function

Public Sub Broadcast(cCode As ResponseTypes, strMessage As String)
For i = 0 To frm_main.ws_data.UBound
    With frm_main.ws_data(i)
        Select Case .State
            Case sckConnected
                .SendData cCode & Sep & strMessage
                DoEvents ' this will allow us to send to all users!!
            Case Else
        End Select
    End With
Next
End Sub

Public Sub CloseNotNeeded()
    For i = 0 To frm_main.ws_data.UBound
        With frm_main.ws_data(i)
            Select Case .State
                Case sckConnected, sckConnected, sckResolvingHost, sckHostResolved
                Case Else
                    .Close
            End Select
        End With
    Next
End Sub

Public Function FindOpenUserSlot() As Integer
    For i = 0 To UBound(users)
        If users(i) = Chr(6) Then
            FindOpenUserSlot = i
            Exit Function
        End If
    Next
    
    ReDim Preserve users(UBound(users) + 1)
        FindOpenUserSlot = (UBound(users))
End Function

Public Function FindUser(ByVal userid As Long) As Integer
    For i = 0 To UBound(users)
        If Val(Parse(users(i), 0)) = userid Then
            FindUser = i
            Exit Function
        End If
    Next
    FindUser = noSuchuser
End Function

Public Sub AddUser(ByVal uTag As String, username As String, ip As String, uindex As String, Optional Acct As String, Optional ByVal islogged As Byte)
users(FindOpenUserSlot) = uTag & Sep & username & Sep & ip & Sep & uindex & Sep & Acct & Sep & islogged
act "CON: " & DateTime.Now & " [" & username & "] " & ip & " < " & uindex
Server.stat_usersT = Server.stat_usersT + 1
Server.stat_users = Server.stat_users + 1
SendUserList
AddLoginsToList
End Sub

Public Sub RemoveUser(uTag As Long)
Dim Usrname As String
Dim Lnum As Long
Dim ln As Long
Dim islogged As Integer
Broadcast quits, Parse(users(FindUser(uTag)), 1)
act "DIS: " & DateTime.Now & " " & Parse(users(FindUser(uTag)), 1) & " [" & uTag & "]"
ln = FindUser(uTag)
Usrname = Parse(users(ln), 4)
Lnum = GetLoginNumFromName(Usrname)
If Lnum <> -255 Then
islogged = Val(Parse(users(FindUser(uTag)), 5))
If islogged = 1 Then
    Logins(Lnum).logged = Logins(Lnum).logged - 1
End If
Else
    Lnum = 0
End If
users(FindUser(uTag)) = Chr(6)
UpdateUserList
SendUserList
AddLoginsToList
End Sub

Public Sub UpdateUserList()
Dim li As ListItem
    With frm_main.UsrList
    .ListItems.Clear
    For i = 0 To UBound(users)
        If users(i) <> Chr(6) Then
            If users(i) = "" Then
                users(i) = Chr(6)
            Else
                Set li = .ListItems.Add(, , Hex(Parse(users(i), 0)), , 1)
                    li.SubItems(1) = Parse(users(i), 1)
                     li.SubItems(2) = Parse(users(i), 2)
                      li.SubItems(3) = Parse(users(i), 3)
                    If Parse(users(i), 4) <> "" Then
                         li.SubItems(4) = Parse(users(i), 4)
                    Else
                         li.SubItems(4) = "None"
                    End If
            End If
        End If
    Next
    End With
End Sub

Public Sub act(sdesc As String)
frm_main.lstActivity.AddItem sdesc
frm_main.lstActivity.ListIndex = frm_main.lstActivity.ListCount - 1
End Sub

Public Sub GetSettings()
With frm_main
    .Check1.Value = GetSetting("ZWSserver", "Settings", "SWelcomeMsg", 0)
    .Text2.Text = GetSetting("ZWSserver", "Settings", "WelcomeMsg", "")
    
        Server.wlcmMsg = .Text2.Text
        Server.showWlcmMsg = .Check1.Value
    With Server
        .stat_bytesT = GetSetting("ZWSserver", "Stats", "bytes", 0)
        .stat_msgsT = GetSetting("ZWSserver", "Stats", "msg", 0)
        .stat_usersT = GetSetting("ZWSserver", "Stats", "users", 0)
        .rest_canMultiUser = GetSetting("ZWSserver", "Restrictions", "MultiUser", False)
        .rest_canRespondCmd = GetSetting("ZWSserver", "Restrictions", "Commands", True)
        .rest_isLogin = GetSetting("ZWSserver", "Restrictions", "UserLogin", False)
        .rest_showGroup = GetSetting("ZWSserver", "Restrictions", "sGroup", False)
        .listenport = GetSetting("ZWSserver", "Server", "Port", 4500)
    End With
        .lstUserRest.Selected(0) = Server.rest_canMultiUser
        .lstUserRest.Selected(1) = Server.rest_canRespondCmd
        .lstUserRest.Selected(2) = Server.rest_isLogin
        .lstUserRest.Selected(3) = Server.rest_showGroup
        .Text4.Text = Server.listenport
End With
End Sub

Public Sub SendWelcomeMsg(ByVal wsIndex As Long)
With frm_main.ws_data(wsIndex)
If Server.showWlcmMsg = True Then
    If Server.wlcmMsg <> "" Then
        .SendData ResponseTypes.notice & Sep & vbCrLf & "ZWS server " & App.Major & "." & App.Minor & "." & App.Revision & " Welcome " & GetUserName(.Tag) & " " & vbCrLf & "     " & Server.wlcmMsg & vbCrLf & "     Runtime: " & Server.runtime & vbCrLf & vbCrLf
    End If
End If
End With
End Sub

Public Sub SendByeMsg(ByVal wsIndex As Long)
With frm_main.ws_data(wsIndex)
If Server.showByeMsg = True Then
    If Server.byeMsg <> "" Then
        .SendData ResponseTypes.notice & Sep & vbCrLf & "ZWS server " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & "     " & Server.byeMsg & vbCrLf & vbCrLf
    End If
End If
End With
End Sub

Public Function GetUserName(ByVal uid As Long) As String
For i = 0 To UBound(users)
    If Val(Parse(users(i), 0)) = uid Then
        GetUserName = Parse(users(i), 1)
        Exit Function
    End If
Next
GetUserName = noSuchUserName
End Function

Public Function GetUserGroup(ByVal userid As Long) As Long
Dim ln As String
For i = 0 To UBound(users)
    If Val(Parse(users(i), 0)) = Val(userid) Then
        ln = Trim(LCase(Parse(users(i), 4)))
            If ln <> "none" Then
                GetUserGroup = GetLoginNumFromName(ln)
            End If
        Exit Function
    End If
Next
GetUserGroup = -255
End Function

Public Function SendUserList()
Dim ul As String
Dim op As String
Dim ln As Long
'form the user list
Dim uIDz As Long
For i = 0 To UBound(users)
If users(i) <> Chr(6) Then
    uIDz = Val(Parse(users(i), 0))
    ln = GetUserGroup(uIDz)
        If ln <> -255 Then ' if ln=-255, the user account must be "none"
            If Logins(ln).Operator = True Then
                op = "@"
            Else
                op = ""
            End If
        End If
        If users(i) <> "" Then
            If Server.rest_showGroup = False Then
                ul = ul & Parse(users(i), 1) & UserListSep1 & Parse(users(i), 0) & UserListSep1 & " " & UserListSep1 & op & UserListSep
            Else
                ul = ul & Parse(users(i), 1) & UserListSep1 & Parse(users(i), 0) & UserListSep1 & _
                Logins(GetUserGroup(uIDz)).auth_user & UserListSep1 & op & UserListSep
            End If
        End If
End If
Next
'send the userlist to everyone...
For i = 0 To frm_main.ws_data.UBound
    With frm_main.ws_data(i)
        Select Case .State
            Case sckConnected
                .SendData ResponseTypes.userlist & Sep & ul
                DoEvents
            Case Else
        End Select
    End With
Next
End Function

Public Function checkUserRestrictions(nRestriction As UserRestV, sUser As String) As Boolean
Select Case nRestriction
Case UserRestV.multiplelogins
    If Server.rest_canMultiUser = True Then
        checkUserRestrictions = True
        Exit Function
    Else
        For i = 0 To UBound(users)
            If Native(Parse(users(i), 1)) = Native(sUser) Then
                checkUserRestrictions = False
                Exit Function
            End If
        Next
    End If
    checkUserRestrictions = True
    Exit Function
Case Else
End Select
End Function

Public Sub respondToCommandRequest(ByVal sIndex As Integer, ByVal comRequest As CommandTypes, Optional cmdPasses As String)
If Server.rest_canRespondCmd Then
Select Case comRequest
    Case CommandTypes.server_stats
        'send the server statistics
        frm_main.ws_data(sIndex).SendData ResponseTypes.commandRequest & _
        Sep & "Server statistics follow..." & vbCrLf & "Bytes served: " & _
        Server.stat_bytesT & vbCrLf & "Users served: " & Server.stat_usersT & _
        vbCrLf & "Messages served: " & Server.stat_msgsT & vbCrLf & _
        "Current users: " & frm_main.StatusBar1.Panels(1).Text & vbCrLf & _
        "Start date: " & Server.StartTime & vbCrLf & "Runtime: " & Server.runtime & _
        vbCrLf & "Last Shutdown: " & Server.lastShutDown & vbCrLf & _
        "Server build: " & Server.build & _
        vbCrLf & "(%) End server stats"
        
    Case Else
        frm_main.ws_data(sIndex).SendData ResponseTypes.commandRequest & Sep & "(%) Invalid remote command " & vbCrLf & "type /help for a list of common commands"
End Select
Else
    frm_main.ws_data(sIndex).SendData ResponseTypes.commandRequest & Sep & "(%) Server is ignoring remote commands..."
End If
End Sub

Public Sub KickUser(ByVal uid As Long, Optional strReason As String)

For i = 0 To UBound(users)
    If Parse(users(i), 0) = uid Then
        frm_main.ws_data(Val(Parse(users(i), 3))).SendData ResponseTypes.kick & Sep & strReason
                DoEvents
            frm_main.ws_data(Val(Parse(users(i), 3))).Close
            act "KICK: " & Parse(users(i), 1) & " (" & Parse(users(i), 0) & ") [" & strReason & "]"
            DoEvents
            CloseNotNeeded
        Exit Sub
    End If
Next

MsgBox "No Such User!"
End Sub

Public Function GetInput(strInputCaption As String, strTarget As String) As inputResponse
inpGotResponse = False
frm_main.enabled = False
frm_input.Show
frm_input.Label1.Caption = strInputCaption
Do
    DoEvents
Loop Until inpGotResponse
    strTarget = inpRtext
    GetInput = inpResp
    frm_main.enabled = True
End Function

Public Function getUserCount() As Long
Dim tmpUsrCntm As Long
For i = 0 To UBound(users)
    If users(i) <> Chr(6) Then
        tmpUsrCntm = tmpUsrCntm + 1
    End If
Next
getUserCount = tmpUsrCntm
End Function

Public Function serverState() As String
With frm_main
    Select Case .ws_listen.State
        Case sckListening
            serverState = "Online"
                Exit Function
        Case sckClosed
            serverState = "Offline"
                Exit Function
    End Select
End With
End Function

Public Sub MassKick(Optional strReason As String)
    For i = 0 To frm_main.ws_data.UBound
        With frm_main.ws_data(i)
            If .State = sckConnected Then
                Send i, kick, strReason
                DoEvents
            End If
            .Close
        End With
    Next
End Sub

Public Function loginExists(authUsername As String) As Boolean
For i = 0 To UBound(Logins)
    With Logins(i)
        If Native(.auth_user) = Native(authUsername) Then
            loginExists = True
            Exit Function
        End If
    End With
Next
loginExists = False
End Function

Public Sub WriteLogins()
Dim file
file = FreeFile
Open LoginFilename For Output As file
Print #file, """"
    For i = 0 To UBound(Logins)
        With Logins(i)
            If .uid <> 0 Then
                Print #file, .uid & Sep & _
                .auth_user & Sep & _
                .auth_pass & Sep & _
                .enabled & Sep & _
                .u_type & Sep & _
                .maxLogins & Sep & _
                .Operator & lSep
            End If
        End With
    Next
Print #file, """"
Close file
End Sub

Public Function ReadLogins(sDestArry() As String) As Boolean
On Error GoTo rderr:
Dim file
file = FreeFile
Dim strLs As String
    Open LoginFilename For Input As #file
        Input #file, strLs
    Close #file
sDestArry = Split(strLs, lSep, -1, vbTextCompare)
ReadLogins = True
Exit Function
rderr:
    ReadLogins = False
    Close #file
End Function

Public Function AddLogin(username As String, password As String, enabled As Boolean, ByVal mType As Byte, ByVal MaxLogi As Integer, lOperator As Boolean) As Boolean
Dim Lind As Long
Lind = (UBound(Logins) + 1)
If loginExists(username) = False Then
    ReDim Preserve Logins(Lind)
    With Logins(UBound(Logins))
        .auth_user = username
        .auth_pass = password
        If enabled = True Then
            .enabled = 1
        Else
            .enabled = 0
        End If
        .Operator = lOperator
        .u_type = mType
        .maxLogins = MaxLogi
        Randomize
        .uid = Int(Rnd * 10000000)
    End With
    AddLogin = True
    WriteLogins
    AddLoginsToList
    Exit Function
Else
    AddLogin = False
End If
End Function

Public Function RemoveLogin(username As String) As Boolean
For i = 0 To UBound(Logins)
    With Logins(i)
        If Native(.auth_user) = Native(username) Then
            .auth_pass = ""
            .auth_user = ""
            .uid = 0
            .enabled = False
            .maxLogins = 0
            .u_type = 3
            RemoveLogin = True
            WriteLogins
            AddLoginsToList
            Exit Function
        End If
    End With
Next
RemoveLogin = False
End Function

Public Sub ParseLogins(loginList() As String)
On Error Resume Next
ReDim Logins(UBound(loginList))
If UBound(loginList) <> 1 Then
For i = 0 To (UBound(loginList))
    If Parse(loginList(i), 0) <> "" Then
        Logins(i).uid = Val(Parse(loginList(i), 0))
        Logins(i).auth_user = Parse(loginList(i), 1)
        Logins(i).auth_pass = Parse(loginList(i), 2)
        Logins(i).enabled = Parse(loginList(i), 3)
        Logins(i).u_type = Parse(loginList(i), 4)
        Logins(i).maxLogins = Parse(loginList(i), 5)
        Logins(i).Operator = Parse(loginList(i), 6)
    End If
Next
Else
    If Parse(loginList(0), 0) <> "" Then
        Logins(i).uid = Val(Parse(loginList(0), 0))
        Logins(i).auth_user = Parse(loginList(0), 1)
        Logins(i).auth_pass = Parse(loginList(0), 2)
        Logins(i).enabled = Parse(loginList(0), 3)
        Logins(i).u_type = Parse(loginList(i), 4)
        Logins(i).maxLogins = Parse(loginList(i), 5)
        Logins(i).Operator = Parse(loginList(i), 6)
    End If
End If
End Sub

Public Sub AddLoginsToList()
Dim picnum As Integer
Dim ndesc As String
frm_main.lstLogin.ListItems.Clear
Dim li As ListItem
For i = 0 To UBound(Logins)
    With Logins(i)
        If .uid <> 0 Then
                Select Case .u_type
                    Case 1
                        picnum = 5
                        ndesc = "User"
                    Case 2
                        picnum = 6
                        ndesc = "Group"
                    Case Else
                        picnum = 3
                        ndesc = "Other"
                End Select
                    If .Operator = True Then
                        Set li = frm_main.lstLogin.ListItems.Add(, , .auth_user, , picnum)
                   Else
                        Set li = frm_main.lstLogin.ListItems.Add(, , .auth_user, , picnum)
                   End If
                           li.SubItems(1) = ndesc
                            If .enabled = True Then
                                li.SubItems(2) = "Yes"
                            Else
                                li.SubItems(2) = "No"
                            End If
                            li.SubItems(3) = .logged & "/" & .maxLogins
        End If
    End With
Next
End Sub

Public Function GetLoginNumFromName(strLoginName As String) As Long
For i = 0 To UBound(Logins)
    With Logins(i)
        If Trim(LCase(.auth_user)) = Trim(LCase(strLoginName)) Then
            GetLoginNumFromName = i
            Exit Function
        End If
    End With
Next
GetLoginNumFromName = -255
End Function

Public Function AuthInfo(indata As String) As String
Dim lItemNum
lItemNum = 1
    Dim arrItems() As String
    arrItems = Split(indata, ParseLoginAuth, , vbTextCompare)
    If lItemNum <= UBound(arrItems) Then
        AuthInfo = arrItems(lItemNum)
    End If
    Erase arrItems
End Function

Public Function IsValidLogin(loginuser As String, Optional loginpass As String, Optional ByVal outLoginNo As Long) As Boolean
For i = 0 To UBound(Logins())
    With Logins(i)
        If Native(.auth_user) = Native(loginuser) Then
            If Logins(GetLoginNumFromName(.auth_user)).auth_pass <> "" Then
                If Native(.auth_pass) = Native(loginpass) Then
                    outLoginNo = i
                     IsValidLogin = True
                    Exit Function
                Else
                    outLoginNo = i
                    IsValidLogin = False
                    Exit Function
                End If
            Else
            outLoginNo = i
            IsValidLogin = True
            Exit Function
            End If
        End If
    End With
Next
 IsValidLogin = False
End Function

Public Function canLoginAccept(loginID As Long) As Boolean
With Logins(loginID)
    If .enabled = True Then
        If .maxLogins = 0 Then
            canLoginAccept = True
        Else
            If .logged < .maxLogins Then
                canLoginAccept = True
                .logged = .logged + 1
            Else
                canLoginAccept = False
            End If
        End If
    Else
        canLoginAccept = False
    End If
End With
End Function

Public Sub AcceptConnect(indata As String, ByVal WinsockIndex As Long, Auth As Boolean)
With frm_main.ws_data(WinsockIndex)
    .Tag = Val(Parse(indata, 1))
    If Auth = False Then
        AddUser Val(Parse(indata, 1)), Parse(indata, 2), .RemoteHostIP, .Index, "none", 255
    Else
        AddUser Val(Parse(indata, 1)), Parse(indata, 2), .RemoteHostIP, .Index, Parse(AuthInfo(indata), 0), 1
    End If
    SendWelcomeMsg (WinsockIndex)
        DoEvents
        Broadcast joins, Parse(indata, 2)
End With
End Sub

Public Function GetLID(strUsername As String) As Long
For i = 0 To UBound(Logins())
    With Logins(i)
        If Native(.auth_user) = Native(strUsername) Then
            GetLID = i
        End If
    End With
Next
End Function

Public Sub ShowAcctSet(strUsername As String)
Dim lld As Long
Dim ln As String

With frm_acctset
    ln = LTrim(RTrim(strUsername))
    lld = GetLoginNumFromName(ln)
    .Acctname = ln
    .Caption = "Account settings [" & ln & "]"
    .acctNum = lld
    .Text1.Text = Logins(lld).maxLogins
    .Text2.Text = LTrim(RTrim(Logins(lld).auth_pass))
        If Logins(lld).Operator = True Then
            .Check4.Value = 1
        Else
            .Check4.Value = 0
        End If
        If Logins(lld).enabled = True Then
            .Check1.Value = 1
        Else
            .Check1.Value = 0
        End If
    .Show
End With
End Sub

Public Sub KickUsersFromAcct(strAcctName As String, Optional strReason As String)
For i = 0 To UBound(users)
    If Native(Parse(users(i), 4)) = Native(strAcctName) Then
        Send GetUserWsk(i), kick, strReason
        DoEvents
        frm_main.ws_data(Val(Parse(users(i), 3))).Close
        DoEvents
        CloseNotNeeded
    End If
Next
End Sub

Public Sub MessageAcct(strAcctName As String, strMessage As String)
For i = 0 To UBound(users)
    If Native(Parse(users(i), 4)) = Native(strAcctName) Then
        Send GetUserWsk(i), message, "Server> " & Sep & strMessage & vbCrLf
    End If
Next
End Sub

Public Sub NoticeAcct(strAcctName As String, strMessage As String)
For i = 0 To UBound(users)
    If Native(Parse(users(i), 4)) = Native(strAcctName) Then
        Send GetUserWsk(i), notice, strMessage
    End If
Next
End Sub

Public Function GetUserWsk(ByVal userIndex As Long) As Long

GetUserWsk = Val(Parse(users(userIndex), 3))
End Function
Public Function GetUserWskFromUID(ByVal userUID As Long) As Long
Dim uid As Long
uid = FindUser(userUID)
If uid <> noSuchuser Then
GetUserWskFromUID = Val(Parse(users(uid), 3))
Else
GetUserWskFromUID = -255
End If
End Function

Public Function Native(inString As String) As String
Native = LTrim(RTrim(LCase(inString)))
End Function
