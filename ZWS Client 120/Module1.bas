Attribute VB_Name = "mod_f"
Public Enum ResponseTypes
    connecting = 1
    disconnecting = 2
    message = 3
    serverstop = 4
    kick = 5
    Notice = 6
    userlist = 7
    commandRequest = 8
    auth = 9
    pm = 10
    Joins = 11
    Quits = 12
    pmclose = 13
    declinepm = 14
    sPage = 15
End Enum
Public Type pmsg
    strUsername As String
    lnguserid As Long
    isChat As Boolean
End Type
Public Type sClient
    serveraddress As String
    serverport As Long
    username As String
    a_loginname As String
    a_loginpass As String
    UserID As Long
    auth As Byte
    PmAcception As PMtype
End Type
Public Enum PMtype
    pmAccept = 0
    pmReject = 1
    pmAsk = 2
End Enum
Public Enum CommandTypes
    server_stats = 1
End Enum

Public Type Userdm
    username As String
    UserID As Long
End Type

Public Type cCol
    MyNick As String
    OthersNicks As String
    Joins As String
    Quits As String
    Notice As String
    NormalText As String
    Chattext As String
    Kicks As String
    PMText As String
    Commands As String
End Type

Public Const Sep As String = vbCrLf & "." & vbCrLf
Public Const lSep As String = vbCrLf & "\«/" & vbCrLf
Public Const noSuchuser As Integer = 32000
Public Const noSuchUserName As String = ""
Public Const UserListSep As String = ""
Public Const UserListSep1 As String = "{«}"
Public Const ParseLoginAuth As String = " & vbcrlf"
Public Client As sClient
Public pmMsg() As pmsg
Public Colors As cCol
Public nUsers() As String
Public pmf() As frm_personalMsg
Public mPath As String
Public Blocks() As String
Public Pages() As String

Public Function Parse(strString As String, ByVal lItemNum As Long) As String
    Dim arrItems() As String
    arrItems = Split(strString, Sep, , vbTextCompare)
    If lItemNum <= UBound(arrItems) Then
        Parse = arrItems(lItemNum)
    End If
    Erase arrItems
End Function
Public Function Parse2(strString As String, ByVal lItemNum As Long, chrSep As String) As String
    Dim arrItems() As String
    arrItems = Split(strString, chrSep, , vbTextCompare)
    If lItemNum <= UBound(arrItems) Then
        Parse2 = arrItems(lItemNum)
    End If
    Erase arrItems
End Function

Public Function GetSettings()
With Client
    .a_loginname = GetSetting("ZWSclient", "login", "user", "guest")
    .a_loginpass = GetSetting("ZWSclient", "login", "pass", "")
    .auth = GetSetting("ZWSclient", "login", "dologin", 1)
    .serveraddress = GetSetting("ZWSclient", "server", "address", "elitewrz.mine.nu")
    .serverport = Val(GetSetting("ZWSclient", "server", "port", "4000"))
    .username = GetSetting("ZWSclient", "user", "username", "Guest " & frm_main.ws_main.LocalIP)
    .PmAcception = GetSetting("ZWSclient", "pm", "mode", 0)
End With
End Function

Public Function SaveSettings()
With Client
    SaveSetting "ZWSclient", "login", "user", .a_loginname
    SaveSetting "ZWSclient", "pm", "mode", .PmAcception
    SaveSetting "ZWSclient", "login", "pass", .a_loginpass
    SaveSetting "ZWSclient", "login", "dologin", .auth
    SaveSetting "ZWSclient", "server", "address", .serveraddress
    SaveSetting "ZWSclient", "server", "port", .serverport
    SaveSetting "ZWSclient", "user", "username", .username
End With
End Function

Public Function isConnect() As Boolean
    If frm_main.ws_main.State = sckConnected Then
        isConnect = True
    Else
        isConnect = False
    End If
End Function

Public Sub sConnect()
frm_main.cht_Window1.ClearUsers
If Not (isConnect) Then
    With Client
        Randomize
        .UserID = Int(Rnd * 1000000000)
        frm_main.ws_main.Connect .serveraddress, .serverport
    End With
End If
End Sub

Public Sub sDisconnect()
frm_main.cht_Window1.ClearUsers
With frm_main.ws_main
    .Close
End With
End Sub

Public Sub AuthUser()
With Client
    If isConnect Then
        snd connecting, .UserID & Sep & .username & Sep & ParseLoginAuth & .a_loginname & Sep & .a_loginpass
    End If
End With
End Sub
Public Sub snd(cCode As ResponseTypes, strdata As String)
    If isConnect Then
        frm_main.ws_main.SendData cCode & Sep & strdata
    End If
End Sub

Public Function IsChattingWith(strUser As String) As Boolean
For i = 0 To UBound(pmMsg)
    If native(pmMsg(i).strUsername) = native(strUser) Then
        If pmMsg(i).isChat = True Then
            IsChattingWith = True
        Else
            IsChattingWith = False
        End If
    Exit Function
    End If
Next
IsChattingWith = False
End Function

Public Function native(instring As String) As String
native = LCase(LTrim(RTrim(instring)))
End Function

Public Function openNewChatWith(ByVal lnguserid As Long) As Integer
For i = 0 To UBound(pmMsg)
    With pmMsg(i)
        If .strUsername = "" Then
            .isChat = True
            .lnguserid = lnguserid
            .strUsername = FindUserName(lnguserid)
            openNewChatWith = i
            Exit Function
        End If
    End With
Next
ReDim Preserve pmMsg(UBound(pmMsg) + 1)
With pmMsg(UBound(pmMsg))
    .isChat = True
    .lnguserid = lnguserid
    .strUsername = FindUserName(lnguserid)
End With
openNewChatWith = UBound(pmMsg)
End Function

Public Sub closeChatWithUSer(ByVal lnguserid As Long)
For i = 0 To UBound(pmMsg)
    With pmMsg(i)
        If .lnguserid = lnguserid Then
            .isChat = False
            .lnguserid = 0
            .strUsername = ""
        End If
    End With
Next
End Sub

Public Function FindUserName(ByVal lnguserid As Long) As String
    For i = 0 To UBound(nUsers)
            If Parse2(nUsers(i), 1, UserListSep1) = lnguserid Then
                FindUserName = Parse2(nUsers(i), 0, UserListSep1)
            Exit Function
            End If
    Next
End Function

Public Function FindUserId(strUsername As String) As Long
    For i = 0 To UBound(nUsers)
            If native(Parse2(nUsers(i), 0, UserListSep1)) = native(strUsername) Then
                FindUserId = Parse2(nUsers(i), 1, UserListSep1)
            Exit Function
            End If
    Next
End Function

Public Sub RefUserList(strUserArry() As String)
ReDim nUsers(UBound(strUserArry))
For i = 0 To UBound(strUserArry)
    nUsers(i) = strUserArry(i)
    strUserArry(i) = Parse2(strUserArry(i), 0, UserListSep1)
Next
End Sub

Public Function FindPM(lnguserid As Long) As Long
For i = 0 To UBound(pmMsg)
    If pmMsg(i).lnguserid = lnguserid Then
        FindPM = i
    Exit Function
    End If
Next
End Function

Public Function AmIChattingWith(lnguserid As Long) As Boolean
On Error Resume Next
For i = 0 To UBound(pmf)
    If pmf(i).Tag = lnguserid Then
        If pmMsg(FindPM(pmf(i).Tag)).isChat = True Then
            AmIChattingWith = True
        Else
            AmIChattingWith = False
        End If
    End If
Next
AmIChattingWith = False
End Function

Public Function FindRelativePMwindow(lnguserid As Long) As Long
'On Error Resume Next
Dim iii As Long
DoEvents

For iii = 0 To UBound(pmf)

    If Val(pmf(iii).Tag) = lnguserid Then
        FindRelativePMwindow = iii
        Exit Function
    End If
Next
FindRelativePMwindow = -255
End Function

Public Sub OpenPm(strUser As String)
On Error Resume Next
Dim ind As Long

For i = 0 To UBound(pmf)
    If pmf(i).Tag = 0 Or pmf(i).Tag = "" Then
        ind = i
        GoTo 1:
    End If
Next
'If UBound(pmf) <> 0 Then
    'make a new form
ind = (UBound(pmf) + 1)
    ReDim Preserve pmf(ind)
'End If
1:
If IsChattingWith(strUser) = False Then

Set pmf(ind) = New frm_personalMsg
Load pmf(ind)
    openNewChatWith (FindUserId(strUser))
    With pmf(ind)
        .Show
        .Tag = FindUserId(strUser)
        .cht.Caption = "Personal message [" & strUser & "]"
        .Caption = "ZWS Client"
        .User = FindUserName(.Tag)
    End With
End If
End Sub

Public Function IsInPMWith(ByVal lnguserid As Long) As Boolean
On Error GoTo er:

If UBound(pmf) = 0 Then
    If Val(pmf(0).Tag) = lnguserid Then
        IsInPMWith = True
    Else
        IsInPMWith = False
    End If
    Exit Function
End If

For i = 0 To UBound(pmf)
    If Val(pmf(i).Tag) = lnguserid Then
        IsInPMWith = True
    End If
Next
IsInPMWith = False
Exit Function
er:
IsInPMWith = False
Resume Next
End Function

Public Sub LoadColors()
With Colors
    .Chattext = GetSetting("ZWSclient", "Colors", "ChatText", RGB(0, 0, 0))
    .Joins = GetSetting("ZWSclient", "Colors", "Joins", RGB(200, 200, 200))
    .Quits = GetSetting("ZWSclient", "Colors", "Quits", RGB(200, 200, 200))
    .Kicks = GetSetting("ZWSclient", "Colors", "Kicks", RGB(200, 100, 100))
    .MyNick = GetSetting("ZWSclient", "Colors", "MyNick", RGB(50, 50, 200))
    .OthersNicks = GetSetting("ZWSclient", "Colors", "ON", RGB(100, 100, 200))
    .Notice = GetSetting("ZWSclient", "Colors", "Notice", RGB(100, 200, 100))
    .PMText = GetSetting("ZWSclient", "Colors", "PMtext", RGB(0, 0, 0))
    .NormalText = GetSetting("ZWSclient", "Colors", "NormalTxt", RGB(0, 0, 0))
    .Commands = GetSetting("ZWSclient", "Colors", "cmd", RGB(150, 150, 150))
End With
End Sub

Public Sub SaveColors()
With Colors
    SaveSetting "ZWSclient", "Colors", "ChatText", .Chattext
    SaveSetting "ZWSclient", "Colors", "Joins", .Joins
    SaveSetting "ZWSclient", "Colors", "Quits", .Quits
    SaveSetting "ZWSclient", "Colors", "Kicks", .Kicks
    SaveSetting "ZWSclient", "Colors", "MyNick", .MyNick
    SaveSetting "ZWSclient", "Colors", "ON", .OthersNicks
    SaveSetting "ZWSclient", "Colors", "Notice", .Notice
    SaveSetting "ZWSclient", "Colors", "PMtext", .PMText
    SaveSetting "ZWSclient", "Colors", "NormalTxt", .NormalText
    SaveSetting "ZWSclient", "Colors", "cmd", .Commands
End With
End Sub

Public Sub LoadColorPreset(strname As String)
On Error GoTo per:
Dim ps As String
    Open mPath & "colors.pst" For Input As #1
        Do While Not EOF(1)
            Input #1, ps
                If Parse(ps, 0) = strname Then
                    With Colors
                        .Chattext = Parse(ps, 1)
                        .Commands = Parse(ps, 2)
                        .Joins = Parse(ps, 3)
                        .Kicks = Parse(ps, 4)
                        .MyNick = Parse(ps, 5)
                        .NormalText = Parse(ps, 6)
                        .Notice = Parse(ps, 7)
                        .OthersNicks = Parse(ps, 8)
                        .PMText = Parse(ps, 9)
                        .Quits = Parse(ps, 10)
                    End With
                    SaveColors
                    Exit Sub
                End If
        DoEvents
        Loop
    Close #1
Exit Sub
per:
    MsgBox "Preset not available"
    Close #1
End Sub
Public Sub SaveColorPreset(strname As String)
On Error GoTo 4:
Dim pmd As String
Dim ps As String
With Colors
ps = strname & Sep _
    & .Chattext & Sep _
    & .Commands & Sep _
    & .Joins & Sep _
    & .Kicks & Sep _
    & .MyNick & Sep _
    & .NormalText & Sep _
    & .Notice & Sep _
    & .OthersNicks & Sep _
    & .PMText & Sep _
    & .Quits
End With
        Open mPath & "colors.pst" For Input As #1
        Do While Not EOF(1)
            Input #1, pmd
            If Parse(pmd, 0) = strname Then
                MsgBox "Preset name exists"
                Close #1
                Exit Sub
            End If
            DoEvents
        Loop
4:
Close #1
    Open mPath & "colors.pst" For Append As #1
        Write #1, ps
    Close #1
End Sub

Public Sub BlockUser(strUsername As String)
Dim ind As Long
For i = 0 To UBound(Blocks)
    If native(Blocks(i)) = "" Then
        ind = i
        GoTo 1:
    End If
Next
ReDim Preserve Blocks(UBound(Blocks) + 1)
ind = UBound(Blocks)
1:
Blocks(ind) = strUsername
SaveBlockList
End Sub

Public Sub UnblockUser(strUsername As String)
For i = 0 To UBound(Blocks)
    If native(Blocks(i)) = native(strUsername) Then
        Blocks(i) = ""
        Exit Sub
    End If
Next
SaveBlockList
End Sub

Public Sub LoadBlockList()
On Error GoTo 1:
Dim Buser As String
Open mPath & "ignore.lst" For Input As #1
    Do While Not EOF(1)
        Input #1, Buser
        If native(Buser) <> "" Then
            Blocks(UBound(Blocks)) = Buser
            ReDim Preserve Blocks(UBound(Blocks) + 1)
        End If
        DoEvents
    Loop
1:
    Close #1

End Sub

Public Sub SaveBlockList()
On Error GoTo 1:
Open mPath & "ignore.lst" For Output As #1
    For i = 0 To UBound(Blocks)
        If native(Blocks(i)) <> "" Then
            Write #1, Blocks(i)
        End If
    Next
1:
Close #1
End Sub

Public Function IsUserBlocked(strUsername As String) As Boolean
For i = 0 To UBound(Blocks)
    If native(Blocks(i)) = native(strUsername) Then
            IsUserBlocked = True
        Exit Function
    End If
Next
IsUserBlocked = False
End Function

'page format: pageid(0),from(1),date(2),subjct(3),message(4)
Public Sub NewPage(ByVal lngPageID As Long, stFrom As String, strDate As String, strSubject As String, strPage As String)
Dim ind As Long
For i = 0 To UBound(Pages)
    If Pages(i) = "" Then
        ind = i
        Pages(ind) = lngPageID & Sep & strFrom & Sep & strDate & Sep & strSubject & Sep & strPage
        Exit Sub
    End If
Next
ReDim Preserve Pages(UBound(Pages) + 1)
ind = UBound(Pages)
Pages(ind) = lngPageID & Sep & strFrom & Sep & strDate & Sep & strSubject & Sep & strPage
End Sub

Public Function findPage(ByVal lngPageID As Long) As Long
For i = 0 To UBound(Pages)
    If Val(Trim(Parse(Pages(i), 0))) = lngPageID Then
        findPage = i
        Exit Function
    End If
Next
findPage = -255
End Function
