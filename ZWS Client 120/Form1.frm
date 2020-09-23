VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_main 
   Caption         =   "ZWS Client"
   ClientHeight    =   3270
   ClientLeft      =   1830
   ClientTop       =   1800
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   4965
   Begin VB.ListBox List1 
      Height          =   2160
      Left            =   1200
      TabIndex        =   3
      Top             =   900
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3960
      Top             =   2100
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   3075
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   344
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin zws_client.cht_Window cht_Window1 
      Height          =   1875
      Left            =   480
      TabIndex        =   0
      Top             =   1020
      Width           =   2415
      _extentx        =   4260
      _extenty        =   3307
      caption         =   "Zws Client"
      text            =   ""
   End
   Begin MSWinsockLib.Winsock ws_main 
      Left            =   3180
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "main_tb_imgs"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Connect"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Disconnect"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Settings..."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Paste Text"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Convert to lamer text"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList main_tb_imgs 
      Left            =   0
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":09F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":109A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuViews 
      Caption         =   "&View"
      Begin VB.Menu mnusetting 
         Caption         =   "&Settings"
         Begin VB.Menu mnuLO 
            Caption         =   "&Login options..."
         End
         Begin VB.Menu mb6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuc 
            Caption         =   "&Colors..."
         End
      End
      Begin VB.Menu mnub378 
         Caption         =   "-"
      End
      Begin VB.Menu mnupgs 
         Caption         =   "&Pager..."
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuPM 
      Caption         =   "&PM"
      Begin VB.Menu mnuPMMMMM 
         Caption         =   "Personal Messages"
         Begin VB.Menu mnuAAA 
            Caption         =   "Auto-Accept All"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnur 
            Caption         =   "Auto-Reject All"
         End
         Begin VB.Menu mnub1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuam 
            Caption         =   "&Ask me for confirmation"
         End
      End
      Begin VB.Menu mnub2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelUsr 
         Caption         =   "&User"
         Begin VB.Menu mnuPMME 
            Caption         =   "&Pm (user)"
         End
         Begin VB.Menu gh5 
            Caption         =   "-"
         End
         Begin VB.Menu mnusndPage 
            Caption         =   "&Page user"
            Visible         =   0   'False
         End
         Begin VB.Menu mb23 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUB 
            Caption         =   "&Unblock User (user)"
         End
         Begin VB.Menu mnublock 
            Caption         =   "&Ignore all from (user)"
         End
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelUser As String
Private Sub cht_Window1_ChatSend(txt As String)
'If isConnect Then
    If Left(LTrim(txt), 1) = "/" Then
        sndCommand (txt)
    Else
        If isConnect Then
            snd message, Client.username & Sep & txt
        End If
    End If

'End If
End Sub

Private Sub cht_Window1_UserClick(strUser As String)
If native(strUser) <> native(Client.username) Then
If IsUserBlocked(strUser) = False Then
    'initiate a new personal message session
    OpenPm (strUser)
End If
End If
End Sub

Private Sub Form_Load()
ReDim Blocks(1)
ReDim Pages(1)
Select Case Right(App.Path, 1)
    Case "\"
        mPath = App.Path
    Case Else
        mPath = App.Path & "\"
End Select
LoadBlockList
GetSettings
DoEvents
 sConnect
 ReDim pmMsg(0)
 ReDim pmf(0)
 Dim bld As Long
'bld = GetSetting("ZWSclient", "Version", "Build", 0)
'bld = bld + 1
bld = 102
Select Case Client.PmAcception
    Case PMtype.pmAccept
        Me.mnuAAA.Checked = True
        Me.mnuam.Checked = False
        Me.mnur.Checked = False
    Case PMtype.pmAsk
        Me.mnuam.Checked = True
        Me.mnur.Checked = False
        Me.mnuAAA.Checked = False
    Case PMtype.pmReject
        Me.mnur.Checked = True
        Me.mnuAAA.Checked = False
        Me.mnuam.Checked = False
End Select
LoadColors
Set pmf(0) = New frm_personalMsg
'SaveSetting "ZWSclient", "Version", "Build", bld
cht_Window1.AddMsgToChat Notice, "", "(*) ZWS client version " & App.Major & "." & App.Minor & "." & App.Revision & " Build " & Format(bld, "000#") & vbCrLf & "(%) Type /help for a list of commands" & vbCrLf
Me.SelUser = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
sDisconnect
For i = 0 To UBound(pmf)
    Unload pmf(i)
Next
End Sub

Private Sub Form_Resize()
On Error Resume Next
cht_Window1.Move 0, Toolbar1.Height, Me.ScaleWidth, Me.ScaleHeight - StatusBar1.Height - Toolbar1.Height
End Sub

Private Sub mnuAAA_Click()
mnuAAA.Checked = True
mnur.Checked = False
mnuam.Checked = False
Client.PmAcception = pmAccept
SaveSettings
End Sub

Private Sub mnuam_Click()
mnuAAA.Checked = False
mnur.Checked = False
mnuam.Checked = True
Client.PmAcception = pmAsk
SaveSettings
End Sub

Private Sub mnublock_Click()
If Me.SelUser <> Client.username Then
    If Me.SelUser <> "" Then
        BlockUser (Me.SelUser)
    End If
End If
End Sub

Private Sub mnuc_Click()
frm_colors.Show
End Sub

Private Sub mnuExit_Click()
Unload Me
End
End Sub

Private Sub mnuLO_Click()
If Not (isConnect) Then
    frm_settings.Show
End If
End Sub

Private Sub mnupgs_Click()
frm_pages.Show
End Sub

Private Sub mnuPMME_Click()
If native(Me.SelUser) <> native(Client.username) Then
    'initiate a new personal message session
    OpenPm (Me.SelUser)
End If
End Sub

Private Sub mnur_Click()
mnuAAA.Checked = False
mnur.Checked = True
mnuam.Checked = False
Client.PmAcception = pmReject
SaveSettings
End Sub

Private Sub mnuUB_Click()
If Me.SelUser <> Client.username Then
    If Me.SelUser <> "" Then
        UnblockUser (Me.SelUser)
    End If
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
cht_Window1.Caption = Client.username & " - " & Client.serveraddress & " - " & cht_Window1.UserCount & " Users"
StatusBar1.Panels(1).Text = getstate
List1.Clear
For i = 0 To UBound(pmf)
    List1.AddItem i & " " & pmf(i).Tag
Next
mnuPMME.Caption = "Pm user"
mnublock.Caption = "Block user"
mnuUB.Caption = "Unblock User"
mnuSelUsr.Caption = Me.SelUser
cht_Window1.UpdateBlks
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    sConnect
Case 2
    sDisconnect
Case 4
    If Not (isConnect) Then
        frm_settings.Show
    End If
Case 6
    Dim cb As String
    cb = Clipboard.GetText
    cht_Window1.SetSendText (cb)
Case 7
    cht_Window1.elitetxtme
End Select
End Sub

Private Sub ws_main_Close()
cht_Window1.AddMsgToChat Notice, "", "Connection closed"
cht_Window1.ClearUsers
End Sub

Private Sub ws_main_Connect()
AuthUser
End Sub

Private Sub ws_main_DataArrival(ByVal bytesTotal As Long)
Dim indata As String
ws_main.GetData indata, vbString
Select Case Parse(indata, 0)
    Case ResponseTypes.message
        'cht_Window1.Text = cht_Window1.Text & Parse(indata, 1) & Parse(indata, 2) & vbCrLf
        cht_Window1.AddMsgToChat message, Parse(indata, 1), Parse(indata, 2)
    Case ResponseTypes.Notice
        'cht_Window1.Text = cht_Window1.Text & Parse(indata, 1) & Parse(indata, 2) & vbCrLf
        cht_Window1.AddMsgToChat Notice, "", Parse(indata, 1) & Parse(indata, 2)
    Case ResponseTypes.kick
        cht_Window1.AddMsgToChat kick, "", "You were kicked (" & Parse(indata, 1) & ")"
        ws_main.Close
    Case ResponseTypes.userlist
        Dim sUsers() As String
            sUsers = Split(Parse(indata, 1), UserListSep, , vbTextCompare)
            RefUserList sUsers
                cht_Window1.ClearUsers
                For i = 0 To UBound(sUsers)
                    If LTrim(RTrim(sUsers(i))) <> "" Then
                        cht_Window1.AddUser sUsers(i)
                    End If
                Next

    Case ResponseTypes.commandRequest
        cht_Window1.AddMsgToChat commandRequest, "", "(%) " & Parse(indata, 1)
    Case ResponseTypes.serverstop
        cht_Window1.AddMsgToChat Notice, "", "Server has stopped"
    Case ResponseTypes.pm

        Dim theChatWindow As Long
        'pm(0),username(1),message(2),userid(3),time(4)
        If IsUserBlocked(Parse(indata, 1)) = False Then
    If IsInPMWith(Val(Parse(indata, 3))) = True Then
        'MsgBox "IS"
        'If frm_personalMsg.Tag <> "" Then
        '
        'On Error Resume Next
            theChatWindow = FindRelativePMwindow(Val(Parse(indata, 3)))
            'Set pmf(theChatWindow) = New frm_personalMsg
            pmf(theChatWindow).cht.AddMsgToChat message, Parse(indata, 1), Parse(indata, 2)
            Exit Sub
        Else
        Dim res As Long
            Select Case Client.PmAcception
                Case PMtype.pmAccept
                Case PMtype.pmAsk
                    res = MsgBox("The user " & Parse(indata, 1) & " has requested " _
                    & "a Personal message session with you, do you accept?" _
                    , vbYesNo + vbInformation, "Accept Message")
                    If res = vbYes Then
                        GoTo 2:
                    Else
                        ws_main.SendData ResponseTypes.pmclose & Sep & Parse(indata, 3) & Sep _
                        & Client.username & Sep & Client.UserID & Sep & "The user " & _
                        "declined your chat request" & Sep & DateTime.Now
                        Exit Sub
                    End If
                Case PMtype.pmReject
                        ws_main.SendData ResponseTypes.pmclose & Sep & Parse(indata, 3) & Sep _
                        & Client.username & Sep & Client.UserID & Sep & "The user " & _
                        "is not accepting personal messages" & Sep & DateTime.Now
                        Exit Sub
            End Select
2:
            'start a new pm chat session
                OpenPm (Parse(indata, 1))
                DoEvents
                theChatWindow = FindRelativePMwindow(Val(Parse(indata, 3)))
            'Set pmf(theChatWindow) = New frm_personalMsg
                pmf(theChatWindow).cht.AddMsgToChat message, Parse(indata, 1), Parse(indata, 2)
        End If
        Else
                ws_main.SendData ResponseTypes.pm & Sep & Parse(indata, 3) & Sep _
                        & Client.username & Sep & Client.UserID & Sep & "The user " & _
                        "is not accepting personal messages" & Sep & DateTime.Now
                        Exit Sub
        End If
    Case ResponseTypes.Joins
        cht_Window1.AddMsgToChat Joins, "", "[Joins] " & Parse(indata, 1)
    Case ResponseTypes.Quits
        cht_Window1.AddMsgToChat Quits, "", "[Quits] " & Parse(indata, 1)
    Case ResponseTypes.pmclose
        pmf(FindRelativePMwindow(Val(Parse(indata, 1)))).Hide
        Unload pmf(FindRelativePMwindow(Val(Parse(indata, 1))))
        cht_Window1.AddMsgToChat Notice, "", "The personal message session was terminated"
            
        'pmf(FindRelativePMWindow(Val(Parse(indata, 1)))).Tag = ""
        'pmf(FindRelativePMWindow(Val(Parse(indata, 1)))).Hide
        'Unload pmf(FindRelativePMWindow(Val(Parse(indata, 1))))
    Case ResponseTypes.sPage
        NewPage Val(Parse(indata, 0)), Parse(indata, 1), Parse(indata, 2), Parse(indata, 3), Parse(indata, 4)
    Case Else
         'notin
End Select
End Sub

Public Sub sndCommand(instring As String)
On Error Resume Next
'MsgBox LCase(Left(LTrim(RTrim(Right(inString, Len(inString) - 1))), 4))
Select Case LCase(Left(LTrim(RTrim(Right(instring, Len(instring) - 1))), 4))
    Case "stat" 'server status, remote command
        If isConnect Then
            SndCmdz server_stats, " "
        End If
    Case "help" 'command list, local command
        cht_Window1.AddMsgToChat ResponseTypes.commandRequest, "", _
        "(%) Commands " & vbCrLf & _
        "/stat - server status report" & vbCrLf & _
        "/help - list of commands" & vbCrLf & _
        "/apms - Active personal messages" & vbCrLf & _
        "(%) end command list"
    Case "apms"
        Dim apms As String
        For i = 0 To List1.ListCount - 1
            List1.ListIndex = i
            apms = apms & List1.Text & vbCrLf
        Next
            cht_Window1.AddMsgToChat ResponseTypes.commandRequest, "", _
        "(%) Active pm's: " & vbCrLf & _
        apms & _
        "(%) end command list"
    Case Else
        cht_Window1.AddMsgToChat ResponseTypes.commandRequest, "", _
        "(%) Invalid command (type /help)"
End Select
End Sub

Public Sub SndCmdz(ByVal cmdTyp As CommandTypes, Optional cmdStr As String)
    ws_main.SendData ResponseTypes.commandRequest & Sep & cmdTyp & Sep & cmdStr
End Sub

Public Function getstate() As String
Select Case ws_main.State
    Case StateConstants.sckClosed
        getstate = "Disconnected"
    Case StateConstants.sckClosing
        getstate = "Disconnecting"
        DoEvents
        ws_main.Close
    Case StateConstants.sckConnected
        getstate = "Connected"
    Case StateConstants.sckConnecting
        getstate = "Connecting"
    Case StateConstants.sckConnectionPending
        getstate = "Connection Pending"
    Case StateConstants.sckError
        getstate = "Error!"
        DoEvents
        ws_main.Close
    Case StateConstants.sckHostResolved
        getstate = "Host Resolved"
    Case StateConstants.sckListening
        getstate = "Listening"
        DoEvents
        ws_main.Close
    Case StateConstants.sckOpen
        getstate = "Open"
        DoEvents
        ws_main.Close
    Case StateConstants.sckResolvingHost
        getstate = "Resolving host"
End Select
End Function

