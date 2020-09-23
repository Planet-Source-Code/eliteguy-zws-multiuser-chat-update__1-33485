VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frm_main 
   Caption         =   "ZWS Server "
   ClientHeight    =   4305
   ClientLeft      =   2235
   ClientTop       =   1965
   ClientWidth     =   5595
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
   ScaleHeight     =   4305
   ScaleWidth      =   5595
   Begin MSWinsockLib.Winsock ws_data 
      Index           =   0
      Left            =   2520
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   4600
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3360
      Top             =   1560
   End
   Begin MSComctlLib.ImageList varToolBars 
      Left            =   3780
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4050
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1164
            MinWidth        =   7
            Text            =   "0 Users"
            TextSave        =   "0 Users"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1085
            MinWidth        =   9
            Text            =   " Offline"
            TextSave        =   " Offline"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock ws_listen 
      Left            =   2940
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   4000
   End
   Begin VB.Frame f_manage 
      Caption         =   "Manage"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   180
      TabIndex        =   5
      Top             =   420
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Frame Frame5 
         Caption         =   "User Restrictions && options"
         Height          =   2415
         Left            =   240
         TabIndex        =   29
         Top             =   660
         Visible         =   0   'False
         Width           =   3855
         Begin VB.ListBox lstUserRest 
            Height          =   2010
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   30
            Top             =   300
            Width           =   3615
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Login Accounts"
         Height          =   2415
         Left            =   240
         TabIndex        =   31
         Top             =   660
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CommandButton lbs 
            Caption         =   "Settings..."
            Height          =   315
            Left            =   180
            TabIndex        =   35
            Top             =   1920
            Width           =   915
         End
         Begin VB.CommandButton lbr 
            Caption         =   "Remove"
            Height          =   315
            Left            =   2280
            TabIndex        =   34
            Top             =   1920
            Width           =   795
         End
         Begin VB.CommandButton lba 
            Caption         =   "Add"
            Height          =   315
            Left            =   3120
            TabIndex        =   33
            Top             =   1920
            Width           =   555
         End
         Begin MSComctlLib.ListView lstLogin 
            Height          =   1995
            Left            =   120
            TabIndex        =   32
            Top             =   300
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   3519
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            Icons           =   "UserPix"
            SmallIcons      =   "UserPix"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Login"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Type"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Enabled"
               Object.Width           =   1323
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Logged"
               Object.Width           =   1323
            EndProperty
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Local Settings"
         Height          =   2415
         Left            =   240
         TabIndex        =   36
         Top             =   660
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CommandButton Command2 
            Caption         =   "Set"
            Height          =   315
            Left            =   1260
            TabIndex        =   39
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox Text4 
            Height          =   315
            Left            =   180
            TabIndex        =   38
            Text            =   "4000"
            Top             =   600
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "Server Port:"
            Height          =   195
            Left            =   180
            TabIndex        =   37
            Top             =   360
            Width           =   1995
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Served"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   240
         TabIndex        =   20
         Top             =   1500
         Visible         =   0   'False
         Width           =   3855
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   525
            Left            =   3060
            TabIndex        =   28
            Top             =   180
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   926
            ButtonWidth     =   847
            ButtonHeight    =   926
            Style           =   1
            ImageList       =   "varToolBars"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Reset"
                  ImageIndex      =   1
               EndProperty
            EndProperty
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000010&
            X1              =   2985
            X2              =   2985
            Y1              =   1545
            Y2              =   120
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            X1              =   2925
            X2              =   2925
            Y1              =   1530
            Y2              =   105
         End
         Begin VB.Label msg 
            Caption         =   "0 users since last restart"
            Height          =   195
            Left            =   180
            TabIndex        =   26
            Top             =   1260
            Width           =   3495
         End
         Begin VB.Label msgt 
            Caption         =   "0 Messages total"
            Height          =   315
            Left            =   180
            TabIndex        =   25
            Top             =   1080
            Width           =   3495
         End
         Begin VB.Label byt 
            Caption         =   "0 bytes since last restart"
            Height          =   315
            Left            =   180
            TabIndex        =   24
            Top             =   840
            Width           =   3495
         End
         Begin VB.Label bytt 
            Caption         =   "0 bytes total"
            Height          =   315
            Left            =   180
            TabIndex        =   23
            Top             =   660
            Width           =   3495
         End
         Begin VB.Label usr 
            Caption         =   "0 users since last restart"
            Height          =   315
            Left            =   180
            TabIndex        =   22
            Top             =   420
            Width           =   3495
         End
         Begin VB.Label usrt 
            Caption         =   "0 users total"
            Height          =   315
            Left            =   180
            TabIndex        =   21
            Top             =   240
            Width           =   3075
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Runtime"
         Height          =   780
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   3855
         Begin VB.Label Label1 
            Caption         =   "0:0:0"
            Height          =   315
            Left            =   180
            TabIndex        =   18
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Send Notice Message"
         Height          =   1095
         Left            =   240
         TabIndex        =   13
         Top             =   1860
         Width           =   3855
         Begin VB.CommandButton cmdSN 
            Caption         =   "Send"
            Height          =   315
            Left            =   2820
            TabIndex        =   19
            Top             =   660
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Left            =   180
            TabIndex        =   16
            Top             =   300
            Width           =   3495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Welcome message"
         Height          =   1095
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   3855
         Begin VB.TextBox Text2 
            Height          =   315
            Left            =   180
            TabIndex        =   15
            Top             =   660
            Width           =   3495
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Enable"
            Height          =   255
            Left            =   180
            TabIndex        =   14
            Top             =   300
            Width           =   3315
         End
      End
      Begin MSComctlLib.TabStrip setting 
         Height          =   2955
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   5212
         HotTracking     =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   5
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Messages"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Statistics"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "User Restrictions"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Login"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Local"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame f_users 
      Caption         =   "Users"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   180
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Timer status_timer 
         Interval        =   1000
         Left            =   3600
         Top             =   1140
      End
      Begin MSComctlLib.ImageList UserPix 
         Left            =   2280
         Top             =   2280
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":02D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0638
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0A0B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0D5D
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":10AF
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1401
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView UsrList 
         Height          =   3015
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "UserPix"
         SmallIcons      =   "UserPix"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Username"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "IP"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Sck"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Acct"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame F_chat 
      Caption         =   "Chat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   180
      TabIndex        =   4
      Top             =   420
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton Command1 
         Caption         =   "Send"
         Height          =   315
         Left            =   3540
         TabIndex        =   9
         Top             =   2940
         Width           =   675
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   2940
         Width           =   3375
      End
      Begin RichTextLib.RichTextBox ChatTxt 
         Height          =   2655
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4683
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Form1.frx":1753
      End
   End
   Begin VB.Frame f_Activity 
      Caption         =   "Activity"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   180
      TabIndex        =   3
      Top             =   420
      Width           =   4335
      Begin VB.ListBox lstActivity 
         Height          =   3000
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4095
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3855
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6800
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Activity"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Users"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Chat"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Manage"
            ImageVarType    =   2
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   1085
      ButtonWidth     =   609
      ButtonHeight    =   926
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.Menu mnuServer 
      Caption         =   "&Server"
      Begin VB.Menu mnuSS 
         Caption         =   "&Start"
      End
      Begin VB.Menu mnuSSD 
         Caption         =   "&Shut Down"
      End
      Begin VB.Menu mnue 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExt 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "&User"
      Begin VB.Menu mnuKick 
         Caption         =   "&Kick"
      End
      Begin VB.Menu mnuMK 
         Caption         =   "&Mass Kick"
         Begin VB.Menu mnuBG 
            Caption         =   "&By Login Account"
         End
         Begin VB.Menu mb4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuaa 
            Caption         =   "&All users"
         End
      End
   End
   Begin VB.Menu mnulogin 
      Caption         =   "&Login"
      Begin VB.Menu mnuadd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuremove 
         Caption         =   "&Remove"
      End
      Begin VB.Menu mnub3 
         Caption         =   "-"
      End
      Begin VB.Menu mnusettt 
         Caption         =   "&Settings"
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public userAcct2 As String
Private Sub ChatTxt_Change()
ChatTxt.SelLength = Len(ChatTxt.Text)
End Sub

Private Sub Check1_Click()
SaveSetting "ZWSserver", "Settings", "SWelcomeMsg", Check1.Value
Server.showWlcmMsg = Check1.Value
End Sub

Private Sub Check2_Click()
SaveSetting "ZWSserver", "Settings", "SByeMsg", Check2.Value
Server.showByeMsg = Check2.Value
End Sub

Private Sub cmdSN_Click()
Broadcast notice, Text3.Text
End Sub

Private Sub Command1_Click()
Broadcast message, "Server> " & Sep & Text1.Text
ChatTxt.Text = ChatTxt.Text & "Server > " & Text1.Text & vbCrLf
Text1.Text = ""
End Sub

Private Sub Command2_Click()
MsgBox "Server will restart on port " & Text4.Text, vbInformation, "Change"
Server.listenport = Text4.Text
SaveSetting "ZWSserver", "Server", "Port", Text4.Text
StopServer
StartServer
End Sub

Private Sub Form_Load()
Dim ffile
Dim llst() As String
ReDim Logins(1)

If Right(App.Path, 1) = "\" Then
    LoginFilename = App.Path & "logins.lst"
Else
    LoginFilename = App.Path & "\logins.lst"
End If
Server.lastShutDown = "Never"
'Dim bld As Long
'bld = GetSetting("ZWSserver", "Version", "Build", 0)
'bld = bld + 1
'SaveSetting "ZWSserver", "Version", "Build", bld
bld = 945
initWsServer
StartServer

Server.build = Format(Str(bld), "000#")
If ReadLogins(llst) = False Then
    ffile = FreeFile
    Open LoginFilename For Append As ffile
    Close ffile
    act "ERR: No logins found (first run?) (OK)"
Else
    ParseLogins llst
    DoEvents
    AddLoginsToList
        act "OK: Login file found, " & (UBound(Logins)) & " Logins parsed"
End If

Me.Caption = "ZWS Server " & App.Major & "." & App.Minor & "." & App.Revision & " Build " & Server.build
lstUserRest.AddItem "Allow Multiple Logins (clone users)"
lstUserRest.AddItem "Respond to remote command requests"
lstUserRest.AddItem "Require users to login"
lstUserRest.AddItem "Display Group in client userlist ex: usr [group]"
GetSettings
Me.Move Me.Left, Me.Top, f_Width + 30, f_Height + 30
Server.StartTime = DateTime.Now

If Server.rest_isLogin = True Then
    act "OK: Login system enabled"
Else
    act "OK: Login system is disabled"
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting "ZWSserver", "Stats", "bytes", Server.stat_bytesT
    SaveSetting "ZWSserver", "Stats", "msg", Server.stat_msgsT
    SaveSetting "ZWSserver", "Stats", "users", Server.stat_usersT
    StopServer
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.Width < (f_Width - 15) And Me.Height < (f_Height - 15) Then
    Me.Move Me.Left, Me.Top, f_Width + 30, f_Height + 30
    GoTo 1:
   Else
    If Me.Width < (f_Width - 15) Then
        Me.Move Me.Left, Me.Top, f_Width + 30, Me.Height
        GoTo 1:
    End If
    If Me.Height < (f_Height - 15) Then
        Me.Move Me.Left, Me.Top, Me.Width, f_Height + 30
        GoTo 1:
    End If
End If

1:
TabStrip1.Move 0, 60, Me.ScaleWidth, Me.ScaleHeight - StatusBar1.Height - 60
f_users.Move 90, TabStrip1.ClientTop, TabStrip1.ClientWidth - 60, TabStrip1.ClientHeight - 30
f_Activity.Move 90, TabStrip1.ClientTop, TabStrip1.ClientWidth - 60, TabStrip1.ClientHeight - 30
F_chat.Move 90, TabStrip1.ClientTop, TabStrip1.ClientWidth - 60, TabStrip1.ClientHeight - 30
f_manage.Move 90, TabStrip1.ClientTop, TabStrip1.ClientWidth - 60, TabStrip1.ClientHeight - 30
UsrList.Move 0 + 90, 0 + 200, f_users.Width - 180, f_users.Height - 290
lstActivity.Move 0 + 90 + 15, 0 + 200 + 15, F_chat.Width - 180 - 30, F_chat.Height - 290 - 30
ChatTxt.Move 90, 200, F_chat.Width - 180, F_chat.Height - 290 - Text1.Height
Text1.Move 90 + 15, ChatTxt.Height + 200, F_chat.Width - Command1.Width - (90 + 15) - 115
Command1.Move Text1.Width + 115, Text1.Top
setting.Move 90, 260, f_manage.Width - 180, f_manage.Height - 260 - 90
If f_manage.Visible = True Then
Call MoveSettings
End If
Label1.Move 180, Label1.Top, Frame1.Width - (180 * 2)
cmdSN.Move Text3.Width - cmdSN.Width + 180
Exit Sub

End Sub

Private Sub lba_Click()
frm_User.Show
End Sub

Private Sub lbr_Click()
On Error Resume Next
If lstLogin.SelectedItem.Text <> "" Then
RemoveLogin (lstLogin.SelectedItem.Text)
End If

End Sub

Private Sub lbs_Click()
On Error Resume Next
ShowAcctSet (lstLogin.SelectedItem.Text)
End Sub

Private Sub lstLogin_DblClick()

ShowAcctSet (lstLogin.SelectedItem.Text)
End Sub


Private Sub lstLogin_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 Then
    PopupMenu mnulogin
End If
End Sub

Private Sub lstUserRest_Click()
With lstUserRest
Select Case .ListIndex
Case 0
    .ToolTipText = "Can multiple clients have the same username?"
        Server.rest_canMultiUser = .Selected(.ListIndex)
        Server.rest_isLogin = False
        .Selected(2) = False
        SaveSetting "ZWSserver", "Restrictions", "UserLogin", False
        SaveSetting "ZWSserver", "Restrictions", "MultiUser", Server.rest_canMultiUser
Case 1
    .ToolTipText = "Will server send out details if requested?"
        Server.rest_canRespondCmd = .Selected(.ListIndex)
        SaveSetting "ZWSserver", "Restrictions", "Commands", Server.rest_canRespondCmd
Case 2
        .ToolTipText = "Enable the login system? multiple users is disabled if this is enabled."
        Server.rest_isLogin = .Selected(.ListIndex)
        Server.rest_canMultiUser = False
        .Selected(0) = False
        SaveSetting "ZWSserver", "Restrictions", "UserLogin", Server.rest_isLogin
        SaveSetting "ZWSserver", "Restrictions", "MultiUser", False
Case 3
        .ToolTipText = "Display the parent group in the client's userlist"
        Server.rest_showGroup = .Selected(.ListIndex)
        SaveSetting "ZWSserver", "Restrictions", "sGroup", Server.rest_showGroup
Case Else
    .ToolTipText = "User restrictions govern the way users interact with the server."
End Select
End With
End Sub

Private Sub mnuaa_Click()
Dim ms As String
If GetInput("Mass kick reason:", ms) = i_ok Then
    Broadcast kick, ms
End If
End Sub

Private Sub mnuadd_Click()
lba_Click
End Sub

Private Sub mnuBG_Click()
Dim ms As String
On Error Resume Next
If UsrList.SelectedItem.Text <> "" Then
    If GetInput("Mass Kick Group " & Logins(GetUserGroup(Val("&H" & UsrList.SelectedItem.Text))).auth_user & " reason:", ms) = i_ok Then
        KickUsersFromAcct Logins(GetUserGroup(Val("&H" & UsrList.SelectedItem.Text))).auth_user, ms
    End If
End If
End Sub

Private Sub mnuExt_Click()
StopServer
DoEvents
Unload Me
End
End Sub

Private Sub mnuKick_Click()
On Error Resume Next
If UsrList.SelectedItem.Text <> "" Then
    Dim ir As String
    If GetInput("Kick Reason:", ir) = i_ok Then
        KickUser (Val("&H" & UsrList.SelectedItem.Text)), ir
    End If
End If
End Sub


Private Sub mnuremove_Click()
lbr_Click
End Sub

Private Sub mnusettt_Click()
lbs_Click
End Sub

Private Sub mnuSS_Click()
StartServer
End Sub

Private Sub mnuSSD_Click()
StopServer
End Sub

Private Sub setting_Click()
Select Case setting.SelectedItem.Index
Case 1
    Frame1.Visible = True
    Frame2.Visible = True
    Frame3.Visible = False
    Frame4.Visible = False
    Frame5.Visible = False
    Frame6.Visible = False
    Frame7.Visible = False
Case 2
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = True
    Frame4.Visible = True
    Frame5.Visible = False
    Frame6.Visible = False
    Frame7.Visible = False
Case 3
    Frame5.Visible = True
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    Frame6.Visible = False
    Frame7.Visible = False
Case 4
    Frame6.Visible = True
    Frame5.Visible = False
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    Frame7.Visible = False
Case 5
    Frame7.Visible = True
    Frame6.Visible = False
    Frame5.Visible = False
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
End Select
End Sub

Private Sub status_timer_Timer()
StatusBar1.Panels(1).Text = getUserCount & " Users"
StatusBar1.Panels(2).Text = " " & serverState & " "
End Sub

Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.Index
    Case 1
        f_users.Visible = False
        f_Activity.Visible = True
        F_chat.Visible = False
        f_manage.Visible = False
        lstActivity.ListIndex = lstActivity.ListCount - 1
    Case 2
        f_users.Visible = True
        f_Activity.Visible = False
        F_chat.Visible = False
        f_manage.Visible = False
    Case 3
        f_users.Visible = False
        f_Activity.Visible = False
        F_chat.Visible = True
        f_manage.Visible = False
            Text1.SetFocus
    Case 4
        MoveSettings
        f_users.Visible = False
        f_Activity.Visible = False
        F_chat.Visible = False
        f_manage.Visible = True
End Select

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1_Click
End If
End Sub

Private Sub Text2_Change()
SaveSetting "ZWSserver", "Settings", "WelcomeMsg", Text2.Text
Server.wlcmMsg = Text2.Text
End Sub


Private Sub Timer1_Timer()
If mServer.runtimeseconds < 59 Then
mServer.runtimeseconds = mServer.runtimeseconds + 1
Else
    mServer.runtimeseconds = 0
        If mServer.runtimeminutes < 59 Then
            mServer.runtimeminutes = mServer.runtimeminutes + 1
        Else
            mServer.runtimeminutes = 0
            mServer.runtimehours = mServer.runtimehours + 1
        End If
End If
Server.runtime = mServer.runtimehours & ":" & mServer.runtimeminutes & ":" & mServer.runtimeseconds
Label1.Caption = Server.runtime & " started on " & Server.StartTime
bytt.Caption = Round((Server.stat_bytesT / 1024), 2) & " bytes total (kb)"
msgt.Caption = Server.stat_msgsT & " messages total"
usrt.Caption = Server.stat_usersT & " users total"
byt.Caption = Server.stat_bytes & " bytes this session"
msg.Caption = Server.stat_msgs & " messages this session"
usr.Caption = Server.stat_users & " users this session"

End Sub

Private Sub Timer2_Timer()
StatusBar1.Panels(1).Text = getUserCount
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    SaveSetting "ZWSserver", "Stats", "bytes", 0
    SaveSetting "ZWSserver", "Stats", "msg", 0
    SaveSetting "ZWSserver", "Stats", "users", 0
    With Server
        .stat_bytesT = 0
        .stat_msgsT = 0
        .stat_usersT = 0
    End With
End Select
End Sub

Private Sub UsrList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim ep As String
If UsrList.SelectedItem.Text <> "" Then
ep = Logins(GetUserGroup(Val("&H" & UsrList.SelectedItem.Text))).auth_user
mnuBG.Caption = "By login account [" & ep & "]"
mnuBG.enabled = True
Select Case Button
    Case 2
        PopupMenu mnuUser
End Select
Else
mnuBG.enabled = False
End If
End Sub

Private Sub ws_data_Close(Index As Integer)
If FindUser(ws_data(Index).Tag) <> noSuchuser Then
    RemoveUser (ws_data(Index).Tag)
Else
End If
End Sub

Private Sub ws_data_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim indata As String
Server.stat_bytesT = Server.stat_bytesT + bytesTotal
Server.stat_bytes = Server.stat_bytes + bytesTotal
With ws_data(Index)
    .GetData indata, vbString
    Select Case GetcCode(indata)
        Case ResponseTypes.message ' a message is coming
            Broadcast message, Parse(indata, 1) & Sep & Parse(indata, 2)
            ChatTxt.Text = ChatTxt.Text & Parse(indata, 1) & "> " & Parse(indata, 2) & vbCrLf
            Server.stat_msgsT = Server.stat_msgsT + 1
            Server.stat_msgs = Server.stat_msgs + 1
        Case ResponseTypes.connecting ' client connect request
            '-----------------------------------------------------------
            'all connection requests should be formed like so:
            'con & sep & userid (long) & sep & nickname & sep & ParseLoginAuth
            '& loginUser & sep & loginPass & sep & ('any random data)
                    .Tag = Val(Parse(indata, 1))
         If RTrim(LTrim(Parse(indata, 1))) <> "" Then
          If checkUserRestrictions(UserRestV.multiplelogins, Parse(indata, 2)) = True Then
            If Server.rest_isLogin = False Then
                    AcceptConnect indata, Index, False
            Else
                'authentication sub
                Dim ath As String
                Dim lid As Long
                ath = AuthInfo(indata) 'authorization string (user & sep & pass)
                If LTrim(RTrim(ath)) <> "" Then
                    If IsValidLogin(Parse(ath, 0), Parse(ath, 1)) = True Then
                        lid = GetLID(Parse(ath, 0))
                        If canLoginAccept(lid) = True Then
                            'login script here
                            AcceptConnect indata, Index, True
                            Send Index, message, vbCrLf & "**User " & Parse(ath, 0) & " Authenticated, welcome"
                        Else
                            Send Index, message, vbCrLf & "**User " & Parse(ath, 0) & " User account is disabled"
                            DoEvents
                            .Close
                        End If
                    Else
                        Send Index, message, vbCrLf & "**User " & Parse(ath, 0) & " invalid authentication"
                        DoEvents
                            .Close
                    End If
                Else
                    Send Index, message, vbCrLf & "**User " & Parse(indata, 2) & " is not a valid account on this server"
                    DoEvents
                    .Close
                End If
            End If
        Else
                Send Index, message, vbCrLf & "User " & Parse(indata, 2) & " exists on the server" & vbCrLf & "Server does not allow multiple logins, please change your username"
                    DoEvents
                    .Close
                    'Exit Sub
            End If
           Else
                    KickUser Val(Parse(indata, 0)), "Invalid username"
            End If
                    DoEvents
                    UpdateUserList
        Case ResponseTypes.disconnecting 'client disconnect request
                .Close
                DoEvents
                UpdateUserList
                'Exit Sub
        Case ResponseTypes.commandRequest
            Dim cmdreq As CommandTypes
            cmdreq = Val(Parse(indata, 1))
            respondToCommandRequest .Index, cmdreq, Parse(indata, 2)
        Case ResponseTypes.Auth
        Case ResponseTypes.pm
            Dim targetUser As Long
            'personal message, only forward to the correct client
            'pm messages are formatted as follows:
            ' pm response code(10) & sep & targetUserId & sep & sourceUserName _
                & sep & source userid & sep & message & sep & timesent
            targetUser = GetUserWskFromUID(Val(Parse(indata, 1)))
            If targetUser <> -255 Then
                ws_data(targetUser).SendData ResponseTypes.pm _
                & Sep & Parse(indata, 2) & Sep & Parse(indata, 4) & Sep & _
                Parse(indata, 3) & Sep & Parse(indata, 5)
            Else
                .SendData ResponseTypes.pmclose & Sep & Parse(indata, 1)
            End If
            'pm(0),username(1),message(2),userid(3),time(4)
        Case ResponseTypes.pmclose
            targetUser = GetUserWskFromUID(Val(Parse(indata, 1)))
            If targetUser <> -255 Then
                ws_data(targetUser).SendData ResponseTypes.pm _
                & Sep & Parse(indata, 2) & Sep & Parse(indata, 4) & Sep & _
                Parse(indata, 3) & Sep & Parse(indata, 5)
            Else
                .SendData ResponseTypes.pmclose & Sep & Parse(indata, 1)
            End If
        Case Else
            'ignore and terminate all other requests.
            indata = ""
            .Close
    End Select
End With
End Sub

Private Sub ws_data_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
mod_WsServer.LogError "ER: " & Number & "-" & Format(DateTime.Now, "MM/DD/YY HH:MM:SS") & "-" & Index & "-" & Description
act "ERR: " & Index & " " & Hex(ws_data(Index).Tag) & " > " & Description
RemoveUser (ws_data(Index).Tag)
ws_data(Index).Close
End Sub

Private Sub ws_listen_ConnectionRequest(ByVal rid As Long)
CloseNotNeeded
ws_data(OpenWsk).Accept rid
End Sub

Public Sub MoveSettings()
Frame1.Move setting.ClientLeft + 90, setting.ClientTop + 90, setting.ClientWidth - 180
Frame2.Move setting.ClientLeft + 90, setting.ClientTop + 90 + Frame1.Height, setting.ClientWidth - 180
Frame3.Move setting.ClientLeft + 90, setting.ClientTop + 90, setting.ClientWidth - 180
Frame4.Move setting.ClientLeft + 90, Frame3.Height + Frame3.Top, setting.ClientWidth - 180
Frame5.Move setting.ClientLeft + 90, setting.ClientTop + 90, setting.ClientWidth - 180, setting.ClientHeight - 160
Frame6.Move setting.ClientLeft + 90, setting.ClientTop + 90, setting.ClientWidth - 180, setting.ClientHeight - 160
Frame7.Move setting.ClientLeft + 90, setting.ClientTop + 90, setting.ClientWidth - 180, setting.ClientHeight - 160
lstLogin.Move 90, 260, Frame6.Width - 180, Frame6.Height - (345 + lbr.Height)
lba.Move lstLogin.Left + (lstLogin.Width - lba.Width) - 15, lstLogin.Top + lstLogin.Height
lbr.Move lba.Left - lbr.Width - 15, lba.Top
lbs.Move lstLogin.Left + 15, lba.Top
lstUserRest.Move 0, 0, Frame5.Width, Frame5.Height
Text2.Move 180, Text2.Top, Frame1.Width - (180 * 2)
Text3.Move 180, Text3.Top, Frame2.Width - (180 * 2)
End Sub
