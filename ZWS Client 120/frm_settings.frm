VERSION 5.00
Begin VB.Form frm_settings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   4500
   ClientLeft      =   3555
   ClientTop       =   1410
   ClientWidth     =   3930
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Defaults"
      Height          =   315
      Left            =   60
      TabIndex        =   15
      Top             =   4080
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2280
      TabIndex        =   14
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   315
      Left            =   3060
      TabIndex        =   13
      Top             =   4080
      Width           =   795
   End
   Begin VB.Frame Frame3 
      Caption         =   "Server"
      Height          =   975
      Left            =   60
      TabIndex        =   8
      Top             =   900
      Width           =   3795
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2820
         TabIndex        =   10
         Text            =   "4000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Text            =   "elitewrz.mine.nu"
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Port"
         Height          =   255
         Left            =   2820
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Server Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nickname"
      Height          =   795
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   3795
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Text            =   "Guest & ip"
         Top             =   300
         Width           =   3555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Authentication"
      Height          =   2055
      Left            =   60
      TabIndex        =   0
      Top             =   1920
      Width           =   3795
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Text            =   "guest"
         Top             =   900
         Width           =   3015
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Server requires me to authenticate"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Value           =   1  'Checked
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   1995
      End
      Begin VB.Label Label1 
         Caption         =   "Username:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   660
         Width           =   2835
      End
   End
End
Attribute VB_Name = "frm_settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Text1.Enabled = Check1.Value
Text2.Enabled = Check1.Value
End Sub

Private Sub Command1_Click()
If LTrim(RTrim(Text3.Text)) = "" Then
MsgBox "Nickname is too short"
Text3.SetFocus
Exit Sub
End If
If LTrim(RTrim(Text4.Text)) = "" Then
MsgBox "Invalid server address"
Text4.SetFocus
Exit Sub
End If
If LTrim(RTrim(Text5.Text)) = "" Then
MsgBox "Invalid port"
Text5.SetFocus
Exit Sub
End If
If Check1.Value = 1 Then
    If LTrim(RTrim(Text1.Text)) = "" Then
        MsgBox "Username is too short"
        Text1.SetFocus
        Exit Sub
    Else
    End If
End If
            With Client
            .username = Text3.Text
            .serveraddress = Text4.Text
            .serverport = Text5.Text
            .auth = Val(Check1.Value)
            .a_loginname = Text1.Text
            .a_loginpass = Text2.Text
            SaveSettings
        End With
        Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Text4.Text = "elitewrz.mine.nu"
Text5.Text = 4000
Check1.Value = 1
Text1.Text = "guest"
Text3.Text = "guest " & frm_main.ws_main.LocalIP
End Sub

Private Sub Form_Load()
Me.Show
        With Client
           Text3.Text = .username
            Text4.Text = .serveraddress
            Text5.Text = .serverport
            Check1.Value = Val(.auth)
            Text1.Text = .a_loginname
             Text2.Text = .a_loginpass
        End With
End Sub

