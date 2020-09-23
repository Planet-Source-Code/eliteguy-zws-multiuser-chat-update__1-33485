VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_acctset 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Account Settings"
   ClientHeight    =   3705
   ClientLeft      =   2310
   ClientTop       =   1590
   ClientWidth     =   4365
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
   ScaleHeight     =   3705
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "Done"
      Height          =   315
      Left            =   3540
      TabIndex        =   11
      Top             =   3300
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Manage"
      Height          =   2595
      Left            =   180
      TabIndex        =   0
      Top             =   540
      Width           =   3975
      Begin VB.CommandButton Command1 
         Caption         =   "Do it"
         Height          =   315
         Left            =   3060
         TabIndex        =   2
         Top             =   2160
         Width           =   795
      End
      Begin VB.ListBox List1 
         Height          =   1800
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   3735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Account History"
      Height          =   2595
      Left            =   180
      TabIndex        =   5
      Top             =   540
      Width           =   3975
      Begin VB.Label Label1 
         Caption         =   "Last login:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   2715
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Account Settings"
      Height          =   2595
      Left            =   180
      TabIndex        =   3
      Top             =   540
      Width           =   3975
      Begin VB.CheckBox Check4 
         Caption         =   "Account is operator"
         Height          =   210
         Left            =   180
         TabIndex        =   14
         Top             =   1140
         Width           =   3555
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   180
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   660
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Apply"
         Height          =   315
         Left            =   3120
         TabIndex        =   10
         Top             =   2220
         Width           =   795
      End
      Begin VB.Frame Frame4 
         Caption         =   "Max logins"
         Height          =   735
         Left            =   0
         TabIndex        =   9
         Top             =   1440
         Width           =   3975
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   180
            MaxLength       =   3
            TabIndex        =   12
            Top             =   300
            Width           =   3615
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Change / Set Password"
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   660
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Account Enabled"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   3315
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3195
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5636
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Manage"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "History"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_acctset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Acctname As String
Public acctNum As Long

Private Sub Command1_Click()
Dim ms As String
Select Case List1.ListIndex
    Case 0
        If GetInput("Kick reason:", ms) = i_ok Then
                MsgBox ms
            KickUsersFromAcct Me.Acctname, ms
        End If

    Case 1
        If GetInput("Kick reason:", ms) = i_ok Then
            KickUsersFromAcct Me.Acctname, ms
        End If
        DoEvents
        With Logins(Me.acctNum)
            .enabled = False
        End With
            WriteLogins
            AddLoginsToList
    Case 2
        If GetInput("Message to group:", ms) = i_ok Then
            MessageAcct Me.Acctname, ms
        End If
    Case 3
        If GetInput("Notice to group:", ms) = i_ok Then
            NoticeAcct Me.Acctname, ms
        End If
    Case 4
        Logins(GetLoginNumFromName(Me.Acctname)).logged = 0
        AddLoginsToList
End Select
End Sub

Private Sub Command2_Click()
With Logins(Me.acctNum)
    .enabled = Check1.Value
    .maxLogins = Val(Text1.Text)
    .Operator = Check4.Value
End With
WriteLogins
AddLoginsToList
End Sub

Private Sub Command3_Click()
With Logins(Me.acctNum)
   .auth_pass = Text2.Text
End With
WriteLogins
AddLoginsToList
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
List1.AddItem "Kick all users on this account"
List1.AddItem "Kick all users on this account; disable account"
List1.AddItem "Message all users on this account"
List1.AddItem "Notice all users on this account"
List1.AddItem "Clear Logins"
End Sub

Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.Index
    Case 1
        Frame1.Visible = True
        Frame3.Visible = False
        Frame2.Visible = False
    Case 2
        Frame2.Visible = True
        Frame1.Visible = False
        Frame3.Visible = False
    Case 3
        Frame3.Visible = True
        Frame1.Visible = False
        Frame2.Visible = False
End Select
End Sub
