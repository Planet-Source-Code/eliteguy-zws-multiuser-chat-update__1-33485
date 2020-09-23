VERSION 5.00
Begin VB.Form frm_User 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Account"
   ClientHeight    =   4020
   ClientLeft      =   3480
   ClientTop       =   1515
   ClientWidth     =   3900
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
   ScaleHeight     =   4020
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Group Account"
      Height          =   1335
      Left            =   60
      TabIndex        =   8
      Top             =   2220
      Width           =   3795
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "1"
         Top             =   840
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "This is a Group Account"
         Height          =   210
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Max Logins:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   2955
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Add another Acct."
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   3660
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2340
      TabIndex        =   4
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   3600
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Login"
      Height          =   2115
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   3795
      Begin VB.CheckBox Check4 
         Caption         =   "Add account as operator"
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   1740
         Width           =   3555
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1260
         Width           =   3315
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Password"
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   2835
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   540
         Width           =   3315
      End
      Begin VB.Label Label1 
         Caption         =   "Username"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   300
         Width           =   2955
      End
   End
End
Attribute VB_Name = "frm_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Text3.enabled = Check1.Value
End Sub

Private Sub Check3_Click()
Text2.enabled = Check3.Value
Label2.enabled = Check3.Value
End Sub

Private Sub Command1_Click()
If LTrim(RTrim(Text1.Text)) = "" Then
    MsgBox "Invalid username"
    Exit Sub
End If
If Check3.Value = 1 Then
    If LTrim(RTrim(Text2.Text)) = "" Then
        MsgBox "Invalid password"
        Exit Sub
    End If
End If
If loginExists(Text1.Text) Then
    MsgBox "Login name exists!"
    Exit Sub
End If

If Check1.Value = 0 Then
    AddLogin Text1.Text, Text2.Text, True, 1, Val(Text3.Text), Check4.Value
Else
    AddLogin Text1.Text, Text2.Text, True, 2, Val(Text3.Text), Check4.Value
End If
If Check2.Value = 1 Then
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
Else
    Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
