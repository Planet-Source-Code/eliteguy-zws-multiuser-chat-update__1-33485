VERSION 5.00
Begin VB.Form frm_input 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Request"
   ClientHeight    =   1650
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   4335
   ControlBox      =   0   'False
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
   ScaleHeight     =   1650
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   1260
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   1380
      TabIndex        =   1
      Top             =   1260
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   300
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Input type:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   2715
   End
End
Attribute VB_Name = "frm_input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  inpResp = i_ok
Unload Me
End Sub

Private Sub Command2_Click()
  inpResp = i_Cancel
Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
inpGotResponse = True
End Sub

Private Sub Text1_Change()
inpRtext = Text1.Text
End Sub
