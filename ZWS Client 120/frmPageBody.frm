VERSION 5.00
Begin VB.Form frmPageBody 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3180
   ClientLeft      =   3675
   ClientTop       =   1620
   ClientWidth     =   4740
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
   ScaleHeight     =   3180
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPage 
      Height          =   3015
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   4635
   End
End
Attribute VB_Name = "frmPageBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
txtPage.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
