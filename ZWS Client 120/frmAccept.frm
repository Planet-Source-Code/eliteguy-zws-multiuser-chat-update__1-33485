VERSION 5.00
Begin VB.Form frmAccept 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Accept message"
   ClientHeight    =   1680
   ClientLeft      =   2370
   ClientTop       =   1695
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3795
      Begin VB.CommandButton Command2 
         Caption         =   "Reject"
         Height          =   315
         Left            =   2880
         TabIndex        =   3
         Top             =   1080
         Width           =   795
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Accept"
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "The user "" "" has requested a private chat session with you, do you accept?"
         Height          =   915
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   3555
      End
   End
End
Attribute VB_Name = "frmAccept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
