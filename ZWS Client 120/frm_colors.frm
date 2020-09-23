VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_colors 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colors"
   ClientHeight    =   3585
   ClientLeft      =   3690
   ClientTop       =   3450
   ClientWidth     =   3780
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
   ScaleHeight     =   3585
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog comCol 
      Left            =   1200
      Top             =   3180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Ok"
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   3180
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colors"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3675
      Begin VB.Frame Frame2 
         Caption         =   "Presets"
         Height          =   795
         Left            =   180
         TabIndex        =   6
         Top             =   2160
         Width           =   3375
         Begin VB.CommandButton Command3 
            Caption         =   "Save"
            Height          =   315
            Left            =   2160
            TabIndex        =   9
            Top             =   300
            Width           =   615
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Set"
            Height          =   315
            Left            =   2820
            TabIndex        =   8
            Top             =   300
            Width           =   435
         End
         Begin VB.ComboBox Combo1 
            Height          =   330
            Left            =   120
            TabIndex        =   7
            Text            =   "Grays"
            Top             =   300
            Width           =   1995
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   315
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   2715
         TabIndex        =   4
         Top             =   1740
         Width           =   2775
      End
      Begin VB.ListBox lstCol 
         Height          =   1050
         IntegralHeight  =   0   'False
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   3315
      End
      Begin VB.Label Label2 
         Caption         =   "Color:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Item:"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frm_colors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
With Colors
Select Case native(Combo1.Text)
Case "zws default"
    .Chattext = RGB(0, 0, 0)
    .Joins = RGB(200, 200, 200)
    .Quits = RGB(200, 200, 200)
    .Kicks = RGB(200, 100, 100)
    .MyNick = RGB(50, 50, 200)
    .OthersNicks = RGB(100, 100, 200)
    .Notice = RGB(100, 200, 100)
    .PMText = RGB(0, 0, 0)
    .NormalText = RGB(0, 0, 0)
    .Commands = RGB(150, 150, 150)
Case "grays"
    .Chattext = RGB(0, 0, 0)
    .Joins = RGB(200, 200, 200)
    .Quits = RGB(200, 200, 200)
    .Kicks = RGB(160, 160, 160)
    .MyNick = RGB(50, 50, 50)
    .OthersNicks = RGB(120, 120, 120)
    .Notice = RGB(23, 23, 23)
    .PMText = RGB(89, 89, 89)
    .NormalText = RGB(0, 0, 0)
    .Commands = RGB(150, 150, 150)
Case Else
    LoadColorPreset (Combo1.Text)
End Select
End With
SaveColors
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim ir As String
ir = InputBox("Enter a name for this preset", "New Preset", "")
If native(ir) <> "" Then
    SaveColorPreset ir
End If
End Sub

Private Sub Form_Load()
Dim pmd As String
lstCol.AddItem "My Nick"
lstCol.AddItem "Others Nicks"
lstCol.AddItem "Joins"
lstCol.AddItem "Quits"
lstCol.AddItem "Notice"
lstCol.AddItem "Normal Text"
lstCol.AddItem "Chat text"
lstCol.AddItem "Kicks"
lstCol.AddItem "PM Text"
lstCol.AddItem "Commands"
Combo1.AddItem "Grays"
Combo1.AddItem "ZWS Default"
On Error GoTo e:
    Open mPath & "colors.pst" For Input As #1
        Do While Not EOF(1)
            Input #1, pmd
                If Parse(pmd, 0) <> "" Then
                    Combo1.AddItem Parse(pmd, 0)
                End If
            DoEvents
        Loop
e:
    Close #1
End Sub

Private Sub lstCol_Click()
With Colors
Select Case lstCol.ListIndex
    Case 0
        Picture1.BackColor = .MyNick
    Case 1
        Picture1.BackColor = .OthersNicks
    Case 2
        Picture1.BackColor = .Joins
    Case 3
        Picture1.BackColor = .Quits
    Case 4
        Picture1.BackColor = .Notice
    Case 5
        Picture1.BackColor = .NormalText
    Case 6
        Picture1.BackColor = .Chattext
    Case 7
        Picture1.BackColor = .Kicks
    Case 8
        Picture1.BackColor = .PMText
    Case 9
        Picture1.BackColor = .Commands
End Select
End With
End Sub

Private Sub lstCol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Colors
Select Case lstCol.ListIndex
    Case 0
        Picture1.BackColor = .MyNick
    Case 1
        Picture1.BackColor = .OthersNicks
    Case 2
        Picture1.BackColor = .Joins
    Case 3
        Picture1.BackColor = .Quits
    Case 4
        Picture1.BackColor = .Notice
    Case 5
        Picture1.BackColor = .NormalText
    Case 6
        Picture1.BackColor = .Chattext
    Case 7
        Picture1.BackColor = .Kicks
    Case 8
        Picture1.BackColor = .PMText
    Case 9
        Picture1.BackColor = .Commands
End Select
End With
End Sub

Private Sub lstCol_Scroll()
With Colors
Select Case lstCol.ListIndex
    Case 0
        Picture1.BackColor = .MyNick
    Case 1
        Picture1.BackColor = .OthersNicks
    Case 2
        Picture1.BackColor = .Joins
    Case 3
        Picture1.BackColor = .Quits
    Case 4
        Picture1.BackColor = .Notice
    Case 5
        Picture1.BackColor = .NormalText
    Case 6
        Picture1.BackColor = .Chattext
    Case 7
        Picture1.BackColor = .Kicks
    Case 8
        Picture1.BackColor = .PMText
    Case 9
        Picture1.BackColor = .Commands
End Select
End With
End Sub

Private Sub Picture1_DblClick()
comCol.ShowColor
Picture1.BackColor = comCol.Color
With Colors
Select Case lstCol.ListIndex
    Case 0
        .MyNick = Picture1.BackColor
    Case 1
        .OthersNicks = Picture1.BackColor
    Case 2
        .Joins = Picture1.BackColor
    Case 3
        .Quits = Picture1.BackColor
    Case 4
        .Notice = Picture1.BackColor
    Case 5
        .NormalText = Picture1.BackColor
    Case 6
        .Chattext = Picture1.BackColor
    Case 7
        .Kicks = Picture1.BackColor
    Case 8
        .PMText = Picture1.BackColor
    Case 9
        .Commands = Picture1.BackColor
End Select
End With
SaveColors
End Sub
