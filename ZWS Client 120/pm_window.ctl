VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.UserControl pm_window 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3195
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   0
      Width           =   4335
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1620
         TabIndex        =   0
         Top             =   2760
         Width           =   1275
      End
      Begin RichTextLib.RichTextBox chatTxt 
         Height          =   1095
         Left            =   1380
         TabIndex        =   2
         Top             =   1380
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1931
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"pm_window.ctx":0000
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
      Begin VB.Label cap 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   2895
      End
      Begin VB.Image im 
         Height          =   225
         Left            =   420
         Picture         =   "pm_window.ctx":0077
         Stretch         =   -1  'True
         Top             =   600
         Width           =   3480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   540
         X2              =   3780
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   3240
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Image ii 
         Height          =   195
         Left            =   300
         Top             =   2100
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image iii 
         Height          =   30
         Left            =   660
         Picture         =   "pm_window.ctx":24BB
         Top             =   2100
         Visible         =   0   'False
         Width           =   9600
      End
   End
End
Attribute VB_Name = "pm_window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event ChatSend(txt As String)

Private Sub chatTxt_Change()
chatTxt.SelLength = Len(chatTxt.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If LTrim(RTrim(Text1.Text)) <> "" Then
        RaiseEvent ChatSend(Text1.Text)
        DoEvents
        Text1.Text = ""
        KeyAscii = 0
    End If
End If
'KeyAscii = 0
End Sub

Private Sub UserControl_Initialize()
ii.Picture = im.Picture
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
cap.Caption = PropBag.ReadProperty("Caption", "Caption")
chatTxt.Text = PropBag.ReadProperty("Text", "ChatTxt")
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
Picture1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
im.Move 0, 0, Picture1.ScaleWidth
cap.Move 0 + 30, 0 + 15, Picture1.ScaleWidth
Line1.Y1 = im.Height
Line1.Y2 = im.Height
Line1.X1 = 0
Line1.X2 = Picture1.ScaleWidth
Line2.Y1 = im.Height + 15
Line2.Y2 = im.Height + 15
Line2.X1 = 0
Line2.X2 = Picture1.ScaleWidth
chatTxt.Move 0, im.Height + 30, Picture1.ScaleWidth, Picture1.ScaleHeight - Text1.Height - im.Height - 30
Text1.Move 0, chatTxt.Height + im.Height + 30 + 15, Picture1.Width
End Sub

Public Function UserCount() As Long
UserCount = lstUsers.ListItems.Count
End Function

Public Property Get Caption() As String
    Caption = cap.Caption
End Property

Public Property Let Caption(ByVal vcap As String)
    cap.Caption = vcap
    PropertyChanged "Caption"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Caption", cap.Caption, "Caption"
PropBag.WriteProperty "Text", chatTxt.Text, "ChatTxt"
End Sub

Public Property Get Text() As String
    Text = chatTxt.Text
End Property

Public Property Let Text(ByVal vstr As String)
    chatTxt.Text = vstr
    PropertyChanged "Text"
End Property

Public Sub SetSendText(txt As String)
    Text1.Text = txt
End Sub

Public Function GetSendText() As String
    GetSendText = Text1.Text
End Function

Public Sub LostFoc()
im.Picture = iii.Picture
End Sub

Public Sub GotFoc()
im.Picture = ii.Picture
End Sub

Public Sub AddMsgToChat(vCode As ResponseTypes, strUsername As String, strMessage As String)

    If strUsername <> "" Then
    chatTxt.SelStart = Len(chatTxt.Text)
    chatTxt.SelStart = Len(chatTxt.Text)
    chatTxt.SelBold = True
    If native(strUsername) = native(Client.username) Then
        chatTxt.SelColor = Colors.MyNick
    Else
        chatTxt.SelColor = Colors.OthersNicks
    End If
    chatTxt.SelFontName = "Arial"
    chatTxt.SelFontSize = 8
    chatTxt.SelItalic = False
    chatTxt.SelStrikeThru = False
    chatTxt.SelUnderline = False
    chatTxt.SelText = strUsername & ": "
    End If


   chatTxt.SelStart = Len(chatTxt.Text)
   chatTxt.SelBold = False
   Select Case vCode
        Case ResponseTypes.message
            chatTxt.SelColor = Colors.Chattext
        Case ResponseTypes.kick
            chatTxt.SelColor = Colors.Kicks
        Case ResponseTypes.Notice
            chatTxt.SelColor = Colors.Notice
            chatTxt.SelBold = True
            strMessage = "**" & strMessage
        Case ResponseTypes.Joins
            chatTxt.SelColor = Colors.Joins
            chatTxt.SelItalic = True
        Case ResponseTypes.Quits
            chatTxt.SelColor = Colors.Quits
            chatTxt.SelItalic = True
        Case ResponseTypes.commandRequest
            chatTxt.SelColor = Colors.Commands
            chatTxt.SelItalic = False
        Case Else
            chatTxt.SelColor = Colors.NormalText
    End Select
    chatTxt.SelFontName = "Arial"
    chatTxt.SelText = strMessage & vbCrLf
End Sub

Public Sub SetTextFocus()

End Sub
