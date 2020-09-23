VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl cht_Window 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin RichTextLib.RichTextBox msgb 
      Height          =   735
      Left            =   1500
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"cht_Window.ctx":0000
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   180
      ScaleHeight     =   3195
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin MSComctlLib.ImageList imgUs 
         Left            =   540
         Top             =   1920
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cht_Window.ctx":0077
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cht_Window.ctx":03C9
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cht_Window.ctx":071B
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lstUsers 
         Height          =   1215
         Left            =   3180
         TabIndex        =   4
         Top             =   1140
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         Icons           =   "imgUs"
         SmallIcons      =   "imgUs"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Users"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1620
         TabIndex        =   3
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
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"cht_Window.ctx":0A6D
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
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   3240
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   540
         X2              =   3780
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label cap 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
      Begin VB.Image im 
         Height          =   225
         Left            =   420
         Picture         =   "cht_Window.ctx":0AE4
         Stretch         =   -1  'True
         Top             =   600
         Width           =   3480
      End
   End
End
Attribute VB_Name = "cht_Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event ChatSend(txt As String)
Public Event UserClick(strUser As String)

Private Sub chatTxt_Change()
chatTxt.SelLength = Len(chatTxt.Text)
End Sub

Private Sub lstUsers_DblClick()
On Error Resume Next
If lstUsers.SelectedItem.Text <> "" Then
    RaiseEvent UserClick(lstUsers.SelectedItem.Text)
End If
End Sub

Private Sub lstUsers_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
frm_main.SelUser = lstUsers.SelectedItem.Text
End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
Case 2
    PopupMenu frm_main.mnuPM
End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If LTrim(RTrim(Text1.Text)) <> "" Then
        RaiseEvent ChatSend(Text1.Text)
        DoEvents
        KeyAscii = 0
        Text1.Text = ""
    End If
End If
'KeyAscii = 0
End Sub


Public Sub UpdateBlks()
For i = 1 To lstUsers.ListItems.Count
If IsUserBlocked(lstUsers.ListItems(i).Text) = True Then
    lstUsers.ListItems(i).SmallIcon = 2
Else
     lstUsers.ListItems(i).SmallIcon = 3
End If
Next
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
cap.Caption = PropBag.ReadProperty("Caption", "Caption")
chatTxt.Text = PropBag.ReadProperty("Text", "ChatTxt")
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
Picture1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
im.Move 0, 0, Picture1.ScaleWidth - lstUsers.Width
cap.Move 0 + 30, 0 + 15, Picture1.ScaleWidth - lstUsers.Width
Line1.Y1 = im.Height
Line1.Y2 = im.Height
Line1.X1 = 0
Line1.X2 = Picture1.ScaleWidth
Line2.Y1 = im.Height + 15
Line2.Y2 = im.Height + 15
Line2.X1 = 0
Line2.X2 = Picture1.ScaleWidth
chatTxt.Move 0, im.Height + 30, Picture1.ScaleWidth - lstUsers.Width, Picture1.ScaleHeight - Text1.Height - im.Height - 30
Text1.Move 0, chatTxt.Height + im.Height + 30 + 15, Picture1.Width - lstUsers.Width - 60
lstUsers.Move chatTxt.Width + 15, -15, 1275, Picture1.ScaleHeight + 30
lstUsers.ColumnHeaders.Item(1).Width = 1275
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

Public Sub AddUser(userstring As String)
If IsUserBlocked(userstring) = True Then
    lstUsers.ListItems.Add , , userstring, , 2
Else
    lstUsers.ListItems.Add , , userstring
End If
End Sub

Public Sub ClearUsers()
lstUsers.ListItems.Clear
End Sub

Public Sub SetSendText(txt As String)
    Text1.Text = txt
End Sub

Public Function GetSendText() As String
    GetSendText = Text1.Text
End Function

Public Sub elitetxtme()

Dim tstr As String
    For i = 0 To Len(Text1.Text)
        Select Case LCase(Right(Left(Text1.Text, i), 1))
            Case "a"
                tstr = tstr & "Å"
            Case "b"
                tstr = tstr & "ß"
            Case "c"
                 tstr = tstr & "©"
            Case "d"
                tstr = tstr & "Ð"
            Case "e"
                tstr = tstr & "ê"
            Case "f"
                tstr = tstr & "ƒ"
            Case "g"
                tstr = tstr & "9"
            Case "h"
                tstr = tstr & "|-|"
            Case "i"
                tstr = tstr & "î"
            Case "j"
                 tstr = tstr & "j"
            Case "k"
                tstr = tstr & "k"
            Case "l"
                tstr = tstr & "£"
            Case "m"
                tstr = tstr & "m"
            Case "n"
                tstr = tstr & "Ñ"
            Case "o"
                tstr = tstr & "Ó"
            Case "p"
                tstr = tstr & "Þ"
            Case "q"
                tstr = tstr & "q"
            Case "r"
                tstr = tstr & "®"
            Case "s"
                tstr = tstr & "§"
            Case "t"
                tstr = tstr & "±"
            Case "u"
                tstr = tstr & "µ"
            Case "v"
                tstr = tstr & "v"
            Case "w"
                tstr = tstr & "w"
            Case "x"
                tstr = tstr & "×"
            Case "y"
                tstr = tstr & "ÿ"
            Case "z"
                tstr = tstr & "ž"
            Case "0"
                tstr = tstr & "°"
            Case "1"
                tstr = tstr & "¹"
            Case "2"
                tstr = tstr & "²"
            Case "3"
                tstr = tstr & "³"
            Case "?"
                tstr = tstr & "¿"
            Case Else
                tstr = tstr & Right(Left(Text1.Text, i), 1)
        End Select
    Next
    Text1.Text = tstr
    
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
            chatTxt.SelColor = Colors.PMText
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

Public Sub RefBlockList()
lstUsers.ListItems.Clear
For i = 1 To lstUsers.ListItems.Count
    AddUser (lstUsers.ListItems(i).Text)
Next
End Sub
