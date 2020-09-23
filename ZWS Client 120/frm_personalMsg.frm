VERSION 5.00
Begin VB.Form frm_personalMsg 
   Caption         =   "Personal Message []"
   ClientHeight    =   3450
   ClientLeft      =   1815
   ClientTop       =   1785
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   5880
   Begin zws_client.pm_window cht 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3855
      _extentx        =   6800
      _extenty        =   5318
      caption         =   "Personal Msg System"
      text            =   ""
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frm_personalMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public UserID As Long
Public User
Private Sub cht_ChatSend(txt As String)
If isConnect Then
    frm_main.ws_main.SendData ResponseTypes.pm & Sep & Me.Tag & Sep _
    & Client.username & Sep & Client.UserID & Sep & txt & Sep & DateTime.Now
End If
cht.AddMsgToChat message, Client.username, txt
'personal message, only forward to the correct client
            'pm messages are formatted as follows:
            ' pm response code(10) & sep & targetUserId & sep & sourceUserName _
                & sep & source userid & sep & message & sep & timesent
' pm,targetUserId,sourceUserName,source userid,message,timesent
End Sub

Private Sub Form_Load()
cht.SetTextFocus
End Sub

Private Sub Form_Resize()
cht.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Me.Tag <> "" Then
    closeChatWithUSer (Me.Tag)
    Me.Tag = 0
End If
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub
