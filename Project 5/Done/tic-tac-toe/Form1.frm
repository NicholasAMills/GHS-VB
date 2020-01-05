VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   4410
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgBlank 
      Height          =   735
      Left            =   7320
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblyscore 
      Caption         =   "0"
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblxscore 
      Caption         =   "0"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblPlayer2 
      Caption         =   "Player 2"
      Height          =   255
      Left            =   5640
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblplayer1 
      Caption         =   "Player 1"
      Height          =   255
      Left            =   4320
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Line Line10 
      X1              =   5400
      X2              =   5520
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line9 
      X1              =   5400
      X2              =   5520
      Y1              =   960
      Y2              =   600
   End
   Begin VB.Line Line8 
      X1              =   5520
      X2              =   5400
      Y1              =   960
      Y2              =   600
   End
   Begin VB.Line Line7 
      X1              =   4320
      X2              =   6480
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line6 
      X1              =   5520
      X2              =   5520
      Y1              =   960
      Y2              =   3120
   End
   Begin VB.Line Line5 
      X1              =   5400
      X2              =   5400
      Y1              =   960
      Y2              =   3120
   End
   Begin VB.Image Image3 
      DragMode        =   1  'Automatic
      Height          =   855
      Left            =   5760
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   855
   End
   Begin VB.Image Image2 
      DragMode        =   1  'Automatic
      Height          =   855
      Left            =   4320
      Picture         =   "Form1.frx":AF77
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   8
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   7
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   6
      Left            =   240
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   5
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   4
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   3
      Left            =   240
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   2
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   600
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   1
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   600
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   0
      Left            =   240
      Stretch         =   -1  'True
      Top             =   600
      Width           =   855
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   3840
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   3840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      X1              =   2640
      X2              =   2640
      Y1              =   480
      Y2              =   3840
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   1320
      Y1              =   480
      Y2              =   3840
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuNewPlayer 
         Caption         =   "New Player"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim turn As Boolean
Dim counter As Single

Private Sub Form_Load()
Dim strp1 As String
Dim strp2 As String
turn = False
Image3.Enabled = False
'entering player names
strp1 = UCase(InputBox("Enter Player 1's name"))
'If strp1 = Cancel Then
    'form_unload
'End If
lblplayer1.Caption = strp1
strp2 = UCase(InputBox("Enter player 2'S name"))
lblPlayer2.Caption = strp2

End Sub

Private Sub Image1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
counter = counter + 1
Image1(Index).Picture = Source.Picture




'turn determiner
If turn = True Then
    Image2.Enabled = True
    Image3.Enabled = False
    turn = False
Else
    Image2.Enabled = False
    Image3.Enabled = True
    turn = True
End If







'gnome
Dim x1 As String
Dim x2 As String
Dim x3 As String
Dim x4 As String
Dim x5 As String
Dim x6 As String
Dim x7 As String
Dim x8 As String
Dim x9 As String

x1 = Image1(0).Picture
x2 = Image1(1).Picture
x3 = Image1(2).Picture
x4 = Image1(3).Picture
x5 = Image1(4).Picture
x6 = Image1(5).Picture
x7 = Image1(6).Picture
x8 = Image1(7).Picture
x9 = Image1(8).Picture



'Osterage
Dim y1 As String
Dim y2 As String
Dim y3 As String
Dim y4 As String
Dim y5 As String
Dim y6 As String
Dim y7 As String
Dim y8 As String
Dim y9 As String

y1 = Image1(0).Picture
y2 = Image1(1).Picture
y3 = Image1(2).Picture
y4 = Image1(3).Picture
y5 = Image1(4).Picture
y6 = Image1(5).Picture
y7 = Image1(6).Picture
y8 = Image1(7).Picture
y9 = Image1(8).Picture



Static yscore As Single
Static xscore As Single



If Image3.Picture = y1 And Image3.Picture = y2 And Image3.Picture = y3 Then
MsgBox "You win!"
yscore = yscore + 1
mnuNewGame_Click
Else
If Image3.Picture = y4 And Image3.Picture = y5 And Image3.Picture = y6 Then
MsgBox "You won play again!"
yscore = yscore + 1
mnuNewGame_Click
Else
If Image3.Picture = y7 And Image3.Picture = y8 And Image3.Picture = y9 Then
MsgBox "Success!"
yscore = yscore + 1
mnuNewGame_Click
Else
If Image3.Picture = y1 And Image3.Picture = y4 And Image3.Picture = y7 Then
MsgBox "You played the game right!"
yscore = yscore + 1
mnuNewGame_Click
Else
If Image3.Picture = y2 And Image3.Picture = y5 And Image3.Picture = y8 Then
MsgBox "Winner!"
yscore = yscore + 1
mnuNewGame_Click
Else
If Image3.Picture = y3 And Image3.Picture = y6 And Image3.Picture = y9 Then
MsgBox "You got the big W!"
yscore = yscore + 1
mnuNewGame_Click
Else
If Image3.Picture = y1 And Image3.Picture = y5 And Image3.Picture = y9 Then
MsgBox "Wow... took ya long enough.."
yscore = yscore + 1
mnuNewGame_Click
Else
If Image3.Picture = y3 And Image3.Picture = y5 And Image3.Picture = y7 Then
MsgBox "Nice, but win faster next time!"
yscore = yscore + 1
mnuNewGame_Click
Else
If Image2.Picture = x1 And Image2.Picture = x2 And Image2.Picture = x3 Then
MsgBox "Winner!"
xscore = xscore + 1
mnuNewGame_Click
Else
If Image2.Picture = x4 And Image2.Picture = x5 And Image2.Picture = x6 Then
MsgBox "Win!"
xscore = xscore + 1
mnuNewGame_Click
Else
If Image2.Picture = x7 And Image2.Picture = x8 And Image2.Picture = x9 Then
MsgBox "You played the game right!"
xscore = xscore + 1
mnuNewGame_Click
Else
If Image2.Picture = x1 And Image2.Picture = x4 And Image2.Picture = x7 Then
MsgBox "You did it!"
xscore = xscore + 1
mnuNewGame_Click
Else
If Image2.Picture = x2 And Image2.Picture = x5 And Image2.Picture = x8 Then
MsgBox "Success!!"
xscore = xscore + 1
mnuNewGame_Click
Else
If Image2.Picture = x3 And Image2.Picture = x6 And Image2.Picture = x9 Then
MsgBox "You got the big W!"
xscore = xscore + 1
mnuNewGame_Click
Else
If Image2.Picture = x1 And Image2.Picture = x5 And Image2.Picture = x9 Then
MsgBox "Wow... took ya long enough.."
xscore = xscore + 1
mnuNewGame_Click
Else
If Image2.Picture = x3 And Image2.Picture = x5 And Image2.Picture = x7 Then
MsgBox "Nice, but win faster next time"
xscore = xscore + 1
mnuNewGame_Click
Else

'tie
If counter = 9 Then
MsgBox "You should be ashamed!!!"
mnuNewGame_Click
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If

lblxscore.Caption = xscore
lblyscore.Caption = yscore

End Sub





Private Sub mnuNewGame_Click()
counter = 0
For i = 0 To 8
Image1(i).Picture = imgBlank.Picture
Next

End Sub

Private Sub mnuNewPlayer_Click()
counter = 0
new_game
mnuNewGame_Click
lblxscore.Caption = 0
lblyscore.Caption = 0
xscore = 0
yscore = 0
End Sub

Private Sub mnuQuit_Click()
Dim snganswer As Single
snganswer = MsgBox("Do you really want to quit?", vbYesNo + vbQuestion, "Exit")
If snganswer = vbYes Then
    'user does want to exit.
    Unload Me
Else
    Form1.Show
    Cancel = 1 'sets the local variable back to 1 because
End If
End Sub

Private Sub new_game()
'entering player names
strp1 = UCase(InputBox("Enter Player 1's name"))
'If strp1 = Cancel Then
    'form_unload
'End If
lblplayer1.Caption = strp1
strp2 = UCase(InputBox("Enter player 2'S name"))
lblPlayer2.Caption = strp2
Image3.Enabled = True
End Sub
