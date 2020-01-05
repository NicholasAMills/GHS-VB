VERSION 5.00
Begin VB.Form frmRPS 
   Caption         =   "Rock Paper Scissors"
   ClientHeight    =   9690
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   12075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   10
      Top             =   9000
      Width           =   1815
   End
   Begin VB.CommandButton CmdFight 
      Caption         =   "FIGHT!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   8
      Top             =   8280
      Width           =   1935
   End
   Begin VB.CheckBox ChkScore 
      Caption         =   "Score/Rules"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   5
      Top             =   8280
      Width           =   1935
   End
   Begin VB.OptionButton OptGun 
      Caption         =   "Gun"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   6600
      Width           =   975
   End
   Begin VB.OptionButton OptMan 
      Caption         =   "Man"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   4680
      Width           =   975
   End
   Begin VB.OptionButton OptSlender 
      Caption         =   "Slenderman"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgComGun 
      Height          =   1335
      Left            =   9240
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Image imgComMan 
      Height          =   1800
      Left            =   9480
      Picture         =   "Form1.frx":3A21
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1395
   End
   Begin VB.Image imgComSlender 
      Height          =   1815
      Left            =   9360
      Picture         =   "Form1.frx":4FDA
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblComp 
      Caption         =   "COMPUTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Slenderman style"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.Image imgPGun 
      Height          =   1335
      Left            =   2160
      Picture         =   "Form1.frx":6405
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Image imgPMan 
      Height          =   1800
      Left            =   2160
      Picture         =   "Form1.frx":9E26
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1395
   End
   Begin VB.Image imgPSlender 
      Height          =   1815
      Left            =   2040
      Picture         =   "Form1.frx":B3DF
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "ROCK PAPER SCISSORS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmRPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ChkScore_Click()
'displaying the score form with a checkbox
If ChkScore = vbChecked Then
    FrmScore.Show
Else
    FrmScore.Hide
End If
End Sub

Private Sub CmdFight_Click()
Dim strComChoice As String
Dim strUserChoice As String




'choice Slender
If OptSlender.Value = True Then
    strUserChoice = "1"
    imgPSlender.Visible = True
Else
    imgPSlender.Visible = False
End If

'choice Man
If OptMan.Value = True Then
    strUserChoice = "2"
    imgPMan.Visible = True
Else
    imgPMan.Visible = False
End If

'choice Gun
If OptGun.Value = True Then
    strUserChoice = "3"
    imgPGun.Visible = True
Else
    imgPGun.Visible = False
End If
Randomize
'randomizing computer values
strCompChoice = Int((3 * Rnd) + 1)

'com choice Slender
If strCompChoice = 1 Then
    imgComSlender.Visible = True
Else
    imgComSlender.Visible = False
End If

'com choice Man
If strCompChoice = 2 Then
    imgComMan.Visible = True
Else
    imgComMan.Visible = False
End If

'com choice Gun
If strCompChoice = 3 Then
    imgComGun.Visible = True
Else
    imgComGun.Visible = False
End If

'adding score if user wins
If strUserChoice = "1" And strCompChoice = "2" Then
    intCountYou = intCountYou + 1
    MsgBox "You Win!"
Else
If strUserChoice = "2" And strCompChoice = "3" Then
    intCountYou = intCountYou + 1
    MsgBox "You Win!"
Else
If strUserChoice = "3" And strCompChoice = "1" Then
    intCountYou = intCountYou + 1
    MsgBox "You Win!"
End If
End If
End If

'adding score if computer wins
If strCompChoice = "1" And strUserChoice = "2" Then
    intCountCom = intCountCom + 1
    MsgBox "You Lost"
Else
If strCompChoice = "2" And strUserChoice = "3" Then
    intCountCom = intCountCom + 1
    MsgBox "You Lost"
Else
If strCompChoice = "3" And strUserChoice = "1" Then
    intCountCom = intCountCom + 1
    MsgBox "You Lost"
End If
End If
End If

If strUserChoice = "1" And strCompChoice = "1" Or strUserChoice = "2" And strCompChoice = "2" Or strUserChoice = "3" And strCompChoice = "3" Then
    MsgBox "There was a tie"
End If

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
'asking user to input their name
Dim strName As String
strName = UCase(InputBox("Please enter your name"))

If strName = Cancel Then
    Form_Unload (0)
End If


'displaying user's name
lblName.Caption = strName

End Sub

Private Sub Form_Unload(Cancel As Integer)
'asking user if they want to exit
Dim bytMessage As Byte
Const conButtons As Integer = vbYesNo + vbDefaultButton2 + vbQuestion + vbApplicationModal
bytMessage = MsgBox("Do you want to quit?", conButtons, "Exit")

If bytMessage = vbYes Then
    End
Else
    Cancel = 1
    frmRPS.Show
End If

End Sub


Private Sub OptGun_Click()
If OptGun.Value = True Then
imgPGun.Visible = True
imgPMan.Visible = False
imgPSlender.Visible = False
End If
End Sub

Private Sub OptMan_Click()
If OptMan.Value = True Then
imgPMan.Visible = True
imgPGun.Visible = False
imgPSlender.Visible = False
End If
End Sub

Private Sub OptSlender_Click()
If OptSlender.Value = True Then
imgPSlender.Visible = True
imgPGun.Visible = False
imgPMan.Visible = False
End If
End Sub
