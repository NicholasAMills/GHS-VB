VERSION 5.00
Begin VB.Form forumFirst 
   Caption         =   "My First Program"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMDexit 
      Caption         =   "E&xit"
      Height          =   1095
      Left            =   9720
      TabIndex        =   3
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton CMDsecond 
      Caption         =   "Press &Second"
      Height          =   1335
      Left            =   9600
      TabIndex        =   2
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton CMDfirst 
      Caption         =   "Press &First"
      Height          =   1215
      Left            =   9600
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Image ImgRobber 
      Height          =   3225
      Left            =   4080
      Picture         =   "forumFirst.frx":0000
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   3660
   End
   Begin VB.Image ImgDino 
      Height          =   4050
      Left            =   120
      Picture         =   "forumFirst.frx":19DB
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   3195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "My First  Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2835
   End
End
Attribute VB_Name = "forumFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_Change()

End Sub

Private Sub CMDexit_Click()
Unload Me
End Sub

Private Sub CMDfirst_Click()
ImgDino.Visible = False
ImgRobber.Visible = True
End Sub

Private Sub CMDsecond_Click()
ImgRobber.Visible = False
ImgDino.Visible = True
End Sub
