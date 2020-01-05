VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Volleyball Assignments"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   615
      Left            =   6120
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   3120
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtPlayer 
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Team 2"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Team 1"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "New Player:"
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
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub List1_DragDrop(Source As Control, X As Single, Y As Single)
Dim strPlayer As String
strPlayer = txtPlayer.Text
List1.AddItem (strPlayer)
txtPlayer.Text = ""
End Sub

Private Sub List2_DragDrop(Source As Control, X As Single, Y As Single)
Dim strPlayer As String
strPlayer = txtPlayer.Text
List2.AddItem (strPlayer)
txtPlayer.Text = ""
End Sub
