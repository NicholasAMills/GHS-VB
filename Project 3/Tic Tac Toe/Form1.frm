VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tic-Tac-Toe"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Quit"
      Height          =   615
      Left            =   7200
      TabIndex        =   30
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label7 
      Height          =   495
      Left            =   1560
      TabIndex        =   37
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label6 
      Height          =   495
      Left            =   600
      TabIndex        =   36
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "P2"
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
      Left            =   1560
      TabIndex        =   35
      Top             =   3000
      Width           =   375
   End
   Begin VB.Line Line6 
      X1              =   600
      X2              =   2160
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   1320
      X2              =   1320
      Y1              =   3000
      Y2              =   4560
   End
   Begin VB.Label Label4 
      Caption         =   "P1"
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
      Left            =   600
      TabIndex        =   34
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   33
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblCount 
      Height          =   855
      Left            =   240
      TabIndex        =   32
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label lblPlayer2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6000
      TabIndex        =   31
      Top             =   1320
      Width           =   105
   End
   Begin VB.Label Label2 
      Caption         =   "VS "
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
      Left            =   4080
      TabIndex        =   29
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblPlayer1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      TabIndex        =   28
      Top             =   1320
      Width           =   105
   End
   Begin VB.Label Label1 
      Caption         =   "TIC-TAC-TOE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   27
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblOpt9 
      Caption         =   "9"
      Height          =   495
      Left            =   5640
      TabIndex        =   26
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label lblOpt8 
      Caption         =   "8"
      Height          =   495
      Left            =   4440
      TabIndex        =   25
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label lblOpt7 
      Caption         =   "7"
      Height          =   495
      Left            =   3360
      TabIndex        =   24
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label lblOpt6 
      Caption         =   "6"
      Height          =   615
      Left            =   5640
      TabIndex        =   23
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label lblOpt5 
      Caption         =   "5"
      Height          =   615
      Left            =   4440
      TabIndex        =   22
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblOpt4 
      Caption         =   "4"
      Height          =   615
      Left            =   3360
      TabIndex        =   21
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblOpt3 
      Caption         =   "3"
      Height          =   615
      Left            =   5520
      TabIndex        =   20
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblOpt2 
      Caption         =   "2"
      Height          =   615
      Left            =   4440
      TabIndex        =   19
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblOpt1 
      Caption         =   "1"
      Height          =   615
      Left            =   3360
      TabIndex        =   18
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblO7 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   17
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblO8 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   16
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblO9 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   15
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblO1 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblO3 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblO4 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblO5 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblO6 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   10
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblX2 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblX3 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblX4 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblX5 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblX6 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblX7 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblX8 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblX9 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblO2 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblX1 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line Line4 
      X1              =   3240
      X2              =   6360
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line3 
      X1              =   3240
      X2              =   6360
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line2 
      X1              =   5400
      X2              =   5400
      Y1              =   2520
      Y2              =   5160
   End
   Begin VB.Line Line1 
      X1              =   4320
      X2              =   4320
      Y1              =   2520
      Y2              =   5160
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intCount As Integer

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim strP1 As String
Dim strP2 As String
'entering Player 1's name
strP1 = UCase(InputBox("Enter Player 1's name"))

If strP1 = Cancel Then
    Form_Unload (0)
End If

lblPlayer1.Caption = strP1

'entering Player 2's name
strP2 = UCase(InputBox("Enter Player 2's name"))
lblPlayer2.Caption = strP2

End Sub

Private Sub Form_Unload(Cancel As Integer)
'asking user if they want to exit
Dim bytMessage As Byte
Const conButtons As Integer = vbYesNo + vbDefaultButton2 + vbQuestion
bytMessage = MsgBox("Do you want to quit?", conButtons, "Exit")

If bytMessage = vbYes Then
    End
Else
    Cancel = 1
    Form1.Show
End If
End Sub

Private Sub lblOpt1_Click()
'displaying if intCount is even it displays X and if odd displays O
intCount = Val(lblCount)
intCount = intCount + 1
If intCount = 0 Or intCount = 2 Or intCount = 4 Or intCount = 6 Or incount = 8 Then
    lblX1.Visible = True
    lblOpt1.Visible = False
Else
If intCount = 1 Or intCount = 3 Or intCount = 5 Or intCount = 7 Or intCount = 9 Then
    lblO1.Visible = True
    lblOpt1.Visible = False
End If
End If
lblCount.Caption = intCount
End Sub

Private Sub lblOpt2_Click()
intCount = Val(lblCount)
intCount = intCount + 1
If intCount = 0 Or intCount = 2 Or intCount = 4 Or intCount = 6 Or incount = 8 Then
    lblX2.Visible = True
    lblOpt2.Visible = False
Else
If intCount = 1 Or intCount = 3 Or intCount = 5 Or intCount = 7 Or intCount = 9 Then
    lblO2.Visible = True
    lblOpt2.Visible = False
End If
End If
lblCount.Caption = intCount
End Sub

Private Sub lblOpt3_Click()
intCount = Val(lblCount)
intCount = intCount + 1
If intCount = 0 Or intCount = 2 Or intCount = 4 Or intCount = 6 Or incount = 8 Then
    lblX3.Visible = True
    lblOpt3.Visible = False
Else
If intCount = 1 Or intCount = 3 Or intCount = 5 Or intCount = 7 Or intCount = 9 Then
    lblO3.Visible = True
    lblOpt3.Visible = False
End If
End If
lblCount.Caption = intCount
End Sub

Private Sub lblOpt4_Click()
intCount = Val(lblCount)
intCount = intCount + 1
If intCount = 0 Or intCount = 2 Or intCount = 4 Or intCount = 6 Or incount = 8 Then
    lblX4.Visible = True
    lblOpt4.Visible = False
Else
If intCount = 1 Or intCount = 3 Or intCount = 5 Or intCount = 7 Or intCount = 9 Then
    lblO4.Visible = True
    lblOpt4.Visible = False
End If
End If
lblCount.Caption = intCount
End Sub

Private Sub lblOpt5_Click()
intCount = Val(lblCount)
intCount = intCount + 1
If intCount = 0 Or intCount = 2 Or intCount = 4 Or intCount = 6 Or incount = 8 Then
    lblX5.Visible = True
    lblOpt5.Visible = False
Else
If intCount = 1 Or intCount = 3 Or intCount = 5 Or intCount = 7 Or intCount = 9 Then
    lblO5.Visible = True
    lblOpt5.Visible = False
End If
End If
lblCount.Caption = intCount
End Sub

Private Sub lblOpt6_Click()
intCount = Val(lblCount)
intCount = intCount + 1
If intCount = 0 Or intCount = 2 Or intCount = 4 Or intCount = 6 Or incount = 8 Then
    lblX6.Visible = True
    lblOpt6.Visible = False
Else
If intCount = 1 Or intCount = 3 Or intCount = 5 Or intCount = 7 Or intCount = 9 Then
    lblO6.Visible = True
    lblOpt6.Visible = False
End If
End If
lblCount.Caption = intCount
End Sub

Private Sub lblOpt7_Click()
intCount = Val(lblCount)
intCount = intCount + 1
If intCount = 0 Or intCount = 2 Or intCount = 4 Or intCount = 6 Or incount = 8 Then
    lblX7.Visible = True
    lblOpt7.Visible = False
Else
If intCount = 1 Or intCount = 3 Or intCount = 5 Or intCount = 7 Or intCount = 9 Then
    lblO7.Visible = True
    lblOpt7.Visible = False
End If
End If
lblCount.Caption = intCount
End Sub

Private Sub lblOpt8_Click()
intCount = Val(lblCount)
intCount = intCount + 1

If intCount = 0 Or intCount = 2 Or intCount = 4 Or intCount = 6 Or incount = 8 Then
    lblX8.Visible = True
    lblOpt8.Visible = False
Else
If intCount = 1 Or intCount = 3 Or intCount = 5 Or intCount = 7 Or intCount = 9 Then
    lblO8.Visible = True
    lblOpt8.Visible = False
End If
End If
lblCount.Caption = intCount
End Sub

Private Sub lblOpt9_Click()
intCount = Val(lblCount)
intCount = intCount + 1
If intCount = 0 Or intCount = 2 Or intCount = 4 Or intCount = 6 Or incount = 8 Then
    lblX9.Visible = True
    lblOpt9.Visible = False
Else
If intCount = 1 Or intCount = 3 Or intCount = 5 Or intCount = 7 Or intCount = 9 Then
    lblO9.Visible = True
    lblOpt9.Visible = False
End If
End If
lblCount.Caption = intCount
End Sub
