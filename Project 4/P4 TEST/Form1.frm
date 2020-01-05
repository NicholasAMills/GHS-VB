VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TEST"
   ClientHeight    =   4665
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   4680
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   615
      Left            =   2880
      TabIndex        =   13
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtScore 
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
      Index           =   3
      Left            =   1680
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtScore 
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
      Index           =   2
      Left            =   1680
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtScore 
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
      Index           =   1
      Left            =   1680
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtScore 
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
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton CmdCal 
      Caption         =   "Calcuate"
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblAveL 
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
      Left            =   4560
      TabIndex        =   17
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblAveP 
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
      Left            =   5760
      TabIndex        =   16
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Average Grade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   15
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblLGrade 
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
      Index           =   3
      Left            =   3120
      TabIndex        =   12
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblLGrade 
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
      Index           =   2
      Left            =   3120
      TabIndex        =   11
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblLGrade 
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
      Index           =   1
      Left            =   3120
      TabIndex        =   10
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblLGrade 
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
      Index           =   0
      Left            =   3120
      TabIndex        =   9
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Test 4"
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
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Test 3"
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
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Test 2"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Test 1"
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
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuCalc 
         Caption         =   "Calculate"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCal_Click()
Dim strGrade As String
Dim intPer As Integer
Dim strAve As String

For I = 0 To 3
intGrade = Val(lblLGrade(I).Caption)
intPer = Val(txtScore(I).Text)
strAve = Val(lblAveL.Caption)
'finding the letter grade
Select Case intPer
Case Is >= 90
strGrade = "A"
Case Is >= 80 And intGrade <= 89
strGrade = "B"

Case Is >= 70 And intGrade <= 79
strGrade = "C"

Case Is >= 60 And intGrade <= 69
strGrade = "D"

Case Is >= 59
strGrade = "F"
End Select

'displaying letter grade
lblLGrade(I).Caption = strGrade

'finding the average grade percent
lblAveP.Caption = Val(lblAveP.Caption) + Val(txtScore(I).Text) / 4

'finding the average letter grade
Select Case lblAveP.Caption
Case Is >= 90
strAve = "A"
Case Is >= 80 And strAve <= 89
strAve = "B"

Case Is >= 70 And strAve <= 79
strAve = "C"

Case Is >= 60 And strAve <= 69
strAve = "D"

Case Is >= 59
strAve = "F"
End Select

lblAveL.Caption = strAve

'correcting errors
If IsNumeric(txtScore(Index).Text) = False Then
MsgBox "please enter only numbers"
txtScore(Index).Text = 0
End If

Next

End Sub

Private Sub CmdClear_Click()
For I = 0 To 3
txtScore(I).Text = 0
lblLGrade(I).Caption = ""
lblAveP = ""
lblAveL = ""
Next
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub mnuCalc_Click()
CmdCal_Click
End Sub

Private Sub mnuClear_Click()
CmdClear_Click
End Sub

Private Sub mnuExit_Click()
CmdExit_Click
End Sub
