VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   Caption         =   "TEST"
   ClientHeight    =   5355
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   6240
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   5880
      TabIndex        =   20
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox TxtName 
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
      Left            =   1320
      TabIndex        =   18
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   4680
      TabIndex        =   14
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   615
      Left            =   2880
      TabIndex        =   13
      Top             =   4320
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
      Top             =   3480
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
      Top             =   2640
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
      Top             =   1800
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
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton CmdCal 
      Caption         =   "Calcuate"
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "(Name)"
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
      Left            =   120
      TabIndex        =   19
      Top             =   240
      Width           =   975
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
      Top             =   3480
      Width           =   735
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
      Top             =   3480
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
      Top             =   2520
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
      Top             =   3480
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
      Top             =   2760
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
      Top             =   1920
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
      Top             =   1200
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
      Top             =   3360
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
      Top             =   2640
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
      Top             =   1800
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
      Top             =   1080
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
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
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
Dim strFilename As String

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

Case Is <= 59
strGrade = "F"
End Select

'displaying letter grade
lblLGrade(I).Caption = strGrade

'finding the average grade percent
lblAveP.Caption = Val(lblAveP.Caption) + txtScore(I).Text / 4

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

Case Is <= 59
strAve = "F"
End Select

lblAveL.Caption = strAve

'correcting errors
If IsNumeric(txtScore(I).Text) = False Then
MsgBox "please enter only numbers"
txtScore(I).Text = ""
End If

Next

End Sub

Private Sub CmdClear_Click()
For I = 0 To 3
txtScore(I).Text = ""
lblLGrade(I).Caption = ""
lblAveP = ""
lblAveL = ""
Next
TxtName.Text = ""
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim strTest As String, strTest1 As String, strTest2 As String, strTest3 As String
Dim strGrade As String, strGrade1 As String, strGrade2 As String, strGrade3 As String
Dim strAveL As String
Dim strAveP As String
 dlgCommon.CancelError = True
dlgCommon.Filter = "Text File (*.txt) *.txt"
dlgCommon.DialogTitle = "Open Text Files"
dlgCommon.ShowOpen

strFilename = dlgCommon.FileName
'For Index = 0 To 3
    Open strFilename For Output As #1       'open the sequential file
    Write #1, TxtName.Text, txtScore(0).Text, txtScore(1).Text, txtScore(2).Text, txtScore(3).Text, lblLGrade(0).Caption, lblLGrade(1).Caption, lblLGrade(2).Caption, lblLGrade(3).Caption, lblAveL.Caption, lblAveP.Caption   ' write the record
    Close #1
    
    
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

Private Sub mnuOpen_Click()

Dim strTest As String, strTest1 As String, strTest2 As String, strTest3 As String
Dim strGrade As String, strGrade1 As String, strGrade2 As String, strGrade3 As String
Dim strAveL As String
Dim strAveP As String

'On Error GoTo Err
dlgCommon.CancelError = True

dlgCommon.Filter = "Text Files (*.txt) * *.txt"
dlgCommon.DialogTitle = "Open Text File"
dlgCommon.ShowOpen
strFilename = dlgCommon.FileName

dlgCommon.CancelError = True


Open strFilename For Input As #1
'Do While Not EOF(1)
Input #1, strTest, strTest1, strTest2, strTest3, strGrade, strGrade1, strGrade2, strGrade3, strinfoget, strAveL, strAveP

'On Error GoTo Err
'TxtName.Text = TxtName.Text & strTest
txtScore(0).Text = txtScore(0).Text & strTest1
txtScore(1).Text = txtScore(1).Text & strTest2
txtScore(2).Text = txtScore(2).Text & strTest3
txtScore(3).Text = txtScore(3).Text & strGrade
lblLGrade(0).Caption = lblLGrade(0).Caption & strGrade1
lblLGrade(1).Caption = lblLGrade(1).Caption & strGrade2
lblLGrade(2).Caption = lblLGrade(2).Caption & strGrade3
lblLGrade(3).Caption = lblLGrade(3).Caption & strinfoget
lblAveL.Caption = lblAveL.Caption & strAveL
lblAveP.Caption = lblAveP.Caption & strAveP
TxtName.Text = TxtName.Text & strTest

Close #1
'Loop

End Sub
