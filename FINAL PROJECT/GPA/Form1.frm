VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form GPACalculator 
   Caption         =   "GPA"
   ClientHeight    =   5970
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   6000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtCredits 
      Height          =   495
      Index           =   6
      Left            =   3000
      TabIndex        =   38
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox txtCredits 
      Height          =   495
      Index           =   5
      Left            =   3000
      TabIndex        =   37
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox txtCredits 
      Height          =   495
      Index           =   4
      Left            =   3000
      TabIndex        =   36
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtCredits 
      Height          =   495
      Index           =   3
      Left            =   3000
      TabIndex        =   35
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtCredits 
      Height          =   495
      Index           =   2
      Left            =   3000
      TabIndex        =   34
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtCredits 
      Height          =   495
      Index           =   1
      Left            =   3000
      TabIndex        =   33
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtGrade 
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
      Index           =   7
      Left            =   13080
      TabIndex        =   32
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCredits 
      Height          =   495
      Index           =   7
      Left            =   12960
      TabIndex        =   31
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCredits 
      Height          =   495
      Index           =   0
      Left            =   3000
      TabIndex        =   30
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Audit"
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
      Left            =   4200
      TabIndex        =   29
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Audit"
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
      Left            =   4200
      TabIndex        =   28
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Audit"
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
      Left            =   4200
      TabIndex        =   27
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Audit"
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
      Left            =   4200
      TabIndex        =   26
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Audit"
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
      Left            =   4200
      TabIndex        =   25
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Audit"
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
      Left            =   4200
      TabIndex        =   24
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Audit"
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
      Left            =   4200
      TabIndex        =   23
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
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
      Left            =   7440
      TabIndex        =   20
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtGrade 
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
      Index           =   6
      Left            =   1680
      TabIndex        =   19
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox txtGrade 
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
      Index           =   5
      Left            =   1680
      TabIndex        =   18
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox txtGrade 
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
      Index           =   4
      Left            =   1680
      TabIndex        =   17
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtGrade 
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
      TabIndex        =   16
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtGrade 
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
      TabIndex        =   15
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtGrade 
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
      TabIndex        =   14
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtGrade 
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
      TabIndex        =   13
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Do not exits ----->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   39
      Top             =   8400
      Width           =   1935
   End
   Begin VB.Label Label19 
      Caption         =   "Credits"
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
      Left            =   3000
      TabIndex        =   22
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Grade"
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
      Left            =   1680
      TabIndex        =   21
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Class 4:"
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
      TabIndex        =   12
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Class 5:"
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
      TabIndex        =   11
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Class 6:"
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
      TabIndex        =   10
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Class 7:"
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
      TabIndex        =   9
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Class 1:"
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
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Class 3:"
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
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Class 2:"
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
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "GPA:"
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
      Left            =   6960
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label lblGPA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblTotCredits 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Total Credits:"
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
      Left            =   6960
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblClass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Classes Taken: "
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
      Left            =   6960
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "New Grade/Clear"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save/Save as"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
   End
End
Attribute VB_Name = "GPACalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strfilename As String

Private Sub Command1_Click()
Dim strGradeClass As String
Dim strCreditClass As String
Dim intCount As Integer
Dim sngTotal As Single
Dim sngGPA(0 To 6) As Single
Dim sngTotal2 As Single
Dim sngGPA1 As Single
sngTotal2 = 0 'resetting to zero every time the user hits "calculate"
sngTotal = 0
sngCount = 0
For I = 0 To 6
sngGPA(I) = Val(txtGrade(I).Text) 'assigning variable to txtGrade.text

If txtGrade(I).Text = "a" Then     'if statements for GPA (lowercase only)
sngGPA(I) = 4
Else
If txtGrade(I).Text = "a-" Then
sngGPA(I) = 3.67
Else
If txtGrade(I).Text = "b+" Then
sngGPA(I) = 3.33
Else
If txtGrade(I).Text = "b" Then
sngGPA(I) = 3
Else
If txtGrade(I).Text = "b-" Then
sngGPA(I) = 2.67
Else
If txtGrade(I).Text = "c+" Then
sngGPA(I) = 2.33
Else
If txtGrade(I).Text = "c" Then
sngGPA(I) = 2
Else
If txtGrade(I).Text = "c-" Then
sngGPA(I) = 1.67
Else
If txtGrade(I).Text = "d+" Then
sngGPA(I) = 1.33
Else
If txtGrade(I).Text = "d" Then
sngGPA(I) = 1
Else
If txtGrade(I).Text = "d-" Then
sngGPA(I) = 0.67
Else
If txtGrade(I).Text = "f" Then
sngGPA(I) = 0

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
sngTotal2 = sngTotal2 + sngGPA(I)

If txtGrade(I).Text = "a" Or txtGrade(I).Text = "a-" Or txtGrade(I).Text = "b+" Or txtGrade(I).Text = "b" Or txtGrade(I).Text = "b-" Or txtGrade(I).Text = "c+" Or txtGrade(I).Text = "c" Or txtGrade(I).Text = "c-" Or txtGrade(I).Text = "d+" Or txtGrade(I).Text = "d" Or txtGrade(I).Text = "d-" Or txtGrade(I).Text = "f" Then
intCountClass = intCountClass + 1 'adding one for each class
Else
intCountClass = intCountClass + 0 'ignoring other letters/numbers
End If



lblTotCredits.Caption = sngTotal  'number of credits

If txtGrade(I).Enabled = True Then
'dividing the GPA amount by the number of classes to get the GPA
sngTotal = sngTotal + Val(txtCredits(I).Text)
End If


lblTotCredits.Caption = sngTotal
lblClass.Caption = Val(intCountClass) 'counting the number of classes


'Select Case txtGrade(Index).Text
'Case Is <> "a", "a-", "b+", "b", "b-", "c+", "c", "c-", "d+", "d+", "d", "d-", "f"
'MsgBox "Please enter your letter grade only."
'End Select
Next
sngGPA1 = sngTotal2 / intCountClass
lblGPA.Caption = Format(sngGPA1, "fixed")
End Sub


Private Sub Command2_Click()
txtGrade(0).Enabled = False
txtCredits(0).Enabled = False
'txtGrade(0).Text = ""
'txtCredits(0).Text = ""
End Sub

Private Sub Command3_Click()
txtGrade(1).Enabled = False
txtCredits(1).Enabled = False
'txtGrade(1).Text = ""
'txtCredits(1).Text = ""
End Sub

Private Sub Command4_Click()
txtGrade(2).Enabled = False
txtCredits(2).Enabled = False
'txtGrade(2).Text = ""
'txtCredits(2).Text = ""
End Sub

Private Sub Command5_Click()
txtGrade(3).Enabled = False
txtCredits(3).Enabled = False
txtGrade(3).Text = ""
txtCredits(3).Text = ""
End Sub

Private Sub Command6_Click()
txtGrade(4).Enabled = False
txtCredits(4).Enabled = False
txtGrade(4).Text = ""
txtCredits(4).Text = ""
End Sub

Private Sub Command7_Click()
txtGrade(5).Enabled = False
txtCredits(5).Enabled = False
txtGrade(5).Text = ""
txtCredits(5).Text = ""
End Sub

Private Sub Command8_Click()
txtGrade(6).Enabled = False
txtCredits(6).Enabled = False
txtGrade(6).Text = ""
txtCredits(6).Text = ""
End Sub

Private Sub mnuHelp_Click()
frmHelp.Show
End Sub

Private Sub mnuNew_Click()
For I = 0 To 6
txtGrade(I).Text = ""
lblClass.Caption = ""
txtCredits(I).Text = ""
lblGPA.Caption = ""
lblTotCredits.Caption = ""
txtGrade(I).Enabled = True
txtCredits(I).Enabled = True

Next
End Sub

Private Sub mnuOpen_Click()
Dim strLGrade As String, strLGrade1 As String, strLGrade2 As String, strLGrade3 As String, strLGrade4 As String, strLGrade5 As String, strLGrade6 As String
Dim strCredit As String, strCredits1 As String, strCredits2 As String, strCredits3 As String, strCredits4 As String, strCredits5 As String, strCredits6 As String
Dim strTClasses As String
Dim strTCredits As String
Dim strGPA As String

dlgCommon.CancelError = True

dlgCommon.Filter = "Text Files (*.txt) * *.txt"
dlgCommon.DialogTitle = "Open Text File"
dlgCommon.ShowOpen
strfilename = dlgCommon.FileName

dlgCommon.CancelError = True


Open strfilename For Input As #1
'Do While Not EOF(1)
Input #1, strLGrade, strLGrade1, strLGrade2, strLGrade3, strLGrade4, strLGrade5, strLGrade6, strCredit, strCredits1, strCredits2, strCredits3, strCredits4, strCredits5, strCredits6, strTClasses, strTCredits, strGPA

txtGrade(0).Text = strLGrade  'inserting letter grades
txtGrade(1).Text = strLGrade1
txtGrade(2).Text = strLGrade2
txtGrade(3).Text = strLGrade3
txtGrade(4).Text = strLGrade4
txtGrade(5).Text = strLGrade5
txtGrade(6).Text = strLGrade6

txtCredits(0).Text = strCredit 'inserting Credits
txtCredits(1).Text = strCredits1
txtCredits(2).Text = strCredits2
txtCredits(3).Text = strCredits3
txtCredits(4).Text = strCredits4
txtCredits(5).Text = strCredits5
txtCredits(6).Text = strCredits6
lblClass.Caption = strTClasses
lblTotCredits.Caption = strTCredits
lblGPA.Caption = lblGPA.Caption & strGPA

Close #1


End Sub

Private Sub mnuSave_Click()
'Dim strLGrade As String, strLGrade1 As String, strLGrade2 As String, strLGrade3 As String, strLGrade4 As String, strLGrade5 As String, strLGrade6 As String
'Dim strCredit As String, strCredits1 As String, strCredits2 As String, strCredits3 As String, strCredits4 As String, strCredits5 As String, strCredits6 As String
'Dim strTClasses As String
'Dim strTCredits As String
'Dim strGPA As String

dlgCommon.CancelError = True
dlgCommon.Filter = "Text File (*.txt) *.txt"
dlgCommon.DialogTitle = "Open Text Files"
dlgCommon.ShowOpen

strfilename = dlgCommon.FileName
Open strfilename For Output As #1
Write #1, txtGrade(0).Text, txtGrade(1).Text, txtGrade(2).Text, txtGrade(3).Text, txtGrade(4).Text, txtGrade(5).Text, txtGrade(6).Text, txtCredits(0).Text, txtCredits(1).Text, txtCredits(2).Text, txtCredits(3).Text, txtCredits(4).Text, txtCredits(5).Text, txtCredits(6).Text, lblClass.Caption, lblTotCredits.Caption, lblGPA.Caption 'Saving all of the slots
Close #1
End Sub

Private Sub txtCredits_Change(Index As Integer)
For I = 0 To 6
'If IsNumeric(txtCredits(i).Text) = False Then
'MsgBox "enter numbers only"
'txtCredits(Index).Text = 0
'End If
Next
End Sub

Private Sub txtGrade_Change(Index As Integer)
'For i = 0 To 6
'If txtGrade(i).Text <> "a" Or txtGrade(i).Text <> "a-" Or txtGrade(i).Text <> "b+" Or txtGrade(i).Text <> "b" Or txtGrade(i).Text <> "b-" Or txtGrade(i).Text <> "c+" Or txtGrade(i).Text <> "c" Or txtGrade(i).Text <> "c-" Or txtGrade(i).Text <> "d+" Or txtGrade(i).Text <> "d" Or txtGrade(i).Text <> "d-" Or txtGrade(i).Text <> "f" Or txtGrade(i).Text <> "" Then
'MsgBox "Please enter your letter grade only."
'txtGrade(Index).Text = ""
'end if
'Next
End Sub
