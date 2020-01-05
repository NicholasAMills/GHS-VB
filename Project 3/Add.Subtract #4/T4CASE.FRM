VERSION 5.00
Begin VB.Form frmMath 
   Caption         =   "Math Practice"
   ClientHeight    =   5100
   ClientLeft      =   1050
   ClientTop       =   1470
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5100
   ScaleWidth      =   8190
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1815
      Left            =   3240
      TabIndex        =   14
      Top             =   2640
      Width           =   2055
      Begin VB.OptionButton optSubtract 
         Caption         =   "Subtraction"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optAdd 
         Caption         =   "&Addition"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1815
      Left            =   600
      TabIndex        =   11
      Top             =   2520
      Width           =   2055
      Begin VB.OptionButton optLevel2 
         Caption         =   "Level &2   (10-100)"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optLevel1 
         Caption         =   "Level &1   (1-10)"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1215
      Left            =   5400
      TabIndex        =   5
      Top             =   2760
      Width           =   2415
      Begin VB.Label lblIncorrect 
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
         Left            =   1440
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblCorrect 
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
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Correct:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label4 
         Caption         =   "Incorrect:"
         Height          =   240
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   7455
      Begin VB.CommandButton cmdVerify 
         Caption         =   "&Verify Answer"
         Default         =   -1  'True
         Height          =   495
         Left            =   5640
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtAnswer 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3840
         TabIndex        =   0
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Equals 
         Caption         =   "8"
         Height          =   375
         Left            =   3840
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblMakes 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   18
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblEquals 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   7200
         TabIndex        =   17
         Top             =   480
         Width           =   15
      End
      Begin VB.Image imgOperator 
         Height          =   495
         Left            =   1560
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblNum1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblNum2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Image imgMinus 
      Height          =   480
      Left            =   1560
      Picture         =   "T4CASE.frx":0000
      Top             =   4440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPlus 
      Height          =   480
      Left            =   960
      Picture         =   "T4CASE.frx":0442
      Top             =   4440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHappy 
      Height          =   480
      Left            =   360
      Picture         =   "T4CASE.frx":0884
      Top             =   4440
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdVerify_Click()
Dim intAnswer As Integer
Dim intCorrect As Integer
Dim intIncorrect As Integer
Dim intNum1 As Integer
Dim intNum2 As Integer

intCorrect = Val(lblCorrect)
intIncorrect = Val(lblIncorrect)
intAnswer = Val(txtAnswer.Text)
intNum1 = Val(lblNum1.Caption)
intNum2 = Val(lblNum2.Caption)

'correcting errors
If IsNumeric(txtAnswer) = False Then
MsgBox "Please enter only numbers"
Exit Sub
End If

'if answer is correct
If intAnswer = Val(Equals.Caption) Then
MsgBox "Correct!"
imgHappy.Visible = True
lblCorrect.Caption = Val(lblCorrect.Caption) + 1
txtAnswer.Text = ""

End If

'if answer is incorrect
If intAnswer <> Equals.Caption Then
MsgBox "Incorrect"
lblIncorrect.Caption = intIncorrect + 1
End If

'if subtraction is selected
If optSubtract.Value = True Then
Equals.Caption = intNum1 - intNum2

End If


'if addition is selected
If optAdd.Value = True Then
Equals.Caption = intNum1 + intNum2

End If

'making 1st number always larger
If intNum1 < intNum2 Then
    intNum1 = intNum2 + 1
End If

End Sub

Private Sub Form_Load()
    frmMath.Top = (Screen.Height - frmMath.Height) / 2
    frmMath.Left = (Screen.Width - frmMath.Width) / 2

End Sub


Private Sub Label1_Click()

End Sub

Private Sub optAdd_Click()
If optAdd.Value = True Then
imgOperator = imgPlus
End If
End Sub

Private Sub optLevel1_Click()
Dim intNum1 As Integer
Dim intNum2 As Integer
intNum1 = Val(lblNum1.Caption)
intNum2 = Val(lblNum2.Caption)

'giving values to lblNum1 and lblNum2
Randomize
If optLevel1.Value = True Then
intNum1 = Int((10 * Rnd) + 0)
intNum2 = Int((intNum1 * Rnd) + 0)
End If

If intNum1 < intNum2 Then
    intNum1 = intNum2 + 1
End If

lblNum1.Caption = intNum1
lblNum2.Caption = intNum2
End Sub

Private Sub optLevel2_Click()

'adding values to lblNum1 and lblNum2
Randomize
If optLevel2.Value = True Then
intNum1 = Int((100 * Rnd) + 0)
intNum2 = Int((intNum1 * Rnd) + 0)
End If
lblNum1.Caption = intNum1
lblNum2.Caption = intNum2

'making 1st number always larger
If intNum1 < intNum2 Then
    intNum1 = intNum2 + 1
End If

End Sub

Private Sub optSubtract_Click()
If optSubtract.Value = True Then
imgOperator = imgMinus
End If


End Sub
