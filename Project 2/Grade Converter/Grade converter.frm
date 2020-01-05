VERSION 5.00
Begin VB.Form FrmGrade 
   Caption         =   "Grade Converter"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Convert to Letter Grade"
      Height          =   735
      Left            =   1920
      TabIndex        =   4
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox TxtGrade 
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
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label LblLGrade 
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
      Left            =   2640
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Letter Grade"
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
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Caption1 
      Caption         =   "Enter Grade"
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
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "FrmGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim sngGrade As Single
Dim strLGrade As String

'check to see if user enters a number in the text box
If IsNumeric(TxtGrade.Text) = False Then 'user did not enter a number
    MsgBox "Pleae enter a number for your percent grade", vbQuestion + vbDefaultButton2, "Input Error"
    'tell the computer to start selecting the string from bad input starting at 0(first letter)
    TxtGrade.SelStart = 0
    'getting the length of the string in the textbox that we need to select
    TxtGrade.SelLength = Len(TxtGrade.Text)  'Len is the length function used to get the length of a string
    'put the cursor in the text box
    TxtGrade.SetFocus
    'exit the sub so user can put in proper input for program to run
    Exit Sub
Else
    'user enter a number and assign that number to sngGrade
    sngGrade = Val(TxtGrade.Text)
End If


    
    
sngGrade = Val(TxtGrade)
strLGrade = Val(LblLGrade)

If sngGrade >= 92 Then
    strLGrade = "A"
Else
    If sngGrade = 90 Or sngGrade = 91 Then
    strLGrade = "A-"
Else

    If sngGrade = 88 Or sngGrade = 89 Then
    strLGrade = "B+"
Else
    If sngGrade >= 82 And sngGrade <= 87 Then
    strLGrade = "B"
Else
    If sngGrade = 80 Or sngGrade = 81 Then
    strLGrade = "B-"
Else
    If sngGrade = 78 Or sngGrade = 79 Then
    strLGrade = "C+"
Else
    If sngGrade >= 72 Or sngGrade <= 77 Then
    strLGrade = "C"
Else
    If sngGrade = 70 Or sngGrade = 71 Then
    strLGrade = "C-"
Else
    If sngGrade = 68 Or sngGrade = 69 Then
    strLGrade = "D+"
Else
    If sngGrade >= 67 And sngGrade <= 68 Then
    strLGrade = "D"
Else
    If sngGrade = 60 Or sngGrade = 61 Then
    strLGrade = "D-"
Else
    If sngGrade <= 59 Then
    strLGrade = "F"
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
LblLGrade.Caption = strLGrade
End Sub

Private Sub Form_Unload(Cancel As Integer)
'ask th euser if they really want to exit
Dim bytAnswer As Byte 'local variable
bytAnswer = MsgBox("Do you want to exit?", vbYesNo + vbQuestion, "Exit")
If bytAnswer = vbYes Then
    'user does want to exit
    Unload Me
Else
    'user does not want to exit
    FrmGrade.Show
    Cancel = 1
End If
















End Sub
