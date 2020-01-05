VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Math Tutor"
   ClientHeight    =   9405
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "12."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   35
      Top             =   8520
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "11."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   34
      Top             =   7800
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "10."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   33
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "9."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   32
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "8."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   31
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "7."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   30
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "6."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   29
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "5."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   28
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "4."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   27
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   26
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   25
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   24
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblAnswerCV 
      BackColor       =   &H00FFFFFF&
      Caption         =   "         6"
      Height          =   375
      Index           =   11
      Left            =   2520
      TabIndex        =   23
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblAnswerCV 
      BackColor       =   &H00FFFFFF&
      Caption         =   "         1"
      Height          =   375
      Index           =   10
      Left            =   2520
      TabIndex        =   22
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblAnswerCV 
      BackColor       =   &H00FFFFFF&
      Caption         =   "       20"
      Height          =   375
      Index           =   9
      Left            =   2520
      TabIndex        =   21
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lblAnswerCV 
      BackColor       =   &H00FFFFFF&
      Caption         =   "       14"
      Height          =   375
      Index           =   8
      Left            =   2520
      TabIndex        =   20
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lblAnswerCV 
      BackColor       =   &H00FFFFFF&
      Caption         =   "      100"
      Height          =   375
      Index           =   7
      Left            =   2520
      TabIndex        =   19
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblAnswerCV 
      BackColor       =   &H00FFFFFF&
      Caption         =   "       36"
      Height          =   375
      Index           =   6
      Left            =   2520
      TabIndex        =   18
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblAnswerCV 
      BackColor       =   &H00FFFFFF&
      Caption         =   "         5"
      Height          =   375
      Index           =   5
      Left            =   2520
      TabIndex        =   17
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblAnswerCV 
      BackColor       =   &H00FFFFFF&
      Caption         =   "         0"
      Height          =   375
      Index           =   4
      Left            =   2520
      TabIndex        =   16
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label lblAnswerCV 
      BackColor       =   &H00FFFFFF&
      Caption         =   "         7"
      Height          =   375
      Index           =   3
      Left            =   2520
      TabIndex        =   15
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label lblAnswerCV 
      BackColor       =   &H00FFFFFF&
      Caption         =   "       18"
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   14
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblAnswerCV 
      BackColor       =   &H00FFFFFF&
      Caption         =   "       12"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   13
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label lblProblem 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "2 x 3"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   11
      Left            =   840
      TabIndex        =   12
      Top             =   8520
      Width           =   345
   End
   Begin VB.Label lblProblem 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "1 x 1"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   10
      Left            =   840
      TabIndex        =   11
      Top             =   7800
      Width           =   345
   End
   Begin VB.Label lblProblem 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "10 x 2"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   9
      Left            =   840
      TabIndex        =   10
      Top             =   7080
      Width           =   435
   End
   Begin VB.Label lblProblem 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "2 x 7"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   8
      Left            =   840
      TabIndex        =   9
      Top             =   6360
      Width           =   345
   End
   Begin VB.Label lblProblem 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "10 x 10"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   7
      Left            =   840
      TabIndex        =   8
      Top             =   5640
      Width           =   525
   End
   Begin VB.Label lblProblem 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "9 x 4"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   6
      Left            =   840
      TabIndex        =   7
      Top             =   4920
      Width           =   345
   End
   Begin VB.Label lblProblem 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "5 x 1"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   5
      Left            =   840
      TabIndex        =   6
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label lblProblem 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0 x 1"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   4
      Left            =   840
      TabIndex        =   5
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label lblProblem 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "7x1"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   3
      Left            =   840
      TabIndex        =   4
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label lblProblem 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "6 x 3"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   2
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   345
   End
   Begin VB.Label lblProblem 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "3 x 4"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   345
   End
   Begin VB.Label lblAnswerCV 
      BackColor       =   &H00FFFFFF&
      Caption         =   "       10"
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblProblem 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "5 x 2"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   345
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuQuiz 
         Caption         =   "Quiz"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuclear 
         Caption         =   "Clear"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intcount As Integer
Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
'unallowable drop; return sign to original location
Source.Visible = True
End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'set sign to invisible when dragging begins
Source.Visible = False
End Sub

Private Sub lblAnswerCV_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
Static intCorrect As Integer

'if statements for correct/incorrect answers
If Source.Index = Index Then
lblAnswerCV(Index).BackColor = vbGreen
intCorrect = intCorrect + 1
Else
MsgBox "Incorrect. Try again."
Source.Visible = True
intcount = intcount + 1
lblAnswerCV(Index).BackColor = vbRed
End If

If intCorrect = 12 Then
MsgBox "Great Job!"
intCorrect = 0
Call mnuclear_Click
End If

For i = 0 To 11
If lblAnswerCV(i).BackColor = vbGreen Then
lblAnswerCV(i).Enabled = False
End If
Next

'exiting form if intCount = 3
If intcount = 3 Then
MsgBox "You have reached your maximum tries. Try again."
mnuclear_Click
End If
End Sub

Private Sub mnuclear_Click()
intcount = 0
For i = 0 To 11
lblProblem(i).Visible = True
lblAnswerCV(i).BackColor = vbWhite
lblAnswerCV(i).Enabled = True
Next

End Sub

Private Sub mnuQuit_Click()
Unload Me
End Sub

Private Sub mnuQuiz_Click()
Const Instruction$ = "Type the letter of the correct response and " & _
                     "click OK. Click Cancel to skip this question. " _
                     & vbNewLine & vbNewLine
Const Choice$ = vbNewLine & vbNewLine & "a. 30" & vbNewLine & _
                "b. 10" & vbNewLine & "c. 100"
Dim QuesNum%, Question$, CorrectAnswers$, Response$
'hide form
Form1.Hide
'loop for three questions
For QuesNum% = 1 To 3
    'assign value to variable Question$ and variable CorrectAndwer$
    Select Case QuesNum%
    Case Is = 1
        Question$ = "1. 5 x 2 = ?" & Choice$
        CorrectAnswer$ = "B"
    Case Is = 2
        Question$ = "2. 10 x 10 = ?" & Choice$
        CorrectAnswer$ = "C"
    Case Is = 3
        Question$ = "3. 5 x 6 = ?" & Choice$
        CorrectAnswer$ = "A"
    End Select
    'display question; assign returned value to variable Response$
    Response$ = InputBox(Instruction$ & Question$, "Math Quiz")
    'begin loop for correct answer or cancel button
    Do Until UCase$(Response$) = CorrectAnswer$ Or Response$ = ""
    'display message box for wrong answer
    MsgBox "Your response was not correct. Please try again", , _
           "Math Quiz"
    'display question again
    Response$ = InputBox(Instruction$ & Question$, _
    "Math Quiz")
    Loop
'add 1 to counter in for...next loop
Next
'display main form after 3rd question
Form1.Show

End Sub
