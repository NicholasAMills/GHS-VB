VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReEnter 
      Caption         =   "Re-Enter Numbers"
      Height          =   615
      Left            =   6240
      TabIndex        =   4
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblSum 
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   6615
   End
   Begin VB.Label Label3 
      Caption         =   "Sum of Nubers"
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
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label lblNumbers 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "Numbers Entered"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sngInupt As Single 'used in load and reset
Dim strInput As String 'used in load and reset
Dim sngSum As Single 'used in load and reset

Private Sub cmdReEnter_Click()
Reset 'calls the reset general procedure
End Sub

Private Sub Form_Load()
'get user input
strInput = UCase(InputBox("please enter a number and 'Q' to quit", "Enter Numbers"))
'loop while the user doesn't enter Q
Do While strInput <> "Q"
'checks if user pressed cancel in the input box
If strInput = Cancel Then
    Exit Do
End If
'checks if user enters a number
If IsNumeric(strInput) = True Then
    snginput = Val(strInput) 'conbert string to number
    'display the number the user inputed
    lblNumbers.Caption = lblNumbers.Caption & " " & snginput
    sngSum = sngSum + snginput 'adds the user input together
Else
If strInput <> "Q" Then 'msgbox to display if user enters anything but Q
    MsgBox "Please enter numbers only and 'Q' to quit."
End If
End If
'gets input from the user in the LOOP
strInput = UCase(InputBox("please enter a number and 'Q' to quit", "Enter Numbers"))
Loop
'displays the sum
lblSum.Caption = sngSum
End Sub

Private Sub Reset()
snginput = 0
strInput = ""
sngSum = 0
lblNumbers.Caption = ""
lblSum.Caption = ""
'calls form load to get new numbers
Form_Load
End Sub

