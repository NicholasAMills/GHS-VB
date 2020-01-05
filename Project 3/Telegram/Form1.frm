VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Telegram"
   ClientHeight    =   4365
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtLetter 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdCal 
      Caption         =   "Calculate Cost"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   6
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblChar 
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
      Left            =   6240
      TabIndex        =   10
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Total Characters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   9
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblTotal 
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
      Left            =   6240
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Total Cost"
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
      Left            =   4080
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblAdd 
      Caption         =   "$0.02"
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
      Left            =   6240
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Additional characters (101+)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblCost 
      Caption         =   "$4.20"
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
      Left            =   6240
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Starting Cost"
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
      Left            =   4080
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCal_Click()
'declaring variables
Dim sngCost As Single
Dim sngAdd As Single
Dim sngTotal As Single
Dim intLetter As Integer

'assigning variables
sngCost = 4.2


'correcting errors
If txtLetter = "" Then
    MsgBox "please write a letter before sending"
End If

intLetter = Len(txtLetter.Text)

If intLetter <= 100 Then
    sngAdd = 0
Else
If intLetter >= 101 Then
   sngAdd = (strLetter - 100) * 0.02
End If
End If

lblChar.Caption = intLetter
sngTotal = sngCost + sngAdd

lblTotal.Caption = sngTotal
lblTotal.Caption = Format(lblTotal.Caption, "currency")
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

