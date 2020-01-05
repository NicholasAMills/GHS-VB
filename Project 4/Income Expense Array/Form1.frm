VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIncome 
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
      Left            =   4440
      TabIndex        =   24
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   615
      Left            =   6480
      TabIndex        =   19
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      Height          =   735
      Left            =   6480
      TabIndex        =   18
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox txt1 
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
      Left            =   4440
      TabIndex        =   11
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txt1 
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
      Left            =   4440
      TabIndex        =   10
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txt1 
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
      Left            =   4440
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txt1 
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
      Left            =   4440
      TabIndex        =   8
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txt1 
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
      Left            =   4440
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblProf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   23
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Profit"
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
      Left            =   2280
      TabIndex        =   22
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label lblTot 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   21
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Total Expenses"
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
      Left            =   2280
      TabIndex        =   20
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label lblExpence 
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
      Index           =   4
      Left            =   6480
      TabIndex        =   17
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label lblExpence 
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
      Index           =   3
      Left            =   6480
      TabIndex        =   16
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label lblExpence 
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
      Index           =   2
      Left            =   6480
      TabIndex        =   15
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lblExpence 
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
      Index           =   1
      Left            =   6480
      TabIndex        =   14
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblExpence 
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
      Index           =   0
      Left            =   6480
      TabIndex        =   13
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Expence %"
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
      Left            =   6480
      TabIndex        =   12
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Other"
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
      Left            =   2280
      TabIndex        =   6
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Food"
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
      Left            =   2280
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Utilities"
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
      Left            =   2280
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Car Payments"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Rent"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Income"
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
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Expenses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalc_Click()
Dim sngTotal As Single
Dim sngIncome As Single
Dim sngExp As Single
Dim sngProf As Single

sngIncome = Val(txtIncome.Text)

'for loop adding expenses together and calculating expence %
For I = 0 To 4
sngTot = sngTot + Val(txt1(I).Text)
sngExp = txt1(I).Text / sngIncome
lblExpence(I).Caption = Format(sngExp, "percent")
Next


sngProf = sngIncome - sngTot

'formatting to currency
lblTot.Caption = Format(sngTot, "currency")
lblProf.Caption = Format(sngProf, "currency")


End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub txt1_Change(Index As Integer)
'correcting errors


If IsNumeric(txt1(Index).Text) = False Then
MsgBox "Please enter only numbers"
txt1(Index).Text = 0
lblTot.Caption = ""
lblProf.Caption = ""
For I = 0 To 4
lblExpence(I).Caption = ""

Next
End If
End Sub

Private Sub txtIncome_Change()
'correcting errors
If IsNumeric(txtIncome.Text) = False Then
MsgBox "Please enter only numbers"
txtIncome.Text = 0
lblTot.Caption = ""
lblProf.Caption = ""
For I = 0 To 4
lblExpence(I).Caption = ""
Next
End If


End Sub
