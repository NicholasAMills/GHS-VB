VERSION 5.00
Begin VB.Form frmCopies 
   Caption         =   "Printing Place"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "C&alculate"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   11
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtCopies 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblcost 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   4800
      TabIndex        =   12
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   8
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label lblCName 
      AutoSize        =   -1  'True
      Caption         =   "WELCOME:"
      Height          =   300
      Left            =   840
      TabIndex        =   7
      Top             =   1080
      Width           =   1290
   End
   Begin VB.Label lblTotal 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblPPCopy 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblLTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total Cost"
      Height          =   300
      Left            =   960
      TabIndex        =   4
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Label lblLPerCopy 
      AutoSize        =   -1  'True
      Caption         =   "Cost per copy"
      Height          =   300
      Left            =   720
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblCopies 
      Caption         =   "Please enter number of copies needed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Printing Place"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2400
      TabIndex        =   0
      Top             =   0
      Width           =   2955
   End
End
Attribute VB_Name = "frmCopies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
'declaring variables
Dim intCoppies As Single
Dim sngCost As Single
Dim sngTot As Single

'correcting errors
If IsNumeric(txtCopies.Text) = False Or txtCopies.Text = "" Then
    MsgBox "please enter the number of coppies"
    lblPPCopy.Caption = ""
    lblTotal.Caption = ""
    Exit Sub
    
End If

'assigning variables
intCoppies = Val(txtCopies)

'using select case
Select Case intCoppies
    Case Is <= 499
        sngCost = 0.3
    Case 500 To 749
        sngCost = 0.28
    Case 750 To 999
        sngCost = 0.27
    Case Is >= 1000
        sngCost = 0.25
End Select
    


'displaying cost per copy
lblPPCopy.Caption = sngCost
lblPPCopy.Caption = Format(lblPPCopy.Caption, "currency")

'calculating total
intTot = intCoppies * sngCost
lblTotal.Caption = intTot
lblTotal.Caption = Format(lblTotal.Caption, "currency")
End Sub

Private Sub cmdClear_Click()
'clear labels
lblPPCopy.Caption = ""
lblTotal.Caption = ""
txtCopies.Text = ""

'reset the variables intCost and intCoppies
sngCost = 0
intCoppies = 0
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblcost.Caption = "0-499      $0.30 per copy" & vbNewLine & _
                  "500-749  $0.28 per copy" & vbNewLine & _
                  "750-999  $0.27 per copy" & vbNewLine & _
                  ">=1000   $0.25 per copy"

'inserting the person's name
Dim strName As String
strName = UCase(InputBox("Please enter your name"))



'display user's name
lblName.Caption = strName
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim bytMessage As Byte
Const conButtons As Integer = vbYesNo + vbDefaultButton2 + vbQuestion + vbApplicationModal
bytMessage = MsgBox("Do you want to quit?", conButtons, "Exit")

If bytMessage = vbYes Then
    End
Else
    Cancel = 1
    frmCopies.Show
End If
End Sub

