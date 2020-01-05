VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Willow Health Club"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   4800
      TabIndex        =   12
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "&Print"
      Height          =   615
      Left            =   2400
      TabIndex        =   11
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton CmdCal 
      Caption         =   "&Calculate Total Due"
      Default         =   -1  'True
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Aditional Fees"
      Height          =   1935
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   1455
      Begin VB.CheckBox ChkRacquetball 
         Caption         =   "&Racquetball"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox ChkGolf 
         Caption         =   "&Golf"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox ChkTennis 
         Caption         =   "&Tennis"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label LblTot 
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
      Left            =   4080
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Total Due:"
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
      Left            =   2400
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label LblAdd 
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
      Left            =   4080
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Additional:"
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
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label LblBasic 
      Caption         =   "80"
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
      Left            =   4080
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Basic Fees:"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCal_Click()
'declaring variables
Dim intTennis As Integer
Dim intGolf As Integer
Dim intRacquetball As Integer
Dim intAdd As Integer
Dim inttot As Integer

'adding values to check boxes
If ChkTennis.Value = vbChecked Then
intTennis = 30
Else
intTenis = 0
End If

If ChkGolf.Value = vbChecked Then
intGolf = 25
Else
intGolf = 0
End If

If ChkRacquetball.Value = vbChecked Then
intRacquetball = 20
Else
intRacquetball = 0
End If

'calculating total
intBasic = Val(LblBasic.Caption)
intAdd = intTennis + intGolf + intRacquetball
LblAdd.Caption = Format(intAdd, "currency")
inttot = intAdd + intBasic
LblTot.Caption = inttot
LblTot.Caption = Format(LblTot.Caption, "currency")

End Sub

Private Sub CmdPrint_Click()
PrintForm
End Sub

Private Sub Command1_Click()
Unload Me
End Sub
