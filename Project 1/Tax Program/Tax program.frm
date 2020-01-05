VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Calculator"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtPrice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   3
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CommandButton CmdCalc 
      Caption         =   "&Calculate"
      Default         =   -1  'True
      Height          =   855
      Left            =   6960
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      Height          =   735
      Left            =   8040
      TabIndex        =   1
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label lblTotalPrice 
      Height          =   975
      Left            =   3120
      TabIndex        =   8
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label TtlPrice 
      Caption         =   "Total Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   7
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Tax"
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
      Left            =   2040
      TabIndex        =   6
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Price Here     ---->"
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
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label LblTxt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   4
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()

End Sub

Private Sub CmdCalc_Click()
'calculating the tax at 5.6%. User input in a textbox.
LblTxt.Caption = Val(TxtPrice.Text) * 0.056
lblTotalPrice.Caption = Val(TxtPrice.Text) + Val(LblTxt.Caption)
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

