VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Calculator"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdStandard 
      Caption         =   "Standard"
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
      Left            =   6600
      TabIndex        =   11
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton CmdFixed 
      Caption         =   "Fixed"
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
      Left            =   6600
      TabIndex        =   10
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton CmdCash 
      Caption         =   "$"
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
      Left            =   6600
      TabIndex        =   9
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton CmdPer 
      Caption         =   "%"
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
      Left            =   6600
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton CmdDiv 
      Caption         =   "/"
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
      Left            =   4800
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton CmdMul 
      Caption         =   "X"
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
      Left            =   3720
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton CmdSub 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox TxtImput2 
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
      Left            =   3960
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox TxtImput1 
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
      Left            =   1200
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7320
      TabIndex        =   7
      Top             =   1440
      Width           =   75
   End
   Begin VB.Label Label1 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   1560
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sngImput1 As Single
Dim sngImput2 As Single
Dim sngTot As Single


Private Sub CmdAdd_Click()
'declaring variables for adding
sngImput1 = Val(TxtImput1)
sngImput2 = Val(TxtImput2)

'finding the total of adding Imput1 and Imput2
sngTot = sngImput1 + sngImput2
LblTot.Caption = sngTot

End Sub

Private Sub CmdCash_Click()
'finding the total as currency
LblTot.Caption = Format(LblTot.Caption, "currency")

End Sub

Private Sub CmdDiv_Click()
'declaring variables for dividing
sngImput1 = Val(TxtImput1)
sngImput2 = Val(TxtImput2)

'finding the total of dividing Imput1 from Imput2
sngTot = sngImput1 / sngImput2
LblTot.Caption = sngTot

End Sub

Private Sub CmdFixed_Click()
'displaying result as fixed numbers
LblTot.Caption = Format(LblTot.Caption, "fixed")
End Sub

Private Sub CmdMul_Click()
'declaring variables for multiplying
sngImput1 = Val(TxtImput1)
sngImput2 = Val(TxtImput2)

'finding the total of multiplying Imput1 and Imput2
sngTot = sngImput1 * sngImput2
LblTot.Caption = sngTot

End Sub

Private Sub CmdPer_Click()
'declaring value of total as a percent
LblTot.Caption = Format(LblTot.Caption, "percent")

End Sub

Private Sub CmdStandard_Click()
'displaying result as standard numbers
LblTot.Caption = Format(LblTot.Caption, "standard")
End Sub

Private Sub CmdSub_Click()
'declaring variables for subtracting
sngImput1 = Val(TxtImput1)
sngImput2 = Val(TxtImput2)

'Finding total of subracting Imput1 from Imput2
sngTot = sngImput1 - sngImput2
LblTot.Caption = sngTot


End Sub

