VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FF80&
   Caption         =   "This was so easy I can't believe i had trouble with this -_-"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
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
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00FFFF80&
      Caption         =   "Circle"
      Height          =   615
      Index           =   3
      Left            =   6480
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H0080FFFF&
      Caption         =   "Oval"
      Height          =   615
      Index           =   2
      Left            =   4680
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H0080C0FF&
      Caption         =   "Rectangle"
      Height          =   615
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H008080FF&
      Caption         =   "Square"
      Height          =   615
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
   Begin VB.Shape Shape 
      Height          =   975
      Index           =   3
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape 
      Height          =   975
      Index           =   2
      Left            =   4560
      Shape           =   2  'Oval
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape 
      Height          =   975
      Index           =   1
      Left            =   2520
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape 
      Height          =   975
      Index           =   0
      Left            =   600
      Shape           =   1  'Square
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Opt_Click(Index As Integer)

For i = 0 To 3
If Opt(i).Value = False Then
Shape(i).Visible = False
End If
Next
If Opt(Index).Value = True Then
Shape(Index).Visible = True
End If
End Sub
