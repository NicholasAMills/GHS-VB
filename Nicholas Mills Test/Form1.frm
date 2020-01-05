VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Feet Inches Yards"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Convert"
      Height          =   735
      Left            =   1800
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox TxtFt 
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
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label LblYds 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Yards"
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
      Left            =   5280
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label LblIn 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Inches"
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
      Left            =   5160
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Feet"
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
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_C()
End Sub

Private Sub Command1_Click()
'coverting feet to inches and yards
LblIn = Val(TxtFt * 12)
LblYds = Val(TxtFt / 3)
End Sub
