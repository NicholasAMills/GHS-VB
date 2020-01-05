VERSION 5.00
Begin VB.Form SLSCOMSN 
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   4215
   ClientTop       =   3540
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   9915
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   975
      Left            =   4200
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Commission"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Sales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "SLSCOMSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Display Amount as Value of User Input * Rate
Text2.Text = Val(Text1.Text) * 0.15
Text1.SetFocus
End Sub

