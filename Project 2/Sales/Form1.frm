VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "State Sales"
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   735
      Left            =   9120
      TabIndex        =   11
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton CmdCalc 
      Caption         =   "Calculate"
      Height          =   735
      Left            =   2640
      TabIndex        =   6
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox TxtFlorida 
      Height          =   615
      Left            =   2640
      TabIndex        =   5
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox TxtMaine 
      Height          =   615
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox TxtNewYork 
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label LblTotal 
      Height          =   375
      Left            =   9120
      TabIndex        =   10
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label LblCom 
      Height          =   495
      Left            =   9120
      TabIndex        =   8
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label 
      Caption         =   "Commissions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label LblFlorida 
      Caption         =   "Florida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label LblMaine 
      Caption         =   "Maine"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label LblNewYork 
      Caption         =   "New York"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCalc_Click()
'declaring variables
Dim intNewYork As Integer
Dim intMaine As Integer
Dim intFlorida As Integer
Dim sngCom As Single
Dim intTotal As Integer

'assigning variables
intNewYork = Val(TxtNewYork)
intMaine = Val(TxtMaine)
intFlorida = Val(TxtFlorida)
intTotal = intNewYork + intMaine + intFlorida
sngCom = intTotal * 0.05

'finding the total sales price and commission
LblTotal.Caption = Format(intTotal, "currency")
LblCom.Caption = Format(sngCom, "currency")

End Sub

Private Sub Command1_Click()
'exiting the project
Unload Me
End Sub
