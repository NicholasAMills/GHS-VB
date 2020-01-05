VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Triangle"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdResults 
      Caption         =   "Results"
      Height          =   735
      Left            =   2880
      TabIndex        =   3
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox Text3 
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
      Left            =   4920
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text2 
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
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
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
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Side 3"
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
      Left            =   6720
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Side 2"
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
      Left            =   480
      TabIndex        =   6
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Side 1"
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
      Left            =   1800
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.Label LblTri 
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
      Left            =   5880
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdResults_Click()
'declaring variables
Dim strTri As String
Dim intSide1 As Integer
Dim intSide2 As Integer
Dim intSide3 As Integer

'correcting errors
If IsNumeric(Text1.Text) = False Then
MsgBox "please enter a valid number"
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Text1.SetFocus
LblTri = ""
Exit Sub
End If


If IsNumeric(Text2.Text) = False Then
MsgBox "please enter a valid number"
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
Text2.SetFocus
LblTri = ""
Exit Sub
End If


If IsNumeric(Text3.Text) = False Then
MsgBox "please enter a valid number"
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
Text3.SetFocus
LblTri = ""
Exit Sub
End If


'assigning variables
intTri = intSide1 + intSide2 + intSide3
intSide1 = Val(Text1)
intSide2 = Val(Text2)
intSide3 = Val(Text3)

'solving for triangle type
If intSide1 + intSide2 < intSide3 Or intSide2 + intSide3 < intSide1 Or intSide1 + intSide3 < intSide2 Then
strTri = "None"
End If

If strTri = "None" Then
MsgBox "non existing triangle"
End If

If intSide1 = intSide2 And intSide3 And intSide2 = intSide1 And intSide3 And intSide3 = intSide1 And intSide2 Then
strTri = "Equilateral"
End If

If intSide1 <> intSide2 And intSide3 Or intSide2 <> intSide1 And intSide3 Or intSide3 <> intSide1 And intSide2 Then
strTri = "Scalene"
End If

If intSide1 = intSide2 And intSide1 <> intSide3 Then
strTri = "Isosceles"
End If


'displaying results
LblTri.Caption = strTri
End Sub
