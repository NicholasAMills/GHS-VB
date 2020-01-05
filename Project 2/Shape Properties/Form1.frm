VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmBorder 
      Caption         =   "BorderStyle"
      Height          =   2175
      Left            =   4320
      TabIndex        =   1
      Top             =   3000
      Width           =   2055
      Begin VB.OptionButton OptDot 
         Caption         =   "Dot"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton OptDash 
         Caption         =   "Dash"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton OptSolid 
         Caption         =   "Solid"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame FrmShape 
      Caption         =   "Shape"
      Height          =   2175
      Left            =   720
      TabIndex        =   0
      Top             =   3000
      Width           =   2055
      Begin VB.OptionButton OptOval 
         Caption         =   "Oval"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   1335
      End
      Begin VB.OptionButton OptCircle 
         Caption         =   "Circle"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton OptRec 
         Caption         =   "Rectangle"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Shape ShpNewShape 
      Height          =   1455
      Left            =   2160
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OptCircle_Click()
ShpNewShape.Shape = 3
End Sub

Private Sub OptDash_Click()
'making the borderstyle dashed
ShpNewShape.BorderStyle = 2
End Sub

Private Sub OptDot_Click()
'making the borderstyle dotted
ShpNewShape.BorderStyle = 3
End Sub

Private Sub OptOval_Click()
ShpNewShape.Shape = 2
End Sub

Private Sub OptRec_Click()
ShpNewShape.Shape = 0
End Sub

Private Sub OptSolid_Click()
'making the borderstyle solid
ShpNewShape.BorderStyle = 1
End Sub
