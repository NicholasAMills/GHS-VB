VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "moving shape"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   3375
      LargeChange     =   200
      Left            =   6960
      Max             =   3600
      Min             =   480
      TabIndex        =   1
      Top             =   1440
      Value           =   480
      Width           =   375
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      LargeChange     =   200
      Left            =   720
      Max             =   3600
      Min             =   480
      TabIndex        =   0
      Top             =   5280
      Value           =   480
      Width           =   6135
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HScroll1_Change()
'making the circle's with bigger or smaller
Shape1.Width = HScroll1.Value
End Sub

Private Sub VScroll1_Change()
'making the shape's height larger or smaller
Shape1.Height = VScroll1.Value
End Sub
