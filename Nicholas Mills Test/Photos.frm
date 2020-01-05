VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   Caption         =   "Photo's Inc. "
   ClientHeight    =   4905
   ClientLeft      =   1200
   ClientTop       =   1515
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4905
   ScaleWidth      =   10560
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   615
      Left            =   8880
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mystery button"
      Height          =   735
      Left            =   7440
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Image imgPhoto 
      Height          =   3495
      Left            =   480
      Picture         =   "Photos.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
imgPhoto.Visible = False
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub imgPhoto_Click()
imgPhoto.Visible = False
End Sub
