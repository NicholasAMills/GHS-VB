VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "Calculate"
      Height          =   735
      Left            =   3840
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd1_Click()
Dim lngCount As Long
Dim lngTot As Long

lngCount = 50000

Do While lngCount >= 50000
If lngCount >= 50000 Then
    lngCount = lngCount * 2
    List1.AddItem (lngCount)
    
End If
If lngCount >= 100000000 Then
MsgBox "You'll reached $100 million in 11 weeks!!!"
Exit Do
End If
Loop

List1.AddItem (lngTot)
End Sub

Private Sub Command1_Click()
Unload Me
End Sub



Private Sub Form_Load()
List1.AddItem (50000)
End Sub

