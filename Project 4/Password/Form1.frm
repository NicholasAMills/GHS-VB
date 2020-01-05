VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "Welcome"
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
      Left            =   3000
      TabIndex        =   0
      Top             =   2040
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim intCount As Integer
'loops unit Password = "Pass"
Do
Password = InputBox$("Please Enter Your Password", "Password Check")
If Password = Cancel Then
End
End If
'if statements for correct/incorrect password
If Password = "Pass" Then
MsgBox "Password Correct. Click OK to Procede"
Else
If Password <> "Pass" Then
MsgBox "Incorrect. Try again"
'adding 1 to intCount if incorrect
intCount = intCount + 1
End If
End If
'if intCount is 3 then form unloads
If intCount = 3 Then
MsgBox "Maximum tries reached. Try again later"
End
End If
Loop Until Password = "Pass"


End Sub
