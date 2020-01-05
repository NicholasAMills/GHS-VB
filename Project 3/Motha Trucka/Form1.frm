VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MileageEx"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton CmdImput 
      Caption         =   "&Input"
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "C&lear"
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton CmdCalculate 
      Caption         =   "&Calculate Pay"
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblPay 
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Mileage Pay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label LblMileage 
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Total Mileage"
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intMile As Integer 'used in cmdCalcualte and Form_Load events

Private Sub CmdCalculate_Click()
Const conRate As Single = 0.32 '32 cents for every mile driven
Dim sngPay As Single 'hold the total reinburstment
'calculate the total reinbursment
sngPay = intMile * conRate
'display the total ammount and formatting to currency
lblPay.Caption = Format(sngPay, "currency")
End Sub

Private Sub CmdClear_Click()
'clear labels
LblMileage.Caption = ""
lblPay.Caption = ""
'reset the variable intMile
intMile = 0
End Sub

Private Sub CmdExit_Click()
Form_Unload (0)
End Sub

Private Sub CmdImput_Click()
'calling the form load event code is all written there
'clear the money label
lblPay.Caption = ""
Form_Load
End Sub

Private Sub Form_Load()
Dim strMile As String 'local to get user input
'creates the input box for user input - ucase for Q
strMile = UCase(InputBox("Please enter the number of miles driven, or enter Q to quit", "Miles Driven"))
'check if user wans to exit
If strMile = "Q" Or strMile = "" Or strMile = Cancel Then
    Form_Unload (0) 'calling Form_unload even
Else
'check if user enter proper input of a number
If IsNumeric(strMile) = True Then
'convert a string to a number use the val function
intMile = Val(strMile)
'display intmile number in the lblmileage on form
LblMileage.Caption = intMile
Else
'user does not enter a number
MsgBox "Please enter a number for total miles driven or a Q to quit"
Form_Load 'calling form load to get a correct input from user


End If
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'variable to hold the user input from a message box
Dim bytMessage As Byte
'constant to create the buttons for the message box
Const conButtons As Integer = vbYesNo + vbDefaultButton2 + vbQuestion + vbApplicationModal
'creates the message box for user to see
bytMessage = MsgBox("Do you want to exit?", conButtons, "Exit")
'if statement to handle user input for Yes or No to exit
If bytMessage = vbYes Then
    End
Else
    'the unload event automatically sets the cancel variable to 0
    '0 will close the program so set cancel to 1 so program does not close
    Cancel = 1
    Form1.Show
End If

End Sub
