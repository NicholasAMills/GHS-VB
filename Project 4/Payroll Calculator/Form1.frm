VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Payroll Calculator"
   ClientHeight    =   4275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
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
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblHours 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblPay 
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
      Left            =   2640
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Total Pay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Pay Rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Total Hours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sngTotal As Single
Dim intHours As Integer
Dim sngOvertime As Single
Dim sngOvertimeTotal As Single

'input boxes for ammount of hours worked
Private Sub Form_Load()
 strDays = Val(InputBox("Enter the ammount of hours and enter a negative number to quit"))
 
 Do While strDays >= 0
 
 If strDays = Cancel Then
    End
End If

 If IsNumeric(strDays) = True Then
    intDays = Val(strDays)
    intHours = intHours + intDays
    lblHours.Caption = intHours
End If



If sngPay <= 5.5 Then
sngTotal = sngTotal + 10
End If

strDays = Val(InputBox("Enter the ammount of hours and enter a negative number to quit"))
Loop
lblPay.Caption = sngTotal
End Sub

Private Sub Text1_Change()
Dim sngPay As Single
Dim sngPayRate As Single

'calculating total hours and pay
sngPayRate = Val(Text1.Text)

sngPay = sngPayRate * intHours
lblPay.Caption = sngPay


'correcting errors
If IsNumeric(Text1.Text) = False Then
MsgBox "Please enter numbers only"
End If



'formatting to currency
lblPay.Caption = Format(lblPay.Caption, "currency")

'if statements for overtime and > $5.50 p/h
If intHours >= 40 Then
sngOvertime = Val(intHours) - 40
sngPay = sngPay * 1.5
sngOvertimeTotal = sngOvertime * sngPay
sngTotal = sngTotal + sngOvertimeTotal
End If
End Sub
