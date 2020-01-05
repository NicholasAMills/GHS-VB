VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Income Expenses"
   ClientHeight    =   9915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   735
      Left            =   5880
      TabIndex        =   18
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton CmdCalc 
      Caption         =   "Calculate"
      Height          =   735
      Left            =   5760
      TabIndex        =   17
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox TxtOthers 
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox TxtFood 
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtUtilities 
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox TxtCarPay 
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox TxtRent 
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox TxtIncome 
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label LblOthersPer 
      Height          =   375
      Left            =   5520
      TabIndex        =   24
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label LblFoodPer 
      Height          =   375
      Left            =   5520
      TabIndex        =   23
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label LblUtilitiesPer 
      Height          =   375
      Left            =   5640
      TabIndex        =   22
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label LblCarPayPer 
      Height          =   375
      Left            =   5640
      TabIndex        =   21
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label LblRentPer 
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Expense %'s"
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
      Left            =   5640
      TabIndex        =   19
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label LblProfit 
      Height          =   615
      Left            =   3120
      TabIndex        =   16
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label LblTot 
      Height          =   615
      Left            =   3120
      TabIndex        =   15
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Profit"
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
      Left            =   960
      TabIndex        =   14
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Total Expenses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   13
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Others"
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
      Left            =   1680
      TabIndex        =   11
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Food"
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
      Left            =   1680
      TabIndex        =   9
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Utilities"
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
      Left            =   1800
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Car Payments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Rent"
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
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Expenses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label LblIncome 
      Caption         =   "Income"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCalc_Click()
'declaring variables
Dim sngIncome As Single
Dim sngRent As Single
Dim sngCarPay As Single
Dim sngUtilities As Single
Dim sngFood As Single
Dim sngOthers As Single

'assigning variables
sngIncome = Val(TxtIncome.Text)
sngRent = Val(TxtRent.Text)
sngCarPay = Val(TxtCarPay.Text)
sngUtilities = Val(txtUtilities.Text)
sngFood = Val(TxtFood.Text)
sngOthers = Val(TxtOthers.Text)
sngTot = Val(LblTot.Caption)
sngProfit = Val(LblProfit.Caption)

'finding the Total Expenses
sngTot = sngRent + sngCarPay + sngUtilities + sngFood + sngOthers
LblTot.Caption = sngTot
LblTot.Caption = Format(LblTot.Caption, "currency")

'finding the profit
sngProfit = sngIncome - sngTot
LblProfit.Caption = sngProfit
LblProfit.Caption = Format(LblProfit.Caption, "currency")

'declaring expense %'s variables
Dim sngRentPer As Single
Dim sngCarPayPer As Single
Dim sngUtilitiesPer As Single
Dim sngFoodPer As Single
Dim sngothersPer As Single

'assigning expense %'s variables
sngRentPer = sngRent / sngIncome
sngCarPayPer = sngCarPay / sngIncome
sngUtilitiesPer = sngUtilities / sngIncome
sngFoodPer = sngFood / sngIncome
sngothersPer = sngOthers / sngIncome

'displaying expense %'s
LblRentPer = sngRentPer
LblRentPer.Caption = Format(LblRentPer.Caption, "percent")
LblCarPayPer.Caption = sngCarPayPer
LblCarPayPer.Caption = Format(LblCarPayPer.Caption, "percent")
LblUtilitiesPer.Caption = sngUtilitiesPer
LblUtilitiesPer.Caption = Format(LblUtilitiesPer.Caption, "percent")
LblFoodPer.Caption = sngFoodPer
LblFoodPer.Caption = Format(LblFoodPer.Caption, "percent")
LblOthersPer.Caption = sngothersPer
LblOthersPer.Caption = Format(LblOthersPer.Caption, "percent")


End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

