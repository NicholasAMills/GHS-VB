VERSION 5.00
Begin VB.Form frmLoanpmt 
   Caption         =   "SavU Loan Analyzer"
   ClientHeight    =   3885
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4605
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsbYears 
      Height          =   255
      LargeChange     =   5
      Left            =   240
      Max             =   30
      Min             =   1
      TabIndex        =   12
      Top             =   3120
      Value           =   1
      Width           =   1695
   End
   Begin VB.HScrollBar hsbRate 
      Height          =   255
      LargeChange     =   10
      Left            =   240
      Max             =   1500
      Min             =   1
      TabIndex        =   11
      Top             =   2160
      Value           =   1
      Width           =   1695
   End
   Begin VB.TextBox TxtAmt 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton CmdAbout 
      Caption         =   "About..."
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton CmdCalc 
      Caption         =   "Calculate"
      Default         =   -1  'True
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblSumpmt 
      Height          =   195
      Left            =   2760
      TabIndex        =   14
      Top             =   1320
      Width           =   1440
   End
   Begin VB.Label LblPayment 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2760
      TabIndex        =   13
      Top             =   600
      Width           =   1485
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "APR"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   330
   End
   Begin VB.Label lblRate 
      AutoSize        =   -1  'True
      Caption         =   "01"
      Height          =   195
      Left            =   1440
      TabIndex        =   5
      Top             =   1680
      Width           =   180
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "YEARS"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "MONTHLY PAYMENT"
      Height          =   195
      Left            =   2760
      TabIndex        =   3
      Top             =   240
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "SUM OF PAYMENTS"
      Height          =   195
      Left            =   2760
      TabIndex        =   2
      Top             =   960
      Width           =   1545
   End
   Begin VB.Label LblYears 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Left            =   1440
      TabIndex        =   1
      Top             =   2640
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "LOAN AMMOUNT"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1320
   End
   Begin VB.Shape Shape3 
      Height          =   1935
      Left            =   2640
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      Height          =   1575
      Left            =   2640
      Top             =   120
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      Height          =   3615
      Left            =   120
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmLoanpmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAmt_Click()

End Sub

Private Sub CmdAbout_Click()
'display modal about dialog form
frmLoanabt.Show 1

End Sub

Private Sub CmdCalc_Click()
'if amount is not a number, then message else perform calculatios
If IsNumeric(TxtAmt.Text) = False Then
    MsgBox "Please Enter Loan Amount In Numbers Only", 48, "SavU Loan Analyzer"
    TxtAmt.Text = ""
    TxtAmt.SetFocus
Else
    montlypmt = Pmt(0.0001 * hsbRate.Value / 12, hsbYears.Value * 12, -1 * TxtAmt.Text, 0.1)
    LblPayment.Caption = Format$(montlypmt, "currency")
    lblSumpmt.Caption = Format$(montlypmt * hsbYears.Value * 12, "curency")
End If

End Sub

Private Sub CmdClear_Click()
'claer input amount and outputs; reset scrollbars to minimums
TxtAmt.Text = ""
hsbYears.Value = 1
hsbRate.Value = 1
LblPayment.Caption = ""
lblSumpmt.Caption = ""
TxtAmt.SetFocus

End Sub

Private Sub hsbRate_Change()
lblRate.Caption = hsbRate.Value * 0.01
End Sub

Private Sub hsbYears_Change()
'update lblYears caption when scrollbox is moved
LblYears.Caption = hsbYears.Value

End Sub

Private Sub hsbYears_Scroll()
'update lblYears caption when scrollbox is moved
LblYears.Caption = hsbYears.Value

End Sub
