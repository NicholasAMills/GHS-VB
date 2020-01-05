VERSION 5.00
Begin VB.Form frmSR 
   Caption         =   "Sales Region"
   ClientHeight    =   5730
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsbSales 
      Height          =   255
      Index           =   3
      LargeChange     =   10000
      Left            =   5040
      Max             =   30000
      SmallChange     =   1000
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.HScrollBar hsbSales 
      Height          =   255
      Index           =   2
      LargeChange     =   10000
      Left            =   3480
      Max             =   30000
      SmallChange     =   1000
      TabIndex        =   14
      Top             =   3360
      Width           =   1215
   End
   Begin VB.HScrollBar hsbSales 
      Height          =   255
      Index           =   1
      LargeChange     =   10000
      Left            =   1920
      Max             =   30000
      SmallChange     =   1000
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   4320
      TabIndex        =   8
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   615
      Left            =   2520
      TabIndex        =   7
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      Height          =   615
      Left            =   720
      TabIndex        =   6
      Top             =   4200
      Width           =   1455
   End
   Begin VB.HScrollBar hsbSales 
      Height          =   255
      Index           =   0
      LargeChange     =   10000
      Left            =   360
      Max             =   30000
      SmallChange     =   1000
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtSales 
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
      Index           =   3
      Left            =   5040
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtSales 
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
      Index           =   2
      Left            =   3480
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtSales 
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
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtSales 
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
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblCom 
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
      Left            =   2760
      TabIndex        =   12
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblSales 
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
      Left            =   2760
      TabIndex        =   11
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Commissions: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Total Sales:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Sales Region"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu mnuCalculate 
      Caption         =   "Calculate"
   End
End
Attribute VB_Name = "frmSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCalc_Click()
mnuCalculate_Click
End Sub

Private Sub CmdClear_Click()
Dim i As Integer 'variable to represent index value
lblSales.Caption = ""
lblCom.Caption = ""
'using For loop to clear scroll bar and text bars
For i = 0 To 3
hsbSales(i).Value = 0
txtSales(i).Text = 0
Next
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub hsbSales_Change(Index As Integer)
'putting the horizontal scrool bars value into the
'text box with the same index number
txtSales(Index).Text = hsbSales(Index).Value
End Sub

Private Sub mnuCalculate_Click()
For i = 0 To 3
curTotal = curTotal + Val(txtSales(i).Text)
Next

lblSales.Caption = Format(curTotal, "currency")

'if statements for commission
If curTotal <= 20000 Then
curCom = 0.01
Else
If curTotal <= 40000 And curTotal > 20000 Then
curCom = 0.03
Else
If curTotal > 40000 Then
curCom = 0.05
End If
End If
End If
curCom = curTotal * curCom
lblCom.Caption = Format(curCom, "currency")
End Sub

Private Sub mnuClear_Click()
CmdClear_Click
End Sub

Private Sub mnuExit_Click()
cmdExit_Click
End Sub

Private Sub txtSales_Change(Index As Integer)
If IsNumeric(txtSales(Index).Text) = False Then
    MsgBox "Enter your total Sales. Numbers only"
    txtSales(Index).SelStart = 0
    txtSales(Index).SelLength = Len(txtSales(Index).Text)
    txtSales(Index).SetFocus
Else
'putting the text box value as the horizontal
'scroll bars value withthe same index number
hsbSales(Index).Value = Val(txtSales(Index).Text)
End If

End Sub

