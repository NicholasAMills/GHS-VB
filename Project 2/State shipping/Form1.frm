VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "State Shipping Costs"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtPrice 
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
      Left            =   1920
      TabIndex        =   5
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton CmdCal 
      Caption         =   "Calculate"
      Height          =   1335
      Left            =   4680
      TabIndex        =   1
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox TxtState 
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
      Left            =   1920
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label LblShipping 
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
      Left            =   2040
      TabIndex        =   10
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Shipping"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label LblTax 
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
      Left            =   2040
      TabIndex        =   8
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Tax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label LblTot 
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
      Left            =   2040
      TabIndex        =   6
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Price of Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Total Price: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   5280
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCal_Click()

'Declaring Variables
Dim strState As String
Dim strPrice As String
Dim strTot As String
Dim ConTax As String
Dim strShipping As String

'declaring Alaska or Hawaii is $15 and anywhere else is $7
strState = LCase(TxtState.Text)
If strState = "alaska" Or strState = "hawaii" Then
strShipping = 15
Else
strShipping = 7
End If

'assigning variables/finding total price
strState = Val(TxtState.Text)
strPrice = Val(TxtPrice.Text)
strtax = Val(TxtPrice) * 0.056
strTot = Val(strPrice) + strShipping + strtax
LblTot.Caption = strTot
LblTax.Caption = strtax
LblShipping.Caption = strShipping

'formatting to currency
LblTot.Caption = Format(LblTot.Caption, "currency")
LblShipping.Caption = Format(LblShipping.Caption, "currency")
LblTax.Caption = Format(LblTax.Caption, "currency")

'correcting false imput from user
If IsNumeric(TxtPrice.Text) = False Then
MsgBox "Please enter a valid number"
TxtPrice.SelStart = 0
TxtPrice.SelLength = Len(TxtPrice.Text)
TxtPrice.SetFocus
End If

If (TxtState.Text) = "" Then
MsgBox "Please enter a state"
TxtState.SelLength = 0
TxtState.SelLength = Len(TxtState.Text)
TxtState.SetFocus
End If
End Sub

