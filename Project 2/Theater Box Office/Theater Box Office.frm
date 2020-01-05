VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   Caption         =   "Theater Box Office"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox frmTheater 
      Height          =   1095
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   4200
      Width           =   3015
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      Height          =   615
      Left            =   2040
      TabIndex        =   12
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "# of Tickets"
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   3015
      Begin VB.OptionButton Option1 
         Caption         =   "1"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton Option2 
         Caption         =   "2"
         Height          =   495
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option3 
         Caption         =   "3"
         Height          =   495
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option4 
         Caption         =   "4"
         Height          =   495
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton Option5 
         Caption         =   "5"
         Height          =   495
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CheckBox chkMatinee 
      Caption         =   "Matinee Discount"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   2655
   End
   Begin VB.ComboBox cboShow 
      Height          =   315
      ItemData        =   "Theater Box Office.frx":0000
      Left            =   360
      List            =   "Theater Box Office.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblAmtdue 
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   3840
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Show Selection"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1110
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   360
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Due"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   3000
      Width           =   885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

End Sub

Private Sub Text1_Change()

End Sub
