VERSION 5.00
Begin VB.Form frmConsolidated 
   Caption         =   "Consolidated Machinery"
   ClientHeight    =   2685
   ClientLeft      =   1350
   ClientTop       =   1515
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2685
   ScaleWidth      =   4470
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "&Verify"
      Default         =   -1  'True
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Product code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1590
   End
End
Attribute VB_Name = "frmConsolidated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdVerify_Click()
    Dim strCode As String
    strCode = Mid(txtCode.Text, 3, 1)   'assign third character to strCode
    Select Case UCase(strCode)
        Case "A", "C"
            lblStatus.Caption = "The product code is valid."
        Case Else
            lblStatus.Caption = "The product code is not valid."
    End Select
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.Text)
    txtCode.SetFocus
End Sub


Private Sub cmdExit_Click()
    Unload frmConsolidated
End Sub

Private Sub Form_Load()
    frmConsolidated.Top = (Screen.Height - frmConsolidated.Height) / 2
    frmConsolidated.Left = (Screen.Width - frmConsolidated.Width) / 2
End Sub

Private Sub txtCode_Change()
    lblStatus.Caption = ""
End Sub
