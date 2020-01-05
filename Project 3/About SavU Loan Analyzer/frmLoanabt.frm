VERSION 5.00
Begin VB.Form frmLoanabt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About SavU Loan Analyzer"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   960
      X2              =   4320
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmLoanabt.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmLoanabt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
'remove about dialog form
Unload frmLoanabt

End Sub

Private Sub Form_Load()
'crate captions for labels
Label1.Caption = _
    "Savu national Bank Loan Analyzer" & vbNewLine & _
    "Operation system 98" & vbNewLine & _
    "Copyright 1999 SavU National Bank Corp."
Label2.Caption = _
    "Developed for SavU National Bank" & vbNewLine & _
    "By Sarah Carter"
Label3.Caption = _
    "Warning: This computer program is protected by" & vbNewLine & _
    "copyright law and international treaties."
End Sub
