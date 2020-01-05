VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   3015
      Left            =   1680
      TabIndex        =   0
      Top             =   3000
      Width           =   4455
      Begin VB.CheckBox ChkItalicize 
         Caption         =   "Italicize"
         Height          =   495
         Left            =   2640
         TabIndex        =   6
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CheckBox ChkUnderline 
         Caption         =   "Underline"
         Height          =   615
         Left            =   2640
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox ChkBold 
         Caption         =   "Bold"
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Opt20 
         Caption         =   "Font 20"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   1335
      End
      Begin VB.OptionButton Opt16 
         Caption         =   "Font 16"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton Opt12 
         Caption         =   "Font 12"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Label LblDisplay 
      Caption         =   "Ice Cream"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()

End Sub

Private Sub ChkBold_Click()
'making the word bold
If ChkBold.Value = vbChecked Then
LblDisplay.FontBold = True
Else
LblDisplay.FontBold = False
End If

End Sub

Private Sub ChkItalicize_Click()
'making the word italicized
If ChkItalicize.Value = vbChecked Then
LblDisplay.FontItalic = True
Else
LblDisplay.FontItalic = False
End If
End Sub

Private Sub ChkUnderline_Click()
'making the word underlined
If ChkUnderline.Value = vbChecked Then
LblDisplay.FontUnderline = True
Else
LblDisplay.FontUnderline = False
End If
End Sub

Private Sub Opt12_Click()
'making the font size 12
LblDisplay.FontSize = 12

End Sub

Private Sub Opt16_Click()
'making the font size 16
LblDisplay.FontSize = 16
End Sub

Private Sub Opt20_Click()
'making the font size 20
LblDisplay.FontSize = 20
End Sub
