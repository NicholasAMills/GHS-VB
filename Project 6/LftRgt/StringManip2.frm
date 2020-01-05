VERSION 5.00
Begin VB.Form frmSweatshirts 
   Caption         =   "Environmental Sweatshirts"
   ClientHeight    =   3480
   ClientLeft      =   1410
   ClientTop       =   1530
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3480
   ScaleWidth      =   5865
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "&Display"
      Default         =   -1  'True
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   1560
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
      Left            =   1920
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblDirections 
      Caption         =   "Enter 100, 200, 300 for Design.         Enter XS, SM, MD, LG, XL for size.  Example: 100XS"
      Height          =   735
      Left            =   480
      TabIndex        =   7
      Top             =   720
      Width           =   5175
   End
   Begin VB.Image imgSun 
      Height          =   480
      Left            =   4920
      Picture         =   "StringManip2.frx":0000
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSnow 
      Height          =   480
      Left            =   4320
      Picture         =   "StringManip2.frx":0442
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEarth 
      Height          =   480
      Left            =   3720
      Picture         =   "StringManip2.frx":0884
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Size:"
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
      Left            =   1800
      TabIndex        =   6
      Top             =   2400
      Width           =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Design:"
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
      TabIndex        =   5
      Top             =   2400
      Width           =   810
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.Image imgDesign 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   360
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   855
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
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1590
   End
End
Attribute VB_Name = "frmSweatshirts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDisplay_Click()
    Const conBtns As Integer = vbOKOnly + vbInformation + vbDefaultButton1 + vbApplicationModal
    Const conMsgl As String = "The first 3 characters must be 100, 200, or 300."
    Const ConMsg2 As String = "The last 2 charactrs must be XS, SM, MD, LG, XL."
    Dim strDesign As String, strSize As String, intRetVal As Integer
    strDesign = Left(txtCode.Text, 3)          'assign leftmost 3 characters
    strSize = UCase(Right(txtCode.Text, 2))     'assign rightmost 2 characters
    Select Case strDesign                       'determine appropriate picture
        Case "100"
            imgDesign.Picture = imgEarth.Picture
        Case "200"
            imgDesign.Picture = imgSnow.Picture
        Case "300"
            imgDesign.Picture = imgSun.Picture
        Case Else
            intRetVal = MsgBox(conMsgl, conBtns, "Design Error")
        End Select
    
    Select Case strSize                         'determine shirt size
        Case "XD", "SM", "MD", "LG", "XL"
            lblSize.Caption = strSize
        Case Else
            intRetVal = MsgBox(ConMsg2, conBtns, "Size Error")
        End Select
        txtCode.SelStart = 0                    'highlight text in text box
        txtCode.SelLength = Len(txtCode.Text)
        txtCode.SetFocus                         'set Focus
            
        
    
 
End Sub


Private Sub cmdExit_Click()
    Unload frmSweatshirts
End Sub

Private Sub Form_Load()
    frmSweatshirts.Top = (Screen.Height - frmSweatshirts.Height) / 2
    frmSweatshirts.Left = (Screen.Width - frmSweatshirts.Width) / 2
End Sub

Private Sub txtCode_Change()
    imgDesign.Picture = LoadPicture()
    lblSize.Caption = ""
End Sub

