VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Enter ID"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblType 
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
      Left            =   4680
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblStats 
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
      Left            =   4680
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Stats:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblInitials 
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
      Left            =   4680
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Initials:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Employee Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Insert ID"
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
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Const conBtns As Integer = vbOKOnly + vbInformation + vbDefaultButton1 + vbApplicationModal
  Const conMsgl As String = "Invalid ID"
  Dim strInitials As String, strStats As String, strType As String, intRetVal As Integer
  strInitials = UCase(Mid(txtID.Text, 2, 2))
  strStats = UCase(Left(txtID.Text, 1))
  strType = UCase(Right(txtID.Text, 1))
  Select Case strStats
   Case "P"
        lblStats.Caption = "Part Time"
    Case "F"
        lblStats.Caption = "Full Time"
    Case Else
        intRetVal = MsgBox(conMsgl, conBtns, "Invalid ID")
        txtID.Text = ""
    Exit Sub
    End Select
Select Case strType
    Case "1"
        lblType.Caption = "New Cars"
    Case "2"
        lblType.Caption = "Used Cars"
    Case Else
        intRetVal = MsgBox(conMsgl, conBtns, "Invalid ID")
       txtID.Text = ""
        Exit Sub
    End Select
    
lblInitials.Caption = strInitials

 

    
  
  
    



End Sub

Private Sub txtID_Change()
lblInitials.Caption = ""
lblStats.Caption = ""
lblType.Caption = ""
End Sub
