VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMovies 
   Caption         =   "Erin's Movie Plaza"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   8280
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkMatinee 
      Caption         =   "Matinee Discount"
      Height          =   735
      Left            =   3240
      TabIndex        =   11
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox cboMovie 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   960
      Width           =   6495
   End
   Begin VB.TextBox txtRecord 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   6480
      Width           =   6375
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   7
      Top             =   8160
      Width           =   2415
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "&Enter"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   6
      Top             =   8160
      Width           =   2415
   End
   Begin VB.Frame fraTickets 
      Caption         =   "# of Tickets"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   840
      TabIndex        =   1
      Top             =   3000
      Width           =   6495
      Begin VB.OptionButton optNumTicket 
         Caption         =   "4"
         Height          =   375
         Index           =   4
         Left            =   5400
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton optNumTicket 
         Caption         =   "3"
         Height          =   375
         Index           =   3
         Left            =   3840
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton optNumTicket 
         Caption         =   "2"
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton optNumTicket 
         Caption         =   "1"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton optNumTicket 
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   2760
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Transaction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      TabIndex        =   9
      Top             =   6000
      Width           =   1470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total Tickets"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   5
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount Due"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   4
      Top             =   4440
      Width           =   1590
   End
   Begin VB.Label lblTotal 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4440
      TabIndex        =   3
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label lblAmtDue 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   840
      TabIndex        =   2
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label lblmovie 
      AutoSize        =   -1  'True
      Caption         =   "Movie Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2010
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuchange 
         Caption         =   "&Change File"
      End
      Begin VB.Menu mnuclear 
         Caption         =   "C&lear"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuenter 
      Caption         =   "&Enter"
      Begin VB.Menu mnuenter2 
         Caption         =   "E&nter Record"
      End
   End
End
Attribute VB_Name = "frmMovies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim intNum As Integer, intTotal As Integer
    Dim curPrice As Currency, curAmtDue As Currency
    Dim strMovie As String
    Dim strFileName As String
    
Private Sub cboMovie_LostFocus()
    If cboMovie.ListIndex = 0 Then
        MsgBox ("Please check a movie.")
        cboMovie.SetFocus
    End If
End Sub

Private Sub cmdEnter_Click()
    lblAmtDue.Caption = curAmtDue
     txtRecord.Text = intNum & vbTab & cboMovie.Text & vbTab & curAmtDue & vbNewLine & txtRecord.Text
    
    'save code goes here
    'Save everything in the text box...number of tickets, the name of the movie
    'and the total price
    Dim strMovie As String
    
    strMovie = txtRecord.Text
    
    Dim infotix As String
    Dim infoTotal As String
    Dim infoMovie As String
    
    On Error GoTo Err
    
    If strFileName = "" Then
        MsgBox " File Not Saved", vbInformation + vbOKOnly, "File Not Saved"
    End If
    
    
    Open strFileName For Append As #1
    
    Write #1, strMovie
    
    Close #1
    
    Exit Sub
Err:
    MsgBox "File Not Saved", vbOKOnly + vbInformation, "Error Handler"
    
    
    lblTotal.Caption = ""
    chkMatinee.Value = vbUnchecked
    lblAmtDue = " "
    
    
    cboMovie.ListIndex = 0
  
    optNumTicket(0).Value = True
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub Form_Load()
    
    strFileName = InputBox("Please enter the path with the file name (.dat)")
    'This is to set the file path for save and open
    cboMovie.AddItem "Click a Flick"
    cboMovie.AddItem "O Brother Where Art Thou"
    cboMovie.AddItem "Remember The Titans"
    cboMovie.AddItem "Beautiful Mind"
    cboMovie.AddItem "Ever After"
    cboMovie.AddItem "Ferris Buller's Day Off"
    cboMovie.ListIndex = 0
    
End Sub

Private Sub mnuchange_Click()
strFileName = InputBox("Enter file name:")
End Sub

Private Sub mnuclear_Click()
txtRecord.Text = ""
End Sub

Private Sub mnuenter2_Click()
Call cmdEnter_Click
End Sub

Private Sub mnuexit_Click()

dlgCommon.ShowSave
dlgCommon.Flags = cdlCFScreenFonts
dlgCommon.ShowFont
dlgCommon.ShowColordlgCommon.ShowPrinter

End

End Sub

Private Sub mnuopen_Click()
    dlgCommon.Filter = "Data Files (*.dat)|*.dat|All Files (*.*)|*.*"
    dlgCommon.FileName = " "
    dlgCommon.ShowOpen
        
    'open code goes here
    'after opening a file the text box should contain the number of tickets, the name of the movie
    'and the total price
    
    Dim strinfoget As String
    On Error GoTo Err
    
    Open strFileName For Input As #1
    Do While Not EOF(1)
    Input #1, strinfoget
    txtRecord.Text = txtRecord.Text & strinfoget
    Loop
    Close #1
    Exit Sub
Err:
    MsgBox "File Not Saved", vbOKOnly + vbInformation, "Error Handler"
    Exit Sub
        
End Sub

Private Sub optNumTicket_Click(Index As Integer)
    intNum = Index
    If chkMatinee.Value = vbChecked Then
        curPrice = 4.5
    Else
        curPrice = 6.5
    End If
    curAmtDue = intNum * curPrice
    lblAmtDue.Caption = Format(curAmtDue, "Currency")
    lblTotal.Caption = intNum
End Sub
