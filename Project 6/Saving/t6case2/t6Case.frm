VERSION 5.00
Begin VB.Form frmPao 
   Caption         =   "Political Awareness Organization"
   ClientHeight    =   4650
   ClientLeft      =   1080
   ClientTop       =   1515
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4650
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstParty 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print &Report"
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "&Display Totals"
      Height          =   495
      Left            =   5760
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "&Enter Information"
      Default         =   -1  'True
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   5055
      Begin VB.Label lblInd 
         Alignment       =   2  'Center
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
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   27
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblRep 
         Alignment       =   2  'Center
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
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   26
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblDem 
         Alignment       =   2  'Center
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
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   25
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblInd 
         Alignment       =   2  'Center
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
         Height          =   375
         Index           =   2
         Left            =   3360
         TabIndex        =   24
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblRep 
         Alignment       =   2  'Center
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
         Height          =   375
         Index           =   2
         Left            =   3360
         TabIndex        =   23
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblDem 
         Alignment       =   2  'Center
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
         Height          =   375
         Index           =   2
         Left            =   3360
         TabIndex        =   22
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblInd 
         Alignment       =   2  'Center
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
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   21
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblRep 
         Alignment       =   2  'Center
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
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   20
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblDem 
         Alignment       =   2  'Center
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
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblInd 
         Alignment       =   2  'Center
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
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   18
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblRep 
         Alignment       =   2  'Center
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
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   17
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblDem 
         Alignment       =   2  'Center
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
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   16
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Over 65"
         Height          =   195
         Left            =   4320
         TabIndex        =   15
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "51 - 65"
         Height          =   195
         Left            =   3360
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "36 - 50"
         Height          =   195
         Left            =   2400
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "18 - 35"
         Height          =   195
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Independent:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Republican:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Democrat:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.ListBox lstAge 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Age:"
      Height          =   195
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Political party:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmPao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDisplay_Click()
    Dim strParty As String, intAge As Integer, intCount As Integer
    For intCount = 0 To 3
        lblDem(intCount).Caption = ""
        lblRep(intCount).Caption = ""
        lblInd(intCount).Caption = ""
    Next intCount
    Open "c:\users\dwing\desktop\Save Open\pao.dat" For Input As #1
    Do While Not EOF(1)
        Input #1, strParty, intAge
    Select Case strParty
        Case "Democrat"
            lblDem(intAge).Caption = Val(lblDem(intAge).Caption) + 1
        Case "Republican"
            lblRep(intAge).Caption = Val(lblRep(intAge).Caption) + 1
        Case "Independent"
            lblInd(intAge).Caption = Val(lblInd(intAge).Caption) + 1
    End Select
Loop
Close #1                              'Closes the file
        
End Sub

Private Sub cmdEnter_Click()
Open "c:\users\dwing\desktop\Save Open\pao.dat" For Append As #1           'open the sequential file
    Write #1, lstParty.Text, lstAge.ListIndex     ' write the record
    Close #1                                      ' close the file
End Sub

Private Sub cmdExit_Click()
    Unload frmPao
End Sub


Private Sub cmdPrint_Click()
    Exit Sub    'you will remove this line in Lesson A
    Dim intX As Integer, intDem As Integer, intRep As Integer
    Dim intInd As Integer, intTotal As Integer
    Dim strFont As String, sngSize As Single
    Dim strPS1 As String * 3, strPS2 As String * 3, strPS3 As String * 3
    Dim strPS4 As String * 3, strPS5 As String * 4
    
    'accumulate totals
    For intX = 0 To 3
        intDem = intDem + Val(lblDem(intX).Caption)
        intRep = intRep + Val(lblRep(intX).Caption)
        intInd = intInd + Val(lblInd(intX).Caption)
    Next intX
    intTotal = intDem + intRep + intInd
    
    strFont = Printer.Font          'save current printer settings
    sngSize = Printer.FontSize
    Printer.Font = "courier new"    'change printer settings
    Printer.FontSize = 10           'print title and headings
    Printer.Print Tab(30); "PAO Information - 1999"
    Printer.Print
    Printer.Print Tab(5); "Party"; Tab(20); "18-35"; Tab(30); "36-50"; _
                  Tab(40); "51-65"; Tab(50); "Over 65"; Tab(60); "Total"
    'align democrat numbers and print
    RSet strPS1 = Format(lblDem(0).Caption, "general number")
    RSet strPS2 = Format(lblDem(1).Caption, "general number")
    RSet strPS3 = Format(lblDem(2).Caption, "general number")
    RSet strPS4 = Format(lblDem(3).Caption, "general number")
    RSet strPS5 = Format(intDem, "general number")
    Printer.Print Tab(5); "Democrat"; Tab(22); strPS1; Tab(32); strPS2; _
                    Tab(42); strPS3; Tab(54); strPS4; Tab(61); strPS5
    'align republican numbers and print
    RSet strPS1 = Format(lblRep(0).Caption, "general number")
    RSet strPS2 = Format(lblRep(1).Caption, "general number")
    RSet strPS3 = Format(lblRep(2).Caption, "general number")
    RSet strPS4 = Format(lblRep(3).Caption, "general number")
    RSet strPS5 = Format(intRep, "general number")
    Printer.Print Tab(5); "Republican"; Tab(22); strPS1; Tab(32); strPS2; _
                    Tab(42); strPS3; Tab(54); strPS4; Tab(61); strPS5
    'align independent numbers and print
    RSet strPS1 = Format(lblInd(0).Caption, "general number")
    RSet strPS2 = Format(lblInd(1).Caption, "general number")
    RSet strPS3 = Format(lblInd(2).Caption, "general number")
    RSet strPS4 = Format(lblInd(3).Caption, "general number")
    RSet strPS5 = Format(intInd, "general number")
    Printer.Print Tab(5); "Independent"; Tab(22); strPS1; Tab(32); strPS2; _
                    Tab(42); strPS3; Tab(54); strPS4; Tab(61); strPS5
    Printer.Print                           'print two blank lines
    Printer.Print
    'print grand total
    RSet strPS5 = Format(intTotal, "general number")
    Printer.Print Tab(41); "Total respondents"; Tab(61); strPS5
    Printer.Print                           'print a blank line
    Printer.Print Tab(5); "End of report"   'print message
    Printer.EndDoc                          'send report to printer
    Printer.Font = strFont
    Printer.FontSize = sngSize
End Sub

Private Sub Form_Load()
    frmPao.Top = (Screen.Height - frmPao.Height) / 2
    frmPao.Left = (Screen.Width - frmPao.Width) / 2
    lstAge.AddItem "18 - 35"
    lstAge.AddItem "36 - 50"
    lstAge.AddItem "51 - 65"
    lstAge.AddItem "Over 65"
    lstParty.AddItem "Republican"
    lstParty.AddItem "Independent"
    lstParty.AddItem "Democrat"
    lstParty.ListIndex = 0
    lstAge.ListIndex = 1
    
End Sub

