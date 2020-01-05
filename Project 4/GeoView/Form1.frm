VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmGeoView 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GeoView"
   ClientHeight    =   4515
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   2355
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   2355
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Users\dwing\Desktop\GeoView\World.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Countries"
      Top             =   4170
      Width           =   2355
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   1
      FontName        =   "MS Sans Serif"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      DataField       =   "Government"
      DataSource      =   "Data1"
      Height          =   195
      Index           =   5
      Left            =   600
      TabIndex        =   5
      Top             =   3600
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Government"
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      DataField       =   "Capital"
      DataSource      =   "Data1"
      Height          =   195
      Index           =   3
      Left            =   600
      TabIndex        =   3
      Top             =   2640
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Capital"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   465
   End
   Begin VB.Image Image1 
      DataField       =   "Flag"
      DataSource      =   "Data1"
      Height          =   960
      Left            =   480
      Stretch         =   -1  'True
      Top             =   0
      Width           =   960
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "&Select"
      Begin VB.Menu mnuNext 
         Caption         =   "&Next"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuPrev 
         Caption         =   "&Previous"
         Shortcut        =   ^P
      End
      Begin VB.Menu Separator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFirst 
         Caption         =   "&First"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuLast 
         Caption         =   "&Last"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuView 
         Caption         =   "&View"
         Begin VB.Menu mnuName 
            Caption         =   "&Name"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuCapital 
            Caption         =   " &Capital"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuGovt 
            Caption         =   "&Government"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuFont 
         Caption         =   "&Font"
      End
   End
End
Attribute VB_Name = "frmGeoView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnuOptions

End Sub

Private Sub mnuExit_Click()
Unload frmGeoView
End Sub

Private Sub mnuFirst_Click()
Data1.Recordset.MoveFirst
End Sub

Private Sub mnuFont_Click()
'display font dialog box
CommonDialog1.ShowFont
'assign font properties to labels in array
For X = 0 To 5
    With Label1(X)
     .FontName = CommonDialog1.FontName
     .FontItalic = CommonDialog1.FontItalic
     .FontBold = CommonDialog1.FontBold
     .FontSize = CommonDialog1.FontSize
    End With
Next
'adjust form width to new label width
Data1.Recordset.MoveFirst
frmGeoView.Width = Label1(5).Left + Label1(5).Width + Label1(4).Left
'center image right to left
Image1.Left = (frmGeoView.Width - Image1.Width) / 2

End Sub

Private Sub mnuGovt_Click()
ToggleView mnuGovt, 5
End Sub

Private Sub mnuLast_Click()
Data1.Recordset.MoveLast
End Sub

Private Sub mnuNext_Click()
Data1.Recordset.MoveNext
If Data1.Recordset.BOF = True Then Data1.Recordset.MoveFirst
End Sub

Private Sub mnuPrev_Click()
Data1.Recordset.MovePrevious
If Data1.Recordset.BOF = True Then Data1.Recordset.MoveLast
End Sub

Public Sub ToggleView(ControlName As Control, LabelIndex As Integer)
'toggle display of the menu check mark
ControName.Checked = Not ControlName.Checked
'toggle display of the label
Label1(Label1Index).Visible = Not Label1(Label1Index).Visible
End Sub
