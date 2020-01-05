VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCommon 
   Caption         =   "Common Dialog Control Examples"
   ClientHeight    =   4230
   ClientLeft      =   1605
   ClientTop       =   1650
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4230
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   240
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtEdit 
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   480
      TabIndex        =   6
      Text            =   "Text"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "&Color Dialog"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "&Font Dialog"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrinter 
      Caption         =   "&Print Dialog"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save &As Dialog"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open Dialog"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFileName As String
Dim strText As String

Private Sub cmdColor_Click()
    dlgCommon.Flags = cdlCCRGBInit
    dlgCommon.Color = txtEdit.ForeColor
    dlgCommon.ShowColor
    txtEdit.ForeColor = dlgCommon.Color
End Sub

Private Sub cmdExit_Click()
    Unload frmCommon
End Sub


Private Sub cmdFont_Click()
    dlgCommon.Flags = cdlCFScreenFonts
        
    dlgCommon.FontName = txtEdit.FontName
    dlgCommon.FontBold = txtEdit.FontBold
    dlgCommon.FontItalic = txtEdit.FontItalic
    dlgCommon.FontSize = txtEdit.FontSize
    
    dlgCommon.ShowFont
    
    txtEdit.FontName = dlgCommon.FontName
    txtEdit.FontBold = dlgCommon.FontBold
    txtEdit.FontItalic = dlgCommon.FontItalic
    txtEdit.FontSize = dlgCommon.FontSize
    
End Sub

Private Sub cmdOpen_Click()
    
    dlgCommon.Filter = "Text Files(*.txt)|*.txt|All Files(*.*)|*.*"
    dlgCommon.FileName = " "
    
    dlgCommon.ShowOpen
    strFileName = dlgCommon.FileName
    txtEdit.Text = ""
    
    Dim strInfoGet As String
    On Error GoTo Err
    Open strFileName For Input As #1
    
    Do While Not EOF(1)
    Input #1, strInfoGet
    txtEdit.Text = txtEdit.Text & strInfoGet
    Loop
    Close #1
    Exit Sub
Err:
    MsgBox "File Not Opened", vbCritical + vbOKOnly, "Error Handler"
    
End Sub

Private Sub cmdPrinter_Click()
    dlgCommon.Flags = cdlPDNoSelection + cdlPDHidePrintToFile
    dlgCommon.ShowPrinter
    
    
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Err
    strText = txtEdit.Text
    dlgCommon.Filter = "Text Files (*.txt)|*.txt"
    dlgCommon.ShowSave
    
    Dim strInfoSave As String
    strFileName = dlgCommon.FileName
    
    If strText = "" Then
        MsgBox "File Not Saved", vbInformation + vbOKOnly, "File Not Saved"
    End If
    
    strInfoSave = txtEdit.Text
    Open strFileName For Output As #1
    Write #1, strInfoSave
    Close #1
    Exit Sub
Err:
    MsgBox "File Not Saved", vbOKOnly + vbInformation, "Error Handler"
    
End Sub

Private Sub Form_Load()
    frmCommon.Top = (Screen.Height - frmCommon.Height) / 2
    frmCommon.Left = (Screen.Width - frmCommon.Width) / 2
End Sub

Private Sub txtEdit_GotFocus()
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit.Text)
End Sub

