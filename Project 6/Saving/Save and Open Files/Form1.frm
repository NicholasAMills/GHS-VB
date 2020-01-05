VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   240
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFile 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'use Form Level
Dim strfileName As String

Private Sub mnuExit_click()
    Unload Me
End Sub

Private Sub mnuOpen_click()
    Dim strInfoGet As String
    On Error GoTo Err
    'Dialog is the name of teh Common Dialog Control on the form
    Dialog.CancelError = True
    'using cancel creates an error so we can catch it
    Dialog.Filter = "Text Files (*.txt)  *.txt"
    'This is the files that apear in the bottom combo box of the  save window
    'so the user can only save text files
    Dialog.DialogTitle = "Open Text File"
    'Set the Dialog Title
    Dialog.ShowOpen
    'Show the Open Dialog
    strfileName = Dialog.FileName
    'strFileName is the title of the file plus the filtered extension
    txtFile.Text = ""
    'Clear the text box so opening is nice and clean
    
    Open strfileName For Input As #1
    'Open strFileName for Input (Input Means input into a box)
    
    Do While Not EOF(1)
    'Looop the 1 in () and the input#1 coincide
    Input #1, strInfoGet
    'Open the file and save the input into a variable
    txtFile.Text = txtFile.Text & vbNewLine & strInfoGet
    'enter info into text box
    Loop
    Close #1
    'Close #1
Exit Sub
'Error Handler
Err:
    MsgBox "File Not Opened", vbCritical + vbOKOnly, "Error Handler"
    
    Exit Sub
End Sub

Private Sub mnuSAve_Click()
Dim strInfoSAve As String
    On Error GoTo Err
    Dialog.CancelError = True
    Dialog.Filter = "Text Files (*.txt)  *.txt"
    Dialog.DialogTitle = "Save Text File"
    Dialog.ShowSave
    
    strfileName = Dialog.FileName
    
    If strfileName = "" Then
        MsgBox "File Not Saved", vbInformation + vbOKOnly, "File Not Saved"
    End If
    
    strInfoSAve = txtFile.Text
    'open the file in the strFileName string for "appened" - the variable it will
    'be saved as is #1 (#1 will be used to refer to the file in the code)
    'Append is used to change the file without erasing it - (editing)
    Open strfileName For Append As #1
    'Write to the file saved in the variable #1 (strFileName) - each argument
    'separated by commas will be writteen to the file (IN THE SAME LINE)
    
    Write #1, strInfoSAve
    'CLOSE the file strFileName (a.k.a. #1)
    Close #1
    Exit Sub
Err:
    MsgBox "File Not Saved", vbOKOnly + vbInformation, "Error Handler"
    
End Sub
    

