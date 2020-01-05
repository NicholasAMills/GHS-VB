VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3690
      Left            =   5160
      TabIndex        =   3
      Top             =   0
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3960
      Width           =   3135
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Browse Song"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   3960
      Width           =   1935
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   8493
      _cy             =   6376
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd1_Click()
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName
End Sub


Private Sub File1_Click()

File1.Path = "C:\Users\Public\Music\Sample Music"
If File1.CausesValidation = True Then
Text1.Text = File1.Path + "\" + File1.FileName
End If
End Sub

Private Sub Text1_Change()
WindowsMediaPlayer1.URL = Text1.Text
End Sub

