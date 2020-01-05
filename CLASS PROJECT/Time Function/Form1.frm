VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   3120
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Record"
      Height          =   855
      Left            =   3840
      TabIndex        =   0
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "List of Lightning Strikes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label2 
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
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Most Recent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim dtmTest As Date
dmtTest = TimeValue(Now)
Label2.Caption = dmtTest
List1.AddItem (Label2.Caption)

End Sub

