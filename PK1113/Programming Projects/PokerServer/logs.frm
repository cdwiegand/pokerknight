VERSION 5.00
Begin VB.Form logs 
   Caption         =   "Log Files"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1620
   ScaleWidth      =   6030
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View the Log"
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Administrator Notes Append:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "None"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Log File:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   4320
      Picture         =   "logs.frx":0000
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "logs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
textview.Show

End Sub

Private Sub form_Load()
Label3.Caption = "pklog." & Format(Now(), "yyy-mm-dd") & ".txt."
End Sub

Private Sub Command2_Click()
Dim sentmess As String
sentmess = "<( Log Append Note )>: " & Text1.Text
writeup sentmess
Text1.Text = ""
End Sub



