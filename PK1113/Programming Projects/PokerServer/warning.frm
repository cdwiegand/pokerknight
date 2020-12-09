VERSION 5.00
Begin VB.Form warning 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   120
      Picture         =   "warning.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton Command1 
         Caption         =   "Accept"
         Height          =   495
         Left            =   3000
         TabIndex        =   2
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Decline"
         Height          =   495
         Left            =   4320
         TabIndex        =   1
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Please use PokerKnight! responsibly and properly."
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   3000
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"warning.frx":5F942
         Height          =   615
         Left            =   2400
         TabIndex        =   4
         Top             =   2280
         Width           =   4335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"warning.frx":5F9F9
         Height          =   1455
         Left            =   2400
         TabIndex        =   3
         Top             =   720
         Width           =   4215
      End
   End
End
Attribute VB_Name = "warning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
frmMain.Show
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Image1_Click()

End Sub
