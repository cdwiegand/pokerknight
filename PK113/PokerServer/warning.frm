VERSION 5.00
Begin VB.Form warning 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   6375
      Left            =   120
      Picture         =   "warning.frx":0000
      ScaleHeight     =   6315
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton Command2 
         Caption         =   "Quit"
         Height          =   495
         Left            =   3960
         TabIndex        =   2
         Top             =   5520
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Proceed!"
         Default         =   -1  'True
         Height          =   495
         Left            =   2640
         TabIndex        =   1
         Top             =   5520
         Width           =   1215
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
