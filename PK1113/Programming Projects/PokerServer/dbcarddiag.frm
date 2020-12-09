VERSION 5.00
Begin VB.Form dbcarddiag 
   Caption         =   "Card Load Diagnostic"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1800
   ScaleWidth      =   6555
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Request"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton Command2 
         Caption         =   "King"
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Card"
      Height          =   1455
      Left            =   4800
      TabIndex        =   0
      Top             =   0
      Width           =   1095
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   120
         Picture         =   "dbcarddiag.frx":0000
         ScaleHeight     =   1035
         ScaleWidth      =   795
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "dbcarddiag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Set Picture1.Picture = loadrespicture(99, 0)
End Sub
