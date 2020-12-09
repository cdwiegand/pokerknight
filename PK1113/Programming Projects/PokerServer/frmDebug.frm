VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Debug Log"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "CardDiag"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
dbcarddiag.Hide
dbcarddiag.Show
End Sub

Private Sub Form_Resize()
    Me.List1.Left = 0
    Me.List1.Width = Me.ScaleWidth
    Me.List1.Top = 0
    Me.List1.Height = Me.ScaleHeight
End Sub

Public Sub AddItem(ByVal sString As String, ByVal iIgnoreThis As Integer)
    Me.List1.AddItem sString, 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
End Sub

