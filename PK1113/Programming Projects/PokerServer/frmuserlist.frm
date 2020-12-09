VERSION 5.00
Begin VB.Form frmuserlist 
   Caption         =   "User List"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6405
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1725
   ScaleWidth      =   6405
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   4560
      Picture         =   "frmuserlist.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connected Users:"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.ListBox lbUsers 
         Height          =   1035
         ItemData        =   "frmuserlist.frx":66B6
         Left            =   120
         List            =   "frmuserlist.frx":66B8
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "NOTE: Only five clients may be connected to PokerKnight at any time."
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   5055
   End
End
Attribute VB_Name = "frmuserlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
End Sub

Private Sub lbUsers_DblClick()
    Dim s As String
    On Error Resume Next
    If lbUsers.ListIndex = -1 Then Exit Sub ' they didn't select anything, boob!
    
    s = lbUsers.List(lbUsers.ListIndex) ' mike @ 1.2.3.4
    s = Left(s, InStr(s, "@") - 1) ' mike
    s = Trim(s) ' mike (trimmed for spaces)
    Dim f As frmServConn
    Set f = frmMain.GetPlayerForm(s)
    If f Is Nothing Then
    Else
        f.Hide
        f.Show
    End If
End Sub



