VERSION 5.00
Begin VB.Form frmuserlist 
   Caption         =   "User List"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1560
   ScaleWidth      =   4785
   Begin VB.Frame Frame1 
      Caption         =   "Connected Users:"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.ListBox lbUsers 
         Height          =   1035
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
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



