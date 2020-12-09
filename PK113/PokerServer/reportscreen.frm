VERSION 5.00
Begin VB.Form reportscreen 
   Caption         =   "Report Log"
   ClientHeight    =   1470
   ClientLeft      =   5250
   ClientTop       =   5685
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1470
   ScaleWidth      =   5880
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "reportscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HideButton_Click()
    Unload Me
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
End Sub

