VERSION 5.00
Begin VB.Form textview 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clipboard"
   ClientHeight    =   3645
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "textview.frx":0000
      Top             =   0
      Width           =   5775
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "textview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub filedump()
Dim iFile As Integer
Dim sfile As String
Dim lineread As String
Text1.Text = ""
textview.Caption = "pklog." & Format(Now(), "yyy-mm-dd") & ".txt."
sfile = App.Path & "\logs\pklog." & Format(Now(), "yyy-mm-dd") & ".txt."
Open sfile For Input As #1
Do While Not EOF(1)
Input #1, lineread
Text1.Text = Text1.Text & lineread _
& vbCrLf & " "
Loop
Close #1
End Sub

Private Sub Command3_Click()
filedump
End Sub

Private Sub form_Load()
textview.Caption = "pklog." & Format(Now(), "yyy-mm-dd") & ".txt."
filedump
End Sub

Private Sub Command1_Click()
Text1.Text = ""
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub
