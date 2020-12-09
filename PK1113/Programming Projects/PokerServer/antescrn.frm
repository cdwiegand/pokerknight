VERSION 5.00
Begin VB.Form antescrn 
   Caption         =   "Ante Information"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1890
   ScaleWidth      =   5865
   Begin VB.CommandButton Command2 
      Caption         =   "Set Ante:"
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
      Left            =   2160
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "500"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.OptionButton Option5 
      Caption         =   "250"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.OptionButton Option4 
      Caption         =   "100"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.OptionButton Option3 
      Caption         =   "50"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      Caption         =   "10"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Automatically Increment Ante Each 5 Hands."
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play Screen"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   4200
      Picture         =   "antescrn.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Automatic Ante Increase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "10"
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Current Ante:"
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
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "antescrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
playscreen.Hide
playscreen.Show
End Sub
