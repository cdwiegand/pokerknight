VERSION 5.00
Begin VB.Form mainmen 
   Caption         =   "Main Menu"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1995
   ScaleWidth      =   5700
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Shuffle"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   360
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   4800
      Picture         =   "mainmen.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   240
      Picture         =   "mainmen.frx":2B62
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Logs"
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Commands"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Deck"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Views"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   3495
      Begin VB.CommandButton Command6 
         Caption         =   "Users"
         Height          =   255
         Left            =   2280
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Report"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Copyright (c) 2002, ShadowCastle Software. All Rights Reserved"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   4695
   End
End
Attribute VB_Name = "mainmen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
logs.Hide
logs.Show
End Sub

Private Sub Command2_Click()
Unload Me
DeckView.Hide
DeckView.Show
End Sub

Private Sub Command5_Click()
Unload Me
reportscreen.Hide
reportscreen.Show
End Sub
