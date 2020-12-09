VERSION 5.00
Begin VB.Form about 
   BackColor       =   &H80000009&
   Caption         =   "About"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2190
   ScaleWidth      =   6270
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   1095
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   3375
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Chris Wiegand"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Mike Brannon"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Programmed by:"
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
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Designed by:"
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
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Mike Brannon"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PokerKnight! Server"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "BETA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1560
      Left            =   0
      Picture         =   "about.frx":0000
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

