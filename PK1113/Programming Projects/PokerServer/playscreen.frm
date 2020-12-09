VERSION 5.00
Begin VB.Form playscreen 
   Caption         =   "Play Screen"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1935
   ScaleWidth      =   5790
   Begin VB.CommandButton Command6 
      Caption         =   "Deal"
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   840
      Width           =   735
   End
   Begin VB.Frame Frame5 
      Caption         =   "Raise"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   18
      Top             =   1080
      Width           =   1095
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Players"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   1695
      Begin VB.ListBox List1 
         Height          =   1035
         ItemData        =   "playscreen.frx":0000
         Left            =   120
         List            =   "playscreen.frx":0013
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Collect the ante..."
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ante"
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cover"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refuse"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Accept"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   3015
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "0"
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Cover Items:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Gold:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Game #1"
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
      Left            =   2760
      TabIndex        =   5
      Top             =   0
      Width           =   3015
      Begin VB.Label Label10 
         Caption         =   "BET"
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "CALL"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "BET"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "DEAL"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "ANTE"
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
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1695
      Left            =   1800
      TabIndex        =   21
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "playscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click()
antescrn.Hide
antescrn.Show
End Sub
