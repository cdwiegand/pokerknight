VERSION 5.00
Begin VB.Form DeckView 
   Caption         =   "Deck Status"
   ClientHeight    =   1590
   ClientLeft      =   5250
   ClientTop       =   5685
   ClientWidth     =   5910
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1590
   ScaleWidth      =   5910
   Begin VB.CommandButton Command1 
      Caption         =   "Reshuffle"
      Height          =   315
      Left            =   120
      TabIndex        =   60
      Top             =   1200
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Deck Stats"
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
      TabIndex        =   56
      Top             =   0
      Width           =   1575
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cards Held:"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cards Played:"
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cards Left:"
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox Diamonds 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   1
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Text            =   "A"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Diamonds 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   13
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Text            =   "K"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Diamonds 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   12
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Text            =   "Q"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Diamonds 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   11
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Text            =   "J"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Diamonds 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   10
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Text            =   "10"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Diamonds 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   9
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Text            =   "9"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Diamonds 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   8
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Text            =   "8"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Diamonds 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   7
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Text            =   "7"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Diamonds 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   6
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Text            =   "6"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Diamonds 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   5
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Text            =   "5"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Diamonds 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   4
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Diamonds 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   3
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Text            =   "3"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Diamonds 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   2
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Text            =   "2"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Hearts 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   1
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Text            =   "A"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Hearts 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   13
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Text            =   "K"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Hearts 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   12
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "Q"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Hearts 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   11
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Text            =   "J"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Hearts 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   10
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Text            =   "10"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Hearts 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   9
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Text            =   "9"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Hearts 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   8
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "8"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Hearts 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   7
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Text            =   "7"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Hearts 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   6
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Text            =   "6"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Hearts 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   5
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "5"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Hearts 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   4
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Hearts 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   3
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "3"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Hearts 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   2
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "2"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Clubs 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   285
      Index           =   1
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "A"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Clubs 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   285
      Index           =   13
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "K"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Clubs 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   285
      Index           =   12
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "Q"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Clubs 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   285
      Index           =   11
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "J"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Clubs 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   285
      Index           =   10
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "10"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Clubs 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   285
      Index           =   9
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "9"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Clubs 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   285
      Index           =   8
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "8"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Clubs 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   285
      Index           =   7
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "7"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Clubs 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   285
      Index           =   6
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "6"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Clubs 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   285
      Index           =   5
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "5"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Clubs 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   285
      Index           =   4
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Clubs 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   285
      Index           =   3
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "3"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Clubs 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   285
      Index           =   2
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "2"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Spades 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "A"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Spades 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "K"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Spades 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Q"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Spades 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "J"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Spades 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "10"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Spades 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "9"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Spades 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "8"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Spades 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "7"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Spades 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "6"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Spades 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "5"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Spades 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Spades 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "3"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Spades 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "2"
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Diamonds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1800
      TabIndex        =   55
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label HeartLabel 
      Caption         =   "Hearts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2040
      TabIndex        =   54
      Top             =   840
      Width           =   615
   End
   Begin VB.Label ClubLabel 
      Caption         =   "Clubs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   2040
      TabIndex        =   53
      Top             =   480
      Width           =   615
   End
   Begin VB.Label SpadeLabel 
      Caption         =   "Spades"
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
      Left            =   2040
      TabIndex        =   52
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "DeckView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const iPlayer1 = &HFF ' red
Private Const iPlayer2 = &HFF00 ' green
Private Const iPlayer3 = &HFF0000 ' blue
Private Const iPlayer4 = &HFF00FF ' purple
Private Const iPlayer5 = &H8080FF ' pink??

Private Sub Command1_Click()
    If MsgBox("Reshuffle Deck? This will invalidate this hand!", vbYesNo + vbDefaultButton2) = vbYes Then frmMain.SetupDeck
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
End Sub


Public Sub RefreshView()
    Dim MyColor As OLE_COLOR
    Dim s As String
    Dim i As Integer
    Dim i2 As Integer
    Dim iplayed As Integer
    Dim ileft As Integer
    Dim iheld As Integer
    
    
    For i = 1 To 52
        s = frmMain.GetCard(i)
        Debug.Print s
        Select Case Mid(s, 3)
        Case "0"
            MyColor = vbWhite
            ileft = ileft + 1
        Case "-"
            MyColor = &H8F8F8F ' grey
            iplayed = iplayed + 1
        Case Else
            MyColor = &H80FFFF ' yellow
            iheld = iheld + 1
        End Select
        
        Select Case Mid(s, 2, 1)
        Case "A"
            i2 = 1
        Case "K"
            i2 = 13
        Case "Q"
            i2 = 12
        Case "J"
            i2 = 11
        Case "0"
            i2 = 10
        Case "2", "3", "4", "5", "6", "7", "8", "9"
            i2 = Mid(s, 2, 1)
        End Select
        
        
        Select Case Left(s, 1)
        Case "S"
            Me.Spades(i2).BackColor = MyColor
        Case "D"
            Me.Diamonds(i2).BackColor = MyColor
        Case "C"
            Me.Clubs(i2).BackColor = MyColor
        Case "H"
            Me.Hearts(i2).BackColor = MyColor
        End Select
        
    Next i
    
    Me.Label1 = "Cards Left: " & ileft
    Me.Label2 = "Cards Played: " & iplayed
    Me.Label4 = "Cards Held: " & iheld
End Sub



