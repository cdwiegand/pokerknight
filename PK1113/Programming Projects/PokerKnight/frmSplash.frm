VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   4050
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   7080
      Begin VB.CommandButton Command1 
         Caption         =   "Go!"
         Height          =   255
         Left            =   5640
         TabIndex        =   7
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   2520
         TabIndex        =   0
         Top             =   3570
         Width           =   2895
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         DrawStyle       =   1  'Dash
         FillColor       =   &H00C0FFC0&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000005&
         Height          =   2535
         Left            =   840
         Picture         =   "frmSplash.frx":000C
         ScaleHeight     =   2535
         ScaleWidth      =   5415
         TabIndex        =   5
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "Enter your name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H80000009&
         Caption         =   "(c) Copyright 2001, ShadowCastle Software"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   2760
         Width           =   5415
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H80000009&
         Caption         =   "Programmed by Chris Wiegand and Mike Brannon"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   3000
         Width           =   5535
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H80000009&
         Caption         =   "Version 1.1, BETA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3120
         TabIndex        =   2
         Top             =   3240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' |   P O K E R K N I G H T   |
' |   The Online Poker Game   |
' |   (c) Copyright 2002      |
' |   ShadowCastle Software   |
' |   All rights reserved.    |
' =-=-=-=-=-=-=-=-=-=-=-=-=-=-=

' =-=-=-=-=-=
' |  About  |
' =-=-=-=-=-=
'
' Form information: Splash Screen Code.
' Purpose: To greet the use and obtain a name.
' Important Notes: A blank name quits the program.


Option Explicit

' -=-=-=-=-=-=-
' | Functions |
' =-=-=-=-=-=-=

' Go Button Function
' ===================
' Written by: Mike
' Function: Launches the application

Private Sub Command1_Click()
  sUserName = Text1.Text
  
  If sUserName = "" Then End
      Unload Me
    InterfacePanel.Show
End Sub

' Hotkey Function
' ===============
' Written by: Chris
' Function: Detects ENTER as accepting  the string

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(Left(vbCrLf, 1)) Then
    sUserName = Text1.Text
    If sUserName = "" Then End
    Command1_Click
    End If
End Sub

