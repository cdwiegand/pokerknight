VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5565
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   5475
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7560
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Press Any Key To Begin"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   2040
         TabIndex        =   4
         Top             =   4920
         Width           =   4215
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H80000009&
         Caption         =   "Version 1.0, BETA"
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
         TabIndex        =   3
         Top             =   4560
         Width           =   1815
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
         TabIndex        =   2
         Top             =   4320
         Width           =   5535
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
         Left            =   1200
         TabIndex        =   1
         Top             =   3960
         Width           =   5415
      End
      Begin VB.Image Image1 
         Height          =   3885
         Left            =   1080
         Picture         =   "frmSplash.frx":000C
         Top             =   120
         Width           =   5700
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    warning.Show
End Sub


Private Sub Frame1_Click()
    Unload Me
    warning.Show
End Sub
