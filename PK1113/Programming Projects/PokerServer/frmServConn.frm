VERSION 5.00
Begin VB.Form frmServConn 
   Caption         =   "Listening"
   ClientHeight    =   720
   ClientLeft      =   3000
   ClientTop       =   3225
   ClientWidth     =   6075
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   48
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   405
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Close Connection Port"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<<Private Message To This User>>"
      Height          =   195
      Left            =   2640
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Label lblWho 
      AutoSize        =   -1  'True
      Caption         =   "-=- Listening Port -=-"
      Height          =   195
      Left            =   285
      TabIndex        =   1
      Top             =   0
      Width           =   1410
   End
End
Attribute VB_Name = "frmServConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' SERVER CONNECTION FORM
'
' =-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' |   P O K E R K N I G H T   |
' |        S E R V E R        |
' |   The Online Poker Game   |
' |   (c) Copyright 2002      |
' |   ShadowCastle Software   |
' |   All rights reserved.    |
' =-=-=-=-=-=-=-=-=-=-=-=-=-=-=

' =-=-=-=-=-=
' |  About  |
' =-=-=-=-=-=
'
' Form information: Server Connection Form
' Purpose: To provide connection to the Server.
' Important Notes: The user list will be generated here.

' Revision log:
' 11/11/02 - Put all these notes in!


' =-=-=-=-=-=-=-=-=
' |  Declarations |
' =-=-=-=-=-=-=-=-=
Option Explicit
Public lID As Long
Public sName As String
Public sIPAddress As String
Public sAvatarID As String
Private WithEvents theForm As frmSocket
Attribute theForm.VB_VarHelpID = -1


' Form Load Function
' ==================
' Written by: Chris
' Function: Form Load Options
'
Private Sub form_Load()
    Me.Caption = "Listening"
    Set theForm = New frmSocket
    theForm.Listen 13597
    lID = frmMain.GetUniqueID()
End Sub


' Disconnect User Function
' ========================
' Written by: Chris
' Function: Drop the indicated user.
'
Private Sub cmdDisconnect_Click()
    ' disconnect client
    If Me.Caption = "Listening" Then
        If MsgBox("Are you sure you want to stop listening? If you stop, you can't start again until you re-open the program.", vbOKCancel + vbDefaultButton2) = vbCancel Then Exit Sub
        theForm.Winsock1.Close
    Else
        If MsgBox("Are you sure you want to disconnect this client?", vbOKCancel + vbDefaultButton2) = vbCancel Then Exit Sub
           ' frmMain.PlayerDropped sName2, "DROPPED" - Suddenly getting an Error. WTF - Mike
        DropUser
    End If
    Unload Me ' put NOTHING After this!!!!
End Sub


' Drop User Function
' ==================
' Written by: Chris
' Function: Removes the user from the list.
'
Public Sub DropUser()
  Dim sName2 As String
    sName2 = sName ' cache it, we're about to clear it
    Me.SendPacket "C:Disconnect"
    sName = "" ' CLEARED, use sName2 for this guy's (or gal's) (or it's) name/id/whatever
    sIPAddress = ""
    theForm.Winsock1.Close
End Sub


' Report View Function
' ====================
' Written by: Mike
' Function: Brings up the Deck View Screen.
'
Private Sub Command1_Click()
reportscreen.Hide
reportscreen.Show
DeckView.Hide
End Sub


' Deck View Function
' ==================
' Written by: Mike
' Function: Brings up the Deck View Screen.
'
Private Sub Command2_Click()
reportscreen.Hide
DeckView.Hide
DeckView.Show
End Sub


' TextPress Function
' ==================
' Written by: Chris
' Function: Accepts Enter as a hotkey.
'
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(Left(vbCrLf, 1)) Then
        SendPacket "T:PRIVATE:" & Me.Text1
        Me.Text1 = ""
    End If
End Sub


' CloseConn Function
' ==================
' Written by: Chris
' Function: Closes a lost connection.
'
Private Sub theForm_ConnectionClosed()
    frmMain.PlayerDropped sName, "LOST"
    Unload Me ' put NOTHING AFTER THIS!!!!!!! -- chris
End Sub


' ConnectMade Function
' ====================
' Written by: Chris
' Function: Accepts the new player
'
Private Sub theForm_ConnectionMade()
    ' spawn new server
    UpdateCaption "Anonymous", theForm.Winsock1.RemoteHostIP
    Dim t As New frmServConn
    t.Show
    frmMain.PlayerAdded Me
End Sub


' UpdateCaption Function
' ======================
' Written by: Chris
' Function: Updates the caption with the User ID.
'
Public Sub UpdateCaption(ByVal psRemoteName As String, Optional ByVal psIP As String = "")
    sName = psRemoteName
    If psIP <> "" Then sIPAddress = psIP
    Me.lblWho = psRemoteName & " @ " & sIPAddress
    Me.Caption = Me.lblWho
    If Me.lblWho = "-=- Listening Port -=-" Then
    cmdDisconnect.Caption = "Close the Port"
    Else
    cmdDisconnect.Caption = "Disconnect this user"
    Label1.Visible = True
    Text1.Visible = True
    End If
End Sub


' Receive Packet Function
' =======================
' Written by: Chris
' Function: ??
'
Private Sub theForm_PacketReceived(ByVal sPacket As String)
    frmMain.PacketReceived sPacket, lID, Me, sName
End Sub


' SendPacket Function
' ===================
' Written by: Chris
' Function: ??
'
Public Sub SendPacket(ByVal sPacket As String)
    theForm.Send sPacket
End Sub
