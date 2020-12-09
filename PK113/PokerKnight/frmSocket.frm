VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSocket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2070
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   2070
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' SOCKET CODE
'
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
' Form information: Form Socket Code.
' Purpose: To handle the connectivity.
' Important Notes: Winsock stuff.

' =-=-=-=-=-=-=-=-=
' |  Declarations |
' =-=-=-=-=-=-=-=-=

Option Explicit
Private sBuffer As String
Public Event ConnectionClosed()
Public Event PacketReceived(ByVal sPacket As String)
Public Event ConnectionMade()

' -=-=-=-=-=-=-
' | Functions |
' =-=-=-=-=-=-=

' Connect Function
' ================
' Written by: Chris
' Function: Attempts to connect to the provided ISP.

Public Function Connect(ByVal sRemoteIP As String, ByVal nRemotePort As Long) As Boolean
    Dim dDate As Date
    Dim bExit As Boolean
    Me.Winsock1.Close
    Me.Winsock1.Connect sRemoteIP, nRemotePort
    dDate = Now
    Do Until bExit
        If Me.Winsock1.State = 7 Then bExit = True
        If Me.Winsock1.State = 6 Then bExit = True
        If Me.Winsock1.State = 9 Then bExit = True
        If Abs(DateDiff("s", Now, dDate)) > 60 Then bExit = True
        DoEvents
        DoEvents
    Loop
    Connect = bExit
End Function


' Send Function
' ================
' Written by: Chris
' Function: Sends a command to the server socket.

Public Sub Send(ByVal sPacket As String)
    If Winsock1.State = 7 Then Me.Winsock1.SendData sPacket & vbCrLf
    DoEvents
End Sub


' Disconnect Function
' ================
' Written by: Chris
' Function: Closes the winsock connection.

Public Sub Disconnect()
    Me.Winsock1.Close
End Sub


' Connect Function
' ================
' Written by: Chris
' Function: Listens for connection.
' Notes: Obsolete. We don't listen here, but this will be useful to keep in case we do a point-to-point connection.

Public Sub Listen(ByVal nLocalPort As Long)
    Me.Winsock1.Close
    Me.Winsock1.LocalPort = nLocalPort
    Me.Winsock1.Listen
End Sub


' Close Winsock Event
' ===================
' Written by: Chris
' Function: Sends a Connection Closed Event.

Private Sub Winsock1_Close()
    RaiseEvent ConnectionClosed
End Sub


' ConnectionRequest Function
' ==========================
' Written by: Chris
' Function: Connects a remote host.
' Notes: Another obsolete function, related to PtP connection.

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    Winsock1.Close
    Winsock1.Accept requestID
    DoEvents
    RaiseEvent ConnectionMade
End Sub


' DataArrival Function
' ==========================
' Written by: Chris
' Function: This receives and dissects a packet from the server.

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim s As String
    Winsock1.GetData s, vbString, bytesTotal
    
    ' take apart the packet...
    sBuffer = sBuffer & s
    
    ' now we use S for a temp string
    While InStr(" " & sBuffer, vbCrLf) > 0
        s = Left(sBuffer, InStr(sBuffer, vbCrLf) - 1)
        If s <> "" Then
            RaiseEvent PacketReceived(s)
        End If
        sBuffer = Mid(sBuffer, InStr(sBuffer, vbCrLf) + 2)
    Wend
End Sub


