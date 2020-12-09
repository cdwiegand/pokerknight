VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "PokerServer"
   ClientHeight    =   2100
   ClientLeft      =   5250
   ClientTop       =   2310
   ClientWidth     =   5985
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu menuMain 
      Caption         =   "Commands"
      Begin VB.Menu mnuMainRD 
         Caption         =   "Reshuffle Deck"
      End
      Begin VB.Menu mnuMainExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menuViews 
      Caption         =   "Menus"
      Begin VB.Menu bringmain 
         Caption         =   "Main Menu"
      End
      Begin VB.Menu mnuViewDeck 
         Caption         =   "View Deck"
      End
      Begin VB.Menu UserList 
         Caption         =   "User List"
      End
   End
   Begin VB.Menu syevents 
      Caption         =   "Records"
      Begin VB.Menu mnuLogsView 
         Caption         =   "Game Logs"
      End
      Begin VB.Menu mnuReports 
         Caption         =   "Report Screen"
      End
      Begin VB.Menu mnuViewsDebug 
         Caption         =   "Debug Log"
      End
   End
   Begin VB.Menu menuUsers 
      Caption         =   "Users"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MAIN PROJECT FORM
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
' Form information: Main Project Form
' Purpose: This is the main interface and input form of the server.
' Important Notes: The deck functions are also executed here.

' Revision log:
' 11/11/02 - Put all these notes in!


' =-=-=-=-=-=-=-=-=
' |  Declarations |
' =-=-=-=-=-=-=-=-=
' Option Explicit
Private Deck(52) As String ' FIXME review me


' Get Card Function
' =================
' Written by: Chris
' Function: This function will return the card at that address.
'
Public Function GetCard(ByVal iCard As Integer) As String
    GetCard = Deck(iCard)
End Function


' Setup Deck Function
' ===================
' Written by: Chris
' Function: This function will generate the deck.
'
Public Sub SetupDeck() ' FIXME review me
' FIXME Duplicate card: Ten of Hearts generated twice within one session
    Dim iSuit As Integer
    Dim iCard As Integer
    Dim sTmp As String
    Dim lDeckBase As Long
    ' generate the deck
    For iSuit = 1 To 4
        For iCard = 1 To 13
            lDeckBase = ((iSuit - 1) * 13)
            Select Case iSuit
                Case 1
                    sTmp = "S" ' spades
                Case 2
                    sTmp = "C" ' clubs
                Case 3
                    sTmp = "H" ' hearts
                Case 4
                    sTmp = "D" ' diamonds
            End Select
            Select Case iCard
                Case 1 ' Ace
                    sTmp = sTmp & "A"
                Case 2, 3, 4, 5, 6, 7, 8, 9
                    sTmp = sTmp & iCard
                Case 10 ' 10 / 0
                    sTmp = sTmp & "0"
                Case 11 ' Jack / J
                    sTmp = sTmp & "J"
                Case 12 ' Queen / Q
                    sTmp = sTmp & "Q"
                Case 13 ' King / K
                    sTmp = sTmp & "K"
            End Select
            sTmp = sTmp & "0" ' off
            Deck(lDeckBase + iCard) = sTmp
        Next iCard
    Next iSuit
    refreshDeckList
End Sub


' Refresh Deck Function
' =====================
' Written by: Chris
' Function: This function will refresb the deck list appearance.
'
Public Function refreshDeckList()
    ' updates DeckView...
DeckView.RefreshView
End Function


' Get Unique ID Function
' ======================
' Written by: Chris
' Function: This function assigns a unique ID to the player.
'
Public Function GetUniqueID() As Long
    Static lUID As Long
    lUID = lUID + 1
    GetUniqueID = lUID
End Function

' Number Of Cards Function
' ======================
' Written by: Chris
' Function: This function counts the # of cards a UID has
Public Function NumberOfCards(ByVal lUID As Long) As Integer
    Dim i As Integer
    Dim sCard As String
    Dim iNumCards As Integer
    For i = 1 To 52
        sCard = Deck(iCard)
        If Len(sCard) >= 4 Then
            If Trim(Mid(sCard, 4)) = Trim(CStr(lUID)) Then iNumCards = iNumCards + 1
        End If
    Next i
    NumberOfCards = iNumCards
End Function

' Get Free Card Function
' ======================
' Written by: Chris
' Function: This function generates a card and assigns it, if it is not already dealt.
'
Public Function GetFreeCard(ByVal lUID As Long) As String '  review me
    Dim sCard As String
    Dim i As Integer
    Dim iCard As Integer
    Dim iSafety As Integer
    Dim suitmark As String
    
    i = DatePart("n", Now()) * DatePart("s", Now())
    Randomize i
    
    While sCard = ""
        iSafety = iSafety + 1
        If iSafety > 5000 Then
            MsgBox "Safety reached - deck full - please reshuffle"
            Exit Function
        End If
        ' card
        iCard = Rnd(52) * 52
        
        ' deck(??) = ABCD
        ' A = deck
        ' B = card #
        ' C = 1/0 in use?
        ' D = 1-5 digit lID of in use user
        
        sCard = Deck(iCard)
        If Mid(sCard, 3, 1) = "1" Then
            sCard = "" ' BAD CARD ALREADY IN USE
        Else
            ' sCard WAS S50, now S5 (spades, 5)
            sCard = Left(sCard, 2) ' don't need to tell them it's 0...
            Deck(iCard) = sCard & "1" & lUID ' tell it it's in use...
        End If
    Wend
    suitmark = "Unknown Suit"
    If Mid(sCard, 1, 1) = "H" Then suitmark = "Hearts"
    If Mid(sCard, 1, 1) = "S" Then suitmark = "Spades"
    If Mid(sCard, 1, 1) = "C" Then suitmark = "Clubs"
    If Mid(sCard, 1, 1) = "D" Then suitmark = "Diamonds"
    writeup "---( Card dealt to " & sNick & ": the " & Mid(sCard, 2) & " of " & suitmark & " )---"
    refreshDeckList
    GetFreeCard = sCard
End Function


' Free Card Function
' ==================
' Written by: Chris
' Function: This function marks a card as unused.
'
Public Sub FreeCard(ByVal sCardName As String) '  review me
    Dim i As Integer
    
    For i = 1 To 52
        If Left(Deck(i), 2) = sCardName Then
            Deck(i) = sCardName & "-"
        End If
    Next i
    
    refreshDeckList
End Sub


' MDI Form Load Function
' ======================
' Written by: Chris
' Function: Preloads the menus and sets up the MDI Form.
'
Private Sub MDIForm_Load()
    Dim t As New frmServConn
    SetupDeck
    t.Show
    mainmen.Show
    frmDebug.Hide ' preloads the form
    DeckView.Hide ' preloads the form
    frmuserlist.Hide
    logs.Hide
    reportscreen.Hide
End Sub


' Report Add Function
' ===================
' Written by: Chris
' Function: Writes to the report screen and logs the message to the file.
'
Private Sub ReportAdd(ByVal sString As String)
    reportscreen.List1.AddItem sString, 0
    writeup sString
End Sub


' MultiCast Function
' ==================
' Written by: Chris
' Function: ??
'
Public Sub SendMulticastPacket(ByVal sPacket As String)
    Dim f As frmServConn
    Dim i As Integer
    On Error Resume Next
    
    frmDebug.AddItem "OUT:" & sPacket, 0
    For i = 0 To Forms.Count - 1
        If Forms(i).Name = "frmServConn" Then
            Set f = Forms(i)
            f.SendPacket sPacket
        End If
    Next i
End Sub


' MultiCast Function
' ==================
' Written by: Chris
' Function: Sends the packet to everyone connected.
'
Public Sub PacketReceived(ByVal sPacket As String, ByVal lUID As Long, ByRef theServConn As frmServConn, ByVal sNick As String)
    ' now we'll send it out to everyone
    frmDebug.AddItem sNick & ": " & sPacket, 0
    Select Case Left(sPacket, 2)
    Case "T:"
        ' public chat
        SendMulticastPacket "T:" & sNick & ":" & Mid(sPacket, 3)
        ReportAdd sNick & ": " & Mid(sPacket, 3)
                
    Case "C:"
        If Left(sPacket, 7) = "C:NICK=" Then
            theServConn.UpdateCaption Mid(sPacket, 8)
            SendMulticastPacket "T:" & "** " & Mid(sPacket, 8) & " joined the game. **"
            PlayerListSync
            frmuserlist.Show ' SHOW the userlist now that they've "signed in"
            ReportAdd "=+= New Player Joined the Game:" & Mid(sPacket, 8) & "=+="
        ElseIf Left(sPacket, 7) = "C:GUID?" Then
            theServConn.SendPacket "C:GUID=" & lUID
        ElseIf Left(sPacket, 10) = "C:DISCARD=" Then
            ' this user discarded a card... will request one next packet
            FreeCard Mid(sPacket, 11, 2)
        ElseIf sPacket = "C:FOLD" Then
            SendMulticastPacket "T:** " & sNick & " has FOLDED! **"
            ReportAdd "=+= " & sNick & ": has FOLDED. =+="
        ElseIf sPacket = "C:GETCARD" Then
            ' get card...
            If NumberOfCards(lUID) >= 5 Then
                theServConn.SendPacket "E:ERROR_TOO_MANY_CARDS"
            Else
                theServConn.SendPacket "C:CARD=" & GetFreeCard(lUID)
                DeckView.RefreshView
            End If
        ElseIf Left(sPacket, 9) = "C:AVATAR=" Then
            ' C:AVATAR=3 or C:AVATAR=Coolies
            ' set avatar
            theServConn.sAvatarID = Mid(sPacket, 10)
            Me.SendMulticastPacket "C:AVASET:" & theServConn.sAvatarID & "=" & sNick
        End If
    Case "P:" ' secret packet
        ' do NOTHING except display in list...
        frmDebug.AddItem sNick & ":(Private!):" & Mid(sPacket, 3), 0
        ReportAdd "<< Private Message from " & sNick & ": " & Mid(sPacket, 3) & " >>"
        ' now send back to THAT client only
        theServConn.SendPacket "T:<TSO>: " & Mid(sPacket, 3) & " <>"
        writeup "<[Server Only from " & sNick & "]>: " & Mid(sPacket, 3)
    End Select
End Sub

' MDI Unload Function
' ===================
' Written by: Chris
' Function: Unloads MDI so that the app can exit.
'
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub


' Exit Function
' =============
' Written by: Chris
' Function: Quits the application.
'
Private Sub mnuMainExit_Click()
    If MsgBox(" Do you want to exit?", vbYesNo + vbDefaultButton2) = vbYes Then End
End Sub

Private Sub mnuLogsView_Click()
logs.Hide
reportscreen.Hide
DeckView.Hide
logs.Show
End Sub

Private Sub mnuDeckView_Click()
    reportscreen.Hide
    DeckView.Show
End Sub

Private Sub mnuMainRD_Click()
    If MsgBox("Reshuffle Deck?", vbYesNo + vbDefaultButton2) = vbYes Then SetupDeck
End Sub

Private Sub mnuReports_Click()
    reportscreen.Hide
    reportscreen.Show
End Sub

Public Sub PlayerDropped(ByVal sName As String, ByVal sMode As String)
    If sMode = "DROPPED" Then
        SendMulticastPacket "T:** " & sName & " has been dropped. **"
        ReportAdd "-*- " & sName & ": has been dropped. -*-"
    Else ' LOST
        SendMulticastPacket "T:** " & sName & " dropped connection. **"
        ReportAdd "=+= " & sName & " dropped connection. =+="
    End If
    ' SendMulticastPacket "C:DROP=" & sName
    SendMulticastPacket "C:PlaySound"
    PlayerListSync

End Sub

Public Sub PlayerAdded(ByVal theServConn As frmServConn)
    If frmuserlist.lbUsers.ListCount >= 5 Then
    ' Send the denial
        theServConn.SendPacket "T:** DENIED! Too many players."
        writeup "=+= User was denied, the game was full. =+="
        theServConn.SendPacket "T:** Try again later.... **"
        theServConn.DropUser
    Else
        theServConn.SendPacket "C:NICK?"
        theServConn.SendPacket "T:** Welcome to PokerServer revision " & App.Revision & " **"
    End If
End Sub

Public Function GetPlayerForm(ByVal sName As String) As frmServConn
    Dim f As frmServConn
    Dim i As Integer
    On Error Resume Next

    For i = 0 To Forms.Count - 1
        If Forms(i).Name = "frmServConn" Then
            Set f = Forms(i)
            
            If f.sName = sName Then
                Set GetPlayerForm = f
                Exit Function
            End If
        End If
    Next i
End Function

Public Sub PlayerListSync()
    Dim f As frmServConn
    Dim i As Integer
    On Error Resume Next
    
    SendMulticastPacket "C:CLSP"
    frmuserlist.lbUsers.Clear
    
    For i = 0 To Forms.Count - 1
        If Forms(i).Name = "frmServConn" Then
            Set f = Forms(i)
            
            If f.sName <> "" Then
            SendMulticastPacket "C:NEWP=" & f.sName
            frmuserlist.lbUsers.AddItem f.sName & " @ " & f.sIPAddress
            End If
        
        End If
    Next i
    frmuserlist.Hide
    frmuserlist.Show
End Sub

Private Sub mnuViewDeck_Click()
    DeckView.Hide
    DeckView.Show
End Sub

Private Sub bringmain_Click()
    mainmen.Hide
    mainmen.Show
End Sub


Private Sub mnuViewsDebug_Click()
    frmDebug.Hide
    frmDebug.Show
End Sub

Private Sub UserList_Click()
    frmuserlist.Hide
    frmuserlist.Show
End Sub

'Private Function EvalHand(ByVal hand As Collection) As Long
'    Dim l As Long
'
'    EvalHand = CheckRoyalFlush(hand)
'    If EvalHand > 0 Then Exit Function ' hand must have hit
'
'    EvalHand = CheckStraightFlush(hand)
'    If EvalHand > 0 Then Exit Function ' hand must have hit
'
'    EvalHand = CheckStraight(hand)
'    If EvalHand > 0 Then Exit Function ' hand must have hit
'
'    EvalHand = CheckFlush(hand)
'    If EvalHand > 0 Then Exit Function ' hand must have hit
'
'    EvalHand = CheckFullHouse(hand)
'    If EvalHand > 0 Then Exit Function ' hand must have hit
'
'    EvalHand = CheckThreeOfAKind(hand)
'    If EvalHand > 0 Then Exit Function ' hand must have hit
'
'    EvalHand = CheckTwoPair(hand)
'    If EvalHand > 0 Then Exit Sub ' hand must have hit
'
'    EvalHand = CheckPair(hand)
'    If EvalHand > 0 Then Exit Sub ' hand must have hit
'
'    EvalHand = CheckHigh(hand)
'    If EvalHand > 0 Then Exit Sub ' hand must have hit
'End Sub

Private Function CheckFlush(ByVal hand As Collection) As Integer
    Dim i As Integer
    Dim j As Integer
    Dim bFlush As Boolean
    For i = 1 To 4
        bFlush = True
        For j = 1 To 5
            If ConvertCardSuitToInt(hand(j)) <> i Then bFlush = False
        Next j
        If bFlush = True Then
            ' made it
            CheckFlush = 500 + CheckHigh(hand) ' get highest card
        End If
    Next i
End Function

Private Function CheckHigh(ByVal hand As Collection) As Integer
    ' run through, find highest card
    Dim i As Integer
    Dim i2 As Integer
    
    CheckHigh = ConvertCardToInt(hand.Item(1))
    For i = 2 To hand.Count
        i2 = ConvertCardToInt(hand.Item(i))
        If i2 > CheckHigh Then CheckHigh = i2
    Next i
End Function

Public Function CheckFullHouse(ByVal hand As Collection)
    Dim iPair As Integer
    Dim i3 As Integer
    iPair = CheckPair(hand)
    i3 = CheckThreeOfAKind(hand)
    If iPair <> i3 And iPair <> 0 And i3 <> 0 Then
        CheckFullHouse = 400 + iPair
    End If
End Function

Public Function CheckPair(ByVal hand As Collection)
    Dim iCard(1 To 5) As Integer
    Dim i As Integer
    Dim j As Integer
    Dim iCount As Integer
    For i = 1 To 5
        iCard(i) = ConvertCardToInt(hand.Item(i))
    Next i
    For i = 2 To 14
        iCount = 0
        For j = 1 To 5
            If iCard(j) = i Then iCount = iCount + 1
        Next j
        If iCount = 2 Then CheckPair = 100 + i ' card number
    Next i
End Function

Public Function CheckThreeOfAKind(ByVal hand As Collection)
    Dim iCard(1 To 5) As Integer
    Dim i As Integer
    Dim j As Integer
    Dim iCount As Integer
    For i = 1 To 5
        iCard(i) = ConvertCardToInt(hand.Item(i))
    Next i
    For i = 2 To 14
        iCount = 0
        For j = 1 To 5
            If iCard(j) = i Then iCount = iCount + 1
        Next j
        ' If iCount = 3 Then CheckPair = 300 + i ' card number
    Next i
End Function

Public Function CheckTwoPair(ByVal hand As Collection)
    Dim iCard(1 To 5) As Integer
    Dim iPairs As Integer
    Dim iHighestPair As Integer
    Dim i As Integer
    Dim j As Integer
    Dim iCount As Integer
    For i = 1 To 5
        iCard(i) = ConvertCardToInt(hand.Item(i))
    Next i
    For i = 2 To 14
        iCount = 0
        For j = 1 To 5
            If iCard(j) = i Then iCount = iCount + 1
        Next j
        If iCount = 2 Then
            iPairs = iPairs + 1
            iHighestPair = i ' highest "card"
        End If
    Next i
    If iPairs = 2 Then
        CheckTwoPair = 200 + iHighestPair
    End If
End Function

Public Function ConvertCardSuitToInt(ByVal sCard As String) As Integer
    ' suit card use
    Dim s As String
    s = Mid(sCard, 1, 1)
    Select Case LCase(s)
        Case "s"
            ConvertCardSuitToInt = 1
        Case "c"
            ConvertCardSuitToInt = 2
        Case "h"
            ConvertCardSuitToInt = 3
        Case "d"
            ConvertCardSuitToInt = 4
    End Select
End Function

Public Function ConvertCardToInt(ByVal sCard As String) As Integer
    ' suit card use
    Dim s As String
    s = Mid(sCard, 2, 1)
    Select Case LCase(s)
        Case "a"
            ConvertCardToInt = 14
        Case "k"
            ConvertCardToInt = 13
        Case "q"
            ConvertCardToInt = 12
        Case "j"
            ConvertCardToInt = 11
        Case 0
            ConvertCardToInt = 10
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9
            ConvertCardToInt = CInt(s)
    End Select
End Function

Private Sub CheckPairs(ByVal hand As Collection)
Dim record(5) As Integer
Dim i As Integer
Dim i2 As Integer

For i = 1 To 5
For i2 = 1 To 5
If hand(i) = hand(i2) Then record(i) = (record(i) + 1)
Next i2
Next i

' Each card will have a tally of how many other cards in the hand
' match it. So if the hand is 2H, 4D, 9S, 6C, 2C, then the array
' will end up reading [2, 0, 0, 0, 2].
' So a 2 in any field of the match array denotes one match.
' A 3 in any of the match array denotes three of a kind.
' More then 1 instances of 2 in the array denotes 3 or 4 of a kind.
' Thus, all like card matches can be returned.
' Return of 1=A pair, 2=2 Pairs, 3=3 of a Kind, 4=4 of a Kind.

' PAIR SCORING
'   Pair - [20] + Pair Val
'   Two Pair - [50] + Pair Vals
'   Three of a Kind - [100] + Total
'   Full House - [200] + total

For i = 1 To 5
If record(i) > 0 Then
match = 1
Else
Exit Sub
End If
Next i

End Sub

'Private Sub CheckFlush(ByVal hand As Collection)
'Dim i As Integer
'Dim i2 As Integer
'Dim flush As Integer
'
'' Ok, we'll actually have to modify this to check the second
'' item in the collection array (the suit).
'' Simple: loop inside a loop. It'll take each card and compare
'' it to every card in the loop, including itself. If all suits
'' aren't identical, no flush.
'
'' Flush Scoring
''   Flush - [300] + High
'
'For i = 1 To 5
'For i2 = 1 To 5
'If hand(i) = hand(i2) Then
'' the cards suits are equal.
'' Break out when we've run thru all and check for a straight flush.
'Else
'' if it gets here, at least one suit is different.
'Exit Sub
'End If
'End Sub
'
'Private Sub CheckStraight(ByVal hand As Collection)
'Dim i As Integer
'
'' Chris, we're going to have to do this one after we've sorted
'' each hand.
'' Once we've done that, we'll just check that each hand increments
'' by one, then we will check to see if all the suits match (straight
'' flush). If those two tests pass, we'll check for 10,J,Q,K,A for a
'' Royal Flush.
'
'' Straight Scoring
''   Straight - [400] + Total
''   Straight Flush - [500] + Total
''   Royal Flush - [800]
'
'
'End Sub
'
'
'
''   High - H Card
'





