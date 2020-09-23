VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00007000&
   Caption         =   "VB Racko"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000C0&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBackHighlight 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   240
      Picture         =   "frmMain.frx":0CCA
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.PictureBox picFrontHighlight 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   3240
      Picture         =   "frmMain.frx":120AC
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   240
      Picture         =   "frmMain.frx":2348E
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.PictureBox picFront 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   3240
      Picture         =   "frmMain.frx":34870
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Shape shpIndicate 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   1575
      Shape           =   3  'Circle
      Top             =   150
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00007000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   13920
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00007000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   11040
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00007000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   8040
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00007000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuNew 
         Caption         =   "New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCPU 
         Caption         =   "CPU Players"
         Begin VB.Menu mnuPlayers 
            Caption         =   "1"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuPlayers 
            Caption         =   "2"
            Index           =   1
         End
         Begin VB.Menu mnuPlayers 
            Caption         =   "3"
            Index           =   2
         End
      End
      Begin VB.Menu mnuAni 
         Caption         =   "Animation Speed"
         Begin VB.Menu mnuSpeed 
            Caption         =   "Fast"
            Index           =   0
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "Normal"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "Slow"
            Index           =   2
         End
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuGuide 
         Caption         =   "User Guide"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Racko
'By Paul Bahlawan
'Jan 2011
'
'----------------
'frmMain
'Main game components

Option Explicit
Option Base 1

Dim Deck(60) As Long
Dim Rack(4, 10) As Long
Dim Discard(20) As Long
Dim DeckPointer As Long
Dim DiscardPointer As Long
Dim Score(4) As Long
Dim numPlayers As Long
Dim GameMode As Long '0=Game Over , 1=Round Over , 2=Player Selecting , 3=Player Swaping , 4=cpu playing
Dim Selected As Long
Dim MaxCards As Long
Dim AniSpeed As Single

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub Form_Load()
    Randomize
    numPlayers = 2
    GameMode = 0
    AniSpeed = 1
End Sub

Private Sub NewGame()
Dim i As Long
    frmMain.Width = 3500 + 200 * numPlayers * 15
    frmMain.Cls
    Caret 0
    
    MaxCards = 20 + numPlayers * 10 '2-player=40, 3-player=50, 4-player=60

    'Set scores to 0
    For i = 1 To 4
        Score(i) = 0
        lblScore(i - 1).Caption = ""
    Next i

    NewRound
End Sub


'Set up for a new round
Private Sub NewRound()
Dim i As Long
Dim j As Long
Dim temp As Long

    frmMain.Cls
    Caret 0
    
    'Show scores
    For i = 1 To numPlayers
        lblScore(i - 1).Caption = Score(i)
    Next i
    
    'Show rack values
    frmMain.FontSize = 8
    frmMain.ForeColor = vbYellow
    For i = 10 To 1 Step -1
        frmMain.CurrentX = 0
        frmMain.CurrentY = 15 + (11 - i) * 30
        frmMain.Print i * 5
    Next i
    frmMain.Refresh
    frmMain.FontSize = 18
    frmMain.ForeColor = vbRed
    
    sndPlaySound App.Path & "\shufflle_cards.wav", 1
    
    'Fresh deck
    For i = 1 To MaxCards
        Deck(i) = i
    Next i
    
    'Shuffel the deck
    For i = 1 To MaxCards
        j = Int(1 + Rnd * MaxCards)
        temp = Deck(i)
        Deck(i) = Deck(j)
        Deck(j) = temp
    Next i
    DeckPointer = 1
    
    'Show deck
    DrawDeck cBack
    frmMain.Refresh
    
    'Deal the cards
    For i = 1 To 10
        For j = 1 To numPlayers
            Rack(j, i) = Deck(DeckPointer)
            incDeckPointer
            If j = 1 Then 'Player card (reveal)
                DrawCard j, i, cFront, Rack(j, i)
            Else 'Computer card (don't reveal)
                DrawCard j, i, cBack, Rack(j, i)
            End If
            frmMain.Refresh
            Sleep (45 - 8 * numPlayers) * AniSpeed
        Next j
    Next i
        
    'One card to the Discard pile
    DiscardPointer = 1
    Discard(DiscardPointer) = Deck(DeckPointer)
    incDeckPointer
    DrawDiscard cFront, Discard(DiscardPointer)
    frmMain.Refresh
    
    Caret 1
    GameMode = 2 'Game on!
End Sub

'User plays here
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim clicked As Long
Dim temp As Long
    
    If GameMode = 1 Then
        NewRound
        Exit Sub
    End If
    
    If GameMode < 2 Or GameMode > 3 Then Exit Sub
    
    clicked = PlayerClicked(x, y)
    
    If GameMode = 2 Then
        If clicked = 11 Then 'User selects the deck
            DrawDeck cFrontHighlight, Deck(DeckPointer)
            Selected = 11
            GameMode = 3
        ElseIf clicked = 12 Then 'User selects the discard
            DrawDiscard cFrontHighlight, Discard(DiscardPointer)
            Selected = 12
            GameMode = 3
        End If
        Exit Sub
    End If
    
    If GameMode = 3 Then
        If clicked >= 1 And clicked <= 10 Then 'users rack
        
            DrawRack 1, clicked, Rack()
            frmMain.Refresh
            Sleep 300 * AniSpeed
            
            If Selected = 11 Then 'Swap DECK to rack and rack to discard
                DiscardPointer = DiscardPointer + 1
                Discard(DiscardPointer) = Rack(1, clicked)
                Rack(1, clicked) = Deck(DeckPointer)
                incDeckPointer
            Else 'Swap DISCARD to rack and rack to discard
                temp = Rack(1, clicked)
                Rack(1, clicked) = Discard(DiscardPointer)
                Discard(DiscardPointer) = temp
            End If
            
            
            DrawRack 1, 0, Rack()
            DrawDiscard cFront, Discard(DiscardPointer)
            DrawDeck cBack
            Selected = 0
            GameMode = 4
        End If
        
        If clicked = 12 And Selected = 11 Then 'sends deck card to discard pile
            DiscardPointer = DiscardPointer + 1
            Discard(DiscardPointer) = Deck(DeckPointer)
            incDeckPointer
            DrawDiscard cFront, Discard(DiscardPointer)
            DrawDeck cBack
            Selected = 0
            GameMode = 4
        End If
    End If
    
    frmMain.Refresh
    Sleep 300 * AniSpeed
    
    If isRacko(1) = 10 Then
        RoundEnd 1
        Exit Sub
    End If
    
    If GameMode = 4 Then
        cpuPlays
    End If
    
End Sub

'Determin what the player has clicked on
' 0 = nothing
' 1-10 = rack
' 11 = deck
' 12 = discard
Private Function PlayerClicked(x As Single, y As Single) As Long
    If x > 10 And x < 205 And y > 30 And y < 419 Then 'Clicked the rack
        PlayerClicked = (y - 15) / 30
        If PlayerClicked > 10 Then
            PlayerClicked = 10
        End If
        Exit Function
    End If
    
    If x > 210 And x < 405 Then 'Clicked the deck
        If y > 60 And y < 179 Then
            PlayerClicked = 11
            Exit Function
        End If
        
        If y > 200 And y < 319 Then ' Clicked the discard
            PlayerClicked = 12
        End If
    End If
End Function

'Increase DeckPointer and, if Deck is finished, move cards from Discard >>to>> Deck and reshuffle
Private Sub incDeckPointer()
Dim i As Long
Dim j As Long
Dim temp As Long
Dim limit As Long

    DeckPointer = DeckPointer + 1
    
    If DeckPointer > MaxCards Then 'is the deck finished?
    
        DrawDeck cNone
        frmMain.Refresh
        
        sndPlaySound App.Path & "\shufflle_cards.wav", 2
       
        'Move all but the top discard card to the deck
        limit = DiscardPointer - 1
        
        For i = 1 To limit
            Deck(1 + MaxCards - i) = Discard(i)
        Next i
        Discard(1) = Discard(DiscardPointer)
        DiscardPointer = 1
        
        DeckPointer = 1 + MaxCards - limit
        
        'Shuffle the deck
        For i = DeckPointer To MaxCards
            j = Int(DeckPointer + Rnd * limit)
            temp = Deck(i)
            Deck(i) = Deck(j)
            Deck(j) = temp
        Next i
        
        DrawDeck cBack
    End If
End Sub

'Computer opponent(s) play here
Private Sub cpuPlays()
Dim cpuNum As Long
Dim rPos As Long
Dim temp As Long

    For cpuNum = 2 To numPlayers
        
        Caret cpuNum
        Sleep 200 * AniSpeed
        
        '//////////// CPU AI should be here (very simple for now)...
        
        Selected = 12 'Check out the discard pile first
        rPos = 10 - Int(Discard(DiscardPointer) / (MaxCards + 1) * 10)
        If 10 - Int(Rack(cpuNum, rPos) / (MaxCards + 1) * 10) = rPos Then
            Selected = 11 'Nothing good in the discard so take a card from the deck
            rPos = 10 - Int(Deck(DeckPointer) / (MaxCards + 1) * 10)
        End If
        
        '//////////// End of CPU AI
        
        DrawRack cpuNum, rPos, Rack()
        
        If Selected = 11 Then 'Swap DECK to rack and rack to discard
            DrawDeck cBackHighlight
            DiscardPointer = DiscardPointer + 1
            Discard(DiscardPointer) = Rack(cpuNum, rPos)
            Rack(cpuNum, rPos) = Deck(DeckPointer)
            incDeckPointer
        Else 'Swap DISCARD to rack and rack to discard
            DrawDiscard cFrontHighlight, Discard(DiscardPointer)
            temp = Rack(cpuNum, rPos)
            Rack(cpuNum, rPos) = Discard(DiscardPointer)
            Discard(DiscardPointer) = temp
        End If

        frmMain.Refresh
        Sleep 500 * AniSpeed
        
        DrawRack cpuNum, 0, Rack()
        DrawDeck cBack
        DrawDiscard cFront, Discard(DiscardPointer)
        frmMain.Refresh
        DoEvents
        
        If isRacko(cpuNum) = 10 Then
            RoundEnd cpuNum
            Exit Sub
        End If
    Next cpuNum
    
    Caret 1
    GameMode = 2
End Sub

'Is Racko?
' 1-9 = number of consecutive cards
' 10 = Racko!
Private Function isRacko(player As Long) As Long
Dim i As Long
    
    isRacko = 1
    For i = 9 To 1 Step -1
        If Rack(player, i) > Rack(player, i + 1) Then
            isRacko = 11 - i
        Else
            Exit For
        End If
    Next i
End Function

Private Sub RoundEnd(player As Long)
Dim i As Long
Dim j As Long
Dim k As Long
Dim points As Long

    GameMode = 1
    
    For i = 1 To numPlayers
        
        'Tally up the points
        k = isRacko(i)
        points = k * 5
        If player = i Then
            points = points + 25
        End If
        lblScore(i - 1).Caption = Score(i) & " +" & points
        Score(i) = Score(i) + points
        
        'Show everyone's cards
        For j = 1 To 10
            If k >= 11 - j Then
                DrawCard i, j, cFrontHighlight, Rack(i, j)
            Else
                DrawCard i, j, cFront, Rack(i, j)
            End If
        Next j
    Next i
    
    DrawRacko
    frmMain.Refresh
    sndPlaySound App.Path & "\proceed.wav", 2

    'Has someone won the game (500 points or more)?
    j = 1
    For i = 2 To numPlayers
        If Score(i) > Score(j) Then
            j = i
        End If
    Next i
    
    If Score(j) >= 500 Then
        Caret j
        
        'Show new scores
        For i = 1 To numPlayers
            lblScore(i - 1).Caption = Score(i)
        Next i

        'Game over
        sndPlaySound App.Path & "\party_horn.wav", 1
        MsgBox " PLAYER " & j & " WINS THE GAME! "
        GameMode = 0
    End If
End Sub

Private Sub Caret(player As Long)
Dim x As Long
    If player = 0 Then
        shpIndicate.Visible = False
    Else
        shpIndicate.Visible = True
    End If
    
    If player = 1 Then
        x = 105
    Else
        x = 105 + 200 * player
    End If

    shpIndicate.Left = x
End Sub




'Menu Stuff.................................................................
Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuGuide_Click()
    ShellExecute Me.hWnd, vbNullString, App.Path & "\Racko.chm", vbNullString, "c:\", 1
End Sub

Private Sub mnuNew_Click()
    If GameMode <> 0 Then
        If MsgBox("This action will end the current game." & Chr(13) & "Do you want to end current game?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
    
    GameMode = 0
    NewGame
End Sub

Private Sub mnuPlayers_Click(Index As Integer)
    If mnuPlayers(Index).Checked Then
        Exit Sub
    End If
    
    If GameMode <> 0 Then
        If MsgBox("This action will end the current game." & Chr(13) & "Do you want to apply the changes and end current game?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
    
    GameMode = 0
    frmMain.Cls
    
    mnuPlayers(0).Checked = False
    mnuPlayers(1).Checked = False
    mnuPlayers(2).Checked = False
    
    numPlayers = Index + 2
    mnuPlayers(Index).Checked = True
    
    NewGame
End Sub

Private Sub mnuSpeed_Click(Index As Integer)
    mnuSpeed(0).Checked = False
    mnuSpeed(1).Checked = False
    mnuSpeed(2).Checked = False
    mnuSpeed(Index).Checked = True
    Select Case Index
        Case 0
            AniSpeed = 0.33
        Case 1
            AniSpeed = 1
        Case 2
            AniSpeed = 2
    End Select
End Sub
