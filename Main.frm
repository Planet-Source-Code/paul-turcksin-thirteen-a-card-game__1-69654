VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "13's"
   ClientHeight    =   7845
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   523
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   658
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrRemoveCard 
      Enabled         =   0   'False
      Interval        =   12
      Left            =   1440
      Top             =   120
   End
   Begin VB.Timer tmrGameOver 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   840
      Top             =   120
   End
   Begin VB.Frame frHelp 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Help"
      Height          =   3735
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Label lblHelp 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Timer tmrInvalidSelection 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblGameOver 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   975
      Index           =   7
      Left            =   6960
      TabIndex        =   10
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblGameOver 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   975
      Index           =   6
      Left            =   6240
      TabIndex        =   9
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblGameOver 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   975
      Index           =   5
      Left            =   5520
      TabIndex        =   8
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblGameOver 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Index           =   4
      Left            =   4800
      TabIndex        =   7
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblGameOver 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   975
      Index           =   3
      Left            =   3720
      TabIndex        =   6
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblGameOver 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   3000
      TabIndex        =   5
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblGameOver 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   975
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblGameOver 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblInvalidSelection 
      BackStyle       =   0  'Transparent
      Caption         =   "Invalid selection"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Menu mnuNewGame 
      Caption         =   "New Game"
   End
   Begin VB.Menu mnuReplay 
      Caption         =   "Replay"
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuSounds 
         Caption         =   "Sounds"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAnimateHeader 
         Caption         =   "Animate"
         Begin VB.Menu mnuAnimate 
            Caption         =   "No"
            Index           =   0
         End
         Begin VB.Menu mnuAnimate 
            Caption         =   "Slow"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuAnimate 
            Caption         =   "Fast"
            Index           =   2
         End
      End
      Begin VB.Menu mnuBacksHeader 
         Caption         =   "Backs"
         Begin VB.Menu mnuBacks 
            Caption         =   "Plaid"
            Index           =   0
            Begin VB.Menu mnuPlaidColor 
               Caption         =   "Red"
               Index           =   0
            End
            Begin VB.Menu mnuPlaidColor 
               Caption         =   "Blue"
               Index           =   1
            End
            Begin VB.Menu mnuPlaidColor 
               Caption         =   "Cyan"
               Index           =   2
            End
            Begin VB.Menu mnuPlaidColor 
               Caption         =   "Yellow"
               Index           =   3
            End
            Begin VB.Menu mnuPlaidColor 
               Caption         =   "Magenta"
               Index           =   4
            End
            Begin VB.Menu mnuPlaidColor 
               Caption         =   "White"
               Index           =   5
            End
         End
         Begin VB.Menu mnuBacks 
            Caption         =   "Sky"
            Index           =   1
         End
         Begin VB.Menu mnuBacks 
            Caption         =   "Blues"
            Index           =   2
         End
         Begin VB.Menu mnuBacks 
            Caption         =   "Fish"
            Index           =   3
         End
         Begin VB.Menu mnuBacks 
            Caption         =   "Frog"
            Index           =   4
         End
         Begin VB.Menu mnuBacks 
            Caption         =   "Wave"
            Index           =   5
         End
         Begin VB.Menu mnuBacks 
            Caption         =   "Island"
            Index           =   6
         End
         Begin VB.Menu mnuBacks 
            Caption         =   "Cross"
            Index           =   7
         End
         Begin VB.Menu mnuBacks 
            Caption         =   "Purple"
            Index           =   8
         End
         Begin VB.Menu mnuBacks 
            Caption         =   "Dune"
            Index           =   9
         End
         Begin VB.Menu mnuBacks 
            Caption         =   "Astronaut"
            Index           =   10
         End
         Begin VB.Menu mnuBacks 
            Caption         =   "Stripes"
            Index           =   11
         End
         Begin VB.Menu mnuBacks 
            Caption         =   "Cars"
            Index           =   12
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   13 - a card game
'
' I saw this game for the first time on PSC, a post by Sehab Veljacic.
' His implementation used multiple arrays of picture boxes, timers, an ImageList
' and in my opinion rather lengthy, repetitive code.
' My challenge: rewrite it from scratch Using the Cards.dll and regions.
'
' Please note hat sounds will only be heard when compiled.
'
' You can use or misuse this code as long as it is done for non-commercial purposes.
'
' Paul Turcksin, November 2007
'
'___________________________________________________________________________________

Option Explicit
'............................ OBJECTS
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'............................ BRUSH
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private hBrush As Long


'............................ REGIONS
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private arRgn(29) As Long
Private rectRemoveCard As RECT
Private iCntRemoveCard As Integer

'............................ CARDS
Private Declare Function cdtInit Lib "Cards.Dll" (Dx As Long, Dy As Long) As Long
Private Declare Function cdtDraw Lib "Cards.Dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal iCard As Long, ByVal iDraw As Long, ByVal clr As Long) As Long
Private Declare Function cdtTerm Lib "Cards.Dll" () As Long

'............................ SOUND
' Sounds are played directly from the resource file.
' !!! This feature ONLY works wwhen compiled.   !!!
Private Declare Function PlayResWAV Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszName&, ByVal dwFlags&) As Long
Private Const SND_RESOURCE = &H40004 'play from resource
Private Const SND_ASYNC = &H1        'play asynchronously or in other words return immediately after beginning the sound

' Card constants: How are we rendering the card?
Private Const ordFaces  As Long = 0          '
Private Const ordBacks  As Long = 1
Private Const ordInvert As Long = 2

' card size
Private lWidth As Long
Private lHeight As Long

' In the following types:
' - Number stands for the card number as defined in Cards.dll
' - Value stands for the card value: 1=ace, 2=2, ... King=13
Private Type PLAYBOARD
   Left As Long
   Top As Long
   Number As Integer
   Value As Integer
   Covered As Integer
   Discarded As Boolean
End Type
Private arPlayboard(29) As PLAYBOARD

Private Type DECK
   Number As Integer        ' card number as difined in Cards.dll
   Value As Integer         ' card alue: 1=ace, 2=2, ... King=13
   Discarded As Boolean
End Type
Private arDeck(23) As DECK
Private iCurrentDeck As Integer

Private Type SELECTED
   Index As Integer
   Number As Integer
   Value As Integer
End Type
Private FirstCard As SELECTED
Private SecondCard As SELECTED
Private iSelected As Integer

Private arCards(51) As Integer
Private arRow(6) As Integer   ' index array is row, value is first slot number in that row
Private iBack As Integer
Private swAnimate As Boolean
Private iAnimateSpeed As Integer
Private iOldAnimateIndex As Integer
Private swSound As Boolean
Private Const cReplay As Boolean = True

Private Sub Form_Load()
   Dim iRow As Integer
   Dim iCol As Integer
   Dim iCnt As Integer
   Dim lLeft As Long
   
' init cards DLL
   cdtInit lWidth, lHeight
   
' set card position on each of the seven rows and create regions
   For iRow = 0 To 6
      lLeft = 300 - (iRow * (lWidth \ 2))
      For iCol = 0 To iRow
         With arPlayboard(iCnt)
            .Left = lLeft + (iCol * lWidth)
            .Top = 75 + (iRow * 25)
            arRgn(iCnt) = CreateRectRgn(.Left, .Top, .Left + lWidth, .Top + lHeight)
            iCnt = iCnt + 1
         End With
      Next iCol
   Next iRow
   
' Finally we create two other regions:
' 1. for the covered deck: clicking it will show the next card from the deck in
' 2. uncovered deck allowing selection of the uncovered card
   arRgn(28) = CreateRectRgn(300, 400, 300 + lWidth, 400 + lHeight)
   arPlayboard(28).Left = 300
   arPlayboard(28).Top = 400
   arRgn(29) = CreateRectRgn(200, 400, 200 + lWidth, 400 + lHeight)
   
' Help
   lblHelp = "The objective of this game is to discard all cards shown on the board. " & _
           "This is done by clicking pairs of cards that total 13. An ace = 1, a 2 = 2, " & _
           "and so on. The valet = 11, queen = 12. As the king has value 13 clicking " & _
           "it will discard it. On the bottom of the board is a deck with the cards not " & _
           "shown on the board. Click on this deck to show the ""hidden"" cards. These " & _
           "can also be used to form pairs of 13." & vbCrLf & _
           "Cards hat are covered by another card cannot be selected. And if you " & _
           "mistakenly selected a card, click it again to de-select it." & vbCrLf & vbCrLf & _
           "Click on this message to hide it and enjoy the game!"
           
' init misc
   hBrush = CreateSolidBrush(Me.BackColor)
   iBack = 54
   arRow(1) = 1
   arRow(2) = 3
   arRow(3) = 6
   arRow(4) = 10
   arRow(5) = 15
   arRow(6) = 21
   Randomize   ' ensure random card shuffle
   swAnimate = True
   iAnimateSpeed = 1
   iOldAnimateIndex = 1
   subGameOverStop
   subNewGame
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim i As Integer
   
' clicked over a region?
   iSelected = -1
   For i = 29 To 0 Step -1
      If PtInRegion(arRgn(i), CLng(X), CLng(Y)) <> 0 _
      And Not arPlayboard(i).Discarded Then
         iSelected = i
         Exit For
      End If
   Next i
   
   Select Case iSelected
   
'     case -1" nothing to do

      Case 0 To 28  ' any uncovered  card with (28) being the deck
         If arPlayboard(iSelected).Covered = 0 Then
            ' is it a king
            If arPlayboard(iSelected).Value = 13 Then
               FirstCard.Index = iSelected
               subRemoveCard FirstCard
               Exit Sub
            End If
            subDrawCard iSelected, ordInvert
           ' first card? show and preserve info
            If FirstCard.Index = -1 Then
               FirstCard.Index = iSelected
               FirstCard.Number = arPlayboard(iSelected).Number
               FirstCard.Value = arPlayboard(iSelected).Value
            Else
               ' second card
               SecondCard.Index = iSelected
               SecondCard.Number = arPlayboard(iSelected).Number
               SecondCard.Value = arPlayboard(iSelected).Value
               subProcessSelection
            End If
         End If
      
      Case 29        ' the covered deck (show next card)
         Do
            iCurrentDeck = iCurrentDeck + 1
            If iCurrentDeck > 23 Then
               iCurrentDeck = -1
               FillRgn Me.hdc, arRgn(28), hBrush
               Me.Refresh
            Else
               If Not arDeck(iCurrentDeck).Discarded Then
                  arPlayboard(28).Number = arDeck(iCurrentDeck).Number
                  arPlayboard(28).Value = arDeck(iCurrentDeck).Value
                 subDrawCard 28, ordFaces
                 Exit Do
              End If
            End If
         Loop Until iCurrentDeck = -1
         
      End Select
   Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim i As Integer
   
' cleanup regions, brush
   For i = 0 To 29
      DeleteObject arRgn(i)
   Next i
   DeleteObject hBrush
' terminate cards.dll
   cdtTerm
   
   Set frmMain = Nothing
End Sub

Private Sub frHelp_Click()
   frHelp.Visible = False
End Sub

Private Sub lblHelp_Click()
   frHelp.Visible = False
End Sub

Private Sub mnuAnimate_Click(Index As Integer)
' uncheck previous and check new
   mnuAnimate(iOldAnimateIndex).Checked = False
   mnuAnimate(Index).Checked = True
   iOldAnimateIndex = Index
   
   swAnimate = True        ' assume animation
   If Index = 0 Then
      swAnimate = False
   Else
      iAnimateSpeed = Index
      Me.DrawWidth = Index
   End If
End Sub

Private Sub mnuBacks_Click(Index As Integer)
   iBack = 53 + Index
   cdtDraw Me.hdc, 200, 400, iBack, ordBacks, 0
   Me.Refresh
End Sub

Private Sub mnuExit_Click()
   subGameOverStop
  Unload Me
End Sub

Private Sub mnuHelp_Click()
   subGameOverStop
   frHelp.Visible = True
End Sub

Private Sub mnuNewGame_Click()
   subGameOverStop
   subNewGame
End Sub

Private Sub mnuPlaidColor_Click(Index As Integer)
' Card (back) 53 uses the Color paramater of API cdtDraw
   Dim iClr As Integer
   
   Select Case Index
      Case 0: iClr = 4
      Case 1: iClr = 1
      Case 2: iClr = 3
      Case 3: iClr = 6
      Case 4: iClr = 5
      Case 5: iClr = 7
   End Select
   cdtDraw Me.hdc, 200, 400, 53, ordBacks, QBColor(iClr)
   Me.Refresh
End Sub

Private Sub mnuReplay_Click()
   subNewGame cReplay
End Sub

Private Sub mnuSounds_Click()
   mnuSounds.Checked = Not mnuSounds.Checked
   swSound = mnuSounds.Checked
End Sub

Private Sub tmrGameOver_Timer()
   Static i As Integer
   
   i = i + 1
   If i > 7 Then
      i = 0
   End If
   
   lblGameOver(i).Visible = Not lblGameOver(i).Visible
End Sub

Private Sub tmrInvalidSelection_Timer()
   Static iCount As Integer
   
   lblInvalidSelection.Visible = Not lblInvalidSelection.Visible
   iCount = iCount + 1
   If iCount > 7 Then
      tmrInvalidSelection.Enabled = False
      iCount = 0
      subDrawCard FirstCard.Index, ordFaces
      subDrawCard SecondCard.Index, ordFaces
      FirstCard.Index = -1
      SecondCard.Index = -1
   End If
End Sub

Private Sub tmrRemoveCard_Timer()
   With rectRemoveCard
      Me.Line (.Left + iCntRemoveCard, .Top + iCntRemoveCard)-(.Right - iCntRemoveCard, .Bottom - iCntRemoveCard), Me.BackColor, B
      Me.Refresh
      End With
   iCntRemoveCard = iCntRemoveCard + iAnimateSpeed
   If iCntRemoveCard > 36 Then
      tmrRemoveCard = False
   End If

End Sub

'======================================================================================
'
'                                 LOCAL PROCEDURES
'______________________________________________________________________________________

Private Sub subDrawCard(sIndex As Integer, sState As Integer)
   If Not arPlayboard(sIndex).Discarded Then
      cdtDraw Me.hdc, arPlayboard(sIndex).Left, arPlayboard(sIndex).Top, arPlayboard(sIndex).Number, sState, vbWhite
      Me.Refresh
   End If
End Sub

Private Sub subGameOverStop()
   Dim i As Integer
   
   tmrGameOver.Enabled = False
   For i = 0 To 7
      lblGameOver(i).Visible = False
   Next i
End Sub

Private Sub subNewGame(Optional sReplay As Boolean)

   Dim Temp       As Integer
   Dim ItemPicked As Integer
   Dim Remaining   As Integer
   Dim i As Integer
   
 ' clear the playboard
    Me.Cls
    FirstCard.Index = -1
    SecondCard.Index = -1
    
 ' Skip shuffle if we replay
    If Not sReplay Then
       ' load values into array , cardnumber(1) = 1, etc
      For i = 0 To 51
         arCards(i) = i
      Next i
   
      ' shuffle this array
      For i = 51 To 1 Step -1
         ItemPicked = Int(Rnd * i)        ' pick a card from cards remaining
         Temp = arCards(i)                ' get bottom card and put it as temp
         arCards(i) = arCards(ItemPicked) ' move picked card to bottom
         arCards(ItemPicked) = Temp       ' put (saved) bottom card
      Next i
   End If
   
'
' first 28 cards are shown on the "playing" area and we save card characteristics
   For i = 0 To 27
      cdtDraw Me.hdc, arPlayboard(i).Left, arPlayboard(i).Top, arCards(i), ordFaces, 0
      arPlayboard(i).Number = arCards(i)
      arPlayboard(i).Value = arCards(i) \ 4 + 1
      ' each card is covered by two cards in the next row, the last row excepted
      arPlayboard(i).Covered = IIf(i < 21, 2, 0)
      arPlayboard(i).Discarded = False
   Next i
   Me.Refresh
' save characteristics of remaining cards (deck)
   For i = 28 To 51
      arDeck(i - 28).Number = arCards(i)
      arDeck(i - 28).Value = arCards(i) \ 4 + 1
      arDeck(i - 28).Discarded = False
   Next i
   
' the covered deck
   cdtDraw Me.hdc, 200, 400, iBack, ordBacks, 0
' the uncovered deck
   FillRgn Me.hdc, arRgn(28), hBrush
   Me.Refresh
   arPlayboard(28).Covered = 0
   iCurrentDeck = -1
End Sub

Private Sub subProcessSelection()

' if the same card has been clicked undo the highligthing and reset
   If FirstCard.Index = SecondCard.Index Then
      subDrawCard FirstCard.Index, ordFaces
      FirstCard.Index = -1
      SecondCard.Index = -1
      Exit Sub
   End If
   
' do both cards add up to 13
   If FirstCard.Value + SecondCard.Value = 13 Then
' Yes: remove first and last card selected in descending sequence
      If FirstCard.Index > SecondCard.Index Then
         subRemoveCard FirstCard
         subRemoveCard SecondCard
      Else
         subRemoveCard SecondCard
         subRemoveCard FirstCard
      End If
      
' No: undo selection
   Else
      If swSound Then
         PlayResWAV 102, SND_ASYNC + SND_RESOURCE
      End If
      tmrInvalidSelection.Enabled = True
   End If
   
End Sub


Private Sub subRemoveCard(sCard As SELECTED)
   Dim iRow As Integer
   Dim iIndex As Integer
   
' preserve sCard.Index and flag it as processed
   iIndex = sCard.Index
   sCard.Index = -1
   If swSound Then
    PlayResWAV 101, SND_ASYNC + SND_RESOURCE
   End If
   If swAnimate Then
      GetRgnBox arRgn(iIndex), rectRemoveCard
      iCntRemoveCard = 0
      tmrRemoveCard.Enabled = True
      Do   ' wait for the timer to finish its job before proceeding
        DoEvents
      Loop Until tmrRemoveCard.Enabled = False
   End If
   
' deck (show underlying card)
   If iIndex = 28 Then
      arDeck(iCurrentDeck).Discarded = True
      Do
         iCurrentDeck = iCurrentDeck - 1
         ' uncovered deck exhausted
         If iCurrentDeck < 0 Then
            FillRgn Me.hdc, arRgn(28), hBrush
           If swSound Then
              PlayResWAV 101, SND_ASYNC + SND_RESOURCE
           End If
          Me.Refresh
            Exit Do
         End If
         If Not arDeck(iCurrentDeck).Discarded Then
            arPlayboard(28).Number = arDeck(iCurrentDeck).Number
            arPlayboard(28).Value = arDeck(iCurrentDeck).Value
            subDrawCard 28, ordFaces
            Exit Do
         End If
      Loop Until iCurrentDeck = -1
      Exit Sub
   End If

' playboard : remove card and region
   If swSound Then
    PlayResWAV 101, SND_ASYNC + SND_RESOURCE
   End If
   If swAnimate Then
      GetRgnBox arRgn(iIndex), rectRemoveCard
      iCntRemoveCard = 0
      tmrRemoveCard.Enabled = True
      Do   ' wait for the timer to finish its job before proceeding
        DoEvents
      Loop Until tmrRemoveCard.Enabled = False
   Else
      FillRgn Me.hdc, arRgn(iIndex), hBrush
   End If
   arPlayboard(iIndex).Discarded = True
   
' now check if last card has been removed
   If iIndex = 0 Then
      Me.Cls
      If swSound Then
         PlayResWAV 103, SND_ASYNC + SND_RESOURCE
      End If
      tmrGameOver.Enabled = True
      Exit Sub
   End If
   
' repaint surroundng cards and set uncovered cards status
   Select Case iIndex
      Case 1, 2:     iRow = 1
      Case 3 To 5:   iRow = 2
      Case 6 To 9:   iRow = 3
      Case 10 To 14: iRow = 4
      Case 15 To 20: iRow = 5
      Case 21 To 27: iRow = 6
   End Select
   
' any slot in a row except last: repaint up rigth and next right
   If iIndex < arRow(iRow) + iRow Then
      subDrawCard iIndex - iRow, ordFaces
      arPlayboard(iIndex - iRow).Covered = arPlayboard(iIndex - iRow).Covered - 1
      subDrawCard iIndex + 1, ordFaces
   End If
   
' any slot in a row except first: repaint up left and next left
   If iIndex > arRow(iRow) Then
      subDrawCard iIndex - iRow - 1, ordFaces
      arPlayboard(iIndex - iRow - 1).Covered = arPlayboard(iIndex - iRow - 1).Covered - 1
      subDrawCard iIndex - 1, ordFaces
   End If
   
' cards below the row of the processed card may have been overwritten
   If iRow < 6 Then
      For iIndex = arRow(iRow + 1) To 27
         subDrawCard iIndex, ordFaces
      Next iIndex
   End If
   
End Sub
