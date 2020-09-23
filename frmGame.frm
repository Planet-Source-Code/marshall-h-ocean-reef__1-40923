VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ocean Reef"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   15270
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   294
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picStatus 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   236
      TabIndex        =   1
      Top             =   2325
      Width           =   3540
      Begin VB.Image invFlippers 
         Height          =   225
         Left            =   3135
         Picture         =   "frmGame.frx":0442
         Top             =   45
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblLevel 
         BackStyle       =   0  'Transparent
         Caption         =   "Level: 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   990
         TabIndex        =   4
         Top             =   45
         Width           =   750
      End
      Begin VB.Image invKey 
         Height          =   105
         Left            =   2850
         Picture         =   "frmGame.frx":0754
         Top             =   105
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Shape shpInv 
         Height          =   315
         Left            =   2775
         Top             =   0
         Width           =   675
      End
      Begin VB.Label lblHealth 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Health: 100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1620
         TabIndex        =   3
         Top             =   45
         Width           =   1095
      End
      Begin VB.Label lblScore 
         BackStyle       =   0  'Transparent
         Caption         =   "Score: 0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   90
         TabIndex        =   2
         Top             =   45
         Width           =   1035
      End
   End
   Begin VB.PictureBox Pb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   2340
      Left            =   0
      Picture         =   "frmGame.frx":08E6
      ScaleHeight     =   156
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   18000
      Begin VB.Image imgYouWon 
         Height          =   1050
         Left            =   -30
         Picture         =   "frmGame.frx":12D9
         Top             =   690
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Image Flippers 
         Height          =   225
         Left            =   6660
         Picture         =   "frmGame.frx":174B
         Top             =   1980
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image FirstAid 
         Height          =   210
         Left            =   6210
         Picture         =   "frmGame.frx":1ACD
         Top             =   1980
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image Fish 
         Height          =   225
         Index           =   15
         Left            =   5565
         Picture         =   "frmGame.frx":1E3E
         Top             =   1695
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Fish 
         Height          =   225
         Index           =   14
         Left            =   1035
         Picture         =   "frmGame.frx":21CF
         Top             =   345
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Fish 
         Height          =   225
         Index           =   13
         Left            =   450
         Picture         =   "frmGame.frx":2560
         Top             =   1890
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Fish 
         Height          =   225
         Index           =   12
         Left            =   9375
         Picture         =   "frmGame.frx":28F1
         Top             =   585
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Fish 
         Height          =   225
         Index           =   11
         Left            =   9480
         Picture         =   "frmGame.frx":2C82
         Top             =   1740
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Fish 
         Height          =   225
         Index           =   10
         Left            =   14895
         Picture         =   "frmGame.frx":3013
         Top             =   1845
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Fish 
         Height          =   225
         Index           =   9
         Left            =   8250
         Picture         =   "frmGame.frx":33A4
         Top             =   510
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Fish 
         Height          =   225
         Index           =   8
         Left            =   7950
         Picture         =   "frmGame.frx":3735
         Top             =   1335
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Fish 
         Height          =   225
         Index           =   7
         Left            =   0
         Picture         =   "frmGame.frx":3AC6
         Top             =   1140
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Fish 
         Height          =   225
         Index           =   6
         Left            =   -150
         Picture         =   "frmGame.frx":3E57
         Top             =   1545
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Fish 
         Height          =   225
         Index           =   5
         Left            =   -105
         Picture         =   "frmGame.frx":41E8
         Top             =   345
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Fish 
         Height          =   225
         Index           =   4
         Left            =   12900
         Picture         =   "frmGame.frx":4579
         Top             =   405
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Fish 
         Height          =   225
         Index           =   3
         Left            =   14310
         Picture         =   "frmGame.frx":490A
         Top             =   405
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Fish 
         Height          =   225
         Index           =   2
         Left            =   11625
         Picture         =   "frmGame.frx":4C9B
         Top             =   1575
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Fish 
         Height          =   225
         Index           =   1
         Left            =   11115
         Picture         =   "frmGame.frx":502C
         Top             =   165
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Fish 
         Height          =   225
         Index           =   0
         Left            =   5355
         Picture         =   "frmGame.frx":53BD
         Top             =   150
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgGameOver 
         Height          =   1050
         Left            =   0
         Picture         =   "frmGame.frx":574E
         Top             =   675
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Image Shark 
         Height          =   345
         Left            =   4035
         Picture         =   "frmGame.frx":5B60
         Top             =   1560
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Image Key 
         Height          =   105
         Left            =   1380
         Picture         =   "frmGame.frx":5F5C
         Top             =   2130
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image Chest 
         Height          =   375
         Left            =   2700
         Picture         =   "frmGame.frx":62B7
         Top             =   1830
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Diver 
         Height          =   330
         Left            =   885
         Picture         =   "frmGame.frx":6691
         Top             =   1110
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "These are the graphics. Notice I made masks for every bitmap."
      Height          =   225
      Left            =   180
      TabIndex        =   5
      Top             =   3975
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Image imgFlippersMask 
      Height          =   225
      Left            =   3135
      Picture         =   "frmGame.frx":6A7D
      Top             =   3270
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgFlippers 
      Height          =   225
      Left            =   3135
      Picture         =   "frmGame.frx":6DDE
      Top             =   3030
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgFirstAidMask 
      Height          =   210
      Left            =   2790
      Picture         =   "frmGame.frx":7160
      Top             =   3255
      Width           =   210
   End
   Begin VB.Image imgFirstAid 
      Height          =   210
      Left            =   2790
      Picture         =   "frmGame.frx":74AE
      Top             =   3030
      Width           =   210
   End
   Begin VB.Image imgFishMask 
      Height          =   225
      Left            =   2760
      Picture         =   "frmGame.frx":781F
      Top             =   3585
      Width           =   240
   End
   Begin VB.Image imgFish 
      Height          =   225
      Left            =   2370
      Picture         =   "frmGame.frx":7B8D
      Top             =   3570
      Width           =   240
   End
   Begin VB.Image imgSharkMask 
      Height          =   345
      Left            =   3510
      Picture         =   "frmGame.frx":7F1E
      Top             =   3525
      Width           =   990
   End
   Begin VB.Image imgShark 
      Height          =   345
      Left            =   3510
      Picture         =   "frmGame.frx":82D9
      Top             =   3060
      Width           =   990
   End
   Begin VB.Image imgKeyMask 
      Height          =   105
      Left            =   2370
      Picture         =   "frmGame.frx":86D5
      Top             =   3255
      Width           =   225
   End
   Begin VB.Image imgKey 
      Height          =   105
      Left            =   2370
      Picture         =   "frmGame.frx":8A2A
      Top             =   3075
      Width           =   225
   End
   Begin VB.Image imgChestMask 
      Height          =   375
      Left            =   1755
      Picture         =   "frmGame.frx":8D85
      Top             =   3510
      Width           =   480
   End
   Begin VB.Image imgChest 
      Height          =   375
      Left            =   1740
      Picture         =   "frmGame.frx":910C
      Top             =   3060
      Width           =   480
   End
   Begin VB.Image imgDiver2Mask 
      Height          =   330
      Left            =   975
      Picture         =   "frmGame.frx":94E6
      Top             =   3525
      Width           =   615
   End
   Begin VB.Image imgDiver2 
      Height          =   330
      Left            =   195
      Picture         =   "frmGame.frx":987C
      Top             =   3540
      Width           =   615
   End
   Begin VB.Image imgDiver1Mask 
      Height          =   330
      Left            =   990
      Picture         =   "frmGame.frx":9C55
      Top             =   3060
      Width           =   615
   End
   Begin VB.Image imgDiver1 
      Height          =   330
      Left            =   195
      Picture         =   "frmGame.frx":9FFC
      Top             =   3060
      Width           =   615
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHighScores 
         Caption         =   "&High Scores"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnumHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************
'**Ocean Reef by Marshall Hecht at GameSoft**
'********************************************
'**------------This is freeware------------**
'********************************************

'Any mistakes are because I'm testing my new
'PC speakers to see how loud they can go.
'I didn't use any stupid timers because they
'screw up under XP.

'If you thought a 12 year old kid like me
'could never make a game, well, tough. :-)

'This is TIGHT CODE - compiles to only 130 KB.

Private Sub Form_Load()
    Me.Width = 3600
    Me.Height = 3345
    SetUp_OceanReef
End Sub

Sub SetUp_OceanReef()
    'set up player stats
    Score = 0
    Health = 100
    Level = 1
    
    imgGameOver.Visible = False
    imgYouWon.Visible = False
    
    SwimAmount = 0.5
    FishSpeed = 0.2
    BossSpeed = 0.8
    
    Diver.Left = 59
    Diver.Top = 74
    
    Flippers.Left = -20
    FirstAid.Left = (Int((600 * Rnd) + 150))
    Chest.Left = 1100
    Chest.Top = 122
    
    Shark.Left = 1050
    Shark.Top = 122
    HaveKey = False
    HaveFlippers = False
    
    BossDir = "UP"
    Key.Left = (Int((500 * Rnd) + 200))
    GroundY = 130
    
    Pb.Left = 0
     
     'update status
    lblHealth.Caption = "Health: " & Health & "%"
    lblScore.Caption = "Score: " & Score
    lblLevel.Caption = "Level: " & Level
    If HaveKey = True Then invKey.Visible = True
    If HaveKey = False Then invKey.Visible = False
    If HaveFlippers = False Then invFlippers.Visible = False
    If HaveFlippers = True Then invFlippers.Visible = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If GameOn = True Then
        If MsgBox("Are you sure you want to quit?", vbYesNo, "Confirm Exit") = vbYes Then
            End 'automatically unloads everything
            Unload frmAbout
            Unload frmGame
            Unload frmHelp
            Unload frmHighScore
        Else
            Cancel = 1
        End If
    ElseIf GameOn = False Then
        Unload frmAbout
        Unload frmGame
        Unload frmHelp
        Unload frmHighScore
        End
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuHelp_Click()
    frmHelp.Show
End Sub

Sub GameLoop()
    Do While GameOn = True
        
        If tTime + ms_Delay <= GetTickCount() Then
        'Reset the timer
        tTime = GetTickCount()
        
        '******MAIN LOOP*******
        DiverMoving = False
        '*******KEYBOARD********
        '---------------------------------------------------------------
        '****Check for keypresses****
        If (GetKeyState(vbKeyUp) And &H1000) Then
            'user pressed up key
            Diver.Top = Diver.Top - SwimAmount
            DiverMoving = True
            'DiverDir = "UP"
        End If
                
        If (GetKeyState(vbKeyDown) And &H1000) Then
            'user pressed down key
            If Diver.Top < GroundY Then Diver.Top = Diver.Top + SwimAmount
            'DiverDir = "DOWN"
            DiverMoving = True
        End If
        
        If (GetKeyState(vbKeyLeft) And &H1000) Then
            'user pressed left key
            Diver.Left = Diver.Left - SwimAmount
            DiverDir = "LEFT"
            If AtBoss = False Then Pb.Left = Pb.Left + SwimAmount 'move picturebox across form
            DiverMoving = True
        End If
        
        If (GetKeyState(vbKeyRight) And &H1000) Then
            'user pressed right key
            Diver.Left = Diver.Left + SwimAmount
            DiverDir = "RIGHT"
            If AtBoss = False Then Pb.Left = Pb.Left - SwimAmount
            DiverMoving = True
        End If
        
        '*******LOGIC*******
        '---------------------------------------------------------------
        If DiverStage > 99 Then
            DiverAction = Not (DiverAction)
            DiverStage = 0
        End If
             
        '*****counter*****
        If DiverMoving = True Then
            DiverStage = DiverStage + 3 'if diver moving then speed up anim
        ElseIf DiverMoving = False Then
            DiverStage = DiverStage + 1 'else count slower
        End If
        
        If Diver.Left > 980 Then AtBoss = True Else AtBoss = False 'if diver is within boundaries, turn boss on
        
        If Diver.Left + Diver.Width < Key.Left + 20 And Diver.Left + Diver.Width > Key.Left - 10 Then
            If Diver.Top > 120 Then
                HaveKey = True 'make key disappear
            End If
        End If
        
        If Diver.Left + Diver.Width < FirstAid.Left + 20 And Diver.Left + Diver.Width > FirstAid.Left - 10 Then
            If Diver.Top > 120 Then
                FirstAid.Left = -20 'make kit disappear
                Health = Health + 20
            End If
        End If
        
        If Diver.Left + Diver.Width < Flippers.Left + 20 And Diver.Left + Diver.Width > Flippers.Left - 10 Then
            If Diver.Top > 120 Then
                Flippers.Left = -20 'make flippers gone
                SwimAmount = 0.8
                HaveFlippers = True
            End If
        End If
        '
        If Diver.Left + Diver.Width < Chest.Left + 20 And Diver.Left + Diver.Width > Chest.Left - 10 Then
            If Diver.Top > 90 Then
                If HaveKey = True Then
                    Chest.Top = 300 'move chest out of screen
                    Score = Score + (Int((22 * Rnd) + 18))
                    NextLevel
                End If
            End If
        End If
        
        '***move shark***
        If HaveKey = True Then
            If BossDir = "UP" And Shark.Top > 10 Then
                Shark.Top = Shark.Top - BossSpeed
            Else
                BossDir = "DOWN"
            End If
            
            If BossDir = "DOWN" And Shark.Top < 120 Then
                Shark.Top = Shark.Top + BossSpeed
            Else
                BossDir = "UP"
            End If
        End If
        
        '***move fishes***
        For f = 0 To 15
            Fish(f).Left = Fish(f).Left - FishSpeed
            If Fish(f).Left < -12 Then Fish(f).Left = (Int((1200 * Rnd) + 1))
        Next f
        
        '*****check 4 collisions******
        If Diver.Left < Shark.Left + Shark.Width And Diver.Left + Diver.Width > Shark.Left Then
            If Diver.Top > Shark.Top - Diver.Height And Diver.Top < Shark.Top + Diver.Height Then
                'diver hit boss
                GameOver
            End If
        End If
        
        For c = 0 To 15
            If Diver.Left < Fish(c).Left + Fish(c).Width And Diver.Left + Diver.Width > Fish(c).Left Then
                If Diver.Top > Fish(c).Top - Diver.Height And Diver.Top + Diver.Height < Fish(c).Top + Diver.Height Then
                    'diver hit fish(c)
                    Health = Health - 1
                End If
            End If
        Next c
        
        '*****GRAPHICS******
        '---------------------------------------------------------------
        Pb.Cls 'clear picturebox
        
        '*******Paint diver*******
        If DiverAction = True Then
            If DiverDir = "RIGHT" Then
                Pb.PaintPicture imgDiver1Mask.Picture, Diver.Left, Diver.Top, , , , , , , vbMergePaint
                Pb.PaintPicture imgDiver1.Picture, Diver.Left, Diver.Top, , , , , , , vbSrcAnd
            ElseIf DiverDir = "LEFT" Then
                Pb.PaintPicture imgDiver1Mask.Picture, Diver.Left + imgDiver1.Width, Diver.Top, -imgDiver1.Width, , , , , , vbMergePaint
                Pb.PaintPicture imgDiver1.Picture, Diver.Left + imgDiver1.Width, Diver.Top, -imgDiver1.Width, , , , , , vbSrcAnd
            End If
        ElseIf DiverAction = False Then
            If DiverDir = "RIGHT" Then
                Pb.PaintPicture imgDiver2Mask.Picture, Diver.Left, Diver.Top, , , , , , , vbMergePaint
                Pb.PaintPicture imgDiver2.Picture, Diver.Left, Diver.Top, , , , , , , vbSrcAnd
            ElseIf DiverDir = "LEFT" Then
                Pb.PaintPicture imgDiver2Mask.Picture, Diver.Left + imgDiver2.Width, Diver.Top, -imgDiver2.Width, , , , , , vbMergePaint
                Pb.PaintPicture imgDiver2.Picture, Diver.Left + imgDiver2.Width, Diver.Top, -imgDiver2.Width, , , , , , vbSrcAnd
            End If
        End If
        
        '*******paint chest*******
        Pb.PaintPicture imgChestMask.Picture, Chest.Left, Chest.Top, , , , , , , vbMergePaint
        Pb.PaintPicture imgChest.Picture, Chest.Left, Chest.Top, , , , , , , vbSrcAnd
        
        '******paint key*******
        If HaveKey = False Then
            Pb.PaintPicture imgKeyMask.Picture, Key.Left, Key.Top, , , , , , , vbMergePaint
            Pb.PaintPicture imgKey.Picture, Key.Left, Key.Top, , , , , , , vbSrcAnd
        End If
        
        '*****paint shark******
        Pb.PaintPicture imgSharkMask.Picture, Shark.Left, Shark.Top, , , , , , , vbMergePaint
        Pb.PaintPicture imgShark.Picture, Shark.Left, Shark.Top, , , , , , , vbSrcAnd
        
        '*****paint fishes*****
        For i = 1 To 15
            Pb.PaintPicture imgFishMask.Picture, Fish(i).Left, Fish(i).Top, , , , , , , vbMergePaint
            Pb.PaintPicture imgFish.Picture, Fish(i).Left, Fish(i).Top, , , , , , , vbSrcAnd
        Next i
        
        '*****paint first aid*****
        Pb.PaintPicture imgFirstAidMask.Picture, FirstAid.Left, FirstAid.Top, , , , , , , vbMergePaint
        Pb.PaintPicture imgFirstAid.Picture, FirstAid.Left, FirstAid.Top, , , , , , , vbSrcAnd
        
        '*****paint flippers*****
        Pb.PaintPicture imgFlippersMask.Picture, Flippers.Left, Flippers.Top, , , , , , , vbMergePaint
        Pb.PaintPicture imgFlippers.Picture, Flippers.Left, Flippers.Top, , , , , , , vbSrcAnd
        
        lblHealth.Caption = "Health: " & Health & "%"
        lblScore.Caption = "Score: " & Score
        lblLevel.Caption = "Level: " & Level
        If HaveKey = True Then invKey.Visible = True
        If HaveKey = False Then invKey.Visible = False
        If HaveFlippers = True Then invFlippers.Visible = True
        If Health = 0 Or Health < 0 Then GameOver
    
        DoEvents 'process system messages so loop won't freeze system
        End If
    Loop
End Sub

Private Sub mnuHighScores_Click()
    frmHighScore.Show
End Sub

Private Sub mnuNewGame_Click()
   DiverDir = "RIGHT"
    GameOn = True
    SetUp_OceanReef
    GameLoop 'starts the main loop
End Sub

Private Sub GameOver()
    Health = 0
    
    'update status
    lblHealth.Caption = "Health: " & Health & "%"
    lblScore.Caption = "Score: " & Score
    lblLevel.Caption = "Level: " & Level
        
    imgGameOver.Left = Diver.Left - 40
    If AtBoss = True Then imgGameOver.Left = Pb.Width - imgGameOver.Width - 60
    imgGameOver.Visible = True
    GameOn = False
    frmHighScore.SaveHiScore "OceanReef", modMain.Score, True
End Sub

Private Sub NextLevel()
    'set up player stats
 
    imgGameOver.Visible = False
    imgYouWon.Visible = False
    
    Level = Level + 1
                                                               
    FishSpeed = FishSpeed + 0.2
    BossSpeed = BossSpeed + 0.1
    Diver.Left = 59
    Diver.Top = 74
    
    If Level = 3 Then Flippers.Left = (Int((600 * Rnd) + 150))
    FirstAid.Left = (Int((600 * Rnd) + 150))
    Chest.Left = 1100
    Chest.Top = 122
    
    Shark.Left = 1050
    Shark.Top = 122
    AtBoss = False
    HaveKey = False
    
    BossDir = "UP"
    Key.Left = (Int((500 * Rnd) + 200))
    GroundY = 130
    
    Pb.Left = 0
     
    'update status
    lblLevel.Caption = "Level: " & Level
    If HaveKey = True Then invKey.Visible = True
    If HaveKey = False Then invKey.Visible = False
    
     If Level = 6 Then
        imgYouWon.Left = Diver.Left - 40
        If AtBoss = True Then imgYouWon.Left = 1000
        imgYouWon.Visible = True
        GameOn = False
        frmHighScore.SaveHiScore "OceanReef", modMain.Score, True
    End If
End Sub


