VERSION 5.00
Begin VB.Form frmmain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AlphaFighter"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9030
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmmain.frx":1CCA
   ScaleHeight     =   395
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   602
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6840
      TabIndex        =   22
      Top             =   5520
      Width           =   855
   End
   Begin VB.PictureBox picblowupmask2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   3840
      Picture         =   "frmmain.frx":3994
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picblowup2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   3360
      Picture         =   "frmmain.frx":4142
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5460
      Left            =   0
      MouseIcon       =   "frmmain.frx":48F0
      MousePointer    =   99  'Custom
      Picture         =   "frmmain.frx":65BA
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   0
      Width           =   9060
      Begin VB.PictureBox piclasermask2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   210
         Left            =   1440
         Picture         =   "frmmain.frx":A493C
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   3
         TabIndex        =   20
         Top             =   3480
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.PictureBox piclaser2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   210
         Left            =   1200
         Picture         =   "frmmain.frx":A49AE
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   3
         TabIndex        =   19
         Top             =   3480
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.PictureBox picblowupmask5 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   435
         Left            =   3480
         Picture         =   "frmmain.frx":A4A68
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   17
         Top             =   2760
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox picblowup5 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   435
         Left            =   3000
         Picture         =   "frmmain.frx":A5216
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   16
         Top             =   2760
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox picblowupmask4 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   435
         Left            =   6000
         Picture         =   "frmmain.frx":A59C4
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   15
         Top             =   4200
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox picblowup4 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   435
         Left            =   5520
         Picture         =   "frmmain.frx":A6172
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   14
         Top             =   4200
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox picblowupmask3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   435
         Left            =   5520
         Picture         =   "frmmain.frx":A6920
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   13
         Top             =   3480
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox picblowup3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   435
         Left            =   5040
         Picture         =   "frmmain.frx":A70CE
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   12
         Top             =   3480
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox picship 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   465
         Left            =   0
         Picture         =   "frmmain.frx":A787C
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   9
         Top             =   4800
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox picshipmask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   465
         Left            =   480
         Picture         =   "frmmain.frx":A80C2
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   8
         Top             =   4800
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox picen 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   435
         Left            =   0
         Picture         =   "frmmain.frx":A8908
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   7
         Top             =   4200
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox picenmask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   435
         Left            =   480
         Picture         =   "frmmain.frx":A90B6
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   6
         Top             =   4200
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox picblowup1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   435
         Left            =   1440
         Picture         =   "frmmain.frx":A9864
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   5
         Top             =   4200
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox picblowupmask1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   435
         Left            =   1920
         Picture         =   "frmmain.frx":AA012
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   4
         Top             =   4200
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox piclaser 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   210
         Left            =   1440
         Picture         =   "frmmain.frx":AA7C0
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   3
         TabIndex        =   3
         Top             =   4800
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.PictureBox piclasermask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   210
         Left            =   1680
         Picture         =   "frmmain.frx":AA87A
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   3
         TabIndex        =   2
         Top             =   4800
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.Label lblwave 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wave:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   360
      Left            =   4440
      TabIndex        =   21
      Top             =   5520
      Width           =   795
   End
   Begin VB.Image imglives 
      Height          =   405
      Index           =   0
      Left            =   0
      Picture         =   "frmmain.frx":AA8EC
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblarmor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Armor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   360
      Left            =   1845
      TabIndex        =   18
      Top             =   5520
      Width           =   855
   End
   Begin VB.Image imglives 
      Height          =   405
      Index           =   3
      Left            =   120
      Picture         =   "frmmain.frx":AB132
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imglives 
      Height          =   405
      Index           =   2
      Left            =   600
      Picture         =   "frmmain.frx":AB978
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imglives 
      Height          =   405
      Index           =   1
      Left            =   1080
      Picture         =   "frmmain.frx":AC1BE
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label lblpoints 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   360
      Left            =   8040
      TabIndex        =   1
      Top             =   5520
      Width           =   900
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
    Running = False
    End
End Sub

Private Sub Form_Load()
    Randomize 'randomizes every thing randomly (redundant, huh?)
    'this starts the initial stuff before the game.
    Me.Show 'show game screen
    Running = True
    SETBLUR = 1 'blur amount set to 1 for no blur
    Blur = SETBLUR 'decides how much blur or 'trailing' there will be in game
    STARSTART 'plot stars
    Points = 0 'set players points to 0
    Ship.Lives = 3 'set players lives to 3
    WaveCount = 150 'time before the first wave starts
    graphicsstyle = 0 'for player ship to alternate graphics to make it look like thrusters are going
    WaveStart = True  'this will be to initiate the wave stuff
    Ship.armor = 3  'this sets players armor to 3
    Initiate = True   'this is yet another thing for the waves before the loops of them moving back and forth starts
    WAVE = 1   'start on first wave
    Dim i As Integer
    Do
        EnemiesDestroyed = True   'set enemies destroyed to true to begin with
        DoEvents 'allows other events to occur
        If WaveCount <> 0 Then          'if the timer for the next wave is not 0, then subtract from it until it is
            WaveCount = WaveCount - 1
        Else
            If WAVE = 1 Then            'load each wave, depending on which wave you are on
                EnemyWave1
            End If
            If WAVE = 2 Then
                EnemyWave2
            End If
            If WAVE = 3 Then  'introduces double layered waves
                EnemyWave3
            End If
            If WAVE = 4 Then 'introduces shooting enemies
                EnemyWave4
            End If
            If WAVE = 5 Then
                EnemyWave5
            End If
            If WAVE = 6 Then  'double layered shooting enemies
                EnemyWave6
            End If
            If WAVE = 7 Then  'a wedged formation
                EnemyWave7
            End If
            If WAVE = 8 Then  'enemies move down
                EnemyWave8
            End If
            If WAVE = 9 Then 'enemies have higher chance of firing, and they move down
                EnemyWave9
            End If
            If WAVE = 10 Then 'enemies move down faster, and higher chance of shooting
                EnemyWave10
            End If
            For i = 1 To NumEnemies
                If Enemies(i).Destroyed = False Then 'check to see if all of the enemies of the wave are destroyed
                    EnemiesDestroyed = False  'if not then enemies destroyed = false and continue with the wave
                End If
            Next
            If EnemiesDestroyed = True Then 'if all enemies are gone, start next wave
                WaveCount = 50          'time until next wave
                WAVE = WAVE + 1            'raise the wave level
                Initiate = True           'sets intial wave settings
                WaveStart = True          'also sets intial wave settings
                For i = 1 To 3
                    ShipShoot(i).X = -10    'all of the players lasers dissapear for
                    ShipShoot(i).Y = -10    'the beginning of the next wave.
                    Ship.Shoot(i) = False
                Next
            End If
        End If
            STARMOVE                        'loops the stars moving
            Collision                       'detects for 3 different types of collision 1 is for player hitting enemy with laser
            Collision2                      '2 is for enemy hitting player
            Collision3                      '3 is for ship collision
            Draw                            'draw EVERYTHING seen to play
    Loop While Running = True
End Sub


'**********************************************************************
'*Draws everything seen on screen (except the bottom game information)
'**********************************************************************

Sub Draw()
    'decides how much blur there will be to game (for certain waves
    Blur = Blur - 1
    If Blur = 0 Then
        picmain.Cls 'clear screen
        Blur = SETBLUR
    End If
    lblarmor.Caption = "Armor:" & Ship.armor 'display armor
    Dim i As Integer
    
    For i = 1 To SN
        picmain.PSet (Stars(i).X, Stars(i).Y), vbWhite 'draw stars
    Next
    
    For i = 1 To NumEnemies
        If Enemies(i).Destroyed = False Then
            RetVal = BitBlt(picmain.hDC, Enemies(i).X, Enemies(i).Y, picen.Width, picen.Height, picenmask.hDC, 0, 0, vbSrcAnd)   'draw the enemy ships
            RetVal = BitBlt(picmain.hDC, Enemies(i).X, Enemies(i).Y, picen.Width, picen.Height, picen.hDC, 0, 0, vbSrcPaint)
        End If
        If Enemies(i).explode >= 1 And Enemies(i).explode < 5 Then
            RetVal = BitBlt(picmain.hDC, Enemies(i).X, Enemies(i).Y, picen.Width, picen.Height, picblowupmask1.hDC, 0, 0, vbSrcAnd) 'show explosion
            RetVal = BitBlt(picmain.hDC, Enemies(i).X, Enemies(i).Y, picen.Width, picen.Height, picblowup1.hDC, 0, 0, vbSrcPaint)   'if they are destroyed
            Enemies(i).explode = Enemies(i).explode + 1
            Enemies(i).Destroyed = True
        ElseIf Enemies(i).explode >= 5 And Enemies(i).explode < 10 Then
            RetVal = BitBlt(picmain.hDC, Enemies(i).X, Enemies(i).Y, picen.Width, picen.Height, picblowupmask2.hDC, 0, 0, vbSrcAnd)
            RetVal = BitBlt(picmain.hDC, Enemies(i).X, Enemies(i).Y, picen.Width, picen.Height, picblowup2.hDC, 0, 0, vbSrcPaint)
            Enemies(i).explode = Enemies(i).explode + 1
        ElseIf Enemies(i).explode >= 10 And Enemies(i).explode < 15 Then
            RetVal = BitBlt(picmain.hDC, Enemies(i).X, Enemies(i).Y, picen.Width, picen.Height, picblowupmask3.hDC, 0, 0, vbSrcAnd)
            RetVal = BitBlt(picmain.hDC, Enemies(i).X, Enemies(i).Y, picen.Width, picen.Height, picblowup3.hDC, 0, 0, vbSrcPaint)
            Enemies(i).explode = Enemies(i).explode + 1
        ElseIf Enemies(i).explode >= 15 And Enemies(i).explode < 20 Then
            RetVal = BitBlt(picmain.hDC, Enemies(i).X, Enemies(i).Y, picen.Width, picen.Height, picblowupmask4.hDC, 0, 0, vbSrcAnd)
            RetVal = BitBlt(picmain.hDC, Enemies(i).X, Enemies(i).Y, picen.Width, picen.Height, picblowup4.hDC, 0, 0, vbSrcPaint)
            Enemies(i).explode = Enemies(i).explode + 1
        ElseIf Enemies(i).explode >= 20 And Enemies(i).explode < 25 Then
            RetVal = BitBlt(picmain.hDC, Enemies(i).X, Enemies(i).Y, picen.Width, picen.Height, picblowupmask5.hDC, 0, 0, vbSrcAnd)
            RetVal = BitBlt(picmain.hDC, Enemies(i).X, Enemies(i).Y, picen.Width, picen.Height, picblowup5.hDC, 0, 0, vbSrcPaint)
            Enemies(i).explode = Enemies(i).explode + 1
        End If
    Next
        Ship.graphicsstyle = Ship.graphicsstyle + 1                   'displays player ship
        If Ship.graphicsstyle <= 10 Then
            picship.Picture = LoadPicture(App.Path & "/ship1.bmp")
            picshipmask.Picture = LoadPicture(App.Path & "/ship1mask.bmp") 'thruster 1
        Else
            picship.Picture = LoadPicture(App.Path & "/ship2.bmp")
            picshipmask.Picture = LoadPicture(App.Path & "/ship2mask.bmp") 'thruster 2
            If Ship.graphicsstyle >= 20 Then
                Ship.graphicsstyle = 0
            End If
        End If
        RetVal = BitBlt(picmain.hDC, Ship.X, Ship.Y, picen.Width, picen.Height, picshipmask.hDC, 0, 0, vbSrcAnd)
        RetVal = BitBlt(picmain.hDC, Ship.X, Ship.Y, picen.Width, picen.Height, picship.hDC, 0, 0, vbSrcPaint) 'actually draw ship
    For i = 1 To 3
        If Ship.Shoot(i) = True Then
            ShipShoot(i).Y = ShipShoot(i).Y - 5
            If ShipShoot(i).Y <= 0 Then
                Ship.Shoot(i) = False
            End If
            RetVal = BitBlt(picmain.hDC, ShipShoot(i).X, ShipShoot(i).Y, piclaser.Width, piclaser.Height, piclasermask.hDC, 0, 0, vbSrcAnd)
            RetVal = BitBlt(picmain.hDC, ShipShoot(i).X, ShipShoot(i).Y, piclaser.Width, piclaser.Height, piclaser.hDC, 0, 0, vbSrcPaint) 'draw lasers
        End If                                                                                                                            'if needed
    Next
    For i = 1 To NumEnemies
        If Enemies(i).Shoot = True Then
            Enshoot(i).Y = Enshoot(i).Y + 5
            If Enshoot(i).Y >= picmain.ScaleHeight Then
                Enemies(i).Shoot = False
            End If                                      'draws enemy lasers
            RetVal = BitBlt(picmain.hDC, Enshoot(i).X, Enshoot(i).Y, piclaser.Width, piclaser.Height, piclasermask2.hDC, 0, 0, vbSrcAnd)
            RetVal = BitBlt(picmain.hDC, Enshoot(i).X, Enshoot(i).Y, piclaser.Width, piclaser.Height, piclaser2.hDC, 0, 0, vbSrcPaint)
        End If
    Next
    lblpoints.Caption = Format(Points, "000000")
    lblwave.Caption = "Wave:" & WAVE
End Sub

Private Sub Form_Terminate()
    Running = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Running = False
End Sub

Private Sub picmain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 1 To 3
        If Ship.Shoot(i) = False Then
            Ship.Shoot(i) = True
            ShipShoot(i).X = Ship.X + 12.5 'if left clicked, then initiate laser beam
            ShipShoot(i).Y = Ship.Y
            Exit Sub
        End If
    Next
End Sub

Private Sub picmain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ship.X = X 'contols the ship
    Ship.Y = Y
End Sub

Public Sub Collision()
    Dim i As Integer
    Dim j As Integer                    'checks for collision of player laser against enemy ship
    For j = 1 To NumEnemies
        For i = 1 To 3
            If ShipShoot(i).X >= Enemies(j).X And ShipShoot(i).X <= Enemies(j).X + 25 And _
            ShipShoot(i).Y >= Enemies(j).Y And ShipShoot(i).Y <= Enemies(j).Y + 25 And Enemies(j).Destroyed = False Then
                Enemies(j).armor = Enemies(j).armor - 1 'subtract armor from enemy
                Ship.Shoot(i) = False
                ShipShoot(i).X = -10  'laser beam disappears
                ShipShoot(i).Y = -10
                If Enemies(j).armor = 0 Then
                    Enemies(j).explode = 1 'explode if no more armor
                    Points = Points + (5 * Enemies(i).PointModifier) 'add points
                End If
            End If
        Next
    Next
End Sub
Public Sub Collision2()
    Dim i As Integer
    For i = 1 To NumEnemies               'checks for collision of enemy laser against player ship
        If Enshoot(i).X >= Ship.X And Enshoot(i).X <= Ship.X + 25 And _
        Enshoot(i).Y >= Ship.Y And Enshoot(i).Y <= Ship.Y + 25 And Ship.Destroyed = False Then
        Enemies(i).Shoot = False
        Enshoot(i).X = -10
        Enshoot(i).Y = -10
        Ship.armor = Ship.armor - 1 'lower ur armor
        If Ship.armor = 0 Then
            Ship.armor = 3                       'if no more armor, lose life reset armor
            Ship.Lives = Ship.Lives - 1
            Ship.explode = 1
            If Ship.Lives = 0 Then
                MsgBox "Sorry, You Lose" 'if no lives, you lose
                End
            End If
            imglives(Ship.Lives + 1).Visible = False
        End If
      End If
    Next
End Sub
Public Sub Collision3()
    Dim i As Integer          'although this doesn't work as good as i wanted it to, it checks for mid air collision (ship to ship)
    For i = 1 To NumEnemies
      If Enemies(i).X >= Ship.X And Enemies(i).X <= Ship.X + 25 And _
      Enemies(i).Y >= Ship.Y And Enemies(i).Y <= Ship.Y + 25 And Ship.Destroyed = False And _
      Enemies(i).Destroyed = False Then
        Enemies(i).armor = Enemies(i).armor - 1     'subtract enemy armor
        Ship.armor = Ship.armor - 1   'subtract ur armor
        If Enemies(i).armor = 0 Then
            Enemies(i).explode = 1
            Points = Points + (5 * Enemies(i).PointModifier) 'add to points if enemy is destroyed
        End If
        If Ship.armor = 0 Then
            Ship.armor = 3
            Ship.Lives = Ship.Lives - 1  'if u are destroyed then subtract lives
            Ship.explode = 1
            If Ship.Lives = 0 Then
                MsgBox "Sorry, You Lose" 'if 0 lives, u lose
                End
            End If
            imglives(Ship.Lives + 1).Visible = False
        End If
      End If
    Next
End Sub


