Attribute VB_Name = "AlphaInfo"
Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
'bitblt stuff (to display graphics)
Public RetVal As Long

Public Type Coordinates           'set coordinate system for stuff like stars and lasers
    X As Single
    Y As Single
End Type

Public Type GameUnits
    X As Single                     'gameunits is ACTUALLY just the enemy ships
    Y As Single   'where ship is at moment
    
    Destroyed As Boolean               'is it destroyed?
    Shoot As Boolean                 'is it shooting?
    armor As Integer                  'sets armor
    explode As Integer              'exploding state
    LeftRight As Boolean           'is it moving left or right
    PointModifier As Integer             'point multiplier (higher levels, more points)
End Type
Public Type Hero
    X As Single
    Y As Single               'this is the hero type
    
    Destroyed As Boolean            'are u destroyed?
    Shoot(1 To 3) As Boolean         'u can only shoot up to 3 lasers at a time
    armor As Integer              'ur current armor
    Lives As Integer             'how many lives u got
    explode As Integer              'explode state (not really implemented)
    graphicsstyle As Integer           'which thruster pic is being shown
End Type

Public Blur As Integer                'sets blur amount of game
Public SETBLUR As Integer             'although not all that great a system, it's just an idea
Public EnemiesDestroyed As Boolean         'are all enemies destroyed for wave?
Public Initiate As Boolean                   'initial wave stuff
Public WaveCount As Integer                'initial wave stuff
Public Points As Long                      'players points
Public ShipShoot(1 To 3) As Coordinates       'players lasers coordinates (altho i coulda stuck it into the type)
Public Enshoot() As Coordinates              'enemy lasers coordinate (altho i coulda stuck it into the type)
Public Stars() As Coordinates                       'star coordinates
Public Const SN As Long = 50             'amount of stars that will appear in background (higher = more)
Public NumEnemies As Integer               'number of enemies per wave
Public Enemies() As GameUnits                           'enemy information
Public Ship As Hero                           'player information
Public WaveStart As Boolean                  'how long till next wave
Public WAVE As Integer                        'wave ur currently on
Public Running As Boolean

Public Sub STARSTART()                              'plot stars initially
    ReDim Stars(1 To SN) As Coordinates
    Dim i As Long
    Dim j As Long
    For i = 1 To SN                        'place them randomly
        Stars(i).X = Int(Rnd * frmmain.picmain.ScaleWidth)
        Stars(i).Y = Int(Rnd * frmmain.picmain.ScaleHeight)
    Next
    
End Sub
Public Sub STARMOVE()
    Dim i As Long
    For i = 1 To SN               'move stars
        Stars(i).Y = Stars(i).Y + 3
        If Stars(i).Y >= frmmain.picmain.ScaleHeight Then
            Stars(i).X = Int(Rnd * frmmain.picmain.ScaleWidth)
            Stars(i).Y = 0
        End If
    Next
End Sub


