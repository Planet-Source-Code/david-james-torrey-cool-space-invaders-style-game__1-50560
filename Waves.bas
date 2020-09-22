Attribute VB_Name = "Waves"
Option Explicit

Public RandomShoot As Long 'random enemy shooting
 
Public Sub EnemyWave1()         'first wave
    Dim i As Integer
    If WaveStart = True Then
        If Initiate = True Then
            NumEnemies = 10                          'set number enemies
            ReDim Enshoot(1 To NumEnemies) As Coordinates
            ReDim Enemies(1 To NumEnemies) As GameUnits
            For i = 1 To NumEnemies
                Enemies(i).armor = 1          'set enemy stats
                Enemies(i).LeftRight = True
                Enemies(i).Destroyed = False
                Enemies(i).Shoot = False
                Enemies(i).PointModifier = 1
                Enemies(i).explode = 0
                Enemies(i).Y = -20      'set initial position
                Enemies(i).X = i * 50
            Next
            Initiate = False 'end all intial stuff
        Else
            For i = 1 To NumEnemies
                Enemies(i).Y = Enemies(i).Y + 2 'have them move into screen slowly
                If Enemies(i).Y >= 30 Then
                    WaveStart = False
                End If
            Next
        End If
    Else
    For i = 1 To NumEnemies
        If Enemies(i).explode = 0 Then
            If Enemies(i).LeftRight = True Then 'makes them move
                Enemies(i).X = Enemies(i).X - 3 'rate they move left and right
            Else
                Enemies(i).X = Enemies(i).X + 3
            End If
        End If
        If Enemies(i).X <= 0 Then            'if they move left and hit wall, move in opposite direction, vice versa
            Enemies(i).LeftRight = False
        ElseIf Enemies(i).X + 25 >= frmmain.picmain.ScaleWidth Then
            Enemies(i).LeftRight = True
        End If
    Next
    
    End If
End Sub



Public Sub EnemyWave2()
    Dim i As Integer                 'same exact wave as first
    If WaveStart = True Then
        If Initiate = True Then
            NumEnemies = 10
            ReDim Enshoot(1 To NumEnemies) As Coordinates
            ReDim Enemies(1 To NumEnemies) As GameUnits
            For i = 1 To NumEnemies
                Enemies(i).armor = 1
                Enemies(i).LeftRight = True
                Enemies(i).PointModifier = 1
                Enemies(i).Destroyed = False
                Enemies(i).Shoot = False
                'MsgBox "reset the stuf"
                Enemies(i).explode = 0
                Enemies(i).Y = -20
                Enemies(i).X = i * 50
            Next
            Initiate = False
        Else
            For i = 1 To NumEnemies
                Enemies(i).Y = Enemies(i).Y + 2
                If Enemies(i).Y >= 30 Then
                    WaveStart = False
                End If
            Next
        End If
    Else
    For i = 1 To NumEnemies
        If Enemies(i).explode = 0 Then
            If Enemies(i).LeftRight = True Then
                Enemies(i).X = Enemies(i).X - 3
            Else
                Enemies(i).X = Enemies(i).X + 3
            End If
        End If
        If Enemies(i).X <= 0 Then
            Enemies(i).LeftRight = False
        ElseIf Enemies(i).X + 25 >= frmmain.picmain.ScaleWidth Then
            Enemies(i).LeftRight = True
        End If
    Next
    
    End If
End Sub



Public Sub EnemyWave3()  'same as 1 and 2 except with an extra layer
    Dim i As Integer
    If WaveStart = True Then
        If Initiate = True Then
            NumEnemies = 20
            ReDim Enshoot(1 To NumEnemies) As Coordinates
            ReDim Enemies(1 To NumEnemies) As GameUnits
            For i = 1 To NumEnemies
                Enemies(i).armor = 1
                Enemies(i).LeftRight = True
                Enemies(i).Destroyed = False
                Enemies(i).PointModifier = 1
                Enemies(i).Shoot = False
                'MsgBox "reset the stuf"
                Enemies(i).explode = 0
                Enemies(i).Y = -20
                
            Next
            For i = 1 To NumEnemies / 2          'puts them in a layered pattern
                Enemies(i).X = i * 50
            Next
            For i = (NumEnemies / 2) + 1 To NumEnemies
                Enemies(i).X = Enemies(i - 10).X
            Next
            Initiate = False
        Else
            For i = 1 To NumEnemies / 2
                Enemies(i).Y = Enemies(i).Y + 2
                If Enemies(i).Y >= 60 Then
                    WaveStart = False
                End If
            Next
            For i = (NumEnemies / 2) + 1 To NumEnemies
                Enemies(i).Y = Enemies(i).Y + 4
                If Enemies(i).Y >= 30 Then
                    WaveStart = False
                End If
            Next
        End If
    Else
    For i = 1 To NumEnemies
        If Enemies(i).explode = 0 Then
            If Enemies(i).LeftRight = True Then
                Enemies(i).X = Enemies(i).X - 3
            Else
                Enemies(i).X = Enemies(i).X + 3
            End If
        End If
        If Enemies(i).X <= 0 Then
            Enemies(i).LeftRight = False
        ElseIf Enemies(i).X + 25 >= frmmain.picmain.ScaleWidth Then
            Enemies(i).LeftRight = True
        End If
    Next
    
    End If
End Sub




Public Sub EnemyWave4() 'like 1, but they can shoot
    Dim i As Integer
    If WaveStart = True Then
        If Initiate = True Then
            NumEnemies = 10
            ReDim Enshoot(1 To NumEnemies) As Coordinates
            ReDim Enemies(1 To NumEnemies) As GameUnits
            For i = 1 To NumEnemies
                Enemies(i).armor = 1
                Enemies(i).PointModifier = 2 'increase point modifier
                Enemies(i).LeftRight = True
                Enemies(i).Destroyed = False
                Enemies(i).Shoot = False
                'MsgBox "reset the stuf"
                Enemies(i).explode = 0
                Enemies(i).Y = -20
                Enemies(i).X = i * 50
            Next
            Initiate = False
        Else
            For i = 1 To NumEnemies
                Enemies(i).Y = Enemies(i).Y + 2
                If Enemies(i).Y >= 30 Then
                    WaveStart = False
                End If
            Next
        End If
    Else
    For i = 1 To NumEnemies
        If Enemies(i).explode = 0 Then
            If Enemies(i).LeftRight = True Then
                Enemies(i).X = Enemies(i).X - 3
            Else
                Enemies(i).X = Enemies(i).X + 3
            End If
        End If
        If Enemies(i).X <= 0 Then
            Enemies(i).LeftRight = False
        ElseIf Enemies(i).X + 25 >= frmmain.picmain.ScaleWidth Then
            Enemies(i).LeftRight = True
        End If
    Next
    
    For i = 1 To NumEnemies
        RandomShoot = Int(Rnd * 100) + 1  'set random shoot
    
        If RandomShoot <= 1 And Enemies(i).Shoot = False And Enemies(i).Destroyed = False Then
            Enemies(i).Shoot = True
            Enshoot(i).X = Enemies(i).X + 12.5     'if they randomshoot is less than or = to 1, they fire
            Enshoot(i).Y = Enemies(i).Y
        End If
    Next
    
    End If
End Sub


Public Sub EnemyWave5() 'same as wave 4
    Dim i As Integer
    If WaveStart = True Then
        If Initiate = True Then
            NumEnemies = 10
            ReDim Enshoot(1 To NumEnemies) As Coordinates
            ReDim Enemies(1 To NumEnemies) As GameUnits
            For i = 1 To NumEnemies
                Enemies(i).armor = 1
                Enemies(i).PointModifier = 2
                Enemies(i).LeftRight = True
                Enemies(i).Destroyed = False
                Enemies(i).Shoot = False
                'MsgBox "reset the stuf"
                Enemies(i).explode = 0
                Enemies(i).Y = -20
                Enemies(i).X = i * 50
            Next
            Initiate = False
        Else
            For i = 1 To NumEnemies
                Enemies(i).Y = Enemies(i).Y + 2
                If Enemies(i).Y >= 30 Then
                    WaveStart = False
                End If
            Next
        End If
    Else
    For i = 1 To NumEnemies
        If Enemies(i).explode = 0 Then
            If Enemies(i).LeftRight = True Then
                Enemies(i).X = Enemies(i).X - 3
            Else
                Enemies(i).X = Enemies(i).X + 3
            End If
        End If
        If Enemies(i).X <= 0 Then
            Enemies(i).LeftRight = False
        ElseIf Enemies(i).X + 25 >= frmmain.picmain.ScaleWidth Then
            Enemies(i).LeftRight = True
        End If
    Next
    
    For i = 1 To NumEnemies
        RandomShoot = Int(Rnd * 100) + 1
    
        If RandomShoot <= 1 And Enemies(i).Shoot = False And Enemies(i).Destroyed = False Then
            Enemies(i).Shoot = True
            Enshoot(i).X = Enemies(i).X + 12.5
            Enshoot(i).Y = Enemies(i).Y
        End If
    Next
    
    End If
End Sub


Public Sub EnemyWave6()        'same as waves 4,5 , only double layered
    Dim i As Integer
    If WaveStart = True Then
        If Initiate = True Then
            NumEnemies = 20
            ReDim Enshoot(1 To NumEnemies) As Coordinates
            ReDim Enemies(1 To NumEnemies) As GameUnits
            For i = 1 To NumEnemies
                Enemies(i).armor = 1
                Enemies(i).LeftRight = True
                Enemies(i).Destroyed = False
                Enemies(i).PointModifier = 2
                Enemies(i).Shoot = False
                'MsgBox "reset the stuf"
                Enemies(i).explode = 0
                Enemies(i).Y = -20
                
            Next
            For i = 1 To NumEnemies / 2
                Enemies(i).X = i * 50
            Next
            For i = (NumEnemies / 2) + 1 To NumEnemies
                Enemies(i).X = Enemies(i - 10).X
            Next
            Initiate = False
        Else
            For i = 1 To NumEnemies / 2
                Enemies(i).Y = Enemies(i).Y + 2
                If Enemies(i).Y >= 60 Then
                    WaveStart = False
                End If
            Next
            For i = (NumEnemies / 2) + 1 To NumEnemies
                Enemies(i).Y = Enemies(i).Y + 4
                If Enemies(i).Y >= 30 Then
                    WaveStart = False
                End If
            Next
        End If
    Else
    For i = 1 To NumEnemies
        If Enemies(i).explode = 0 Then
            If Enemies(i).LeftRight = True Then
                Enemies(i).X = Enemies(i).X - 3
            Else
                Enemies(i).X = Enemies(i).X + 3
            End If
        End If
        If Enemies(i).X <= 0 Then
            Enemies(i).LeftRight = False
        ElseIf Enemies(i).X + 25 >= frmmain.picmain.ScaleWidth Then
            Enemies(i).LeftRight = True
        End If
    Next
    
     For i = 1 To NumEnemies
        RandomShoot = Int(Rnd * 100) + 1
    
        If RandomShoot <= 1 And Enemies(i).Shoot = False And Enemies(i).Destroyed = False Then
            Enemies(i).Shoot = True
            Enshoot(i).X = Enemies(i).X + 12.5
            Enshoot(i).Y = Enemies(i).Y
        End If
    Next
    
    End If
End Sub


Public Sub EnemyWave7()        'this wave i experimented with a different pattern (wedged)
    Dim i As Integer
    If WaveStart = True Then
        If Initiate = True Then
            NumEnemies = 10
            ReDim Enshoot(1 To NumEnemies) As Coordinates
            ReDim Enemies(1 To NumEnemies) As GameUnits
            For i = 1 To NumEnemies
                Enemies(i).armor = 1
                Enemies(i).LeftRight = True
                Enemies(i).Destroyed = False
                Enemies(i).PointModifier = 2
                Enemies(i).Shoot = False
                'MsgBox "reset the stuf"
                Enemies(i).explode = 0
                
            Next
            Enemies(1).Y = -20
            Enemies(2).Y = Enemies(1).Y - 10
            Enemies(3).Y = Enemies(2).Y - 10
            Enemies(4).Y = Enemies(3).Y - 10
            Enemies(5).Y = Enemies(4).Y - 10     'set them up
            Enemies(6).Y = Enemies(5).Y
            Enemies(7).Y = Enemies(6).Y + 10
            Enemies(8).Y = Enemies(7).Y + 10
            Enemies(9).Y = Enemies(8).Y + 10
            Enemies(10).Y = Enemies(9).Y + 10
            For i = 1 To NumEnemies
                Enemies(i).X = 50 * i
            Next
            Initiate = False
        Else
              
            For i = 1 To NumEnemies
                Enemies(i).Y = Enemies(i).Y + 4
                If Enemies(i).Y >= 50 Then
                    WaveStart = False
                End If
            Next
        End If
    Else
    For i = 1 To NumEnemies
        If Enemies(i).explode = 0 Then
            If Enemies(i).LeftRight = True Then
                Enemies(i).X = Enemies(i).X - 3
            Else
                Enemies(i).X = Enemies(i).X + 3
            End If
        End If
        If Enemies(i).X <= 0 Then
            Enemies(i).LeftRight = False
        ElseIf Enemies(i).X + 25 >= frmmain.picmain.ScaleWidth Then
            Enemies(i).LeftRight = True
        End If
    Next
    
     For i = 1 To NumEnemies
        RandomShoot = Int(Rnd * 100) + 1
    
        If RandomShoot <= 5 And Enemies(i).Shoot = False And Enemies(i).Destroyed = False Then
            Enemies(i).Shoot = True
            Enshoot(i).X = Enemies(i).X + 12.5
            Enshoot(i).Y = Enemies(i).Y
        End If
    Next
    
    End If
End Sub

Public Sub EnemyWave8() 'back to normal pattern, they move down slightly
    Dim i As Integer
    If WaveStart = True Then
        If Initiate = True Then
            NumEnemies = 10
            ReDim Enshoot(1 To NumEnemies) As Coordinates
            ReDim Enemies(1 To NumEnemies) As GameUnits
            For i = 1 To NumEnemies
                Enemies(i).armor = 1
                Enemies(i).PointModifier = 3
                Enemies(i).LeftRight = True
                Enemies(i).Destroyed = False
                Enemies(i).Shoot = False
                'MsgBox "reset the stuf"
                Enemies(i).explode = 0
                Enemies(i).Y = -20
                Enemies(i).X = i * 50
            Next
            Initiate = False
        Else
            For i = 1 To NumEnemies
                Enemies(i).Y = Enemies(i).Y + 2
                If Enemies(i).Y >= 30 Then
                    WaveStart = False
                End If
            Next
        End If
    Else
    For i = 1 To NumEnemies
        If Enemies(i).explode = 0 Then
            If Enemies(i).LeftRight = True Then
                Enemies(i).X = Enemies(i).X - 3
            Else
                Enemies(i).X = Enemies(i).X + 3
            End If
        End If
        If Enemies(i).X <= 0 Then
            Enemies(i).LeftRight = False
            Enemies(i).Y = Enemies(i).Y + 5 'move them down 5 pixels when they bump into a wall
        ElseIf Enemies(i).X + 25 >= frmmain.picmain.ScaleWidth Then
            Enemies(i).LeftRight = True
            Enemies(i).Y = Enemies(i).Y + 5
        End If
        If Enemies(i).Y + 25 >= frmmain.picmain.ScaleHeight Then
            Enemies(i).Destroyed = True 'if they reach the end, make them dissapear
        End If
    Next
    
    For i = 1 To NumEnemies
        RandomShoot = Int(Rnd * 100) + 1
    
        If RandomShoot <= 3 And Enemies(i).Shoot = False And Enemies(i).Destroyed = False Then
            Enemies(i).Shoot = True
            Enshoot(i).X = Enemies(i).X + 12.5
            Enshoot(i).Y = Enemies(i).Y
        End If
    Next
    
    End If
End Sub

Public Sub EnemyWave9() 'same as 8 only fire quicker, and move down faster
    Dim i As Integer
    If WaveStart = True Then
        If Initiate = True Then
            NumEnemies = 20
            ReDim Enshoot(1 To NumEnemies) As Coordinates
            ReDim Enemies(1 To NumEnemies) As GameUnits
            For i = 1 To NumEnemies
                Enemies(i).armor = 1
                Enemies(i).LeftRight = True
                Enemies(i).Destroyed = False
                Enemies(i).PointModifier = 2
                Enemies(i).Shoot = False
                'MsgBox "reset the stuf"
                Enemies(i).explode = 0
                Enemies(i).Y = -20
                
            Next
            For i = 1 To NumEnemies / 2
                Enemies(i).X = i * 50
            Next
            For i = (NumEnemies / 2) + 1 To NumEnemies
                Enemies(i).X = Enemies(i - 10).X
            Next
            Initiate = False
        Else
            For i = 1 To NumEnemies / 2
                Enemies(i).Y = Enemies(i).Y + 2
                If Enemies(i).Y >= 60 Then
                    WaveStart = False
                End If
            Next
            For i = (NumEnemies / 2) + 1 To NumEnemies
                Enemies(i).Y = Enemies(i).Y + 4
                If Enemies(i).Y >= 30 Then
                    WaveStart = False
                End If
            Next
        End If
    Else
    For i = 1 To NumEnemies
        If Enemies(i).explode = 0 Then
            If Enemies(i).LeftRight = True Then
                Enemies(i).X = Enemies(i).X - 3
            Else
                Enemies(i).X = Enemies(i).X + 3
            End If
        End If
        If Enemies(i).X <= 0 Then
            Enemies(i).LeftRight = False
            Enemies(i).Y = Enemies(i).Y + 10 'rate they move down
        ElseIf Enemies(i).X + 25 >= frmmain.picmain.ScaleWidth Then
            Enemies(i).LeftRight = True
            Enemies(i).Y = Enemies(i).Y + 10
        End If
        If Enemies(i).Y + 25 >= frmmain.picmain.ScaleHeight Then
            Enemies(i).Destroyed = True
        End If
    Next
    
     For i = 1 To NumEnemies
        RandomShoot = Int(Rnd * 100) + 1
    
        If RandomShoot <= 5 And Enemies(i).Shoot = False And Enemies(i).Destroyed = False Then
            Enemies(i).Shoot = True
            Enshoot(i).X = Enemies(i).X + 12.5
            Enshoot(i).Y = Enemies(i).Y
        End If
    Next
    
    End If
End Sub

Public Sub EnemyWave10() 'same as 8 and 9, double layered, move down FAST
    Dim i As Integer
    If WaveStart = True Then
        If Initiate = True Then
            NumEnemies = 20
            ReDim Enshoot(1 To NumEnemies) As Coordinates
            ReDim Enemies(1 To NumEnemies) As GameUnits
            For i = 1 To NumEnemies
                Enemies(i).armor = 2 'armor increase
                Enemies(i).LeftRight = True
                Enemies(i).Destroyed = False
                Enemies(i).PointModifier = 3
                Enemies(i).Shoot = False
                'MsgBox "reset the stuf"
                Enemies(i).explode = 0
                Enemies(i).Y = -20
                
            Next
            For i = 1 To NumEnemies / 2
                Enemies(i).X = i * 50
            Next
            For i = (NumEnemies / 2) + 1 To NumEnemies
                Enemies(i).X = Enemies(i - 10).X
            Next
            Initiate = False
        Else
            For i = 1 To NumEnemies / 2
                Enemies(i).Y = Enemies(i).Y + 2
                If Enemies(i).Y >= 60 Then
                    WaveStart = False
                End If
            Next
            For i = (NumEnemies / 2) + 1 To NumEnemies
                Enemies(i).Y = Enemies(i).Y + 4
                If Enemies(i).Y >= 30 Then
                    WaveStart = False
                End If
            Next
        End If
    Else
    For i = 1 To NumEnemies
        If Enemies(i).explode = 0 Then
            If Enemies(i).LeftRight = True Then
                Enemies(i).X = Enemies(i).X - 3
            Else
                Enemies(i).X = Enemies(i).X + 3
            End If
        End If
        If Enemies(i).X <= 0 Then
            Enemies(i).LeftRight = False
            Enemies(i).Y = Enemies(i).Y + 15
        ElseIf Enemies(i).X + 25 >= frmmain.picmain.ScaleWidth Then
            Enemies(i).LeftRight = True
            Enemies(i).Y = Enemies(i).Y + 15 'rate in which they move down
        End If
        If Enemies(i).Y + 25 >= frmmain.picmain.ScaleHeight Then
            Enemies(i).Destroyed = True
        End If
    Next
    
     For i = 1 To NumEnemies
        RandomShoot = Int(Rnd * 100) + 1
    
        If RandomShoot <= 10 And Enemies(i).Shoot = False And Enemies(i).Destroyed = False Then
            Enemies(i).Shoot = True
            Enshoot(i).X = Enemies(i).X + 12.5
            Enshoot(i).Y = Enemies(i).Y
        End If
    Next
    
    End If
End Sub

