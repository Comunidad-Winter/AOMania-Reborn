Attribute VB_Name = "SistemaTeletransporte"
Option Explicit

Private WarpMapaCiudad As Integer
Private WarpNpcTp As Integer
Private TeleportX As Integer
Private TeleportY As Integer

Sub AmuTeleport(ByVal UserIndex As String)
    Dim Ciudad As Integer

    Ciudad = RandomNumber(1, 11)

    If RestriTele(Ciudad) Then
        Call WarpUserChar(UserIndex, WarpMapaCiudad, TeleportX, TeleportY, True)

    End If

End Sub

Private Function RestriTele(Ciudad As Integer)

    Select Case Ciudad

    Case "1"
        WarpMapaCiudad = 1
        TeleportX = 50
        TeleportY = 49
        RestriTele = True
        Exit Function

    Case "2"
        WarpMapaCiudad = 34
        TeleportX = 34
        TeleportY = 83
        RestriTele = True
        Exit Function

    Case "3"
        WarpMapaCiudad = 59
        TeleportX = 50
        TeleportY = 50
        RestriTele = True
        Exit Function

    Case "4"
        WarpMapaCiudad = 132
        TeleportX = 50
        TeleportY = 49
        RestriTele = True
        Exit Function

    Case "5"
        WarpMapaCiudad = 86
        TeleportX = 50
        TeleportY = 44
        RestriTele = True
        Exit Function

    Case "6"
        WarpMapaCiudad = 84
        TeleportX = 45
        TeleportY = 51
        RestriTele = True
        Exit Function

    Case "7"
        WarpMapaCiudad = 20
        TeleportX = 57
        TeleportY = 37
        RestriTele = True
        Exit Function

    Case "8"
        WarpMapaCiudad = 62
        TeleportX = 71
        TeleportY = 41
        RestriTele = True
        Exit Function

    Case "9"
        WarpMapaCiudad = 95
        TeleportX = 50
        TeleportY = 51
        RestriTele = True
        Exit Function

    Case "10"
        WarpMapaCiudad = 149
        TeleportX = 49
        TeleportY = 52
        RestriTele = True
        Exit Function

    Case "11"
        WarpMapaCiudad = 138
        TeleportX = 22
        TeleportY = 67
        RestriTele = True
        Exit Function

    End Select

    RestriTele = False

End Function

Sub NpcTeleport(UserIndex As Integer)
    Dim Tp As Integer

    Tp = RandomNumber(1, 10)

    If RestriNpcTp(Tp) Then
        Call WarpUserChar(UserIndex, WarpNpcTp, TeleportX, TeleportY, True)

    End If

End Sub

Private Function RestriNpcTp(Tp As Integer)

    Select Case Tp

    Case 1
        WarpNpcTp = 35
        TeleportX = 50
        TeleportY = 45
        RestriNpcTp = True
        Exit Function

    Case 2
        WarpNpcTp = 5
        TeleportX = 55
        TeleportY = 51
        RestriNpcTp = True
        Exit Function

    Case 3
        WarpNpcTp = 66
        TeleportX = 71
        TeleportY = 79
        RestriNpcTp = True
        Exit Function

    Case 4
        WarpNpcTp = 61
        TeleportX = 50
        TeleportY = 69
        RestriNpcTp = True
        Exit Function

    Case 5
        WarpNpcTp = 131
        TeleportX = 33
        TeleportY = 22
        RestriNpcTp = True
        Exit Function

    Case 6
        WarpNpcTp = 62
        TeleportX = 87
        TeleportY = 47
        RestriNpcTp = True
        Exit Function

    Case 7
        WarpNpcTp = 92
        TeleportX = 89
        TeleportY = 40
        RestriNpcTp = True
        Exit Function

    Case 8
        WarpNpcTp = 128
        TeleportX = 51
        TeleportY = 47
        RestriNpcTp = True
        Exit Function

    Case 9
        WarpNpcTp = 108
        TeleportX = 50
        TeleportY = 50
        RestriNpcTp = True
        Exit Function

    Case 10
        WarpNpcTp = 115
        TeleportX = 50
        TeleportY = 50
        RestriNpcTp = True
        Exit Function

    Case 11
        WarpNpcTp = 172
        TeleportX = 31
        TeleportY = 50
        RestriNpcTp = True
        Exit Function

    Case 12
        WarpNpcTp = 172
        TeleportX = 31
        TeleportY = 50
        RestriNpcTp = True
        Exit Function

    Case 13
        WarpNpcTp = 8
        TeleportX = 50
        TeleportY = 50
        RestriNpcTp = True
        Exit Function

    Case 14
        WarpNpcTp = 129
        TeleportX = 50
        TeleportY = 50
        RestriNpcTp = True
        Exit Function

    Case 15
        WarpNpcTp = 172
        TeleportX = 31
        TeleportY = 50
        RestriNpcTp = True
        Exit Function

    Case 16
        WarpNpcTp = 3
        TeleportX = 53
        TeleportY = 50
        RestriNpcTp = True
        Exit Function

    Case 17
        WarpNpcTp = 16
        TeleportX = 50
        TeleportY = 50
        RestriNpcTp = True
        Exit Function

    Case 18
        WarpNpcTp = 15
        TeleportX = 50
        TeleportY = 50
        RestriNpcTp = True
        Exit Function

    Case 11
        WarpNpcTp = 172
        TeleportX = 31
        TeleportY = 50
        RestriNpcTp = True
        Exit Function

    Case 11
        WarpNpcTp = 172
        TeleportX = 31
        TeleportY = 50
        RestriNpcTp = True
        Exit Function

    End Select

    RestriNpcTp = False

End Function
