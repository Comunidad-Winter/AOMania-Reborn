Attribute VB_Name = "ModRanking"

Public Sub EnviaRank(ByVal UserIndex As Integer)

    SendData SendTarget.ToIndex, UserIndex, 0, "BINMODEPT" & Ranking.MaxOro.UserName & "," & Ranking.MaxOro.value & "," & _
            Ranking.MaxTrofeos.UserName & "," & Ranking.MaxTrofeos.value & "," & Ranking.MaxUsuariosMatados.UserName & "," & _
            Ranking.MaxUsuariosMatados.value & "," & Ranking.MaxTorneos.UserName & "," & Ranking.MaxTorneos.value & "," & _
            Ranking.MaxDeaths.UserName & "," & Ranking.MaxDeaths.value & "," & Ranking.MaxRetos.UserName & "," & Ranking.MaxRetos.value & "," & _
            Ranking.MaxDuelos.UserName & "," & Ranking.MaxDuelos.value & "," & Ranking.MaxPlantes.UserName & "," & Ranking.MaxPlantes.value

End Sub

Public Sub EnviaPuntos(ByVal UserIndex As Integer)

    SendData SendTarget.ToIndex, UserIndex, 0, "WETA" & UserList(UserIndex).Stats.PuntosDeath & "," & UserList(UserIndex).Stats.PuntosDuelos & "," _
            & UserList(UserIndex).Stats.PuntosPlante & "," & UserList(UserIndex).Stats.PuntosRetos & "," & UserList(UserIndex).Stats.PuntosTorneo

End Sub

Public Sub CompruebaOro(ByVal UserIndex As Integer)

    ' actualiza el ranking de oro si el usuario tiene mas oro que el mayor del ranking
    If UserList(UserIndex).Stats.GLD > Ranking.MaxOro.value And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
        Ranking.MaxOro.value = UserList(UserIndex).Stats.GLD
        Ranking.MaxOro.UserName = UserList(UserIndex).Name

    End If

End Sub

Public Sub CompruebaTrofeos(ByVal UserIndex As Integer)

    ' actualiza el ranking de trofeos si el usuario tiene mas trofeos que el mayor del ranking
    If UserList(UserIndex).Stats.TrofOro > Ranking.MaxTrofeos.value And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
        Ranking.MaxTrofeos.value = UserList(UserIndex).Stats.TrofOro
        Ranking.MaxTrofeos.UserName = UserList(UserIndex).Name

    End If

End Sub

Public Sub CompruebaUserDies(ByVal UserIndex As Integer)

    ' actualiza el ranking de muertes si el usuario tiene mas muertes que el mayor del ranking
    If UserList(UserIndex).Stats.UsuariosMatados > Ranking.MaxUsuariosMatados.value And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
        Ranking.MaxUsuariosMatados.value = UserList(UserIndex).Stats.UsuariosMatados
        Ranking.MaxUsuariosMatados.UserName = UserList(UserIndex).Name

    End If

End Sub

Public Sub CompruebaDuelos(ByVal UserIndex As Integer)

    ' actualiza el ranking de duelos si el usuario tiene mas duelos que el mayor del ranking
    If UserList(UserIndex).Stats.PuntosDuelos > Ranking.MaxDuelos.value And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
        Ranking.MaxDuelos.value = UserList(UserIndex).Stats.PuntosDuelos
        Ranking.MaxDuelos.UserName = UserList(UserIndex).Name

    End If

End Sub

Public Sub CompruebaRetos(ByVal UserIndex As Integer)

    ' actualiza el ranking de duelos si el usuario tiene mas duelos que el mayor del ranking
    If UserList(UserIndex).Stats.PuntosRetos > Ranking.MaxRetos.value And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
        Ranking.MaxRetos.value = UserList(UserIndex).Stats.PuntosRetos
        Ranking.MaxRetos.UserName = UserList(UserIndex).Name

    End If

End Sub

Public Sub CompruebaPlantes(ByVal UserIndex As Integer)

    ' actualiza el ranking de plantes si el usuario tiene mas plantes que el mayor del ranking
    If UserList(UserIndex).Stats.PuntosPlante > Ranking.MaxPlantes.value And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
        Ranking.MaxPlantes.value = UserList(UserIndex).Stats.PuntosPlante
        Ranking.MaxPlantes.UserName = UserList(UserIndex).Name

    End If

End Sub

Public Sub CompruebaTorneos(ByVal UserIndex As Integer)

    ' actualiza el ranking de torneos si el usuario tiene mas torneos que el mayor del ranking
    If UserList(UserIndex).Stats.PuntosTorneo > Ranking.MaxTorneos.value And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
        Ranking.MaxTorneos.value = UserList(UserIndex).Stats.PuntosTorneo
        Ranking.MaxTorneos.UserName = UserList(UserIndex).Name

    End If

End Sub

Public Sub CompruebaDeaths(ByVal UserIndex As Integer)

    ' actualiza el ranking de deaths si el usuario tiene mas deaths que el mayor del ranking
    If UserList(UserIndex).Stats.PuntosDeath > Ranking.MaxDeaths.value And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
        Ranking.MaxDeaths.value = UserList(UserIndex).Stats.PuntosDeath
        Ranking.MaxDeaths.UserName = UserList(UserIndex).Name

    End If

End Sub
