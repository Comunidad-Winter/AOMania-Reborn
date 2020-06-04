Attribute VB_Name = "RETOSPLANTE"

Public Sub ComensarDueloPlantes(ByVal UserIndex As Integer, ByVal TIndex As Integer)

    YaHayPlante = True
    UserList(UserIndex).flags.EstaDueleando1 = True
    UserList(UserIndex).flags.Oponente1 = TIndex
    UserList(TIndex).flags.EstaDueleando1 = True
    Call WarpUserChar(TIndex, 1, 75, 33)
    UserList(TIndex).flags.Oponente1 = UserIndex
    Call WarpUserChar(UserIndex, 1, 76, 33)
    Call SendData(toall, 0, 0, "||Plantes: " & UserList(TIndex).Name & " y " & UserList(UserIndex).Name & " van a competir en un Reto de plantes." _
            & FONTTYPE_PLANTE)
    Plante1 = UserList(TIndex).Name
    Plante2 = UserList(UserIndex).Name

End Sub

Public Sub ResetDueloPlantes(ByVal UserIndex As Integer, ByVal TIndex As Integer)

    On Error GoTo errrorxaoo

    UserList(UserIndex).flags.EsperandoDuelo1 = False
    UserList(UserIndex).flags.Oponente1 = 0
    UserList(UserIndex).flags.EstaDueleando1 = False
    Dim NuevaPos  As WorldPos
    Dim FuturePos As WorldPos
    FuturePos.Map = 1
    FuturePos.X = 50
    FuturePos.Y = 50
    Call ClosestLegalPos(FuturePos, NuevaPos)

    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
    Call ClosestLegalPos(FuturePos, NuevaPos)

    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(TIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
    UserList(TIndex).flags.EsperandoDuelo1 = False
    UserList(TIndex).flags.Oponente1 = 0
    UserList(TIndex).flags.EstaDueleando1 = False
    YaHayPlante = False
errrorxaoo:
    YaHayPlante = False

End Sub

Public Sub TerminarDueloPlantes(ByVal Ganador As Integer, ByVal Perdedor As Integer)

    On Error GoTo errorxao

    Call SendData(toall, Ganador, 0, "||Plantes: " & UserList(Ganador).Name & " venció a " & UserList(Perdedor).Name & " en un reto de plantes." & _
            FONTTYPE_PLANTE)

    If UserList(Perdedor).Stats.GLD >= entrarPlante Then
        UserList(Perdedor).Stats.GLD = UserList(Perdedor).Stats.GLD - entrarPlante

    End If

    UserList(Ganador).Stats.GLD = UserList(Ganador).Stats.GLD + entrarPlante
    UserList(Ganador).Stats.PuntosPlante = UserList(Ganador).Stats.PuntosPlante + 1
    
    Call SendUserStatsBox(Perdedor)
    Call SendUserStatsBox(Ganador)
    Call ResetDueloPlantes(Ganador, Perdedor)
    YaHayPlante = False
errorxao:
    YaHayPlante = False

End Sub

Public Sub DesconectarDueloPlantes(ByVal Ganador As Integer, ByVal Perdedor As Integer)

    On Error GoTo errorxaoo

    Call SendData(toall, Ganador, 0, "||Plantes: El reto ha sido cancelado por la desconexión de " & UserList(Perdedor).Name & "." & FONTTYPE_PLANTE)
    Call ResetDueloPlantes(Ganador, Perdedor)
    YaHayPlante = False
errorxaoo:

End Sub
