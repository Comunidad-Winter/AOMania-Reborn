Attribute VB_Name = "RETOS"

Public Sub ComensarDuelo(ByVal UserIndex As Integer, ByVal TIndex As Integer)

    UserList(UserIndex).flags.EstaDueleando = True
    UserList(UserIndex).flags.Oponente = TIndex
    UserList(TIndex).flags.EstaDueleando = True
    Call WarpUserChar(TIndex, 78, 41, 50)
    UserList(TIndex).flags.Oponente = UserIndex
    Call WarpUserChar(UserIndex, 78, 60, 50)
    Call SendData(toall, 0, 0, "||Retos: " & UserList(TIndex).Name & " y " & UserList(UserIndex).Name & " van a competir en un Reto." & _
            FONTTYPE_RETOS)
    Retos1 = UserList(TIndex).Name
    Retos2 = UserList(UserIndex).Name

End Sub

Public Sub ResetDuelo(ByVal UserIndex As Integer, ByVal TIndex As Integer)

    On Error GoTo errrorxaoo

    UserList(UserIndex).flags.EsperandoDuelo = False
    UserList(UserIndex).flags.Oponente = 0
    UserList(UserIndex).flags.EstaDueleando = False
    Dim NuevaPos  As WorldPos
    Dim FuturePos As WorldPos
    FuturePos.Map = 1
    FuturePos.X = 50
    FuturePos.Y = 50
    Call ClosestLegalPos(FuturePos, NuevaPos)

    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)

    Call ClosestLegalPos(FuturePos, NuevaPos)

    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(TIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
    UserList(TIndex).flags.EsperandoDuelo = False
    UserList(TIndex).flags.Oponente = 0
    UserList(TIndex).flags.EstaDueleando = False
errrorxaoo:

End Sub

Public Sub TerminarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)

    On Error GoTo errorxao

    Call SendData(toall, Ganador, 0, "||Retos: " & UserList(Ganador).Name & " venció a " & UserList(Perdedor).Name & " en un reto." & FONTTYPE_RETOS)

    If UserList(Perdedor).Stats.GLD >= entrarReto Then
        UserList(Perdedor).Stats.GLD = UserList(Perdedor).Stats.GLD - entrarReto

    End If

    UserList(Ganador).Stats.GLD = UserList(Ganador).Stats.GLD + entrarReto
    UserList(Ganador).Stats.PuntosRetos = UserList(Ganador).Stats.PuntosRetos + 1
    
    Call SendUserStatsBox(Perdedor)
    Call SendUserStatsBox(Ganador)
    Call ResetDuelo(Ganador, Perdedor)
errorxao:

End Sub

Public Sub DesconectarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)

    On Error GoTo errorxaoo

    Call SendData(toall, Ganador, 0, "||Retos: El reto ha sido cancelado por la desconexión de " & UserList(Perdedor).Name & "." & FONTTYPE_RETOS)
    Call ResetDuelo(Ganador, Perdedor)
errorxaoo:

End Sub
