Attribute VB_Name = "MdosVdos"

Public Type Teamduel

    Activado As Boolean
    EnCurso As Boolean
    SonDos As Boolean
    Pj1 As Integer
    Pj2 As Integer
    pj3 As Integer
    pj4 As Integer

End Type

Public Team As Teamduel

Public Sub VerificarDosVDos(ByVal UserIndex As Integer)

    On Error GoTo errorh:

    UserList(UserIndex).flags.ParejaMuerta = True

    If UserList(Team.Pj1).flags.ParejaMuerta = True And UserList(Team.Pj2).flags.ParejaMuerta = True Then
        UserList(Team.pj3).Stats.GLD = UserList(Team.pj3).Stats.GLD + entrarReto2v2
        UserList(Team.pj4).Stats.GLD = UserList(Team.pj4).Stats.GLD + entrarReto2v2
        UserList(Team.pj4).Stats.PuntosRetos = UserList(Team.pj4).Stats.PuntosRetos + 1
        UserList(Team.pj3).Stats.PuntosRetos = UserList(Team.pj3).Stats.PuntosRetos + 1

        If UserList(Team.Pj2).Stats.GLD >= entrarReto2v2 Then
            UserList(Team.Pj2).Stats.GLD = UserList(Team.Pj2).Stats.GLD - entrarReto2v2

        End If

        If UserList(Team.Pj1).Stats.GLD >= entrarReto2v2 Then
            UserList(Team.Pj1).Stats.GLD = UserList(Team.Pj1).Stats.GLD - entrarReto2v2

        End If

        Call SendUserStatsBox(Team.pj3)
        Call SendUserStatsBox(Team.pj4)
        Call SendUserStatsBox(Team.Pj2)
        Call SendUserStatsBox(Team.Pj1)
        Call SendData(ToAll, UserIndex, 0, "||2Vs2: " & UserList(Team.Pj1).Name & " y " & UserList(Team.Pj2).Name & " han perdido contra " & _
                                           UserList(Team.pj3).Name & " y " & UserList(Team.pj4).Name & FONTTYPE_RETOS2V2)
        Call TerminoDosVDos

    ElseIf UserList(Team.pj3).flags.ParejaMuerta = True And UserList(Team.pj4).flags.ParejaMuerta = True Then

        If UserList(Team.pj3).Stats.GLD >= entrarReto2v2 Then
            UserList(Team.pj3).Stats.GLD = UserList(Team.pj3).Stats.GLD - entrarReto2v2

        End If

        If UserList(Team.pj4).Stats.GLD >= entrarReto2v2 Then
            UserList(Team.pj4).Stats.GLD = UserList(Team.pj4).Stats.GLD - entrarReto2v2

        End If

        UserList(Team.Pj2).Stats.GLD = UserList(Team.Pj2).Stats.GLD + entrarReto2v2
        UserList(Team.Pj1).Stats.GLD = UserList(Team.Pj1).Stats.GLD + entrarReto2v2
        UserList(Team.Pj2).Stats.PuntosRetos = UserList(Team.Pj2).Stats.PuntosRetos + 1
        UserList(Team.Pj1).Stats.PuntosRetos = UserList(Team.Pj1).Stats.PuntosRetos + 1

        Call SendUserStatsBox(Team.pj3)
        Call SendUserStatsBox(Team.pj4)
        Call SendUserStatsBox(Team.Pj2)
        Call SendUserStatsBox(Team.Pj1)
        Call SendData(ToAll, UserIndex, 0, "||2Vs2: " & UserList(Team.pj3).Name & " y " & UserList(Team.pj4).Name & " han perdido contra " & _
                                           UserList(Team.Pj1).Name & " y " & UserList(Team.Pj2).Name & FONTTYPE_RETOS2V2)

        Call TerminoDosVDos

    End If

errorh:

End Sub

Public Sub TerminoDosVDos()

    On Error GoTo errorh:

    UserList(Team.Pj1).flags.EnDosVDos = False
    UserList(Team.Pj1).flags.envioSol = False
    UserList(Team.Pj1).flags.RecibioSol = False
    UserList(Team.Pj2).flags.EnDosVDos = False
    UserList(Team.Pj2).flags.envioSol = False
    UserList(Team.Pj2).flags.RecibioSol = False
    UserList(Team.pj3).flags.EnDosVDos = False
    UserList(Team.pj3).flags.envioSol = False
    UserList(Team.pj3).flags.RecibioSol = False
    UserList(Team.pj4).flags.EnDosVDos = False
    UserList(Team.pj4).flags.envioSol = False
    UserList(Team.pj4).flags.RecibioSol = False
    UserList(Team.Pj1).flags.ParejaMuerta = False
    UserList(Team.Pj2).flags.ParejaMuerta = False
    UserList(Team.pj3).flags.ParejaMuerta = False
    UserList(Team.pj4).flags.ParejaMuerta = False
    Dim NuevaPos As WorldPos
    Dim FuturePos As WorldPos
    FuturePos.Map = 1
    FuturePos.X = 50
    FuturePos.Y = 50
    Call ClosestLegalPos(FuturePos, NuevaPos)

    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Team.Pj1, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
    Call ClosestLegalPos(FuturePos, NuevaPos)

    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Team.Pj2, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
    Call ClosestLegalPos(FuturePos, NuevaPos)

    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Team.pj3, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
    Call ClosestLegalPos(FuturePos, NuevaPos)

    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Team.pj4, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
    Team.EnCurso = False
    Team.SonDos = False
    Team.Pj1 = 0
    Team.Pj2 = 0
    Team.pj3 = 0
    Team.pj4 = 0
errorh:

End Sub

Sub CerroEnDuelo(ByVal UserIndex As Integer)

    On Error GoTo errorh

    Call TerminoDosVDos

    Call SendData(ToAll, 0, 0, "||2Vs2: El reto se cancela porque " & UserList(UserIndex).Name & " desconectó. Se le penaliza con 2kk de oro." & _
                               FONTTYPE_RETOS2V2)

    If UserList(UserIndex).Stats.GLD >= 2000000 Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 2000000
        SendUserStatsBox (UserIndex)

    End If

errorh:

End Sub
