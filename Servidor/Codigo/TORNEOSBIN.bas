Attribute VB_Name = "TORNEOSBIN"
Option Explicit
Public Torneo_Activo As Boolean
Public Torneo_Esperando As Boolean
Private Torneo_Rondas As Integer
Private Torneo_Luchadores() As Integer

Private Const mapatorneo As Integer = 160
' esquinas superior isquierda del ring
Private Const esquina1x As Integer = 42
Private Const esquina1y As Integer = 51
' esquina inferior derecha del ring
Private Const esquina2x As Integer = 59
Private Const esquina2y As Integer = 51
' Donde esperan los hijos de puta
Private Const esperax As Integer = 27
Private Const esperay As Integer = 32
' Mapa desconecta
Private Const mapa_fuera As Integer = 34
Private Const fueraesperay As Integer = 57
Private Const fueraesperax As Integer = 21
Private Const x1 As Integer = 36
Private Const x2 As Integer = 65
Private Const Y1 As Integer = 24
Private Const Y2 As Integer = 30

Private Const MIngreso1 As Byte = 34
Private Const MIngreso2 As Byte = 84

Public RondaTorneo As Byte
Private Const PremioTorneo As Integer = 10000

Function PermiteIngresoTorneo(ByVal UserIndex As Integer) As Boolean

    With UserList(UserIndex)

        Select Case .pos.Map

        Case MIngreso1
            PermiteIngresoTorneo = True
            Exit Function

        Case MIngreso2
            PermiteIngresoTorneo = True
            Exit Function

        End Select

    End With

End Function

Sub CommandParticipar(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        '  If .flags.Privilegios > PlayerType.User Then Exit Sub

        If Not PermiteIngresoTorneo(UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "||Tienes que estar en " & Trim$(UCase$(MapInfo(MIngreso1).Name)) & " ó en " & Trim$(UCase$(MapInfo(MIngreso2).Name)) & " para participar." & FONTTYPE_INFO)
            Exit Sub
        End If

        If .flags.EstaDueleando1 = True Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes ir a torneos estando plantes!." & FONTTYPE_WARNING)
            Exit Sub

        End If

        If .flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas muerto!!!" & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserIndex = Team.Pj1 Or UserIndex = Team.Pj2 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
            Exit Sub

        End If

        If .Stats.ELV < lvlTorneo Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes ser lvl " & lvlTorneo & " o mas para entrar al torneo!" & FONTTYPE_INFO)
            Exit Sub
        End If

        Call Torneos_Entra(UserIndex)

    End With

End Sub

Sub Torneos_Entra(ByVal UserIndex As Integer)

    On Error GoTo errorh

    Dim i As Integer

    If (Not Torneo_Activo) Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ningun torneo!." & FONTTYPE_INFO)
        Exit Sub

    End If

    If (Not Torneo_Esperando) Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El torneo ya ha empezado, te quedaste fuera!." & FONTTYPE_INFO)
        Exit Sub
    End If

    For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)

        If (Torneo_Luchadores(i) = UserIndex) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya estas dentro!" & FONTTYPE_WARNING)
            Exit Sub

        End If

    Next i

    For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)

        If (Torneo_Luchadores(i) = -1) Then

            If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto Then
                UserList(UserIndex).Counters.Invisibilidad = IntervaloInvisible
                UserList(UserIndex).flags.Invisible = 0
                UserList(UserIndex).Counters.Ocultando = 0
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "INVI0")
            End If

            Torneo_Luchadores(i) = UserIndex
            Dim NuevaPos As WorldPos
            Dim FuturePos As WorldPos
            FuturePos.Map = mapatorneo
            FuturePos.X = esperax
            FuturePos.Y = esperay
            Call ClosestLegalPos(FuturePos, NuevaPos)

            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
            UserList(Torneo_Luchadores(i)).flags.automatico = True

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas dentro del torneo!" & FONTTYPE_INFO)

            '  Call SendData(SendTarget.toall, 0, 0, "||Torneo: Entra el participante " & UserList(userindex).name & FONTTYPE_INFO)
            If (i = UBound(Torneo_Luchadores)) Then
                xao = 100
                Call SendData(SendTarget.ToAll, 0, 0, "||Torneo: Empieza el torneo!" & FONTTYPE_GUILD)
                Torneo_Esperando = False
                Call Rondas_Combate(1)

            End If

            Exit Sub

        End If

    Next i

errorh:

End Sub

Sub Torneoauto_Cancela()

    On Error GoTo errorh:

    If (Not Torneo_Activo And Not Torneo_Esperando) Then Exit Sub
    Torneo_Activo = False
    Torneo_Esperando = False
    Call SendData(SendTarget.ToAll, 0, 0, "||Torneo: Torneo Automatico cancelado por falta de participantes." & FONTTYPE_GUILD)
    Dim i As Integer

    For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)

        If (Torneo_Luchadores(i) <> -1) Then
            Dim NuevaPos As WorldPos
            Dim FuturePos As WorldPos
            FuturePos.Map = mapa_fuera
            FuturePos.X = fueraesperax
            FuturePos.Y = fueraesperay
            Call ClosestLegalPos(FuturePos, NuevaPos)

            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
            UserList(Torneo_Luchadores(i)).flags.automatico = False

        End If

    Next i

errorh:

End Sub

Sub Rondas_Cancela()

    On Error GoTo errorh

    If (Not Torneo_Activo And Not Torneo_Esperando) Then Exit Sub
    Torneo_Activo = False
    Torneo_Esperando = False
    xao = 0
    Call SendData(SendTarget.ToAll, 0, 0, "||Torneo: Torneo Automatico cancelado por Game Master" & FONTTYPE_GUILD)
    Dim i As Integer

    For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)

        If (Torneo_Luchadores(i) <> -1) Then
            Dim NuevaPos As WorldPos
            Dim FuturePos As WorldPos
            FuturePos.Map = mapa_fuera
            FuturePos.X = fueraesperax
            FuturePos.Y = fueraesperay
            Call ClosestLegalPos(FuturePos, NuevaPos)

            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
            UserList(Torneo_Luchadores(i)).flags.automatico = False

        End If

    Next i

errorh:

End Sub

Sub Rondas_UsuarioMuere(ByVal UserIndex As Integer, Optional Real As Boolean = True, Optional CambioMapa As Boolean = False)

    On Error GoTo rondas_usuariomuere_errorh

    Dim i As Integer, pos As Integer, j As Integer
    Dim combate As Integer, LI1 As Integer, LI2 As Integer
    Dim UI1 As Integer, UI2 As Integer
    Dim b As Byte

    If (Not Torneo_Activo) Then
        Exit Sub
    ElseIf (Torneo_Activo And Torneo_Esperando) Then

        For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)

            If (Torneo_Luchadores(i) = UserIndex) Then
                Torneo_Luchadores(i) = -1
                Call WarpUserChar(UserIndex, mapa_fuera, fueraesperay, fueraesperax, True)
                UserList(UserIndex).flags.automatico = False
                Exit Sub

            End If

        Next i

        Exit Sub

    End If

    For pos = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)

        If (Torneo_Luchadores(pos) = UserIndex) Then Exit For
    Next pos

    ' si no lo ha encontrado
    If (Torneo_Luchadores(pos) <> UserIndex) Then Exit Sub

    ' comprueba si esta esperando peliar, para asi no enviar otra pelea
    If UserList(UserIndex).pos.X >= x1 And UserList(UserIndex).pos.X <= x2 And UserList(UserIndex).pos.Y >= Y1 And UserList(UserIndex).pos.Y <= Y2 _
       Then
        Call SendData(SendTarget.ToAll, 0, 0, "||Torneo: " & UserList(UserIndex).Name & " se fue del torneo mientras esperaba pelear.!" & _
                                              FONTTYPE_TALK)
        Call WarpUserChar(UserIndex, mapa_fuera, fueraesperax, fueraesperay, True)
        UserList(UserIndex).flags.automatico = False
        Torneo_Luchadores(pos) = -1
        Exit Sub

    End If

    combate = 1 + (pos - 1) \ 2

    'ponemos li1 y li2 (luchador index) de los que combatian
    LI1 = 2 * (combate - 1) + 1
    LI2 = LI1 + 1

    'se informa a la gente
    If (Real) Then
        Call SendData(SendTarget.ToAll, 0, 0, "||Torneo: " & UserList(UserIndex).Name & " pierde el combate!" & FONTTYPE_TALK)
    Else
        Call SendData(SendTarget.ToAll, 0, 0, "||Torneo: " & UserList(UserIndex).Name & " se fue del combate!" & FONTTYPE_TALK)

    End If

    'se le teleporta fuera si murio
    If (Real) Then
        Call WarpUserChar(UserIndex, mapa_fuera, fueraesperax, fueraesperay, True)
        UserList(UserIndex).flags.automatico = False
    ElseIf (Not CambioMapa) Then
        'haz la mierda para ke se le guarde otro sitio en la ficha
        Call WarpUserChar(UserIndex, mapa_fuera, fueraesperax, fueraesperay, True)
        UserList(UserIndex).flags.automatico = False

    End If

    'se le borra de la lista y se mueve el segundo a li1
    If (Torneo_Luchadores(LI1) = UserIndex) Then
        Torneo_Luchadores(LI1) = Torneo_Luchadores(LI2)    'cambiamos slot
        Torneo_Luchadores(LI2) = -1
    Else
        Torneo_Luchadores(LI2) = -1

    End If

    'si es la ultima ronda
    If (Torneo_Rondas = 1) Then
        If RondaTorneo < "4" Then
            b = 2
        Else
            b = 3
        End If
        Call WarpUserChar(Torneo_Luchadores(LI1), mapa_fuera, fueraesperax, fueraesperay, True)
        Call SendData(SendTarget.ToAll, 0, 0, "||GANADOR DEL TORNEO: " & UserList(Torneo_Luchadores(LI1)).Name & FONTTYPE_GUILD)
        Call SendData(SendTarget.ToAll, 0, 0, "||PREMIO: " & Left$(val(PremioTorneo * val(2 ^ RondaTorneo)), b) & "k de oro." & FONTTYPE_GUILD)
        UserList(Torneo_Luchadores(LI1)).Stats.GLD = UserList(Torneo_Luchadores(LI1)).Stats.GLD + val(PremioTorneo * val(2 ^ RondaTorneo))
        UserList(Torneo_Luchadores(LI1)).Stats.PuntosTorneo = UserList(Torneo_Luchadores(LI1)).Stats.PuntosTorneo + 1
        UserList(Torneo_Luchadores(LI1)).flags.automatico = False
        Call SendUserStatsBox(Torneo_Luchadores(LI1))
        Torneo_Activo = False
        xao = 0
        Exit Sub
    Else
        'a su compañero se le teleporta dentro, condicional por seguridad
        Call WarpUserChar(Torneo_Luchadores(LI1), mapatorneo, esperax, esperay, True)

    End If

    'si es el ultimo combate de la ronda
    If (2 ^ Torneo_Rondas = 2 * combate) Then

        Call SendData(SendTarget.ToAll, 0, 0, "||Torneo: Siguiente ronda!" & FONTTYPE_GUILD)
        Torneo_Rondas = Torneo_Rondas - 1

        'antes de llamar a la proxima ronda hay q copiar a los putos xD
        For i = 1 To 2 ^ Torneo_Rondas
            UI1 = Torneo_Luchadores(2 * (i - 1) + 1)
            UI2 = Torneo_Luchadores(2 * i)

            If (UI1 = -1) Then UI1 = UI2
            Torneo_Luchadores(i) = UI1
        Next i

        ReDim Preserve Torneo_Luchadores(1 To 2 ^ Torneo_Rondas) As Integer
        Call Rondas_Combate(1)
        Exit Sub

    End If

    'vamos al siguiente combate
    Call Rondas_Combate(combate + 1)
rondas_usuariomuere_errorh:

End Sub

Sub Rondas_UsuarioDesconecta(ByVal UserIndex As Integer)

    On Error GoTo errorh

    Call SendData(SendTarget.ToAll, 0, 0, "||Torneo: " & UserList(UserIndex).Name & _
                                        " Ha desconectado en Torneo Automatico!!" & FONTTYPE_TALK)

    Call SendUserStatsBox(UserIndex)
    Call Rondas_UsuarioMuere(UserIndex, False, False)
errorh:

End Sub

Sub Rondas_UsuarioCambiamapa(ByVal UserIndex As Integer)

    On Error GoTo errorh

    Call Rondas_UsuarioMuere(UserIndex, False, True)
errorh:

End Sub

Sub torneos_auto(ByVal rondas As Integer)

    On Error GoTo errorh

    If (Torneo_Activo) Then

        Exit Sub

    End If

    'Call SendData(SendTarget.ToAll, 0, 0, "||Torneo: Esta empezando un nuevo torneo 1v1 de " & val(2 ^ rondas) & _
     " participantes!! para participar pon /PARTICIPAR - (No cae inventario)" & FONTTYPE_GUILD)
    Call SendData(SendTarget.ToAll, 0, 0, "TW48")

    Torneo_Rondas = rondas
    Torneo_Activo = True
    Torneo_Esperando = True

    ReDim Torneo_Luchadores(1 To 2 ^ rondas) As Integer
    Dim i As Integer

    For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
        Torneo_Luchadores(i) = -1
    Next i

errorh:

End Sub

Sub Torneos_Inicia(ByVal UserIndex As Integer, ByVal rondas As Integer)

    On Error GoTo errorh

    If (Torneo_Activo) Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya hay un torneo!." & FONTTYPE_INFO)
        Exit Sub

    End If

    Call SendData(SendTarget.ToAll, 0, 0, "||Torneo: Esta empezando un nuevo torneo 1v1 de " & val(2 ^ rondas) & _
                                        " participantes!! para participar pon /PARTICIPAR - (No cae inventario)" & FONTTYPE_GUILD)
    Call SendData(SendTarget.ToAll, 0, 0, "TW48")

    Torneo_Rondas = rondas
    Torneo_Activo = True
    Torneo_Esperando = True

    ReDim Torneo_Luchadores(1 To 2 ^ rondas) As Integer
    Dim i As Integer

    For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
        Torneo_Luchadores(i) = -1
    Next i

errorh:

End Sub

Sub Rondas_Combate(combate As Integer)

    On Error GoTo errorh

    Dim UI1 As Integer, UI2 As Integer
    UI1 = Torneo_Luchadores(2 * (combate - 1) + 1)
    UI2 = Torneo_Luchadores(2 * combate)

    If (UI2 = -1) Then
        UI2 = Torneo_Luchadores(2 * (combate - 1) + 1)
        UI1 = Torneo_Luchadores(2 * combate)

    End If

    If (UI1 = -1) Then
        Call SendData(SendTarget.ToAll, 0, 0, "||Torneo: Combate anulado porque un participante involucrado se desconecto" & FONTTYPE_TALK)

        If (Torneo_Rondas = 1) Then
            If (UI2 <> -1) Then
                Call SendData(SendTarget.ToAll, 0, 0, "||Torneo: Torneo terminado. Ganador del torneo por eliminacion: " & UserList(UI2).Name & _
                                                      FONTTYPE_GUILD)
                UserList(UI2).flags.automatico = False
                ' dale_recompensa()
                Torneo_Activo = False
                Exit Sub

            End If

            Call SendData(SendTarget.ToAll, 0, 0, "||Torneo: Torneo terminado. No hay ganador porque todos se fueron :(" & FONTTYPE_GUILD)
            Exit Sub

        End If

        If (UI2 <> -1) Then Call SendData(SendTarget.ToAll, 0, 0, "||Torneo: " & UserList(UI2).Name & " pasa a la siguiente ronda!" & FONTTYPE_TALK)

        If (2 ^ Torneo_Rondas = 2 * combate) Then
            Call SendData(SendTarget.ToAll, 0, 0, "||Torneo: Siguiente ronda!" & FONTTYPE_GUILD)
            Torneo_Rondas = Torneo_Rondas - 1
            'antes de llamar a la proxima ronda hay q copiar a los putos xD
            Dim i As Integer, j As Integer

            For i = 1 To 2 ^ Torneo_Rondas
                UI1 = Torneo_Luchadores(2 * (i - 1) + 1)
                UI2 = Torneo_Luchadores(2 * i)

                If (UI1 = -1) Then UI1 = UI2
                Torneo_Luchadores(i) = UI1
            Next i

            ReDim Preserve Torneo_Luchadores(1 To 2 ^ Torneo_Rondas) As Integer
            Call Rondas_Combate(1)
            Exit Sub

        End If

        Call Rondas_Combate(combate + 1)
        Exit Sub

    End If

    Call SendData(SendTarget.ToAll, 0, 0, "||Torneo: " & UserList(UI1).Name & " versus " & UserList(UI2).Name & ". Esquinas!! Peleen!" & _
                                          FONTTYPE_GUILD)

    Call WarpUserChar(UI1, mapatorneo, esquina1x, esquina1y, True)
    Call WarpUserChar(UI2, mapatorneo, esquina2x, esquina2y, True)
errorh:

End Sub

