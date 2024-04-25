Attribute VB_Name = "modAsedio"
'Recordar cambiar posicion muralla corX y corY
'Mapear el mapa castillo


Option Explicit
Public Const ItemMuralla As Integer = 2063
Public Const ReyNPC As Integer = 725
Public Const MurallaNPC As Integer = 726

Private Const Muralla_Max As Integer = 17356
Private Const Muralla_Medio As Integer = 17358
Private Const Muralla_Min As Integer = 17393

Private Muralla_Position(1 To 4) As tAsedioPos

Public ReyIndex As Integer

Public Muralla(0 To 6, 1 To 4) As Integer

Public UserAsedio() As Integer
Public Enum Equipos
    Verde = 1
    Negro = 2
    Azul = 3
    Rojo = 4
End Enum

Public Enum AStatus
    Finalizada = 0
    Inscripcion = 1
    Curso = 2
End Enum

Public Type tAsedio
    Estado As AStatus
    MaxSlots As Integer
    Slots As Integer
    Costo As Long
    Premio As Long
    Tiempo As Long
End Type

Public Type flagsAsedio
    Participando As Boolean
    Slot As Integer
    Team As Integer
End Type

Private Type tAsedioPos
    Map As Byte
    X As Byte
    Y As Byte
End Type

Public ReyTeam As Byte

Public Asedio As tAsedio
Public Sub WarpUserCharX(ByVal UserIndex As Integer, ByVal Mapa As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)
    Dim NuevaPos As WorldPos
    Dim FuturePos As WorldPos
    FuturePos.Map = Mapa
    FuturePos.X = X
    FuturePos.Y = Y
    Call ClosestLegalPos(FuturePos, NuevaPos)

    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, FX)
End Sub
Public Sub Iniciar_Asedio(ByVal UserIndex As Integer, ByVal MaxSlot As Integer, ByVal Costo As Long, ByVal Tiempo As Long)

    Select Case Asedio.Estado
    Case AStatus.Inscripcion
        If UserIndex > 0 Then Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Las inscripciones están abiertas!" & FONTTYPE_Motd4)
        Exit Sub
    Case AStatus.Curso
        If UserIndex > 0 Then Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡El evento ya ha comenzado!" & FONTTYPE_Motd4)
        Exit Sub
    End Select

    '    If MaxSlot Mod 4 <> 0 Then
    '        Call SendData(SendTarget.toIndex, UserIndex, 0, "||La cantidad de participantes tienen que ser múltiplos de 4!" & FONTTYPE_Motd4)
    '        Exit Sub
    '    End If

    If Tiempo < 5 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||!El tiempo minimo es de 15 minutos¡" & FONTTYPE_Motd4)
        Exit Sub
    End If

    Asedio.MaxSlots = MaxSlot
    Asedio.Slots = 0
    Asedio.Costo = Costo
    Asedio.Premio = 1000000
    Asedio.Estado = AStatus.Curso
    Asedio.Tiempo = Tiempo

    If ReyIndex > 0 Then
        If Npclist(ReyIndex).Numero = ReyNPC Then
            Call QuitarNPC(ReyIndex)
        End If
    End If


    '[seth]cambiar posición muralla

    With Muralla_Position(Equipos.Azul)
        .Map = 192
        .X = 14
        .Y = 52
    End With

    With Muralla_Position(Equipos.Verde)
        .Map = 192
        .X = 47
        .Y = 76
    End With

    With Muralla_Position(Equipos.Rojo)
        .Map = 192
        .X = 80
        .Y = 51
    End With

    With Muralla_Position(Equipos.Negro)
        .Map = 192
        .X = 47
        .Y = 27
    End With
    '[/seth]

    Dim i As Byte
    Dim j As Byte
    Dim Position As WorldPos

    For i = 0 To 6
        For j = 1 To 4
            Position.Map = Muralla_Position(j).Map
            Position.X = Muralla_Position(j).X + i
            Position.Y = Muralla_Position(j).Y
            Muralla(i, j) = SpawnNpc(MurallaNPC, Position, False, False)
            Npclist(Muralla(i, j)).MurallaEquipo = j
            Npclist(Muralla(i, j)).MurallaIndex = i
            Call CalcularGrafico(Muralla(i, j))
        Next j
    Next i

    Dim PosRey As WorldPos
    PosRey.Map = 115
    PosRey.X = 46
    PosRey.Y = 58

    ReyIndex = SpawnNpc(ReyNPC, PosRey, False, False)

    ReDim UserAsedio(1 To Asedio.MaxSlots, 1 To 4) As Integer

    Call SendData(SendTarget.ToAll, 0, 0, "||Se ha dado comienzo al Evento Asedio, para ingresar escribe /ASEDIO" & FONTTYPE_INFON)
    Call SendData(SendTarget.ToAll, 0, 0, "||Precio de Inscripción: " & Asedio.Costo & FONTTYPE_INFON)
    Call SendData(SendTarget.ToAll, 0, 0, "||Duración del Evento: " & Asedio.Tiempo & FONTTYPE_INFON)
    Call SendData(SendTarget.ToAll, 0, 0, "TW48")

End Sub
Public Sub Inscribir_Asedio(ByVal UserIndex As Integer)


    If UserList(UserIndex).flags.EstaDueleando1 = True Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No puedes ir al evento si estás duelando!" & FONTTYPE_WARNING)
        Exit Sub
    End If


    If UserList(UserIndex).pos.Map = 160 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No puedes ir al evento si estás en un torneo!" & FONTTYPE_WARNING)
        Exit Sub
    End If

    If UserList(UserIndex).pos.Map = 48 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No puedes ir al evento si estás en la prisión! ¡Cumple la condena malhechor!" & FONTTYPE_WARNING)
        Exit Sub
    End If
    If UserList(UserIndex).Asedio.Participando Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Si ya estás dentro! ¿Qué quieres?" & FONTTYPE_WARNING)
        Exit Sub
    End If

    If Asedio.Slots = Asedio.MaxSlots Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Mira, si yo te dejaría entrar, pero no cabe ni un alfiler!" & FONTTYPE_WARNING)
        Exit Sub
    End If

    If UserList(UserIndex).Stats.GLD - Asedio.Costo < 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No te digo que seas pobre, pero te falta oro para pagar la inscripción" & FONTTYPE_WARNING)
        Exit Sub
    End If

    Static NumTeam As Integer
    Dim i As Long

    If NumTeam = 4 Then NumTeam = 0
    NumTeam = NumTeam + 1

    For i = 1 To Asedio.MaxSlots
        If UserAsedio(i, NumTeam) = 0 Then
            UserList(UserIndex).Asedio.Slot = i
            UserList(UserIndex).Asedio.Team = NumTeam
            UserAsedio(i, NumTeam) = UserIndex
            Exit For
        End If
    Next i


    Asedio.Slots = Asedio.Slots + 1    'Primero prueba el asedio con 100 slots
    UserList(UserIndex).Asedio.Participando = True
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Asedio.Costo
    Asedio.Premio = Asedio.Premio + Asedio.Costo
    Call SendUserStatsBox(UserIndex)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has ingresado al evento! ¡Estás en el equipo " & NombreEquipo(NumTeam) & "!" & FONTTYPE_WARNING)

    Dim User_Position As tAsedioPos
    User_Position = PosBase(NumTeam)
    With User_Position
        Call WarpUserCharX(UserIndex, .Map, .X, .Y, True)
    End With
    Call EnviarAsedio("SSED" & Asedio.Tiempo)
    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, val(UserIndex), UserList(UserIndex).char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)
End Sub

Private Function PosBase(ByVal Team As Byte) As tAsedioPos
    If ReyTeam = Team Then
        PosBase.Map = 115
        PosBase.X = 51
        PosBase.Y = 15
        Exit Function
    End If

    Select Case Team
    Case 1
        PosBase.Map = 192
        PosBase.X = 80
        PosBase.Y = 40
    Case 2
        PosBase.Map = 192
        PosBase.X = 14
        PosBase.Y = 60
    Case 3
        PosBase.Map = 192
        PosBase.X = 48
        PosBase.Y = 17
    Case 4
        PosBase.Map = 192
        PosBase.X = 48
        PosBase.Y = 82
    End Select
End Function
Private Function NombreEquipo(ByVal Team As Byte) As String
    Select Case Team
    Case Equipos.Azul
        NombreEquipo = "Azul"
    Case Equipos.Negro
        NombreEquipo = "Negro"
    Case Equipos.Rojo
        NombreEquipo = "Rojo"
    Case Equipos.Verde
        NombreEquipo = "Verde"
    End Select
End Function
Public Sub MuereUser(ByVal UserIndex As Integer)
    UserList(UserIndex).flags.Muerto = 0
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    Call DarCuerpoDesnudo(UserIndex)

    '[MaTeO 9]
    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).pos.Map, val(UserIndex), UserList(UserIndex).char.Body, _
                        UserList(UserIndex).OrigChar.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, _
                        UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)
    '[/MaTeO 9]

    Call SendUserStatsBox(UserIndex)

    Dim User_Position As tAsedioPos
    User_Position = PosBase(UserList(UserIndex).Asedio.Team)
    With User_Position
        Call WarpUserCharX(UserIndex, .Map, .X, .Y, True)
    End With
End Sub
Public Sub MuereRey(ByVal UserIndex As Integer)
    If UserIndex > 0 Then
        With UserList(UserIndex)
            If .Asedio.Team > 0 Then
                Select Case .Asedio.Team
                Case Equipos.Negro
                    Npclist(ReyIndex).char.Body = 697
                Case Equipos.Verde
                    Npclist(ReyIndex).char.Body = 698
                Case Equipos.Azul
                    Npclist(ReyIndex).char.Body = 699
                Case Equipos.Rojo
                    Npclist(ReyIndex).char.Body = 63
                End Select
                Npclist(ReyIndex).Stats.MinHP = Npclist(ReyIndex).Stats.MaxHP
                Call ChangeNPCChar(ToMap, 0, Npclist(ReyIndex).pos.Map, ReyIndex, Npclist(ReyIndex).char.Body, Npclist(ReyIndex).char.Head, Npclist(ReyIndex).char.heading)
                ReyTeam = .Asedio.Team
                Call LogAsedio("El rey ahora es del equipo " & ReyTeam)
            End If
        End With
        Call EnviarAsedio("||¡El rey ahora es del equipo " & NombreEquipo(ReyTeam) & FONTTYPE_TALK)
    End If
End Sub
Public Sub DoTimerAsedio()
    If Asedio.Estado <> AStatus.Curso Then Exit Sub
    Asedio.Tiempo = Asedio.Tiempo - 1
    If Asedio.Tiempo = 0 Then
        'If ReyTeam = 0 Then
        ' Call EnviarAsedio("||Se ha finalizado el tiempo del evento, pero al no tener un ganador se agregan 5 minutos." & FONTTYPE_WARNING)
        ' Asedio.Tiempo = Asedio.Tiempo + 5
        '  Else
        Call SendData(SendTarget.ToAll, 0, 0, "||¡Ha finalizado el evento y el ganador es el equipo " & NombreEquipo(ReyTeam) & "!" & FONTTYPE_WARNING)
        Dim i As Long
        Dim j As Long
        Dim Participantes As Long
        For i = 1 To Asedio.MaxSlots

            If UserAsedio(i, ReyTeam) > 0 Then
                If UserList(UserAsedio(i, ReyTeam)).Asedio.Team = ReyTeam Then
                    Participantes = Participantes + 1
                End If
            End If
        Next i
        Dim PremioxP As Long
        If Participantes <> 0 Then
            PremioxP = Asedio.Premio / Participantes
        End If
        For i = 1 To Asedio.MaxSlots
            For j = 1 To 4
                If UserAsedio(i, j) > 0 Then
                    ' Call WarpUserCharX(UserAsedio(i, j), 1, 50, 50, False) 'PRUEBALO
                    Call LogAsedio("Damos premio a equipo: " & ReyTeam)
                    If j = ReyTeam Then
                        UserList(UserAsedio(i, j)).Stats.GLD = UserList(UserAsedio(i, j)).Stats.GLD + PremioxP
                        Call SendData(SendTarget.ToIndex, UserAsedio(i, j), 0, "||!Has ganado " & PremioxP & " de oro¡ " & FONTTYPE_INFO)
                        UserList(UserAsedio(i, j)).AoMCanjes = UserList(UserAsedio(i, j)).AoMCanjes + 1
                        Call SendData(SendTarget.ToIndex, UserAsedio(i, j), 0, "||¡Has ganado 1 AoMCanje!" & FONTTYPE_INFO)
                        Call SendUserStatsBox(UserAsedio(i, j))
                    End If
                    Call ResetFlagsAsedio(UserAsedio(i, j))
                End If
            Next j
        Next i

        If ReyIndex > 0 Then
            Call QuitarNPC(ReyIndex)
            ReyIndex = 0
        End If
        Asedio.Costo = 0
        Asedio.Estado = Finalizada
        Asedio.MaxSlots = 0
        Asedio.Tiempo = 0
        Asedio.Slots = 0
        Asedio.Premio = 0
        ReyTeam = 0

        'End If
    End If
    Call EnviarAsedio("SSED" & Asedio.Tiempo)
End Sub
Public Sub EnviarAsedio(ByRef rData As String)
    Call SendData(SendTarget.ToMap, 0, 114, rData)
    Call SendData(SendTarget.ToMap, 0, 115, rData)
    'Debug.Print "Envio: " & rData
End Sub
Public Sub ResetFlagsAsedio(ByVal UserIndex As Integer)
    With UserList(UserIndex).Asedio
        If .Slot <> 0 And .Team <> 0 Then
            UserAsedio(.Slot, .Team) = 0
        End If
        If .Participando Then
            Call WarpUserCharX(UserIndex, 34, 30, 50, False)
        End If
        .Participando = False
        .Slot = 0
        .Team = 0
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(UserIndex).char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, val(UserIndex), UserList(UserIndex).char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)
        Else
            Call DarCuerpoDesnudo(UserIndex, False)
        End If
    End With

End Sub
Public Sub CancelAsedio()
    Call SendData(SendTarget.ToAll, 0, 0, "||El Asedio ha sido cancelado. ¡El rey ha huido como un cobarde!" & FONTTYPE_WARNING)
    Dim i As Long
    Dim j As Long
    For i = 1 To Asedio.MaxSlots
        For j = 1 To 4
            If UserAsedio(i, j) > 0 Then
                'Call WarpUserCharX(UserAsedio(i, j), 1, 50, 50, False)
                UserList(UserAsedio(i, j)).Stats.GLD = UserList(UserAsedio(i, j)).Stats.GLD + Asedio.Costo
                Call SendUserStatsBox(UserAsedio(i, j))
                Call ResetFlagsAsedio(UserAsedio(i, j))
            End If
        Next j
    Next i

    For i = 0 To 6
        For j = 1 To 4
            If Muralla(i, j) > 0 Then
                Call QuitarNPC(Muralla(i, j))
                Muralla(i, j) = 0
            End If
        Next j
    Next i

    If ReyIndex > 0 Then
        Call QuitarNPC(ReyIndex)
        ReyIndex = 0
    End If

    Asedio.Costo = 0
    Asedio.Estado = Finalizada
    Asedio.MaxSlots = 0
    Asedio.Tiempo = 0
    Asedio.Slots = 0
    Asedio.Premio = 0
End Sub
Public Sub CalcularGrafico(ByVal NpcIndex As Integer)
    Dim Vida As Long
    Dim TeamNPC As Byte
    TeamNPC = Npclist(NpcIndex).MurallaEquipo

    If Npclist(NpcIndex).Stats.MinHP <= 0 Then
        MapData(Muralla_Position(TeamNPC).Map, Muralla_Position(TeamNPC).X + 3, Muralla_Position(TeamNPC).Y).OBJInfo.ObjIndex = 0
    Else
        Vida = Fix(((Npclist(NpcIndex).Stats.MinHP / 100) / (Npclist(NpcIndex).Stats.MaxHP / 100)) * 100) + 1

        MapData(Muralla_Position(TeamNPC).Map, Muralla_Position(TeamNPC).X + 3, Muralla_Position(TeamNPC).Y).OBJInfo.ObjIndex = ItemMuralla + TeamNPC - 1
    End If
    Select Case Vida
    Case 80 To 101    'Intacta
        ObjData(ItemMuralla + TeamNPC - 1).GrhIndex = Muralla_Max
    Case 35 To 79    'Maso maso
        ObjData(ItemMuralla + TeamNPC - 1).GrhIndex = Muralla_Medio
    Case 1 To 34    'Casi destruida
        ObjData(ItemMuralla + TeamNPC - 1).GrhIndex = Muralla_Min
    End Select
    If Vida = 0 Then
        Call SendToAreaByPos(Muralla_Position(TeamNPC).Map, Muralla_Position(TeamNPC).X + 3, Muralla_Position(TeamNPC).Y, "BO" & Muralla_Position(TeamNPC).X + 3 & "," & Muralla_Position(TeamNPC).Y)
    Else
        Call ModAreas.SendToAreaByPos(Muralla_Position(TeamNPC).Map, Muralla_Position(TeamNPC).X + 3, Muralla_Position(TeamNPC).Y, "HO" & ObjData(ItemMuralla + TeamNPC - 1).GrhIndex & "," & Muralla_Position(TeamNPC).X + 3 & "," & Muralla_Position(TeamNPC).Y)
    End If
End Sub


