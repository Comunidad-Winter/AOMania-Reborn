Attribute VB_Name = "NPCs"
Option Explicit

Public Const IntervaloDragonAlado As Integer = 3600

Public Const NpcDragonAlado       As Integer = 672

Public NpcDragonAladoVive         As Boolean

Public Const MapaGenios           As Byte = 56

Public Sub LoadDragonAlado()
    NpcDragonAladoVive = False

End Sub

Public Sub SpawnDragonAlado()

    Dim MiPos As WorldPos

    MiPos.Map = MapaGenios
    MiPos.X = RandomNumber(12, 89)
    MiPos.Y = RandomNumber(81, 90)

    Call SpawnNpc(NpcDragonAlado, MiPos, True, False)
    NpcDragonAladoVive = True

    Call SendData(SendTarget.ToAll, 0, 0, "||Dragon Alado Aparecio en Aomania." & FONTTYPE_GUILD)
    Call SendData(SendTarget.ToAll, 0, 0, "TW3")

End Sub

Sub QuitarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    Dim i As Integer

    UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1

    For i = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasIndex(i) = NpcIndex Then
            UserList(UserIndex).MascotasIndex(i) = 0
            UserList(UserIndex).MascotasType(i) = 0
            Exit For

        End If

    Next i

End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer, ByVal Mascota As Integer)
    Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1

End Sub

Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

    If UserList(UserIndex).Quest.Start = 1 Then
        If UserList(UserIndex).Quest.NumNpc > 0 Then

            Call MuereNpcQuest(UserIndex, NpcIndex, UserList(UserIndex).Quest.Quest)

        End If

    End If

    If Npclist(NpcIndex).Numero = ReyNPC And NpcIndex = ReyIndex Then

        Call modAsedio.MuereRey(UserIndex)
        Exit Sub

    End If

    If Npclist(NpcIndex).Numero = MurallaNPC Then

        Call modAsedio.CalcularGrafico(NpcIndex)

    End If

    Dim MiNPC As npc

    Dim Map   As Integer

    Dim X     As Integer

    Dim Y     As Integer

    MiNPC = Npclist(NpcIndex)

    If MiNPC.Numero = NpcThorn And Npclist(NpcIndex).Pos.Map = MapaCasaAbandonada1 Then

        NpcThornVive = False

    End If

    If MiNPC.Numero = NpcDragonAlado And Npclist(NpcIndex).Pos.Map = MapaGenios Then

        NpcDragonAladoVive = False

    End If

    If MiNPC.Numero = 661 And Npclist(NpcIndex).Pos.Map = mapainvo Then

        StatusInvo = False
        ConfInvo = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||�Has ganado 250000 puntos de experiencia!" & FONTTYPE_FIGHT)
        Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " ha matado al " & Npclist(NpcIndex).Name & "!! Felicidades!! Gana 250000 de experiencia." & FONTTYPE_GUILD)
        UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + 250000
        Call EnviarExp(UserIndex)
        Call CheckUserLevel(UserIndex)

    End If
    
    If MiNPC.Numero = 705 And Npclist(NpcIndex).Pos.Map = mapahades Then

        StatusHades = True
        ConfHades = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||�Has ganado 250000 puntos de experiencia!" & FONTTYPE_FIGHT)
        Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " ha matado al " & Npclist(NpcIndex).Name & "!! Felicidades!! Gana 250000 de experiencia." & FONTTYPE_GUILD)
        UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + 250000
        Call EnviarExp(UserIndex)
        Call CheckUserLevel(UserIndex)

    End If

    If MiNPC.Numero = 253 Then

        Call SendData(ToAll, 0, 0, "||Angeles ganaron la guerra entre bandas, reciben como premio experiencia!!!" & FONTTYPE_GUERRA)
        Call RespGuerrasAngeles
        Call Ban_Angeles

    End If

    If MiNPC.Numero = 254 Then

        Call SendData(ToAll, 0, 0, "||Demonios ganaron la guerra entre bandas, reciben como premio experiencia!!!" & FONTTYPE_GUERRA)
        Call RespGuerrasDemonio
        Call Ban_Demonios

    End If

    If MiNPC.Numero = NpcCorsarios Then

        Call SendData(ToAll, 0, 0, "||Piratas ganaron la batalla de medusas, reciben experiencia como premio!!!" & FONTTYPE_GUERRA)
        Call RespGuerrasCorsarios
        Call Med_Piratas

    End If

    If MiNPC.Numero = NpcPiratas Then

        Call SendData(ToAll, 0, 0, "||Corsarios ganaron la batalla de medusas, reciben experiencia como premio!!!" & FONTTYPE_GUERRA)
        Call RespGuerrasPiratas
        Call Med_Corsarios

    End If

    If MiNPC.Numero = NpcNosfe Then

        NickMataNosfe = UserIndex
        MataNosfe = True

    End If

    If (esPretoriano(NpcIndex) = 4) Then

        'seteamos todos estos 'flags' acorde para que cambien solos de alcoba
        Dim i    As Integer

        Dim j    As Integer

        Dim NPCI As Integer

        For i = 8 To 90
            For j = 8 To 90

                NPCI = MapData(Npclist(NpcIndex).Pos.Map, i, j).NpcIndex

                If NPCI > 0 Then
                    If esPretoriano(NPCI) > 0 Then

                        Npclist(NPCI).Invent.ArmourEqpSlot = IIf(Npclist(NpcIndex).Pos.X > 50, 1, 5)

                    End If

                End If

            Next j
        Next i

        Call CrearClanPretoriano(MAPA_PRETORIANO, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y)
    ElseIf esPretoriano(NpcIndex) > 0 Then
        Npclist(NpcIndex).Invent.ArmourEqpSlot = 0

    End If

    If MiNPC.Pos.Map = mapainvo Then MapInfo(mapainvo).criatinv = 0

    'Quitamos el npc
    Call QuitarNPC(NpcIndex)

    If UserIndex > 0 Then    ' Lo mato un usuario?

        Call AccionNpcCastillos(MiNPC.Numero, UserIndex)

        If MiNPC.flags.Snd3 > 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & MiNPC.flags.Snd3)

        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun

        'El user que lo mato tiene mascotas?
        If UserList(UserIndex).NroMacotas > 0 Then

            Dim T As Integer

            For T = 1 To MAXMASCOTAS

                If UserList(UserIndex).MascotasIndex(T) > 0 Then
                    If Npclist(UserList(UserIndex).MascotasIndex(T)).TargetNpc = NpcIndex Then

                        Call FollowAmo(UserList(UserIndex).MascotasIndex(T))

                    End If

                End If

            Next T

        End If

        If MiNPC.MaestroUser = 0 Then

            'Tiramos el oro
            Call NPCTirarOro(MiNPC, UserIndex)

            Call EnviarOro(UserIndex)

            'Tiramos el inventario
            Call NPC_TIRAR_ITEMS(MiNPC, UserIndex)

        End If

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has matado la criatura!" & FONTTYPE_INFO)

        If UserList(UserIndex).Stats.NPCsMuertos < 32000 Then UserList(UserIndex).Stats.NPCsMuertos = UserList(UserIndex).Stats.NPCsMuertos + 1

        If MiNPC.Stats.Alineacion = 0 Then
            If MiNPC.Numero = Guardias Then

                Call VolverCriminal(UserIndex)

            End If

            If Not EsDios(UserList(UserIndex).Name) Then Call AddtoVar(UserList(UserIndex).Reputacion.AsesinoRep, vlASESINO, MAXREP)

        ElseIf MiNPC.Stats.Alineacion = 1 Then
            Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlCAZADOR, MAXREP)
        ElseIf MiNPC.Stats.Alineacion = 2 Then
            Call AddtoVar(UserList(UserIndex).Reputacion.NobleRep, vlASESINO / 2, MAXREP)
        ElseIf MiNPC.Stats.Alineacion = 4 Then
            Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlCAZADOR, MAXREP)

        End If

        Call CheckUserLevel(UserIndex)
    End If    ' Userindex > 0

    'ReSpawn o no
    Call RespawnNPC(MiNPC)

End Sub

Sub ResetNpcFlags(ByVal NpcIndex As Integer)
    'Clear the npc's flags

    With Npclist(NpcIndex).flags
        .AfectaParalisis = 0
        .AguaValida = 0
        .AttackedBy = ""
        .Attacking = 0
        .BackUp = 0
        .Bendicion = 0
        .Domable = 0
        .Envenenado = 0
        .Faccion = 0
        .Follow = False
        .LanzaSpells = 0
        .GolpeExacto = 0
        .Invisible = 0
        .Maldicion = 0
        .OldHostil = 0
        .OldMovement = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Respawn = 0
        .RespawnOrigPos = 0
        .Snd1 = 0
        .Snd2 = 0
        .Snd3 = 0
        .TierraInvalida = 0
        .AtacaAPJ = 0
        .AtacaANPC = 0
        .AIAlineacion = e_Alineacion.ninguna
        .AIPersonalidad = e_Personalidad.ninguna

    End With

End Sub

Sub ResetNpcCounters(ByVal NpcIndex As Integer)

    Npclist(NpcIndex).Contadores.Paralisis = 0
    Npclist(NpcIndex).Contadores.TiempoExistencia = 0

End Sub

Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)

    Npclist(NpcIndex).char.Body = 0
    Npclist(NpcIndex).char.CascoAnim = 0
    Npclist(NpcIndex).char.CharIndex = 0
    Npclist(NpcIndex).char.FX = 0
    Npclist(NpcIndex).char.Head = 0
    Npclist(NpcIndex).char.heading = 0
    Npclist(NpcIndex).char.loops = 0
    Npclist(NpcIndex).char.ShieldAnim = 0
    Npclist(NpcIndex).char.WeaponAnim = 0

End Sub

Sub ResetNpcCriatures(ByVal NpcIndex As Integer)

    Dim j As Integer

    For j = 1 To Npclist(NpcIndex).NroCriaturas
        Npclist(NpcIndex).Criaturas(j).NpcIndex = 0
        Npclist(NpcIndex).Criaturas(j).NpcName = ""
    Next j

    Npclist(NpcIndex).NroCriaturas = 0

End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)

    Dim j As Integer

    For j = 1 To Npclist(NpcIndex).NroExpresiones
        Npclist(NpcIndex).Expresiones(j) = ""
    Next j

    Npclist(NpcIndex).NroExpresiones = 0

End Sub

Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)

    Npclist(NpcIndex).Attackable = 0
    Npclist(NpcIndex).CanAttack = 0
    Npclist(NpcIndex).Comercia = 0
    Npclist(NpcIndex).GiveEXP = 0
    Npclist(NpcIndex).GiveGLD = 0
    Npclist(NpcIndex).Hostile = 0
    Npclist(NpcIndex).Inflacion = 0
    Npclist(NpcIndex).InvReSpawn = 0
    Npclist(NpcIndex).level = 0

    If Npclist(NpcIndex).MaestroUser > 0 Then Call QuitarMascota(Npclist(NpcIndex).MaestroUser, NpcIndex)
    If Npclist(NpcIndex).MaestroNpc > 0 Then Call QuitarMascotaNpc(Npclist(NpcIndex).MaestroNpc, NpcIndex)

    Npclist(NpcIndex).MaestroUser = 0
    Npclist(NpcIndex).MaestroNpc = 0

    Npclist(NpcIndex).Mascotas = 0
    Npclist(NpcIndex).Movement = 0
    Npclist(NpcIndex).Name = "NPC SIN INICIAR"
    Npclist(NpcIndex).NPCtype = 0
    Npclist(NpcIndex).Numero = 0
    Npclist(NpcIndex).Orig.Map = 0
    Npclist(NpcIndex).Orig.X = 0
    Npclist(NpcIndex).Orig.Y = 0
    Npclist(NpcIndex).PoderAtaque = 0
    Npclist(NpcIndex).PoderEvasion = 0
    Npclist(NpcIndex).Pos.Map = 0
    Npclist(NpcIndex).Pos.X = 0
    Npclist(NpcIndex).Pos.Y = 0
    Npclist(NpcIndex).SkillDomar = 0
    Npclist(NpcIndex).Target = 0
    Npclist(NpcIndex).TargetNpc = 0
    Npclist(NpcIndex).TipoItems = 0
    Npclist(NpcIndex).Veneno = 0
    Npclist(NpcIndex).Desc = ""

    Dim j As Integer

    For j = 1 To Npclist(NpcIndex).NroSpells
        Npclist(NpcIndex).Spells(j) = 0
    Next j

    Call ResetNpcCharInfo(NpcIndex)
    Call ResetNpcCriatures(NpcIndex)
    Call ResetExpresiones(NpcIndex)

End Sub

Sub QuitarNPC(ByVal NpcIndex As Integer)

    On Error GoTo errhandler

    With Npclist(NpcIndex)
        .flags.NPCActive = False

        If InMapBounds(.Pos.Map, .Pos.X, .Pos.Y) Then
            Call EraseNPCChar(NpcIndex)

        End If

    End With

    'Nos aseguramos de que el inventario sea removido...
    'asi los lobos no volveran a tirar armaduras ;))
    Call ResetNpcInv(NpcIndex)
    Call ResetNpcFlags(NpcIndex)
    Call ResetNpcCounters(NpcIndex)
    Call ResetNpcMainInfo(NpcIndex)

    If NpcIndex = LastNPC Then

        Do Until Npclist(LastNPC).flags.NPCActive
            LastNPC = LastNPC - 1

            If LastNPC < 1 Then Exit Do
        Loop

    End If

    If NumNPCs <> 0 Then
        NumNPCs = NumNPCs - 1

    End If

    Exit Sub

errhandler:
    Npclist(NpcIndex).flags.NPCActive = False
    Call LogError("Error en QuitarNPC")

End Sub

Function TestSpawnTrigger(Pos As WorldPos) As Boolean

    If LegalPos(Pos.Map, Pos.X, Pos.Y) Then
        TestSpawnTrigger = MapData(Pos.Map, Pos.X, Pos.Y).Trigger <> 3 And MapData(Pos.Map, Pos.X, Pos.Y).Trigger <> 2 And MapData(Pos.Map, Pos.X, Pos.Y).Trigger <> 1

    End If

End Function

Sub CrearNPC(NroNPC As Integer, Mapa As Integer, OrigPos As WorldPos)

    'Call LogTarea("Sub CrearNPC")
    'Crea un NPC del tipo NRONPC
    On Error GoTo err

    Dim error          As Integer

    Dim Pos            As WorldPos

    Dim newpos         As WorldPos

    Dim nIndex         As Integer

    Dim PosicionValida As Boolean

    Dim Iteraciones    As Long

    Dim Map            As Integer

    Dim X              As Integer

    Dim Y              As Integer

    error = 1

    If (Mapa = 0) Then Exit Sub

    error = 2
    nIndex = OpenNPC(NroNPC) 'Conseguimos un indice

    If nIndex > MAXNPCS Then Exit Sub

    error = 3

    'Necesita ser respawned en un lugar especifico
    If InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
        error = 4
        Map = OrigPos.Map
        X = OrigPos.X
        Y = OrigPos.Y
        Npclist(nIndex).Orig = OrigPos
        Npclist(nIndex).Pos = OrigPos
   
    Else
        error = 5
  
        Pos.Map = Mapa 'mapa
    
        Do While Not PosicionValida
            error = 6
            Randomize (Timer)
            Pos.X = CInt(Rnd * 100 + 1) 'Obtenemos posicion al azar en x
            Pos.Y = CInt(Rnd * 100 + 1) 'Obtenemos posicion al azar en y
            error = 7
            Call ClosestLegalPosNpc(Pos, newpos, IIf(Npclist(nIndex).flags.AguaValida = 1, True, False)) 'Nos devuelve la posicion valida mas cercana
            error = 8

            'Si X e Y son iguales a 0 significa que no se encontro posicion valida
            If LegalPosNPC(newpos.Map, newpos.X, newpos.Y, IIf(Npclist(nIndex).flags.AguaValida = 1, True, False)) And Not HayPCarea(newpos) And TestSpawnTrigger(newpos) Then
                error = 9
                'Asignamos las nuevas coordenas solo si son validas
                Npclist(nIndex).Pos.Map = newpos.Map
                Npclist(nIndex).Pos.X = newpos.X
                Npclist(nIndex).Pos.Y = newpos.Y
                PosicionValida = True
            Else
                error = 10
                newpos.X = 0
                newpos.Y = 0
        
            End If
        
            error = 11
            'for debug
            Iteraciones = Iteraciones + 1

            If Iteraciones > MAXSPAWNATTEMPS Then
                Call QuitarNPC(nIndex)
                Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & Mapa & " NroNpc:" & NroNPC)
                Exit Sub

            End If

        Loop
    
        error = 12
        'asignamos las nuevas coordenas
        Map = newpos.Map
        X = Npclist(nIndex).Pos.X
        Y = Npclist(nIndex).Pos.Y

    End If

    error = 13
    'Crea el NPC
    Call MakeNPCChar(ToMap, 0, Map, nIndex, Map, X, Y)

    Exit Sub
err:
    LogError ("ERROR en CREARNPC :" & err.Description & " " & " --->" & error)

End Sub

Sub MakeNPCChar(sndRoute As Byte, _
                sndIndex As Integer, _
                sndMap As Integer, _
                NpcIndex As Integer, _
                ByVal Map As Integer, _
                ByVal X As Integer, _
                ByVal Y As Integer)

    Dim CharIndex As Integer

    If Npclist(NpcIndex).char.CharIndex = 0 Then
        CharIndex = NextOpenCharIndex
        Npclist(NpcIndex).char.CharIndex = CharIndex
        CharList(CharIndex) = NpcIndex

    End If

    MapData(Map, X, Y).NpcIndex = NpcIndex

    If sndRoute = SendTarget.ToMap Then
        Call AgregarNpc(NpcIndex)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "BC" & Npclist(NpcIndex).char.Body & "," & Npclist(NpcIndex).char.Head & "," & Npclist(NpcIndex).char.heading & "," & Npclist(NpcIndex).char.CharIndex & "," & X & "," & Y)
        Call SendData(sndRoute, sndIndex, sndMap, "XN" & Npclist(NpcIndex).char.CharIndex & "," & Npclist(NpcIndex).NPCtype & "," & Npclist(NpcIndex).Name & "," & Npclist(NpcIndex).Hostile)

    End If

End Sub

Sub ChangeNPCChar(ByVal sndRoute As Byte, _
                  ByVal sndIndex As Integer, _
                  ByVal sndMap As Integer, _
                  ByVal NpcIndex As Integer, _
                  ByVal Body As Integer, _
                  ByVal Head As Integer, _
                  ByVal heading As eHeading)

    If NpcIndex > 0 Then
        Npclist(NpcIndex).char.Body = Body
        Npclist(NpcIndex).char.Head = Head
        Npclist(NpcIndex).char.heading = heading
        Call SendData(sndRoute, sndIndex, sndMap, "CP" & Npclist(NpcIndex).char.CharIndex & "," & Body & "," & Head & "," & heading)

    End If

End Sub

Sub EraseNPCChar(ByVal NpcIndex As Integer)

    If Npclist(NpcIndex).char.CharIndex <> 0 Then CharList(Npclist(NpcIndex).char.CharIndex) = 0

    If Npclist(NpcIndex).char.CharIndex = LastChar Then

        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1

            If LastChar <= 1 Then Exit Do
        Loop

    End If

    'Quitamos del mapa
    MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

    'Actualizamos los clientes
    Call SendData(SendTarget.ToMap, 0, Npclist(NpcIndex).Pos.Map, "BP" & Npclist(NpcIndex).char.CharIndex)

    'Update la lista npc
    Npclist(NpcIndex).char.CharIndex = 0

    'update NumChars
    NumChars = NumChars - 1

End Sub

Sub MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte)

    On Error GoTo errh

    Dim nPos As WorldPos

    nPos = Npclist(NpcIndex).Pos
    Call HeadtoPos(nHeading, nPos)

    'Es mascota ????
    If Npclist(NpcIndex).MaestroUser > 0 Then

        ' es una posicion legal
        If LegalPos(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y, Npclist(NpcIndex).flags.AguaValida = 1, EsElemental(NpcIndex)) Then

            If Npclist(NpcIndex).flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub

            If Npclist(NpcIndex).flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub

            Call SendToNpcArea(NpcIndex, "*" & Npclist(NpcIndex).char.CharIndex & "," & nPos.X & "," & nPos.Y)

            'Update map and user pos
            MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0
            Npclist(NpcIndex).Pos = nPos
            Npclist(NpcIndex).char.heading = nHeading
            MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = NpcIndex
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)

        End If

    Else    ' No es mascota

        ' Controlamos que la posicion sea legal, los npc que
        ' no son mascotas tienen mas restricciones de movimiento.
        If LegalPosNPC(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y, Npclist(NpcIndex).flags.AguaValida) Then

            If Npclist(NpcIndex).flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub

            If Npclist(NpcIndex).flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub

            '[Alejo-18-5]
            'server

            Call SendToNpcArea(NpcIndex, "*" & Npclist(NpcIndex).char.CharIndex & "," & nPos.X & "," & nPos.Y)

            'Update map and user pos
            MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0
            Npclist(NpcIndex).Pos = nPos
            Npclist(NpcIndex).char.heading = nHeading
            MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = NpcIndex

            Call CheckUpdateNeededNpc(NpcIndex, nHeading)

        Else

            If Npclist(NpcIndex).Movement = TipoAI.NpcPathfinding Then
                'Someone has blocked the npc's way, we must to seek a new path!
                Npclist(NpcIndex).PFINFO.PathLenght = 0

            End If

        End If

    End If

    Exit Sub

errh:
    LogError ("Error en move npc " & NpcIndex)

End Sub

Function NextOpenNPC() As Integer
    'Call LogTarea("Sub NextOpenNPC")

    On Error GoTo errhandler

    Dim LoopC As Integer

    For LoopC = 1 To MAXNPCS + 1

        If LoopC > MAXNPCS Then Exit For
        If Not Npclist(LoopC).flags.NPCActive Then Exit For
    Next LoopC

    NextOpenNPC = LoopC

    Exit Function
errhandler:
    Call LogError("Error en NextOpenNPC")

End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer)

    Dim n As Integer

    n = RandomNumber(1, 100)

    If n < 30 Then
        UserList(UserIndex).flags.Envenenado = 1
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||��La criatura te ha envenenado!!" & FONTTYPE_Motd4)

    End If

End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, _
                  Pos As WorldPos, _
                  ByVal FX As Boolean, _
                  ByVal Respawn As Boolean) As Integer
    'Crea un NPC del tipo Npcindex
    'Call LogTarea("Sub SpawnNpc")

    Dim newpos         As WorldPos

    Dim nIndex         As Integer

    Dim PosicionValida As Boolean

    Dim Map            As Integer

    Dim X              As Integer

    Dim Y              As Integer

    Dim it             As Integer

    nIndex = OpenNPC(NpcIndex, Respawn)   'Conseguimos un indice

    it = 0

    If nIndex > MAXNPCS Then
        SpawnNpc = nIndex
        Exit Function

    End If

    Do While Not PosicionValida
        
        Call ClosestLegalPosNpc(Pos, newpos, IIf(Npclist(nIndex).flags.AguaValida = 1, True, False))  'Nos devuelve la posicion valida mas cercana

        'Si X e Y son iguales a 0 significa que no se encontro posicion valida
        If LegalPosNPC(newpos.Map, newpos.X, newpos.Y, Npclist(nIndex).flags.AguaValida) Then
            'Asignamos las nuevas coordenas solo si son validas
            Npclist(nIndex).Pos.Map = newpos.Map
            Npclist(nIndex).Pos.X = newpos.X
            Npclist(nIndex).Pos.Y = newpos.Y
            PosicionValida = True
        Else
            newpos.X = 0
            newpos.Y = 0

        End If
        
        it = it + 1
        
        If it > MAXSPAWNATTEMPS Then
            Call QuitarNPC(nIndex)
            SpawnNpc = MAXNPCS
            'Call LogError("Mas de " & MAXSPAWNATTEMPS & " iteraciones en SpawnNpc Mapa:" & Pos.Map & " Index:" & NpcIndex)
            Exit Function

        End If

    Loop
    
    'asignamos las nuevas coordenas
    Map = newpos.Map
    X = Npclist(nIndex).Pos.X
    Y = Npclist(nIndex).Pos.Y

    'Crea el NPC
    Call MakeNPCChar(ToMap, 0, Map, nIndex, Map, X, Y)

    If FX Then
        'Call SendData(ToMap, 0, Map, "TW" & SND_WARP)
        Call SendData(ToMap, 0, Map, "TW" & SND_WARP & "," & Npclist(nIndex).char.CharIndex)
        Call SendData(ToMap, 0, Map, "CFX" & Npclist(nIndex).char.CharIndex & "," & FXWARP & "," & 0)

    End If

    SpawnNpc = nIndex

End Function


Sub RespawnNPC(MiNPC As npc)

    If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.Pos.Map, MiNPC.Orig)

End Sub

'Devuelve el nro de enemigos que hay en el Mapa Map
Function NPCHostiles(ByVal Map As Integer) As Integer

    Dim NpcIndex As Integer

    Dim cont     As Integer

    'Contador
    cont = 0

    For NpcIndex = 1 To LastNPC

        '�esta vivo?
        If Npclist(NpcIndex).flags.NPCActive And Npclist(NpcIndex).Pos.Map = Map And Npclist(NpcIndex).Hostile = 1 And Npclist(NpcIndex).Stats.Alineacion = 2 Then
            cont = cont + 1

        End If

    Next NpcIndex

    NPCHostiles = cont

End Function

Sub NPCTirarOro(MiNPC As npc, UserIndex As Integer)

    'SI EL NPC TIENE ORO LO TIRAMOS
    If MiNPC.GiveGLD > 0 And MiNPC.GiveGLD < MaxOro Then

        Dim MiAux As Long

        Dim MiObj As Obj

        MiAux = MiNPC.GiveGLD

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La criatura ha dropeado " & MiAux & " monedas de oro." & FONTTYPE_Motd4)

        Do While MiAux > MAX_INVENTORY_OBJS
            MiObj.Amount = MAX_INVENTORY_OBJS
            MiObj.ObjIndex = iORO
            Call TirarItemAlPiso(MiNPC.Pos, MiObj)
            MiAux = MiAux - MAX_INVENTORY_OBJS
        Loop

        If MiAux > 0 Then
            MiObj.Amount = MiAux
            MiObj.ObjIndex = iORO
            Call TirarItemAlPiso(MiNPC.Pos, MiObj)

        End If

    ElseIf MiNPC.GiveGLD > MaxOro Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MiNPC.GiveGLD

    End If

    If MiNPC.GiveGLD = 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La criatura no ha dejado oro." & FONTTYPE_Motd4)

    End If

End Sub

Function OpenNPC(ByVal NPCNumber As Integer, Optional ByVal Respawn = True) As Integer

    '###################################################
    '#               ATENCION PELIGRO                  #
    '###################################################
    '
    '    ���� NO USAR GetVar PARA LEER LOS NPCS !!!!
    '
    'El que ose desafiar esta LEY, se las tendr� que ver
    'con migo. Para leer los NPCS se deber� usar la
    'nueva clase clsLeerInis.
    '
    'Alejo
    '
    '###################################################

    Dim NpcIndex As Integer

    Dim npcfile  As String

    Dim Leer     As clsIniManager
    
    If NPCNumber > 499 Then
        Set Leer = LeerNPCsHostil
    Else
        Set Leer = LeerNPCs

    End If
    
    NpcIndex = NextOpenNPC

    If NpcIndex > MAXNPCS Then    'Limite de npcs
        OpenNPC = NpcIndex
        Exit Function

    End If

    Npclist(NpcIndex).Numero = NPCNumber
    Npclist(NpcIndex).Name = Leer.GetValue("NPC" & NPCNumber, "Name")
    Npclist(NpcIndex).Desc = Leer.GetValue("NPC" & NPCNumber, "Desc")

    Npclist(NpcIndex).Movement = val(Leer.GetValue("NPC" & NPCNumber, "Movement"))
    Npclist(NpcIndex).flags.OldMovement = Npclist(NpcIndex).Movement

    Npclist(NpcIndex).flags.AguaValida = val(Leer.GetValue("NPC" & NPCNumber, "AguaValida"))
    Npclist(NpcIndex).flags.TierraInvalida = val(Leer.GetValue("NPC" & NPCNumber, "TierraInValida"))
    Npclist(NpcIndex).flags.Faccion = val(Leer.GetValue("NPC" & NPCNumber, "Faccion"))

    Npclist(NpcIndex).NPCtype = val(Leer.GetValue("NPC" & NPCNumber, "NpcType"))

    Npclist(NpcIndex).char.Body = val(Leer.GetValue("NPC" & NPCNumber, "Body"))
    Npclist(NpcIndex).char.Head = val(Leer.GetValue("NPC" & NPCNumber, "Head"))
    Npclist(NpcIndex).char.heading = val(Leer.GetValue("NPC" & NPCNumber, "Heading"))

    Npclist(NpcIndex).Attackable = val(Leer.GetValue("NPC" & NPCNumber, "Attackable"))
    Npclist(NpcIndex).Comercia = val(Leer.GetValue("NPC" & NPCNumber, "Comercia"))
    Npclist(NpcIndex).Hostile = val(Leer.GetValue("NPC" & NPCNumber, "Hostile"))
    Npclist(NpcIndex).flags.OldHostil = Npclist(NpcIndex).Hostile

    If DiaEspecialExp = True Then
        Npclist(NpcIndex).GiveEXP = Round((val(Leer.GetValue("NPC" & NPCNumber, "GiveEXP")) * Multexp) * LoteriaCriatura)
    Else

        If SistemaCriatura.ExpCriatura = True Then
            If Npclist(NpcIndex).Numero = NpcCriatura Then
                Npclist(NpcIndex).GiveEXP = Round((val(Leer.GetValue("NPC" & NPCNumber, "GiveEXP")) * Multexp) * LoteriaCriatura)
            Else
                Npclist(NpcIndex).GiveEXP = val(Leer.GetValue("NPC" & NPCNumber, "GiveEXP")) * Multexp

            End If

        Else
            Npclist(NpcIndex).GiveEXP = val(Leer.GetValue("NPC" & NPCNumber, "GiveEXP")) * Multexp

        End If

    End If

    Dim LagaNDrop

    If Npclist(NpcIndex).Hostile = 1 Then
        Npclist(NpcIndex).Drops.NumDrop = val(Leer.GetValue("NPC" & NPCNumber, "NROITEMS"))

        For LagaNDrop = 1 To Npclist(NpcIndex).Drops.NumDrop

            Npclist(NpcIndex).Drops.DropIndex(LagaNDrop) = val(readfield2(1, Leer.GetValue("NPC" & NPCNumber, "Obj" & LagaNDrop & ""), 45))
            Npclist(NpcIndex).Drops.Amount(LagaNDrop) = val(readfield2(2, Leer.GetValue("NPC" & NPCNumber, "Obj" & LagaNDrop & ""), 45))
            Npclist(NpcIndex).Drops.Porcentaje(LagaNDrop) = val(Leer.GetValue("NPC" & NPCNumber, "Prob" & LagaNDrop & ""))
        Next LagaNDrop

    End If

    'Npclist(NpcIndex).flags.ExpDada = Npclist(NpcIndex).GiveEXP
    Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).GiveEXP

    Npclist(NpcIndex).Veneno = val(Leer.GetValue("NPC" & NPCNumber, "Veneno"))

    Npclist(NpcIndex).flags.Domable = val(Leer.GetValue("NPC" & NPCNumber, "Domable"))

    If DiaEspecialOro = True Then
        Npclist(NpcIndex).GiveGLD = Round(val(Leer.GetValue("NPC" & NPCNumber, "GiveGLD")) * MultOro) * LoteriaCriatura
    Else

        If SistemaCriatura.OroCriatura = True Then
            If Npclist(NpcIndex).Numero = NpcCriatura Then
                Npclist(NpcIndex).GiveGLD = Round(val(Leer.GetValue("NPC" & NPCNumber, "GiveGLD")) * MultOro) * LoteriaCriatura
            Else
                Npclist(NpcIndex).GiveGLD = val(Leer.GetValue("NPC" & NPCNumber, "GiveGLD")) * MultOro

            End If

        Else
            Npclist(NpcIndex).GiveGLD = val(Leer.GetValue("NPC" & NPCNumber, "GiveGLD")) * MultOro

        End If

    End If

    Npclist(NpcIndex).PoderAtaque = val(Leer.GetValue("NPC" & NPCNumber, "PoderAtaque"))
    Npclist(NpcIndex).PoderEvasion = val(Leer.GetValue("NPC" & NPCNumber, "PoderEvasion"))

    Npclist(NpcIndex).InvReSpawn = val(Leer.GetValue("NPC" & NPCNumber, "InvReSpawn"))

    Npclist(NpcIndex).Stats.MaxHP = val(Leer.GetValue("NPC" & NPCNumber, "MaxHP"))
    Npclist(NpcIndex).Stats.MinHP = val(Leer.GetValue("NPC" & NPCNumber, "MinHP"))
    Npclist(NpcIndex).Stats.MaxHit = val(Leer.GetValue("NPC" & NPCNumber, "MaxHIT"))
    Npclist(NpcIndex).Stats.MinHit = val(Leer.GetValue("NPC" & NPCNumber, "MinHIT"))
    Npclist(NpcIndex).Stats.def = val(Leer.GetValue("NPC" & NPCNumber, "DEF"))
    Npclist(NpcIndex).Stats.Alineacion = val(Leer.GetValue("NPC" & NPCNumber, "Alineacion"))

    Dim LoopC As Integer

    Dim ln    As String

    Npclist(NpcIndex).Invent.NroItems = val(Leer.GetValue("NPC" & NPCNumber, "NROITEMS"))

    Dim Lenght As Integer

    If Npclist(NpcIndex).Invent.NroItems > MAX_INVENTORY_SLOTS Then
        Npclist(NpcIndex).Invent.NroItems = MAX_INVENTORY_SLOTS

    End If

    For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
        ln = Leer.GetValue("NPC" & NPCNumber, "Obj" & LoopC)

        If Len(ln) > 0 Then
            Lenght = InStr(1, ln, "'")

            If Lenght > 0 Then    ' Esto hasta que le diga al wachin que datee bien las cosas xD
                ln = mid$(ln, 1, Lenght)
                ln = Replace(ln, "'", "-")
            Else
                ln = ln + "-0"

            End If

            Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(readfield2(1, ln, 45))
            Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(readfield2(2, ln, 45))
            Npclist(NpcIndex).Invent.Object(LoopC).ProbTirar = val(readfield2(3, ln, 45))

        End If

    Next LoopC

    Npclist(NpcIndex).flags.LanzaSpells = val(Leer.GetValue("NPC" & NPCNumber, "LanzaSpells"))

    If Npclist(NpcIndex).flags.LanzaSpells > 0 Then ReDim Npclist(NpcIndex).Spells(1 To Npclist(NpcIndex).flags.LanzaSpells)

    For LoopC = 1 To Npclist(NpcIndex).flags.LanzaSpells
        Npclist(NpcIndex).Spells(LoopC) = val(Leer.GetValue("NPC" & NPCNumber, "Sp" & LoopC))
    Next LoopC

    If Npclist(NpcIndex).NPCtype = eNPCType.Entrenador Then
        Npclist(NpcIndex).NroCriaturas = val(Leer.GetValue("NPC" & NPCNumber, "NroCriaturas"))
        ReDim Npclist(NpcIndex).Criaturas(1 To Npclist(NpcIndex).NroCriaturas) As tCriaturasEntrenador

        For LoopC = 1 To Npclist(NpcIndex).NroCriaturas
            Npclist(NpcIndex).Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NPCNumber, "CI" & LoopC)
            Npclist(NpcIndex).Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NPCNumber, "CN" & LoopC)
        Next LoopC

    End If

    Npclist(NpcIndex).Inflacion = val(Leer.GetValue("NPC" & NPCNumber, "Inflacion"))

    Npclist(NpcIndex).flags.NPCActive = True

    If Respawn Then
        Npclist(NpcIndex).flags.Respawn = val(Leer.GetValue("NPC" & NPCNumber, "ReSpawn"))
    Else
        Npclist(NpcIndex).flags.Respawn = 1

    End If

    Npclist(NpcIndex).flags.BackUp = val(Leer.GetValue("NPC" & NPCNumber, "BackUp"))
    Npclist(NpcIndex).flags.RespawnOrigPos = val(Leer.GetValue("NPC" & NPCNumber, "OrigPos"))
    Npclist(NpcIndex).flags.AfectaParalisis = val(Leer.GetValue("NPC" & NPCNumber, "AfectaParalisis"))
    Npclist(NpcIndex).flags.GolpeExacto = val(Leer.GetValue("NPC" & NPCNumber, "GolpeExacto"))
    Npclist(NpcIndex).flags.Magiainvisible = val(Leer.GetValue("NPC" & NPCNumber, "Magiainvisible"))
    
    Npclist(NpcIndex).flags.Snd1 = val(Leer.GetValue("NPC" & NPCNumber, "Snd1"))
    Npclist(NpcIndex).flags.Snd2 = val(Leer.GetValue("NPC" & NPCNumber, "Snd2"))
    Npclist(NpcIndex).flags.Snd3 = val(Leer.GetValue("NPC" & NPCNumber, "Snd3"))

    '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

    Dim aux As String

    aux = Leer.GetValue("NPC" & NPCNumber, "NROEXP")

    If aux = "" Then
        Npclist(NpcIndex).NroExpresiones = 0
    Else
        Npclist(NpcIndex).NroExpresiones = val(aux)
        ReDim Npclist(NpcIndex).Expresiones(1 To Npclist(NpcIndex).NroExpresiones) As String

        For LoopC = 1 To Npclist(NpcIndex).NroExpresiones
            Npclist(NpcIndex).Expresiones(LoopC) = Leer.GetValue("NPC" & NPCNumber, "Exp" & LoopC)
        Next LoopC

    End If

    '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

    'Tipo de items con los que comercia
    Npclist(NpcIndex).TipoItems = val(Leer.GetValue("NPC" & NPCNumber, "TipoItems"))

    'Update contadores de NPCs
    If NpcIndex > LastNPC Then LastNPC = NpcIndex
    NumNPCs = NumNPCs + 1

    'Devuelve el nuevo Indice
    OpenNPC = NpcIndex

End Function

Sub EnviarListaCriaturas(ByVal UserIndex As Integer, ByVal NpcIndex)

    Dim SD As String

    Dim k  As Integer

    SD = SD & Npclist(NpcIndex).NroCriaturas & ","

    For k = 1 To Npclist(NpcIndex).NroCriaturas
        SD = SD & Npclist(NpcIndex).Criaturas(k).NpcName & ","
    Next k

    SD = "LSTCRI" & SD
    Call SendData(SendTarget.ToIndex, UserIndex, 0, SD)

End Sub

Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)

    Dim i As Integer

    With Npclist(NpcIndex)

        If .flags.Follow Then
            .flags.AttackedBy = vbNullString
            .flags.Follow = False
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
            .char.heading = 3
            .MaestroUser = 0
            
            Call SendData(SendTarget.ToNPCArea, UserList(NameIndex(UserName)).flags.TargetNpc, Npclist(UserList(NameIndex(UserName)).flags.TargetNpc).Pos.Map, "||" & vbWhite & "�" & "Aqui me quedo" & "�" & CStr(Npclist(UserList(NameIndex(UserName)).flags.TargetNpc).char.CharIndex))

            For i = 1 To LastUser
                Call UpdateUserMap(i)
            Next i
            
        Else

            .flags.AttackedBy = UserName
            .flags.Follow = True
            .Movement = TipoAI.SigueAmo   'follow
            .Hostile = 0
            .MaestroUser = 1
            
            Call SendData(SendTarget.ToNPCArea, UserList(NameIndex(UserName)).flags.TargetNpc, Npclist(UserList(NameIndex(UserName)).flags.TargetNpc).Pos.Map, "||" & vbWhite & "�" & "Te sigo " & UserName & "�" & CStr(Npclist(UserList(NameIndex(UserName)).flags.TargetNpc).char.CharIndex))

        End If

    End With

End Sub

Sub FollowAmo(ByVal NpcIndex As Integer)

    Npclist(NpcIndex).flags.Follow = True
    Npclist(NpcIndex).Movement = TipoAI.SigueAmo    'follow
    Npclist(NpcIndex).Hostile = 0
    Npclist(NpcIndex).Target = 0
    Npclist(NpcIndex).TargetNpc = 0

End Sub

Sub LeerNpc(ByVal num As Long, ByVal Search As String, ByVal UserIndex As Integer)

    Dim Name As String

    Name = GetVar(App.Path & "\Dat\Npcs.Dat", "NPC" & num, "Name")
    
    If Name = "" Then Exit Sub

    If Search = "" Then
        CountNpc = CountNpc + 1
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "NNHS" & num & "#" & Name)
    Else

        If InStr(LCase(Name), LCase(Search)) Then
            CountNpc = CountNpc + 1
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "NNHS" & num & "#" & Name)

        End If

    End If

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "NNHC" & CountNpc)

End Sub

Sub LeerNpcH(ByVal num As Long, ByVal Search As String, ByVal UserIndex As Integer)

    Dim Name As String

    Name = GetVar(App.Path & "\Dat\NPCs-HOSTILES.Dat", "NPC" & num, "Name")
    
    If Name = "" Then Exit Sub

    If Search = "" Then
        CountNpcH = CountNpcH + 1
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "NHHS" & num & "#" & Name)
    Else

        If InStr(LCase(Name), LCase(Search)) Then
            CountNpcH = CountNpcH + 1
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "NHHS" & num & "#" & Name)

        End If

    End If

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "NHHC" & CountNpcH)

End Sub
