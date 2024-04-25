Attribute VB_Name = "AI"
Option Explicit

Public Enum TipoAI

    ESTATICO = 1
    MueveAlAzar = 2
    NpcMaloAtacaUsersBuenos = 3
    NPCDEFENSA = 4
    GuardiasAtacanCriminales = 5
    GuardiasAtacanCiudadanos = 6
    SigueAmo = 8
    NpcAtacaNpc = 9
    NpcPathfinding = 10
    NpcSeth = 23
    NpcDragon = 25

End Enum

Public Const ELEMENTALFUEGO  As Integer = 93
Public Const ELEMENTALFUEGOII As Integer = 619
Public Const ELEMENTALTIERRA As Integer = 94
Public Const ELEMENTALTIERRAII As Integer = 620
Public Const ELEMENTALTIERRAM As Integer = 166
Public Const ELEMENTALAGUA   As Integer = 92
Public Const ELEMENTALAGUAII As Integer = 618
Public Const ELEMENTALVIENTO As Integer = 242
Public Const ELEMENTALTEMPLARIO As Integer = 693
Public Const ELEMENTALTORTUGA As Integer = 721
Public Const ELEMENTALFATUO As Integer = 89

'Damos a los NPCs el mismo rango de visión que un PJ
Public Const RANGO_VISION_X  As Byte = 10
Public Const RANGO_VISION_Y  As Byte = 10

Public Enum e_Alineacion

    ninguna = 0
    Real = 1
    Caos = 2
    Neutro = 3

End Enum

Public Enum e_Personalidad

    ''Inerte: no tiene objetivos de ningun tipo (npcs vendedores, curas, etc)
    ''Agresivo no magico: Su objetivo es acercarse a las victimas para atacarlas
    ''Agresivo magico: Su objetivo es mantenerse lo mas lejos posible de sus victimas y atacarlas con magia
    ''Mascota: Solo ataca a quien ataque a su amo.
    ''Pacifico: No ataca.
    ninguna = 0
    Inerte = 1
    AgresivoNoMagico = 2
    AgresivoMagico = 3
    Macota = 4
    Pacifico = 5

End Enum

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo AI_NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'AI de los NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Private Sub HandleAlineacion(ByVal NpcIndex As Integer)

    Dim Al          As e_Alineacion
    Dim Pe          As e_Personalidad
    Dim TargetPJ    As Integer
    Dim TargetNpc   As Integer
    Dim TieneTarget As Boolean
    Dim EsNpc       As Boolean

    TieneTarget = False
    Al = Npclist(NpcIndex).flags.AIAlineacion
    TargetPJ = Npclist(NpcIndex).flags.AtacaAPJ
    TargetNpc = Npclist(NpcIndex).flags.AtacaANPC
    
    Select Case Al

        Case e_Alineacion.Caos

            If TargetPJ > 0 Then
                If InRangoVisionNPC(NpcIndex, UserList(TargetPJ).pos.X, UserList(TargetPJ).pos.Y) Then
                    If Not Criminal(TargetPJ) Then
                        TieneTarget = True
                    Else
                        Npclist(NpcIndex).flags.AtacaAPJ = 0

                    End If

                Else
                    Npclist(NpcIndex).flags.AtacaAPJ = 0

                End If

            End If

            If TargetNpc > 0 Then
                If InRangoVisionNPC(NpcIndex, Npclist(TargetNpc).pos.X, Npclist(TargetNpc).pos.Y) Then
                    TieneTarget = True
                Else
                    Npclist(NpcIndex).flags.AtacaANPC = 0

                End If

            End If

        Case e_Alineacion.Neutro

            If TargetPJ > 0 Then
                If InRangoVisionNPC(NpcIndex, UserList(TargetPJ).pos.X, UserList(TargetPJ).pos.Y) Then
                    TieneTarget = True
                Else
                    Npclist(NpcIndex).flags.AtacaAPJ = 0

                End If

            End If

            If TargetNpc > 0 Then
                If InRangoVisionNPC(NpcIndex, Npclist(TargetNpc).pos.X, Npclist(TargetNpc).pos.Y) Then
                    TieneTarget = True
                Else
                    Npclist(NpcIndex).flags.AtacaANPC = 0

                End If

            End If

        Case e_Alineacion.ninguna
            Exit Sub

        Case e_Alineacion.Real

            If TargetPJ > 0 Then
                If InRangoVisionNPC(NpcIndex, UserList(TargetPJ).pos.X, UserList(TargetPJ).pos.Y) Then
                    If Criminal(TargetPJ) Then
                        TieneTarget = True
                    Else
                        Npclist(NpcIndex).flags.AtacaAPJ = 0

                    End If

                Else
                    Npclist(NpcIndex).flags.AtacaAPJ = 0

                End If

            End If

            If TargetNpc > 0 Then
                If InRangoVisionNPC(NpcIndex, Npclist(TargetNpc).pos.X, Npclist(TargetNpc).pos.Y) Then
                    TieneTarget = True
                Else
                    Npclist(NpcIndex).flags.AtacaANPC = 0

                End If

            End If

    End Select
    
    If Not TieneTarget Then
    
    End If

End Sub

Private Function AcquireNewTargetForAlignment(ByVal NpcIndex As Integer, ByRef EsNpc As Boolean) As Integer

    Dim r             As Byte
    Dim NPCPosX       As Byte
    Dim NPCPosY       As Byte
    Dim NpcBestTarget As Integer
    Dim PJBestTarget  As Integer
    Dim PJ            As Integer
    Dim npc           As Integer

    Dim X             As Integer
    Dim Y             As Integer
    Dim m             As Integer

    NPCPosX = Npclist(NpcIndex).pos.X
    NPCPosY = Npclist(NpcIndex).pos.Y
    m = Npclist(NpcIndex).pos.Map
    
    For r = 1 To MinYBorder
        For X = NPCPosX - r To NPCPosX + r
            For Y = NPCPosY - r To NPCPosY + r
                PJ = MapData(m, X, Y).UserIndex
                npc = MapData(m, X, Y).NpcIndex
                
                If PJ > 0 Then

                    Select Case Npclist(NpcIndex).flags.AIAlineacion

                        Case e_Alineacion.Caos

                            If Not Criminal(PJ) And Not UserList(PJ).flags.Muerto And Not UserList(PJ).flags.Invisible And Not UserList( _
                                    PJ).flags.Oculto And UserList(PJ).flags.Privilegios = PlayerType.User Then
                                PJBestTarget = PJ

                            End If

                        Case e_Alineacion.Real
                        
                        Case e_Alineacion.Neutro
                    
                    End Select
                
                End If

                If MapData(m, X, Y).NpcIndex > 0 Then
                
                End If

            Next Y
        Next X

        If PJBestTarget > 0 Then
            EsNpc = False
            AcquireNewTargetForAlignment = PJBestTarget
            Exit Function

        End If

        If NpcBestTarget > 0 Then
            EsNpc = True
            AcquireNewTargetForAlignment = NpcBestTarget
            Exit Function

        End If
        
    Next r

End Function

Private Sub GuardiasAI(ByVal NpcIndex As Integer, Optional ByVal DelCaos As Boolean = False)

    Dim nPos        As WorldPos
    Dim headingloop As Byte
    Dim tHeading    As Byte
    Dim Y           As Integer
    Dim X           As Integer
    Dim UI          As Integer

    For headingloop = eHeading.NORTH To eHeading.WEST
        nPos = Npclist(NpcIndex).pos

        If Npclist(NpcIndex).flags.Inmovilizado = 0 Or headingloop = Npclist(NpcIndex).char.Heading Then
            Call HeadtoPos(headingloop, nPos)

            If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex

                If UI > 0 Then
                    If UserList(UI).flags.Muerto = 0 Then

                        '¿ES CRIMINAL?
                        If Not DelCaos Then
                            If Criminal(UI) Then
                                If NpcAtacaUser(NpcIndex, UI) Then
                                    Call ChangeNPCChar(SendTarget.ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).char.Body, Npclist( _
                                            NpcIndex).char.Head, headingloop)

                                End If

                                Exit Sub
                            ElseIf Npclist(NpcIndex).flags.AttackedBy = UserList(UI).Name And Not Npclist(NpcIndex).flags.Follow Then
                                  
                                If NpcAtacaUser(NpcIndex, UI) Then
                                    Call ChangeNPCChar(SendTarget.ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).char.Body, Npclist( _
                                            NpcIndex).char.Head, headingloop)

                                End If

                                Exit Sub

                            End If

                        Else

                            If Not Criminal(UI) Then
                                   
                                If NpcAtacaUser(NpcIndex, UI) Then
                                    Call ChangeNPCChar(SendTarget.ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).char.Body, Npclist( _
                                            NpcIndex).char.Head, headingloop)

                                End If

                                Exit Sub
                            ElseIf Npclist(NpcIndex).flags.AttackedBy = UserList(UI).Name And Not Npclist(NpcIndex).flags.Follow Then
                                  
                                If NpcAtacaUser(NpcIndex, UI) Then
                                    Call ChangeNPCChar(SendTarget.ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).char.Body, Npclist( _
                                            NpcIndex).char.Head, headingloop)

                                End If

                                Exit Sub

                            End If

                        End If

                    End If

                End If

            End If

        End If  'not inmovil

    Next headingloop

    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub HostilMalvadoAI(ByVal NpcIndex As Integer)

    Dim nPos        As WorldPos
    Dim headingloop As Byte
    Dim tHeading    As Byte
    Dim Y           As Integer
    Dim X           As Integer
    Dim UI          As Integer
    Dim NPCI        As Integer
    Dim atacoPJ     As Boolean

    atacoPJ = False

    For headingloop = eHeading.NORTH To eHeading.WEST
        nPos = Npclist(NpcIndex).pos

        If Npclist(NpcIndex).flags.Inmovilizado = 0 Or Npclist(NpcIndex).char.Heading = headingloop Then
            Call HeadtoPos(headingloop, nPos)

            If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                NPCI = MapData(nPos.Map, nPos.X, nPos.Y).NpcIndex

                If UI > 0 And Not atacoPJ Then
                    If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Privilegios = PlayerType.User Then
                        atacoPJ = True

                        If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then
                            Call NpcLanzaUnSpell(NpcIndex, UI)

                        End If

                        If NpcAtacaUser(NpcIndex, MapData(nPos.Map, nPos.X, nPos.Y).UserIndex) Then
                            Call ChangeNPCChar(SendTarget.ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).char.Body, Npclist(NpcIndex).char.Head, _
                                    headingloop)

                        End If

                        Exit Sub

                    End If

                ElseIf NPCI > 0 Then

                    If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
                        Call ChangeNPCChar(SendTarget.ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).char.Body, Npclist(NpcIndex).char.Head, _
                                headingloop)
                      '  Call SistemaCombate.NpcAtacaNpc(NpcIndex, NPCI, False)
                        Exit Sub

                    End If

                End If

            End If

        End If  'inmo

    Next headingloop

    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub HostilBuenoAI(ByVal NpcIndex As Integer)

    Dim nPos        As WorldPos
    Dim headingloop As eHeading
    Dim tHeading    As Byte
    Dim Y           As Integer
    Dim X           As Integer
    Dim UI          As Integer

    For headingloop = eHeading.NORTH To eHeading.WEST
        nPos = Npclist(NpcIndex).pos

        If Npclist(NpcIndex).flags.Inmovilizado = 0 Or Npclist(NpcIndex).char.Heading = headingloop Then
            Call HeadtoPos(headingloop, nPos)

            If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex

                If UI > 0 Then
                    If UserList(UI).Name = Npclist(NpcIndex).flags.AttackedBy Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Privilegios = PlayerType.User Then

                            If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                Call NpcLanzaUnSpell(NpcIndex, UI)

                            End If

                            If NpcAtacaUser(NpcIndex, UI) Then
                                Call ChangeNPCChar(SendTarget.ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).char.Body, Npclist( _
                                        NpcIndex).char.Head, headingloop)

                            End If

                            Exit Sub

                        End If

                    End If

                End If

            End If

        End If

    Next headingloop

    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub IrUsuarioCercano(ByVal NpcIndex As Integer)
On Error GoTo fallo
    Dim nPos        As WorldPos
    Dim headingloop As Byte
    Dim tHeading    As Byte
    Dim Y           As Integer
    Dim X           As Integer
    Dim UI          As Integer
    Dim SignoNS     As Integer
    Dim SignoEO     As Integer
    Dim UserIndex As Integer
    
    For UserIndex = 1 To NumUsers

    If Npclist(NpcIndex).Numero = 254 And UserList(UserIndex).flags.Angel = True Then
          Exit Sub
    End If

    If Npclist(NpcIndex).Numero = 253 And UserList(UserIndex).flags.Demonio = True Then
          Exit Sub
    End If

    If UserList(UserIndex).flags.Corsarios = True Then
        If Npclist(NpcIndex).Numero = NpcCorsarios Then
          Exit Sub
        End If
    End If

    If UserList(UserIndex).flags.Piratas = True Then
        If Npclist(NpcIndex).Numero = NpcPiratas Then
          Exit Sub
        End If
    End If

    Next UserIndex

        For Y = Npclist(NpcIndex).pos.Y - RANGO_VISION_Y To Npclist(NpcIndex).pos.Y + RANGO_VISION_Y
            For X = Npclist(NpcIndex).pos.X - RANGO_VISION_X To Npclist(NpcIndex).pos.X + RANGO_VISION_X

                If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                    UI = MapData(Npclist(NpcIndex).pos.Map, X, Y).UserIndex

                    If UI > 0 Then

                       If UserList(UI).flags.Muerto = 0 And NpcVeInvi(NpcIndex) And UserList(UI).flags.Privilegios = PlayerType.User Then
                       If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                           tHeading = FindDirection(Npclist(NpcIndex).pos, UserList(MapData(Npclist( _
                             NpcIndex).pos.Map, X, Y).UserIndex).pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub
                        End If

 
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList( _
                               UI).flags.Privilegios = PlayerType.User Then

                            If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                            tHeading = FindDirection(Npclist(NpcIndex).pos, UserList(MapData(Npclist( _
                             NpcIndex).pos.Map, X, Y).UserIndex).pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                           Exit Sub
                       End If

                    End If

                End If

            Next X
       Next Y



  Call RestoreOldMovement(NpcIndex)
Exit Sub
fallo:
Call LogError("IrUsuarioCercano " & Err.Number & " D: " & Err.Description)
    
End Sub

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)

    Dim nPos        As WorldPos
    Dim headingloop As Byte
    Dim tHeading    As Byte
    Dim Y           As Integer
    Dim X           As Integer
    Dim UI          As Integer

    Dim SignoNS     As Integer
    Dim SignoEO     As Integer

    With Npclist(NpcIndex)

        If .flags.Inmovilizado = 1 Then

            Select Case .char.Heading

                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0

                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1

                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0

                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0

            End Select
    
            For Y = .pos.Y To .pos.Y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)

                For X = .pos.X To .pos.X + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)

                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        UI = MapData(.pos.Map, X, Y).UserIndex

                        If UI > 0 Then
                            If UserList(UI).Name = .flags.AttackedBy Then
                                If .MaestroUser > 0 Then
                                    If Not Criminal(.MaestroUser) And Not Criminal(UI) And (UserList(.MaestroUser).flags.Seguro Or UserList( _
                                            .MaestroUser).Faccion.ArmadaReal = 1) Then
                                        Call SendData(SendTarget.toindex, .MaestroUser, 0, _
                                                "||La mascota no atacará a ciudadanos si eres miembro de la Armada Real o tienes el seguro activado" _
                                                & FONTTYPE_INFO)
                                        .flags.AttackedBy = vbNullString
                                        Exit Sub

                                    End If

                                End If

                                If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 And _
                                        UserList(UI).flags.Privilegios = PlayerType.User Then

                                    If .flags.LanzaSpells > 0 Then
                                        Call NpcLanzaUnSpell(NpcIndex, UI)

                                    End If

                                    Exit Sub

                                End If

                            End If

                        End If

                    End If

                Next X
            Next Y

        Else

            For Y = .pos.Y - RANGO_VISION_Y To .pos.Y + RANGO_VISION_Y
                For X = .pos.X - RANGO_VISION_X To .pos.X + RANGO_VISION_X

                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        UI = MapData(.pos.Map, X, Y).UserIndex

                        If UI > 0 Then
                            If UserList(UI).Name = .flags.AttackedBy Then
                                If .MaestroUser > 0 Then
                                    If Not Criminal(.MaestroUser) And Not Criminal(UI) And (UserList(.MaestroUser).flags.Seguro Or UserList( _
                                            .MaestroUser).Faccion.ArmadaReal = 1) Then
                                        Call SendData(SendTarget.toindex, .MaestroUser, 0, _
                                                "||La mascota no atacará a ciudadanos si eres miembro de la Armada Real o tienes el seguro activado" _
                                                & FONTTYPE_INFO)
                                        .flags.AttackedBy = vbNullString
                                        Call FollowAmo(NpcIndex)
                                        Exit Sub

                                    End If

                                End If

                                If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 Then

                                    If .flags.LanzaSpells > 0 Then
                                        Call NpcLanzaUnSpell(NpcIndex, UI)

                                    End If

                                    tHeading = FindDirection(.pos, UserList(UI).pos)
                                    'Call MoveNPCChar(NpcIndex, tHeading)
                                    Exit Sub
                                    
                                    If Not Npclist(NpcIndex).PFINFO.PathLenght > 0 Then tHeading = FindDirection(Npclist(NpcIndex).pos, UserList( _
                                            MapData(Npclist(NpcIndex).pos.Map, X, Y).UserIndex).pos)

                                    If tHeading = 0 Then
                                        If ReCalculatePath(NpcIndex) Then
                                            Call PathFindingAI(NpcIndex)

                                            'Existe el camino?
                                            If Npclist(NpcIndex).PFINFO.NoPath Then 'Si no existe nos movemos al azar
                                                'Move randomly
                                                Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))

                                            End If

                                        Else

                                            If Not PathEnd(NpcIndex) Then
                                                Call FollowPath(NpcIndex)
                                            Else
                                                Npclist(NpcIndex).PFINFO.PathLenght = 0

                                            End If

                                        End If

                                    Else

                                        If Not Npclist(NpcIndex).PFINFO.PathLenght > 0 Then Call MoveNPCChar(NpcIndex, tHeading)
                                        Exit Sub

                                    End If
                            
                                    Exit Sub

                                End If

                            End If

                        End If

                    End If

                Next X
            Next Y

        End If

        Call RestoreOldMovement(NpcIndex)

    End With

End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)

    If Npclist(NpcIndex).MaestroUser = 0 Then
        Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
        Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
        Npclist(NpcIndex).flags.AttackedBy = ""

    End If

End Sub

Private Sub PersigueCiudadano(ByVal NpcIndex As Integer)
 On Error GoTo fallo
Dim UI As Integer
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
For Y = Npclist(NpcIndex).pos.Y - RANGO_VISION_Y To Npclist(NpcIndex).pos.Y + RANGO_VISION_Y
        For X = Npclist(NpcIndex).pos.X - RANGO_VISION_X To Npclist(NpcIndex).pos.X + RANGO_VISION_X

        If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
           UI = MapData(Npclist(NpcIndex).pos.Map, X, Y).UserIndex
             If UI > 0 Then
                If Not Criminal(UI) And UserList(UI).flags.Privilegios = 0 Then
                   If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                              Dim k As Integer
                              k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                              Call NpcLanzaUnSpell(NpcIndex, UI)
                        End If
                   If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminInvisible = 0 And UserList(UI).flags.Privilegios = 0 Then
                        
                       Call MoveNPCChar(NpcIndex, FindDirectionEAO(Npclist(NpcIndex).pos, UserList(UI).pos, (Npclist(NpcIndex).flags.AguaValida)))
                           
                        Exit Sub
                   End If
                End If
           End If
        End If
    Next X
Next Y

If Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then
    Call AI_Volver(NpcIndex)
End If
Call RestoreOldMovement(NpcIndex)
Exit Sub
fallo:
Call LogError("PERSIGUECIUDADANO " & Err.Number & " D: " & Err.Description)

End Sub

Private Sub PersigueCriminal(ByVal NpcIndex As Integer)
 On Error GoTo fallo
Dim UI As Integer
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
    For Y = Npclist(NpcIndex).pos.Y - RANGO_VISION_Y To Npclist(NpcIndex).pos.Y + RANGO_VISION_Y
        For X = Npclist(NpcIndex).pos.X - RANGO_VISION_X To Npclist(NpcIndex).pos.X + RANGO_VISION_X

       If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
           UI = MapData(Npclist(NpcIndex).pos.Map, X, Y).UserIndex
           If UI > 0 Then
                If Criminal(UI) And UserList(UI).flags.Privilegios = PlayerType.User Then
            
                   If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                              Dim k As Integer
                              k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                              Call NpcLanzaUnSpell(NpcIndex, UI)
                        End If
                   If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.AdminInvisible = 0 And UserList(UI).flags.Privilegios = 0 Then
                
                          Call MoveNPCChar(NpcIndex, FindDirectionEAO(Npclist(NpcIndex).pos, UserList(UI).pos, (Npclist(NpcIndex).flags.AguaValida)))
                           Exit Sub
                   End If
                End If
           End If
           
        End If
    Next X
Next Y

If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
    Call AI_Volver(NpcIndex)
End If
Call RestoreOldMovement(NpcIndex)
Exit Sub
fallo:
Call LogError("PERSIGUECRIMINAL " & Err.Number & " D: " & Err.Description)

End Sub

Private Sub NpcDragonAI(ByVal NpcIndex As Integer)

       
    Dim nPos        As WorldPos
    Dim headingloop As Byte
    Dim tHeading    As Byte
    Dim Y           As Integer
    Dim X           As Integer
    Dim UI          As Integer
    
    For Y = Npclist(NpcIndex).pos.Y - RANGO_VISION_Y To Npclist(NpcIndex).pos.Y + RANGO_VISION_Y
        For X = Npclist(NpcIndex).pos.X - RANGO_VISION_X To Npclist(NpcIndex).pos.X + RANGO_VISION_X

            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
              
                    UI = MapData(Npclist(NpcIndex).pos.Map, X, Y).UserIndex

                    If UI > 0 Then
                    
                       If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Privilegios = PlayerType.User Then
                            If RandomNumber(1, 10) < 5 Then
                                If RandomNumber(1, 10) < 3 Then
                                    Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))
                                End If
                                Else
                                If RandomNumber(1, 10) > 3 Then
                                    Call MoveNPCChar(NpcIndex, FindDirectionEAO(Npclist(NpcIndex).pos, UserList(UI).pos, (Npclist(NpcIndex).flags.AguaValida)))
                                End If
                            End If
                            Exit Sub
                        ElseIf UserList(UI).flags.Muerto = 1 And UserList(UI).flags.Privilegios = PlayerType.User Then
                            If RandomNumber(1, 10) < 3 Then
                                Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))
                            End If
                            Exit Sub
                       End If
                   
                    
          End If
        End If
    
      Next X
    Next Y
                       If RandomNumber(1, 10) < 3 Then
                                Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))
                       End If
    
       
End Sub


Private Sub SeguirAmo(ByVal NpcIndex As Integer)

    Dim nPos        As WorldPos
    Dim headingloop As Byte
    Dim tHeading    As Byte
    Dim Y           As Integer
    Dim X           As Integer
    Dim UI          As Integer

    For Y = Npclist(NpcIndex).pos.Y - RANGO_VISION_Y To Npclist(NpcIndex).pos.Y + RANGO_VISION_Y
        For X = Npclist(NpcIndex).pos.X - RANGO_VISION_X To Npclist(NpcIndex).pos.X + RANGO_VISION_X

            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                If Npclist(NpcIndex).Target = 0 And Npclist(NpcIndex).TargetNpc = 0 Then
                    UI = MapData(Npclist(NpcIndex).pos.Map, X, Y).UserIndex

                    If UI > 0 Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 And UI = Npclist( _
                                NpcIndex).MaestroUser And Distancia(Npclist(NpcIndex).pos, UserList(UI).pos) > 3 Then
                            tHeading = FindDirection(Npclist(NpcIndex).pos, UserList(MapData(Npclist( _
                             NpcIndex).pos.Map, X, Y).UserIndex).pos)

                            'Call MoveNPCChar(NpcIndex, tHeading)
                            If Not Npclist(NpcIndex).PFINFO.PathLenght > 0 Then tHeading = FindDirection(Npclist(NpcIndex).pos, UserList(MapData( _
                                    Npclist(NpcIndex).pos.Map, X, Y).UserIndex).pos)

                            If tHeading = 0 Then
                                If ReCalculatePath(NpcIndex) Then
                                    Call PathFindingAI(NpcIndex)

                                    'Existe el camino?
                                    If Npclist(NpcIndex).PFINFO.NoPath Then 'Si no existe nos movemos al azar
                                        'Move randomly
                                        Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))

                                    End If

                                Else

                                    If Not PathEnd(NpcIndex) Then
                                        Call FollowPath(NpcIndex)
                                    Else
                                        Npclist(NpcIndex).PFINFO.PathLenght = 0

                                    End If

                                End If

                            Else

                                If Not Npclist(NpcIndex).PFINFO.PathLenght > 0 Then Call MoveNPCChar(NpcIndex, tHeading)
                                Exit Sub

                            End If
                            
                            Exit Sub

                        End If

                    End If

                End If

            End If

        Next X
    Next Y

    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub AiNpcAtacaNpc(ByVal NpcIndex As Integer)

    Dim tHeading As Byte
    Dim Y        As Integer
    Dim X        As Integer
    Dim NI       As Integer
    Dim bNoEsta  As Boolean

    Dim SignoNS  As Integer
    Dim SignoEO  As Integer

    With Npclist(NpcIndex)

        If .flags.Inmovilizado = 1 Then

            Select Case .char.Heading

                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0

                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1

                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0

                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0

            End Select
    
            For Y = .pos.Y To .pos.Y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
                For X = .pos.X To .pos.X + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)

                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        NI = MapData(.pos.Map, X, Y).NpcIndex

                        If NI > 0 Then
                            If .TargetNpc = NI Then
                                bNoEsta = True

                                If .Numero = ELEMENTALFUEGO Then
                                    Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)

                                    If Npclist(NI).NPCtype = eNPCType.dragon Then
                                        Npclist(NI).CanAttack = 1
                                        Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)

                                    End If

                                Else
                                    'aca verificamosss la distancia de ataque
                            
                                   Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
                            
                                End If

                                Exit Sub

                            End If

                        End If

                    End If

                Next X
            Next Y

        Else

            For Y = .pos.Y - RANGO_VISION_Y To .pos.Y + RANGO_VISION_Y
                For X = .pos.X - RANGO_VISION_X To .pos.X + RANGO_VISION_X

                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        NI = MapData(.pos.Map, X, Y).NpcIndex

                        If NI > 0 Then
                            If .TargetNpc = NI Then
                                bNoEsta = True

                                If .Numero = ELEMENTALFUEGO Then
                                    Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)

                                    If Npclist(NI).NPCtype = eNPCType.dragon Then
                                        Npclist(NI).CanAttack = 1
                                        Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)

                                    End If

                                Else
                                    'aca verificamosss la distancia de ataque
                            
                                    Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
                     
                                End If

                                If .flags.Inmovilizado = 1 Then Exit Sub
                                If .TargetNpc = 0 Then Exit Sub
                                tHeading = FindDirection(.pos, Npclist(MapData(Npclist( _
                                 NpcIndex).pos.Map, X, Y).NpcIndex).pos)

                              ' Call MoveNPCChar(NpcIndex, tHeading)
                                If Not .PFINFO.PathLenght > 0 Then
                                    tHeading = FindDirection(.pos, Npclist(MapData(.pos.Map, X, Y).NpcIndex).pos)

                                End If

                                If tHeading = 0 Then
                                    If ReCalculatePath(NpcIndex) Then
                                        Call PathFindingAI(NpcIndex)

                                        'Existe el camino?
                                        If .PFINFO.NoPath Then 'Si no existe nos movemos al azar
                                            'Move randomly
                                            Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))

                                        End If

                                    Else

                                        If Not PathEnd(NpcIndex) Then
                                            Call FollowPath(NpcIndex)
                                        Else
                                            .PFINFO.PathLenght = 0

                                        End If

                                    End If

                                Else

                                    If Not .PFINFO.PathLenght > 0 Then Call MoveNPCChar(NpcIndex, tHeading)
                                    Exit Sub

                                End If
                            
                                Exit Sub

                            End If

                        End If

                    End If

                Next X
            Next Y

        End If

        If Not bNoEsta Then
            If .MaestroUser > 0 Then
                Call FollowAmo(NpcIndex)
            Else
                .Movement = .flags.OldMovement
                .Hostile = .flags.OldHostil

            End If

        End If

    End With

End Sub

Function NPCAI(ByVal NpcIndex As Integer)

    On Error GoTo ErrorHandler

    '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
    If Npclist(NpcIndex).MaestroUser = 0 Then

        'Busca a alguien para atacar
        '¿Es un guardia?
        If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            Call GuardiasAI(NpcIndex)
        ElseIf Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then
            Call GuardiasAI(NpcIndex, True)
        ElseIf Npclist(NpcIndex).Hostile And Npclist(NpcIndex).Stats.Alineacion <> 0 Then
            Call HostilMalvadoAI(NpcIndex)
        ElseIf Npclist(NpcIndex).Hostile And Npclist(NpcIndex).Stats.Alineacion = 0 Then
            Call HostilBuenoAI(NpcIndex)

        End If

    Else

        If False Then Exit Function

        'Evitamos que ataque a su amo, a menos
        'que el amo lo ataque.
        'Call HostilBuenoAI(NpcIndex)
    End If
        
    '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
    Select Case Npclist(NpcIndex).Movement

        Case TipoAI.MueveAlAzar

            If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Function
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                If RandomNumber(1, 12) = 3 Then
                    Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                End If

                Call PersigueCriminal(NpcIndex)
            ElseIf Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then

                If RandomNumber(1, 12) = 3 Then
                    Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))

                End If

                Call PersigueCiudadano(NpcIndex)
            Else

                If RandomNumber(1, 12) = 3 Then
                    Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))

                End If

            End If

            'Va hacia el usuario cercano
        Case TipoAI.NpcMaloAtacaUsersBuenos
            Call IrUsuarioCercano(NpcIndex)

            'Va hacia el usuario que lo ataco(FOLLOW)
        Case TipoAI.NPCDEFENSA
            Call SeguirAgresor(NpcIndex)

            'Persigue criminales
        Case TipoAI.GuardiasAtacanCriminales
                Call PersigueCriminal(NpcIndex)
         
         Case TipoAI.GuardiasAtacanCiudadanos
                Call PersigueCiudadano(NpcIndex)

        Case TipoAI.SigueAmo

            If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Function
            Call SeguirAmo(NpcIndex)

            If RandomNumber(1, 12) = 3 Then
                Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))

            End If

        Case TipoAI.NpcAtacaNpc
            Call AiNpcAtacaNpc(NpcIndex)

        Case TipoAI.NpcPathfinding

            If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Function
            If ReCalculatePath(NpcIndex) Then
                Call PathFindingAI(NpcIndex)

                'Existe el camino?
                If Npclist(NpcIndex).PFINFO.NoPath Then 'Si no existe nos movemos al azar
                    'Move randomly
                    Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))

                End If

            Else

                If Not PathEnd(NpcIndex) Then
                    Call FollowPath(NpcIndex)
                Else
                    Npclist(NpcIndex).PFINFO.PathLenght = 0

                End If

            End If
        
       Case TipoAI.NpcSeth
            If RandomNumber(1, 12) = 3 Then
                    Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
            End If
            
        Case TipoAI.NpcDragon
            Call NpcDragonAI(NpcIndex)

    End Select

    Exit Function

ErrorHandler:
    Call LogError("NPCAI " & Npclist(NpcIndex).Name & " " & Npclist(NpcIndex).MaestroUser & " " & Npclist(NpcIndex).MaestroNpc & " mapa:" & Npclist( _
            NpcIndex).pos.Map & " x:" & Npclist(NpcIndex).pos.X & " y:" & Npclist(NpcIndex).pos.Y & " Mov:" & Npclist(NpcIndex).Movement & _
            " TargU:" & Npclist(NpcIndex).Target & " TargN:" & Npclist(NpcIndex).TargetNpc)
    Dim MiNPC As npc
    MiNPC = Npclist(NpcIndex)
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)
    
End Function

Function UserNear(ByVal NpcIndex As Integer) As Boolean
    '#################################################################
    'Returns True if there is an user adjacent to the npc position.
    '#################################################################
    UserNear = Not Int(Distance(Npclist(NpcIndex).pos.X, Npclist(NpcIndex).pos.Y, UserList(Npclist(NpcIndex).PFINFO.TargetUser).pos.X, UserList( _
            Npclist(NpcIndex).PFINFO.TargetUser).pos.Y)) > 1

End Function

Function ReCalculatePath(ByVal NpcIndex As Integer) As Boolean

    '#################################################################
    'Returns true if we have to seek a new path
    '#################################################################
    If Npclist(NpcIndex).PFINFO.PathLenght = 0 Then
        ReCalculatePath = True
    ElseIf Not UserNear(NpcIndex) And Npclist(NpcIndex).PFINFO.PathLenght = Npclist(NpcIndex).PFINFO.CurPos - 1 Then
        ReCalculatePath = True

    End If

End Function

Function SimpleAI(ByVal NpcIndex As Integer) As Boolean
    '#################################################################
    'Old Ore4 AI function
    '#################################################################
    Dim nPos        As WorldPos
    Dim headingloop As Byte
    Dim tHeading    As Byte
    Dim Y           As Integer
    Dim X           As Integer

    For Y = Npclist(NpcIndex).pos.Y - RANGO_VISION_Y To Npclist(NpcIndex).pos.Y + RANGO_VISION_Y    'Makes a loop that looks at
        For X = Npclist(NpcIndex).pos.X - RANGO_VISION_X To Npclist(NpcIndex).pos.X + RANGO_VISION_X   '5 tiles in every direction

            'Make sure tile is legal
            If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then

                'look for a user
                If MapData(Npclist(NpcIndex).pos.Map, X, Y).UserIndex > 0 Then
                    'Move towards user
                    tHeading = FindDirection(Npclist(NpcIndex).pos, UserList(MapData(Npclist( _
                     NpcIndex).pos.Map, X, Y).UserIndex).pos)

                    'MoveNPCChar NpcIndex, tHeading
                    'Leave
                    If Not Npclist(NpcIndex).PFINFO.PathLenght > 0 Then tHeading = FindDirection(Npclist(NpcIndex).pos, UserList(MapData(Npclist( _
                            NpcIndex).pos.Map, X, Y).UserIndex).pos)

                    If tHeading = 0 Then
                        If ReCalculatePath(NpcIndex) Then
                            Call PathFindingAI(NpcIndex)

                            'Existe el camino?
                            If Npclist(NpcIndex).PFINFO.NoPath Then 'Si no existe nos movemos al azar
                                'Move randomly
                                Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))

                            End If

                        Else

                            If Not PathEnd(NpcIndex) Then
                                Call FollowPath(NpcIndex)
                            Else
                                Npclist(NpcIndex).PFINFO.PathLenght = 0

                            End If

                        End If

                    Else

                        If Not Npclist(NpcIndex).PFINFO.PathLenght > 0 Then Call MoveNPCChar(NpcIndex, tHeading)
                        Exit Function

                    End If
                           
                    Exit Function

                End If

            End If

        Next X
    Next Y

End Function

Function PathEnd(ByVal NpcIndex As Integer) As Boolean
    '#################################################################
    'Coded By Gulfas Morgolock
    'Returns if the npc has arrived to the end of its path
    '#################################################################
    PathEnd = Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.PathLenght

End Function

Function FollowPath(ByVal NpcIndex As Integer) As Boolean
    '#################################################################
    'Coded By Gulfas Morgolock
    'Moves the npc.
    '#################################################################

    Dim tmpPos   As WorldPos
    Dim tHeading As Byte

    tmpPos.Map = Npclist(NpcIndex).pos.Map
    tmpPos.X = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).Y ' invertí las coordenadas
    tmpPos.Y = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).X

    'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"

    tHeading = FindDirection(Npclist(NpcIndex).pos, tmpPos)
    MoveNPCChar NpcIndex, tHeading

    Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.CurPos + 1

End Function

Function PathFindingAI(ByVal NpcIndex As Integer) As Boolean
 On Error GoTo fallo

Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer

For Y = Npclist(NpcIndex).pos.Y - RANGO_VISION_Y To Npclist(NpcIndex).pos.Y + RANGO_VISION_Y    'Makes a loop that looks at
     For X = Npclist(NpcIndex).pos.X - RANGO_VISION_X To Npclist(NpcIndex).pos.X + RANGO_VISION_X   '5 tiles in every direction

         'Make sure tile is legal
         If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
         
             'look for a user
             If MapData(Npclist(NpcIndex).pos.Map, X, Y).UserIndex > 0 Then
                 'pluto:2.11
                If UserList(MapData(Npclist(NpcIndex).pos.Map, X, Y).UserIndex).flags.Privilegios > 0 Or UserList(MapData(Npclist(NpcIndex).pos.Map, X, Y).UserIndex).flags.Muerto > 0 Then GoTo yop
                 
                 'Move towards user
                  Dim tmpUserIndex As Integer
                  tmpUserIndex = MapData(Npclist(NpcIndex).pos.Map, X, Y).UserIndex
                  
                  'We have to invert the coordinates, this is because
                  'ORE refers to maps in converse way of my pathfinding
                  'routines.
                  Npclist(NpcIndex).PFINFO.Target.X = UserList(tmpUserIndex).pos.Y
                  Npclist(NpcIndex).PFINFO.Target.Y = UserList(tmpUserIndex).pos.X 'ops!
                  Npclist(NpcIndex).PFINFO.TargetUser = tmpUserIndex
                                      'pluto:2.10
                   

                   If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, tmpUserIndex)
                   

                  Call SeekPath(NpcIndex)
                  Exit Function
             End If
             
         End If
yop:
     Next X
 Next Y

Exit Function
fallo:
Call LogError("PATHFINDINGAI " & Err.Number & " D: " & Err.Description)


End Function
Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

    
         If NpcVeInvi(NpcIndex) Then
         
         Else
         If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Or Not UserList(UserIndex).flags.Privilegios = _
            PlayerType.User Then Exit Sub
         End If

    Dim k As Integer
    k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    If Npclist(NpcIndex).Spells(k) > 0 Then
    Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, Npclist(NpcIndex).Spells(k))
    End If

End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNpc As Integer)

    Dim k As Integer
    k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNpc, Npclist(NpcIndex).Spells(k))

End Sub

Function NpcVeInvi(ByVal NpcIndex As Integer) As Boolean

       Select Case Npclist(NpcIndex).Numero
                   Case "657", "663", "645", "654", "562", "590", "584", "672", "569", "577", "578", "674"
                   NpcVeInvi = True
                   Exit Function
       End Select
       
       
       NpcVeInvi = False
       
End Function

Public Sub AI_Volver(ByVal NpcIndex As Integer)

    With Npclist(NpcIndex)
         
         If .pos.X = .Orig.X And .pos.Y = .Orig.Y Then
             .char.Heading = eHeading.SOUTH
            Call SendToNpcArea(NpcIndex, "*" & Npclist(NpcIndex).char.CharIndex & "," & .pos.X & "," & .pos.Y)
             Exit Sub
        End If
        
    
         Call MoveNPCChar(NpcIndex, FindDirectionEAO(.pos, .Orig, (Npclist(NpcIndex).flags.AguaValida)))
                    
        
    End With
    
End Sub
