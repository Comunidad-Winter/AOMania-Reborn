Attribute VB_Name = "modHechizos"
Option Explicit

Public Const HELEMENTAL_FUEGO As Integer = 26
Public Const HELEMENTAL_TIERRA As Integer = 28
Public Const SUPERANILLO As Integer = 700

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, _
                           ByVal UserIndex As Integer, _
                           ByVal Spell As Integer)

    Dim AmuletoDaño As Integer

    If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
    If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
    If UserList(UserIndex).flags.Invisible = 1 And Npclist(NpcIndex).flags.Magiainvisible = 0 Then Exit Sub
    Npclist(NpcIndex).CanAttack = 0

    Dim Daño As Integer

    If Hechizos(Spell).SubeHP = 1 Then

        Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + Daño

        If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha quitado " & Daño & " puntos de vida." & FONTTYPE_Motd4)
        Call EnviarHP(val(UserIndex))

    ElseIf Hechizos(Spell).SubeHP = 2 Then

        If UserList(UserIndex).flags.Privilegios = PlayerType.User Then

            Daño = Daño - Porcentaje(Daño, Int(((UserList(UserIndex).Stats.UserSkills(Resistencia) + 1) / 4) + ResistenciaClase(UserList(UserIndex).Clase)))
            Call SubirSkill(UserIndex, Resistencia)

            Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)

            If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                Daño = Daño - RandomNumber(ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMax)

            End If

            If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                Daño = Daño - RandomNumber(ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).DefensaMagicaMax)

            End If

            If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
                Daño = Daño - RandomNumber(ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)

            End If

            If UserList(UserIndex).Invent.AmuletoEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.AmuletoEqpObjIndex).AmuletoDefensa.TipoBonifica = eAmuleto.otMagia Then

                    AmuletoDaño = RandomNumber(1, ObjData(UserList(UserIndex).Invent.AmuletoEqpObjIndex).AmuletoDefensa.Bonifica)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu Amuleto te ha protegido de " & AmuletoDaño & " puntos de Daño." & FONTTYPE_TALKMSG)
                    Daño = Daño - AmuletoDaño

                End If

            End If

            If Daño < 0 Then Daño = 0

            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & Hechizos(Spell).WAV)
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - Daño

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha quitado " & Daño & " puntos de vida." & FONTTYPE_Motd4)
            Call EnviarHP(val(UserIndex))

            'Muere
            If UserList(UserIndex).Stats.MinHP < 1 Then
                UserList(UserIndex).Stats.MinHP = 0

                If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                    RestarCriminalidad (UserIndex)

                End If

                MuereSpell = Hechizos(Spell).FXgrh
                LoopSpell = Hechizos(Spell).loops

                Call UserDie(UserIndex)

                '[Barrin 1-12-03]
                If Npclist(NpcIndex).MaestroUser > 0 Then
                    Call ContarMuerte(UserIndex, Npclist(NpcIndex).MaestroUser)
                    Call ActStats(UserIndex, Npclist(NpcIndex).MaestroUser)

                End If

                '[/Barrin]
            End If

        End If

    End If

    If Hechizos(Spell).Paraliza = 1 Then
        If UserList(UserIndex).flags.Paralizado = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & Hechizos(Spell).WAV)
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

            'If UserList(UserIndex).Invent.HerramientaEqpObjIndex = SUPERANILLO Then
            '   Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Tu anillo rechaza los efectos del hechizo." & FONTTYPE_Motd4)
            '   Exit Sub
            'End If

            UserList(UserIndex).flags.Paralizado = 1
            UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "PARADOW")
            
            If UCase(Hechizos(Spell).nombre) = "PARALIZAR" Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "PARADO2")
            End If
            
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).pos.X & "," & UserList(UserIndex).pos.Y)
            Call Corr_ActualizarPosicion(UserIndex, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y)
        End If

    End If
    
     If Hechizos(Spell).Ceguera = 1 Then
        
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & Hechizos(Spell).WAV)
    
        If UCase$(UserList(UserIndex).Clase) <> "BARDO" And UserList(UserIndex).flags.Angel = False And UserList(UserIndex).flags.Demonio = False Then
            
            UserList(UserIndex).flags.Ceguera = 1
            UserList(UserIndex).Counters.Ceguera = IntervaloCeguera / 4

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "CEGU")
            
        Else
          
            Call SendData(ToIndex, UserIndex, 0, "|| " & Npclist(NpcIndex).Name & " te ha intentado cegar, pero eres INMUNE!!" & FONTTYPE_FIGHT)

        End If

    End If

    If Hechizos(Spell).Estupidez = 1 Then
        
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & Hechizos(Spell).WAV)
        
        If UCase$(UserList(UserIndex).Clase) <> "BARDO" And UserList(UserIndex).flags.Angel = False And UserList(UserIndex).flags.Demonio = False Then
            
            UserList(UserIndex).flags.Estupidez = 1
            UserList(UserIndex).Counters.Ceguera = IntervaloCeguera

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "DUMB")

            
        Else
        
            Call SendData(ToIndex, UserIndex, 0, "|| " & Npclist(NpcIndex).Name & " te ha intentado volver estúpido, pero eres INMUNE!!" & FONTTYPE_FIGHT)
        
        End If

    End If

End Sub

Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNpc As Integer, ByVal Spell As Integer)
'solo hechizos ofensivos!

    If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
    Npclist(NpcIndex).CanAttack = 0

    Dim Daño As Integer

    If Hechizos(Spell).SubeHP = 2 Then

        Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call SendData(SendTarget.ToNPCArea, TargetNpc, Npclist(TargetNpc).pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(SendTarget.ToNPCArea, TargetNpc, Npclist(TargetNpc).pos.Map, "CFX" & Npclist(TargetNpc).char.CharIndex & "," & Hechizos( _
                                                                                   Spell).FXgrh & "," & Hechizos(Spell).loops)

        Npclist(TargetNpc).Stats.MinHP = Npclist(TargetNpc).Stats.MinHP - Daño

        'Muere
        If Npclist(TargetNpc).Stats.MinHP < 1 Then
            Npclist(TargetNpc).Stats.MinHP = 0

            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call MuereNpc(TargetNpc, Npclist(NpcIndex).MaestroUser)
            Else
                Call MuereNpc(TargetNpc, 0)

            End If

        End If

    End If

End Sub

Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean

    On Error GoTo errhandler

    Dim j As Integer

    For j = 1 To MAXUSERHECHIZOS

        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function

        End If

    Next

    Exit Function
errhandler:

End Function

Sub AgregarHechizoEspecial(ByVal UserIndex As Integer, ByVal hIndex As Integer)
    Dim j As Integer

    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j

    If Not TieneHechizo(hIndex, UserIndex) Then
        If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes espacio para mas hechizos." & FONTTYPE_INFO)
        Else
            UserList(UserIndex).Stats.UserHechizos(j) = hIndex
            Call UpdateUserHechizos(False, UserIndex, CByte(j))
        End If

    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya tienes ese hechizo." & FONTTYPE_INFO)
    End If

End Sub


Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
    Dim hIndex As Integer
    Dim j As Integer
    hIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).HechizoIndex

    If Not TieneHechizo(hIndex, UserIndex) Then

        'Buscamos un slot vacio
        For j = 1 To MAXUSERHECHIZOS

            If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
        Next j

        If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes espacio para mas hechizos." & FONTTYPE_INFO)
        Else
            UserList(UserIndex).Stats.UserHechizos(j) = hIndex
            Call UpdateUserHechizos(False, UserIndex, CByte(j))
            'Quitamos del inv el item
            Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)

        End If

    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya tienes ese hechizo." & FONTTYPE_INFO)
    End If

End Sub

Sub DecirPalabrasMagicas(ByVal s As String, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim ind As String
    ind = UserList(UserIndex).char.CharIndex

    If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbCyan & "°" & s & "°" & ind)
        Exit Sub
    End If

    If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbRed & "°" & s & "°" & ind)
        Exit Sub
    End If

    If UserList(UserIndex).Faccion.Templario = 1 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbWhite & "°" & s & "°" & ind)
        Exit Sub
    End If

    If UserList(UserIndex).Faccion.Nemesis = 1 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & "&H808080" & "°" & s & "°" & ind)
        Exit Sub
    End If


    If Criminal(UserIndex) Then
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbRed & "°" & s & "°" & ind)
        Exit Sub
    End If

    If Not Criminal(UserIndex) Then
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbCyan & "°" & s & "°" & ind)
        Exit Sub
    End If

End Sub

Function ClasePuedeLanzarHechizo(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean
    On Error GoTo manejador


    Dim i As Integer

    For i = 1 To NUMCLASES

        If Hechizos(HechizoIndex).ClaseProhibida(i) = Chr(34) & UCase$(UserList(UserIndex).Clase) & Chr(34) Then
            ClasePuedeLanzarHechizo = False
            Exit Function
        End If

    Next i


    ClasePuedeLanzarHechizo = True
    Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarHechizo")
End Function

Function FaccionPuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean


    If Hechizos(HechizoIndex).Real = 1 Then
        FaccionPuedeLanzar = (UserList(UserIndex).Faccion.ArmadaReal = 1)
        Exit Function

    ElseIf Hechizos(HechizoIndex).Caos = 1 Then
        FaccionPuedeLanzar = (UserList(UserIndex).Faccion.FuerzasCaos = 1)
        Exit Function

    ElseIf Hechizos(HechizoIndex).Nemes = 1 Then
        FaccionPuedeLanzar = (UserList(UserIndex).Faccion.Nemesis = 1)
        Exit Function

    ElseIf Hechizos(HechizoIndex).Templ = 1 Then
        FaccionPuedeLanzar = (UserList(UserIndex).Faccion.Templario = 1)
        Exit Function

    Else
        FaccionPuedeLanzar = False

    End If

End Function


Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean

    If UserList(UserIndex).flags.Privilegios = PlayerType.User Then

        If HechizoIndex = "34" Or HechizoIndex = "67" Or HechizoIndex = "56" Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu no eres GameMaster." & FONTTYPE_INFO)
            Exit Function
        End If

        If Hechizos(HechizoIndex).Real = 1 Or Hechizos(HechizoIndex).Caos = 1 Or Hechizos(HechizoIndex).Nemes = 1 Or Hechizos(HechizoIndex).Real = 1 Then
            If Not FaccionPuedeLanzar(UserIndex, HechizoIndex) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu faccion no puede lanzar este hechizo." & FONTTYPE_INFO)
                Exit Function
            End If
        End If


        If LenB(UCase(Hechizos(HechizoIndex).ExclusivoClase)) > 0 And UCase(Hechizos(HechizoIndex).ExclusivoClase) <> UCase(UserList(UserIndex).Clase) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tú clase no puede lanzar este hechizo." & FONTTYPE_INFO)
            PuedeLanzar = False
            Exit Function
        End If

        If LenB(UCase(Hechizos(HechizoIndex).ExclusivoClase)) = 0 And ClasePuedeLanzarHechizo(UserIndex, HechizoIndex) = False Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tú clase no puede lanzar este hechizo." & FONTTYPE_INFO)
            PuedeLanzar = False
            Exit Function
        End If

    End If


    If UserList(UserIndex).flags.Muerto = 0 Then
        Dim wp2 As WorldPos
        wp2.Map = UserList(UserIndex).flags.TargetMap
        wp2.X = UserList(UserIndex).flags.TargetX
        wp2.Y = UserList(UserIndex).flags.TargetY

        If Hechizos(HechizoIndex).NeedStaff > 0 Then
            If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, _
                                      "||Tu Báculo no es lo suficientemente poderoso para que puedas lanzar el conjuro." & FONTTYPE_INFO)
                        PuedeLanzar = False
                        Exit Function
                    End If

                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes lanzar este conjuro sin la ayuda de un báculo." & FONTTYPE_INFO)
                    PuedeLanzar = False
                    Exit Function

                End If

            End If

        End If

        If Hechizos(HechizoIndex).NeedStaff > 0 Then
            If UCase$(UserList(UserIndex).Clase) = "CLERIGO" Then
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu espada no es lo suficientemente fuerte para lanzar este hechizo." & _
                                                                        FONTTYPE_INFO)
                        PuedeLanzar = False
                        Exit Function

                    End If

                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes lanzar este conjuro sin la ayuda de una Espada Argentum." & _
                                                                    FONTTYPE_INFO)
                    PuedeLanzar = False
                    Exit Function

                End If

            End If

        End If

        If UserList(UserIndex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
            If UserList(UserIndex).Stats.UserSkills(eSkill.Magia) >= Hechizos(HechizoIndex).MinSkill Then
                If UserList(UserIndex).Stats.MinSta >= Hechizos(HechizoIndex).StaRequerido Then
                    PuedeLanzar = True
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z1")
                    PuedeLanzar = False

                End If

            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z2")
                PuedeLanzar = False

            End If

        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z3")
            PuedeLanzar = False

        End If

    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z4")
        PuedeLanzar = False

    End If

End Function

Function ResistenciaClase(Clase As String) As Integer
   Dim Cuan As Integer

    Select Case UCase$(Clase)

        Case "MAGO"
            Cuan = 6

        Case "BRUJO"
            Cuan = 7

        Case "DRUIDA"
            Cuan = 6 '2

        Case "CLERIGO"
            Cuan = 1

        Case "BARDO"
            Cuan = 5

        Case Else
            Cuan = 0

    End Select

    ResistenciaClase = Cuan

End Function

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)
    Dim PosCasteadaX As Integer
    Dim PosCasteadaY As Integer
    Dim PosCasteadaM As Integer
    Dim h As Integer
    Dim TempX As Integer
    Dim TempY As Integer

    PosCasteadaX = UserList(UserIndex).flags.TargetX
    PosCasteadaY = UserList(UserIndex).flags.TargetY
    PosCasteadaM = UserList(UserIndex).flags.TargetMap

    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

    If Hechizos(h).RemueveInvisibilidadParcial = 1 Then
        b = True

        For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
            For TempY = PosCasteadaY - 8 To PosCasteadaY + 8

                If InMapBounds(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then

                        'hay un user
                        If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.Invisible = 1 Or UserList(MapData(PosCasteadaM, TempX, _
                                                                                                                           TempY).UserIndex).flags.Oculto = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.Privilegios = _
                                                                                                                           PlayerType.User Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(MapData(PosCasteadaM, _
                                                                                                                                TempX, TempY).UserIndex).char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)

                        End If

                    End If

                End If

            Next TempY
        Next TempX

        Call InfoHechizo(UserIndex)

    End If

End Sub

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef b As Boolean)

    If UserList(UserIndex).NroMacotas >= MAXMASCOTAS Then Exit Sub

    'No permitimos se invoquen criaturas en zonas seguras
    If MapInfo(UserList(UserIndex).pos.Map).Pk = False Or MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList( _
                                                                                                                          UserIndex).pos.Y).Trigger = eTrigger.ZONASEGURA Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z5")
        Exit Sub

    End If

    'If UserList(UserIndex).pos.Map = 75 Then
    '    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Aqui no puedes invocar mascotas!." & FONTTYPE_INFO)
    '     Exit Sub
    ' End If

    If UserList(UserIndex).pos.Map = MapaCasaAbandonada1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Aquí dentro no puedes crear mascotas...." & FONTTYPE_TALK)
        Exit Sub
    End If

    If UserList(UserIndex).pos.Map = 96 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El Demonio no deja invocar Mascotas." & FONTTYPE_TALK)
        Exit Sub
    End If

    If UserList(UserIndex).pos.Map = 98 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Una fuerza oscura te impide lanzar este hechizo." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UserList(UserIndex).pos.Map = 99 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Una fuerza oscura te impide lanzar este hechizo." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UserList(UserIndex).pos.Map = 100 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Una fuerza oscura te impide lanzar este hechizo." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UserList(UserIndex).pos.Map = 101 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Una fuerza oscura te impide lanzar este hechizo." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UserList(UserIndex).pos.Map = 154 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Aquí dentro no puedes crear mascotas...." & FONTTYPE_INFO)
        Exit Sub
    End If

    Dim h As Integer, j As Integer, ind As Integer, Index As Integer
    Dim TargetPos As WorldPos

    TargetPos.Map = UserList(UserIndex).flags.TargetMap
    TargetPos.X = UserList(UserIndex).flags.TargetX
    TargetPos.Y = UserList(UserIndex).flags.TargetY

    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

    For j = 1 To Hechizos(h).Cant

        If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
            ind = SpawnNpc(Hechizos(h).NumNpc, TargetPos, True, False)

            If ind > 0 Then
                UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1

                Index = FreeMascotaIndex(UserIndex)

                UserList(UserIndex).MascotasIndex(Index) = ind
                UserList(UserIndex).MascotasType(Index) = Npclist(ind).Numero

                Npclist(ind).MaestroUser = UserIndex
                Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
                Npclist(ind).GiveGLD = 0

                Call FollowAmo(ind)

            End If

        Else
            Exit For

        End If

    Next j

    Call InfoHechizo(UserIndex)
    b = True

End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal uh As Integer)

    Dim b As Boolean

    If UserList(UserIndex).Counters.TimerAttack > 0 Then Exit Sub

    Select Case Hechizos(uh).Tipo

    Case TipoHechizo.uInvocacion    '
        Call HechizoInvocacion(UserIndex, b)

    Case TipoHechizo.uEstado
        Call HechizoTerrenoEstado(UserIndex, b)

    Case TipoHechizo.uArea
        Call HechizoAreaUsuario(UserIndex, b)
    End Select

    If b Then
        Call SubirSkill(UserIndex, Magia)
        'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido

        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido

        If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
        Call EnviarMn(UserIndex)
        Call EnviarSta(UserIndex)

    End If

End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)

    Dim b As Boolean

    If UserList(UserIndex).Counters.TimerAttack > 0 Then Exit Sub

    Select Case Hechizos(uh).Tipo

    Case TipoHechizo.uEstado    ' Afectan estados (por ejem : Envenenamiento)
        Call HechizoEstadoUsuario(UserIndex, b)

    Case TipoHechizo.uPropiedades    ' Afectan HP,MANA,STAMINA,ETC
        Call HechizoPropUsuario(UserIndex, b)

    Case TipoHechizo.uArea
        Call HechizoEstadoUsuario(UserIndex, b)

    End Select



    If b Then
        Call SubirSkill(UserIndex, Magia)
        'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido

        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido

        If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
        Call EnviarSta(UserIndex)
        Call EnviarMn(UserIndex)
        Call EnviarHP(UserList(UserIndex).flags.TargetUser)
        UserList(UserIndex).flags.TargetUser = 0

    End If

End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal uh As Integer)

    If Not HechizoNpcCastillo(UserIndex, uh) Then Exit Sub

    If UserList(UserIndex).flags.Demonio = True Then
        If Npclist(UserList(UserIndex).flags.TargetNpc).Numero = 253 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes acatar tu rey!!" & FONTTYPE_INFO)
            Exit Sub
        End If
    End If

    If UserList(UserIndex).flags.Angel = True Then
        If Npclist(UserList(UserIndex).flags.TargetNpc).Numero = 254 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes acatar tu rey!!" & FONTTYPE_INFO)
            Exit Sub
        End If
    End If

    If UserList(UserIndex).flags.Corsarios = True Then
        If Npclist(UserList(UserIndex).flags.TargetNpc).Numero = NpcCorsarios Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes acatar tu rey!!" & FONTTYPE_INFO)
            Exit Sub
        End If
    End If

    If UserList(UserIndex).flags.Piratas = True Then
        If Npclist(UserList(UserIndex).flags.TargetNpc).Numero = NpcPiratas Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes acatar tu rey!!" & FONTTYPE_INFO)
            Exit Sub
        End If
    End If

    If UserList(UserIndex).pos.Map = MapaCasaAbandonada1 Then
        Call Efecto_AccionCasaEncantada(UserIndex, UserList(UserIndex).flags.TargetNpc)
    End If

    Dim b As Boolean

    If UserList(UserIndex).Counters.TimerAttack > 0 Then Exit Sub

    Select Case Hechizos(uh).Tipo

    Case TipoHechizo.uEstado    ' Afectan estados (por ejem : Envenenamiento)
        Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNpc, uh, b, UserIndex)

    Case TipoHechizo.uPropiedades    ' Afectan HP,MANA,STAMINA,ETC
        Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNpc, UserIndex, b)

    End Select

    If b Then
        Call SubirSkill(UserIndex, Magia)
        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido

        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido

        If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
        Call EnviarMn(UserIndex)
        Call EnviarSta(UserIndex)

    End If

End Sub

Sub LanzarHechizo(Index As Integer, UserIndex As Integer)

    Dim uh As Integer
    Dim exito As Boolean

    uh = UserList(UserIndex).Stats.UserHechizos(Index)

    If UserList(UserIndex).pos.Map = MapaMedusa Then
        If UserList(UserIndex).pos.X >= 45 And UserList(UserIndex).pos.X <= 56 _
           And UserList(UserIndex).pos.Y >= 37 And UserList(UserIndex).pos.Y <= 42 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar en la zona de reclutamiento!!" & FONTTYPE_INFO)
            Exit Sub
        End If
    End If

    If uh = H_Demonio Or uh = H_DemonioII Then
        If UserList(UserIndex).Metamorfosis.Demonio = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No eres demonio." & FONTTYPE_INFO)
            Exit Sub
        End If
    ElseIf uh = H_Angel Or uh = H_AngelII Then
        If UserList(UserIndex).Metamorfosis.Angel = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No eres angel." & FONTTYPE_INFO)
            Exit Sub
        End If
    End If

    'if UserList(UserIndex).flags.Desnudo = 1 Then
    '    Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes atacar sin ropa." & FONTTYPE_WARNING)
    '    Exit Sub
    'End If

    If PuedeLanzar(UserIndex, uh) Then

        Select Case Hechizos(uh).Target

        Case TargetType.uUsuarios

            If UserList(UserIndex).flags.TargetUser > 0 Then
                Call HandleHechizoUsuario(UserIndex, uh)

            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Este hechizo actua solo sobre usuarios." & FONTTYPE_INFO)

            End If

        Case TargetType.uNPC

            If UserList(UserIndex).flags.TargetNpc > 0 Then
                Call HandleHechizoNPC(UserIndex, uh)

            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Este hechizo solo afecta a los npcs." & FONTTYPE_INFO)

            End If

        Case TargetType.uUsuariosYnpc

            If UserList(UserIndex).flags.TargetUser > 0 Then
                Call HandleHechizoUsuario(UserIndex, uh)

            ElseIf UserList(UserIndex).flags.TargetNpc > 0 Then
                Call HandleHechizoNPC(UserIndex, uh)

            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Objetivo inválido." & FONTTYPE_INFO)

            End If

        Case TargetType.uTerreno
            Call HandleHechizoTerreno(UserIndex, uh)

        End Select

    End If

    If UserList(UserIndex).Counters.Trabajando Then UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

End Sub

Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)

    Dim h As Integer, TU As Integer

    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    TU = UserList(UserIndex).flags.TargetUser
    
    If UserList(UserIndex).pos.Map = 192 Then
        If UserList(UserIndex).flags.SuPareja = TU Then Exit Sub
    End If

    If UserList(UserIndex).flags.Demonio = True And UserList(TU).flags.Demonio = True And Not Hechizos(h).RemoverParalisis = 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).flags.Angel = True And UserList(TU).flags.Angel = True And Not Hechizos(h).RemoverParalisis = 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).flags.Corsarios = True And UserList(TU).flags.Corsarios = True And Not Hechizos(h).RemoverParalisis = 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).flags.Piratas = True And UserList(TU).flags.Piratas = True And Not Hechizos(h).RemoverParalisis = 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
        Exit Sub

    End If

    If Hechizos(h).Invisibilidad = 1 Then

        If UserList(UserIndex).pos.Map = MAPADUELO Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No puedes echarte invisibilidad en esta sala!." & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(UserIndex).pos.Map = MapaMedusa Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No puedes usar este hechizo en este mapa!." & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(UserIndex).pos.Map = MapaBan Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No puedes usar este hechizo en este mapa!." & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(UserIndex).pos.Map = mapainvo Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No puedes echarte invisibilidad en esta sala!." & FONTTYPE_INFO)
            Exit Sub

        End If

        'If UserList(TU).flags.Muerto = 1 Then
        '   Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Está muerto!" & FONTTYPE_INFO)
        '   b = False
        '    Exit Sub
        ' End If

        If Criminal(TU) And Not Criminal(UserIndex) Then
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z6")
                Exit Sub
            Else
                'Aqui se hace criminales ciudadano vs ciudadano cuando pegan con hechizos.
                Call VolverCriminal(UserIndex)

            End If

        End If

        UserList(TU).flags.Invisible = 1

        Call SendData(SendTarget.ToMap, 0, UserList(TU).pos.Map, "NOVER" & UserList(TU).char.CharIndex & ",1," & UserList(TU).PartyIndex)

        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).Mimetiza = 1 Then
        If UserList(TU).flags.Muerto = 1 Then
            Exit Sub

        End If

        If UserList(TU).flags.Navegando = 1 Then
            Exit Sub

        End If

        If UserList(UserIndex).flags.Navegando = 1 Then
            Exit Sub

        End If

        If UserList(TU).flags.Privilegios >= PlayerType.Consejero Then
            Exit Sub

        End If

        If UserList(UserIndex).flags.Mimetizado = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya te encuentras transformado. El hechizo no ha tenido efecto" & FONTTYPE_INFO)
            Exit Sub

        End If

        'copio el char original al mimetizado

        With UserList(UserIndex)
            .CharMimetizado.Body = .char.Body
            .CharMimetizado.Head = .char.Head
            .CharMimetizado.CascoAnim = .char.CascoAnim

            .CharMimetizado.ShieldAnim = .char.ShieldAnim
            .CharMimetizado.WeaponAnim = .char.WeaponAnim

            .flags.Mimetizado = 1

            'ahora pongo local el del enemigo
            .char.Body = UserList(TU).char.Body
            .char.Head = UserList(TU).char.Head

            .char.CascoAnim = UserList(TU).char.CascoAnim
            .char.ShieldAnim = UserList(TU).char.ShieldAnim
            .char.WeaponAnim = UserList(TU).char.WeaponAnim

            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

        End With

        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).Metamorfosis.Status = 1 Then

        If UserList(TU).flags.Metamorfosis = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡El Usuario ya está transformado!" & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(TU).flags.Mimetizado = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡El Usuario ya está transformado!" & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(TU).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Esta muerto!!" & FONTTYPE_INFO)
            Exit Sub

        End If

        With UserList(TU)

            .CharMimetizado.Body = .char.Body
            .CharMimetizado.Head = .char.Head
            .CharMimetizado.CascoAnim = .char.CascoAnim

            .CharMimetizado.ShieldAnim = .char.ShieldAnim
            .CharMimetizado.WeaponAnim = .char.WeaponAnim

            .CharMimetizado.Alas = .char.Alas

            .CharMimetizado.Fuerza = .Stats.UserAtributos(eAtributos.Fuerza)
            .CharMimetizado.Agilidad = .Stats.UserAtributos(eAtributos.Agilidad)
            .CharMimetizado.Inteligencia = .Stats.UserAtributos(eAtributos.Inteligencia)

            If Left$(Hechizos(h).Metamorfosis.Fuerza, 1) = "-" Then
                .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) - Right(Hechizos(h).Metamorfosis.Fuerza, Len(.Stats.UserAtributos(eAtributos.Fuerza)) - 1)
            Else
                .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + Hechizos(h).Metamorfosis.Fuerza

            End If

            If Left$(Hechizos(h).Metamorfosis.Agilidad, 1) = "-" Then
                .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) - Right(Hechizos(h).Metamorfosis.Agilidad, Len(.Stats.UserAtributos(eAtributos.Agilidad)) - 1)
            Else
                .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + Hechizos(h).Metamorfosis.Agilidad

            End If

            If Left$(Hechizos(h).Metamorfosis.Inteligencia, 1) = "-" Then
                .Stats.UserAtributos(eAtributos.Inteligencia) = .Stats.UserAtributos(eAtributos.Inteligencia) - Right(Hechizos(h).Metamorfosis.Inteligencia, Len(.Stats.UserAtributos(eAtributos.Inteligencia)) - 1)
            Else
                .Stats.UserAtributos(eAtributos.Inteligencia) = .Stats.UserAtributos(eAtributos.Inteligencia) + Hechizos(h).Metamorfosis.Inteligencia

            End If

            .flags.Metamorfosis = 1

            .char.Body = Hechizos(h).Metamorfosis.Body
            .char.Head = 0
            .char.WeaponAnim = 2
            .char.ShieldAnim = 2
            .char.Alas = 0

            Call ChangeUserChar(SendTarget.ToMap, 0, .pos.Map, UserIndex, .char.Body, .char.Head, .char.heading, .char.WeaponAnim, .char.ShieldAnim, .char.CascoAnim, .char.Alas)

            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & 1 & "," & 1)

            .Counters.Metamorfosis = IntervaloMetamorfosis

        End With

        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).Envenena = 2 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)

        End If

        UserList(TU).flags.HechizoVeneno = 1
        UserList(TU).flags.Envenenado = 1
        UserList(TU).TipoVeneno = 2
        UserList(TU).AumentoVeneno = UserList(TU).AumentoVeneno + 13
        Call InfoHechizo(UserIndex)
        Call EnviarMn(UserIndex)
        b = True

    ElseIf Hechizos(h).Envenena = 20 Then

        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)

        End If

        UserList(TU).flags.HechizoVeneno = 1
        UserList(TU).flags.Envenenado = 1
        UserList(TU).TipoVeneno = 20
        UserList(TU).AumentoVeneno = UserList(TU).AumentoVeneno + 31
        Call InfoHechizo(UserIndex)
        Call EnviarMn(UserIndex)
        b = True

    End If

    If Hechizos(h).CuraVeneno = 1 Then
        If UserList(TU).flags.Muerto = 1 And UserIndex <> TU Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Está muerto!!" & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(TU).flags.Envenenado = 0 And UserIndex Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No estás envenenado!!" & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(TU).flags.Envenenado = 0 And UserIndex <> TU Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No está envenenado!!" & FONTTYPE_INFO)
            Exit Sub

        End If

        UserList(TU).flags.Envenenado = 0

        If UserList(TU).flags.HechizoVeneno = 1 Then
            UserList(TU).flags.HechizoVeneno = 0
            UserList(TU).TipoVeneno = 0
            UserList(TU).AumentoVeneno = 0

        End If

        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).Maldicion = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)

        End If

        UserList(TU).flags.Maldicion = 1
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).RemoverMaldicion = 1 Then
        UserList(TU).flags.Maldicion = 0
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).Bendicion = 1 Then
        UserList(TU).flags.Bendicion = 1
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).Paraliza = 1 Or Hechizos(h).Inmoviliza = 1 Then

        Dim PuedeInmo As Long

        If UserList(TU).flags.Paralizado = 0 Then
            If Not PuedeAtacar(UserIndex, TU) Then Exit Sub

            If UserIndex <> TU Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TU)

            End If

            Call InfoHechizo(UserIndex)
            b = True

            PuedeInmo = RandomNumber(1, 100)

            If UCase$(UserList(TU).Clase) = "DRUIDA" Then
                If PuedeInmo > 45 Then
                    Call SendData(ToIndex, TU, 0, "||Has logrado escapar de la paralisis!" & FONTTYPE_Motd4)
                    Exit Sub

                End If

            End If

            UserList(TU).flags.Paralizado = 1
            UserList(TU).Counters.Paralisis = IntervaloParalizado
            Call SendData(SendTarget.ToIndex, TU, 0, "PARADOW")
            
            If UCase$(Hechizos(h).nombre) = "PARALIZAR" Then
                Call SendData(SendTarget.ToIndex, TU, 0, "PARADO2")

            End If
            
            Call SendData(SendTarget.ToIndex, TU, 0, "PU" & UserList(TU).pos.X & "," & UserList(TU).pos.Y)
            Call Corr_ActualizarPosicion(TU, UserList(TU).pos.X, UserList(TU).pos.Y)

        End If

    End If

    If Hechizos(h).RemoverParalisis = 1 Then
        If UserList(TU).flags.Paralizado = 1 Then
        
            UserList(TU).flags.Paralizado = 0
            'no need to crypt this
            Call SendData(SendTarget.ToIndex, TU, 0, "PARADOW")
            Call Corr_ActualizarPosicion(TU, UserList(TU).pos.X, UserList(TU).pos.Y)
            Call InfoHechizo(UserIndex)
            b = True

        End If

    End If

    If Hechizos(h).RemoverEstupidez = 1 Then
        If Not UserList(TU).flags.Estupidez = 0 Then
            UserList(TU).flags.Estupidez = 0
            'no need to crypt this
            Call SendData(SendTarget.ToIndex, TU, 0, "NESTUP")
            Call InfoHechizo(UserIndex)
            b = True

        End If

    End If

    If Hechizos(h).Revivir = 1 Then

        If UserList(TU).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z6")

        End If

        If UserList(TU).flags.Muerto = 0 And UserIndex Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No te puedes resucitar si estás vivo!!" & FONTTYPE_INFO)
            Exit Sub
        ElseIf UserList(TU).flags.Muerto = 0 And UserIndex <> TU Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes resucitar a los vivos!!" & FONTTYPE_INFO)
            Exit Sub

        End If

        '/Juan Maraxus
        If Not Criminal(TU) Then
            If TU <> UserIndex Then
                Call AddtoVar(UserList(UserIndex).Reputacion.NobleRep, 500, MAXREP)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Los Dioses te sonrien, has ganado 500 puntos de nobleza!." & FONTTYPE_INFO)

            End If

        End If

        'UserList(TU).Stats.MinMAN = 0
        Call EnviarMn(UserIndex)
        '/Pablo Toxic Waste

        b = True
        Call InfoHechizo(UserIndex)
        Call RevivirUsuario(TU)

   ' ElseIf Hechizos(h).Revivir = 3 Then
   '     b = True
   '     Call InfoHechizo(UserIndex)
   '     Call RevivirUsuario(TU)

    End If

    If Hechizos(h).Ceguera = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)

        End If
        
        If UserList(TU).flags.Ceguera = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||El usuario ya esta ciego, el hechizo no tendria efecto alguno." & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(TU).pos.Map = MAPADUELO Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes tirarle el hechizo al usuario estando en Duelo, el hechizo no tendria efecto alguno." & FONTTYPE_FIGHT)
            Exit Sub

        End If

        If UserList(TU).pos.Map = CastilloNorte Or UserList(TU).pos.Map = CastilloOeste Or UserList(TU).pos.Map = CastilloEste Or UserList(TU).pos.Map = CastilloSur Or UserList(TU).pos.Map = MapaFortaleza Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes tirarle el hechizo al usuario estando en conquista de los castillos, el hechizo no tendria efecto alguno." & FONTTYPE_FIGHT)
            Exit Sub

        End If
        
        If UCase$(UserList(TU).Clase) <> "BARDO" And UserList(TU).flags.Angel = False And UserList(TU).flags.Demonio = False Then
            
            UserList(TU).flags.Ceguera = 1
            UserList(TU).Counters.Ceguera = IntervaloCeguera / 4

            Call SendData(SendTarget.ToIndex, TU, 0, "CEGU")

            Call InfoHechizo(UserIndex)
            b = True
            
        Else
          
            Call SendData(ToIndex, TU, 0, "|| " & UserList(UserIndex).Name & " te ha intentado cegar, pero eres INMUNE!!" & FONTTYPE_FIGHT)
            Call SendData(ToIndex, UserIndex, 0, "|| " & UserList(TU).Name & " es INMUNE!!" & FONTTYPE_FIGHT)
        
        End If

    End If

    If Hechizos(h).Estupidez = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)

        End If
        
        If UserList(TU).flags.Estupidez = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||El usuario ya esta Estupido, el hechizo no tendria efecto alguno." & FONTTYPE_FIGHT)
            Exit Sub

        End If

        If UserList(TU).pos.Map = MAPADUELO Then
            Call SendData(ToIndex, UserIndex, 0, "||¡No puedes Echar estupidez estando en Duelo!." & FONTTYPE_FIGHT)
            Exit Sub

        End If

        If UserList(TU).pos.Map = CastilloNorte Or UserList(TU).pos.Map = CastilloOeste Or UserList(TU).pos.Map = CastilloEste Or UserList(TU).pos.Map = CastilloSur Or UserList(TU).pos.Map = MapaFortaleza Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes tirarle el hechizo al usuario estando en conquista de los castillos, el hechizo no tendria efecto alguno." & FONTTYPE_FIGHT)
            Exit Sub

        End If
        
        If UCase$(UserList(TU).Clase) <> "BARDO" And UserList(TU).flags.Angel = False And UserList(TU).flags.Demonio = False Then
            
            UserList(TU).flags.Estupidez = 1
            UserList(TU).Counters.Ceguera = IntervaloCeguera

            Call SendData(SendTarget.ToIndex, TU, 0, "DUMB")

            Call InfoHechizo(UserIndex)
            b = True
            
        Else
        
            Call SendData(ToIndex, TU, 0, "|| " & UserList(UserIndex).Name & " te ha intentado volver estúpido, pero eres INMUNE!!" & FONTTYPE_FIGHT)
            Call SendData(ToIndex, UserIndex, 0, "|| " & UserList(TU).Name & " es INMUNE!!" & FONTTYPE_FIGHT)
        
        End If

    End If

End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)

    If Npclist(NpcIndex).Numero = 616 And Not Hechizos(hIndex).SubeHP = 1 Then
        If Npclist(NpcIndex).Stats.MinHP > 14000 Then
            Call SendData(ToAll, 0, 0, "LEMU")

        End If

        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
            Call SendData(ToAll, 0, 0, "LEMU")

        End If

    End If

    If Npclist(NpcIndex).Numero = 617 And Not Hechizos(hIndex).SubeHP = 1 Then
        If Npclist(NpcIndex).Stats.MinHP > 14000 Then
            Call SendData(ToAll, 0, 0, "TALE")

        End If

        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
            Call SendData(ToAll, 0, 0, "TALE")

        End If

    End If

    If Npclist(NpcIndex).Numero = 910 And Not Hechizos(hIndex).SubeHP = 1 Then
        If Npclist(NpcIndex).Stats.MinHP > 14000 Then
            Call SendData(ToAll, 0, 0, "NIX")

        End If

        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
            Call SendData(ToAll, 0, 0, "NIX")

        End If

    End If

    If Hechizos(hIndex).Invisibilidad = 1 Then
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Invisible = 1
        b = True

    End If

    If Hechizos(hIndex).Envenena = 1 Then
        If Npclist(NpcIndex).Attackable = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z7")
            Exit Sub

        End If

        If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z8")
                Exit Sub
            Else
                UserList(UserIndex).Reputacion.NobleRep = 0
                UserList(UserIndex).Reputacion.PlebeRep = 0
                UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep + 200

                If UserList(UserIndex).Reputacion.AsesinoRep > MAXREP Then UserList(UserIndex).Reputacion.AsesinoRep = MAXREP

            End If

        End If

        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Envenenado = 1
        b = True

    End If

    If Hechizos(hIndex).CuraVeneno = 1 Then
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Envenenado = 0
        b = True

    End If

    If Hechizos(hIndex).ParalisisArea = 1 Then
        Call InfoHechizo(UserIndex)

        Dim Map As Integer
        Dim X As Integer
        Dim Y As Integer
        Dim TempX As Integer
        Dim TempY As Integer

        Map = UserList(UserIndex).pos.Map
        X = UserList(UserIndex).pos.X
        Y = UserList(UserIndex).pos.Y

        For TempX = X - 8 To X + 8
            For TempY = Y - 8 To Y + 8

                If InMapBounds(Map, TempX, TempY) Then
                    If MapData(Map, TempX, TempY).NpcIndex > 0 Then

                        Npclist(MapData(Map, TempX, TempY).NpcIndex).flags.Paralizado = 1
                        Npclist(MapData(Map, TempX, TempY).NpcIndex).flags.Inmovilizado = 0
                        Npclist(MapData(Map, TempX, TempY).NpcIndex).Contadores.Paralisis = IntervaloParalizado

                        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & Npclist(MapData(Map, TempX, TempY).NpcIndex).char.CharIndex & "," & Hechizos(hIndex).FXgrh & "," & Hechizos(hIndex).loops)


                    End If

                End If



            Next TempY
        Next TempX

        Call DecirPalabrasMagicas(Hechizos(hIndex).PalabrasMagicas, UserIndex)



        b = True
    End If

    If Hechizos(hIndex).Maldicion = 1 Then
        If Npclist(NpcIndex).Attackable = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z7")
            Exit Sub

        End If

        If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z8")
                Exit Sub
            Else
                UserList(UserIndex).Reputacion.NobleRep = 0
                UserList(UserIndex).Reputacion.PlebeRep = 0
                UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep + 200

                If UserList(UserIndex).Reputacion.AsesinoRep > MAXREP Then UserList(UserIndex).Reputacion.AsesinoRep = MAXREP

            End If

        End If

        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Maldicion = 1
        b = True

    End If

    If Hechizos(hIndex).RemoverMaldicion = 1 Then
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Maldicion = 0
        b = True

    End If

    If Hechizos(hIndex).Bendicion = 1 Then
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Bendicion = 1
        b = True

    End If

    If Hechizos(hIndex).Paraliza = 1 Then

        If Npclist(NpcIndex).flags.Paralizado = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El npc esta Paralizado." & FONTTYPE_Motd4)
            Exit Sub
        End If

        If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                If UserList(UserIndex).flags.Seguro Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z8")
                    Exit Sub
                Else
                    UserList(UserIndex).Reputacion.NobleRep = 0
                    UserList(UserIndex).Reputacion.PlebeRep = 0
                    UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep + 500

                    If UserList(UserIndex).Reputacion.AsesinoRep > MAXREP Then UserList(UserIndex).Reputacion.AsesinoRep = MAXREP

                End If

            End If

            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 1
            Npclist(NpcIndex).flags.Inmovilizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
            b = True
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z9")

        End If

    End If

    '[Barrin 16-2-04]
    If Hechizos(hIndex).RemoverParalisis = 1 Then

        If Npclist(NpcIndex).flags.Paralizado = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El npc no está paralizado." & FONTTYPE_INFO)
        End If

        If Npclist(NpcIndex).flags.Paralizado = 1 Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
        End If

    End If

    '[/Barrin]

    If Hechizos(hIndex).Inmoviliza = 1 Then
        If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                If UserList(UserIndex).flags.Seguro Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z8")
                    Exit Sub
                Else
                    UserList(UserIndex).Reputacion.NobleRep = 0
                    UserList(UserIndex).Reputacion.PlebeRep = 0
                    UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep + 500

                    If UserList(UserIndex).Reputacion.AsesinoRep > MAXREP Then UserList(UserIndex).Reputacion.AsesinoRep = MAXREP

                End If

            End If

            Npclist(NpcIndex).flags.Inmovilizado = 1
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
            Call InfoHechizo(UserIndex)
            b = True
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z9")

        End If

    End If

End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean)

    Dim Daño As Long
    
    Dim Bono           As Integer

    Dim bonificaMinimo As Integer

    bonificaMinimo = 0

    If UserList(UserIndex).Invent.WeaponEqpSlot > 0 Then
        If ObjData(UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.WeaponEqpSlot).ObjIndex).Subtipo = 4 Then
            bonificaMinimo = (Hechizos(hIndex).MaxHP - Hechizos(hIndex).MinHP) / 4
        End If
    End If

    'Quitamos stamina
    If UserList(UserIndex).Stats.MinSta >= 10 Then
        If Not UserList(UserIndex).flags.SeguroCombate Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, _
                          "||No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate." & FONTTYPE_Motd4)
            Exit Sub
        End If
    Else
        Exit Sub
    End If


    If Not GolpeNpcCastillo(UserIndex, NpcIndex) Then Exit Sub

    If Npclist(NpcIndex).Numero = 663 And UserList(UserIndex).pos.Map = 102 Then
        If Npclist(NpcIndex).Stats.MinHP <= 15000 And Npclist(NpcIndex).Stats.MinHP > 5000 Then
            Call SendData(ToAll, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " esta atacando el castillo Fortaleza." & _
                                       FONTTYPE_CONSEJOCAOSVesA)

        End If

        If Npclist(NpcIndex).Stats.MinHP < 5000 Then
            Call SendData(ToAll, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & _
                                     " esta apunto de conquistar castillo Fortaleza." & FONTTYPE_CONSEJOCAOSVesA)

        End If

    End If

    If Npclist(NpcIndex).Numero = 657 And UserList(UserIndex).pos.Map = 101 Then
        If Npclist(NpcIndex).Stats.MinHP <= 15000 And Npclist(NpcIndex).Stats.MinHP > 5000 Then
            Call SendData(ToAll, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " esta atacando el castillo Oeste." & _
                                       FONTTYPE_CONSEJOCAOSVesA)

        End If

        If Npclist(NpcIndex).Stats.MinHP < 5000 Then
            Call SendData(ToAll, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & _
                                     " esta apunto de conquistar castillo Oeste." & FONTTYPE_CONSEJOCAOSVesA)

        End If

    End If

    If Npclist(NpcIndex).Numero = 657 And UserList(UserIndex).pos.Map = 100 Then
        If Npclist(NpcIndex).Stats.MinHP <= 15000 And Npclist(NpcIndex).Stats.MinHP > 5000 Then
            Call SendData(ToAll, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " esta atacando el castillo Este." & _
                                       FONTTYPE_CONSEJOCAOSVesA)

        End If

        If Npclist(NpcIndex).Stats.MinHP < 5000 Then
            Call SendData(ToAll, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & _
                                     " esta apunto de conquistar castillo Este." & FONTTYPE_CONSEJOCAOSVesA)

        End If

    End If

    If Npclist(NpcIndex).Numero = 657 And UserList(UserIndex).pos.Map = 98 Then
        If Npclist(NpcIndex).Stats.MinHP <= 15000 And Npclist(NpcIndex).Stats.MinHP > 5000 Then
            Call SendData(ToAll, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " esta atacando el castillo Norte." & _
                                       FONTTYPE_CONSEJOCAOSVesA)

        End If

        If Npclist(NpcIndex).Stats.MinHP < 5000 Then
            Call SendData(ToAll, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & _
                                     " esta apunto de conquistar castillo Norte." & FONTTYPE_CONSEJOCAOSVesA)

        End If

    End If

    If Npclist(NpcIndex).Numero = 657 And UserList(UserIndex).pos.Map = 99 Then
        If Npclist(NpcIndex).Stats.MinHP <= 15000 And Npclist(NpcIndex).Stats.MinHP > 5000 Then
            Call SendData(ToAll, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " esta atacando el castillo Sur." & _
                                       FONTTYPE_CONSEJOCAOSVesA)

        End If

        If Npclist(NpcIndex).Stats.MinHP < 5000 Then
            Call SendData(ToAll, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " esta apunto de conquistar castillo Sur." _
                                     & FONTTYPE_CONSEJOCAOSVesA)

        End If

    End If

    'Salud

    If Hechizos(hIndex).SubeHP = 1 Then
        Daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
        Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
       

        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP + Daño

        If Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP Then Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " (" & Npclist(NpcIndex).Stats.MinHP & "/" & Npclist( _
                                                        NpcIndex).Stats.MaxHP & ")." & FONTTYPE_Motd4)

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has curado " & Daño & " puntos de salud a la criatura." & FONTTYPE_Motd4)
        b = True
        
    ElseIf Hechizos(hIndex).SubeHP = 2 Then

        If Npclist(NpcIndex).Attackable = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z7")
            b = False
            Exit Sub

        End If

        If Npclist(NpcIndex).NPCtype = 2 And UserList(UserIndex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z8")
            b = False
            Exit Sub

        End If

        If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
            b = False
            Exit Sub

        End If

        Daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)

        'daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)

        If UserList(UserIndex).GranPoder = 1 And GranPoder.TipoAura = hGranPoder.Daño Then

            If UserList(UserIndex).Invent.WeaponEqpObjIndex = ObjVaraNormal Then

                Daño = (Daño * 0.7) + (((UserList(UserIndex).Stats.ELV * 4.2) / 100) * (RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)) * _
                                       0.8) + 100 * 2
            Else
                Daño = (Daño * 0.7) + (((UserList(UserIndex).Stats.ELV * 4.2) / 100) * (RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)) * _
                                       0.8) * 2
            End If

        Else

            If UserList(UserIndex).Invent.WeaponEqpObjIndex = ObjVaraNormal Then

                Daño = (Daño * 0.7) + (((UserList(UserIndex).Stats.ELV * 4.2) / 100) * (RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)) * _
                                       0.8) + 100
            Else
                Daño = (Daño * 0.7) + (((UserList(UserIndex).Stats.ELV * 4.2) / 100) * (RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)) * _
                                       0.8)
            End If

        End If

        If Npclist(NpcIndex).DefensaMagica = 1 Then
            Daño = Daño - RandomNumber(100, 150)

        End If

        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).VaraDragon = 1 And Npclist(NpcIndex).NPCtype = eNPCType.DRAGON Then
                Daño = Daño * 40

            End If

        End If

        If Hechizos(hIndex).StaffAffected Then
            If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    Daño = (Daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                    'Aumenta daño segun el staff-
                    'Daño = (Daño* (80 + BonifBáculo)) / 100
                Else
                    Daño = Daño * 0.7    'Baja daño a 80% del original

                End If

            End If

        End If

        If UserList(UserIndex).Invent.HerramientaEqpObjIndex = LAUDMAGICO Then
            Daño = Daño * 1.04  'laud magico de los bardos

        End If

        Call InfoHechizo(UserIndex)
        b = True
        Call NpcAtacado(NpcIndex, UserIndex)

        If Npclist(NpcIndex).flags.Snd2 > 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & Npclist( _
                                                                                                                            NpcIndex).flags.Snd2)

        '  Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbYellow & "°- " & Daño & "!" & "°" & str(Npclist( _
           NpcIndex).char.CharIndex))
        If Npclist(NpcIndex).NPCtype = eNPCType.DRAGON Then Daño = 1
           
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - Daño


        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Le has causado " & Daño & " puntos de daño a la criatura!" & FONTTYPE_Motd4)

        If Npclist(NpcIndex).Stats.MinHP < 0 Then
            Daño = Npclist(NpcIndex).Stats.MinHP + Daño
            Npclist(NpcIndex).Stats.MinHP = 0
        End If

        Call CalcularDarExp(UserIndex, NpcIndex, Daño)

        If Npclist(NpcIndex).Stats.MinHP < 1 Then
            Npclist(NpcIndex).Stats.MinHP = 0
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " le queda " & Npclist(NpcIndex).Stats.MinHP & " / " & Npclist( _
                                                            NpcIndex).Stats.MaxHP & " de vida." & FONTTYPE_Motd4)
            Call MuereNpc(NpcIndex, UserIndex)

        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " le queda " & Npclist(NpcIndex).Stats.MinHP & " / " & Npclist( _
                                                            NpcIndex).Stats.MaxHP & " de vida." & FONTTYPE_Motd4)
            'Mascotas atacan a la criatura.
            'Call CheckPets(NpcIndex, UserIndex, True)
        End If

    End If

End Sub

Sub InfoHechizo(ByVal UserIndex As Integer)

    Dim h As Integer
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

    Call DecirPalabrasMagicas(Hechizos(h).PalabrasMagicas, UserIndex)

    If UserList(UserIndex).flags.TargetUser > 0 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserList( _
                                                                                                    UserIndex).flags.TargetUser).char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)
        Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, UserList(UserIndex).pos.Map, "TW" & Hechizos(h).WAV)
    ElseIf UserList(UserIndex).flags.TargetNpc > 0 Then
        Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNpc, Npclist(UserList(UserIndex).flags.TargetNpc).pos.Map, "CFX" & _
                                                                                                                                       Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)
        Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNpc, UserList(UserIndex).pos.Map, "TW" & Hechizos(h).WAV)

    End If

    If UserList(UserIndex).flags.TargetUser > 0 Then
        If UserIndex <> UserList(UserIndex).flags.TargetUser Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Hechizos(h).HechizeroMsg & " " & UserList(UserList( _
                                                                                                             UserIndex).flags.TargetUser).Name & FONTTYPE_Motd4)
            Call SendData(SendTarget.ToIndex, UserList(UserIndex).flags.TargetUser, 0, "||" & UserList(UserIndex).Name & " " & Hechizos( _
                                                                                       h).TargetMsg & FONTTYPE_Motd4)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Hechizos(h).PropioMsg & FONTTYPE_Motd4)

        End If

    ElseIf UserList(UserIndex).flags.TargetNpc > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Hechizos(h).HechizeroMsg & " " & "la criatura." & FONTTYPE_Motd4)

    End If

End Sub

Sub HechizoAreaUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
    Dim h As Integer
    Dim Daño As Integer
    Dim tempChr As Integer
    Dim TempX As Integer
    Dim TempY As Integer

    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

    Select Case UserList(UserIndex).pos.Map
    Case "96", "98", "99", "100", "101", "102", "154"
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Una fuerza oscura te impide lanzar este hechizo." & FONTTYPE_INFO)
        Exit Sub
    End Select

    If UserList(UserIndex).Stats.UserSkills(eSkill.Magia) < Hechizos(h).MinSkill Then
        Call SendData(SendTarget.ToIndex, UserIndex, "0", "||Necesitarás " & Hechizos(h).MinSkill & _
                                                        " Skills en Magia para utilizarlo. " & FONTTYPE_INFO)
    End If

    With UserList(UserIndex)

        For TempX = .pos.X - 8 To .pos.X + 8
            For TempY = .pos.Y - 8 To .pos.Y + 8


                If InMapBounds(.pos.Map, TempX, TempY) Then
                    If MapData(.pos.Map, TempX, TempY).UserIndex > 0 And MapData(.pos.Map, TempX, TempY).UserIndex <> UserIndex Then
                        tempChr = MapData(.pos.Map, TempX, TempY).UserIndex
                        If UserList(MapData(.pos.Map, TempX, TempY).UserIndex).flags.Muerto = 0 And UserList(MapData(.pos.Map, TempX, TempY).UserIndex).flags.Privilegios = PlayerType.User Then

                            If Hechizos(h).SubeHP = 2 And Not MismaParty(UserIndex, MapData(.pos.Map, TempX, TempY).UserIndex) And Not MismoClan(UserIndex, MapData(.pos.Map, TempX, TempY).UserIndex) Then
                                If UserList(tempChr).Stats.MinHP < 1 Then
                                    Call ContarMuerte(tempChr, UserIndex)
                                    UserList(tempChr).Stats.MinHP = 0
                                    Call ActStats(tempChr, UserIndex)
                                    MuereSpell = Hechizos(h).FXgrh
                                    LoopSpell = Hechizos(h).loops
                                    Call UserDie(tempChr)

                                Else
                                    Daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
                                    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - Daño
                                    Call EnviarHP(tempChr)
                                End If
                                Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & .Name & " te ha quitado " & Daño & " puntos de vida." & FONTTYPE_Motd4)
                                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Hechizos(h).HechizeroMsg & UserList(tempChr).Name & "." & FONTTYPE_Motd4)
                            End If
                        End If

                        If Hechizos(h).Revivir = 1 Then
                            If UserList(tempChr).flags.Muerto = 1 Then
                                Call RevivirUsuario(tempChr)
                            End If
                            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP

                            Call EnviarHP(UserIndex)
                        End If


                    End If
                End If

                If InMapBounds(.pos.Map, TempX, TempY) Then
                    If MapData(.pos.Map, TempX, TempY).UserIndex > 0 Then
                        tempChr = MapData(.pos.Map, TempX, TempY).UserIndex
                        If UserList(MapData(.pos.Map, TempX, TempY).UserIndex).flags.Muerto = 0 And UserList(MapData(.pos.Map, TempX, TempY).UserIndex).flags.Privilegios = PlayerType.User Then

                            If Hechizos(h).SubeHP = 1 Then
                                Daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
                                If val(UserList(tempChr).Stats.MinHP + Daño) >= UserList(tempChr).Stats.MaxHP Then
                                    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP
                                    Call EnviarHP(tempChr)
                                Else
                                    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP + Daño
                                    Call EnviarHP(tempChr)
                                End If

                                If UserIndex <> tempChr Then
                                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de vida a " & UserList(tempChr).Name & _
                                                                                    FONTTYPE_Motd4)
                                    Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de vida." & _
                                                                                  FONTTYPE_Motd4)
                                Else
                                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de vida." & FONTTYPE_Motd4)
                                End If

                            End If
                        End If
                    End If
                End If

            Next TempY
        Next TempX



        If .Stats.MinMAN <= Hechizos(h).ManaRequerido Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes suficiente mana." & FONTTYPE_INFO)
        Else
            .Stats.MinMAN = .Stats.MinMAN - Hechizos(h).ManaRequerido
            Call EnviarMn(UserIndex)
        End If
    End With

    Call DecirPalabrasMagicas(Hechizos(h).PalabrasMagicas, UserIndex)
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & Hechizos(h).WAV)
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserList( _
                                                                                                UserIndex).flags.TargetUser).char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)


End Sub

Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)

    Dim h As Integer

    Dim Daño As Integer

    Dim tempChr As Integer

    Dim AmuletoDaño As Integer

    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    tempChr = UserList(UserIndex).flags.TargetUser
    
    If UserList(UserIndex).pos.Map = 192 Then
        If UserList(UserIndex).flags.SuPareja = tempChr Then Exit Sub
    End If

    If UserList(UserIndex).flags.Demonio = True And UserList(tempChr).flags.Demonio = True And Not Hechizos(h).SubeFuerza = 1 And Not Hechizos(h).SubeAgilidad = 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).flags.Angel = True And UserList(tempChr).flags.Angel = True And Not Hechizos(h).SubeFuerza = 1 And Not Hechizos(h).SubeAgilidad = 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
        Exit Sub

    End If

    'Quitamos stamina
    If UserList(UserIndex).Stats.MinSta >= 10 Then
        If Not UserList(UserIndex).flags.SeguroCombate Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate." & FONTTYPE_Motd4)
            Exit Sub

        End If

    Else
        Exit Sub

    End If

    'Hambre
    If Hechizos(h).SubeHam = 1 Then

        Call InfoHechizo(UserIndex)

        Daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)

        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + Daño

        If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam

        If UserIndex <> tempChr Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_Motd4)
            Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de hambre." & FONTTYPE_Motd4)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de hambre." & FONTTYPE_Motd4)

        End If

        Call EnviarHambreYsed(tempChr)
        b = True

    ElseIf Hechizos(h).SubeHam = 2 Then

        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub

        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
        Else
            Exit Sub

        End If

        Call InfoHechizo(UserIndex)

        Daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)

        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - Daño

        If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0

        If UserIndex <> tempChr Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_Motd4)
            Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de hambre." & FONTTYPE_Motd4)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te has quitado " & Daño & " puntos de hambre." & FONTTYPE_Motd4)

        End If

        Call EnviarHambreYsed(tempChr)

        b = True

        If UserList(tempChr).Stats.MinHam < 1 Then
            UserList(tempChr).Stats.MinHam = 0
            UserList(tempChr).flags.Hambre = 1

        End If

    End If

    'Sed
    If Hechizos(h).SubeSed = 1 Then

        Daño = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)

        Call InfoHechizo(UserIndex)

        UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + Daño

        If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU

        If UserIndex <> tempChr Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_Motd4)
            Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de sed." & FONTTYPE_Motd4)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de sed." & FONTTYPE_Motd4)

        End If

        b = True
        Call EnviarHambreYsed(tempChr)

    End If

    Dim MXATRIBUTOS As String

    ' <-------- Agilidad ---------->
    If Hechizos(h).SubeAgilidad = 1 Then

        Call InfoHechizo(UserIndex)
        Daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)

        UserList(tempChr).flags.DuracionEfectoAmarillas = 4800
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(3, 4)

        MXATRIBUTOS = val(MAXATRIBUTOS + UserList(UserIndex).flags.EspecialAgilidad)

        If UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) > MXATRIBUTOS Then UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = MXATRIBUTOS
        UserList(tempChr).flags.TomoPocionAmarilla = True
        Call EnviarAmarillas(tempChr)
        b = True

    ElseIf Hechizos(h).SubeAgilidad = 2 Then

        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub

        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If

        Call InfoHechizo(UserIndex)

        UserList(tempChr).flags.TomoPocionAmarilla = True
        Call EnviarAmarillas(tempChr)
        Daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
        UserList(tempChr).flags.DuracionEfectoAmarillas = 4800
        UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - Daño

        If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
        b = True

    End If

    ' <-------- Fuerza ---------->

    If Hechizos(h).SubeFuerza = 1 Then

        Call InfoHechizo(UserIndex)
        Daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)

        UserList(tempChr).flags.DuracionEfectoVerdes = 4800

        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(3, 4)

        MXATRIBUTOS = val(MAXATRIBUTOS + UserList(UserIndex).flags.EspecialFuerza)

        If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) > MXATRIBUTOS Then UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = MXATRIBUTOS

        UserList(tempChr).flags.TomoPocionVerde = True
        Call EnviarVerdes(tempChr)
        b = True

    ElseIf Hechizos(h).SubeFuerza = 2 Then

        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub

        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If

        Call InfoHechizo(UserIndex)

        UserList(tempChr).flags.TomoPocionVerde = True
        Call EnviarVerdes(tempChr)

        Daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
        UserList(tempChr).flags.DuracionEfectoVerdes = 4800
        UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - Daño

        If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
        b = True

    End If

    'Salud
    If Hechizos(h).SubeHP = 1 Then

        If UserList(UserIndex).flags.Muerto = 1 And UserIndex = tempChr Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡¡Estás muerto!!!." & FONTTYPE_INFO)
            Exit Sub
        ElseIf UserList(tempChr).flags.Muerto = 1 And UserIndex <> tempChr Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡¡No puedes curar a los muertos!!!." & FONTTYPE_INFO)
            Exit Sub
        ElseIf UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP And UserIndex = tempChr Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Parece que no estás herido." & FONTTYPE_INFO)
            Exit Sub
        ElseIf UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP And UserIndex <> tempChr Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Parece que no está herido." & FONTTYPE_INFO)
            Exit Sub
        Else

            Daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
            Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)

            Call InfoHechizo(UserIndex)

            UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP + Daño

            If UserList(tempChr).Stats.MinHP > UserList(tempChr).Stats.MaxHP Then UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP

            If UserIndex <> tempChr Then

                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de vida a " & UserList(tempChr).Name & FONTTYPE_Motd4)
                Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de vida." & FONTTYPE_Motd4)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de vida." & FONTTYPE_Motd4)

            End If

        End If

        b = True
    ElseIf Hechizos(h).SubeHP = 2 Then

        'If UserList(UserIndex).flags.SeguroClan Then
        'If Guilds(UserList(tempChr).GuildIndex).GuildName = Guilds(UserList(UserIndex).GuildIndex).GuildName And Guilds(UserList(UserIndex).GuildIndex).GuildName <> "" Then
        '    Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes atacar a tu propio Clan con el seguro activado, escribe /SEGCLAN para desactivarlo." & FONTTYPE_motd4)
        '    Exit Sub
        'End If
        'End If

        If UserIndex = tempChr Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacarte a ti mismo." & FONTTYPE_Motd4)
            Exit Sub

        End If

        Daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
        'daño = daño - Porcentaje(daño, Int(((UserList(tempChr).Stats.UserSkills(Resistencia) + 1) / 4) + ResistenciaClase(UserList(tempChr).Clase)))

        'Nueva formula
        If UserList(UserIndex).GranPoder = 1 And GranPoder.TipoAura = hGranPoder.Daño Then
            Daño = (Daño * 0.7) + (((UserList(UserIndex).Stats.ELV * 4.2) / 100) * (Daño) * 0.8) * 2
        Else
            Daño = (Daño * 0.7) + (((UserList(UserIndex).Stats.ELV * 4.2) / 100) * (Daño) * 0.8)

        End If

        Daño = Daño - (Daño * (((UserList(tempChr).Stats.UserSkills(Resistencia) + ((UserList(UserIndex).Stats.ELV / 55) * 10)) / 5) / 100))

        If Hechizos(h).StaffAffected Then
            If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    Daño = (Daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                Else
                    Daño = Daño * 0.7    'Baja daño a 70% del original

                End If

            End If

        End If

        If UserList(UserIndex).Invent.HerramientaEqpObjIndex = LAUDMAGICO Then
            Daño = Daño * 1.04  'laud magico de los bardos

        End If

        'cascos antimagia
        If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
            Daño = Daño - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax)

        End If

        If UserList(tempChr).Invent.ArmourEqpObjIndex > 0 Then
            Daño = Daño - RandomNumber(ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).DefensaMagicaMax)

        End If

        'anillos
        If (UserList(tempChr).Invent.HerramientaEqpObjIndex > 0) Then
            Daño = Daño - RandomNumber(ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)

        End If

        If (UserList(tempChr).Invent.AmuletoEqpObjIndex > 0) Then

            If ObjData(UserList(tempChr).Invent.AmuletoEqpObjIndex).AmuletoDefensa.TipoBonifica = eAmuleto.otMagia Then

                AmuletoDaño = RandomNumber(1, ObjData(UserList(tempChr).Invent.AmuletoEqpObjIndex).AmuletoDefensa.Bonifica)
                Call SendData(SendTarget.ToIndex, tempChr, 0, "||Tu Amuleto te ha protegido de " & AmuletoDaño & " puntos de Daño." & FONTTYPE_TALKMSG)
                Daño = Daño - AmuletoDaño

            End If

        End If

        If Daño < 0 Then Daño = 0

        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub

        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If

        Call SubirSkill(tempChr, Resistencia)
        Call InfoHechizo(UserIndex)

        UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - Daño
        'Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbRed & "°- " & Daño & "!" & "°" & str(UserList( _
         tempChr).char.CharIndex))
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de vida a " & UserList(tempChr).Name & FONTTYPE_Motd4)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de vida." & FONTTYPE_Motd4)

        'Muere
        If UserList(tempChr).Stats.MinHP < 1 Then
            Call ContarMuerte(tempChr, UserIndex)
            UserList(tempChr).Stats.MinHP = 0
            Call ActStats(tempChr, UserIndex)

            MuereSpell = Hechizos(h).FXgrh
            LoopSpell = Hechizos(h).loops

            Call UserDie(tempChr)

        End If

        b = True

    ElseIf Hechizos(h).SubeHP = 4 And Not MismaParty(UserIndex, tempChr) And Not MismoClan(UserIndex, tempChr) And UserIndex <> tempChr Then

        Dim Map   As Integer

        Dim X     As Integer

        Dim Y     As Integer

        Dim TempX As Integer

        Dim TempY As Integer

        Map = UserList(UserIndex).pos.Map
        X = UserList(UserIndex).pos.X
        Y = UserList(UserIndex).pos.Y

        For TempX = X - 8 To X + 8
            For TempY = Y - 8 To Y + 8

                If InMapBounds(Map, TempX, TempY) Then
                    If MapData(Map, TempX, TempY).UserIndex > 0 And UserIndex <> MapData(Map, TempX, TempY).UserIndex Then

                        'hay un user
                        If UserList(MapData(Map, TempX, TempY).UserIndex).flags.Privilegios = PlayerType.User And Not MismaParty(UserIndex, MapData(Map, TempX, TempY).UserIndex) And Not MismoClan(UserIndex, MapData(Map, TempX, TempY).UserIndex) Then

                            Daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)

                            UserList(MapData(Map, TempX, TempY).UserIndex).Stats.MinHP = UserList(MapData(Map, TempX, TempY).UserIndex).Stats.MinHP - Daño
                            Call EnviarHP(MapData(Map, TempX, TempY).UserIndex)

                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de vida a " & UserList(MapData(Map, TempX, TempY).UserIndex).Name & FONTTYPE_Motd4)
                            Call SendData(SendTarget.ToIndex, MapData(Map, TempX, TempY).UserIndex, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de vida." & FONTTYPE_Motd4)

                            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(MapData(Map, TempX, TempY).UserIndex).char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)

                            If UserList(MapData(Map, TempX, TempY).UserIndex).Stats.MinHP < 1 Then
                                Call ContarMuerte(MapData(Map, TempX, TempY).UserIndex, UserIndex)
                                UserList(MapData(Map, TempX, TempY).UserIndex).Stats.MinHP = 0
                                Call ActStats(MapData(Map, TempX, TempY).UserIndex, UserIndex)

                                MuereSpell = Hechizos(h).FXgrh
                                LoopSpell = Hechizos(h).loops

                                Call UserDie(MapData(Map, TempX, TempY).UserIndex)

                            End If

                        End If

                    End If

                End If

            Next TempY
        Next TempX

        'Call InfoHechizo(UserIndex)
        Call DecirPalabrasMagicas(Hechizos(h).PalabrasMagicas, UserIndex)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Quieres pegar hechizo angelico" & FONTTYPE_INFO)
        b = True

    End If

    'Mana
    If Hechizos(h).SubeMana = 1 Then

        Call InfoHechizo(UserIndex)

        Daño = RandomNumber(Hechizos(h).MinMana, Hechizos(h).ManMana)

        UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + Daño

        If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN

        If UserIndex <> tempChr Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_Motd4)
            Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de mana." & FONTTYPE_Motd4)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de mana." & FONTTYPE_Motd4)

        End If

        b = True

    ElseIf Hechizos(h).SubeMana = 2 Then

        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub

        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If

        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_Motd4)
            Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de mana." & FONTTYPE_Motd4)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te has quitado " & Daño & " puntos de mana." & FONTTYPE_Motd4)

        End If

        UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - Daño

        If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
        b = True

    End If

    'Stamina
    If Hechizos(h).SubeSta = 1 Then
        Call InfoHechizo(UserIndex)
        UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + Daño

        If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta

        If UserIndex <> tempChr Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_Motd4)
            Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de vitalidad." & FONTTYPE_Motd4)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de vitalidad." & FONTTYPE_Motd4)

        End If

        b = True

    End If

End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

'Call LogTarea("Sub UpdateUserHechizos")

    Dim LoopC As Byte

    'Actualiza un solo slot
    If Not UpdateAll Then

        'Actualiza el inventario
        If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
            Call ChangeUserHechizo(UserIndex, Slot, UserList(UserIndex).Stats.UserHechizos(Slot))
        Else
            Call ChangeUserHechizo(UserIndex, Slot, 0)

        End If

    Else

        'Actualiza todos los slots
        For LoopC = 1 To MAXUSERHECHIZOS

            'Actualiza el inventario
            If UserList(UserIndex).Stats.UserHechizos(LoopC) > 0 Then
                Call ChangeUserHechizo(UserIndex, LoopC, UserList(UserIndex).Stats.UserHechizos(LoopC))
            Else
                Call ChangeUserHechizo(UserIndex, LoopC, 0)

            End If

        Next LoopC

    End If

End Sub

Sub EnviarOlvidoHechizos(ByVal UserIndex As Integer)
    Dim i As Integer
    Dim Hechizo As Integer

    For i = 1 To MAXUSERHECHIZOS

        Hechizo = UserList(UserIndex).Stats.UserHechizos(i)

        If Hechizo > 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "LSTH" & Hechizos(Hechizo).nombre)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "LSTH" & "(None)")
        End If

    Next i

End Sub

Sub OlvidaHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal PalabraSecreta As String)

    Dim Hechizo As Integer
    Dim Obj As Obj
    Dim ObjHechizo As Integer
    Dim i As Integer

    If UCase$(PalabraSecreta) = UCase$(UserList(UserIndex).PalabraSecreta) Then

        Hechizo = UserList(UserIndex).Stats.UserHechizos(Slot)

        If Hechizo > 0 Then

            For i = 1 To NumObjDatas
                If ObjData(i).HechizoIndex = Hechizo Then
                    ObjHechizo = i
                End If
            Next i

            UserList(UserIndex).Stats.UserHechizos(Slot) = 0
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - OroHechizo

            Obj.ObjIndex = ObjHechizo
            Obj.Amount = 1
            Call MeterItemEnInventario(UserIndex, Obj)

            Call EnviarOro(UserIndex)
            Call UpdateUserHechizos(False, UserIndex, CByte(Slot))

        End If

    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La Respuesta secreta que nos proporciono, no coincide con la del registro." & FONTTYPE_Motd4)
    End If

End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

'Call LogTarea("ChangeUserHechizo")

    UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo

    If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "SHS" & Slot & "," & Hechizo & "," & Hechizos(Hechizo).nombre)

    Else

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "SHS" & Slot & "," & "0" & "," & "(None)")

    End If

End Sub

Public Sub MoverHechizo(ByVal UserIndex As Integer, ByVal LastSlot As Integer, ByVal NewSlot As Integer)

    If Not (LastSlot >= 1 And LastSlot <= MAXUSERHECHIZOS) Then Exit Sub
    If Not (NewSlot >= 1 And NewSlot <= MAXUSERHECHIZOS) Then Exit Sub

    Dim TempHechizo As Integer

    With UserList(UserIndex).Stats

        TempHechizo = .UserHechizos(NewSlot)
        .UserHechizos(NewSlot) = .UserHechizos(LastSlot)
        .UserHechizos(LastSlot) = TempHechizo

    End With

    Call UpdateUserHechizos(False, UserIndex, LastSlot)
    Call UpdateUserHechizos(False, UserIndex, NewSlot)

End Sub

Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)

    If Not (Dire >= 1 And Dire <= 2) Then Exit Sub
    If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

    Dim TempHechizo As Integer

    If Dire = 1 Then    'Mover arriba
        If CualHechizo = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes mover el hechizo en esa direccion." & FONTTYPE_INFO)
            Exit Sub
        Else
            TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo

            Call UpdateUserHechizos(False, UserIndex, CualHechizo - 1)

        End If

    Else    'mover abajo

        If CualHechizo = MAXUSERHECHIZOS Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes mover el hechizo en esa direccion." & FONTTYPE_INFO)
            Exit Sub
        Else
            TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo

            Call UpdateUserHechizos(False, UserIndex, CualHechizo + 1)

        End If

    End If

    Call UpdateUserHechizos(True, UserIndex, CualHechizo)

End Sub

Public Sub DisNobAuBan(ByVal UserIndex As Integer, NoblePts As Long, BandidoPts As Long)

'disminuye la nobleza NoblePts puntos y aumenta el bandido BandidoPts puntos

'Si estamos en la arena no hacemos nada
    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Trigger = 6 Then Exit Sub

    'pierdo nobleza...
    UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep - NoblePts

    If UserList(UserIndex).Reputacion.NobleRep < 0 Then
        UserList(UserIndex).Reputacion.NobleRep = 0

    End If

    'gano bandido...
    UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep + BandidoPts

    If UserList(UserIndex).Reputacion.BandidoRep > MAXREP Then UserList(UserIndex).Reputacion.BandidoRep = MAXREP
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "PN")

    'If Criminal(UserIndex) Then
    '    If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    '        Call ExpulsarFaccionReal(UserIndex)
    '    End If
    '    If UserList(UserIndex).Faccion.Templario = 1 Then
    '        Call ExpulsarFaccionTemplario(UserIndex)
    '    End If
    'End If

End Sub
