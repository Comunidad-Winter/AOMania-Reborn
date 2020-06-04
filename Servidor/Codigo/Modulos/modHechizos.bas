Attribute VB_Name = "modHechizos"
Option Explicit

Public Const HELEMENTAL_FUEGO  As Integer = 26
Public Const HELEMENTAL_TIERRA As Integer = 28
Public Const SUPERANILLO       As Integer = 700

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer)

    If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
    If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Sub

    Npclist(NpcIndex).CanAttack = 0
    Dim Daño As Integer

    If Hechizos(Spell).SubeHP = 1 Then

        Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & Hechizos( _
            Spell).FXgrh & "," & Hechizos(Spell).loops)

        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + Daño

        If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha quitado " & Daño & " puntos de vida." & _
            FONTTYPE_FIGHT)
        Call EnviarHP(val(UserIndex))

    ElseIf Hechizos(Spell).SubeHP = 2 Then
    
        If UserList(UserIndex).flags.Privilegios = PlayerType.User Then
        
            Daño = Daño - Porcentaje(Daño, Int(((UserList(UserIndex).Stats.UserSkills(Resistencia) + 1) / 4) + ResistenciaClase(UserList( _
                UserIndex).Clase)))
            Call SubirSkill(UserIndex, Resistencia)
              
            Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        
            If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                Daño = Daño - RandomNumber(ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList( _
                    UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMax)

            End If
        
            If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                Daño = Daño - RandomNumber(ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(UserList( _
                    UserIndex).Invent.ArmourEqpObjIndex).DefensaMagicaMax)

            End If
        
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
                Daño = Daño - RandomNumber(ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList( _
                    UserIndex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)

            End If
        
            If Daño < 0 Then Daño = 0
        
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & Hechizos(Spell).WAV)
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & Hechizos( _
                Spell).FXgrh & "," & Hechizos(Spell).loops)
    
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - Daño
        
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha quitado " & Daño & " puntos de vida." & _
                FONTTYPE_FIGHT)
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
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & Hechizos( _
                Spell).FXgrh & "," & Hechizos(Spell).loops)
          
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = SUPERANILLO Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Tu anillo rechaza los efectos del hechizo." & FONTTYPE_FIGHT)
                Exit Sub

            End If

            UserList(UserIndex).flags.Paralizado = 1
            UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
            Call SendData(SendTarget.toIndex, UserIndex, 0, "PARADOW")
            Call SendData(SendTarget.toIndex, UserIndex, 0, "PU" & UserList(UserIndex).pos.X & "," & UserList(UserIndex).pos.Y)

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

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
    Dim hIndex As Integer
    Dim j      As Integer
    hIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).HechizoIndex

    If Not TieneHechizo(hIndex, UserIndex) Then

        'Buscamos un slot vacio
        For j = 1 To MAXUSERHECHIZOS

            If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
        Next j
        
        If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No tenes espacio para mas hechizos." & FONTTYPE_INFO)
        Else
            UserList(UserIndex).Stats.UserHechizos(j) = hIndex
            Call UpdateUserHechizos(False, UserIndex, CByte(j))
            'Quitamos del inv el item
            Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)

        End If

    Else
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Ya tenes ese hechizo." & FONTTYPE_INFO)

    End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal s As String, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim ind As String
    ind = UserList(UserIndex).char.CharIndex

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
               
        If Hechizos(HechizoIndex).ClaseProhibida(i) = UCase$(UserList(UserIndex).Clase) Then
            ClasePuedeLanzarHechizo = False
            Exit Function
        End If
               
    Next i
            
        
    ClasePuedeLanzarHechizo = True
    Exit Function
    
manejador:
    LogError ("Error en ClasePuedeUsarHechizo")
End Function

Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean
     
    If UCase$(UserList(UserIndex).Clase) <> "DIOS" Then
     
        If LenB(UCase(Hechizos(HechizoIndex).ExclusivoClase)) > 0 And UCase(Hechizos(HechizoIndex).ExclusivoClase) <> UCase(UserList(UserIndex).Clase) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tú clase no puede lanzar este hechizo." & FONTTYPE_INFO)
            PuedeLanzar = False
            Exit Function
        End If
    
        If LenB(UCase(Hechizos(HechizoIndex).ExclusivoClase)) = 0 And ClasePuedeLanzarHechizo(UserIndex, HechizoIndex) = False Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tú clase no puede lanzar este hechizo." & FONTTYPE_INFO)
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
                        Call SendData(SendTarget.toIndex, UserIndex, 0, _
                            "||Tu Báculo no es lo suficientemente poderoso para que puedas lanzar el conjuro." & FONTTYPE_INFO)
                        PuedeLanzar = False
                        Exit Function

                    End If

                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes lanzar este conjuro sin la ayuda de un báculo." & FONTTYPE_INFO)
                    PuedeLanzar = False
                    Exit Function

                End If

            End If

        End If
          
        If Hechizos(HechizoIndex).NeedStaff > 0 Then
            If UCase$(UserList(UserIndex).Clase) = "CLERIGO" Then
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu espada no es lo suficientemente fuerte para lanzar este hechizo." & _
                            FONTTYPE_INFO)
                        PuedeLanzar = False
                        Exit Function

                    End If

                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes lanzar este conjuro sin la ayuda de una Espada Argentum." & _
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
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "Z1")
                    PuedeLanzar = False

                End If
                
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z2")
                PuedeLanzar = False

            End If

        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "Z3")
            PuedeLanzar = False

        End If

    Else
        Call SendData(SendTarget.toIndex, UserIndex, 0, "Z4")
        PuedeLanzar = False

    End If

End Function

Function ResistenciaClase(Clase As String) As Integer
    Dim Cuan As Integer

    Select Case UCase$(Clase)

        Case "MAGO"
            Cuan = 1

        Case "DRUIDA"
            Cuan = 2

        Case "CLERIGO"
            Cuan = 1

        Case "BARDO"
            Cuan = 1

        Case Else
            Cuan = 0

    End Select

    ResistenciaClase = Cuan

End Function

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)
    Dim PosCasteadaX As Integer
    Dim PosCasteadaY As Integer
    Dim PosCasteadaM As Integer
    Dim H            As Integer
    Dim TempX        As Integer
    Dim TempY        As Integer

    PosCasteadaX = UserList(UserIndex).flags.TargetX
    PosCasteadaY = UserList(UserIndex).flags.TargetY
    PosCasteadaM = UserList(UserIndex).flags.TargetMap
    
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    If Hechizos(H).RemueveInvisibilidadParcial = 1 Then
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
                                TempX, TempY).UserIndex).char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)

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
        UserIndex).pos.Y).trigger = eTrigger.ZONASEGURA Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "Z5")
        Exit Sub

    End If

    If UserList(UserIndex).pos.Map = 75 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Aqui no puedes invocar mascotas!." & FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).pos.Map = 77 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Aqui no puedes invocar mascotas!." & FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).pos.Map = 96 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Aqui no puedes invocar mascotas!." & FONTTYPE_INFO)
        Exit Sub

    End If
 
    Dim H         As Integer, j As Integer, ind As Integer, Index As Integer
    Dim TargetPos As WorldPos

    TargetPos.Map = UserList(UserIndex).flags.TargetMap
    TargetPos.X = UserList(UserIndex).flags.TargetX
    TargetPos.Y = UserList(UserIndex).flags.TargetY

    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    For j = 1 To Hechizos(H).Cant
    
        If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
            ind = SpawnNpc(Hechizos(H).NumNpc, TargetPos, True, False)

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

    Select Case Hechizos(uh).Tipo

        Case TipoHechizo.uInvocacion '
            Call HechizoInvocacion(UserIndex, b)

        Case TipoHechizo.uEstado
            Call HechizoTerrenoEstado(UserIndex, b)
    
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

    Select Case Hechizos(uh).Tipo

        Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
            Call HechizoEstadoUsuario(UserIndex, b)

        Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
            Call HechizoPropUsuario(UserIndex, b)

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
     
    If UserList(UserIndex).flags.demonio = True Then
        If Npclist(UserList(UserIndex).flags.TargetNpc).Numero = 253 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes acatar tu rey!!" & FONTTYPE_INFO)
            Exit Sub
        End If
    End If

    If UserList(UserIndex).flags.angel = True Then
        If Npclist(UserList(UserIndex).flags.TargetNpc).Numero = 254 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes acatar tu rey!!" & FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    If UserList(UserIndex).flags.Corsarios = True Then
        If Npclist(UserList(UserIndex).flags.TargetNpc).Numero = NpcCorsarios Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes acatar tu rey!!" & FONTTYPE_INFO)
            Exit Sub
        End If
    End If
   
    If UserList(UserIndex).flags.Piratas = True Then
        If Npclist(UserList(UserIndex).flags.TargetNpc).Numero = NpcPiratas Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes acatar tu rey!!" & FONTTYPE_INFO)
            Exit Sub
        End If
    End If

    Dim b As Boolean

    Select Case Hechizos(uh).Tipo

        Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
            Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNpc, uh, b, UserIndex)

        Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
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

    Dim uh    As Integer
    Dim exito As Boolean

    uh = UserList(UserIndex).Stats.UserHechizos(Index)
    
    If UserList(UserIndex).pos.Map = MapaMedusa Then
        If UserList(UserIndex).pos.X >= 45 And UserList(UserIndex).pos.X <= 56 _
            And UserList(UserIndex).pos.Y >= 37 And UserList(UserIndex).pos.Y <= 42 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar en la zona de reclutamiento!!" & FONTTYPE_INFO)
            Exit Sub
        End If
    End If

    'if UserList(UserIndex).flags.Desnudo = 1 Then
    '    Call SendData(SendTarget.toindex, UserIndex, 0, "||No podés atacar sin ropa." & FONTTYPE_WARNING)
    '    Exit Sub
    'End If

    If PuedeLanzar(UserIndex, uh) Then

        Select Case Hechizos(uh).Target
        
            Case TargetType.uUsuarios

                If UserList(UserIndex).flags.TargetUser > 0 Then
                    If Abs(UserList(UserList(UserIndex).flags.TargetUser).pos.Y - UserList(UserIndex).pos.Y) <= RANGO_VISION_Y Then
                        Call HandleHechizoUsuario(UserIndex, uh)
                    Else
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)

                    End If

                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Este hechizo actua solo sobre usuarios." & FONTTYPE_INFO)

                End If

            Case TargetType.uNPC

                If UserList(UserIndex).flags.TargetNpc > 0 Then
                    If Abs(Npclist(UserList(UserIndex).flags.TargetNpc).pos.Y - UserList(UserIndex).pos.Y) <= RANGO_VISION_Y Then
                        Call HandleHechizoNPC(UserIndex, uh)
                    Else
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)

                    End If

                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Este hechizo solo afecta a los npcs." & FONTTYPE_INFO)

                End If

            Case TargetType.uUsuariosYnpc

                If UserList(UserIndex).flags.TargetUser > 0 Then
                    If Abs(UserList(UserList(UserIndex).flags.TargetUser).pos.Y - UserList(UserIndex).pos.Y) <= RANGO_VISION_Y Then
                        Call HandleHechizoUsuario(UserIndex, uh)
                    Else
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)

                    End If

                ElseIf UserList(UserIndex).flags.TargetNpc > 0 Then

                    If Abs(Npclist(UserList(UserIndex).flags.TargetNpc).pos.Y - UserList(UserIndex).pos.Y) <= RANGO_VISION_Y Then
                        Call HandleHechizoNPC(UserIndex, uh)
                    Else
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)

                    End If

                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "Z26")

                End If

            Case TargetType.uTerreno
                Call HandleHechizoTerreno(UserIndex, uh)

        End Select
    
    End If

    If UserList(UserIndex).Counters.Trabajando Then UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1
    
End Sub

Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)

    Dim H As Integer, TU As Integer
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    TU = UserList(UserIndex).flags.TargetUser
    

    If UserList(UserIndex).flags.demonio = True And UserList(TU).flags.demonio = True And Not Hechizos(H).RemoverParalisis = 1 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
        Exit Sub
    End If

    If UserList(UserIndex).flags.angel = True And UserList(TU).flags.angel = True And Not Hechizos(H).RemoverParalisis = 1 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.Corsarios = True And UserList(TU).flags.Corsarios = True And Not Hechizos(H).RemoverParalisis = 1 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.Piratas = True And UserList(TU).flags.Piratas = True And Not Hechizos(H).RemoverParalisis = 1 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
        Exit Sub
    End If
     
    If Hechizos(H).Invisibilidad = 1 Then
   
        If UserList(UserIndex).pos.Map = MAPADUELO Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes lanzar invisibilidad en duelo!" & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).pos.Map = MapaMedusa Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes lanzar invisibilidad en Guerra de Medusa!" & FONTTYPE_INFO)
            Exit Sub
        End If
   
        If UserList(UserIndex).pos.Map = mapainvo Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes lanzar invisibilidad en sala de invocaciones!" & FONTTYPE_INFO)
            Exit Sub

        End If
      
        If UserList(TU).flags.Muerto = 1 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Está muerto!" & FONTTYPE_INFO)
            b = False
            Exit Sub

        End If
    
        If Criminal(TU) And Not Criminal(UserIndex) Then
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z6")
                Exit Sub
            Else

                Call VolverCriminal(UserIndex)
        
            End If

        End If
    
        UserList(TU).flags.Invisible = 1
  
        Call SendData(SendTarget.ToMap, 0, UserList(TU).pos.Map, "NOVER" & UserList(TU).char.CharIndex & ",1," & UserList(UserIndex).PartyIndex)
                    
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(H).Mimetiza = 1 Then
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
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Ya te encuentras transformado. El hechizo no ha tenido efecto" & FONTTYPE_INFO)
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
    
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                UserIndex).OrigChar.Head, UserList(UserIndex).char.Heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)
    
        End With
   
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(H).Envenena = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)

        End If

        UserList(TU).flags.Envenenado = 1
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(H).CuraVeneno = 1 Then
        UserList(TU).flags.Envenenado = 0
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(H).Maldicion = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)

        End If

        UserList(TU).flags.Maldicion = 1
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(H).RemoverMaldicion = 1 Then
        UserList(TU).flags.Maldicion = 0
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(H).Bendicion = 1 Then
        UserList(TU).flags.Bendicion = 1
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(H).Paraliza = 1 Or Hechizos(H).Inmoviliza = 1 Then
        If UserList(TU).flags.Paralizado = 0 Then
            If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
            
            If UserIndex <> TU Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TU)

            End If
            
            Call InfoHechizo(UserIndex)
            b = True

            If UserList(TU).Invent.HerramientaEqpObjIndex = SUPERANILLO Then
                Call SendData(SendTarget.toIndex, TU, 0, "|| Tu anillo rechaza los efectos del hechizo." & FONTTYPE_FIGHT)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| ¡El hechizo no tiene efecto!" & FONTTYPE_FIGHT)
                Exit Sub

            End If
            
            UserList(TU).flags.Paralizado = 1
            UserList(TU).Counters.Paralisis = IntervaloParalizado
            Call SendData(SendTarget.toIndex, TU, 0, "PARADOW")
            Call SendData(SendTarget.toIndex, TU, 0, "PU" & UserList(TU).pos.X & "," & UserList(TU).pos.Y)

        End If

    End If

    If Hechizos(H).RemoverParalisis = 1 Then
        If UserList(TU).flags.Paralizado = 1 Then
            If Criminal(TU) And Not Criminal(UserIndex) Then
                If UserList(UserIndex).flags.Seguro Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "Z6")
                    Exit Sub
                Else
                    Call VolverCriminal(UserIndex)

                End If

            End If
        
            UserList(TU).flags.Paralizado = 0
            'no need to crypt this
            Call SendData(SendTarget.toIndex, TU, 0, "PARADOW")
            Call InfoHechizo(UserIndex)
            b = True

        End If

    End If

    If Hechizos(H).RemoverEstupidez = 1 Then
        If Not UserList(TU).flags.Estupidez = 0 Then
            UserList(TU).flags.Estupidez = 0
            'no need to crypt this
            Call SendData(SendTarget.toIndex, TU, 0, "NESTUP")
            Call InfoHechizo(UserIndex)
            b = True

        End If

    End If

    If Hechizos(H).Revivir = 1 Then

        If UserList(TU).flags.Muerto = 1 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "Z6")
        End If

        '   If UCase$(UserList(UserIndex).Clase) = "CLERIGO" Then
        '       If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        '           If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(H).NeedStaff Then
        '               Call SendData(SendTarget.toindex, UserIndex, 0, "||Necesitas una mejor espada para este hechizo" & FONTTYPE_INFO)
        '               b = False
        '               Exit Sub
        ''
        '                   End If
        '
        '                End If
        '
        '            End If
        '
        '            'revisamos si necesita vara
        '            If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
        '                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        '                    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(H).NeedStaff Then
        '                        Call SendData(SendTarget.toindex, UserIndex, 0, "||Necesitas un mejor báculo para este hechizo" & FONTTYPE_INFO)
        '                        b = False
        '                        Exit Sub
        '
        '                    End If
        '
        '                End If
        '
        '            ElseIf UCase$(UserList(UserIndex).Clase) = "BARDO" Then
        '
        '                If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> LAUDMAGICO Then
        '                    Call SendData(SendTarget.toindex, UserIndex, 0, "||Necesitas un instrumento mágico para devolver la vida" & FONTTYPE_INFO)
        '                   b = False
        '                   Exit Sub
        '
        '                End If
        '
        '            End If
        '
        '/Juan Maraxus
        If Not Criminal(TU) Then
            If TU <> UserIndex Then
                UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep + 500

                If UserList(UserIndex).Reputacion.NobleRep > MAXREP Then UserList(UserIndex).Reputacion.NobleRep = MAXREP
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Los Dioses te sonrien, has ganado 500 puntos de nobleza!." & FONTTYPE_INFO)

            End If

        End If

        'UserList(TU).Stats.MinMAN = 0
        Call EnviarMn(TU)
        '/Pablo Toxic Waste
        
        b = True
        Call InfoHechizo(UserIndex)
        Call RevivirUsuario(TU)
    Else
        b = False

    End If


    If Hechizos(H).Ceguera = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)

        End If

        UserList(TU).flags.Ceguera = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado / 3

        Call SendData(SendTarget.toIndex, TU, 0, "CEGU")

        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(H).Estupidez = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)

        End If

        UserList(TU).flags.Estupidez = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado
   
        Call SendData(SendTarget.toIndex, TU, 0, "DUMB")
           
        Call InfoHechizo(UserIndex)
        b = True

    End If

End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)

    If Npclist(NpcIndex).Numero = 616 And Not Hechizos(hIndex).SubeHP = 1 Then
        If Npclist(NpcIndex).Stats.MinHP > 14000 Then
            Call SendData(toall, 0, 0, "LEMU")

        End If

        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
            Call SendData(toall, 0, 0, "LEMU")

        End If

    End If
    
    If Npclist(NpcIndex).Numero = 617 And Not Hechizos(hIndex).SubeHP = 1 Then
        If Npclist(NpcIndex).Stats.MinHP > 14000 Then
            Call SendData(toall, 0, 0, "TALE")

        End If

        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
            Call SendData(toall, 0, 0, "TALE")

        End If

    End If

    If Npclist(NpcIndex).Numero = 910 And Not Hechizos(hIndex).SubeHP = 1 Then
        If Npclist(NpcIndex).Stats.MinHP > 14000 Then
            Call SendData(toall, 0, 0, "NIX")

        End If

        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
            Call SendData(toall, 0, 0, "NIX")

        End If

    End If
    
    If Hechizos(hIndex).Invisibilidad = 1 Then
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Invisible = 1
        b = True

    End If

    If Hechizos(hIndex).Envenena = 1 Then
        If Npclist(NpcIndex).Attackable = 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "Z7")
            Exit Sub

        End If
   
        If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z8")
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

    If Hechizos(hIndex).Maldicion = 1 Then
        If Npclist(NpcIndex).Attackable = 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "Z7")
            Exit Sub

        End If
   
        If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z8")
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
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El npc esta Paralizado." & FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                If UserList(UserIndex).flags.Seguro Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "Z8")
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
            Call SendData(SendTarget.toIndex, UserIndex, 0, "Z9")

        End If

    End If

    '[Barrin 16-2-04]
    If Hechizos(hIndex).RemoverParalisis = 1 Then
    
        If Npclist(NpcIndex).flags.Paralizado = 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Este NPC no está paralizado.!!" & FONTTYPE_INFO)
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
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "Z8")
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
            Call SendData(SendTarget.toIndex, UserIndex, 0, "Z9")

        End If

    End If

End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean)

    Dim Daño As Long
    
    If Npclist(NpcIndex).Numero = 663 And UserList(UserIndex).pos.Map = 102 Then
        If Npclist(NpcIndex).Stats.MinHP <= 15000 And Npclist(NpcIndex).Stats.MinHP > 5000 Then
            Call SendData(toall, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " esta atacando el castillo Fortaleza." & _
                FONTTYPE_CONSEJOCAOSVesA)

        End If

        If Npclist(NpcIndex).Stats.MinHP < 5000 Then
            Call SendData(toall, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & _
                " esta apunto de conquistar castillo Fortaleza." & FONTTYPE_CONSEJOCAOSVesA)

        End If

    End If
    
    If Npclist(NpcIndex).Numero = 657 And UserList(UserIndex).pos.Map = 101 Then
        If Npclist(NpcIndex).Stats.MinHP <= 15000 And Npclist(NpcIndex).Stats.MinHP > 5000 Then
            Call SendData(toall, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " esta atacando el castillo Oeste." & _
                FONTTYPE_CONSEJOCAOSVesA)

        End If

        If Npclist(NpcIndex).Stats.MinHP < 5000 Then
            Call SendData(toall, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & _
                " esta apunto de conquistar castillo Oeste." & FONTTYPE_CONSEJOCAOSVesA)

        End If

    End If
    
    If Npclist(NpcIndex).Numero = 657 And UserList(UserIndex).pos.Map = 100 Then
        If Npclist(NpcIndex).Stats.MinHP <= 15000 And Npclist(NpcIndex).Stats.MinHP > 5000 Then
            Call SendData(toall, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " esta atacando el castillo Este." & _
                FONTTYPE_CONSEJOCAOSVesA)

        End If

        If Npclist(NpcIndex).Stats.MinHP < 5000 Then
            Call SendData(toall, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & _
                " esta apunto de conquistar castillo Este." & FONTTYPE_CONSEJOCAOSVesA)

        End If

    End If
    
    If Npclist(NpcIndex).Numero = 657 And UserList(UserIndex).pos.Map = 98 Then
        If Npclist(NpcIndex).Stats.MinHP <= 15000 And Npclist(NpcIndex).Stats.MinHP > 5000 Then
            Call SendData(toall, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " esta atacando el castillo Norte." & _
                FONTTYPE_CONSEJOCAOSVesA)

        End If

        If Npclist(NpcIndex).Stats.MinHP < 5000 Then
            Call SendData(toall, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & _
                " esta apunto de conquistar castillo Norte." & FONTTYPE_CONSEJOCAOSVesA)

        End If

    End If
    
    If Npclist(NpcIndex).Numero = 657 And UserList(UserIndex).pos.Map = 99 Then
        If Npclist(NpcIndex).Stats.MinHP <= 15000 And Npclist(NpcIndex).Stats.MinHP > 5000 Then
            Call SendData(toall, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " esta atacando el castillo Sur." & _
                FONTTYPE_CONSEJOCAOSVesA)

        End If

        If Npclist(NpcIndex).Stats.MinHP < 5000 Then
            Call SendData(toall, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " esta apunto de conquistar castillo Sur." _
                & FONTTYPE_CONSEJOCAOSVesA)

        End If

    End If

    'Salud

    If Hechizos(hIndex).SubeHP = 1 Then
        Daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
        Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbCyan & "°+ " & Daño & "!" & "°" & str(Npclist( _
            NpcIndex).char.CharIndex))

        ' If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        '     If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).VaraDragon = 1 And Npclist( _
        '             NpcIndex).NPCtype = eNPCType.DRAGON Then
        '         Daño = Daño * 40
        '
        '            End If
        '
        '        End If
    
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP + Daño

        If Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP Then Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
                
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " (" & Npclist(NpcIndex).Stats.MinHP & "/" & Npclist( _
            NpcIndex).Stats.MaxHP & ")." & FONTTYPE_FIGHT)
                
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has curado " & Daño & " puntos de salud a la criatura." & FONTTYPE_FIGHT)
        b = True
    ElseIf Hechizos(hIndex).SubeHP = 2 Then
    
        If Npclist(NpcIndex).Attackable = 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "Z7")
            b = False
            Exit Sub

        End If
    
        If Npclist(NpcIndex).NPCtype = 2 And UserList(UserIndex).flags.Seguro Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "Z8")
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
                    Daño = Daño * 0.7 'Baja daño a 80% del original

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
    
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbYellow & "°- " & Daño & "!" & "°" & str(Npclist( _
            NpcIndex).char.CharIndex))
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - Daño
        
        If Npclist(NpcIndex).Stats.MinHP < 0 Then Npclist(NpcIndex).Stats.MinHP = 0
        
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " (" & Npclist(NpcIndex).Stats.MinHP & "/" & Npclist( _
            NpcIndex).Stats.MaxHP & ")." & FONTTYPE_FIGHT)
        
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Le has causado " & Daño & " puntos de daño a la criatura!" & FONTTYPE_FIGHT)
                
        Call CalcularDarExp(UserIndex, NpcIndex, Daño)
        
    
        If Npclist(NpcIndex).Stats.MinHP < 1 Then
            Npclist(NpcIndex).Stats.MinHP = 0
            Call MuereNpc(NpcIndex, UserIndex)
        Else
            'Mascotas atacan a la criatura.
            Call CheckPets(NpcIndex, UserIndex, True)
        End If

    End If

End Sub

Sub InfoHechizo(ByVal UserIndex As Integer)

    Dim H As Integer
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    Call DecirPalabrasMagicas(Hechizos(H).PalabrasMagicas, UserIndex)
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserList( _
            UserIndex).flags.TargetUser).char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
        Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, UserList(UserIndex).pos.Map, "TW" & Hechizos(H).WAV)
    ElseIf UserList(UserIndex).flags.TargetNpc > 0 Then
        Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNpc, Npclist(UserList(UserIndex).flags.TargetNpc).pos.Map, "CFX" & _
            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
        Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNpc, UserList(UserIndex).pos.Map, "TW" & Hechizos(H).WAV)

    End If
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
        If UserIndex <> UserList(UserIndex).flags.TargetUser Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & Hechizos(H).HechizeroMsg & " " & UserList(UserList( _
                UserIndex).flags.TargetUser).Name & FONTTYPE_FIGHT)
            Call SendData(SendTarget.toIndex, UserList(UserIndex).flags.TargetUser, 0, "||" & UserList(UserIndex).Name & " " & Hechizos( _
                H).TargetMsg & FONTTYPE_FIGHT)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & Hechizos(H).PropioMsg & FONTTYPE_FIGHT)

        End If

    ElseIf UserList(UserIndex).flags.TargetNpc > 0 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & Hechizos(H).HechizeroMsg & " " & "la criatura." & FONTTYPE_FIGHT)

    End If

End Sub

Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)

    Dim H As Integer
    Dim Daño As Integer
    Dim tempChr As Integer
    
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    tempChr = UserList(UserIndex).flags.TargetUser
      
    If UserList(UserIndex).flags.demonio = True And UserList(tempChr).flags.demonio = True And Not Hechizos(H).SubeFuerza = 1 And Not Hechizos( _
        H).SubeAgilidad = 1 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).flags.angel = True And UserList(tempChr).flags.angel = True And Not Hechizos(H).SubeFuerza = 1 And Not Hechizos( _
        H).SubeAgilidad = 1 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
        Exit Sub
    End If
      
    'Hambre
    If Hechizos(H).SubeHam = 1 Then
    
        Call InfoHechizo(UserIndex)
    
        Daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + Daño

        If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
    
        If UserIndex <> tempChr Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de hambre a " & UserList(tempChr).Name & _
                FONTTYPE_FIGHT)
            Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de hambre." & _
                FONTTYPE_FIGHT)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)

        End If
    
        Call EnviarHambreYsed(tempChr)
        b = True
    
    ElseIf Hechizos(H).SubeHam = 2 Then

        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
        Else
            Exit Sub

        End If
    
        Call InfoHechizo(UserIndex)
    
        Daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - Daño
    
        If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
        If UserIndex <> tempChr Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de hambre a " & UserList(tempChr).Name & _
                FONTTYPE_FIGHT)
            Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de hambre." & _
                FONTTYPE_FIGHT)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Te has quitado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)

        End If
    
        Call EnviarHambreYsed(tempChr)
    
        b = True
    
        If UserList(tempChr).Stats.MinHam < 1 Then
            UserList(tempChr).Stats.MinHam = 0
            UserList(tempChr).flags.Hambre = 1

        End If
    
    End If

    'Sed
    If Hechizos(H).SubeSed = 1 Then
            
        Daño = RandomNumber(Hechizos(H).MinSed, Hechizos(H).MaxSed)
    
        Call InfoHechizo(UserIndex)
    
        UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + Daño

        If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
         
        If UserIndex <> tempChr Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de sed a " & UserList(tempChr).Name & _
                FONTTYPE_FIGHT)
            Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de sed." & _
                FONTTYPE_FIGHT)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)

        End If
    
        b = True
        Call EnviarHambreYsed(tempChr)
        
        
    End If
    
    Dim MXATRIBUTOS As String
    ' <-------- Agilidad ---------->
    If Hechizos(H).SubeAgilidad = 1 Then
    
        Call InfoHechizo(UserIndex)
        Daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    
        UserList(tempChr).flags.DuracionEfectoAmarillas = 4800
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + _
            RandomNumber(3, 4)
                    
        MXATRIBUTOS = val(MAXATRIBUTOS + UserList(UserIndex).flags.EspecialAgilidad)
                    
        If UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) > MXATRIBUTOS Then UserList(UserIndex).Stats.UserAtributos( _
            eAtributos.Agilidad) = MXATRIBUTOS
        UserList(tempChr).flags.TomoPocionAmarilla = True
        Call EnviarAmarillas(tempChr)
        b = True
       
    ElseIf Hechizos(H).SubeAgilidad = 2 Then
    
        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If
    
        Call InfoHechizo(UserIndex)
    
        UserList(tempChr).flags.TomoPocionAmarilla = True
        Call EnviarAmarillas(tempChr)
        Daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
        UserList(tempChr).flags.DuracionEfectoAmarillas = 4800
        UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - Daño

        If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos( _
            eAtributos.Agilidad) = MINATRIBUTOS
        b = True

    End If

    ' <-------- Fuerza ---------->
  
    If Hechizos(H).SubeFuerza = 1 Then
    
        Call InfoHechizo(UserIndex)
        Daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    
        UserList(tempChr).flags.DuracionEfectoVerdes = 4800

        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + _
            RandomNumber(3, 4)
                    
        MXATRIBUTOS = val(MAXATRIBUTOS + UserList(UserIndex).flags.EspecialFuerza)
                                        
        If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) > MXATRIBUTOS Then UserList(UserIndex).Stats.UserAtributos( _
            eAtributos.Fuerza) = MXATRIBUTOS
                            
        UserList(tempChr).flags.TomoPocionVerde = True
        Call EnviarVerdes(tempChr)
        b = True
        
    ElseIf Hechizos(H).SubeFuerza = 2 Then

        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If
    
        Call InfoHechizo(UserIndex)
    
        UserList(tempChr).flags.TomoPocionVerde = True
        Call EnviarVerdes(tempChr)
    
        Daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
        UserList(tempChr).flags.DuracionEfectoVerdes = 4800
        UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - Daño

        If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = _
            MINATRIBUTOS
        b = True

    End If

    'Salud
    If Hechizos(H).SubeHP = 1 Then
 
        If UserList(UserIndex).flags.Muerto = 1 And UserIndex = tempChr Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡¡Estás muerto!!!." & FONTTYPE_INFO)
        ElseIf UserList(tempChr).flags.Muerto = 1 And UserIndex <> tempChr Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡¡No puedes curar a los muertos!!!." & FONTTYPE_INFO)
        ElseIf UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP And UserIndex = tempChr Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Parece que no estás herido." & FONTTYPE_INFO)
        ElseIf UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP And UserIndex <> tempChr Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Parece que no está herido." & FONTTYPE_INFO)
        
        Else
    
            Daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
            Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
    
            Call InfoHechizo(UserIndex)

            UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP + Daño

            If UserList(tempChr).Stats.MinHP > UserList(tempChr).Stats.MaxHP Then UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP
    
            If UserIndex <> tempChr Then
    
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de vida a " & UserList(tempChr).Name & _
                    FONTTYPE_FIGHT)
                Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de vida." & _
                    FONTTYPE_FIGHT)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de vida." & FONTTYPE_FIGHT)

            End If
        End If
    
        b = True
    ElseIf Hechizos(H).SubeHP = 2 Then
    
        'If UserList(UserIndex).flags.SeguroClan Then
        'If Guilds(UserList(tempChr).GuildIndex).GuildName = Guilds(UserList(UserIndex).GuildIndex).GuildName And Guilds(UserList(UserIndex).GuildIndex).GuildName <> "" Then
        '    Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes atacar a tu propio Clan con el seguro activado, escribe /SEGCLAN para desactivarlo." & FONTTYPE_FIGHT)
        '    Exit Sub
        'End If
        'End If
    
        If UserIndex = tempChr Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No podes atacarte a vos mismo." & FONTTYPE_FIGHT)
            Exit Sub
        End If
    
        Daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
        'daño = daño - Porcentaje(daño, Int(((UserList(tempChr).Stats.UserSkills(Resistencia) + 1) / 4) + ResistenciaClase(UserList(tempChr).Clase)))
        'Nueva formula
        If UserList(UserIndex).GranPoder = 1 And GranPoder.TipoAura = hGranPoder.Daño Then
            Daño = (Daño * 0.7) + (((UserList(UserIndex).Stats.ELV * 4.2) / 100) * (RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)) * 0.8) * 2
        Else
            Daño = (Daño * 0.7) + (((UserList(UserIndex).Stats.ELV * 4.2) / 100) * (RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)) * 0.8)
        End If
        
        'daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
        'Resistencia Magica
        Daño = Daño - ((UserList(UserIndex).Stats.ELV * 1.1) / 100) * UserList(tempChr).Stats.UserSkills(Resistencia)
          
        If Hechizos(H).StaffAffected Then
            If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    Daño = (Daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                Else
                    Daño = Daño * 0.7 'Baja daño a 70% del original

                End If

            End If

        End If
    
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex = LAUDMAGICO Then
            Daño = Daño * 1.04  'laud magico de los bardos

        End If
    
        'cascos antimagia
        If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
            Daño = Daño - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList( _
                tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax)

        End If
    
        If UserList(tempChr).Invent.ArmourEqpObjIndex > 0 Then
            Daño = Daño - RandomNumber(ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(UserList( _
                tempChr).Invent.ArmourEqpObjIndex).DefensaMagicaMax)

        End If
    
        'anillos
        If (UserList(tempChr).Invent.HerramientaEqpObjIndex > 0) Then
            Daño = Daño - RandomNumber(ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList( _
                tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)

        End If
    
        If Daño < 0 Then Daño = 0
    
        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If
        
        Call SubirSkill(tempChr, Resistencia)
        Call InfoHechizo(UserIndex)
    
        UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - Daño
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbRed & "°- " & Daño & "!" & "°" & str(UserList( _
            tempChr).char.CharIndex))
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de vida a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de vida." & _
            FONTTYPE_FIGHT)
    
        'Muere
        If UserList(tempChr).Stats.MinHP < 1 Then
            Call ContarMuerte(tempChr, UserIndex)
            UserList(tempChr).Stats.MinHP = 0
            Call ActStats(tempChr, UserIndex)
            
            MuereSpell = Hechizos(H).FXgrh
            LoopSpell = Hechizos(H).loops
            
            Call UserDie(tempChr)

        End If
    
        b = True

    End If

    'Mana
    If Hechizos(H).SubeMana = 1 Then
    
        Call InfoHechizo(UserIndex)

        Daño = RandomNumber(Hechizos(H).MinMana, Hechizos(H).ManMana)
        
        UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + Daño
        

        If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN
     
        If UserIndex <> tempChr Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de mana a " & UserList(tempChr).Name & _
                FONTTYPE_FIGHT)
            Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de mana." & _
                FONTTYPE_FIGHT)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)

        End If
    
        b = True
    
    ElseIf Hechizos(H).SubeMana = 2 Then

        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If
    
        Call InfoHechizo(UserIndex)
    
        If UserIndex <> tempChr Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de mana a " & UserList(tempChr).Name & _
                FONTTYPE_FIGHT)
            Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de mana." & _
                FONTTYPE_FIGHT)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Te has quitado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)

        End If
    
        UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - Daño

        If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
        b = True
    
    End If

    'Stamina
    If Hechizos(H).SubeSta = 1 Then
        Call InfoHechizo(UserIndex)
        UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + Daño

        If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta

        If UserIndex <> tempChr Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de vitalidad a " & UserList(tempChr).Name & _
                FONTTYPE_FIGHT)
            Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de vitalidad." & _
                FONTTYPE_FIGHT)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)

        End If

        b = True

    End If

End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

    'Call LogTarea("Sub UpdateUserHechizos")

    Dim loopc As Byte

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
        For loopc = 1 To MAXUSERHECHIZOS

            'Actualiza el inventario
            If UserList(UserIndex).Stats.UserHechizos(loopc) > 0 Then
                Call ChangeUserHechizo(UserIndex, loopc, UserList(UserIndex).Stats.UserHechizos(loopc))
            Else
                Call ChangeUserHechizo(UserIndex, loopc, 0)

            End If

        Next loopc

    End If

End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

    'Call LogTarea("ChangeUserHechizo")

    UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo

    If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then

        Call SendData(SendTarget.toIndex, UserIndex, 0, "SHS" & Slot & "," & Hechizo & "," & Hechizos(Hechizo).nombre)

    Else

        Call SendData(SendTarget.toIndex, UserIndex, 0, "SHS" & Slot & "," & "0" & "," & "(Vacío)")

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

    If Dire = 1 Then 'Mover arriba
        If CualHechizo = 1 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes mover el hechizo en esa direccion." & FONTTYPE_INFO)
            Exit Sub
        Else
            TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo
        
            Call UpdateUserHechizos(False, UserIndex, CualHechizo - 1)

        End If

    Else 'mover abajo

        If CualHechizo = MAXUSERHECHIZOS Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes mover el hechizo en esa direccion." & FONTTYPE_INFO)
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
    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).trigger = 6 Then Exit Sub
    
    'pierdo nobleza...
    UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep - NoblePts

    If UserList(UserIndex).Reputacion.NobleRep < 0 Then
        UserList(UserIndex).Reputacion.NobleRep = 0

    End If
    
    'gano bandido...
    UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep + BandidoPts

    If UserList(UserIndex).Reputacion.BandidoRep > MAXREP Then UserList(UserIndex).Reputacion.BandidoRep = MAXREP
    Call SendData(SendTarget.toIndex, UserIndex, 0, "PN")

    'If Criminal(UserIndex) Then
    '    If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    '        Call ExpulsarFaccionReal(UserIndex)
    '    End If
    '    If UserList(UserIndex).Faccion.Templario = 1 Then
    '        Call ExpulsarFaccionTemplario(UserIndex)
    '    End If
    'End If

End Sub
