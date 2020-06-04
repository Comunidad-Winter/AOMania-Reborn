Attribute VB_Name = "SistemaCombate"
Option Explicit

Public Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer

    If a > b Then
        MinimoInt = b
    Else
        MinimoInt = a

    End If

End Function

Public Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer

    If a > b Then
        MaximoInt = a
    Else
        MaximoInt = b

    End If

End Function

Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long

    With UserList(UserIndex)
        PoderEvasionEscudo = (.Stats.UserSkills(eSkill.Defensa) * ModClase(ClaseToByte(.Clase)).Escudo) * 0.5

    End With

End Function

Function PoderEvasion(ByVal UserIndex As Integer) As Long

    Dim lTemp As Long

    With UserList(UserIndex)
        lTemp = (.Stats.UserSkills(eSkill.Tacticas) + .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * _
                ModClase(ClaseToByte(.Clase)).Evasion

        PoderEvasion = lTemp + 2.5 * MaximoInt(.Stats.ELV - 12, 0)

    End With

End Function

Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long

    Dim PoderAtaqueTemp As Long
    Dim Skill As Byte
    Dim Modificador As Double

    With UserList(UserIndex)

        Skill = .Stats.UserSkills(eSkill.Armas)
        Modificador = ModClase(ClaseToByte(.Clase)).AtaqueArmas

        If Skill < 31 Then
            PoderAtaqueTemp = Skill * Modificador
        ElseIf Skill < 61 Then
            PoderAtaqueTemp = (Skill + .Stats.UserAtributos(eAtributos.Agilidad)) * Modificador
        ElseIf Skill < 91 Then
            PoderAtaqueTemp = (Skill + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * Modificador
        Else
            PoderAtaqueTemp = (Skill + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * Modificador
        End If

        PoderAtaqueArma = PoderAtaqueTemp + 2.5 * MaximoInt(.Stats.ELV - 12, 0)

    End With

End Function

Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long

    Dim PoderAtaqueTemp As Long
    Dim Skill As Byte
    Dim Modificador As Double

    With UserList(UserIndex)

        Skill = .Stats.UserSkills(eSkill.Proyectiles)
        Modificador = ModClase(ClaseToByte(.Clase)).AtaqueProyectiles

        If Skill < 31 Then
            PoderAtaqueTemp = Skill * Modificador
        ElseIf Skill < 61 Then
            PoderAtaqueTemp = (Skill + .Stats.UserAtributos(eAtributos.Agilidad)) * Modificador
        ElseIf Skill < 91 Then
            PoderAtaqueTemp = (Skill + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * Modificador
        Else
            PoderAtaqueTemp = (Skill + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * Modificador
        End If

        PoderAtaqueProyectil = PoderAtaqueTemp + 2.5 * MaximoInt(.Stats.ELV - 12, 0)

    End With

End Function

Function PoderAtaqueWresterling(ByVal UserIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    Dim Skill As Byte
    Dim Modificador As Double

    With UserList(UserIndex)

        Skill = .Stats.UserSkills(eSkill.Wresterling)
        Modificador = ModClase(ClaseToByte(.Clase)).AtaqueWrestling

        If Skill < 31 Then
            PoderAtaqueTemp = Skill * Modificador
        ElseIf Skill < 61 Then
            PoderAtaqueTemp = Skill + .Stats.UserAtributos(eAtributos.Agilidad) * Modificador
        ElseIf Skill < 91 Then
            PoderAtaqueTemp = (Skill + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * Modificador
        Else
            PoderAtaqueTemp = (Skill + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * Modificador
        End If

        PoderAtaqueWresterling = PoderAtaqueTemp + 2.5 * MaximoInt(.Stats.ELV - 12, 0)

    End With

End Function

Public Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean

    Dim PoderAtaque As Long
    Dim Arma As Integer
    Dim Skill As eSkill
    Dim ProbExito As Long

    Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex

    If Arma > 0 Then    'Usando un arma
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(UserIndex)
            Skill = eSkill.Proyectiles
        Else
            PoderAtaque = PoderAtaqueArma(UserIndex)
            Skill = eSkill.Armas

        End If

    Else    'Peleando con puños
        PoderAtaque = PoderAtaqueWresterling(UserIndex)
        Skill = eSkill.Wresterling

    End If

    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))

    UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

    If UserImpactoNpc Then
        Call SubirSkill(UserIndex, Skill)
    Else
        Call SubirSkill(UserIndex, Skill)

    End If

End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Revisa si un NPC logra impactar a un user o no
'03/15/2006 Maraxus - Evité una división por cero que eliminaba NPCs
'*************************************************
    Dim Rechazo As Boolean
    Dim ProbRechazo As Long
    Dim ProbExito As Long
    Dim UserEvasion As Long
    Dim NpcPoderAtaque As Long
    Dim PoderEvasioEscudo As Long
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long

    UserEvasion = PoderEvasion(UserIndex)
    NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
    PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)

    SkillTacticas = UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(UserIndex).Stats.UserSkills(eSkill.Defensa)

    'Esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo

    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))

    NpcImpacto = (RandomNumber(1, 100) <= ProbExito)

    ' el usuario esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        If Not NpcImpacto Then
            If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
                ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
                Rechazo = (RandomNumber(1, 100) <= ProbRechazo)

                If Rechazo = True Then
                    'Se rechazo el ataque con el escudo
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_ESCUDO)
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "EW" & UserList(UserIndex).char.CharIndex)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "7")
                    Call SubirSkill(UserIndex, Defensa)

                End If

            End If

        End If

    End If

End Function

Public Function CalcularDaño(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long

    Dim DañoArma As Long, DañoUsuario As Long, Arma As ObjData, ModifClase As Single

    Dim proyectil As ObjData

    Dim DañoMaxArma As Long

    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        Arma = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex)
    
        ' Ataca a un npc?
        If NpcIndex > 0 Then
        
            'Usa la mata dragones?
            If Arma.Subtipo = eSubtipo.otMatadragones Then ' Usa la matadragones?
                ModifClase = modDañoArma(UserList(UserIndex).Clase)

                If Npclist(NpcIndex).NPCtype = eNPCType.DRAGON Then 'Ataca dragon?
                    DañoArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                    DañoMaxArma = Arma.MaxHit
                Else ' Si no es dragon daño es 1
                    DañoArma = 1
                    DañoMaxArma = 1

                End If

            Else ' daño comun
        
                If Arma.proyectil = 1 Then
                    If UserList(UserIndex).Sagrada.Enabled = 0 Then
                        ModifClase = modDañoProyectil(UserList(UserIndex).Clase)
                        DañoArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                        DañoMaxArma = Arma.MaxHit

                        If Arma.Municion = 1 Then
                            proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                            DañoArma = DañoArma + RandomNumber(proyectil.MinHit, proyectil.MaxHit)
                            DañoMaxArma = Arma.MaxHit

                        End If

                    ElseIf UserList(UserIndex).Sagrada.Enabled = 1 Then
                        ModifClase = modDañoProyectil(UserList(UserIndex).Clase)
                        DañoArma = RandomNumber(UserList(UserIndex).Sagrada.MinHit, UserList(UserIndex).Sagrada.MaxHit)
                        DañoMaxArma = UserList(UserIndex).Sagrada.MaxHit

                        If Arma.Municion = 1 Then
                            proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                            DañoArma = DañoArma + RandomNumber(proyectil.MinHit, proyectil.MaxHit)
                            DañoMaxArma = Arma.MaxHit

                        End If

                    End If
                
                Else

                    If UserList(UserIndex).Sagrada.Enabled = 0 Then
                        ModifClase = modDañoArma(UserList(UserIndex).Clase)
                        DañoArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                        DañoMaxArma = Arma.MaxHit
                    ElseIf UserList(UserIndex).Sagrada.Enabled = 1 Then
                        ModifClase = modDañoArma(UserList(UserIndex).Clase)
                        DañoArma = RandomNumber(UserList(UserIndex).Sagrada.MinHit, UserList(UserIndex).Sagrada.MaxHit)
                        DañoMaxArma = UserList(UserIndex).Sagrada.MaxHit

                    End If

                End If

            End If
    
        Else ' Ataca usuario

            If Arma.Subtipo = eSubtipo.otMatadragones Then
                ModifClase = modDañoArma(UserList(UserIndex).Clase)
                DañoArma = 1 ' Si usa la espada matadragones daño es 1
                DañoMaxArma = 1
            Else

                If Arma.proyectil = 1 Then
                    If UserList(UserIndex).Sagrada.Enabled = 0 Then
                        ModifClase = modDañoProyectil(UserList(UserIndex).Clase)
                        DañoArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                        DañoMaxArma = Arma.MaxHit

                        If Arma.Municion = 1 Then
                            proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                            DañoArma = DañoArma + RandomNumber(proyectil.MinHit, proyectil.MaxHit)
                            DañoMaxArma = Arma.MaxHit

                        End If

                    ElseIf UserList(UserIndex).Sagrada.Enabled = 1 Then
                        ModifClase = modDañoProyectil(UserList(UserIndex).Clase)
                        DañoArma = RandomNumber(UserList(UserIndex).Sagrada.MinHit, UserList(UserIndex).Sagrada.MaxHit)
                        DañoMaxArma = UserList(UserIndex).Sagrada.MaxHit

                        If Arma.Municion = 1 Then
                            proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                            DañoArma = DañoArma + RandomNumber(proyectil.MinHit, proyectil.MaxHit)
                            DañoMaxArma = Arma.MaxHit

                        End If

                    End If

                Else

                    If UserList(UserIndex).Sagrada.Enabled = 0 Then
                        ModifClase = modDañoArma(UserList(UserIndex).Clase)
                        DañoArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                        DañoMaxArma = Arma.MaxHit
                    ElseIf UserList(UserIndex).Sagrada.Enabled = 1 Then
                        ModifClase = modDañoArma(UserList(UserIndex).Clase)
                        DañoArma = RandomNumber(UserList(UserIndex).Sagrada.MinHit, UserList(UserIndex).Sagrada.MaxHit)
                        DañoMaxArma = UserList(UserIndex).Sagrada.MaxHit

                    End If

                End If

            End If

        End If

    End If

    DañoUsuario = RandomNumber(UserList(UserIndex).Stats.MinHit, UserList(UserIndex).Stats.MaxHit)
    
    If UserList(UserIndex).GranPoder = 1 And GranPoder.TipoAura = hGranPoder.Daño Then
         CalcularDaño = (((3 * DañoArma) + ((DañoMaxArma / 5) * Maximo(0, (UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) - 15))) + DañoUsuario) * ModifClase) * 2
         Else
         CalcularDaño = (((3 * DañoArma) + ((DañoMaxArma / 5) * Maximo(0, (UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) - 15))) + DañoUsuario) * ModifClase)
    End If
    
    'Debug.Print "FUERZA: " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza)
End Function

Function Maximo(ByVal a As Long, ByVal b As Long) As Long
If a > b Then
    Maximo = a
    Else: Maximo = b
End If
End Function

Function modDañoArma(ByVal Clase As String) As Single
    Select Case UCase$(Clase)
        Case "GUERRERO", "GLADIADOR MAGICO"
            modDañoArma = 1.1
        Case "ASESINO"
            modDañoArma = 0.9
        Case "PALADIN"
            modDañoArma = 0.88
        Case "PIRATA", "HERRERO MAGICO"
            modDañoArma = 0.8
        Case "LADRON", "THESAUROS", "CLERIGO"
            modDañoArma = 0.75
        Case "CAZADOR", "BARDO", "DRUIDA", "ARQUERO"
            modDañoArma = 0.7
        Case "HERRERO"
            modDañoArma = 0.65
        Case "MAGO", "BRUJO"
            modDañoArma = 0.6
        Case Else
            modDañoArma = 0.5
    End Select
End Function

Function modDañoProyectil(ByVal Clase As String) As Single
    Select Case UCase$(Clase)
        Case "ARQUERO"
            modDañoProyectil = 1.5
        Case "CAZADOR", "GLADIADOR MAGICO"
            modDañoProyectil = 1.2
        Case "GUERRERO"
            modDañoProyectil = 0.9
        Case "PALADIN"
            modDañoProyectil = 0.8
        Case "ASESINO", "PIRATA", "HERRERO MAGICO"
            modDañoProyectil = 0.75
        Case "CLERIGO", "BARDO", "LADRON", "THESAUROS"
            modDañoProyectil = 0.7
        Case "HERRERO", "MINERO"
            modDañoProyectil = 0.65
        Case "MAGO", "BRUJO", "DRUIDA"
            modDañoProyectil = 0.6
        Case Else
            modDañoProyectil = 0.5
    End Select
End Function

Public Sub UserDañoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    Dim Daño As Long
    Dim TeCritico As Byte
    Dim Arma As Long
    Dim Municion As Long

    TeCritico = RandomNumber(1, 8)

    Daño = CalcularDaño(UserIndex, NpcIndex)

    If Npclist(NpcIndex).Numero = 616 Then
        If Npclist(NpcIndex).Stats.MinHP > 14000 Then
            Call SendData(ToAll, 0, 0, "LEMU")

        End If

        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
            Call SendData(ToAll, 0, 0, "LEMU")

        End If

    End If

    If Npclist(NpcIndex).Numero = 617 Then
        If Npclist(NpcIndex).Stats.MinHP > 14000 Then
            Call SendData(ToAll, 0, 0, "TALE")

        End If

        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
            Call SendData(ToAll, 0, 0, "TALE")

        End If

    End If

    If Npclist(NpcIndex).Numero = 910 Then
        If Npclist(NpcIndex).Stats.MinHP > 14000 Then
            Call SendData(ToAll, 0, 0, "NIX")

        End If

        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
            Call SendData(ToAll, 0, 0, "NIX")

        End If

    End If

    'Peto
    'esta navegando? si es asi le sumamos el daño del barco
    If UserList(UserIndex).flags.Navegando = 1 Then Daño = Daño + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHit, ObjData( _
                                                                                                                                         UserList(UserIndex).Invent.BarcoObjIndex).MaxHit)

    Daño = Daño - Npclist(NpcIndex).Stats.def

    If Daño < 0 Then Daño = 0

    If UserList(UserIndex).pos.Map = MapaCasaAbandonada1 Then
        Call Efecto_AccionCasaEncantada(UserIndex, NpcIndex)
    End If

    ' animacion daño sobre 100
    If Daño >= 100 Then
        'Call SendData(SendTarget.ToNPCArea, NpcIndex, Npclist(NpcIndex).pos.Map, "CFX" & Npclist(NpcIndex).char.CharIndex & "," & 38 & "," & 0)
    End If

    ' animacion daño bajo 100
    If Daño < 100 Then
        'Call SendData(SendTarget.ToNPCArea, NpcIndex, Npclist(NpcIndex).pos.Map, "CFX" & Npclist(NpcIndex).char.CharIndex & "," & 14 & "," & 0)
    End If

    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
        If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
            Municion = UserList(UserIndex).Invent.MunicionEqpObjIndex
            If ObjData(Municion).Paraliza > 0 And Npclist(NpcIndex).flags.Paralizado = 0 Then
                If RandomNumber(1, 100) <= ObjData(Municion).Paraliza Then
                    Npclist(NpcIndex).flags.Paralizado = 1
                    Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
                End If
            End If
        Else
            If ObjData(Arma).Paraliza > 0 And Npclist(NpcIndex).flags.Paralizado = 0 Then
                If RandomNumber(1, 100) <= ObjData(Arma).Paraliza Then
                    Npclist(NpcIndex).flags.Paralizado = 1
                    Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
                End If
            End If
        End If
    End If

    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - Daño

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Le has pegado a la criatura por " & Daño & " !!" & FONTTYPE_Motd4)

    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbCyan & "°" & Daño & "°" & CStr(Npclist( _
                                                                                                                       NpcIndex).char.CharIndex))

    If Npclist(NpcIndex).Stats.MinHP < 0 Then
        Daño = Npclist(NpcIndex).Stats.MinHP + Daño
        Npclist(NpcIndex).Stats.MinHP = 0
    End If

    Call CalcularDarExp(UserIndex, NpcIndex, Daño)

    If Npclist(NpcIndex).Stats.MinHP > 0 Then

        'Trata de apuñalar por la espalda al enemigo
        If PuedeApuñalar(UserIndex) Then
            Call DoApuñalar(UserIndex, NpcIndex, 0, Daño)
            Call SubirSkill(UserIndex, Apuñalar)
        End If

        'Mascotas atacan a la criatura.
        'Call CheckPets(NpcIndex, UserIndex, True)

    End If

    If Npclist(NpcIndex).Stats.MinHP <= 0 Then

        ' Para que las mascotas no sigan intentando luchar y
        ' comiencen a seguir al amo

        Dim j As Integer

        For j = 1 To MAXMASCOTAS

            If UserList(UserIndex).MascotasIndex(j) > 0 Then
                If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNpc = NpcIndex Then Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNpc = 0
                Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.SigueAmo

            End If

        Next j

        Npclist(NpcIndex).Stats.MinHP = 0

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " le queda " & Npclist(NpcIndex).Stats.MinHP & " / " & Npclist( _
                                                        NpcIndex).Stats.MaxHP & " de vida." & FONTTYPE_Motd4)
        Call MuereNpc(NpcIndex, UserIndex)

    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " le queda " & Npclist(NpcIndex).Stats.MinHP & " / " & Npclist( _
                                                        NpcIndex).Stats.MaxHP & " de vida." & FONTTYPE_Motd4)
    End If

End Sub

Public Sub NpcDaño(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

    Dim Daño As Integer, Lugar As Integer, absorbido As Integer, npcfile As String
    Dim antdaño As Integer, defbarco As Integer, defArmadura As Integer, defEscudo As Integer
    Dim Obj As ObjData, AmuletoDaño As Integer

    Daño = RandomNumber(Npclist(NpcIndex).Stats.MinHit, Npclist(NpcIndex).Stats.MaxHit)
    antdaño = Daño

    If UserList(UserIndex).flags.Navegando = 1 Then
        Obj = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
        defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)

    End If

    Lugar = RandomNumber(1, 6)

    Select Case Lugar

    Case PartesCuerpo.bCabeza

        'Si tiene casco absorbe el golpe
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
            Obj = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex)
            absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
            absorbido = absorbido + defbarco
            Daño = Daño - absorbido

            If Daño < 1 Then Daño = 1

        End If

    Case Else

        'Si tiene armadura absorbe el golpe
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            Obj = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex)
            defArmadura = RandomNumber(Obj.MinDef, Obj.MaxDef)

        End If

        'Si tiene escudo absorbe el golpe
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
            Obj = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex)
            defEscudo = RandomNumber(Obj.MinDef, Obj.MaxDef)

        End If

        If defArmadura > 0 And defEscudo > 0 Then
            absorbido = Int((defArmadura + defEscudo) * 0.5) + defbarco
        ElseIf defArmadura > 0 And defEscudo <= 0 Then
            absorbido = defArmadura + defbarco
        ElseIf defArmadura <= 0 And defEscudo > 0 Then
            absorbido = defEscudo + defbarco
        End If

        If UserList(UserIndex).Invent.AmuletoEqpObjIndex > 0 Then
            If ObjData(UserList(UserIndex).Invent.AmuletoEqpObjIndex).AmuletoDefensa.TipoBonifica = eAmuleto.otFisico Then
                AmuletoDaño = RandomNumber(1, ObjData(UserList(UserIndex).Invent.AmuletoEqpObjIndex).AmuletoDefensa.Bonifica)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu Amuleto te ha protegido de " & AmuletoDaño & " puntos de Daño." & FONTTYPE_TALKMSG)

            End If
        End If

        Daño = (Daño - absorbido) - AmuletoDaño

        Select Case UCase$(UserList(UserIndex).Clase)
        Case "GUERRERO"
            Daño = Porcentaje(Daño, "50")
        Case "ARQUERO"
            Daño = Porcentaje(Daño, "30")
        End Select

        If Daño < 1 Then Daño = 1

    End Select

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "N2" & Lugar & "," & Daño)

    If UserList(UserIndex).flags.Privilegios = PlayerType.User Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - Daño

    'Muere el usuario
    If UserList(UserIndex).Stats.MinHP <= 0 Then

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "6")    ' Le informamos que ha muerto ;)

        'Si lo mato un guardia
        If Criminal(UserIndex) And Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            Call RestarCriminalidad(UserIndex)

            'If Not Criminal(UserIndex) Then
            '    If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
            '        Call ExpulsarFaccionCaos(UserIndex)
            '    End If

            '    If UserList(UserIndex).Faccion.Nemesis = 1 Then
            '        Call ExpulsarFaccionNemesis(UserIndex)
            '    End If
            'End If

        End If

        If Npclist(NpcIndex).MaestroUser > 0 Then
            Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
        Else

            'Al matarlo no lo sigue mas
            If Npclist(NpcIndex).Stats.Alineacion = 0 Then
                Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
                Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
                Npclist(NpcIndex).flags.AttackedBy = ""

            End If

        End If

        Call UserDie(UserIndex)

    End If

End Sub

Public Sub RestarCriminalidad(ByVal UserIndex As Integer)

    If UserList(UserIndex).Reputacion.AsesinoRep > 0 Then
        UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep - vlASESINO
        If UserList(UserIndex).Reputacion.AsesinoRep < 0 Then UserList(UserIndex).Reputacion.AsesinoRep = 0

    ElseIf UserList(UserIndex).Reputacion.BandidoRep > 0 Then
        UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep - vlASALTO
        If UserList(UserIndex).Reputacion.BandidoRep < 0 Then UserList(UserIndex).Reputacion.BandidoRep = 0

    ElseIf UserList(UserIndex).Reputacion.LadronesRep > 0 Then
        UserList(UserIndex).Reputacion.LadronesRep = UserList(UserIndex).Reputacion.LadronesRep - (vlCAZADOR * 10)
        If UserList(UserIndex).Reputacion.LadronesRep < 0 Then UserList(UserIndex).Reputacion.LadronesRep = 0
    End If

End Sub

Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, Optional ByVal CheckElementales As Boolean = True)

    Dim j As Long, Mascota As Integer

    With UserList(UserIndex)

        For j = 1 To MAXMASCOTAS

            Mascota = .MascotasIndex(j)

            If Mascota > 0 Then
                If Mascota <> NpcIndex Then
                    If CheckElementales Or (Npclist(Mascota).Numero <> ELEMENTALFUEGO And Npclist(Mascota).Numero <> ELEMENTALTIERRA) Then

                        If Npclist(Mascota).TargetNpc = 0 Then Npclist(Mascota).TargetNpc = NpcIndex
                        Npclist(Mascota).Movement = TipoAI.NpcAtacaNpc

                    End If

                End If

            End If

        Next j

    End With

End Sub

Public Sub AllFollowAmo(ByVal UserIndex As Integer)

    Dim j As Integer

    For j = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasIndex(j) > 0 Then
            Call FollowAmo(UserList(UserIndex).MascotasIndex(j))

        End If

    Next j

End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

    If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function

    If Npclist(NpcIndex).Numero = 253 And UserList(UserIndex).flags.Demonio Then
        Exit Function
    End If

    If Npclist(NpcIndex).Numero = 254 And UserList(UserIndex).flags.Angel Then
        Exit Function
    End If

    ' El npc puede atacar ???
    If Npclist(NpcIndex).CanAttack = 1 Then
        NpcAtacaUser = True
        Call CheckPets(NpcIndex, UserIndex, False)

        If Npclist(NpcIndex).Target = 0 Then Npclist(NpcIndex).Target = UserIndex

        If UserList(UserIndex).flags.AtacadoPorNpc = 0 And UserList(UserIndex).flags.AtacadoPorUser = 0 Then UserList( _
           UserIndex).flags.AtacadoPorNpc = NpcIndex
    Else
        NpcAtacaUser = False
        Exit Function

    End If

    Npclist(NpcIndex).CanAttack = 0

    If Npclist(NpcIndex).flags.Snd1 > 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & Npclist( _
                                                                                                                        NpcIndex).flags.Snd1)

    If NpcImpacto(NpcIndex, UserIndex) Then
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_IMPACTO)

        If UserList(UserIndex).flags.Meditando = False Then
            If UserList(UserIndex).flags.Navegando = 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & _
                                                                                                                                       UserList(UserIndex).char.CharIndex & "," & FXSANGRE & "," & 0)

        End If

        Call NpcDaño(NpcIndex, UserIndex)

        If Npclist(NpcIndex).Name = "Tornado" Then
            Call NpcTeleport(UserIndex)

        End If

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ASH" & UserList(UserIndex).Stats.MinHP)

        '¿Puede envenenar?
        If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(UserIndex)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "N1")

    End If

    '-----Tal vez suba los skills------
    Call SubirSkill(UserIndex, Tacticas)

    'Controla el nivel del usuario
    Call CheckUserLevel(UserIndex)
    Call EnviarHP(UserIndex)

End Function

Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
    Dim PoderAtt As Long, PoderEva As Long, dif As Long
    Dim ProbExito As Long

    PoderAtt = Npclist(Atacante).PoderAtaque
    PoderEva = Npclist(Victima).PoderEvasion
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtt - PoderEva) * 0.4)))
    NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

End Function

Public Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)

    Dim Daño As Long

    Dim ANpc As npc, DNpc As npc

    ANpc = Npclist(Atacante)

    Daño = RandomNumber(ANpc.Stats.MinHit, ANpc.Stats.MaxHit)
    Npclist(Victima).Stats.MinHP = Npclist(Victima).Stats.MinHP - Daño
    Call CalcularDarExp(Npclist(Atacante).MaestroUser, Victima, Daño)

    If Npclist(Victima).Stats.MinHP < 1 Then Npclist(Victima).Stats.MinHP = 0
    Call SendData(ToIndex, Npclist(Atacante).MaestroUser, 0, "||" & Npclist(Victima).Name & " le queda " & Npclist(Victima).Stats.MinHP & "/" & Npclist(Victima).Stats.MaxHP & " de vida." & FONTTYPE_Motd4)

    Npclist(Victima).Stats.MinHP = Npclist(Victima).Stats.MinHP - Daño

    If Npclist(Victima).Stats.MinHP < 1 Then
        
        If Npclist(Atacante).flags.AttackedBy <> "" Then
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
            Npclist(Atacante).Hostile = Npclist(Atacante).flags.OldHostil
        Else
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement

        End If
        
        Call FollowAmo(Atacante)
        
        Call MuereNpc(Victima, Npclist(Atacante).MaestroUser)

    End If

End Sub

Function ProtecNPC(ByVal npc As Integer) As Boolean

    Select Case npc
    Case ELEMENTALTORTUGA, ELEMENTALFUEGO, ELEMENTALFUEGOII, ELEMENTALTIERRA, ELEMENTALTIERRAII, ELEMENTALTIERRAM, ELEMENTALAGUA, ELEMENTALAGUAII, ELEMENTALVIENTO, ELEMENTALTEMPLARIO, ELEMENTALFATUO
        ProtecNPC = True
        Exit Function
    End Select

    ProtecNPC = False

End Function

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)

    ' El npc puede atacar ???
    If Npclist(Atacante).CanAttack = 1 Then
        Npclist(Atacante).CanAttack = 0
        Npclist(Victima).TargetNpc = Atacante
        Npclist(Victima).Movement = TipoAI.NpcAtacaNpc
    Else
        Exit Sub
    End If

    If Npclist(Atacante).flags.Snd1 > 0 Then Call SendData(SendTarget.ToNPCArea, Atacante, Npclist(Atacante).pos.Map, "TW" & Npclist( _
                                                                                                                      Atacante).flags.Snd1)

    If NpcImpactoNpc(Atacante, Victima) Then

        If Npclist(Victima).flags.Snd2 > 0 Then
            Call SendData(ToMap, Victima, Npclist(Victima).pos.Map, "TW" & Npclist(Victima).flags.Snd2)
        Else
            Call SendData(ToMap, Victima, Npclist(Victima).pos.Map, "TW" & SND_IMPACTO2)

        End If

        If Npclist(Atacante).MaestroUser > 0 Then
            Call SendData(ToMap, Atacante, Npclist(Atacante).pos.Map, "TW" & SND_IMPACTO)
        Else
            Call SendData(ToMap, Victima, Npclist(Victima).pos.Map, "TW" & SND_IMPACTO)

        End If

        Call NpcDañoNpc(Atacante, Victima)

    Else

        If Npclist(Atacante).MaestroUser > 0 Then
            Call SendData(ToMap, Atacante, Npclist(Atacante).pos.Map, "TW" & SND_SWING)
        Else
            Call SendData(ToMap, Victima, Npclist(Victima).pos.Map, "TW" & SND_SWING)

        End If

    End If
    
   ' Call SendData(SendTarget.ToIndex, Npclist(Atacante).MaestroUser, 0, "||" & Npclist(Victima).Name & " le queda " & Npclist(Victima).Stats.MinHP & " / " & Npclist( _
                                                            Victima).Stats.MaxHP & " de vida." & FONTTYPE_Motd4)

End Sub

Public Sub UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    Call RestarCriminalidad(UserIndex)

    If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then Exit Sub

    If Abs(Npclist(NpcIndex).pos.X - UserList(UserIndex).pos.X) > RANGO_VISION_X Or Abs(Npclist(NpcIndex).pos.Y - UserList(UserIndex).pos.Y) > RANGO_VISION_Y Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estás muy lejos para disparar." & FONTTYPE_INFO)
        Exit Sub

    End If
    
    If UserList(UserIndex).flags.Silenciado = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||Estás Silenciado, No puedes atacar a Ningun NPC." & FONTTYPE_FIGHT)
                Exit Sub
            End If

    If Not GolpeNpcCastillo(UserIndex, NpcIndex) Then Exit Sub

    If UserList(UserIndex).flags.Demonio = True Then
        If Npclist(NpcIndex).Numero = 253 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar a la medusa de tu equipo ¬¬" & FONTTYPE_INFO)
            Exit Sub
        End If
    End If

    If UserList(UserIndex).flags.Angel = True Then
        If Npclist(NpcIndex).Numero = 254 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar a la medusa de tu equipo ¬¬" & FONTTYPE_INFO)
            Exit Sub
        End If
    End If

    If UserList(UserIndex).flags.Corsarios = True Then
        If Npclist(NpcIndex).Numero = NpcCorsarios Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar a la medusa de tu equipo ¬¬" & FONTTYPE_INFO)
            Exit Sub
        End If
    End If

    If UserList(UserIndex).flags.Piratas = True Then
        If Npclist(NpcIndex).Numero = NpcPiratas Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar a la medusa de tu equipo ¬¬" & FONTTYPE_INFO)
            Exit Sub
        End If
    End If

    Call NpcAtacado(NpcIndex, UserIndex)

    If UserImpactoNpc(UserIndex, NpcIndex) Then

        If Npclist(NpcIndex).flags.Snd2 > 0 Then
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & Npclist(NpcIndex).flags.Snd2)

        Else
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_IMPACTO2)

        End If

        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "FG" & UserList(UserIndex).char.CharIndex)
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "SFX" & Npclist(NpcIndex).char.CharIndex & "-1") ' kalii

        Call UserDañoNpc(UserIndex, NpcIndex)

    Else
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_SWING)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "FG" & UserList(UserIndex).char.CharIndex)

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "U1")
        'Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbRed & "°" & "Fallé" & "!" & "°" & str(UserList( _
         UserIndex).char.CharIndex))

    End If

End Sub

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)

'If UserList(UserIndex).flags.PuedeAtacar = 1 Then
    If IntervaloPermiteAtacar(UserIndex) Then

        'Quitamos stamina
        If UserList(UserIndex).Stats.MinSta >= 10 Then
            Call QuitarSta(UserIndex, RandomNumber(1, 10))
            Call EnviarSta(UserIndex)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas muy cansado para luchar." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(UserIndex).flags.Privilegios = PlayerType.User Then

            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Gm = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar con items de los GameMasters." & FONTTYPE_INFO)
                    Exit Sub
                End If
            End If

            If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex).Gm = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar con items de los GameMasters." & FONTTYPE_INFO)
                    Exit Sub
                End If
            End If

            If UserList(UserIndex).Invent.AlaEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.AlaEqpObjIndex).Gm = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar con items de los GameMasters." & FONTTYPE_INFO)
                    Exit Sub
                End If
            End If

            If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Gm = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar con items de los GameMasters." & FONTTYPE_INFO)
                    Exit Sub
                End If
            End If

            If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).Gm = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar con items de los GameMasters." & FONTTYPE_INFO)
                    Exit Sub
                End If
            End If

            If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).Gm = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar con items de los GameMasters." & FONTTYPE_INFO)
                    Exit Sub
                End If
            End If

            If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).Gm = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar con items de los GameMasters." & FONTTYPE_INFO)
                    Exit Sub
                End If
            End If

        End If

        'UserList(UserIndex).flags.PuedeAtacar = 0

        Dim AttackPos As WorldPos
        AttackPos = UserList(UserIndex).pos
        Call HeadtoPos(UserList(UserIndex).char.heading, AttackPos)

        'Exit if not legal
        If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_SWING)
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "FG" & UserList(UserIndex).char.CharIndex)
            Exit Sub

        End If

        Dim Index As Integer
        Index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex

        'Look for user
        If Index > 0 Then
            Call UsuarioAtacaUsuario(UserIndex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex)
            Call EnviarHP(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex)
            Exit Sub

        End If

        'Look for NPC
        If MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex > 0 Then

            If Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).Attackable Then

                If Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).MaestroUser > 0 And MapInfo(Npclist(MapData(AttackPos.Map, _
                                                                                                                                  AttackPos.X, AttackPos.Y).NpcIndex).pos.Map).Pk = False Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar mascotas en zonas seguras" & FONTTYPE_INFO)
                    Exit Sub

                End If

                Call UsuarioAtacaNpc(UserIndex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex)

            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar a este NPC" & FONTTYPE_INFO)

            End If

            Exit Sub

        End If

        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_SWING)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "FG" & UserList(UserIndex).char.CharIndex)

    End If

    If UserList(UserIndex).Counters.Trabajando Then UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean

    Dim ProbRechazo As Long
    Dim Rechazo As Boolean
    Dim ProbExito As Long
    Dim PoderAtaque As Long
    Dim UserPoderEvasion As Long
    Dim UserPoderEvasionEscudo As Long
    Dim Arma As Integer
    Dim proyectil As Boolean
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long

    SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(eSkill.Defensa)

    Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex

    If Arma > 0 Then
        proyectil = ObjData(Arma).proyectil = 1
    Else
        proyectil = False

    End If

    'Calculamos el poder de evasion...
    UserPoderEvasion = PoderEvasion(VictimaIndex)

    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
        UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
        UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
    Else
        UserPoderEvasionEscudo = 0

    End If

    'Esta usando un arma ???
    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then

        If proyectil Then
            PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
        Else
            PoderAtaque = PoderAtaqueArma(AtacanteIndex)

        End If

        ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - UserPoderEvasion) * 0.4)))

    Else
        PoderAtaque = PoderAtaqueWresterling(AtacanteIndex)
        ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - UserPoderEvasion) * 0.4)))

    End If

    UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)

    ' el usuario esta usando un escudo ???
    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then

        'Fallo ???
        If UsuarioImpacto = False Then
            ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)

            If Rechazo = True Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "TW" & SND_ESCUDO)
                Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "EW" & UserList(VictimaIndex).char.CharIndex)
                Call SendData(SendTarget.ToIndex, AtacanteIndex, 0, "8")
                Call SendData(SendTarget.ToIndex, VictimaIndex, 0, "7")
                Call SubirSkill(VictimaIndex, Defensa)

            End If

        End If

    End If

    If UsuarioImpacto Then
        If Arma > 0 Then
            If Not proyectil Then
                Call SubirSkill(AtacanteIndex, Armas)
            Else
                Call SubirSkill(AtacanteIndex, Proyectiles)

            End If

        Else
            Call SubirSkill(AtacanteIndex, Wresterling)

        End If

    End If

End Function

Public Sub UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)

    If UserList(AtacanteIndex).flags.Demonio = True And UserList(VictimaIndex).flags.Demonio = True Then
        Call SendData(SendTarget.ToIndex, AtacanteIndex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(AtacanteIndex).flags.Angel = True And UserList(VictimaIndex).flags.Angel = True Then
        Call SendData(SendTarget.ToIndex, AtacanteIndex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
        Exit Sub

    End If
    
    If UserList(AtacanteIndex).pos.Map = 192 Then
        If UserList(AtacanteIndex).flags.SuPareja = VictimaIndex Then Exit Sub
    End If

    If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Sub

    If Abs(UserList(AtacanteIndex).pos.X - UserList(VictimaIndex).pos.X) > RANGO_VISION_X Or Abs(UserList(AtacanteIndex).pos.Y - UserList(VictimaIndex).pos.Y) > RANGO_VISION_Y Then
        Call SendData(SendTarget.ToIndex, AtacanteIndex, 0, "||Estás muy lejos para disparar." & FONTTYPE_INFO)
        Exit Sub

    End If

    Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)

    If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
        Call SendData(SendTarget.ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "TW" & SND_IMPACTO)

        Call UserDañoUser(AtacanteIndex, VictimaIndex)
    Else
        Call SendData(SendTarget.ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "TW" & SND_SWING)
        Call SendData(SendTarget.ToIndex, AtacanteIndex, 0, "U1")
        'Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "||" & vbRed & "°" & "Fallé" & "!" & "°" & str(UserList( _
         AtacanteIndex).char.CharIndex))
        Call SendData(SendTarget.ToIndex, VictimaIndex, 0, "U3" & UserList(AtacanteIndex).Name)

    End If

    Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "FG" & UserList(AtacanteIndex).char.CharIndex)

    If UCase$(UserList(AtacanteIndex).Clase) = "LADRON" Then Call Desarmar(AtacanteIndex, VictimaIndex)

End Sub

Public Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)

    On Error GoTo ErrorHandler

    Dim Daño As Long, antdaño As Integer
    Dim Lugar As Integer, absorbido As Long
    Dim defbarco As Integer, defArmadura As Integer, defEscudo As Integer
    Dim TeCritico As Byte, AmuletoDaño As Integer
    Dim Arma As Long, Municion As Long

    TeCritico = RandomNumber(1, 10)

    Dim Obj As ObjData

    Daño = CalcularDaño(AtacanteIndex)
    antdaño = Daño

    If UserList(VictimaIndex).Invent.AmuletoEqpObjIndex > 0 Then
        If ObjData(UserList(VictimaIndex).Invent.AmuletoEqpObjIndex).AmuletoDefensa.TipoBonifica = eAmuleto.otFisico Then
            AmuletoDaño = RandomNumber(1, ObjData(UserList(VictimaIndex).Invent.AmuletoEqpObjIndex).AmuletoDefensa.Bonifica)
            Call SendData(SendTarget.ToIndex, VictimaIndex, 0, "||Tu Amuleto te ha protegido de " & AmuletoDaño & " puntos de Daño." & FONTTYPE_TALKMSG)
            Daño = Daño - AmuletoDaño
        End If
    End If

    If Daño >= 200 Then
        If UserList(VictimaIndex).flags.Navegando = 0 Then Call SendData(SendTarget.ToPCArea, VictimaIndex, UserList(VictimaIndex).pos.Map, "CFX" & _
                                                                                                                                            UserList(VictimaIndex).char.CharIndex & "," & 38 & "," & 0)
    End If

    If Daño < 200 Then
        If UserList(VictimaIndex).flags.Navegando = 0 Then Call SendData(SendTarget.ToPCArea, VictimaIndex, UserList(VictimaIndex).pos.Map, "CFX" & _
                                                                                                                                            UserList(VictimaIndex).char.CharIndex & "," & 14 & "," & 0)

    End If

    Call UserEnvenena(AtacanteIndex, VictimaIndex)

    If UserList(AtacanteIndex).flags.Navegando = 1 Then
        Obj = ObjData(UserList(AtacanteIndex).Invent.BarcoObjIndex)
        Daño = Daño + RandomNumber(Obj.MinHit, Obj.MaxHit)

    End If

    If UserList(VictimaIndex).flags.Navegando = 1 Then
        Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
        defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)

    End If

    '[MaTeO 9]
    If UserList(VictimaIndex).char.Alas > 0 Then
        Obj = ObjData(UserList(VictimaIndex).Invent.AlaEqpObjIndex)
        defbarco = defbarco + RandomNumber(Obj.MinDef, Obj.MaxDef)

    End If

    '[/MaTeO 9]

    Dim Resist As Byte

    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
        Resist = ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).Refuerzo

    End If

    Lugar = RandomNumber(1, 6)

    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
        If RandomNumber(1, 100) <= ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).Pegadoble Then

            Select Case Lugar

            Case bCabeza

                'Si tiene casco absorbe el golpe
                If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
                    absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    absorbido = absorbido + defbarco - Resist
                    Daño = Daño - absorbido

                    If Daño < 0 Then Daño = 1

                End If

            Case Else

                'Si tiene armadura absorbe el golpe
                If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
                    defArmadura = RandomNumber(Obj.MinDef, Obj.MaxDef)

                End If

                'Si tiene escudo absorbe el golpe
                If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
                    defEscudo = RandomNumber(Obj.MinDef, Obj.MaxDef)

                End If

                If defArmadura > 0 And defEscudo > 0 Then
                    absorbido = Int((defArmadura + defEscudo) * 0.5) + defbarco - Resist
                ElseIf defArmadura > 0 And defEscudo <= 0 Then
                    absorbido = defArmadura + defbarco - Resist
                ElseIf defArmadura <= 0 And defEscudo > 0 Then
                    absorbido = defEscudo + defbarco - Resist
                End If

                Daño = Daño - absorbido

                If Daño < 0 Then Daño = 1

            End Select

            If TeCritico = 10 Then
                ' Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "||" & vbRed & "°- " & Round(Daño * 1.1, 0) & "!" & "°" & _
                  str(UserList(VictimaIndex).char.CharIndex))
                Call SendData(SendTarget.ToIndex, AtacanteIndex, 0, "N5" & Lugar & "," & Round(Daño * 1.1, 0) & "," & UserList(VictimaIndex).Name)
                Call SendData(SendTarget.ToIndex, VictimaIndex, 0, "N4" & Lugar & "," & Daño & "," & UserList(AtacanteIndex).Name)
                UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - Round(Daño * 1.1, 0)
            Else
                'Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "||" & vbRed & "°- " & Daño & "!" & "°" & str(UserList( _
                 VictimaIndex).char.CharIndex))
                Call SendData(SendTarget.ToIndex, AtacanteIndex, 0, "N5" & Lugar & "," & Daño & "," & UserList(VictimaIndex).Name)
                Call SendData(SendTarget.ToIndex, VictimaIndex, 0, "N4" & Lugar & "," & Daño & "," & UserList(AtacanteIndex).Name)
                UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - Daño

            End If

            Select Case Lugar

            Case bCabeza

                'Si tiene casco absorbe el golpe
                If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
                    absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    absorbido = absorbido + defbarco - Resist
                    Daño = Daño - absorbido

                    If Daño < 0 Then Daño = 1

                End If

            Case Else

                'Si tiene armadura absorbe el golpe
                If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
                    defArmadura = RandomNumber(Obj.MinDef, Obj.MaxDef)

                End If

                'Si tiene escudo absorbe el golpe
                If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
                    defEscudo = RandomNumber(Obj.MinDef, Obj.MaxDef)

                End If

                If defArmadura > 0 And defEscudo > 0 Then
                    absorbido = Int((defArmadura + defEscudo) * 0.5) + defbarco - Resist
                ElseIf defArmadura > 0 And defEscudo <= 0 Then
                    absorbido = defArmadura + defbarco - Resist
                ElseIf defArmadura <= 0 And defEscudo > 0 Then
                    absorbido = defEscudo + defbarco - Resist
                End If

                Daño = Daño - absorbido

                If Daño < 0 Then Daño = 1

            End Select

            If TeCritico = 10 Then
                'Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "||" & vbRed & "°- " & Round(Daño * 1.1, 0) & "!" & "°" & _
                 str(UserList(VictimaIndex).char.CharIndex))
                Call SendData(SendTarget.ToIndex, AtacanteIndex, 0, "N5" & Lugar & "," & Round(Daño * 1.1, 0) & "," & UserList(VictimaIndex).Name)
                Call SendData(SendTarget.ToIndex, VictimaIndex, 0, "N4" & Lugar & "," & Round(Daño * 1.1, 0) & "," & UserList(AtacanteIndex).Name)
                UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - Round(Daño * 1.1, 0)
            Else
                'Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "||" & vbRed & "°- " & Daño & "!" & "°" & str(UserList( _
                 VictimaIndex).char.CharIndex))
                Call SendData(SendTarget.ToIndex, AtacanteIndex, 0, "N5" & Lugar & "," & Daño & "," & UserList(VictimaIndex).Name)
                Call SendData(SendTarget.ToIndex, VictimaIndex, 0, "N4" & Lugar & "," & Daño & "," & UserList(AtacanteIndex).Name)
                UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - Daño

            End If

        Else

            Select Case Lugar

            Case bCabeza

                'Si tiene casco absorbe el golpe
                If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
                    absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    absorbido = absorbido + defbarco - Resist
                    Daño = Daño - absorbido

                    If Daño < 0 Then Daño = 1

                End If

            Case Else

                'Si tiene armadura absorbe el golpe
                If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
                    defArmadura = RandomNumber(Obj.MinDef, Obj.MaxDef)

                End If

                'Si tiene escudo absorbe el golpe
                If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
                    defEscudo = RandomNumber(Obj.MinDef, Obj.MaxDef)

                End If

                If defArmadura > 0 And defEscudo > 0 Then
                    absorbido = Int((defArmadura + defEscudo) * 0.5) + defbarco - Resist
                ElseIf defArmadura > 0 And defEscudo <= 0 Then
                    absorbido = defArmadura + defbarco - Resist
                ElseIf defArmadura <= 0 And defEscudo > 0 Then
                    absorbido = defEscudo + defbarco - Resist

                End If

                Daño = Daño - absorbido

                If Daño < 0 Then Daño = 1

            End Select

            If TeCritico = 10 Then
                'Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "||" & vbRed & "°- " & Round(Daño * 1.1, 0) & "!" & "°" & _
                 str(UserList(VictimaIndex).char.CharIndex))
                Call SendData(SendTarget.ToIndex, AtacanteIndex, 0, "N5" & Lugar & "," & Round(Daño * 1.1, 0) & "," & UserList(VictimaIndex).Name)
                Call SendData(SendTarget.ToIndex, VictimaIndex, 0, "N4" & Lugar & "," & Round(Daño * 1.1, 0) & "," & UserList(AtacanteIndex).Name)
                UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - Daño * 1.1
            Else
                'Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "||" & vbRed & "°- " & Daño & "!" & "°" & str(UserList( _
                 VictimaIndex).char.CharIndex))
                Call SendData(SendTarget.ToIndex, AtacanteIndex, 0, "N5" & Lugar & "," & Daño & "," & UserList(VictimaIndex).Name)
                Call SendData(SendTarget.ToIndex, VictimaIndex, 0, "N4" & Lugar & "," & Daño & "," & UserList(AtacanteIndex).Name)
                UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - Daño
            End If

        End If

    Else

        Select Case Lugar

        Case bCabeza

            'Si tiene casco absorbe el golpe
            If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
                Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
                absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                absorbido = absorbido + defbarco - Resist
                Daño = Daño - absorbido

                If Daño < 0 Then Daño = 1

            End If

        Case Else

            'Si tiene armadura absorbe el golpe
            If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
                Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
                defArmadura = RandomNumber(Obj.MinDef, Obj.MaxDef)

            End If

            'Si tiene escudo absorbe el golpe
            If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
                Obj = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
                defEscudo = RandomNumber(Obj.MinDef, Obj.MaxDef)

            End If

            If defArmadura > 0 And defEscudo > 0 Then
                absorbido = Int((defArmadura + defEscudo) * 0.5) + defbarco - Resist
            ElseIf defArmadura > 0 And defEscudo <= 0 Then
                absorbido = defArmadura + defbarco - Resist
            ElseIf defArmadura <= 0 And defEscudo > 0 Then
                absorbido = defEscudo + defbarco - Resist

            End If

            Daño = Daño - absorbido

            If Daño < 0 Then Daño = 1

        End Select

        'Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "||" & vbRed & "°- " & Daño & "!" & "°" & str(UserList( _
         VictimaIndex).char.CharIndex))
        Call SendData(SendTarget.ToIndex, AtacanteIndex, 0, "N5" & Lugar & "," & Daño & "," & UserList(VictimaIndex).Name)
        Call SendData(SendTarget.ToIndex, VictimaIndex, 0, "N4" & Lugar & "," & Daño & "," & UserList(AtacanteIndex).Name)
        UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - Daño

    End If

    If UserList(AtacanteIndex).flags.Hambre = 0 And UserList(AtacanteIndex).flags.Sed = 0 Then

        'Si usa un arma quizas suba "Combate con armas"
        If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
            Call SubirSkill(AtacanteIndex, Armas)
        Else
            'sino tal vez lucha libre
            Call SubirSkill(AtacanteIndex, Wresterling)

        End If

        Call SubirSkill(AtacanteIndex, Tacticas)

        'Trata de apuñalar por la espalda al enemigo
        If PuedeApuñalar(AtacanteIndex) Then
            Call DoApuñalar(AtacanteIndex, 0, VictimaIndex, Daño)
            Call SubirSkill(AtacanteIndex, Apuñalar)

        End If

    End If

    If UserList(VictimaIndex).Stats.MinHP <= 0 Then

        Call ContarMuerte(VictimaIndex, AtacanteIndex)

        ' Para que las mascotas no sigan intentando luchar y
        ' comiencen a seguir al amo
        Dim j As Integer

        For j = 1 To MAXMASCOTAS

            If UserList(AtacanteIndex).MascotasIndex(j) > 0 Then
                If Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = VictimaIndex Then Npclist(UserList(AtacanteIndex).MascotasIndex( _
                                                                                                        j)).Target = 0
                Call FollowAmo(UserList(AtacanteIndex).MascotasIndex(j))

            End If

        Next j

        Call ActStats(VictimaIndex, AtacanteIndex)
    Else

        If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
            Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
            If UserList(AtacanteIndex).Invent.MunicionEqpObjIndex > 0 Then
                Municion = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
                If ObjData(Municion).Paraliza > 0 And UserList(VictimaIndex).flags.Paralizado = 0 Then
                    If RandomNumber(1, 100) <= ObjData(Municion).Paraliza Then
                        UserList(VictimaIndex).flags.Paralizado = 1
                        UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado
                        Call SendData(SendTarget.ToIndex, VictimaIndex, 0, "PARADOW")
                    End If
                End If
            Else
                If ObjData(Arma).Paraliza > 0 And UserList(VictimaIndex).flags.Paralizado = 0 Then
                    If RandomNumber(1, 100) <= ObjData(Arma).Paraliza Then
                        UserList(VictimaIndex).flags.Paralizado = 1
                        UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado
                        Call SendData(SendTarget.ToIndex, VictimaIndex, 0, "PARADOW")
                    End If
                End If
            End If
        End If


        'Está vivo - Actualizamos el HP
        Call SendData(SendTarget.ToIndex, VictimaIndex, 0, "ASH" & UserList(VictimaIndex).Stats.MinHP)
    End If

    'Controla el nivel del usuario
    Call CheckUserLevel(AtacanteIndex)
ErrorHandler:

    '   Call LogError("Error en SUB USERDAÑOUSER. Error: " & Err.Number & " Descripción: " & Err.Description)
End Sub

Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)

    If TriggerZonaPelea(AttackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub

    If Not Criminal(AttackerIndex) And Not Criminal(VictimIndex) Then
        Call VolverCriminal(AttackerIndex)

    End If

    If UserList(VictimIndex).pos.Map = "154" Or UserList(AttackerIndex).Faccion.ArmadaReal = "1" Or UserList(AttackerIndex).Faccion.Templario = "1" Then
        Exit Sub
    Else
        If Not Criminal(VictimIndex) Then
            UserList(AttackerIndex).Reputacion.BandidoRep = UserList(AttackerIndex).Reputacion.BandidoRep + vlASALTO

            If UserList(AttackerIndex).Reputacion.BandidoRep > MAXREP Then UserList(AttackerIndex).Reputacion.BandidoRep = MAXREP
        Else
            UserList(AttackerIndex).Reputacion.NobleRep = UserList(AttackerIndex).Reputacion.NobleRep + vlNoble

            If UserList(AttackerIndex).Reputacion.NobleRep > MAXREP Then UserList(AttackerIndex).Reputacion.NobleRep = MAXREP

        End If
    End If

    Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)

    If UserList(AttackerIndex).flags.EstaDueleando = True And UserList(VictimIndex).flags.EstaDueleando = True Then
        Exit Sub

    End If

    If UserList(AttackerIndex).flags.EstaDueleando1 = True And UserList(VictimIndex).flags.EstaDueleando1 = True Then
        Exit Sub

    End If

End Sub

Sub AllMascotasAtacanUser(ByVal Victim As Integer, ByVal Maestro As Integer)
'Reaccion de las mascotas
    Dim iCount As Integer

    For iCount = 1 To MAXMASCOTAS

        If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(Victim).Name
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1

        End If

    Next iCount

End Sub

Public Function ZonaDuelos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

    If Map = 35 Then
        If X >= 40 And X <= 61 And Y >= 76 And Y <= 89 Then
            ZonaDuelos = True
            Exit Function
        End If
    End If

    If Map = 150 Then
        If X >= 21 And X <= 46 And Y >= 21 And Y <= 38 Then
            ZonaDuelos = True
            Exit Function
        End If

        If X >= 16 And X <= 47 And Y >= 57 And Y <= 79 Then
            ZonaDuelos = True
            Exit Function
        End If

        If X >= 62 And X <= 87 And Y >= 20 And Y <= 38 Then
            ZonaDuelos = True
            Exit Function
        End If

        If X >= 59 And X <= 90 And Y >= 57 And Y <= 84 Then
            ZonaDuelos = True
            Exit Function
        End If
    End If

    If Map = 154 Then
        If X >= 38 And X <= 63 And Y >= 41 And Y <= 60 Then
            ZonaDuelos = True
            Exit Function
        End If
    End If

    If Map = 160 Then
        If X >= 35 And X <= 67 And Y >= 40 And Y <= 62 Then
            ZonaDuelos = True
            Exit Function
        End If
    End If

    If Map = 161 Then
        If X >= 34 And X <= 67 And Y >= 37 And Y <= 63 Then
            ZonaDuelos = True
            Exit Function
        End If
    End If

    If Map = 190 Then
        If X >= 16 And X <= 47 And Y >= 57 And Y <= 79 Then
            ZonaDuelos = True
            Exit Function
        End If
    End If

    ZonaDuelos = False

End Function

Public Function ProhibidoAtacar(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

    If Map = "35" Then

        If X >= "40" And X <= "61" And Y >= 76 And Y <= 78 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= "40" And X <= "41" And Y >= 79 And Y <= 81 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X = "40" And Y >= 82 And Y <= 84 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= "40" And X <= "41" And Y >= 85 And Y <= 87 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= "40" And X <= "61" And Y >= 88 And Y <= 89 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= "60" And X <= "61" And Y >= 79 And Y <= 87 Then
            ProhibidoAtacar = True
            Exit Function
        End If

    End If

    If Map = 150 Then

        If X >= 21 And X <= 46 And Y >= 21 And Y <= 24 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 21 And X <= 46 And Y >= 36 And Y <= 38 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 21 And X <= 24 And Y >= 25 And Y <= 28 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 21 And X <= 24 And Y >= 32 And Y <= 35 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 21 And X <= 22 And Y >= 29 And Y <= 31 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 43 And X <= 46 And Y >= 25 And Y <= 28 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 43 And X <= 46 And Y >= 32 And Y <= 35 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 45 And X <= 45 And Y >= 29 And Y <= 31 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 16 And X <= 47 And Y >= 57 And Y <= 60 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 16 And X <= 18 And Y >= 61 And Y <= 75 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X = 19 And Y >= 64 And Y <= 72 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 16 And X <= 47 And Y >= 76 And Y <= 79 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 45 And X <= 47 And Y >= 61 And Y <= 75 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X = 44 And Y >= 64 And Y <= 72 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 62 And X <= 87 And Y >= 20 And Y <= 24 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 62 And X <= 65 And Y >= 25 And Y <= 28 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 62 And X <= 65 And Y >= 32 And Y <= 35 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 62 And X <= 63 And Y >= 29 And Y <= 31 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 62 And X <= 87 And Y >= 36 And Y <= 38 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 84 And X <= 87 And Y >= 25 And Y <= 28 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 84 And X <= 87 And Y >= 32 And Y <= 35 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 86 And X <= 87 And Y >= 29 And Y <= 31 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 59 And X <= 90 And Y >= 57 And Y <= 62 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 59 And X <= 90 And Y >= 78 And Y <= 84 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 59 And X <= 61 And Y >= 63 And Y <= 77 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X = 62 And Y >= 66 And Y <= 74 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 89 And X <= 90 And Y >= 63 And Y <= 77 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X = 88 And Y >= 66 And Y <= 74 Then
            ProhibidoAtacar = True
            Exit Function
        End If

    End If

    If Map = 160 Then

        If X >= 38 And X <= 64 And Y >= 43 And Y <= 44 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 38 And X <= 64 And Y >= 58 And Y <= 59 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 38 And X <= 41 And Y >= 45 And Y <= 49 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 38 And X <= 41 And Y >= 53 And Y <= 57 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 38 And X <= 40 And Y >= 50 And Y <= 52 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 60 And X <= 64 And Y >= 45 And Y <= 49 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 60 And X <= 64 And Y >= 53 And Y <= 57 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 61 And X <= 64 And Y >= 50 And Y <= 52 Then
            ProhibidoAtacar = True
            Exit Function
        End If
    End If

    If Map = 161 Then

        If X >= 34 And X <= 67 And Y >= 37 And Y <= 44 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 34 And X <= 67 And Y >= 57 And Y <= 63 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 34 And X <= 40 And Y >= 45 And Y <= 56 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 61 And X <= 67 And Y >= 45 And Y <= 56 Then
            ProhibidoAtacar = True
            Exit Function
        End If

    End If

    If Map = 190 Then

        If X >= 16 And X <= 47 And Y >= 57 And Y <= 60 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 16 And X <= 18 And Y >= 61 And Y <= 75 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X = 19 And Y >= 64 And Y <= 72 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 16 And X <= 47 And Y >= 76 And Y <= 79 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X >= 45 And X <= 47 And Y >= 61 And Y <= 75 Then
            ProhibidoAtacar = True
            Exit Function
        End If

        If X = 44 And Y >= 64 And Y <= 72 Then
            ProhibidoAtacar = True
            Exit Function
        End If
    End If

    ProhibidoAtacar = False
End Function

Public Function PuedeAtacar(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean

    On Error GoTo errhandler

    Dim T As eTrigger6

    'Estas en modo Combate?
    If Not UserList(AttackerIndex).flags.SeguroCombate Then
        Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate." & _
                                                            FONTTYPE_Motd4)
        PuedeAtacar = False
        Exit Function
    End If

    If UserList(VictimIndex).flags.Muerto = 1 Then
        SendData SendTarget.ToIndex, AttackerIndex, 0, "||No puedes atacar a un espiritu" & FONTTYPE_INFO
        PuedeAtacar = False
        Exit Function
    End If

    If UserList(VictimIndex).Asedio.Team = UserList(AttackerIndex).Asedio.Team And UserList(AttackerIndex).Asedio.Team <> 0 Then
        SendData SendTarget.ToIndex, AttackerIndex, 0, "||¡No puedes atacar a los miembros de tu equipo!" & FONTTYPE_WARNING
        Exit Function
    End If

    If ProhibidoAtacar(UserList(AttackerIndex).pos.Map, UserList(AttackerIndex).pos.X, UserList(AttackerIndex).pos.Y) Then
        Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||Atacar en esta zona está prohibido." & FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    ElseIf ProhibidoAtacar(UserList(VictimIndex).pos.Map, UserList(VictimIndex).pos.X, UserList(VictimIndex).pos.Y) Then
        Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||Atacar en esta zona está prohibido." & FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If


    'A la pareja no?
    ' If UserList(VictimIndex).Name = UserList(AttackerIndex).Pareja Then
    '     Call SendData(SendTarget.toIndex, AttackerIndex, 0, "||No puedes atacar a tu pareja." & FONTTYPE_WARNING)
    '     PuedeAtacar = False
    '     Exit Function
    ' End If

    If UserList(AttackerIndex).flags.Seguro Then
        If Not Criminal(VictimIndex) Then
            Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||No puedes atacar ciudadanos con el seguro activado." & FONTTYPE_WARNING)
            Exit Function
        End If
    End If

    'Se asegura que la victima no es un GM
    If UserList(VictimIndex).flags.Privilegios >= PlayerType.Consejero Then
        SendData SendTarget.ToIndex, AttackerIndex, 0, "||No puedes atacar a los dioses!" & FONTTYPE_WARNING
        PuedeAtacar = False
        Exit Function
    End If

    If UserList(AttackerIndex).flags.SeguroClan = True Then
        If Guilds(UserList(AttackerIndex).GuildIndex).GuildName <> "" Then
            If Guilds(UserList(VictimIndex).GuildIndex).GuildName = Guilds(UserList(AttackerIndex).GuildIndex).GuildName Then
                Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||Para atacar a tu propio clan presiona la tecla W." & FONTTYPE_INFO)
                PuedeAtacar = False
                Exit Function
            End If
        End If
    End If

    If UserList(AttackerIndex).PartyIndex > 0 Then
        If UserList(AttackerIndex).PartyIndex = UserList(VictimIndex).PartyIndex Then
            Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||No puedes atacar a tu compañero de party." & FONTTYPE_INFO)
            PuedeAtacar = False
            Exit Function
        End If
    End If

    T = TriggerZonaPelea(AttackerIndex, VictimIndex)

    If T = TRIGGER6_PERMITE Then
        PuedeAtacar = True
        Exit Function
    ElseIf T = TRIGGER6_PROHIBE Then
        PuedeAtacar = False
        Exit Function

    End If

    If MapInfo(UserList(VictimIndex).pos.Map).Pk = False And Not MapData(UserList(AttackerIndex).pos.Map, UserList(AttackerIndex).pos.X, UserList(AttackerIndex).pos.Y).Trigger = 2 Then

        If esTemplario(AttackerIndex) Then
            Select Case UserList(VictimIndex).pos.Map
            Case "95"
                PuedeAtacar = True
                Exit Function
            End Select
        End If

        If esNemesis(AttackerIndex) Then
            Select Case UserList(VictimIndex).pos.Map
            Case "84"
                PuedeAtacar = True
                Exit Function

            Case "20"
                PuedeAtacar = True
                Exit Function
            End Select
        End If

        If esArmada(AttackerIndex) Then
            Select Case UserList(VictimIndex).pos.Map
            Case "58"
                PuedeAtacar = True
                Exit Function
            Case "59"
                PuedeAtacar = True
                Exit Function
            Case "60"
                PuedeAtacar = True
                Exit Function
            Case "61"
                PuedeAtacar = True
                Exit Function

            End Select
        End If

        If esCaos(AttackerIndex) Then
            Select Case UserList(VictimIndex).pos.Map
            Case "132"
                PuedeAtacar = True
                Exit Function
            End Select
        End If

        If esTemplario(AttackerIndex) Or esNemesis(AttackerIndex) Then
            Select Case UserList(VictimIndex).pos.Map
            Case "86"
                PuedeAtacar = True
                Exit Function
            End Select
        End If

        If Not Criminal(AttackerIndex) And Criminal(VictimIndex) Then
            Select Case UserList(VictimIndex).pos.Map
            Case "1"
                PuedeAtacar = True
                Exit Function
            End Select
        End If

        Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||En zona segura no se pueden atacar otros usuarios." & FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If

    If MapData(UserList(VictimIndex).pos.Map, UserList(VictimIndex).pos.X, UserList(VictimIndex).pos.Y).Trigger = eTrigger.ZONASEGURA Or MapData( _
       UserList(AttackerIndex).pos.Map, UserList(AttackerIndex).pos.X, UserList(AttackerIndex).pos.Y).Trigger = eTrigger.ZONASEGURA Then
        Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||No puedes pelear aqui." & FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If

    'If (Not Criminal(VictimIndex)) And (UserList(AttackerIndex).Faccion.ArmadaReal = 1) Then
    '    Call SendData(SendTarget.toIndex, AttackerIndex, 0, "||Los soldados del Ejercito Real tienen prohibido atacar ciudadanos." & FONTTYPE_WARNING)
    '    PuedeAtacar = False
    '    Exit Function
    'End If

    'If (Not Criminal(VictimIndex)) And (UserList(AttackerIndex).Faccion.Templario = 1) Then
    '    Call SendData(SendTarget.toIndex, AttackerIndex, 0, "||Los soldados del Ejercito Templario tienen prohibido atacar ciudadanos." & _
         '            FONTTYPE_WARNING)
    '    PuedeAtacar = False
    '    Exit Function
    'End If

    If UserList(VictimIndex).flags.Demonio = True And UserList(AttackerIndex).flags.Demonio = True Then
        Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||No puedes atacar a tu bando!." & FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function

    End If

    If UserList(VictimIndex).flags.Angel = True And UserList(AttackerIndex).flags.Angel = True Then
        Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||No puedes atacar a tu bando!." & FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function

    End If


    If UserList(AttackerIndex).flags.Privilegios = PlayerType.Consejero Then
        PuedeAtacar = False
        Exit Function
    End If

    If UserList(VictimIndex).flags.EstaDueleando = True And UserList(AttackerIndex).flags.EstaDueleando = True Then
        PuedeAtacar = True
        Exit Function
    End If

    If UserList(VictimIndex).flags.EstaDueleando1 = True And UserList(AttackerIndex).flags.EstaDueleando1 = True Then
        PuedeAtacar = True
        Exit Function
    End If

    If UserList(AttackerIndex).flags.Muerto = 1 Then
        SendData SendTarget.ToIndex, AttackerIndex, 0, "||No puedes atacar porque estas muerto" & FONTTYPE_INFO
        PuedeAtacar = False
        Exit Function
    End If

    PuedeAtacar = True
errhandler:
    PuedeAtacar = True

End Function

Public Function PuedeAtacarNPC(ByVal AttackerIndex As Integer, ByVal NpcIndex As Integer) As Boolean

'Estas en modo Combate?
    If Not UserList(AttackerIndex).flags.SeguroCombate Then
        Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate." & _
                                                            FONTTYPE_Motd4)
        PuedeAtacarNPC = False
        Exit Function
    End If

    If Npclist(NpcIndex).MaestroUser > 0 Then
        If Not Criminal(AttackerIndex) And Not Criminal(Npclist(NpcIndex).MaestroUser) Then
            If UserList(AttackerIndex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||Para atacar mascotas de ciudadanos debes quitarte el seguro" & FONTTYPE_FIGHT)
                PuedeAtacarNPC = False
                Exit Function
            End If
        End If
    End If

    If NpcIndex = ReyIndex And ReyTeam = UserList(AttackerIndex).Asedio.Team Then
        Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||¡No puedes atacar a tu propio Rey!" & FONTTYPE_FIGHT)
        PuedeAtacarNPC = False
        Exit Function
    End If

    If UserList(AttackerIndex).flags.Muerto = 1 Then
        SendData SendTarget.ToIndex, AttackerIndex, 0, "Z12"
        PuedeAtacarNPC = False
        Exit Function
    End If

    If UserList(AttackerIndex).flags.Privilegios = PlayerType.Consejero Then
        PuedeAtacarNPC = False
        Exit Function
    End If

    PuedeAtacarNPC = True

End Function

'[KEVIN]
'
'[Alejo]
'Modifique un poco el sistema de exp por golpe, ahora
'son 2/3 de la exp mientras esta vivo, el resto se
'obtiene al matarlo.
'Ahora además
Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)

    With Npclist(NpcIndex)
        Dim ExpaDar As Long
        Dim DivExp As Long
        Dim MulExp As Long
        Dim Div2Exp As Long
        Dim CantMiembro As Byte


        If .Stats.MaxHP <= 0 Then Exit Sub
        If ElDaño <= 0 Then ElDaño = 0
        If Npclist(NpcIndex).MurallaEquipo > 0 Then
            Dim i As Long
            Dim IndexM As Integer
            For i = 0 To 6
                If i <> Npclist(NpcIndex).MurallaIndex Then
                    IndexM = Muralla(i, Npclist(NpcIndex).MurallaEquipo)
                    Npclist(IndexM).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP
                    If Npclist(IndexM).Stats.MinHP <= 0 Then Call MuereNpc(IndexM, UserIndex)
                End If
            Next i
            Call modAsedio.CalcularGrafico(NpcIndex)
        End If
        ' If ElDaño > .Stats.MinHP Then ElDaño = .Stats.MinHP

        '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia

        'ExpaDar = CLng(ElDaño * (.GiveEXP / .Stats.MaxHP))

        If .Stats.MinHP > 0 Then
            ExpaDar = CLng((((val(.GiveEXP) / val(.Stats.MaxHP)) * val(.Stats.MinHP)) / val(.Stats.MinHP)) * ElDaño)
        Else
            ExpaDar = CLng((((val(.GiveEXP) / val(.Stats.MaxHP)) * .Stats.MaxHP) / .Stats.MaxHP) * ElDaño)
        End If

        If UserIndex > 0 Then
            If UserList(UserIndex).GranPoder = 1 And GranPoder.TipoAura = hGranPoder.Experencia Then
                ExpaDar = CLng(ExpaDar * ("1," & GranPoder.Cantidad))
            End If
        End If

        'If ExpaDar <= 0 Then Exit Sub
        '
        '      If ExpaDar > .flags.ExpCount Then
        '         ' ExpaDar = .flags.ExpCount
        '         .flags.ExpCount = 0
        '       Else
        '           .flags.ExpCount = .flags.ExpCount - ExpaDar
        ''
        '     End If

        If ExpaDar <> 0 Then

            If UserList(UserIndex).PartyIndex > 0 Then
                Call mdParty.ObtenerExito(UserIndex, ExpaDar, .pos.Map, .pos.X, .pos.Y)
            Else
                UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar

                If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP

                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has ganado " & ExpaDar & " puntos de experiencia!" & FONTTYPE_Motd4)

            End If

            Call CheckUserLevel(UserIndex)
            Call EnviarExp(UserIndex)

        End If

    End With

End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6

    If Origen > 0 And Destino > 0 And Origen <= UBound(UserList) And Destino <= UBound(UserList) Then
        If MapData(UserList(Origen).pos.Map, UserList(Origen).pos.X, UserList(Origen).pos.Y).Trigger = eTrigger.ZONAPELEA Or MapData(UserList( _
                                                                                                                                     Destino).pos.Map, UserList(Destino).pos.X, UserList(Destino).pos.Y).Trigger = eTrigger.ZONAPELEA Then

            If (MapData(UserList(Origen).pos.Map, UserList(Origen).pos.X, UserList(Origen).pos.Y).Trigger = MapData(UserList(Destino).pos.Map, _
                                                                                                                    UserList(Destino).pos.X, UserList(Destino).pos.Y).Trigger) Then
                TriggerZonaPelea = TRIGGER6_PERMITE
            Else
                TriggerZonaPelea = TRIGGER6_PROHIBE

            End If

        Else
            TriggerZonaPelea = TRIGGER6_AUSENTE

        End If

    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE

    End If

End Function

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    Dim ArmaObjInd As Integer, ObjInd As Integer
    Dim num As Long

    ArmaObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    ObjInd = 0

    If ArmaObjInd > 0 Then
        If ObjData(ArmaObjInd).proyectil = 0 Then
            ObjInd = ArmaObjInd
        Else
            ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex

        End If

        If ObjInd > 0 Then
            If (ObjData(ObjInd).Envenena = 1) Then
                num = RandomNumber(1, 100)

                If num < 60 Then
                    UserList(VictimaIndex).flags.Envenenado = 1
                    Call SendData(SendTarget.ToIndex, VictimaIndex, 0, "||" & UserList(AtacanteIndex).Name & " te ha envenenado!!" & FONTTYPE_Motd4)
                    Call SendData(SendTarget.ToIndex, AtacanteIndex, 0, "||Has envenenado a " & UserList(VictimaIndex).Name & "!!" & FONTTYPE_Motd4)

                End If

            End If

        End If

    End If

End Sub
