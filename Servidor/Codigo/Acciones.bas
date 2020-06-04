Attribute VB_Name = "Acciones"

'Pablo Ignacio Márquez

Option Explicit

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''eva
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Ys

Sub Accion(ByVal UserIndex As Integer, _
           ByVal Map As Integer, _
           ByVal X As Integer, _
           ByVal Y As Integer)

    On Error Resume Next

    '¿Posicion valida?
    If InMapBounds(Map, X, Y) Then

        Dim FoundChar      As Byte

        Dim FoundSomething As Byte

        Dim TempCharIndex  As Integer

        'USUARIO
        '        If MapData(Map, X, Y).UserIndex > 0 And UserList(UserIndex).flags.Privilegios <> PlayerType.User Then
        '
        '
        '                UserList(UserIndex).flags.SeleccioneA = UserList(MapData(Map, X, Y).UserIndex).Name
        '                UserList(UserList(UserIndex).flags.TargetUser).flags.EstoySelec = 1
        '
        '                Call SendData(SendTarget.ToIndex, UserList(UserIndex).flags.TargetUser, 0, "||El GM te a seleccionado para teletransportarte." & _
        '                        FONTTYPE_INFO)
        '
        '                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Seleccionaste a: " & UserList(UserIndex).flags.SeleccioneA & FONTTYPE_INFO)
        '                Call SendData(SendTarget.ToIndex, UserIndex, 0, "TX")
        '
        '        End If

        'USUARIO

        '¿Es un obj?
        If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.ObjIndex

            Select Case ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ObjType

                Case eOBJType.otPuertas    'Es una puerta
                    Call AccionParaPuerta(Map, X, Y, UserIndex)

                Case eOBJType.otCARTELES    'Es un cartel
                    Call AccionParaCartel(Map, X, Y, UserIndex)

                Case eOBJType.otFOROS    'Foro
                    Call AccionParaForo(Map, X, Y, UserIndex)
                
                Case eOBJType.otTumba
                     Call AccionParaTumba(Map, X, Y, UserIndex)

                Case eOBJType.otLeña    'Leña

                    If MapData(Map, X, Y).OBJInfo.ObjIndex = FOGATA_APAG And UserList(UserIndex).flags.Muerto = 0 Then
                        Call AccionParaRamita(Map, X, Y, UserIndex)

                    End If

            End Select

            '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
        ElseIf MapData(Map, X + 1, Y).OBJInfo.ObjIndex > 0 And MapData(Map, X, Y).NpcIndex = 0 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y).OBJInfo.ObjIndex
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).ObjType & "," & ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).Name & "," & "OBJ")

            Select Case ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).ObjType

                Case 6    'Es una puerta
                    Call AccionParaPuerta(Map, X + 1, Y, UserIndex)

            End Select

        ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex > 0 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).ObjType & "," & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).Name & "," & "OBJ")

            Select Case ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).ObjType

                Case 6    'Es una puerta
                    Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)

            End Select

        ElseIf MapData(Map, X, Y + 1).OBJInfo.ObjIndex > 0 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y + 1).OBJInfo.ObjIndex
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).ObjType & "," & ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).Name & "," & "OBJ")

            Select Case ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).ObjType

                Case 6    'Es una puerta
                    Call AccionParaPuerta(Map, X, Y + 1, UserIndex)

            End Select

        ElseIf MapData(Map, X, Y).UserIndex > 0 Then

            If UserList(UserIndex).flags.Privilegios > User Then Call SendData(SendTarget.ToIndex, UserIndex, 0, "TEST" & UserList(UserList(UserIndex).flags.TargetUser).Name & ", " & UserList(UserList(UserIndex).flags.TargetUser).Stats.UsuariosMatados & ", " & UserList(UserList(UserIndex).flags.TargetUser).Clase)

        ElseIf MapData(Map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
            'Set the target NPC
            UserList(UserIndex).flags.TargetNpc = MapData(Map, X, Y).NpcIndex

            If Npclist(MapData(Map, X, Y).NpcIndex).Comercia = 1 Then

                If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Creditos Then

                    If Distancia(Npclist(MapData(Map, X, Y).NpcIndex).Pos, UserList(UserIndex).Pos) > 4 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                        Exit Sub

                    End If

                    If UserList(UserIndex).flags.Montado = True Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes comerciar estando arriba de tu Mascota!" & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    'A depositar de una
                    Call Mod_Monedas.IniciarComercioCreditos(UserIndex)
                    Exit Sub

                ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Canjes Then

                    If Distancia(Npclist(MapData(Map, X, Y).NpcIndex).Pos, UserList(UserIndex).Pos) > 4 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                        Exit Sub

                    End If

                    If UserList(UserIndex).flags.Montado = True Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes comerciar estando arriba de tu Mascota!" & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    'A depositar de una
                    Call Mod_Monedas.IniciarComercioCanjes(UserIndex)
                    Exit Sub

                End If

                If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Montado = True Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes comerciar estando arriba de tu Mascota!" & FONTTYPE_INFO)
                    Exit Sub

                End If

                'Iniciamos la rutina pa' comerciar.
                Call IniciarCOmercioNPC(UserIndex)

            ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Banquero Then

                If Distancia(Npclist(MapData(Map, X, Y).NpcIndex).Pos, UserList(UserIndex).Pos) > 4 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                    Exit Sub

                End If

                If UserList(UserIndex).flags.ValidBank = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "VLDB")
                    Exit Sub
                ElseIf UserList(UserIndex).flags.ValidBank = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "BANP" & UserList(UserIndex).Stats.Banco & "," & UserList(UserIndex).Stats.GLD & "," & UserList(UserIndex).BancoInvent.NroItems)
                    Exit Sub

                End If

            ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Revividor Then

                If Distancia(UserList(UserIndex).Pos, Npclist(MapData(Map, X, Y).NpcIndex).Pos) > 10 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z32")
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Envenenado = 1 Then
                    UserList(UserIndex).flags.Envenenado = 0
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te has curado del envenenamiento." & FONTTYPE_INFO)

                End If

                'Revivimos si es necesario
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call RevivirUsuario(UserIndex)

                End If

                'curamos totalmente
                UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP

                Call EnviarHP(UserIndex)

            ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.OlvidarHechizo Then

                If Distancia(UserList(UserIndex).Pos, Npclist(MapData(Map, X, Y).NpcIndex).Pos) > 5 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                    Exit Sub

                End If

                If OroHechizo > UserList(UserIndex).Stats.GLD Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes oro suficiente." & FONTTYPE_TALKMSG)
                    Exit Sub

                End If

                Call SendData(SendTarget.ToIndex, UserIndex, 0, "HECA" & FONTTYPE_INFO)

            ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Cirujia Then

                If Distancia(UserList(UserIndex).Pos, Npclist(MapData(Map, X, Y).NpcIndex).Pos) > 5 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                    Exit Sub

                End If

                If OroCirujia > UserList(UserIndex).Stats.GLD Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes suficientes monedas de oro para la cirujía." & FONTTYPE_INFO)
                    Exit Sub

                End If

                Call IniciarChangeHead(UserIndex)

            End If
            
            If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Misiones Then
                Call ClickMisionesQuest(UserIndex)
            End If
            
            If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.nQuest Then
                Call IniciarVentanaQuest(UserIndex)
           End If
           
           If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Subasta Then
              
              If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estás muerto!!" & FONTTYPE_INFO)
                  Exit Sub
              End If
               
                Call IniciarVentanaSubasta(UserIndex)
           End If

        Else
            
            UserList(UserIndex).flags.TargetNpc = 0
            UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0

        End If

    End If

End Sub

Sub AccionParaForo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim Pos As WorldPos
    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y

    If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
        Exit Sub

    End If

    '¿Hay mensajes?
    Dim f As String, tit As String, men As String, base As String, auxcad As String
    f = App.Path & "\foros\" & UCase$(ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ForoID) & ".for"

    If FileExist(f, vbNormal) Then
        Dim num As Integer
        num = val(GetVar(f, "INFO", "CantMSG"))
        base = Left$(f, Len(f) - 4)
        Dim i As Integer
        Dim n As Integer

        For i = 1 To num
            n = FreeFile
            f = base & i & ".for"
            Open f For Input Shared As #n
            Input #n, tit
            men = ""
            auxcad = ""

            Do While Not EOF(n)
                Input #n, auxcad
                men = men & vbCrLf & auxcad
            Loop
            Close #n
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "FMSG" & tit & Chr(176) & men)

        Next

    End If

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "MFOR")

End Sub

Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim MiObj As Obj
    Dim wp As WorldPos

    If Not (Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2) Then
        If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Llave = 0 Then
            If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Cerrada = 1 Then

                'Abre la puerta
                If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Llave = 0 Then

                    MapData(Map, X, Y).OBJInfo.ObjIndex = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).IndexAbierta

                    Call ModAreas.SendToAreaByPos(Map, X, Y, "HO" & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).GrhIndex & "," & X & "," & Y)

                    'Desbloquea
                    MapData(Map, X, Y).Blocked = 0
                    MapData(Map, X - 1, Y).Blocked = 0

                    'Bloquea todos los mapas
                    Call Bloquear(SendTarget.ToMap, 0, Map, Map, X, Y, 0)
                    Call Bloquear(SendTarget.ToMap, 0, Map, Map, X - 1, Y, 0)

                    'Sonido
                    SendData SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_PUERTA

                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La puerta esta cerrada con llave." & FONTTYPE_INFO)

                End If

            Else
                'Cierra puerta
                MapData(Map, X, Y).OBJInfo.ObjIndex = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).IndexCerrada

                Call ModAreas.SendToAreaByPos(Map, X, Y, "HO" & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).GrhIndex & "," & X & "," & Y)

                MapData(Map, X, Y).Blocked = 1
                MapData(Map, X - 1, Y).Blocked = 1

                Call Bloquear(SendTarget.ToMap, 0, Map, Map, X - 1, Y, 1)
                Call Bloquear(SendTarget.ToMap, 0, Map, Map, X, Y, 1)

                SendData SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_PUERTA

            End If

            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.ObjIndex
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La puerta esta cerrada con llave." & FONTTYPE_INFO)

        End If

    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")

    End If

End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim MiObj As Obj

    If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ObjType = 8 Then

        If Len(ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).texto) > 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MCAR" & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).texto & Chr(176) & ObjData( _
                                                            MapData(Map, X, Y).OBJInfo.ObjIndex).GrhSecundario)

        End If

    End If

End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim Suerte As Byte
    Dim exito As Byte
    Dim Obj As Obj
    Dim raise As Integer

    Dim Pos As WorldPos
    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y

    If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
        Exit Sub

    End If

    If MapData(Map, X, Y).Trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Pk = False Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||En zona segura no puedes hacer fogatas." & FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).Stats.UserSkills(Supervivencia) > 1 And UserList(UserIndex).Stats.UserSkills(Supervivencia) < 6 Then
        Suerte = 3
    ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 10 Then
        Suerte = 2
    ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 10 And UserList(UserIndex).Stats.UserSkills(Supervivencia) Then
        Suerte = 1

    End If

    exito = RandomNumber(1, Suerte)

    If exito = 1 Then
        If MapInfo(UserList(UserIndex).Pos.Map).Zona <> Ciudad Then
            Obj.ObjIndex = FOGATA
            Obj.Amount = 1

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has prendido la fogata." & FONTTYPE_INFO)
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "FO")

            Call MakeObj(ToMap, 0, Map, Obj, Map, X, Y)

            'Las fogatas prendidas se deben eliminar
            Dim Fogatita As New cGarbage
            Fogatita.Map = Map
            Fogatita.X = X
            Fogatita.Y = Y
            Call TrashCollector.Add(Fogatita)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La ley impide realizar fogatas en las ciudades." & FONTTYPE_INFO)
            Exit Sub

        End If

    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No has podido hacer fuego." & FONTTYPE_INFO)

    End If

    'Sino tiene hambre o sed quizas suba el skill supervivencia
    If UserList(UserIndex).flags.Hambre = 0 And UserList(UserIndex).flags.Sed = 0 Then
        Call SubirSkill(UserIndex, Supervivencia)

    End If

End Sub

Sub AccionParaTumba(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ObjType = eOBJType.otTumba Then
    
    If MapData(Map, X, Y).OBJInfo.ObjIndex = 1773 Then
        Call AccionParaTumbaTotan(Map, X, Y, UserIndex)
        Exit Sub
    End If
    
    If MapData(Map, X, Y).OBJInfo.ObjIndex = 1774 Then
        Call AccionParaTumbaSaturos(Map, X, Y, UserIndex)
        Exit Sub
    End If
    
End If

End Sub

Sub AccionParaTumbaTotan(ByVal Map As Integer, _
                         ByVal X As Integer, _
                         ByVal Y As Integer, _
                         ByVal UserIndex As Integer)

    On Error Resume Next

    Dim texto As String

    Dim Name  As String

    Name = ObjData(objtotan(objetoTotan)).Name

    If quitarObjetoTumba(UserIndex, objtotan(objetoTotan)) Then

        Dim MiObj As Obj

        'titraroro
        MiObj.ObjIndex = iORO

        Dim n As Integer

        For n = 1 To objetoTotan
            MiObj.Amount = 10000
            Call TirarItemAlPiso(posTumbaTotan, MiObj)
        Next n
    
        Call EraseObj(ToMap, 0, posTumbaTotan.Map, 1, posTumbaTotan.Map, posTumbaTotan.X, posTumbaTotan.Y)
        Call EraseObj(ToMap, 0, posTumbaSatu.Map, 1, posTumbaSatu.Map, posTumbaSatu.X, posTumbaSatu.Y)
        posTumbaTotan.Map = 0
        posTumbaSatu.Map = 0
        TumbaOn = False
    Else
        texto = "HOLA VIAJERO SOY TOTAN EL DE LAS TORTUGAS, TRAEME 1 " & Name & " Y TE DARE UNA RECOMPENSA..."

    End If

    If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ObjType = eOBJType.otTumba Then
  
        Call SendData(ToIndex, UserIndex, 0, "MCAR" & texto & Chr(176) & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).GrhSecundario)

    End If

End Sub

Sub AccionParaTumbaSaturos(ByVal Map As Integer, _
                           ByVal X As Integer, _
                           ByVal Y As Integer, _
                           ByVal UserIndex As Integer)

    On Error Resume Next

    Dim texto As String

    Dim Name  As String

    texto = "COMO OSAS MOLESTARME , NO SABES QUIEN SOY?, SATUROS LA PERSONA MAS VIL Y TRAIDORA DE AOMANIA ...."

    Dim MiObj As Obj

    If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ObjType = eOBJType.otTumba Then
  
        Call SendData(ToIndex, UserIndex, 0, "MCAR" & texto & Chr(176) & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).GrhSecundario)

    End If
        
    If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
      
    If UserList(UserIndex).flags.Paralizado = 0 Then

        Call SendData(ToMap, UserIndex, UserList(UserIndex).Pos.Map, "TW16," & UserList(UserIndex).char.CharIndex)
        'Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & ",8,1")
        
        UserList(UserIndex).flags.Paralizado = 1
        UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
          
        'Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
          
        Call SendData(ToIndex, UserIndex, 0, "PARADOK")
        Call SendData(ToIndex, UserIndex, 0, "PARADO2")
           
        'Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
        Call Corr_ActualizarPosicion(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
        
        Call SendData(ToIndex, UserIndex, 0, "||Saturos te ha paralizado ..." & FONTTYPE_Motd4)

    End If
            
    If TieneObjetosRobables(UserIndex) Then

        Call robarObjetoTumba(UserIndex)

    End If
        
    UserList(UserIndex).Counters.Veneno = 0
    UserList(UserIndex).flags.Envenenado = 100
    Call SendData(ToIndex, UserIndex, 0, "||Saturos te ha envenenado por 100 ..." & FONTTYPE_VERDE)

End Sub

Function quitarObjetoTumba(VictimaIndex As Integer, Obj As Integer)

    quitarObjetoTumba = False

    Dim i As Integer

    i = 1

    Do While i <= MAX_INVENTORY_SLOTS

        'Hay objeto en este slot?
        If UserList(VictimaIndex).Invent.Object(i).ObjIndex = Obj Then
            If ObjEsRobable(VictimaIndex, i) Then
                                 
                Dim MiObj As Obj

                Dim num   As Byte

                'Cantidad al azar
                num = 1
                
                MiObj.Amount = num
                MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
    
                UserList(VictimaIndex).Invent.Object(i).Amount = UserList(VictimaIndex).Invent.Object(i).Amount - num
                
                If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then

                    Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)

                End If
            
                Call UpdateUserInv(False, VictimaIndex, CByte(i))
                Call SendData(ToIndex, VictimaIndex, 0, "||Gracias aqui tienes tu recompensa...." & FONTTYPE_INFO)
                i = 100
                quitarObjetoTumba = True
                    
            Else
                'i = 100
                Call SendData(ToIndex, VictimaIndex, 0, "||No puedes llevar equipado mi objeto...." & FONTTYPE_INFO)
                i = 100

            End If
        
        End If

        i = i + 1
    Loop

End Function

Sub robarObjetoTumba(VictimaIndex As Integer)

Dim i As Integer
i = 1
Do While i <= MAX_INVENTORY_SLOTS
        'Hay objeto en este slot?
        If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
           If ObjEsRobable(VictimaIndex, i) Then
                                 
                    Dim MiObj As Obj
                    Dim num As Byte
                    'Cantidad al azar
                    num = RandomNumber(1, 5)
                
                    If num > UserList(VictimaIndex).Invent.Object(i).Amount Then
                        num = UserList(VictimaIndex).Invent.Object(i).Amount
                    End If
                
                    MiObj.Amount = num
                    MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
    
                    UserList(VictimaIndex).Invent.Object(i).Amount = UserList(VictimaIndex).Invent.Object(i).Amount - num
                
                    If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
                            Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
                    End If
            
                    Call UpdateUserInv(False, VictimaIndex, CByte(i))
                    Call SendData(ToIndex, VictimaIndex, 0, "||Saturos te ha robado ..." & FONTTYPE_INFO)
                    'i = 100
           End If
        
        End If
        i = i + 1
    Loop

End Sub
