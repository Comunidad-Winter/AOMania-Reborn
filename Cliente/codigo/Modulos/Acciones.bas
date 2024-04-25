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

Sub Accion(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

    On Error Resume Next

    '¿Posicion valida?
    If InMapBounds(Map, X, Y) Then
   
        Dim FoundChar      As Byte
        Dim FoundSomething As Byte
        Dim TempCharIndex  As Integer
        
          
          If MapData(Map, X, Y).UserIndex > 0 Then
        
        If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Creditos Then

            If Distancia(Npclist(MapData(Map, X, Y).NpcIndex).pos, UserList(UserIndex).pos) > 4 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
                Exit Sub

            End If

            If UserList(UserIndex).flags.Montado = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes comerciar estando arriba de tu Mascota!" & FONTTYPE_INFO)
                Exit Sub

            End If

            'A depositar de una
            Call Mod_Monedas.IniciarComercioCreditos(UserIndex)
            Exit Sub
            
        ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Canjes Then
                 
            If Distancia(Npclist(MapData(Map, X, Y).NpcIndex).pos, UserList(UserIndex).pos) > 4 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
                Exit Sub

            End If

            If UserList(UserIndex).flags.Montado = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes comerciar estando arriba de tu Mascota!" & FONTTYPE_INFO)
                Exit Sub

            End If

            'A depositar de una
            Call Mod_Monedas.IniciarComercioCanjes(UserIndex)
            Exit Sub
                 
        End If
       
        '¿Es un obj?
        If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.ObjIndex
        
            Select Case ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType
            
                Case eOBJType.otPuertas 'Es una puerta
                    Call AccionParaPuerta(Map, X, Y, UserIndex)

                Case eOBJType.otCARTELES 'Es un cartel
                    Call AccionParaCartel(Map, X, Y, UserIndex)

                Case eOBJType.otFOROS 'Foro
                    Call AccionParaForo(Map, X, Y, UserIndex)

                Case eOBJType.otLeña    'Leña

                    If MapData(Map, X, Y).OBJInfo.ObjIndex = FOGATA_APAG And UserList(UserIndex).flags.Muerto = 0 Then
                        Call AccionParaRamita(Map, X, Y, UserIndex)

                    End If

            End Select

            '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
        ElseIf MapData(Map, X + 1, Y).OBJInfo.ObjIndex > 0 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y).OBJInfo.ObjIndex
            Call SendData(SendTarget.toIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).OBJType & "," & ObjData( _
                MapData(Map, X + 1, Y).OBJInfo.ObjIndex).Name & "," & "OBJ")

            Select Case ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).OBJType
            
                Case 6 'Es una puerta
                    Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
            
            End Select
            
            'USUARIO
        ElseIf MapData(Map, X, Y).UserIndex > 0 Then

            If UserList(UserIndex).flags.Privilegios <> PlayerType.User Then
                UserList(UserIndex).flags.SeleccioneA = UserList(MapData(Map, X, Y).UserIndex).Name
                UserList(UserList(UserIndex).flags.TargetUser).flags.EstoySelec = 1
        
                Call SendData(SendTarget.toIndex, UserList(UserIndex).flags.TargetUser, 0, "||El GM te a seleccionado para teletransportarte." & _
                    FONTTYPE_INFO)
        
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Seleccionaste a: " & UserList(UserIndex).flags.SeleccioneA & FONTTYPE_INFO)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "TX")

            End If

            'USUARIO

        ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex > 0 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex
            Call SendData(SendTarget.toIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).OBJType & "," & ObjData( _
                MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).Name & "," & "OBJ")

            Select Case ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).OBJType
            
                Case 6 'Es una puerta
                    Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)
            
            End Select

        ElseIf MapData(Map, X, Y + 1).OBJInfo.ObjIndex > 0 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y + 1).OBJInfo.ObjIndex
            Call SendData(SendTarget.toIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).OBJType & "," & ObjData( _
                MapData(Map, X, Y + 1).OBJInfo.ObjIndex).Name & "," & "OBJ")

            Select Case ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).OBJType
            
                Case 6 'Es una puerta
                    Call AccionParaPuerta(Map, X, Y + 1, UserIndex)
            
            End Select
            
    
        ElseIf MapData(Map, X, Y).UserIndex > 0 Then
    
        ElseIf MapData(Map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
            'Set the target NPC
            UserList(UserIndex).flags.TargetNpc = MapData(Map, X, Y).NpcIndex
        
            If Npclist(MapData(Map, X, Y).NpcIndex).Comercia = 1 Then
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 3 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
                    Exit Sub
                End If
                
                If UserList(UserIndex).flags.Montado = True Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes comerciar estando arriba de tu Mascota!" & FONTTYPE_INFO)
                    Exit Sub

                End If
                                   
                'Iniciamos la rutina pa' comerciar.
                Call IniciarCOmercioNPC(UserIndex)
                    
            ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Banquero Then

                If Distancia(Npclist(MapData(Map, X, Y).NpcIndex).pos, UserList(UserIndex).pos) > 4 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Montado = True Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes usar la boveda estando arriba de tu Mascota!" & FONTTYPE_INFO)
                    Exit Sub

                End If

                'A depositar de una
                Call IniciarDeposito(UserIndex)
        
            ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Revividor Then

                If Distancia(UserList(UserIndex).pos, Npclist(MapData(Map, X, Y).NpcIndex).pos) > 10 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "Z32")
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Envenenado = 1 Then
                    UserList(UserIndex).flags.Envenenado = 0
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Te has curado del envenenamiento." & FONTTYPE_INFO)

                End If

                'Revivimos si es necesario
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call RevivirUsuario(UserIndex)

                End If
            
                'curamos totalmente
                UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
            
                Call EnviarHP(UserIndex)

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

    Dim pos As WorldPos
    pos.Map = Map
    pos.X = X
    pos.Y = Y

    If Distancia(pos, UserList(UserIndex).pos) > 2 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
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
            Call SendData(SendTarget.toIndex, UserIndex, 0, "FMSG" & tit & Chr(176) & men)
        
        Next

    End If

    Call SendData(SendTarget.toIndex, UserIndex, 0, "MFOR")

End Sub

Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim MiObj As Obj
    Dim wp    As WorldPos

    If Not (Distance(UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y, X, Y) > 2) Then
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
                    SendData SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_PUERTA
                    
                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||La puerta esta cerrada con llave." & FONTTYPE_INFO)

                End If

            Else
                'Cierra puerta
                MapData(Map, X, Y).OBJInfo.ObjIndex = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).IndexCerrada
                
                Call ModAreas.SendToAreaByPos(Map, X, Y, "HO" & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).GrhIndex & "," & X & "," & Y)
                
                MapData(Map, X, Y).Blocked = 1
                MapData(Map, X - 1, Y).Blocked = 1
                
                Call Bloquear(SendTarget.ToMap, 0, Map, Map, X - 1, Y, 1)
                Call Bloquear(SendTarget.ToMap, 0, Map, Map, X, Y, 1)
                
                SendData SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_PUERTA

            End If
        
            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.ObjIndex
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La puerta esta cerrada con llave." & FONTTYPE_INFO)

        End If

    Else
        Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")

    End If

End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim MiObj As Obj

    If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = 8 Then
  
        If Len(ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).texto) > 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "MCAR" & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).texto & Chr(176) & ObjData( _
                MapData(Map, X, Y).OBJInfo.ObjIndex).GrhSecundario)

        End If
  
    End If

End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim Suerte As Byte
    Dim exito  As Byte
    Dim Obj    As Obj
    Dim raise  As Integer

    Dim pos    As WorldPos
    pos.Map = Map
    pos.X = X
    pos.Y = Y

    If Distancia(pos, UserList(UserIndex).pos) > 2 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
        Exit Sub

    End If

    If MapData(Map, X, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Pk = False Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||En zona segura no puedes hacer fogatas." & FONTTYPE_INFO)
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
        If MapInfo(UserList(UserIndex).pos.Map).Zona <> Ciudad Then
            Obj.ObjIndex = FOGATA
            Obj.Amount = 1
        
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has prendido la fogata." & FONTTYPE_INFO)
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "FO")
        
            Call MakeObj(ToMap, 0, Map, Obj, Map, X, Y)
        
            'Las fogatas prendidas se deben eliminar
            Dim Fogatita As New cGarbage
            Fogatita.Map = Map
            Fogatita.X = X
            Fogatita.Y = Y
            Call TrashCollector.Add(Fogatita)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La ley impide realizar fogatas en las ciudades." & FONTTYPE_INFO)
            Exit Sub

        End If

    Else
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No has podido hacer fuego." & FONTTYPE_INFO)

    End If

    'Sino tiene hambre o sed quizas suba el skill supervivencia
    If UserList(UserIndex).flags.Hambre = 0 And UserList(UserIndex).flags.Sed = 0 Then
        Call SubirSkill(UserIndex, Supervivencia)

    End If

End Sub
