Attribute VB_Name = "TCP_HandleData2"

Option Explicit

Public Sub HandleData_2(ByVal UserIndex As Integer, rData As String, ByRef Procesado As Boolean)

    Dim LoopC    As Integer
    Dim nPos     As WorldPos
    Dim tStr     As String
    Dim tInt     As Integer
    Dim tLong    As Long
    Dim TIndex   As Integer
    Dim tName    As String
    Dim tMessage As String
    Dim AuxInd   As Integer
    Dim Arg1     As String
    Dim Arg2     As String
    Dim Arg3     As String
    Dim Arg4     As String
    Dim Ver      As String
    Dim encpass  As String
    Dim Pass     As String
    Dim Mapa     As Integer
    Dim Name     As String
    Dim ind
    Dim n        As Integer
    Dim wpaux    As WorldPos
    Dim mifile   As Integer
    Dim X        As Integer
    Dim Y        As Integer
    Dim DummyInt As Integer
    Dim T()      As String
    Dim i        As Integer
    Dim GuildName As String

    Procesado = True
        
    If UCase$(Left(rData, 3)) = "/SI" Then
        If Encuesta.ACT = 0 Then Exit Sub
        If UserList(UserIndex).flags.VotEnc = True Then Exit Sub
        Encuesta.EncSI = Encuesta.EncSI + 1
        Call SendData(SendTarget.toindex, UserIndex, 0, "||Has votado exitosamente." & FONTTYPE_INFO)
        UserList(UserIndex).flags.VotEnc = True
        Exit Sub
    End If

    If UCase$(Left(rData, 3)) = "/NO" Then
        If Encuesta.ACT = 0 Then Exit Sub
        If UserList(UserIndex).flags.VotEnc = True Then Exit Sub
        Encuesta.EncNO = Encuesta.EncNO + 1
        Call SendData(SendTarget.toindex, UserIndex, 0, "||Has votado exitosamente." & FONTTYPE_INFO)
        UserList(UserIndex).flags.VotEnc = True
        Exit Sub
    End If
    
    If UCase(Left(rData, 11)) = "/TELEPATIA " Then
        
        rData = Right$(rData, Len(rData) - 11)
      
        tName = ReadField(1, rData, 32)
        tMessage = Right$(rData, Len(rData) - (1 + Len(tName)))
        TIndex = NameIndex(tName)
    
        If UserList(UserIndex).Telepatia = 1 Then
            If TIndex <= 0 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(TIndex).Telepatia = 0 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Este usuario no sabe usar la telepatia." & FONTTYPE_INFO)
                Exit Sub
            End If
            
        ElseIf UserList(UserIndex).Telepatia = 0 Then
              Call SendData(SendTarget.toindex, UserIndex, 0, "||A�n no sabes usar la telepatia." & FONTTYPE_INFO)
        End If
    
        If UserList(TIndex).flags.Privilegios = PlayerType.User Then

            If UserList(UserIndex).Telepatia = 1 Then
    
                Call SendData(SendTarget.toindex, TIndex, 0, "||< " & UserList(UserIndex).Name & " > te dice: " & tMessage & FONTTYPE_SERVER)
  
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Le has mandado a " & tName & " : " & tMessage & FONTTYPE_SERVER)
                
                Call LogTelepatia(UserList(UserIndex).Name, tName, tMessage)

            End If

        End If

    End If
    
    Select Case UCase$(rData)

        Case "/MAYOR"
               Call CommandMayor(UserIndex)
         Exit Sub
         
        Case "/ONLINE"
        
            'No se envia m�s la lista completa de usuarios
            n = 0
            tStr = vbNullString
             
            For LoopC = 1 To LastUser

                If Len(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.Privilegios <= PlayerType.Consejero Then
                    n = n + 1
                    tStr = tStr & UserList(LoopC).Name & ", "

                End If

            Next LoopC
          
            If n > 0 Then
                tStr = Left$(tStr, Len(tStr) - 2)
                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & tStr & "." & FONTTYPE_INFO)
                Call SendData(SendTarget.toindex, UserIndex, 0, "||N�mero de usuarios: " & n & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toindex, UserIndex, 0, "||No hay usuarios Online." & FONTTYPE_INFO)

            End If
             
            Exit Sub
            
        Case "/RANKCLAN"
             Call modGuilds.UpdateRankGuild(UserIndex)
            Exit Sub

        Case "/CASTILLOS"
            Call SendInfoCastillos(UserIndex)

        Case "/CASTILLO ESTE"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Estas muerto!" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call WarpCastillo(UserIndex, "ESTE")
            Exit Sub

        Case "/CASTILLO OESTE"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Estas muerto!" & FONTTYPE_INFO)
                Exit Sub
            End If
            Call WarpCastillo(UserIndex, "OESTE")
            Exit Sub

        Case "/CASTILLO NORTE"
           If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Estas muerto!" & FONTTYPE_INFO)
                Exit Sub
            End If
            Call WarpCastillo(UserIndex, "NORTE")
            Exit Sub

        Case "/CASTILLO SUR"
        If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Estas muerto!" & FONTTYPE_INFO)
                Exit Sub
            End If
            Call WarpCastillo(UserIndex, "SUR")
            Exit Sub

        Case "/FORTALEZA"
        If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Estas muerto!" & FONTTYPE_INFO)
                Exit Sub
            End If
            Call WarpCastillo(UserIndex, "FORTALEZA")
            Exit Sub
        
        Case "/FORTALEZAFUERTE"
        
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Estas muerto!" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            GuildName = Guilds(UserList(UserIndex).GuildIndex).GuildName
            
            If GuildName = Norte And GuildName = Sur And GuildName = Oeste And GuildName = Este And GuildName = Fortaleza Then
            
              Call WarpUserChar(UserIndex, 164, 45, 52, True)
            
              Else
              
              Call SendData(SendTarget.toindex, UserIndex, 0, "||Debes tener todos los castillos y la fortaleza para ir a fortaleza fuerte." & FONTTYPE_INFO)
            
            End If
             
            Exit Sub

        Case "/DUELOS"

            If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Primero tienes que seleccionar un personaje, hace click izquierdo sobre el." & _
                        FONTTYPE_INFO)
                Exit Sub
            End If

            Dim JuanpaDuelosMap As Integer
            JuanpaDuelosMap = MAPADUELO
            Dim JuanpaDuelosX As Integer
            JuanpaDuelosX = RandomNumber(43, 58)
            Dim JuanpaDuelosY As Integer
            JuanpaDuelosY = RandomNumber(45, 56)

            '�El NPC puede comerciar?
           
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Duelos Then
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "Z27")
                    Exit Sub
                End If

                If UserList(UserIndex).flags.Navegando = 1 Then
 
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbYellow & "�" & "No puedes entrar a duelos estando navegando!!!" & "�" _
                            & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                    Exit Sub

                End If
            
                If UserList(UserIndex).flags.Privilegios = PlayerType.Dios Or UserList(UserIndex).flags.Privilegios = PlayerType.SemiDios Or _
                        UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbCyan & "�" & _
                            "��No puedes entrar a duelos eres GM teletransportate al mapa " & JuanpaDuelosMap & "!!" & "�" & CStr(Npclist(UserList( _
                            UserIndex).flags.TargetNpc).char.CharIndex))
                    Exit Sub

                End If
            
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||�� Est�s muerto !!" & FONTTYPE_INFO)
                    Exit Sub
                ElseIf UserList(UserIndex).Stats.ELV < 25 Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes hacer duelos siendo menor a nivel 25." & FONTTYPE_INFO)
                    Exit Sub
                ElseIf MapInfo(JuanpaDuelosMap).NumUsers >= 2 Then
                
                    If MapInfo(JuanpaDuelosMap).NumUsers = 2 Then
                
                        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Duelos Then
                            Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbYellow & "�" & _
                                    "��El mapa de duelos est� ocupado ahora mismo.!!" & "�" & CStr(Npclist(UserList( _
                                    UserIndex).flags.TargetNpc).char.CharIndex))
                            Exit Sub

                        End If
                
                        Exit Sub

                    End If

                ElseIf MapInfo(UserList(UserIndex).pos.Map).Pk = True Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||Estas en una zona insegura." & FONTTYPE_WARNING)
                    Exit Sub
                ElseIf MapInfo(JuanpaDuelosMap).NumUsers = 1 Then
                    duelosreta = UserIndex
                 
                    If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto Then
                        UserList(UserIndex).Counters.Invisibilidad = IntervaloInvisible
                        UserList(UserIndex).flags.Invisible = 0
                        UserList(UserIndex).Counters.Ocultando = 0
                        Call SendData(SendTarget.toindex, UserIndex, 0, "INVI0")

                    End If
                 
                    Call WarpUserChar(UserIndex, JuanpaDuelosMap, JuanpaDuelosX, JuanpaDuelosY, True)
       
                    Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos: ���" & UserList(duelosreta).Name & " acepto el desafio!!!" & _
                            FONTTYPE_TALK)
                    Exit Sub
                ElseIf MapInfo(JuanpaDuelosMap).NumUsers = 0 Then
                    duelosespera = UserIndex
                
                    If UserList(UserIndex).flags.Oculto = 1 Then
                        UserList(UserIndex).Counters.Ocultando = 0
                        UserList(UserIndex).flags.Oculto = 0
                        Call SendData(SendTarget.toindex, UserIndex, 0, "INVI0")

                    End If
                
                    If UserList(UserIndex).flags.Invisible = 1 Then
                        UserList(UserIndex).Counters.Invisibilidad = IntervaloInvisible
                        UserList(UserIndex).flags.Invisible = 0
                        UserList(UserIndex).Counters.Ocultando = 0
                        Call SendData(SendTarget.toindex, UserIndex, 0, "INVI0")

                    End If
                 
                    Call WarpUserChar(UserIndex, JuanpaDuelosMap, JuanpaDuelosX, JuanpaDuelosY, True)
                    Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos: " & UserList(duelosespera).Name & _
                            " espera rival en la sala de torneos." & FONTTYPE_TALK)

                End If
            
            End If

            Exit Sub
        
        Case "/SALIR"

            If UserList(UserIndex).flags.Montado = True Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes salir estando en montado en tu mascota!." & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).flags.Paralizado = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes salir estando paralizado." & FONTTYPE_WARNING)
                Exit Sub

            End If

            ''mato los comercios seguros
            If UserList(UserIndex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                        Call SendData(SendTarget.toindex, UserList(UserIndex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & _
                                FONTTYPE_TALK)
                        Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)

                    End If

                End If

                Call SendData(SendTarget.toindex, UserIndex, 0, "||Comercio cancelado. " & FONTTYPE_TALK)
                Call FinComerciarUsu(UserIndex)

            End If

            Call Cerrar_Usuario(UserIndex)
            Exit Sub

        Case "/SALIRCLAN"
            'obtengo el guildindex
            tInt = m_EcharMiembroDeClan(UserIndex, UserList(UserIndex).Name, False)
            
            If tInt > 0 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Dejas el clan." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(UserIndex).Name & " deja el clan." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Tu no puedes salir de ning�n clan." & FONTTYPE_GUILD)

            End If
            
            Exit Sub
            
        Case "/BALANCE"

            '�Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z30")
                Exit Sub

            End If

            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 3 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z27")
                Exit Sub

            End If

            Select Case Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype

                Case eNPCType.Banquero

                    If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
                        Call SendData(SendTarget.toindex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                        CloseSocket (UserIndex)
                        Exit Sub

                    End If

                    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "�" & "Tienes " & UserList(UserIndex).Stats.Banco & _
                            " monedas de oro en tu cuenta." & "�" & Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex & FONTTYPE_INFO)

                Case eNPCType.Timbero

                    If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                        tLong = Apuestas.Ganancias - Apuestas.Perdidas
                        n = 0

                        If tLong >= 0 And Apuestas.Ganancias <> 0 Then
                            n = Int(tLong * 100 / Apuestas.Ganancias)

                        End If

                        If tLong < 0 And Apuestas.Perdidas <> 0 Then
                            n = Int(tLong * 100 / Apuestas.Perdidas)

                        End If

                        Call SendData(SendTarget.toindex, UserIndex, 0, "||Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & _
                                " Ganancia Neta: " & tLong & " (" & n & "%) Jugadas: " & Apuestas.Jugadas & FONTTYPE_INFO)

                    End If

            End Select

            Exit Sub

        Case "/QUIETO" ' << Comando a mascotas

            '�Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z30")
                Exit Sub

            End If

            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z27")
                Exit Sub

            End If

            If Npclist(UserList(UserIndex).flags.TargetNpc).MaestroUser <> UserIndex Then Exit Sub
            Npclist(UserList(UserIndex).flags.TargetNpc).Movement = TipoAI.ESTATICO
            Call Expresar(UserList(UserIndex).flags.TargetNpc, UserIndex)
            Exit Sub
            
        Case "/NAVEGAR"

            '�Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||��Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            '�El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
                Exit Sub

            End If

            '�El NPC puede comerciar?
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 15 Then
                If Len(Npclist(UserList(UserIndex).flags.TargetNpc).Desc) > 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, UserList( _
                        UserIndex).pos.Map, "||" & vbWhite & "�" & "Yo no administro las navegaciones." & "�" & str(Npclist(UserList( _
                        UserIndex).flags.TargetNpc).char.CharIndex))
                Exit Sub

            End If

            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 4 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                Exit Sub

            End If

            Dim iui       As Integer
            Dim TienePass As Boolean

            If Barcos.Pasajeros >= MAX_PASAJEROS Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & _
                        "�Lo lamento pero el barco esta lleno, deberas esperar hasta la pr�xima embarcaci�n." & "�" & Npclist(UserList( _
                        UserIndex).flags.TargetNpc).char.CharIndex & FONTTYPE_INFO)

            End If

            For iui = 1 To MAX_INVENTORY_SLOTS

                If UserList(UserIndex).Invent.Object(iui).ObjIndex > 0 Then
                    If ObjData(UserList(UserIndex).Invent.Object(iui).ObjIndex).OBJType = eOBJType.otPasaje Then

                        If ObjData(UserList(UserIndex).Invent.Object(iui).ObjIndex).Zona = Barcos.Zona Then
                            If Barcos.TiempoRest > 0 And Barcos.TiempoRest < 11 Then
                                UserList(UserIndex).Invent.Object(iui).Amount = UserList(UserIndex).Invent.Object(iui).Amount - 1

                                If (UserList(UserIndex).Invent.Object(iui).Amount <= 0) Then
                                    UserList(UserIndex).Invent.Object(iui).Amount = 0
                                    UserList(UserIndex).Invent.Object(iui).ObjIndex = 0

                                End If

                                If Not InMapBounds(245, 50, 50) Then Exit Sub
                                Call WarpUserChar(UserIndex, 245, 50, 50, False)
                                UserList(UserIndex).flags.Embarcado = 1
                                UserList(UserIndex).Zona = Barcos.Zona
                                Barcos.Pasajeros = Barcos.Pasajeros + 1
                                Call UpdateUserInv(False, UserIndex, iui)
                            ElseIf Barcos.TiempoRest < 1 Then
                                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "�Lo lamento pero la embarcacion ya a partido." & _
                                        "�" & Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex & FONTTYPE_INFO)
                            Else
                                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & _
                                        "�La embarcacion partira en un rato, mientras ve a pasear." & "�" & Npclist(UserList( _
                                        UserIndex).flags.TargetNpc).char.CharIndex & FONTTYPE_INFO)

                            End If

                            Exit Sub
                        Else
                            TienePass = True

                        End If

                    End If

                End If

            Next iui
        
            If Not TienePass Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "�Tu no tienes pasaje." & "�" & Npclist(UserList( _
                        UserIndex).flags.TargetNpc).char.CharIndex & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "�Ese pasaje no es para esta embarcacion." & "�" & Npclist( _
                        UserList(UserIndex).flags.TargetNpc).char.CharIndex & FONTTYPE_INFO)

            End If

            Exit Sub

        Case "/ENTRENAR"

            '�Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z30")
                Exit Sub

            End If

            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z27")
                Exit Sub

            End If

            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> eNPCType.Entrenador Then Exit Sub
            Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNpc)
            Exit Sub
        
        Case "/DESCANSAR"

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            If HayOBJarea(UserList(UserIndex).pos, FOGATA) Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "DOK")

                If Not UserList(UserIndex).flags.Descansar Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||Te acomodas junto a la fogata y comenzas a descansar." & FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)

                End If

                UserList(UserIndex).flags.Descansar = Not UserList(UserIndex).flags.Descansar
            Else

                If UserList(UserIndex).flags.Descansar Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)
                        
                    UserList(UserIndex).flags.Descansar = False
                    Call SendData(SendTarget.toindex, UserIndex, 0, "DOK")
                    Exit Sub

                End If

                Call SendData(SendTarget.toindex, UserIndex, 0, "||No hay ninguna fogata junto a la cual descansar." & FONTTYPE_INFO)

            End If

            Exit Sub

        Case "/MEDITAR"

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            If UserList(UserIndex).Stats.MaxMAN = 0 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Solo las clases m�gicas conocen el arte de la meditaci�n" & FONTTYPE_INFO)
                Exit Sub

            End If

            'If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
            'UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
            'Call SendData(SendTarget.toindex, UserIndex, 0, "||Mana restaurado" & FONTTYPE_VENENO)
            'Call EnviarMn(UserIndex)
            'Exit Sub

            'End If
            
            If UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN Then
                Exit Sub

            End If
            
            Call SendData(SendTarget.toindex, UserIndex, 0, "MEDOK")

            If Not UserList(UserIndex).flags.Meditando Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z23")
            Else
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z16")

            End If

            UserList(UserIndex).flags.Meditando = Not UserList(UserIndex).flags.Meditando

            'Barrin 3/10/03 Tiempo de inicio al meditar
            If UserList(UserIndex).flags.Meditando Then
                UserList(UserIndex).Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z37")
                
                UserList(UserIndex).char.loops = LoopAdEternum
                Call FxDoMeditar(UserIndex)

            Else
                UserList(UserIndex).Counters.bPuedeMeditar = False
                
                UserList(UserIndex).char.FX = 0
                UserList(UserIndex).char.loops = 0
                Call SendData(SendTarget.ToMap, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & 0 & "," _
                        & 0)

            End If

            Exit Sub

        Case "/PARTICIPAR"

            If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
                Exit Sub

            End If
            
            If UserList(UserIndex).flags.Invisible = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes ir a eventos estando invisible!." & FONTTYPE_WARNING)
                Exit Sub

            End If
      
            If UserList(UserIndex).flags.Oculto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes ir a eventos estando invisible!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EstaDueleando1 = True Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes ir a torneos estando plantes!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Estas muerto!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserIndex = Team.Pj1 Or UserIndex = Team.Pj2 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
                Exit Sub

            End If
            
            If UserList(UserIndex).Stats.ELV < lvlTorneo Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Debes ser lvl " & lvlTorneo & " o mas para entrar al torneo!" & FONTTYPE_INFO)
                Exit Sub

            End If
       
            Call Torneos_Entra(UserIndex)
            Exit Sub

       Case "/GUERRA"
        
            If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, _
                        "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
                Exit Sub

            End If
       
            Call CommandGuerra(UserIndex)
            Exit Sub
            
        Case "/MEDUSA"
        
            If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, _
                        "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
                Exit Sub

            End If
       
            Call CommandMedusa(UserIndex)
            Exit Sub

        Case "/MEZCLAR"
               
               If Not TieneObjetos(Plumas.Ampere, 1, UserIndex) Or Not TieneObjetos(Plumas.Bassinger, 1, UserIndex) Or Not TieneObjetos(Plumas.Seth, 1, UserIndex) Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||Para realizar una mezcla necesitas " _
                     & Chr(147) & ObjData(Plumas.Ampere).Name & Chr(147) & ", " & Chr(147) & ObjData(Plumas.Bassinger).Name & Chr(147) & " y " & Chr(147) & ObjData(Plumas.Seth).Name & Chr(147) & "." & FONTTYPE_TALK)
                    Exit Sub
               End If
               
               With UserList(UserIndex)
                     
                     If .Faccion.ArmadaReal = 0 And .Faccion.FuerzasCaos = 0 And .Faccion.Nemesis = 0 And .Faccion.Templario = 0 Then
                         Call SendData(SendTarget.toindex, UserIndex, 0, "||Solo las armadas pueden crear alas de faccion." & FONTTYPE_INFO)
                         Exit Sub
                     End If
                     
                      If .Faccion.ArmadaReal = 1 Then
                        
                        If TieneObjetos(AlasReal.Four, 1, UserIndex) Then
                            Call SendData(SendTarget.toindex, UserIndex, 0, "||Ya tienes las alas mejoradas al maximo." & FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If TieneObjetos(AlasReal.One, 1, UserIndex) Then
                            Call MezclarAlas(UserIndex, AlasReal.One, AlasReal.Second)
                            Exit Sub
                        End If
                        
                        If TieneObjetos(AlasReal.Second, 1, UserIndex) Then
                            Call MezclarAlas(UserIndex, AlasReal.Second, AlasReal.Thir)
                            Exit Sub
                        End If
                        
                        If TieneObjetos(AlasReal.Thir, 1, UserIndex) Then
                            Call MezclarAlas(UserIndex, AlasReal.Thir, AlasReal.Four)
                            Exit Sub
                        End If
                        
                        If Not TieneObjetos(AlasReal.One, 1, UserIndex) And Not TieneObjetos(AlasReal.Second, 1, UserIndex) And Not TieneObjetos(AlasReal.Thir, 1, UserIndex) And Not TieneObjetos(AlasReal.Four, 1, UserIndex) Then
                            Call MezclarAlas(UserIndex, "0", AlasReal.One)
                            Exit Sub
                        End If
                     End If
                     
                     If .Faccion.FuerzasCaos = 1 Then
                        
                        If TieneObjetos(AlasCaos.Four, 1, UserIndex) Then
                            Call SendData(SendTarget.toindex, UserIndex, 0, "||Ya tienes las alas mejoradas al maximo." & FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If TieneObjetos(AlasCaos.One, 1, UserIndex) Then
                            Call MezclarAlas(UserIndex, AlasCaos.One, AlasCaos.Second)
                            Exit Sub
                        End If
                        
                        If TieneObjetos(AlasCaos.Second, 1, UserIndex) Then
                            Call MezclarAlas(UserIndex, AlasCaos.Second, AlasCaos.Thir)
                            Exit Sub
                        End If
                        
                        If TieneObjetos(AlasCaos.Thir, 1, UserIndex) Then
                            Call MezclarAlas(UserIndex, AlasCaos.Thir, AlasCaos.Four)
                            Exit Sub
                        End If
                        
                        If Not TieneObjetos(AlasCaos.One, 1, UserIndex) And Not TieneObjetos(AlasCaos.Second, 1, UserIndex) And Not TieneObjetos(AlasCaos.Thir, 1, UserIndex) And Not TieneObjetos(AlasCaos.Four, 1, UserIndex) Then
                            Call MezclarAlas(UserIndex, "0", AlasCaos.One)
                            Exit Sub
                        End If
                     End If
                     
                      If .Faccion.Templario = 1 Then
                        
                        If TieneObjetos(AlasTemplario.Four, 1, UserIndex) Then
                            Call SendData(SendTarget.toindex, UserIndex, 0, "||Ya tienes las alas mejoradas al maximo." & FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If TieneObjetos(AlasTemplario.One, 1, UserIndex) Then
                            Call MezclarAlas(UserIndex, AlasTemplario.One, AlasTemplario.Second)
                            Exit Sub
                        End If
                        
                        If TieneObjetos(AlasTemplario.Second, 1, UserIndex) Then
                            Call MezclarAlas(UserIndex, AlasTemplario.Second, AlasTemplario.Thir)
                            Exit Sub
                        End If
                        
                        If TieneObjetos(AlasTemplario.Thir, 1, UserIndex) Then
                            Call MezclarAlas(UserIndex, AlasTemplario.Thir, AlasTemplario.Four)
                            Exit Sub
                        End If
                        
                        If Not TieneObjetos(AlasTemplario.One, 1, UserIndex) And Not TieneObjetos(AlasTemplario.Second, 1, UserIndex) And Not TieneObjetos(AlasTemplario.Thir, 1, UserIndex) And Not TieneObjetos(AlasTemplario.Four, 1, UserIndex) Then
                            Call MezclarAlas(UserIndex, "0", AlasTemplario.One)
                            Exit Sub
                        End If
                     End If
                     
                      If .Faccion.Nemesis = 1 Then
                        
                        If TieneObjetos(AlasNemesis.Four, 1, UserIndex) Then
                            Call SendData(SendTarget.toindex, UserIndex, 0, "||Ya tienes las alas mejoradas al maximo." & FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If TieneObjetos(AlasNemesis.One, 1, UserIndex) Then
                            Call MezclarAlas(UserIndex, AlasNemesis.One, AlasNemesis.Second)
                            Exit Sub
                        End If
                        
                        If TieneObjetos(AlasNemesis.Second, 1, UserIndex) Then
                            Call MezclarAlas(UserIndex, AlasNemesis.Second, AlasNemesis.Thir)
                            Exit Sub
                        End If
                        
                        If TieneObjetos(AlasNemesis.Thir, 1, UserIndex) Then
                            Call MezclarAlas(UserIndex, AlasNemesis.Thir, AlasNemesis.Four)
                            Exit Sub
                        End If
                        
                        If Not TieneObjetos(AlasNemesis.One, 1, UserIndex) And Not TieneObjetos(AlasNemesis.Second, 1, UserIndex) And Not TieneObjetos(AlasNemesis.Thir, 1, UserIndex) And Not TieneObjetos(AlasNemesis.Four, 1, UserIndex) Then
                            Call MezclarAlas(UserIndex, "0", AlasNemesis.One)
                            Exit Sub
                        End If
                     End If
                     
               End With

        Exit Sub
        
        Case "/PROMEDIO"
            Dim Promedio
            Promedio = Round(UserList(UserIndex).Stats.MaxHP / UserList(UserIndex).Stats.ELV, 2)
            Call SendData(SendTarget.toindex, UserIndex, 0, "||El Promedio de vida de tu Personaje es de " & Promedio & FONTTYPE_ORO)
            Exit Sub
   
        Case "/AYUDA"
            Call SendHelp(UserIndex)
            Exit Sub
        
        Case "/SEG"

            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "OFFOFS")
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Aviso: Has desactivado el seguro." & FONTTYPE_RETOS)
            Else
                Call SendData(SendTarget.toindex, UserIndex, 0, "ONONS")
                 Call SendData(SendTarget.toindex, UserIndex, 0, "||Aviso: Has activado el seguro." & FONTTYPE_RETOS)
            End If

            UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
            Exit Sub
            
        Case "/SEGCLAN"

            If UserList(UserIndex).flags.SeguroClan Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "SEGCO99")
                UserList(UserIndex).flags.SeguroClan = False
                 Call SendData(SendTarget.toindex, UserIndex, 0, "||Aviso: Has activado el seguro de clan." & FONTTYPE_RETOS)
            Else
                Call SendData(SendTarget.toindex, UserIndex, 0, "SEG108")
                UserList(UserIndex).flags.SeguroClan = True
                 Call SendData(SendTarget.toindex, UserIndex, 0, "||Aviso: Has desactivado el seguro de clan." & FONTTYPE_RETOS)
            End If

            Exit Sub
            
        Case "/SEGCMBT"

            If UserList(UserIndex).flags.SeguroCombate Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "SEG11")
                UserList(UserIndex).flags.SeguroCombate = False
                 Call SendData(SendTarget.toindex, UserIndex, 0, "||Has salido del modo de combate." & FONTTYPE_RETOS2V2)
            Else
                Call SendData(SendTarget.toindex, UserIndex, 0, "SEG10")
                UserList(UserIndex).flags.SeguroCombate = True
                  Call SendData(SendTarget.toindex, UserIndex, 0, "||Has pasado al modo de combate." & FONTTYPE_RETOS2V2)
                  End If

            Exit Sub
            
        Case "/SEGOBJT"

            If UserList(UserIndex).flags.SeguroObjetos Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "SEG13")
                UserList(UserIndex).flags.SeguroObjetos = False
                 Call SendData(SendTarget.toindex, UserIndex, 0, "||Aviso: Seguro de objeto activado." & FONTTYPE_RETOS)
            Else
                Call SendData(SendTarget.toindex, UserIndex, 0, "SEG12")
                UserList(UserIndex).flags.SeguroObjetos = True
                 Call SendData(SendTarget.toindex, UserIndex, 0, "||Aviso: Seguro de objeto desactivado." & FONTTYPE_RETOS)
            End If

            Exit Sub
            
        Case "/SEGHZS"

            If Not UserList(UserIndex).flags.SeguroHechizos Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "SEG15")
                UserList(UserIndex).flags.SeguroHechizos = True
                   Call SendData(SendTarget.toindex, UserIndex, 0, "||Aviso: Seguro de mover hechizos activado." & FONTTYPE_RETOS)
            
            Else
                Call SendData(SendTarget.toindex, UserIndex, 0, "SEG14")
                UserList(UserIndex).flags.SeguroHechizos = False
                  Call SendData(SendTarget.toindex, UserIndex, 0, "||Aviso: Seguro de mover hechizos desactivado." & FONTTYPE_RETOS)
                End If
            Exit Sub
         
        Case "/COMERCIAR"

            '�Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            If UserList(UserIndex).flags.Montado = True Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||�Debes Demontarte para poder Comerciar!.!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).flags.Comerciando Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Ya est�s comerciando" & FONTTYPE_INFO)
                Exit Sub

            End If
            
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                Exit Sub

            End If

            '�El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNpc > 0 Then

                '�El NPC puede comerciar?
                If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                    If Len(Npclist(UserList(UserIndex).flags.TargetNpc).Desc) > 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, UserList( _
                            UserIndex).pos.Map, "||" & vbWhite & "�" & "No tengo ningun interes en comerciar." & "�" & CStr(Npclist(UserList( _
                            UserIndex).flags.TargetNpc).char.CharIndex))
                    Exit Sub

                End If

                If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 3 Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "Z27")
                    Exit Sub

                End If
 
                'Iniciamos la rutina pa' comerciar.
                 If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Creditos Then
                    Call Mod_Monedas.IniciarComercioCreditos(UserIndex)
                Else
                'Iniciamos la rutina pa' comerciar.
                Call IniciarCOmercioNPC(UserIndex)
                End If
                '[Alejo]
            ElseIf UserList(UserIndex).flags.TargetUser > 0 Then

                'Comercio con otro usuario
                'Puede comerciar ?
                If ComerciarAc = True Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||��El comercio con usuarios esta deshabilitado.!!" & FONTTYPE_INFO)
                    Exit Sub

                End If

                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||��No puedes comerciar con los muertos!!" & FONTTYPE_INFO)
                    Exit Sub

                End If

                'soy yo ?
                If UserList(UserIndex).flags.TargetUser = UserIndex Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes comerciar con vos mismo..." & FONTTYPE_INFO)
                    Exit Sub

                End If

                'ta muy lejos ?
                If Distancia(UserList(UserList(UserIndex).flags.TargetUser).pos, UserList(UserIndex).pos) > 3 Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "Z13")
                    Exit Sub

                End If

                'Ya ta comerciando ? es conmigo o con otro ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Comerciando = True And UserList(UserList( _
                        UserIndex).flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes comerciar con el usuario en este momento." & FONTTYPE_INFO)
                    Exit Sub

                End If

                'inicializa unas variables...
                UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
                UserList(UserIndex).ComUsu.DestNick = UserList(UserList(UserIndex).flags.TargetUser).Name
                UserList(UserIndex).ComUsu.Cant = 0
                UserList(UserIndex).ComUsu.Objeto = 0
                UserList(UserIndex).ComUsu.Acepto = False
                
                'Rutina para comerciar con otro usuario
                Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)
            Else
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z31")

            End If

            Exit Sub

        Case "/BANCO"

            '�Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            '�El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNpc > 0 Then
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 3 Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "Z27")
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Montado = True Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes usar el banco estando arriba de tu Mascota!" & FONTTYPE_INFO)
                    Exit Sub

                End If

                If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Banquero Then
                         Call SendData(SendTarget.toindex, UserIndex, 0, "BANP" & UserList(UserIndex).Stats.Banco & "," & UserList(UserIndex).Stats.GLD & "," & UserList(UserIndex).BancoInvent.NroItems)
                End If

            Else
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z31")

            End If

            Exit Sub
    
        Case "/ENLISTAR"

            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z30")
                Exit Sub

            End If
           
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 5 Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Debes acercarte m�s." & FONTTYPE_INFO)
                 Exit Sub
            End If

            Select Case Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion
          
                Case 0
                    Call EnlistarArmadaReal(UserIndex)

                Case 1
                    Call EnlistarCaos(UserIndex)

                Case 3
                    Call EnlistarTemplarios(UserIndex)

                Case 5
                    Call EnlistarNemesis(UserIndex)

            End Select
           
            Exit Sub
    
        Case "/CIRUGIA"
    
           If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Cirujia Then
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 5 Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "Z27")
                    Exit Sub
                End If
                 
                 If OroCirujia > UserList(UserIndex).Stats.GLD Then
                     Call SendData(SendTarget.toindex, UserIndex, 0, "||No tienes suficientes monedas de oro para la cirug�a." & FONTTYPE_INFO)
                     Exit Sub
                End If
                 
                Call IniciarChangeHead(UserIndex)
            End If
            Exit Sub
           
        Case "/RECOMPENSA"

            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z30")
                Exit Sub

            End If

            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 5 Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub

            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z32")
                Exit Sub
            End If

            Select Case Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion
            
                Case 0

                    If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbBlue & "�" & "No perteneces a la Armada del Credo, vete de aqu� o te ahogaras en tu insolencia!!" & "�" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                        Exit Sub

                    End If

                    Call RecompensaArmadaReal(UserIndex)
                    Exit Sub

                Case 1

                    If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbRed & "�" & "No perteneces a la legi�n oscura!!!" & "�" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                        Exit Sub

                    End If

                    Call RecompensaCaos(UserIndex)
                
                    Exit Sub
                
                Case 3

                    If UserList(UserIndex).Faccion.Templario = 0 Then
                        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "�" & "No perteneces a la Orden Templaria, vete de aqu� o volaras al vacio de tu ignorancia!!!" & "�" & _
                                CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                        Exit Sub

                    End If

                    Call RecompensaTemplario(UserIndex)
                    Exit Sub

                Case 5

                    If UserList(UserIndex).Faccion.Nemesis = 0 Then
                        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "�" & "No perteneces a los Caballeros de las Tinieblas, vete de aqu� o te enterraremos vivo!!!" & "�" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                        Exit Sub

                    End If

                    Call RecompensaNemesis(UserIndex)
                    Exit Sub

            End Select

            Exit Sub
           
         Exit Sub
                    
        Case "/SALIRPARTY"
            Call mdParty.SalirDeParty(UserIndex)
            Exit Sub
        
    End Select
    
    If UCase$(Left$(rData, 14)) = "/CAMBIARBARCO " Then
           rData = val(Right$(rData, Len(rData) - 14))
           
           If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Clero Then
           
           If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z27")
                Exit Sub
            End If
            
                Select Case rData
                   
                   Case 1
                     Call CambiarBarcoClero(rData, UserIndex)
                   Case 2
                     Call CambiarBarcoClero(rData, UserIndex)
                   Case 3
                    Call CambiarBarcoClero(rData, UserIndex)
                   Case 4
                     Call CambiarBarcoClero(rData, UserIndex)
                   Case 5
                     Call CambiarBarcoClero(rData, UserIndex)
                   Case 6
                     Call CambiarBarcoClero(rData, UserIndex)
                   Case Else
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbBlue & "�" & "No entiendo, Pon /CambiarBarco 1,2,3 o 4,5,6 (para recuperarlo)" & "�" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                   
              End Select
                   
           End If
           
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Abbadon Then
           
           If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z27")
                Exit Sub
            End If
            
                Select Case rData
                   
                   Case 1
                     Call CambiarBarcoAbbadon(rData, UserIndex)
                   Case 2
                     Call CambiarBarcoAbbadon(rData, UserIndex)
                   Case 3
                    Call CambiarBarcoAbbadon(rData, UserIndex)
                   Case 4
                     Call CambiarBarcoAbbadon(rData, UserIndex)
                   Case 5
                     Call CambiarBarcoAbbadon(rData, UserIndex)
                   Case 6
                     Call CambiarBarcoAbbadon(rData, UserIndex)
                   Case Else
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbRed & "�" & "No entiendo, Pon /CambiarBarco 1,2,3 o 4,5,6 (para recuperarlo)" & "�" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                   
              End Select
                   
           End If
           
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Tiniebla Then
           
           If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z27")
                Exit Sub
            End If
            
                Select Case rData
                   
                   Case 1
                     Call CambiarBarcoTiniebla(rData, UserIndex)
                   Case 2
                     Call CambiarBarcoTiniebla(rData, UserIndex)
                   Case 3
                    Call CambiarBarcoTiniebla(rData, UserIndex)
                   Case 4
                     Call CambiarBarcoTiniebla(rData, UserIndex)
                   Case 5
                     Call CambiarBarcoTiniebla(rData, UserIndex)
                   Case 6
                     Call CambiarBarcoTiniebla(rData, UserIndex)
                   Case Else
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "�" & "No entiendo, Pon /CambiarBarco 1,2,3 o 4,5,6 (para recuperarlo)" & "�" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                   
              End Select
                   
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Templario Then
           
           If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z27")
                Exit Sub
            End If
            
                Select Case rData
                   
                   Case 1
                     Call CambiarBarcoTemplario(rData, UserIndex)
                   Case 2
                     Call CambiarBarcoTemplario(rData, UserIndex)
                   Case 3
                    Call CambiarBarcoTemplario(rData, UserIndex)
                   Case 4
                     Call CambiarBarcoTemplario(rData, UserIndex)
                   Case 5
                     Call CambiarBarcoTemplario(rData, UserIndex)
                   Case 6
                     Call CambiarBarcoTemplario(rData, UserIndex)
                   Case Else
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "�" & "No entiendo, Pon /CambiarBarco 1,2,3 o 4,5,6 (para recuperarlo)" & "�" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                   
              End Select
                   
           End If
           
    End If
  
    If UCase$(Left$(rData, 6)) = "/CLAN " Then
        'clanesnuevo
        rData = Right$(rData, Len(rData) - 6)

        If UserList(UserIndex).GuildIndex > 0 Then
            Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, 0, "|+MiembroClan: " & UserList(UserIndex).Name & " dice: " & _
                    rData & FONTTYPE_GUILDMSG)
            FrmUserhablan.hClan (now & " Mensaje de " & UserList(UserIndex).Name & ":>" & rData)
            Call addConsolee(UserList(UserIndex).Name & ": " & rData, 255, 0, 0, True, False)
            Call LogUser(UserList(UserIndex).Name, "Dice en Clan: " & rData)

        End If
        
        Exit Sub

    End If
    
     If UCase$(Left$(rData, 9)) = "/LLEVAME " Then
            Dim Destino As Integer
            
            Destino = UCase$(Right$(rData, Len(rData) - 9))
            
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Teleport Then
             
               If UserList(UserIndex).Stats.GLD < 20000 Then
                   Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbYellow & "�" & "��Necesitas 20000 monedas de oro para pagar el teletransporte!!" & "�" _
                            & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                    Exit Sub
               End If
               
               Select Case Destino
                    
                     Case "1"
                       Call WarpUserChar(UserIndex, 61, 52, 60, True)
                       UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - "20000"
                       Call EnviarOro(UserIndex)
                     Exit Sub
                    
                    Case "2"
                      Call WarpUserChar(UserIndex, 34, 23, 75, True)
                      UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - "20000"
                      Call EnviarOro(UserIndex)
                     Exit Sub
                     
                    Case "3"
                       Call WarpUserChar(UserIndex, 131, 35, 23, True)
                       UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - "20000"
                      Call EnviarOro(UserIndex)
                     Exit Sub
                     
                     Case Else
                      Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbYellow & "�" & "�A ese destino no puedes ir! Solo puedes ir a /llevame 1, 2 o 3" & "�" _
                            & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                      Exit Sub
               End Select
               
            End If
            
            Exit Sub
    End If
        
    If UCase$(Left$(rData, 6)) = "/MMSG " Then
        rData = Right$(rData, Len(rData) - 6)
       
        Dim tRespuesta As String
        tRespuesta = rData
       
        If NameIndex(UserList(UserIndex).Pareja) <= 0 Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||Pareja offline." & FONTTYPE_INFO)
            Exit Sub

        End If
           
        If Len(tRespuesta) <= 0 Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||No has escrito un mensaje." & FONTTYPE_INFO)
            Exit Sub

        End If
       
        Call SendData(SendTarget.toindex, UserList(UserIndex).Pareja, 0, "||(Pareja) " & UserList(UserIndex).Name & ": " & tRespuesta & FONTTYPE_INFO)
        Call SendData(SendTarget.toindex, UserIndex, 0, "||(Pareja) " & UserList(UserIndex).Name & ": " & tRespuesta & FONTTYPE_INFO)
        Exit Sub

    End If
    
    If UCase$(Left$(rData, 6)) = "/PMSG " Then
        If Len(rData) > 6 Then
            
            If UserList(UserIndex).PartyIndex = 0 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||����No estas en Party!!!!" & FONTTYPE_VENENO)
                Exit Sub
            End If
        
            Call mdParty.BroadCastParty(UserIndex, "MiembroParty: " & UserList(UserIndex).Name & " dice: " & mid$(rData, 7) & FONTTYPE_PARTY)
            rData = Right$(rData, Len(rData) - 6)
            FrmUserhablan.hParty (now & " Mensaje de " & UserList(UserIndex).Name & ":>" & rData)
            Call LogUser(UserList(UserIndex).Name, "Dice en Party: " & rData)

        End If

        Exit Sub

    End If
    
    If UCase$(Left$(rData, 11)) = "/CENTINELA " Then

        'Evitamos overflow y underflow
        If val(Right$(rData, Len(rData) - 11)) > &H7FFF Or val(Right$(rData, Len(rData) - 11)) < &H8000 Then Exit Sub
        
        tInt = val(Right$(rData, Len(rData) - 11))
        Call CentinelaCheckClave(UserIndex, tInt)
        Exit Sub

    End If
    
    If UCase$(rData) = "/ONLINECLAN" Then
        tStr = modGuilds.m_ListaDeMiembrosOnline(UserIndex, UserList(UserIndex).GuildIndex)

        If UserList(UserIndex).GuildIndex <> 0 Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||" & Guilds(UserList(UserIndex).GuildIndex).GuildName & ": " & tStr & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toindex, UserIndex, 0, "||No pertences a ning�n clan." & FONTTYPE_INFO)

        End If

        Exit Sub

    End If
    
    If UCase$(rData) = "/ONLINEPARTY" Then
        Call mdParty.OnlineParty(UserIndex)
        Exit Sub

    End If
    
    If UCase$(Left$(rData, 6)) = "/BMSG " Then
        rData = Right$(rData, Len(rData) - 6)

        If UserList(UserIndex).flags.PertAlCons = 1 Then
            Call SendData(SendTarget.ToConsejo, UserIndex, 0, "|| (Consejero) " & UserList(UserIndex).Name & "> " & rData & FONTTYPE_CONSEJO)

        End If

        If UserList(UserIndex).flags.PertAlConsCaos = 1 Then
            Call SendData(SendTarget.ToConsejoCaos, UserIndex, 0, "|| (Consejero) " & UserList(UserIndex).Name & "> " & rData & FONTTYPE_CONSEJOCAOS)

        End If

        Exit Sub

    End If

    
    If UCase$(Left$(rData, 5)) = "/ROL " Then
        rData = Right$(rData, Len(rData) - 5)
        Call SendData(SendTarget.toindex, 0, 0, "|| " & "Su solicitud ha sido enviada" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToRolesMasters, 0, 0, "|| " & LCase$(UserList(UserIndex).Name) & " PREGUNTA ROL: " & rData & FONTTYPE_GUILDMSG)
        Exit Sub

    End If
    
    'Mensaje del servidor a GMs - Lo ubico aqui para que no se confunda con /GM [Gonzalo]
    If UCase$(Left$(rData, 3)) = "/G " And UserList(UserIndex).flags.Privilegios > PlayerType.User Then
        rData = Right$(rData, Len(rData) - 3)
        Call LogGM(UserList(UserIndex).Name, "Dice en GM Chat:" & rData)

        If rData <> "" Then
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & "> " & rData & "~255~255~255~0~1")

        End If

        Exit Sub

    End If
    
    If UCase$(Left$(rData, 4)) = "/SOS" Then
        
        If UserList(UserIndex).flags.Privilegios = PlayerType.Dios Or UserList(UserIndex).flags.Privilegios = PlayerType.SemiDios Or UserList( _
                UserIndex).flags.Privilegios = PlayerType.Consejero Then
            Exit Sub

        End If
        
        Call SendData(SendTarget.toindex, UserIndex, 0, "CSOS")
        
        Exit Sub

    End If
    
        
    If UCase$(Left$(rData, 9)) = "/SHOW_SOS" Then
            Dim SSRev  As Long
            Dim SSSuma As Long
            Dim SSName As String
            Dim SSMsg  As String
            Dim SSFH   As String
            rData = Right$(rData, Len(rData) - 10)

            If rData = "" Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||El mensaje SOS no ha sido enviado, revisa el mensaje." & FONTTYPE_INFO)
            Else
                SSName = UserList(UserIndex).Name
                SSMsg = rData
                SSFH = now
                SSRev = val(GetVar(App.Path & "\Logs\Show\SOS\" & SSName & ".ini", "Config", "NumMsg"))
                SSSuma = SSRev + "1"
              
                Call SendData(SendTarget.toindex, UserIndex, 0, "||El mensaje SOS ha sido enviado, ahora solo debes esperar que un gm te responda." _
                        & FONTTYPE_INFO)
                Call SendData(SendTarget.ToAdmins, 0, 0, "||Nuevo SOS del Usuario: " & UserList(UserIndex).Name & FONTTYPE_INFO)
              
                Call WriteVar(App.Path & "\Logs\Show\SOS\" & SSName & ".ini", "Config", "NumMsg", SSSuma)
                Call WriteVar(App.Path & "\Logs\Show\SOS\" & SSName & ".ini", "Mensaje" & SSSuma, "Mensaje", SSMsg)
                Call WriteVar(App.Path & "\Logs\Show\SOS\" & SSName & ".ini", "Mensaje" & SSSuma, "HoraFecha", SSFH)

            End If
        
        End If
       
        If UCase$(Left$(rData, 14)) = "/SHOW_DENUNCIA" Then
            Dim SDRev  As Long
            Dim SDSuma As Long
            Dim SDName As String
            Dim SDMsg  As String
            Dim SDFH   As String
            rData = Right$(rData, Len(rData) - 15)

            If rData = "" Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||El mensaje DENUNCIA no ha sido enviado, revisa el mensaje." & FONTTYPE_INFO)
            Else

                SDName = UserList(UserIndex).Name
                SDMsg = rData
                SDFH = now
                SDRev = val(GetVar(App.Path & "\Logs\Show\DENUNCIA\" & SDName & ".ini", "Config", "NumMsg"))
                SDSuma = SDRev + "1"
              
                Call SendData(SendTarget.toindex, UserIndex, 0, _
                        "||El mensaje DENUNCIA ha sido enviado, ahora solo debes esperar que un gm revise el caso." & FONTTYPE_INFO)
                Call SendData(SendTarget.ToAdmins, 0, 0, "||Nueva DENUNCIA del Usuario: " & UserList(UserIndex).Name & FONTTYPE_INFO)
              
                Call WriteVar(App.Path & "\Logs\Show\DENUNCIA\" & SDName & ".ini", "Config", "NumMsg", SDSuma)
                Call WriteVar(App.Path & "\Logs\Show\DENUNCIA\" & SDName & ".ini", "Mensaje" & SDSuma, "Mensaje", SDMsg)
                Call WriteVar(App.Path & "\Logs\Show\DENUNCIA\" & SDName & ".ini", "Mensaje" & SDSuma, "HoraFecha", SDFH)
           
            End If

        End If
       
        If UCase$(Left$(rData, 9)) = "/SHOW_BUG" Then
            Dim SBRev  As Long
            Dim SBSuma As Long
            Dim SBName As String
            Dim SBMsg  As String
            Dim SBFH   As String
            rData = Right$(rData, Len(rData) - 10)

            If rData = "" Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||El mensaje BUG no ha sido enviado, revisa el mensaje." & FONTTYPE_INFO)
            Else
               
                SBName = UserList(UserIndex).Name
                SBMsg = rData
                SBFH = now
                SBRev = val(GetVar(App.Path & "\Logs\Show\BUG\" & SBName & ".ini", "Config", "NumMsg"))
                SBSuma = SBRev + "1"
              
                Call SendData(SendTarget.toindex, UserIndex, 0, "||El mensaje BUG ha sido enviado, ahora un gm revisar� el bug enviado �Gracias!." _
                        & FONTTYPE_INFO)
                Call SendData(SendTarget.ToAdmins, 0, 0, "||Nuevo BUG del Usuario: " & UserList(UserIndex).Name & FONTTYPE_INFO)
              
                Call WriteVar(App.Path & "\Logs\Show\BUG\" & SBName & ".ini", "Config", "NumMsg", SBSuma)
                Call WriteVar(App.Path & "\Logs\Show\BUG\" & SBName & ".ini", "Mensaje" & SBSuma, "Mensaje", SBMsg)
                Call WriteVar(App.Path & "\Logs\Show\BUG\" & SBName & ".ini", "Mensaje" & SBSuma, "HoraFecha", SBFH)
           
            End If

        End If
        
        If UCase$(Left$(rData, 16)) = "/SHOW_SUGERENCIA" Then
            rData = Right$(rData, Len(rData) - 17)
            Dim SGRev  As Long
            Dim SGSuma As Long
            Dim SGName As String
            Dim SGMsg  As String
            Dim SGFH   As String

            If rData = "" Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||El mensaje SUGERENCIA no ha sido enviado, revisa el mensaje." & FONTTYPE_INFO)
            Else
               
                SGName = UserList(UserIndex).Name
                SGMsg = rData
                SGFH = now
                SGRev = val(GetVar(App.Path & "\Logs\Show\SUGERENCIA\" & SGName & ".ini", "Config", "NumMsg"))
                SGSuma = SGRev + "1"
              
                Call SendData(SendTarget.toindex, UserIndex, 0, _
                        "||El mensaje SUGERENCIA ha sido enviado, el staff debatira su sugerencia �Gracias!." & FONTTYPE_INFO)
                Call SendData(SendTarget.ToAdmins, 0, 0, "|| Nueva SUGERENCIA del Usuario: " & UserList(UserIndex).Name & FONTTYPE_INFO)
              
                Call WriteVar(App.Path & "\Logs\Show\SUGERENCIA\" & SGName & ".ini", "Config", "NumMsg", SGSuma)
                Call WriteVar(App.Path & "\Logs\Show\SUGERENCIA\" & SGName & ".ini", "Mensaje" & SGSuma, "Mensaje", SGMsg)
                Call WriteVar(App.Path & "\Logs\Show\SUGERENCIA\" & SGName & ".ini", "Mensaje" & SGSuma, "HoraFecha", SGFH)
           
            End If

        End If
        
    
    If UCase$(Left$(rData, 9)) = "/GM_QUEST" Then
        rData = Right$(rData, Len(rData) - 9)
         
        If UserList(UserIndex).flags.Quest = 1 Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||Debes esperar a tu turno para que el GM te haga teletransporte." & FONTTYPE_INFO)
            Exit Sub
        Else
            Call SendData(SendTarget.toindex, UserIndex, 0, _
                    "||Has enviado GM QUEST ahora debes esperar a tu turno para que el GM te haga teletransporte." & FONTTYPE_INFO)
            UserList(UserIndex).flags.Quest = 1
            Call Quest.Push(rData, UserList(UserIndex).Name)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||Nuevo QUEST del Usuario: " & UserList(UserIndex).Name & FONTTYPE_INFO)
              
        End If

    End If
    
    Select Case UCase$(Left$(rData, 3))

        Case "/GM"
            rData = Right$(rData, Len(rData) - 4)
        
            Dim GMRev  As Long
            Dim GMSuma As Long
            Dim GMName As String
            Dim GMMsg  As String
            Dim GMFH   As String
                
            If rData = "" Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||El mensaje SOS no ha sido enviado, revisa el mensaje." & FONTTYPE_INFO)
            Else
                GMName = UserList(UserIndex).Name
                GMMsg = rData
                GMFH = now
                GMRev = val(GetVar(App.Path & "\Logs\Consultas\" & GMName & ".ini", "Config", "NumMsg"))
                GMSuma = GMRev + "1"
              
                Call SendData(SendTarget.toindex, UserIndex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que un gm te responda." & _
                        FONTTYPE_INFO)
                Call SendData(SendTarget.ToAdmins, 0, 0, "|| Nuevo SOS del Usuario: " & UserList(UserIndex).Name & FONTTYPE_INFO)
              
                Call WriteVar(App.Path & "\Logs\Consultas\" & GMName & ".ini", "Config", "NumMsg", GMSuma)
                Call WriteVar(App.Path & "\Logs\Consultas\" & GMName & ".ini", "Mensaje" & GMSuma, "Mensaje", GMMsg)
                Call WriteVar(App.Path & "\Logs\Consultas\" & GMName & ".ini", "Mensaje" & GMSuma, "HoraFecha", GMFH)

            End If
        
            Exit Sub

    End Select
    
    Select Case UCase$(Left$(rData, 6))

        Case "/DESC "

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z12" & FONTTYPE_INFO)
                Exit Sub

            End If

            rData = Right$(rData, Len(rData) - 6)

            If Not AsciiValidos(rData) Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||La descripcion tiene caracteres invalidos." & FONTTYPE_INFO)
                Exit Sub

            End If

            UserList(UserIndex).Desc = Trim$(rData)
            Call SendData(SendTarget.toindex, UserIndex, 0, "||La descripcion a cambiado." & FONTTYPE_INFO)
            Exit Sub

        Case "/VOTO "
            rData = Right$(rData, Len(rData) - 6)

            If Not modGuilds.v_UsuarioVota(UserIndex, rData, tStr) Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Voto NO contabilizado: " & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Voto contabilizado." & FONTTYPE_GUILD)

            End If

            Exit Sub

    End Select
    
    Select Case UCase$(Left$(rData, 8))

        Case "/PASSWD "
            rData = Right$(rData, Len(rData) - 8)

            If Len(rData) < 6 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||El password debe tener al menos 6 caracteres." & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toindex, UserIndex, 0, "||El password ha sido cambiado." & FONTTYPE_INFO)
                UserList(UserIndex).Password = MD5String(rData)
                
                #If MYSQL = 1 Then
                   Call Add_DataBase(UserIndex, "Account")
                #End If
                
            End If

            Exit Sub

    End Select
    
    Select Case UCase$(Left$(rData, 9))
        
        Case "/APOSTAR "
            rData = Right(rData, Len(rData) - 9)
            tLong = CLng(val(rData))

            If tLong > 32000 Then tLong = 32000
            n = tLong

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z12")
            ElseIf UserList(UserIndex).flags.TargetNpc = 0 Then
                'Se asegura que el target es un npc
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z30")
            ElseIf Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "Z27")
            ElseIf Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> eNPCType.Timbero Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "�" & "No tengo ningun interes en apostar." & "�" & str(Npclist( _
                        UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            ElseIf n < 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "�" & "El minimo de apuesta es 1 moneda." & "�" & str(Npclist( _
                        UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            ElseIf n > 5000 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "�" & "El maximo de apuesta es 5000 monedas." & "�" & str(Npclist( _
                        UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            ElseIf UserList(UserIndex).Stats.GLD < n Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "�" & "No tienes esa cantidad." & "�" & str(Npclist(UserList( _
                        UserIndex).flags.TargetNpc).char.CharIndex))
            Else

                If RandomNumber(1, 100) <= 47 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + n
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "�" & "Felicidades! Has ganado " & CStr(n) & _
                            " monedas de oro!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                    
                    Apuestas.Perdidas = Apuestas.Perdidas + n
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - n
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "�" & "Lo siento, has perdido " & CStr(n) & " monedas de oro." _
                            & "�" & str(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                
                    Apuestas.Ganancias = Apuestas.Ganancias + n
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))

                End If

                Apuestas.Jugadas = Apuestas.Jugadas + 1
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
                
                Call EnviarOro(UserIndex)

            End If

            Exit Sub

    End Select
    
    Select Case UCase$(Left$(rData, 11))
            
        Case "/CERRARCLAN"

            If Not UserList(UserIndex).GuildIndex >= 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||No perteneces a ning�n clan." & FONTTYPE_GUILD)
                Exit Sub

            End If

            If UCase$(Guilds(UserList(UserIndex).GuildIndex).Fundador) <> UCase$(UserList(UserIndex).Name) Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||No eres l�der del clan." & FONTTYPE_GUILD)
                Exit Sub

            End If

            If Guilds(UserList(UserIndex).GuildIndex).CantidadDeMiembros > 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Debes hechar a todos los miembros del clan para cerrarlo." & FONTTYPE_GUILD)
                Exit Sub

            End If

            'If UserList(UserIndex).flags.YaCerroClan = 1 Then
            'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya has cerrado un clan antes" & FONTTYPE_GUILD)
            'Exit Sub
            'End If

            Call SendData(SendTarget.toall, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " cerr�." & FONTTYPE_GUILD)

            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Founder", "NADIE")
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "GuildName", Guilds(UserList( _
                    UserIndex).GuildIndex).GuildName & "(CLAN CERRADO)")
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex1", "CLAN CERRADO")
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex2", "CLAN CERRADO")
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex3", "CLAN CERRADO")
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex4", "CLAN CERRADO")
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Leader", "NADIE")

            Call Guilds(UserList(UserIndex).GuildIndex).DesConectarMiembro(UserIndex)
            Call Guilds(UserList(UserIndex).GuildIndex).ExpulsarMiembro(UserList(UserIndex).Name)
            UserList(UserIndex).GuildIndex = 0
            'UserList(UserIndex).flags.YaCerroClan = 1
            Call WarpUserChar(UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y)
            Exit Sub
                    
        Case "/FUNDARCLAN"
            rData = Right$(rData, Len(rData) - 11)
            
            If UserList(UserIndex).Clan.FundoClan > 0 Then
                 Call SendData(SendTarget.toindex, UserIndex, 0, "||Ya has fundado un clan, s�lo se puede fundar uno por personaje." & FONTTYPE_INFO)
                Exit Sub
            End If

            If modGuilds.PuedeFundarUnClan(UserIndex, tStr) Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "SHOWFUN")
            Else
                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)

            End If
           
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 12))

        Case "/ECHARPARTY "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)

            If tInt > 0 Then
                Call mdParty.ExpulsarDeParty(UserIndex, tInt)
            Else
                Call SendData(SendTarget.toindex, UserIndex, 0, "|| El personaje no est� online." & FONTTYPE_INFO)

            End If

            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 14))
            
        Case "/MIEMBROSCLAN "
            rData = Trim(Right(rData, Len(rData) - 14))
            Name = Replace(rData, "\", "")
            Name = Replace(rData, "/", "")
    
            If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
                Exit Sub

            End If
            
            tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
            
            For i = 1 To tInt
                tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
                'tstr es la victima
                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & tStr & "<" & rData & ">." & FONTTYPE_INFO)
            Next i
        
            Exit Sub

    End Select
    
     If UCase$(Left$(rData, 10)) = "/PETICION " Then

        rData = Right$(rData, Len(rData) - 10)
        
        Dim Obj As Obj
        
         If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Primero tienes que seleccionar un personaje, hace click izquierdo sobre el." & _
                        FONTTYPE_INFO)
                Exit Sub
         End If
         
            TIndex = NameIndex(rData)
         
          If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Casamiento Then
                
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "Z27")
                    Exit Sub
                End If
                
                If TIndex <= 0 Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(UserIndex).flags.Muerto = 1 Then
                     Call SendData(SendTarget.toindex, UserIndex, 0, "||���Estas muerto!!!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(UserIndex).Stats.ELV <= 20 Then
                   Call SendData(SendTarget.toindex, UserIndex, 0, "||Debes ser por lo menos nivel 20 para casarte." & FONTTYPE_INFO)
                   Exit Sub
                ElseIf UserList(TIndex).Stats.ELV <= 20 Then
                   Call SendData(SendTarget.toindex, UserIndex, 0, "||El otro usuario debe ser por lo menos nivel 20 para casarse." & FONTTYPE_INFO)
                   Exit Sub
                End If
                
                If UserIndex = TIndex Then
                     Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes casarte contigo mismo..." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(UserIndex).Genero = UserList(TIndex).Genero Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes casarte con un usuario de tu mismo g�nero..." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(UserIndex).flags.Casado = 1 Then
                     Call SendData(SendTarget.toindex, UserIndex, 0, "||Ya est�s casado!!!." & FONTTYPE_INFO)
                    Exit Sub
                ElseIf UserList(TIndex).flags.Casado = 1 Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||El usuario ya est� casado!!!." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(UserIndex).flags.Casandose = True And UCase$(UserList(UserIndex).flags.SolicitudC) = UCase$(UserList(UserIndex).Name) Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||Ya has mandado una petici�n a " & UserList(UserIndex).flags.QuienName & " si la quieres cancelar tipea el comando /RECHAZARPETICION." & FONTTYPE_INFO)
                     Exit Sub
                    ElseIf UserList(UserIndex).flags.Casandose = True And UCase$(UserList(UserIndex).flags.SolicitudC) <> UCase$(UserList(TIndex).Name) Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||El usuario ya tiene una petici�n de " & UserList(TIndex).flags.QuienName & " pendiente." & FONTTYPE_INFO)
                    Exit Sub
                    ElseIf UserList(UserIndex).flags.Casandose = True And UCase$(UserList(UserIndex).flags.SolicitudC) <> UCase$(UserList(TIndex).Name) Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "||Este usuario no es el que te envi� solicitud para matrimonio." & FONTTYPE_INFO)
                    Exit Sub
                    ElseIf UserList(UserIndex).flags.Casandose = True And UCase$(UserList(UserIndex).flags.SolicitudC) = UCase$(UserList(TIndex).Name) Then
                      
                      Call SendData(SendTarget.toall, 0, 0, "||" & UserList(UserIndex).Name & " ha aceptado la petici�n de " & UserList(TIndex).Name & " para un matrimonio." & FONTTYPE_TALKMSG)
                      Call SendData(SendTarget.toall, 0, 0, "||El server AoMania declara a " & UserList(UserIndex).Name & " y " & UserList(TIndex).Name & " marido y mujer." & FONTTYPE_TALKMSG)
                      
                      UserList(UserIndex).flags.Casado = 1
                      UserList(TIndex).flags.Casado = 1
                      UserList(UserIndex).Pareja = UserList(TIndex).Name
                      UserList(TIndex).Pareja = UserList(UserIndex).Name
                      
                      Select Case UserList(UserIndex).Genero
                           
                           Case "Hombre"
                              
                              Select Case UCase$(UserList(UserIndex).Raza)
                                    
                                    Case "HOBBIT"
                                      
                                      Obj.ObjIndex = 1646
                                      Obj.Amount = 1
                                      Call MeterItemEnInventario(UserIndex, Obj)
                                      
                                      
                                    Case "ENANO", "GNOMO"
                                    
                                       Obj.ObjIndex = 1645
                                       Obj.Amount = 1
                                       Call MeterItemEnInventario(UserIndex, Obj)
                                    
                                    Case Else
                                       
                                       Obj.ObjIndex = 1647
                                       Obj.Amount = 1
                                       Call MeterItemEnInventario(UserIndex, Obj)
                                        
                              End Select
                              
                              
                           Case "Mujer"
                                
                                Select Case UCase$(UserList(TIndex).Raza)
                                     Case "HOBBIT"
                                       Obj.ObjIndex = 1648
                                       Obj.Amount = 1
                                       Call MeterItemEnInventario(UserIndex, Obj)
                                       
                                     Case "ENANO", "GNOMO"
                                       Obj.ObjIndex = 1649
                                       Obj.Amount = 1
                                       Call MeterItemEnInventario(UserIndex, Obj)
                                       
                                     Case Else
                                        Obj.ObjIndex = 1650
                                        Obj.Amount = 1
                                        Call MeterItemEnInventario(UserIndex, Obj)
                                        
                                End Select
                        
                      End Select
                      
                      Select Case UserList(TIndex).Genero
                           
                           Case "Hombre"
                              
                              Select Case UCase$(UserList(TIndex).Raza)
                                    
                                    Case "HOBBIT"
                                      
                                      Obj.ObjIndex = 1646
                                      Obj.Amount = 1
                                      Call MeterItemEnInventario(TIndex, Obj)
                                      
                                      
                                    Case "ENANO", "GNOMO"
                                    
                                       Obj.ObjIndex = 1645
                                       Obj.Amount = 1
                                       Call MeterItemEnInventario(TIndex, Obj)
                                    
                                    Case Else
                                       
                                       Obj.ObjIndex = 1647
                                       Obj.Amount = 1
                                       Call MeterItemEnInventario(TIndex, Obj)
                                        
                              End Select
                              
                              
                           Case "Mujer"
                                
                                Select Case UCase$(UserList(TIndex).Raza)
                                     Case "HOBBIT"
                                       Obj.ObjIndex = 1648
                                       Obj.Amount = 1
                                       Call MeterItemEnInventario(TIndex, Obj)
                                       
                                     Case "ENANO", "GNOMO"
                                       Obj.ObjIndex = 1649
                                       Obj.Amount = 1
                                       Call MeterItemEnInventario(TIndex, Obj)
                                       
                                     Case Else
                                        Obj.ObjIndex = 1650
                                        Obj.Amount = 1
                                        Call MeterItemEnInventario(TIndex, Obj)
                                        
                                End Select
                        
                      End Select
                      
                    Exit Sub
                End If
                
                  UserList(UserIndex).flags.Casandose = True
                  UserList(TIndex).flags.Casandose = True
                  UserList(UserIndex).flags.Quien = TIndex
                  UserList(TIndex).flags.Quien = UserIndex
                  UserList(UserIndex).flags.QuienName = UserList(TIndex).Name
                  UserList(TIndex).flags.QuienName = UserList(UserIndex).Name
                  UserList(UserIndex).flags.SolicitudC = UserList(UserIndex).Name
                  UserList(TIndex).flags.SolicitudC = UserList(UserIndex).Name
                  
                   Call SendData(SendTarget.toall, 0, 0, "||" & UserList(UserIndex).Name & " le ha propuesto matrimonio a " & UserList(TIndex).Name & ". Ahora se debe esperar la respuesta. Si se quiere cancelar el matrimonio Mandar /RECHAZARPETICION � desconectar.." & FONTTYPE_TALKMSG)
           
        End If
     
        Exit Sub
     End If
     
     If UCase$(Left$(rData, 12)) = "/DIVORCIARSE" Then
        
         If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Primero tienes que seleccionar un personaje, hace click izquierdo sobre el." & _
                        FONTTYPE_INFO)
                Exit Sub
         End If
         
            TIndex = NameIndex(rData)
         
          If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Casamiento Then
          
            If UserList(UserIndex).flags.Casado = 0 Then
                  Call SendData(SendTarget.toindex, UserIndex, 0, "||�� No est�s casado !!" & FONTTYPE_INFO)
                  Exit Sub
            End If
            
            If UserList(UserIndex).Stats.GLD < OroDivorciarse Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Para poder divorciarte es necesario pagar " & OroDivorciarse & " monedas de oro." & FONTTYPE_INFO)
                Exit Sub
            End If
                        
            TIndex = NameIndex(UserList(UserIndex).Pareja)
            
            If TIndex = 0 Then
                 Call SendData(SendTarget.toall, 0, 0, "||���" & UCase$(UserList(UserIndex).Name) & " y " & UCase$(UserList(UserIndex).Pareja) & "  SE DIVORCIARON!!!" & FONTTYPE_AMARILLON)
                 Call SendData(SendTarget.toall, 0, 0, "TW154")
                 Call WriteVar(App.Path & "\charfile\" & UCase$(UserList(UserIndex).Pareja) & ".chr", "INIT", "PAREJA", "")
                 Call WriteVar(App.Path & "\charfile\" & UCase$(UserList(UserIndex).Pareja) & ".chr", "FLAGS", "CASADO", "0")
                 UserList(UserIndex).flags.Casado = 0
                 UserList(UserIndex).Pareja = ""
            Else
                 Call SendData(SendTarget.toall, 0, 0, "||���" & UCase$(UserList(UserIndex).Name) & " y " & UCase$(UserList(TIndex).Name) & "  SE DIVORCIARON!!!" & FONTTYPE_AMARILLON)
                 Call SendData(SendTarget.toall, 0, 0, "TW154")
                 UserList(UserIndex).flags.Casado = 0
                 UserList(TIndex).flags.Casado = 0
                 UserList(UserIndex).Pareja = ""
                 UserList(TIndex).Pareja = ""
            End If
            
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - OroDivorciarse
            Call EnviarOro(UserIndex)
            
            
          End If
        
        Exit Sub
     End If
     
     If UCase$(Left$(rData, 17)) = "/RECHAZARPETICION" Then

        rData = Right$(rData, Len(rData) - 17)
        
        If UserList(UserIndex).flags.Casandose = True Then
            
            If UCase$(UserList(UserIndex).Name) = UCase$(UserList(UserIndex).flags.SolicitudC) Then
            If UCase$(UserList(UserList(UserIndex).flags.Quien).Name) = UCase$(UserList(UserIndex).flags.QuienName) Then
               TIndex = UserList(UserIndex).flags.Quien
           Else
               TIndex = NameIndex(UserList(UserIndex).flags.QuienName)
           End If
           
           If TIndex <= 0 Then
              Call SendData(SendTarget.toindex, UserIndex, 0, "||El otro usuario no est� online." & FONTTYPE_INFO)
              Exit Sub
           End If
           
           Call SendData(SendTarget.toall, 0, 0, "||" & UserList(UserIndex).Name & " cancelo la peticion de " & UserList(UserIndex).flags.QuienName & " para un matrimonio." & FONTTYPE_TALK)
            
           UserList(UserIndex).flags.Casandose = False
           UserList(TIndex).flags.Casandose = False
           UserList(UserIndex).flags.Quien = 0
           UserList(TIndex).flags.Quien = 0
           UserList(UserIndex).flags.QuienName = ""
           UserList(TIndex).flags.QuienName = ""

            Exit Sub
            Else
            If UCase$(UserList(UserList(UserIndex).flags.Quien).Name) = UCase$(UserList(UserIndex).flags.QuienName) Then
               TIndex = UserList(UserIndex).flags.Quien
           Else
               TIndex = NameIndex(UserList(UserIndex).flags.QuienName)
           End If
           
           If TIndex <= 0 Then
              Call SendData(SendTarget.toindex, UserIndex, 0, "||El otro usuario no est� online." & FONTTYPE_INFO)
              Exit Sub
           End If
           
           Call SendData(SendTarget.toall, 0, 0, "||" & UserList(UserIndex).Name & " rechaz� la peticion de " & UserList(UserIndex).flags.QuienName & " para un matrimonio." & FONTTYPE_TALK)
            
           UserList(UserIndex).flags.Casandose = False
           UserList(TIndex).flags.Casandose = False
           UserList(UserIndex).flags.Quien = 0
           UserList(TIndex).flags.Quien = 0
           UserList(UserIndex).flags.QuienName = ""
           UserList(TIndex).flags.QuienName = ""
           End If
         Else
           Call SendData(SendTarget.toindex, UserIndex, 0, "||��No tienes nada que rechazar!!!" & FONTTYPE_INFO)
           Exit Sub
        End If
        
        
        Exit Sub
    End If
     
     Call HandleData_3(UserIndex, rData, Procesado)
     
    Procesado = False
           
End Sub

Public Sub ActGM()
    
    frmMain.Gms.Clear
    
    Dim LoopC     As Integer
    Dim UserIndex As Integer
    
    For LoopC = 1 To LastUser

        'Tiene nombre? Es GM? Si es Dios o Admin, nosotros lo somos tambi�n??
        If (UserList(LoopC).Name <> "") And UserList(LoopC).flags.Privilegios > PlayerType.User And (UserList(LoopC).flags.Privilegios < _
                PlayerType.Dios Or UserList(LoopC).flags.Privilegios >= PlayerType.Dios) Then
            frmMain.Gms.AddItem (UserList(LoopC).Name)

        End If

    Next LoopC
   
End Sub

Public Sub ActUser()
    Dim LoopC As Integer
    frmMain.User.Clear

    For LoopC = 1 To LastUser

        If Len(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.Privilegios <= PlayerType.Consejero Then
            
            frmMain.User.AddItem (UserList(LoopC).Name)

        End If

    Next LoopC

End Sub

Sub MostrarTimeOnline()

    frmMain.CantOnMin.caption = "Minutos Online: " & OnMin
    frmMain.CantOnHor.caption = "Horas Online: " & OnHor
    frmMain.CantOnDay.caption = "Dias Online: " & OnDay

End Sub

Public Sub RegUser()
    Dim LoopC As Integer
    Dim tStr  As String
    Dim Count As Long

    For LoopC = 1 To NumUsers

        If UserList(LoopC).flags.Privilegios = PlayerType.User Then
             
            tStr = UserList(LoopC).Name & "," & tStr
            
            Count = Count + "1"
            
            frmMain.CantUsuarios.caption = "N�mero de usuarios: " & Count

        End If

    Next LoopC
    
    If Len(tStr) = 0 Then
        frmMain.CantUsuarios.caption = "N�mero de usuarios: 0"

    End If

End Sub

Public Sub RegGM()
    Dim LoopC As Integer
    Dim tStr  As String
    Dim Count As Long

    For LoopC = 1 To NumUsers

        If UserList(LoopC).flags.Privilegios > PlayerType.User Then
             
            tStr = UserList(LoopC).Name & "," & tStr
            Count = Count + "1"
            frmMain.CantNumGM.caption = "N�mero de gms: " & Count

        End If

    Next LoopC
    
    If Len(tStr) = 0 Then
        frmMain.CantNumGM.caption = "N�mero de gms: 0"

    End If

End Sub
