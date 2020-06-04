Attribute VB_Name = "TCP_HandleData2"

Option Explicit

Public Sub HandleData_2(ByVal UserIndex As Integer, rData As String, ByRef Procesado As Boolean)

    Dim loopc    As Integer
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

    Procesado = True 'ver al final del sub

    If UCase$(Left$(rData, 9)) = "/REALMSG " Then

        rData = Right$(rData, Len(rData) - 9)

        If UserList(UserIndex).flags.Privilegios > PlayerType.SemiDios Or UserList(UserIndex).flags.PertAlCons = 1 Then

            If rData <> "" Then
                Call SendData(SendTarget.ToRealYRMs, 0, 0, "||" & UserList(UserIndex).Name & ">" & rData & FONTTYPE_CONSEJOVesA)
                Call LogGM(UserList(UserIndex).Name, "Comando: /REALMSG " & rData)

            End If

        End If

        Exit Sub

    End If
    
    If UCase$(Left$(rData, 9)) = "/CAOSMSG " Then

        rData = Right$(rData, Len(rData) - 9)

        If UserList(UserIndex).flags.Privilegios > PlayerType.SemiDios Or UserList(UserIndex).flags.PertAlConsCaos = 1 Then

            If rData <> "" Then
                Call SendData(SendTarget.ToCaosYRMs, 0, 0, "||" & UserList(UserIndex).Name & ">" & rData & FONTTYPE_WETAS)
                Call LogGM(UserList(UserIndex).Name, "Comando: /CAOSMSG " & rData)

            End If

        End If

        Exit Sub

    End If
    
    If UCase$(Left$(rData, 8)) = "/CIUMSG " Then

        rData = Right$(rData, Len(rData) - 8)

        'Solo dioses, admins y RMS
        If UserList(UserIndex).flags.Privilegios > PlayerType.SemiDios Or UserList(UserIndex).flags.PertAlCons = 1 Then

            If rData <> "" Then
                Call SendData(SendTarget.ToCiudadanosYRMs, 0, 0, "||" & UserList(UserIndex).Name & ">" & rData & FONTTYPE_CONSEJOVesA)
                Call LogGM(UserList(UserIndex).Name, "Comando: /CIUMSG " & rData)

            End If

        End If

        Exit Sub

    End If

    '#################### LISTA DE AMIGOS by GALLE ######################
    If UCase$(Left$(rData, 3)) = "/MP" Then
        Dim Mensaje As String
        Dim MPname  As String
        rData = Right$(rData, Len(rData) - 3)
        MPname = ReadField(2, rData, 64)
        Mensaje = ReadField(3, rData, 64)
        TIndex = NameIndex(MPname)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " dice: " & Mensaje & FONTTYPE_TALK)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario recibio el Mensaje." & FONTTYPE_INFO)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 8)) = "/CRIMSG " Then

        rData = Right$(rData, Len(rData) - 8)

        If UserList(UserIndex).flags.Privilegios > PlayerType.SemiDios Or UserList(UserIndex).flags.PertAlConsCaos = 1 Then

            If rData <> "" Then
                Call SendData(SendTarget.ToCriminalesYRMs, 0, 0, "||" & UserList(UserIndex).Name & ">" & rData & FONTTYPE_WETAS)
                Call LogGM(UserList(UserIndex).Name, "Comando: /CRIMSG " & rData)

            End If

        End If

        Exit Sub

    End If

    If UCase$(Left(rData, 3)) = "/SI" Then
        If Encuesta.ACT = 0 Then Exit Sub
        If UserList(UserIndex).flags.VotEnc = True Then Exit Sub
        Encuesta.EncSI = Encuesta.EncSI + 1
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has votado exitosamente." & FONTTYPE_INFO)
        UserList(UserIndex).flags.VotEnc = True
        Exit Sub

    End If

    If UCase$(Left(rData, 3)) = "/NO" Then
        If Encuesta.ACT = 0 Then Exit Sub
        If UserList(UserIndex).flags.VotEnc = True Then Exit Sub
        Encuesta.EncNO = Encuesta.EncNO + 1
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has votado exitosamente." & FONTTYPE_INFO)
        UserList(UserIndex).flags.VotEnc = True
        Exit Sub

    End If

    If UCase$(Left$(rData, 9)) = "/REGALAR " Then
        Dim Cantidad As Long
        Cantidad = UserList(UserIndex).Stats.GLD
        rData = Right$(rData, Len(rData) - 8)
        TIndex = NameIndex(ReadField(1, rData, 32))
        Arg1 = ReadField(2, rData, 32)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub

        End If

        If Distancia(UserList(UserIndex).pos, UserList(TIndex).pos) > 3 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estas Demasiado Lejos" & FONTTYPE_WARNING)
            Exit Sub

        End If

        If val(Arg1) > Cantidad Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No tenes esa cantidad de oro" & FONTTYPE_WARNING)
        ElseIf val(Arg1) < 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No podes transferir cantidades negativas" & FONTTYPE_WARNING)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Le regalaste " & val(Arg1) & " monedas de oro a " & UserList(TIndex).Name & "!" & _
                FONTTYPE_ORO)
            Call SendData(SendTarget.toIndex, TIndex, 0, "||¡" & UserList(UserIndex).Name & " te regalo " & val(Arg1) & " monedas de oro!" & _
                FONTTYPE_ORO)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(Arg1)
            UserList(TIndex).Stats.GLD = UserList(TIndex).Stats.GLD + val(Arg1)
            Call EnviarOro(TIndex)
            Call EnviarOro(UserIndex)
            Exit Sub

        End If

        Exit Sub

    End If
    
    If UCase(Left(rData, 11)) = "/TELEPATIA " Then
        
        rData = Right$(rData, Len(rData) - 11)
      
        tName = ReadField(1, rData, 32)
        tMessage = Right$(rData, Len(rData) - (1 + Len(tName)))
        TIndex = NameIndex(tName)
    
        If UserList(UserIndex).Telepatia = 1 Then
            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
                Exit Sub

            End If

        End If
    
        If UserList(TIndex).flags.Privilegios = PlayerType.User Then

            If UserList(UserIndex).Telepatia = 1 Then
    
                Call SendData(SendTarget.toIndex, TIndex, 0, "||< " & UserList(UserIndex).Name & " > te dice: " & tMessage & FONTTYPE_SERVER)
  
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Le has mandado a " & tName & " : " & tMessage & FONTTYPE_SERVER)
                
                Call LogTelepatia(UserList(UserIndex).Name, tName, tMessage)

            End If

        End If

    End If
    
    Select Case UCase$(rData)
        
        Case "/MOV"

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If
               
            If UserList(UserIndex).flags.TargetUser = 0 Then Exit Sub
               
            If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 0 Then Exit Sub
  
            If Distancia(UserList(UserIndex).pos, UserList(UserList(UserIndex).flags.TargetUser).pos) > 2 Then Exit Sub
  
            Dim CadaverUltPos As WorldPos
            CadaverUltPos.Y = UserList(UserList(UserIndex).flags.TargetUser).pos.Y + 1
            CadaverUltPos.X = UserList(UserList(UserIndex).flags.TargetUser).pos.X
            CadaverUltPos.Map = UserList(UserList(UserIndex).flags.TargetUser).pos.Map
                    
            Dim CadaverUltPos2 As WorldPos
            CadaverUltPos2.Y = UserList(UserList(UserIndex).flags.TargetUser).pos.Y
            CadaverUltPos2.X = UserList(UserList(UserIndex).flags.TargetUser).pos.X + 1
            CadaverUltPos2.Map = UserList(UserList(UserIndex).flags.TargetUser).pos.Map
                    
            Dim CadaverUltPos3 As WorldPos
            CadaverUltPos3.Y = UserList(UserList(UserIndex).flags.TargetUser).pos.Y - 1
            CadaverUltPos3.X = UserList(UserList(UserIndex).flags.TargetUser).pos.X
            CadaverUltPos3.Map = UserList(UserList(UserIndex).flags.TargetUser).pos.Map
                    
            Dim CadaverUltPos4 As WorldPos
            CadaverUltPos4.Y = UserList(UserList(UserIndex).flags.TargetUser).pos.Y
            CadaverUltPos4.X = UserList(UserList(UserIndex).flags.TargetUser).pos.X - 1
            CadaverUltPos4.Map = UserList(UserList(UserIndex).flags.TargetUser).pos.Map
                
            If LegalPos(CadaverUltPos.Map, CadaverUltPos.X, CadaverUltPos.Y, False) Then
                Call WarpUserChar(UserList(UserIndex).flags.TargetUser, CadaverUltPos.Map, CadaverUltPos.X, CadaverUltPos.Y, False)
            ElseIf LegalPos(CadaverUltPos2.Map, CadaverUltPos2.X, CadaverUltPos2.Y, False) Then
                Call WarpUserChar(UserList(UserIndex).flags.TargetUser, CadaverUltPos2.Map, CadaverUltPos2.X, CadaverUltPos2.Y, False)
            ElseIf LegalPos(CadaverUltPos3.Map, CadaverUltPos3.X, CadaverUltPos3.Y, False) Then
                Call WarpUserChar(UserList(UserIndex).flags.TargetUser, CadaverUltPos3.Map, CadaverUltPos3.X, CadaverUltPos3.Y, False)
            ElseIf LegalPos(CadaverUltPos4.Map, CadaverUltPos4.X, CadaverUltPos4.Y, False) Then
                Call WarpUserChar(UserList(UserIndex).flags.TargetUser, CadaverUltPos4.Map, CadaverUltPos4.X, CadaverUltPos4.Y, False)
            Else
                Call WarpUserChar(UserList(UserIndex).flags.TargetUser, 1, 58, 45, True)

            End If

            UserList(UserIndex).flags.TargetUser = 0
            Exit Sub
    
            'Case "/HOGAR"
            '
            '            If UserList(UserIndex).pos.Map = 87 Then
            '                Call SendData(SendTarget.toindex, UserIndex, 0, "||¡¡No puedes usar este comando en retos 2v2!!!!." & FONTTYPE_INFO)
            '                Exit Sub
            '
            '            End If
            '
            ''            If UserList(UserIndex).pos.Map = 66 Then
            '               Call SendData(SendTarget.toindex, UserIndex, 0, "||¡¡En guerras no puedes usar este comando!!!!." & FONTTYPE_INFO)
            '               Exit Sub
            '
            '            End If
            '
            '            If EsNewbie(UserIndex) Then
            '                Call SendData(SendTarget.toindex, UserIndex, 0, "||¡¡Los Newbies no Pueden Utilizar este Comando!!!." & FONTTYPE_INFO)
            '                Exit Sub
            '
            '            End If
            '
            '            If UserList(UserIndex).flags.Muerto = 0 Then
            '                Call SendData(SendTarget.toindex, UserIndex, 0, "||¡¡Tenes que estar muerto para poder usar este comando!!!." & FONTTYPE_INFO)
            '                Exit Sub
            '
            '            End If
            '
            '            If UserList(UserIndex).Counters.Pena >= 1 Then
            '                Call SendData(SendTarget.toindex, UserIndex, 0, "||¡¡No podes usar este comando estando encarcelado!!!." & FONTTYPE_INFO)
            '                Exit Sub
            '
            '            End If
            '
            '            Call WarpUserChar(UserIndex, 34, 30, 50, True)
            '
            '            Exit Sub

        Case "/MAYOR"
            Call CommandMayor(UserIndex)
            Exit Sub
         
        Case "/ONLINE"
        
            'No se envia más la lista completa de usuarios
            n = 0
            tStr = vbNullString
             
            For loopc = 1 To LastUser

                If Len(UserList(loopc).Name) <> 0 And UserList(loopc).flags.Privilegios <= PlayerType.Consejero Then
                    n = n + 1
                    tStr = tStr & UserList(loopc).Name & ", "

                End If

            Next loopc
          
            If n > 0 Then
                tStr = Left$(tStr, Len(tStr) - 2)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & tStr & "." & FONTTYPE_INFO)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Número de usuarios: " & LastUser & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No hay usuarios Online." & FONTTYPE_INFO)

            End If
             
            Exit Sub
            
        Case "/RANKCLAN"
            Call modGuilds.UpdateRankGuild(UserIndex)
            Exit Sub
  

        Case "/CASTILLOS"
            Call SendInfoCastillos(UserIndex)

        Case "/CASTILLO ESTE"
            Call WarpCastillo(UserIndex, "ESTE")
            Exit Sub

        Case "/CASTILLO OESTE"
            Call WarpCastillo(UserIndex, "OESTE")
            Exit Sub

        Case "/CASTILLO NORTE"
            Call WarpCastillo(UserIndex, "NORTE")
            Exit Sub

        Case "/CASTILLO SUR"
            Call WarpCastillo(UserIndex, "SUR")
            Exit Sub

        Case "/FORTALEZA"
            Call WarpCastillo(UserIndex, "FORTALEZA")
            Exit Sub

        Case "/DUELOS"

            If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Primero tienes que seleccionar un personaje, hace click izquierdo sobre el." & _
                    FONTTYPE_INFO)
                Exit Sub

            End If

            Dim JuanpaDuelosMap As Integer
            JuanpaDuelosMap = MAPADUELO
            Dim JuanpaDuelosX As Integer
            JuanpaDuelosX = RandomNumber(43, 58)
            Dim JuanpaDuelosY As Integer
            JuanpaDuelosY = RandomNumber(45, 56)

            '¿El NPC puede comerciar?
           
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Duelos Then
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Navegando = 1 Then
 
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbYellow & "°" & "No puedes entrar a duelos estando navegando!!!" & "°" _
                        & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                    Exit Sub

                End If
            
                If UserList(UserIndex).flags.Privilegios = PlayerType.Dios Or UserList(UserIndex).flags.Privilegios = PlayerType.SemiDios Or _
                    UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbCyan & "°" & _
                        "¡¡No puedes entrar a duelos eres GM teletransportate al mapa " & JuanpaDuelosMap & "!!" & "°" & CStr(Npclist(UserList( _
                        UserIndex).flags.TargetNpc).char.CharIndex))
                    Exit Sub

                End If
            
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡ Estás muerto !!" & FONTTYPE_INFO)
                    Exit Sub
                ElseIf UserList(UserIndex).Stats.ELV < 25 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes hacer duelos siendo menor a nivel 25." & FONTTYPE_INFO)
                    Exit Sub
                ElseIf MapInfo(JuanpaDuelosMap).NumUsers >= 2 Then
                
                    If MapInfo(JuanpaDuelosMap).NumUsers = 2 Then
                
                        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Duelos Then
                            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbYellow & "°" & _
                                "¡¡El mapa de duelos está ocupado ahora mismo.!!" & "°" & CStr(Npclist(UserList( _
                                UserIndex).flags.TargetNpc).char.CharIndex))
                            Exit Sub

                        End If
                
                        Exit Sub

                    End If

                ElseIf MapInfo(UserList(UserIndex).pos.Map).Pk = True Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estas en una zona insegura." & FONTTYPE_WARNING)
                    Exit Sub
                ElseIf MapInfo(JuanpaDuelosMap).NumUsers = 1 Then
                    duelosreta = UserIndex
                 
                    If UserList(UserIndex).flags.Invisible = 1 Then
                        UserList(UserIndex).Counters.Invisibilidad = IntervaloInvisible
                        UserList(UserIndex).flags.Invisible = 0
                        UserList(UserIndex).Counters.Ocultando = 0
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "INVI0")

                    End If
                 
                    If UserList(UserIndex).flags.Oculto = 1 Then
                        UserList(UserIndex).Counters.Ocultando = 0
                        UserList(UserIndex).flags.Oculto = 0
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "INVI0")

                    End If
                 
                    Call WarpUserChar(UserIndex, JuanpaDuelosMap, JuanpaDuelosX, JuanpaDuelosY, True)
       
                    Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos: ¡¡¡" & UserList(duelosreta).Name & " acepto el desafio!!!" & _
                        FONTTYPE_TALK)
                    Exit Sub
                ElseIf MapInfo(JuanpaDuelosMap).NumUsers = 0 Then
                    duelosespera = UserIndex
                
                    If UserList(UserIndex).flags.Oculto = 1 Then
                        UserList(UserIndex).Counters.Ocultando = 0
                        UserList(UserIndex).flags.Oculto = 0
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "INVI0")

                    End If
                
                    If UserList(UserIndex).flags.Invisible = 1 Then
                        UserList(UserIndex).Counters.Invisibilidad = IntervaloInvisible
                        UserList(UserIndex).flags.Invisible = 0
                        UserList(UserIndex).Counters.Ocultando = 0
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "INVI0")

                    End If
                 
                    Call WarpUserChar(UserIndex, JuanpaDuelosMap, JuanpaDuelosX, JuanpaDuelosY, True)
                    Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos: " & UserList(duelosespera).Name & _
                        " espera rival en la sala de torneos." & FONTTYPE_TALK)

                End If
            
            End If

            Exit Sub
        
        Case "/SALIR"

            If UserList(UserIndex).flags.Montado = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes salir estando en montado en tu mascota!." & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).flags.Paralizado = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes salir estando paralizado." & FONTTYPE_WARNING)
                Exit Sub

            End If

            ''mato los comercios seguros
            If UserList(UserIndex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                        Call SendData(SendTarget.toIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & _
                            FONTTYPE_TALK)
                        Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)

                    End If

                End If

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Comercio cancelado. " & FONTTYPE_TALK)
                Call FinComerciarUsu(UserIndex)

            End If

            Call Cerrar_Usuario(UserIndex)
            Exit Sub

        Case "/SALIRCLAN"
            'obtengo el guildindex
            tInt = m_EcharMiembroDeClan(UserIndex, UserList(UserIndex).Name, False)
            
            If tInt > 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Dejas el clan." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(UserIndex).Name & " deja el clan." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu no puedes salir de ningún clan." & FONTTYPE_GUILD)

            End If
            
            Exit Sub
            
        Case "/BALANCE"

            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z30")
                Exit Sub

            End If

            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 3 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
                Exit Sub

            End If

            Select Case Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype

                Case eNPCType.Banquero

                    If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                        CloseSocket (UserIndex)
                        Exit Sub

                    End If

                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & _
                        " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex & FONTTYPE_INFO)

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

                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & _
                            " Ganancia Neta: " & tLong & " (" & n & "%) Jugadas: " & Apuestas.Jugadas & FONTTYPE_INFO)

                    End If

            End Select

            Exit Sub

        Case "/QUIETO" ' << Comando a mascotas

            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z30")
                Exit Sub

            End If

            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
                Exit Sub

            End If

            If Npclist(UserList(UserIndex).flags.TargetNpc).MaestroUser <> UserIndex Then Exit Sub
            Npclist(UserList(UserIndex).flags.TargetNpc).Movement = TipoAI.ESTATICO
            Call Expresar(UserList(UserIndex).flags.TargetNpc, UserIndex)
            Exit Sub

        Case "/ACOMPAÑAR"

            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z30")
                Exit Sub

            End If

            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
                Exit Sub

            End If

            If Npclist(UserList(UserIndex).flags.TargetNpc).MaestroUser <> UserIndex Then Exit Sub
            Call FollowAmo(UserList(UserIndex).flags.TargetNpc)
            Call Expresar(UserList(UserIndex).flags.TargetNpc, UserIndex)
            Exit Sub
            
            '[/Alejo]
        Case "/BARCOS"

            If Barcos.TiempoRest > 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||La embarcación a " & Zonas(Barcos.Zona).nombre & " sarpara en " & _
                    Barcos.TiempoRest & " minutos." & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||La embarcación a " & Zonas(Barcos.Zona).nombre & _
                    " ya a sarpado, llegara a tierra en " & -(TIEMPO_LLEGADA - Barcos.TiempoRest) & " minutos." & FONTTYPE_INFO)

            End If

            Exit Sub
            
        Case "/NAVEGAR"

            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
                Exit Sub

            End If

            '¿El NPC puede comerciar?
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 15 Then
                If Len(Npclist(UserList(UserIndex).flags.TargetNpc).Desc) > 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, UserList( _
                    UserIndex).pos.Map, "||" & vbWhite & "°" & "Yo no administro las navegaciones." & "°" & str(Npclist(UserList( _
                    UserIndex).flags.TargetNpc).char.CharIndex))
                Exit Sub

            End If

            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 4 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                Exit Sub

            End If

            Dim iui       As Integer
            Dim TienePass As Boolean

            If Barcos.Pasajeros >= MAX_PASAJEROS Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & _
                    "°Lo lamento pero el barco esta lleno, deberas esperar hasta la próxima embarcación." & "°" & Npclist(UserList( _
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
                                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°Lo lamento pero la embarcacion ya a partido." & _
                                    "°" & Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex & FONTTYPE_INFO)
                            Else
                                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & _
                                    "°La embarcacion partira en un rato, mientras ve a pasear." & "°" & Npclist(UserList( _
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
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°Tu no tienes pasaje." & "°" & Npclist(UserList( _
                    UserIndex).flags.TargetNpc).char.CharIndex & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°Ese pasaje no es para esta embarcacion." & "°" & Npclist( _
                    UserList(UserIndex).flags.TargetNpc).char.CharIndex & FONTTYPE_INFO)

            End If

            Exit Sub

        Case "/ENTRENAR"

            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z30")
                Exit Sub

            End If

            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
                Exit Sub

            End If

            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> eNPCType.Entrenador Then Exit Sub
            Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNpc)
            Exit Sub
        
        Case "/DESCANSAR"

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            If HayOBJarea(UserList(UserIndex).pos, FOGATA) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "DOK")

                If Not UserList(UserIndex).flags.Descansar Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Te acomodas junto a la fogata y comenzas a descansar." & FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)

                End If

                UserList(UserIndex).flags.Descansar = Not UserList(UserIndex).flags.Descansar
            Else

                If UserList(UserIndex).flags.Descansar Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)
                        
                    UserList(UserIndex).flags.Descansar = False
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "DOK")
                    Exit Sub

                End If

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No hay ninguna fogata junto a la cual descansar." & FONTTYPE_INFO)

            End If

            Exit Sub

        Case "/MEDITAR"

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            If UserList(UserIndex).Stats.MaxMAN = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Solo las clases mágicas conocen el arte de la meditación" & FONTTYPE_INFO)
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
            
            Call SendData(SendTarget.toIndex, UserIndex, 0, "MEDOK")

            If Not UserList(UserIndex).flags.Meditando Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z23")
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z16")

            End If

            UserList(UserIndex).flags.Meditando = Not UserList(UserIndex).flags.Meditando

            'Barrin 3/10/03 Tiempo de inicio al meditar
            If UserList(UserIndex).flags.Meditando Then
                UserList(UserIndex).Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z37")
                
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

        Case "/ACEPTAR"

            If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EstaDueleando1 = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes ir a duelos estando plantes!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EnDosVDos = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya estas en reto!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EsperandoDuelo = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, _
                    "||¡¡Ya has retado antes, espera que acepten tu desafio anterior para poder aceptar uno nuevo.!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserList(UserIndex).flags.TargetUser).flags.EnDosVDos = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserIndex).flags.Muerto = 1 Or UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡No se puede retar muerto!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserList(UserIndex).flags.TargetUser).flags.EstaDueleando1 = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡No se puede retar porque esta en plante!!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserIndex).Stats.GLD < entrarReto Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡Debes tener al meno" & entrarReto & " de oro!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If MapInfo(78).NumUsers >= 2 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya hay un Reto!" & FONTTYPE_TALK)
                Exit Sub

            End If
  
            Call ComensarDuelo(UserIndex, UserList(UserIndex).flags.Oponente)
            'error:     Call SendData(SendTarget.toindex, userindex, 0, "||¡No te han retado!!" & FONTTYPE_TALK)
            Exit Sub

        Case "/ACEPTO"

            If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EnDosVDos = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya estas en reto!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EstaDueleando = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya estas en reto!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EsperandoDuelo1 = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, _
                    "||¡¡Yas has invitado a retar a un usuario. Termina tu reto para poder aceptar uno nuevo!!" & FONTTYPE_TALK)
                Exit Sub

            End If
         
            If UserList(UserIndex).flags.Muerto = 1 Or UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡No se puede retar muerto!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserList(UserIndex).flags.TargetUser).flags.EstaDueleando = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserList(UserIndex).flags.TargetUser).flags.EnDosVDos = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserIndex).Stats.GLD < entrarPlante Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡Debes tener al menos " & entrarPlante & " de oro!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If YaHayPlante = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya hay un Reto!" & FONTTYPE_TALK)
                Exit Sub

            End If
 
            Call ComensarDueloPlantes(UserIndex, UserList(UserIndex).flags.Oponente1)
            'error:     Call SendData(SendTarget.toindex, userindex, 0, "||¡No te han retado!!" & FONTTYPE_TALK)
            Exit Sub

            'RETOS 2V2

        Case "/DUAL"

            If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If Team.EnCurso = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los 2vs2 Están ocupados!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserIndex = Team.Pj1 Or UserIndex = Team.Pj2 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Ya estas emparejado :$!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            Dim Pj2 As Integer
            Pj2 = UserList(UserIndex).flags.TargetUser

            If Team.Activado = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los Retos 2 v 2 están desactivados!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.TargetUser = Team.Pj1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Ya tiene pareja!!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.TargetUser = Team.Pj2 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Ya tiene pareja!!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserList(UserIndex).flags.TargetUser).Clase = UserList(UserIndex).Clase Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes conformar pareja con alguien de tu misma clase!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EstaDueleando1 = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes ir a retos estando plantes!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EstaDueleando = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes ir a duelos estando retos!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).Stats.GLD < entrarReto2v2 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Debes tener almenos " & entrarReto2v2 & " para poder retar." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserList(UserIndex).flags.TargetUser).flags.EstaDueleando = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserList(UserIndex).flags.TargetUser).flags.EstaDueleando1 = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
                Exit Sub

            End If
                 
            If UserList(UserIndex).flags.TargetUser = UserIndex Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Debes seleccionar a un personaje!." & FONTTYPE_INFO)
                Exit Sub

            End If
    
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario esta muerto" & FONTTYPE_INFO)
                Exit Sub

            End If
 
            If UserList(UserIndex).flags.TargetUser <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Debes seleccionar a un usuario." & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(Pj2).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario esta muerto!." & FONTTYPE_INFO)
                Exit Sub

            End If
   
            If Distancia(UserList(UserList(UserIndex).flags.TargetUser).pos, UserList(UserIndex).pos) > 5 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estás demasiado lejos!" & FONTTYPE_INFO)
                Exit Sub

            End If
    
            If Team.EnCurso = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los 2vs2 Están ocupados!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If Team.Activado = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los retos 2vs2 están desactivados!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EnDosVDos = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Ya estás en 2vs2!!!" & FONTTYPE_INFO)

            End If
        
            Call SendData(SendTarget.toIndex, Pj2, 0, "||" & UserList(UserIndex).Name & _
                " desea jugar un 2vs2. Haz click sobre tu pareja y escribe /SDUAL para aceptar." & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Pediste a " & UserList(Pj2).Name & " que sea tu pareja." & FONTTYPE_INFO)
            UserList(UserIndex).flags.envioSol = True
            UserList(Pj2).flags.RecibioSol = True
            UserList(Pj2).flags.compa = UserIndex
        
            Exit Sub
        
        Case "/SDUAL"

            If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes aceptar si estas muerto!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If Team.EnCurso = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los 2vs2 Están ocupados!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If Team.Activado = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los Retos 2 v 2 están desactivados!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserIndex = Team.Pj1 Or UserIndex = Team.Pj2 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Ya estas emparejado :$!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserList(UserIndex).flags.TargetUser).Clase = UserList(UserIndex).Clase Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes conformar pareja con alguien de tu misma clase!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.TargetUser = Team.Pj1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Ya tiene pareja!!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.TargetUser = Team.Pj2 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Ya tiene pareja!!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EstaDueleando1 = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes ir a duelos estando plantes!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EstaDueleando = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes ir a duelos estando retos!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).Stats.GLD < entrarReto2v2 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Debes tener almenos " & entrarReto & " para poder retar." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserList(UserIndex).flags.TargetUser).flags.EstaDueleando = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserList(UserIndex).flags.TargetUser).flags.EstaDueleando1 = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Esta muerto!")
                Exit Sub

            End If

            If Team.EnCurso = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los 2vs2 Están ocupados!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If Team.Activado = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los retos 2vs2 están desactivados!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).flags.RecibioSol = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Nadie te invitó a como pareja!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EnDosVDos = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Ya estás en 2vs2!!!" & FONTTYPE_INFO)

            End If

            If Team.SonDos = True Then
                Team.pj3 = UserIndex
                Team.pj4 = UserList(UserIndex).flags.compa
                'Warpeo
                Call WarpUserChar(Team.Pj1, 87, 41, 50)
                Call WarpUserChar(Team.Pj2, 87, 41, 51)
                Call WarpUserChar(Team.pj3, 87, 60, 50)
                Call WarpUserChar(Team.pj4, 87, 60, 51)
                UserList(Team.Pj1).flags.EnDosVDos = True
                UserList(Team.Pj2).flags.EnDosVDos = True
                UserList(Team.pj3).flags.EnDosVDos = True
                UserList(Team.pj4).flags.EnDosVDos = True
                Team.EnCurso = True
                Call SendData(toall, UserIndex, 0, "||2Vs2: " & UserList(Team.Pj1).Name & " y " & UserList(Team.Pj2).Name & " VS " & UserList( _
                    Team.pj3).Name & " y " & UserList(Team.pj4).Name & " que gane el mejor!" & FONTTYPE_RETOS2V2)
           
            ElseIf Team.SonDos = False Then
                Team.SonDos = True
                Team.Pj1 = UserIndex
                Team.Pj2 = UserList(UserIndex).flags.compa
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu pareja es ahora " & UserList(Team.Pj2).Name & " , espera contrincantes." & _
                    FONTTYPE_INFO)
                Call SendData(SendTarget.toIndex, Team.Pj2, 0, "||Tu pareja es ahora " & UserList(UserIndex).Name & " , espera contrincantes." & _
                    FONTTYPE_INFO)
                Call SendData(SendTarget.toall, 0, 0, "||2Vs2: La pareja " & UserList(UserIndex).Name & "(" & UserList(UserIndex).Clase & ")" & _
                    " y " & UserList(Team.Pj2).Name & "(" & UserList(Team.Pj2).Clase & ")" & " Retan por 1KK !!." & FONTTYPE_RETOS2V2)

            End If

            Exit Sub
            
            'TERMINA RETOS 2V2
        Case "/RETAR"

            If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EstaDueleando1 = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes ir a retos estando plantes!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserIndex = Team.Pj1 Or UserIndex = Team.Pj2 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).flags.TargetUser = Team.Pj1 Or UserList(UserIndex).flags.TargetUser = Team.Pj2 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puede participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EnDosVDos = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya estas en reto!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡Estas Muerto!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserIndex).flags.TargetUser > 0 Then
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡El usuario con el que quieres retar está muerto!" & FONTTYPE_TALK)
                    Exit Sub

                End If

                If UserList(UserList(UserIndex).flags.TargetUser).flags.EstaDueleando = True Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
                    Exit Sub

                End If

                If UserList(UserList(UserIndex).flags.TargetUser).flags.EstaDueleando1 = True Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
                    Exit Sub

                End If

                If UserList(UserList(UserIndex).flags.TargetUser).flags.EnDosVDos = True Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
                    Exit Sub

                End If
    
                If UserList(UserIndex).Stats.GLD < entrarReto Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡Debes tener al menos " & entrarReto & " de oro!!" & FONTTYPE_TALK)
                    Exit Sub

                End If

                If MapInfo(78).NumUsers >= 2 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya hay un reto!." & FONTTYPE_TALK)
                    Exit Sub

                End If

                If UserList(UserIndex).flags.TargetUser = UserIndex Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes retarte a ti mismo." & FONTTYPE_TALK)
                    Exit Sub

                End If

                If UserList(UserList(UserIndex).flags.TargetUser).flags.EsperandoDuelo = True Then
                    If UserList(UserList(UserIndex).flags.TargetUser).flags.Oponente = UserIndex Then
                        Call ComensarDuelo(UserIndex, UserList(UserIndex).flags.TargetUser)
                        Exit Sub

                    End If

                    If UserList(UserList(UserIndex).flags.TargetUser).flags.EsperandoDuelo = True Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, _
                            "||El usuario que intentas retar ya ha retado a otro usuario, espera que termine su reto!." & FONTTYPE_TALK)
                        Exit Sub

                    End If

                Else
                    Call SendData(SendTarget.toIndex, UserList(UserIndex).flags.TargetUser, 0, "|| " & UserList(UserIndex).Name & _
                        " Te ha retado por " & entrarReto & " , si quieres aceptar haz click sobre tu oponente y pon /ACEPTAR." & FONTTYPE_TALK)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has retado por " & entrarReto & " a " & UserList(UserList( _
                        UserIndex).flags.TargetUser).Name & FONTTYPE_TALK)
                    UserList(UserIndex).flags.EsperandoDuelo = True
                    UserList(UserIndex).flags.Oponente = UserList(UserIndex).flags.TargetUser
                    UserList(UserList(UserIndex).flags.TargetUser).flags.Oponente = UserIndex
                    Exit Sub

                End If

            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_TALK)

            End If

            Exit Sub
    
        Case "/PLANTAR"

            If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EnDosVDos = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya estas en reto!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EstaDueleando = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya estas en reto!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserIndex = Team.Pj1 Or UserIndex = Team.Pj2 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).flags.TargetUser = Team.Pj1 Or UserList(UserIndex).flags.TargetUser = Team.Pj2 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puede participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡Estas Muerto!!" & FONTTYPE_TALK)
                Exit Sub

            End If

            If UserList(UserIndex).flags.TargetUser > 0 Then
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡El usuario con el que quieres retar está muerto!" & FONTTYPE_TALK)
                    Exit Sub

                End If

                If UserList(UserList(UserIndex).flags.TargetUser).flags.EnDosVDos = True Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
                    Exit Sub

                End If

                If UserList(UserList(UserIndex).flags.TargetUser).flags.EstaDueleando1 = True Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya hay un reto!" & FONTTYPE_TALK)
                    Exit Sub

                End If

                If UserList(UserList(UserIndex).flags.TargetUser).flags.EstaDueleando = True Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
                    Exit Sub

                End If

                If UserList(UserIndex).Stats.GLD < entrarPlante Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡Debes tener al menos " & entrarPlante & " de oro!!" & FONTTYPE_TALK)
                    Exit Sub

                End If

                If YaHayPlante = True Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡¡'Ya hay un reto!!!!" & FONTTYPE_TALK)
                    Exit Sub

                End If

                If UserList(UserIndex).flags.TargetUser = UserIndex Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes retarte a ti mismo." & FONTTYPE_TALK)
                    Exit Sub

                End If

                If UserList(UserList(UserIndex).flags.TargetUser).flags.EsperandoDuelo1 = True Then
                    If UserList(UserList(UserIndex).flags.TargetUser).flags.Oponente1 = UserIndex Then
                        Call ComensarDueloPlantes(UserIndex, UserList(UserIndex).flags.TargetUser)
                        Exit Sub

                    End If

                    If UserList(UserList(UserIndex).flags.TargetUser).flags.EsperandoDuelo1 = True Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, _
                            "||El usuario que intentas retar ya ha retado a otro usuario, espera que termine su reto!." & FONTTYPE_TALK)
                        Exit Sub

                    End If

                Else
                    Call SendData(SendTarget.toIndex, UserList(UserIndex).flags.TargetUser, 0, "|| " & UserList(UserIndex).Name & _
                        " Te ha retado a Plantar por " & entrarPlante & " , si quieres aceptar haz click sobre tu oponente y pon /ACEPTO." & _
                        FONTTYPE_TALK)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has retado a Plantar por " & entrarPlante & " a " & UserList(UserList( _
                        UserIndex).flags.TargetUser).Name & FONTTYPE_TALK)
                    UserList(UserIndex).flags.EsperandoDuelo1 = True
                    UserList(UserIndex).flags.Oponente1 = UserList(UserIndex).flags.TargetUser
                    UserList(UserList(UserIndex).flags.TargetUser).flags.Oponente1 = UserIndex
                    Exit Sub

                End If

            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_TALK)

            End If

            Exit Sub

        Case "/GANADOR"

            If UserList(UserIndex).flags.death = True Then
                If terminodeat = True Then
                    Call WarpUserChar(UserIndex, 1, 50, 50, True)
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + 1000000
                    UserList(UserIndex).Stats.PuntosDeath = UserList(UserIndex).Stats.PuntosDeath + 1
                    Call SendUserStatsBox(UserIndex)
                    Call SendData(toall, UserIndex, 0, "||GANADOR DEATHMATCH: " & UserList(UserIndex).Name & FONTTYPE_DEATH)
                    Call SendData(toall, UserIndex, 0, "||PREMIO: 1.000.000, Equipo Recaudado y 1 punto de DeathMatch." & FONTTYPE_DEATH)
                    UserList(UserIndex).flags.death = False
                    terminodeat = False
                    deathesp = False
                    deathac = False
                    Cantidad = 0

                End If

            End If

            Exit Sub
   
        Case "/VERS"
            Call EnviarResp(UserIndex)
            SendData SendTarget.toIndex, UserIndex, 0, "INITRES"
            Exit Sub

        Case "/RESETSOP"
            Call ResetSop(UserIndex)
            Exit Sub

            ' Case "/XAOPEPELVL"
            'If UserList(userindex).Stats.ELV = 55 Then
            'Exit Sub
            'End If
            'Dim lvl As Integer
            ' For lvl = 1 To 55
            ' UserList(userindex).Stats.Exp = UserList(userindex).Stats.ELU
            'Call CheckUserLevel(userindex)
            ' Call SendData(SendTarget.toindex, userindex, 0, "||Has Subido un nivel!" & FONTTYPE_APU)
            ' Next
            'Exit Sub
        
            'Case "/XAOPEPEORO"
            'If UserList(userindex).Stats.GLD >= 50000000 Then Exit Sub
            'UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + 50000000
            'Call SendUserStatsBox(userindex)
            'Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado 50.000.000 monedas de ORO!" & FONTTYPE_ORO)
            'Exit Sub
        Case "/PARTICIPAR"

            If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
                Exit Sub

            End If
            
            If UserList(UserIndex).flags.Invisible = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes ir a eventos estando invisible!." & FONTTYPE_WARNING)
                Exit Sub

            End If
      
            If UserList(UserIndex).flags.Oculto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes ir a eventos estando invisible!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EstaDueleando1 = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes ir a torneos estando plantes!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estas muerto!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserIndex = Team.Pj1 Or UserIndex = Team.Pj2 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
                Exit Sub

            End If
            
            If UserList(UserIndex).Stats.ELV < lvlTorneo Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Debes ser lvl " & lvlTorneo & " o mas para entrar al torneo!" & FONTTYPE_INFO)
                Exit Sub

            End If
       
            Call Torneos_Entra(UserIndex)
            Exit Sub

        Case "/DEATH"

            If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.Invisible = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes ir a eventos estando invisible!." & FONTTYPE_WARNING)
                Exit Sub

            End If
      
            If UserList(UserIndex).flags.Oculto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes ir a eventos estando invisible!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EstaDueleando1 = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes ir a deathmatch estando plantes!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserIndex = Team.Pj1 Or UserIndex = Team.Pj2 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estas muerto!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).Stats.ELV < lvlDeath Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Debes ser lvl " & lvlDeath & " o mas para entrar al deathmatch!" & FONTTYPE_INFO)
                Exit Sub

            End If
       
            Call death_entra(UserIndex)
            Exit Sub

        Case "/GUERRA"
        
            If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, _
                    "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
                Exit Sub

            End If
       
            Call CommandGuerra(UserIndex)
            Exit Sub
            
        Case "/MEDUSA"
        
            If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, _
                    "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
                Exit Sub

            End If
       
            Call CommandMedusa(UserIndex)
            Exit Sub

        Case "/PUNTOS"
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Actualmente Tienes:" & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Puntos de Torneo: " & UserList(UserIndex).Stats.PuntosTorneo & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Puntos de Deathmatch: " & UserList(UserIndex).Stats.PuntosDeath & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Puntos de Retos : " & UserList(UserIndex).Stats.PuntosRetos & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Puntos de Duelos: " & UserList(UserIndex).Stats.PuntosDuelos & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Puntos de Plantes: " & UserList(UserIndex).Stats.PuntosPlante & FONTTYPE_INFO)
            Exit Sub

        Case "/TIEMPOS"
            Dim tiempo1   As Integer
            Dim tiempo2   As Integer
            Dim tiempo3   As Integer
            Dim demonioql As Integer
            Dim arcangel  As Integer
            Dim torneoql  As Integer
            Dim mascotaql As Integer
            Dim deatmaql  As Integer
            Dim GAOMania  As Integer
            GAOMania = 48
            deatmaql = 63
            mascotaql = 480 'mascota
            tiempo1 = 360 ' demonio
            tiempo2 = 380 ' arcangel
            tiempo3 = 94 ' torneo
            GAOMania = val(GAOMania) - val(bandasqls)
            demonioql = val(tiempo1) - val(ContReSpawnNpc)
            arcangel = val(tiempo2) - val(ContReSpawnNpc)
            torneoql = val(tiempo3) - val(xao)
            mascotaql = val(mascotaql) - val(mariano)
            deatmaql = val(deatmaql) - val(tukiql)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Quedan " & demonioql & " minutos para que renasca el Espectro Infernal!." & _
                FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Quedan " & arcangel & " minutos para que renasca el Arcangel!." & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Quedan " & mascotaql & " minutos para que renasca el Domador!." & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Quedan " & torneoql & " minutos para el próximo torneo automatico!." & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Quedan " & deatmaql & " minutos para el próximo deathmatch automatico!." & _
                FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Quedan " & GAOMania & " minutos para la próxima Guerra AOMania!." & FONTTYPE_INFO)
            Exit Sub

        Case "/MEZCLAR"
            Dim AlasItems(2) As Integer
            AlasItems(0) = 4023
            AlasItems(1) = 4024
            AlasItems(2) = 4025

            Dim AlasLvl1(1) As Integer
            AlasLvl1(0) = 4301 'Ciudadano
            AlasLvl1(1) = 4302 'Criminal

            Dim AlasLvl2(1) As Integer
            AlasLvl2(0) = 4305 'Ciudadano
            AlasLvl2(1) = 4306 'Criminal

            Dim AlasLvl3(1) As Integer
            AlasLvl3(0) = 4309 'Ciudadano
            AlasLvl3(1) = 4310 'Criminal

            Dim AlasLvl4(1) As Integer
            AlasLvl4(0) = 4313 'Ciudadano
            AlasLvl4(1) = 4314 'Criminal

            Dim HasObjects As Boolean
            Dim H          As Long
            HasObjects = True

            For H = 0 To UBound(AlasItems)

                If Not TieneObjetos(AlasItems(H), 1, UserIndex) Then
                    HasObjects = False
                    Exit For

                End If

            Next H

            If HasObjects Then

                For H = 0 To UBound(AlasItems)
                    Call QuitarObjetos(AlasItems(H), 1, UserIndex)
                Next H

                Dim NoFallaAlas As Boolean
                NoFallaAlas = RandomNumber(1, 3) = 2
                Dim MiObj As Obj
                MiObj.Amount = 1
                Dim alasQuitar As Integer
                'Nunca intente :$ ahora lo hago, a mi esa mierda me da sospecha a lentitud jaja pero mariano quiere cada mierda

                If TieneObjetos(AlasLvl4(0), 1, UserIndex) Then
                    Exit Sub
                ElseIf TieneObjetos(AlasLvl4(1), 1, UserIndex) Then
                    Exit Sub
                ElseIf TieneObjetos(AlasLvl3(0), 1, UserIndex) Then
                    MiObj.ObjIndex = AlasLvl4(0)
                    alasQuitar = AlasLvl3(0)
                ElseIf TieneObjetos(AlasLvl3(1), 1, UserIndex) Then
                    MiObj.ObjIndex = AlasLvl4(1)
                    alasQuitar = AlasLvl3(1)
                ElseIf TieneObjetos(AlasLvl2(0), 1, UserIndex) Then
                    MiObj.ObjIndex = AlasLvl3(0)
                    alasQuitar = AlasLvl2(0)
                ElseIf TieneObjetos(AlasLvl2(1), 1, UserIndex) Then
                    MiObj.ObjIndex = AlasLvl3(1)
                    alasQuitar = AlasLvl2(1)
                ElseIf TieneObjetos(AlasLvl1(0), 1, UserIndex) Then
                    MiObj.ObjIndex = AlasLvl2(0)
                    alasQuitar = AlasLvl1(0)
                ElseIf TieneObjetos(AlasLvl1(1), 1, UserIndex) Then
                    MiObj.ObjIndex = AlasLvl2(1)
                    alasQuitar = AlasLvl1(1)
                Else

                    If Criminal(UserIndex) Then
                        MiObj.ObjIndex = AlasLvl1(1)
                    Else
                        MiObj.ObjIndex = AlasLvl1(0)

                    End If

                End If

                If NoFallaAlas Then
                    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                        Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)

                    End If

                    Call QuitarObjetos(alasQuitar, 1, UserIndex)
                    Call SendData(SendTarget.toall, 0, 0, "||El usuario " & UserList(UserIndex).Name & " ha creado """ & ObjData( _
                        MiObj.ObjIndex).Name & """ exitosamente. ~255~255~255~1~0~")
                    Call SendData(SendTarget.toall, UserIndex, UserList(UserIndex).pos.Map, "TW122")
                    Call Alas(UserList(UserIndex).Name & " ha creado alas")
                Else
                    Call SendData(SendTarget.toall, 0, 0, "||El usuario " & UserList(UserIndex).Name & " ha fallado en crear """ & ObjData( _
                        MiObj.ObjIndex).Name & """ y ha perdido los items de la mezcla. ~255~255~255~1~0~")
                    Call SendData(SendTarget.toall, UserIndex, UserList(UserIndex).pos.Map, "TW45")

                End If

            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Para realizar una mezcla necesitas """ & ObjData(AlasItems(0)).Name & """, """ & _
                    ObjData(AlasItems(1)).Name & """ y """ & ObjData(AlasItems(2)).Name & "~255~255~255~0~0~")

            End If

            Exit Sub

            'Case "/XAOPEPESKILLS"
            'Dim satu  As Integer
            ' For satu = 1 To NUMSKILLS
            '         UserList(userindex).Stats.UserSkills(satu) = 100
            '  Next
            ' Call SendData(SendTarget.toindex, userindex, 0, "||Tienes todos tus skills al maximo" & FONTTYPE_ORO)
            ' Exit Sub
        
        Case "/PROMEDIO"
            Dim Promedio
            Promedio = Round(UserList(UserIndex).Stats.MaxHP / UserList(UserIndex).Stats.ELV, 2)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El Promedio de vida de tu Personaje es de " & Promedio & FONTTYPE_ORO)
            Exit Sub
        
            ' FIANZA CULIA
        Case "/SDFAGASSATUROS" ' CHOTS | Sistema de Fianzas
            Dim fianza As Double

            If UserList(UserIndex).Counters.Pena = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No estas en la carcel, o tienes pena permanente!." & FONTTYPE_INFO)
                Exit Sub

            End If

            fianza = val((UserList(UserIndex).Counters.Pena) * 20000) 'CHOTS | 200k por minuto asi le re kb

            If UserList(UserIndex).Stats.GLD < fianza Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Necesitas " & fianza & " monedas de oro!." & FONTTYPE_INFO)
                Exit Sub

            End If

            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(fianza)
            Call EnviarOro(UserIndex)
            UserList(UserIndex).Counters.Pena = 0
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has sido liberado bajo fianza!" & FONTTYPE_INFO)
            Call WarpUserChar(UserIndex, Libertad.Map, Libertad.X, Libertad.Y, True)

            Exit Sub 'CHOTS | Sistema de Fianzas
               
        Case "/COLADESHURA11"

            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z30")
                Exit Sub

            End If

            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> eNPCType.Revividor Or UserList(UserIndex).flags.Muerto <> 1 Then Exit Sub

            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 10 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
                Exit Sub

            End If

            Call RevivirUsuario(UserIndex)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "Z40")
   
            Exit Sub

        Case "/SEMANTICOZ23"

            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z30")
                Exit Sub

            End If

            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> eNPCType.Revividor Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub

            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 10 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z32")
                Exit Sub

            End If

            If UserList(UserIndex).flags.Envenenado = 1 Then
                UserList(UserIndex).flags.Envenenado = 0
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Te has curado el envenenamiento!" & FONTTYPE_INFO)
         
            End If

            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
          
            Call EnviarHP(UserIndex)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "Z41")
            Exit Sub
   
        Case "/AYUDA"
            Call SendHelp(UserIndex)
            Exit Sub
                  
        Case "/EST"
            Call SendUserStatsTxt(UserIndex, UserIndex)
            Exit Sub
        
            'Proponer casamiento, tarararan
        Case "/PROPONER"

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            'Es un usuario el click (?
            If UserList(UserIndex).flags.TargetUser > 0 Then
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡No puedes casarte con un muerto!" & FONTTYPE_INFO)
                    Exit Sub

                End If

                'soy yo ?
                If UserList(UserIndex).flags.TargetUser = UserIndex Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes casarte contigo mismo..." & FONTTYPE_INFO)
                    Exit Sub

                End If

                'somos gays? ?
                If UserList(UserList(UserIndex).flags.TargetUser).Genero = UserList(UserIndex).Genero Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes con personas de tu mismo sexo." & FONTTYPE_INFO)
                    Exit Sub

                End If

                'ta muy lejos ?
                If Distancia(UserList(UserList(UserIndex).flags.TargetUser).pos, UserList(UserIndex).pos) > 3 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estas demasiado lejos del usuario." & FONTTYPE_INFO)
                    Exit Sub

                End If

                'Actual? Jaja ?
                If UserList(UserList(UserIndex).flags.TargetUser).Pareja = UserList(UserIndex).Name Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario es tu pareja actual. ¿No quieres más a tu pareja? /DIVORCIARSE." _
                        & FONTTYPE_INFO)
                    Exit Sub

                End If

                'Tiene pareja ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Casado = 1 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario ya tiene pareja." & FONTTYPE_INFO)
                    Exit Sub

                End If

                'Ya ta casandote ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Casandose = True And UserList(UserList( _
                    UserIndex).flags.TargetUser).flags.Quien <> UserIndex Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario esta intentado casarse ..." & FONTTYPE_INFO)
                    Exit Sub

                End If
               
                UserList(UserIndex).flags.Quien = UserList(UserIndex).flags.TargetUser
               
                'Si ambos pusieron /PROPONER entonces
                If UserList(UserIndex).flags.Quien = UserList(UserIndex).flags.TargetUser And UserList(UserList( _
                    UserIndex).flags.TargetUser).flags.Quien = UserIndex Then
                   
                    UserList(UserIndex).flags.Casado = 1
                    UserList(UserIndex).Pareja = UserList(UserList(UserIndex).flags.TargetUser).Name
                   
                    UserList(UserList(UserIndex).flags.TargetUser).flags.Casado = 1
                    UserList(UserList(UserIndex).flags.TargetUser).Pareja = UserList(UserIndex).Name
                   
                    'SE CASARON
                    Call SendData(SendTarget.toall, 0, 0, "||Se han unido en matrimonio " & UserList(UserIndex).Name & UserList(UserList( _
                        UserIndex).flags.TargetUser).Name & "  ¡Felicidades! " & FONTTYPE_GUILD & ENDC)
                Else
                    UserList(UserIndex).flags.Casandose = True
                    UserList(UserList(UserIndex).flags.TargetUser).flags.Casandose = True
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Le has propuesto casamiento a " & UserList(UserList( _
                        UserIndex).flags.TargetUser).Name & " espera su respuesta..." & FONTTYPE_TALK)
                    Call SendData(SendTarget.toIndex, UserList(UserIndex).flags.TargetUser, 0, "||" & UserList(UserIndex).Name & _
                        " te ha propuesto matrimonio. Si deseas aceptar, Escribe /PROPONER, caso contrario /NEGAR." & FONTTYPE_TALK)
                    UserList(UserList(UserIndex).flags.TargetUser).flags.TargetUser = UserIndex

                End If

            End If

            Exit Sub
   
        Case "/NEGAR"

            'Es un usuario el click (?
            If UserList(UserIndex).flags.TargetUser > 0 Then
                If UserList(UserIndex).flags.Quien = UserList(UserIndex).flags.TargetUser Then
                    UserList(UserIndex).flags.Casandose = False
                    UserList(UserList(UserIndex).flags.TargetUser).flags.Casandose = True
                    Call SendData(SendTarget.toIndex, UserList(UserIndex).flags.TargetUser, 0, "||" & UserList(UserIndex).Name & _
                        " no desea unirse en matrimonio." & FONTTYPE_TALK)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Le niegas la propuesta de casamiento a " & UserList(UserList( _
                        UserIndex).flags.TargetUser).Name & ". (Rompecorazones :@ )" & FONTTYPE_TALK)

                End If

            End If

            Exit Sub
        
        Case "/SEG"

            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "OFFOFS")
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "ONONS")

            End If

            UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
            Exit Sub
            
        Case "/SEGCLAN"

            If UserList(UserIndex).flags.SeguroClan Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "SEGCO99")
                UserList(UserIndex).flags.SeguroClan = False
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "SEG108")
                UserList(UserIndex).flags.SeguroClan = True

            End If

            Exit Sub
            
        Case "/SEGCMBT"

            If UserList(UserIndex).flags.SeguroCombate Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "SEG11")
                UserList(UserIndex).flags.SeguroCombate = False
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "SEG10")
                UserList(UserIndex).flags.SeguroCombate = True

            End If

            Exit Sub
            
        Case "/SEGOBJT"

            If UserList(UserIndex).flags.SeguroObjetos Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "SEG13")
                UserList(UserIndex).flags.SeguroObjetos = False
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "SEG12")
                UserList(UserIndex).flags.SeguroObjetos = True

            End If

            Exit Sub
            
        Case "/SEGHZS"

            If UserList(UserIndex).flags.SeguroHechizos Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "SEG15")
                UserList(UserIndex).flags.SeguroHechizos = False
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "SEG14")
                UserList(UserIndex).flags.SeguroHechizos = True

            End If

            Exit Sub
         
        Case "/COMERCIAR"

            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            If UserList(UserIndex).flags.Montado = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Debes Demontarte para poder Comerciar!.!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).flags.Comerciando Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Ya estás comerciando" & FONTTYPE_INFO)
                Exit Sub

            End If
            
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                Exit Sub

            End If

            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNpc > 0 Then

                '¿El NPC puede comerciar?
                If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                    If Len(Npclist(UserList(UserIndex).flags.TargetNpc).Desc) > 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, UserList( _
                        UserIndex).pos.Map, "||" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & CStr(Npclist(UserList( _
                        UserIndex).flags.TargetNpc).char.CharIndex))
                    Exit Sub

                End If

                If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 3 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
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
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡El comercio con usuarios esta deshabilitado.!!" & FONTTYPE_INFO)
                    Exit Sub

                End If

                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡No puedes comerciar con los muertos!!" & FONTTYPE_INFO)
                    Exit Sub

                End If

                'soy yo ?
                If UserList(UserIndex).flags.TargetUser = UserIndex Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes comerciar con vos mismo..." & FONTTYPE_INFO)
                    Exit Sub

                End If

                'ta muy lejos ?
                If Distancia(UserList(UserList(UserIndex).flags.TargetUser).pos, UserList(UserIndex).pos) > 3 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "Z13")
                    Exit Sub

                End If

                'Ya ta comerciando ? es conmigo o con otro ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Comerciando = True And UserList(UserList( _
                    UserIndex).flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes comerciar con el usuario en este momento." & FONTTYPE_INFO)
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
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z31")

            End If

            Exit Sub

            '[KEVIN]------------------------------------------
        Case "/BOVEDA"

            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNpc > 0 Then
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 3 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Montado = True Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes usar la boveda estando arriba de tu Mascota!" & FONTTYPE_INFO)
                    Exit Sub

                End If

                If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Banquero Then
                    Call IniciarDeposito(UserIndex)

                End If

            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z31")

            End If

            Exit Sub
            '[/KEVIN]------------------------------------
    
        Case "/ENLISTAR"

            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z30")
                Exit Sub

            End If
           
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 5 Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Debes acercarte más." & FONTTYPE_INFO)
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

        Case "/INFORMACION"

            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z30")
                Exit Sub

            End If
           
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 5 Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
                Exit Sub

            End If
           
            If Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str( _
                        Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                    Exit Sub

                End If

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & _
                    "Tu deber es combatir criminales, cada 100 criminales que derrotes te dare una recompensa." & "°" & str(Npclist(UserList( _
                    UserIndex).flags.TargetNpc).char.CharIndex))
            Else

                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str( _
                        Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                    Exit Sub

                End If

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & _
                    "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te dare una recompensa." & "°" & str( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))

            End If

            Exit Sub
    
        Case "/ROSTRO"
    
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡Estas muerto!! Debes resucitarte para poder cambiar tu rostro!!" & FONTTYPE_ORO)
                Exit Sub

            End If
                
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z30")
                Exit Sub

            End If
               
            'Para que te cobre el dinero..

            If UserList(UserIndex).Stats.GLD < 20000 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Para cambiarte de rostro necesitas 20.000 monedas de oro." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).Stats.GLD >= 20000 Then
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 20000
                Call SendUserStatsBox(UserIndex)

            End If
              
            '¿El target es un NPC valido?
            If Not Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = 9 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Debes seleccionar el NPC correspondiente" & FONTTYPE_INFO)
                Exit Sub
            Else

                If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 10 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No podes hacer la cirujia plastica debido a que estas demasiado lejos." & _
                        FONTTYPE_INFO)
                    Exit Sub

                End If

            End If
        
            If UserList(UserIndex).Genero = "Hombre" Then

                Select Case UCase$(UserList(UserIndex).Raza)
                    Dim u As Integer

                    Case "HUMANO"
                        u = CInt(RandomNumber(1, 30))

                        If u > 30 Then u = 11

                    Case "ELFO"
                        u = CInt(RandomNumber(1, 12)) + 100

                        If u > 112 Then u = 104

                    Case "ELFO OSCURO"
                        u = CInt(RandomNumber(1, 9)) + 200

                        If u > 209 Then u = 203

                    Case "ENANO"
                        u = RandomNumber(1, 5) + 300

                        If u > 305 Then u = 304

                    Case "GNOMO"
                        u = RandomNumber(1, 6) + 400

                        If u > 406 Then u = 404
                        
                    Case "HOBBIT"
                        u = RandomNumber(608, 611)
                        
                        If u > 608 Then u = 611
                    
                    Case "ORCO"
                        u = RandomNumber(601, 605)
                        
                        If u > 601 Then u = 605
                        
                    Case "LICANTROPO"
                        u = RandomNumber(1, 30)
                        
                        If u > 1 Then u = 30

                    Case "VAMPIRO"
                        u = RandomNumber(710, 712)
                        
                        If u >= 710 Then u = 712

                    Case "CICLOPE"
                        u = RandomNumber(530, 532)
                        
                        If u > 530 Then u = 532
                    
                    Case Else
                        u = 1

                End Select

            End If

            'mujer
            If UserList(UserIndex).Genero = "Mujer" Then

                Select Case UCase$(UserList(UserIndex).Raza)

                    Case "HUMANO"
                        u = CInt(RandomNumber(1, 7)) + 69

                        If u > 76 Then u = 74

                    Case "ELFO"
                        u = CInt(RandomNumber(1, 7)) + 166

                        If u > 177 Then u = 172

                    Case "ELFO OSCURO"
                        u = CInt(RandomNumber(1, 11)) + 269

                        If u > 280 Then u = 265

                    Case "GNOMO"
                        u = RandomNumber(1, 5) + 469

                        If u > 474 Then u = 472

                    Case "ENANO"
                        u = RandomNumber(1, 3) + 369

                        If u > 372 Then u = 372

                    Case "HOBBIT"
                        u = RandomNumber(612, 615)
                        
                        If u > 612 Then u = 615
                    
                    Case "ORCO"
                        u = RandomNumber(606, 607)
                        
                        If u > 606 Then u = 607

                    Case "LICANTROPO"
                        u = CInt(RandomNumber(1, 7)) + 69

                        If u > 76 Then u = 74

                    Case "VAMPIRO"
                        u = RandomNumber(710, 712)

                        If u > 710 Then u = 712

                    Case "CICLOPE"
                        u = RandomNumber(533, 535)

                        If u > 533 Then u = 535

                    Case Else
                        u = 1

                End Select

            End If

            UserList(UserIndex).char.Head = u
            UserList(UserIndex).OrigChar.Head = u
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & "Espero que te guste tu nuevo rostro!!" & FONTTYPE_APU)
            '[MaTeO 9]
            Call ChangeUserChar(ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, val(u), UserList( _
                UserIndex).char.Heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList( _
                UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)
            '[/MaTeO 9]
            Exit Sub
           
        Case "/RECOMPENSA"

            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z30")
                Exit Sub

            End If

            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 5 Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub

            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z32")
                Exit Sub
            End If

            Select Case Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion
            
                Case 0

                    If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "No perteneces a la Armada del Credo, vete de aquí o te ahogaras en tu insolencia!!" & "°" & CStr( _
                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                        Exit Sub

                    End If

                    Call RecompensaArmadaReal(UserIndex)
                    Exit Sub

                Case 1

                    If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "No perteneces a la legión oscura!!!" & "°" & CStr( _
                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                        Exit Sub

                    End If

                    Call RecompensaCaos(UserIndex)
                
                    Exit Sub
                
                Case 3

                    If UserList(UserIndex).Faccion.Templario = 0 Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a la Orden Templaria, vete de aquí o volaras al vacio de tu ignorancia!!!" & "°" & _
                            CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                        Exit Sub

                    End If

                    Call RecompensaTemplario(UserIndex)
                    Exit Sub

                Case 5

                    If UserList(UserIndex).Faccion.Nemesis = 0 Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & "&H808080" & "°" & "No perteneces a los Caballeros de las Tinieblas, vete de aquí o te enterraremos vivo!!!" & "°" & CStr( _
                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                        Exit Sub

                    End If

                    Call RecompensaNemesis(UserIndex)
                    Exit Sub

            End Select

            Exit Sub
            
           
            Exit Sub
           
        Case "/MOTD"
            Call SendMOTD(UserIndex)
            Exit Sub
                    
        Case "/SALIRPARTY"
            Call mdParty.SalirDeParty(UserIndex)
            Exit Sub
        
        Case "/CREARPARTY"

            If Not mdParty.PuedeCrearParty(UserIndex) Then Exit Sub
            Call mdParty.CrearParty(UserIndex)
            Exit Sub

        Case "/PARTY"
            Call mdParty.SolicitarIngresoAParty(UserIndex)
            Exit Sub

    End Select
    
    If UCase$(Left$(rData, 14)) = "/CAMBIARBARCO " Then
        rData = val(Right$(rData, Len(rData) - 14))
           
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Clero Then
           
            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
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
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "No entiendo, Pon /CambiarBarco 1,2,3 o 4,5,6 (para recuperarlo)" & "°" & CStr( _
                        Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                   
            End Select
                   
        End If
           
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Abbadon Then
           
            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
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
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "No entiendo, Pon /CambiarBarco 1,2,3 o 4,5,6 (para recuperarlo)" & "°" & CStr( _
                        Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                   
            End Select
                   
        End If
           
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Tiniebla Then
           
            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
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
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & "&H808080" & "°" & "No entiendo, Pon /CambiarBarco 1,2,3 o 4,5,6 (para recuperarlo)" & "°" & CStr( _
                        Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                   
            End Select
                   
        End If
           
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Templario Then
           
            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
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
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "No entiendo, Pon /CambiarBarco 1,2,3 o 4,5,6 (para recuperarlo)" & "°" & CStr( _
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
            FrmUserhablan.hClan (Now & " Mensaje de " & UserList(UserIndex).Name & ":>" & rData)
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
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbYellow & "°" & "¡¡Necesitas 20000 monedas de oro para pagar el teletransporte!!" & "°" _
                    & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If
               
            Select Case Destino
                    
                Case "1"
                    Call WarpUserChar(UserIndex, 34, 23, 75, True)
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - "20000"
                    Call EnviarOro(UserIndex)
                    Exit Sub
                    
                Case "2"
                    Call WarpUserChar(UserIndex, 61, 52, 60, True)
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - "20000"
                    Call EnviarOro(UserIndex)
                    Exit Sub
                     
                Case "3"
                    Call WarpUserChar(UserIndex, 131, 35, 23, True)
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - "20000"
                    Call EnviarOro(UserIndex)
                    Exit Sub
                     
                Case Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbYellow & "°" & "¡A ese destino no puedes ir! Solo puedes ir a /llevame 1, 2 o 3" & "°" _
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
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Pareja offline." & FONTTYPE_INFO)
            Exit Sub

        End If
           
        If Len(tRespuesta) <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No has escrito un mensaje." & FONTTYPE_INFO)
            Exit Sub

        End If
       
        Call SendData(SendTarget.toIndex, UserList(UserIndex).Pareja, 0, "||(Pareja) " & UserList(UserIndex).Name & ": " & tRespuesta & FONTTYPE_INFO)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||(Pareja) " & UserList(UserIndex).Name & ": " & tRespuesta & FONTTYPE_INFO)
        Exit Sub

    End If
    
    If UCase$(Left$(rData, 6)) = "/PMSG " Then
        If Len(rData) > 6 Then
            Call mdParty.BroadCastParty(UserIndex, mid$(rData, 7))
            Call SendData(SendTarget.ToPartyArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbYellow & "°< " & mid$(rData, 7) & " >°" & CStr( _
                UserList(UserIndex).char.CharIndex))
            rData = Right$(rData, Len(rData) - 6)
            FrmUserhablan.hParty (Now & " Mensaje de " & UserList(UserIndex).Name & ":>" & rData)
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
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & Guilds(UserList(UserIndex).GuildIndex).GuildName & ": " & tStr & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No pertences a ningún clan." & FONTTYPE_INFO)

        End If

        Exit Sub

    End If
    
    If UCase$(rData) = "/ONLINEPARTY" Then
        Call mdParty.OnlineParty(UserIndex)
        Exit Sub

    End If
    
    'DIVORCIARSE
    If UCase$(Left$(rData, 13)) = "/DIVORCIARSE " Then
        rData = Right$(rData, Len(rData) - 13)
        Dim Pareja As String
        Pareja = UserList(UserIndex).Pareja
   
        TIndex = NameIndex(rData)
           
        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub

        End If
   
        If NameIndex(rData) <> UserList(UserIndex).Pareja Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No es tu pareja." & FONTTYPE_INFO)
            Exit Sub

        End If
               
        UserList(UserIndex).flags.Casado = 0
        UserList(TIndex).flags.Casado = 0
   
        UserList(UserIndex).Pareja = ""
        UserList(TIndex).Pareja = ""
   
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Te has divorciado de " & UserList(TIndex).Name & "!" & FONTTYPE_WARNING)
        Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " declaro el divorcio, tu matrimonio se anuló." & _
            FONTTYPE_WARNING)
        Exit Sub

    End If
    
    '[yb]
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

    '[/yb]
    
    If UCase$(Left$(rData, 5)) = "/ROL " Then
        rData = Right$(rData, Len(rData) - 5)
        Call SendData(SendTarget.toIndex, 0, 0, "|| " & "Su solicitud ha sido enviada" & FONTTYPE_INFO)
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
        
        Call SendData(SendTarget.toIndex, UserIndex, 0, "CSOS")
        
        Exit Sub

    End If
    
    If UCase$(Left$(rData, 5)) = "/SHOW" Then
        
        If UserList(UserIndex).flags.Privilegios = PlayerType.Dios Or UserList(UserIndex).flags.Privilegios = PlayerType.SemiDios Or UserList( _
            UserIndex).flags.Privilegios = PlayerType.Consejero Then
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
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El mensaje SOS no ha sido enviado, revisa el mensaje." & FONTTYPE_INFO)
            Else
                SSName = UserList(UserIndex).Name
                SSMsg = rData
                SSFH = Now
                SSRev = val(GetVar(App.Path & "\Logs\Show\SOS\" & SSName & ".ini", "Config", "NumMsg"))
                SSSuma = SSRev + "1"
              
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El mensaje SOS ha sido enviado, ahora solo debes esperar que un gm te responda." _
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
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El mensaje DENUNCIA no ha sido enviado, revisa el mensaje." & FONTTYPE_INFO)
            Else

                SDName = UserList(UserIndex).Name
                SDMsg = rData
                SDFH = Now
                SDRev = val(GetVar(App.Path & "\Logs\Show\DENUNCIA\" & SDName & ".ini", "Config", "NumMsg"))
                SDSuma = SDRev + "1"
              
                Call SendData(SendTarget.toIndex, UserIndex, 0, _
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
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El mensaje BUG no ha sido enviado, revisa el mensaje." & FONTTYPE_INFO)
            Else
               
                SBName = UserList(UserIndex).Name
                SBMsg = rData
                SBFH = Now
                SBRev = val(GetVar(App.Path & "\Logs\Show\BUG\" & SBName & ".ini", "Config", "NumMsg"))
                SBSuma = SBRev + "1"
              
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El mensaje BUG ha sido enviado, ahora un gm revisará el bug enviado ¡Gracias!." _
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
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El mensaje SUGERENCIA no ha sido enviado, revisa el mensaje." & FONTTYPE_INFO)
            Else
               
                SGName = UserList(UserIndex).Name
                SGMsg = rData
                SGFH = Now
                SGRev = val(GetVar(App.Path & "\Logs\Show\SUGERENCIA\" & SGName & ".ini", "Config", "NumMsg"))
                SGSuma = SGRev + "1"
              
                Call SendData(SendTarget.toIndex, UserIndex, 0, _
                    "||El mensaje SUGERENCIA ha sido enviado, el staff debatira su sugerencia ¡Gracias!." & FONTTYPE_INFO)
                Call SendData(SendTarget.ToAdmins, 0, 0, "|| Nueva SUGERENCIA del Usuario: " & UserList(UserIndex).Name & FONTTYPE_INFO)
              
                Call WriteVar(App.Path & "\Logs\Show\SUGERENCIA\" & SGName & ".ini", "Config", "NumMsg", SGSuma)
                Call WriteVar(App.Path & "\Logs\Show\SUGERENCIA\" & SGName & ".ini", "Mensaje" & SGSuma, "Mensaje", SGMsg)
                Call WriteVar(App.Path & "\Logs\Show\SUGERENCIA\" & SGName & ".ini", "Mensaje" & SGSuma, "HoraFecha", SGFH)
           
            End If

        End If
        
    End If
    
    If UCase$(Left$(rData, 9)) = "/GM_QUEST" Then
        rData = Right$(rData, Len(rData) - 9)
         
        If UserList(UserIndex).flags.Quest = 1 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Debes esperar a tu turno para que el GM te haga teletransporte." & FONTTYPE_INFO)
            Exit Sub
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, _
                "||Has enviado GM QUEST ahora debes esperar a tu turno para que el GM te haga teletransporte." & FONTTYPE_INFO)
            UserList(UserIndex).flags.Quest = 1
            Call Quest.Push(rData, UserList(UserIndex).Name)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||Nuevo QUEST del Usuario: " & UserList(UserIndex).Name & FONTTYPE_INFO)
              
        End If

    End If
    
    Select Case UCase$(Left$(rData, 7))

        ' vaya mierda de codigo, solamente sumonea JAJA
        Case "/TORNEO"

            If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserList(UserIndex).flags.EstaDueleando1 = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes ir a torneo estando plantes!." & FONTTYPE_WARNING)
                Exit Sub

            End If

            If UserIndex = Team.Pj1 Or UserIndex = Team.Pj2 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If Hay_Torneo = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No hay ningún torneo disponible." & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).Stats.ELV < Torneo_Nivel_Minimo Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu nivel es: " & UserList(UserIndex).Stats.ELV & ".El requerido es: " & _
                    Torneo_Nivel_Minimo & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).Stats.ELV > Torneo_Nivel_Maximo Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu nivel es: " & UserList(UserIndex).Stats.ELV & ".El máximo es: " & _
                    Torneo_Nivel_Maximo & FONTTYPE_INFO)
                Exit Sub

            End If

            If Torneo_Inscriptos >= Torneo_Cantidad Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El cupo ya ha sido alcanzado." & FONTTYPE_INFO)
                Exit Sub

            End If

            For i = 1 To 8

                If UCase$(UserList(UserIndex).Clase) = UCase$(Torneo_Clases_Validas(i)) And Torneo_Clases_Validas2(i) = 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu clase no es válida en este torneo." & FONTTYPE_INFO)
                    Exit Sub

                End If

            Next
            
            Dim NuevaPos As WorldPos
            
            'Old, si entras no salis =P
            If Not Torneo.Existe(UserList(UserIndex).Name) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estás en la lista de espera del torneo. Estás en el puesto nº " & _
                    Torneo.Longitud + 1 & FONTTYPE_INFO)
                Call Torneo.Push(rData, UserList(UserIndex).Name)
                
                Call SendData(SendTarget.ToAdmins, 0, 0, "||/TORNEO [" & UserList(UserIndex).Name & "]" & FONTTYPE_INFOBOLD)
                Torneo_Inscriptos = Torneo_Inscriptos + 1

                If Torneo_Inscriptos = Torneo_Cantidad Then
                    Call SendData(SendTarget.toall, 0, 0, "||Cupo alcanzado." & FONTTYPE_CELESTE_NEGRITA)

                End If

                If Torneo_SumAuto = 1 Then
                    Dim FuturePos As WorldPos
                    FuturePos.Map = Torneo_Map
                    FuturePos.X = Torneo_X
                    FuturePos.Y = Torneo_Y
                    Call ClosestLegalPos(FuturePos, NuevaPos)

                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)

                End If

            Else
                '                Call Torneo.Quitar(UserList(Userindex).Name)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Ya estás en la lista de espera del torneo." & FONTTYPE_INFO)

                '                Torneo_Inscriptos = Torneo_Inscriptos - 1
                '                If Torneo_SumAuto = 1 Then
                '                    Call WarpUserChar(Userindex, 1, 50, 50, True)
                '                End If
            End If

            Exit Sub

    End Select
    
    Select Case UCase$(Left$(rData, 3))

        Case "/GM"
            rData = Right$(rData, Len(rData) - 4)
        
            Dim GMRev  As Long
            Dim GMSuma As Long
            Dim GMName As String
            Dim GMMsg  As String
            Dim GMFH   As String
                
            If rData = "" Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El mensaje SOS no ha sido enviado, revisa el mensaje." & FONTTYPE_INFO)
            Else
                GMName = UserList(UserIndex).Name
                GMMsg = rData
                GMFH = Now
                GMRev = val(GetVar(App.Path & "\Logs\Consultas\" & GMName & ".ini", "Config", "NumMsg"))
                GMSuma = GMRev + "1"
              
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que un gm te responda." & _
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
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z12" & FONTTYPE_INFO)
                Exit Sub

            End If

            rData = Right$(rData, Len(rData) - 6)

            If Not AsciiValidos(rData) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||La descripcion tiene caracteres invalidos." & FONTTYPE_INFO)
                Exit Sub

            End If

            UserList(UserIndex).Desc = Trim$(rData)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La descripcion a cambiado." & FONTTYPE_INFO)
            Exit Sub

        Case "/VOTO "
            rData = Right$(rData, Len(rData) - 6)

            If Not modGuilds.v_UsuarioVota(UserIndex, rData, tStr) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Voto NO contabilizado: " & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Voto contabilizado." & FONTTYPE_GUILD)

            End If

            Exit Sub

    End Select
    
    If UCase$(Left$(rData, 7)) = "/PENAS " Then
        Name = Right$(rData, Len(rData) - 7)

        If Name = "" Then Exit Sub
        
        Name = Replace(Name, "\", "")
        Name = Replace(Name, "/", "")
        
        If FileExist(CharPath & Name & ".chr", vbNormal) Then
            tInt = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))

            If tInt = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Sin prontuario.." & FONTTYPE_INFO)
            Else

                While tInt > 0

                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & tInt & "- " & GetVar(CharPath & Name & ".chr", "PENAS", "P" & tInt) & _
                        FONTTYPE_INFO)
                    tInt = tInt - 1
                Wend

            End If

        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Personaje """ & Name & """ inexistente." & FONTTYPE_INFO)

        End If

        Exit Sub

    End If
    
    Select Case UCase$(Left$(rData, 8))

        Case "/PASSWD "
            rData = Right$(rData, Len(rData) - 8)

            If Len(rData) < 6 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El password debe tener al menos 6 caracteres." & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El password ha sido cambiado." & FONTTYPE_INFO)
                UserList(UserIndex).Password = MD5String(rData)
                
#If MYSQL = 1 Then
                Call Add_DataBase(UserIndex, "Account")
#End If
                
            End If

            Exit Sub

    End Select
    
    Select Case UCase$(Left$(rData, 9))
            
        'Comando /APOSTAR basado en la idea de DarkLight,
        'pero con distinta probabilidad de exito.
        
        Case "/APOSTAR "
            rData = Right(rData, Len(rData) - 9)
            tLong = CLng(val(rData))

            If tLong > 32000 Then tLong = 32000
            n = tLong

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z12")
            ElseIf UserList(UserIndex).flags.TargetNpc = 0 Then
                'Se asegura que el target es un npc
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z30")
            ElseIf Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
            ElseIf Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> eNPCType.Timbero Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "No tengo ningun interes en apostar." & "°" & str(Npclist( _
                    UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            ElseIf n < 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "El minimo de apuesta es 1 moneda." & "°" & str(Npclist( _
                    UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            ElseIf n > 5000 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "El maximo de apuesta es 5000 monedas." & "°" & str(Npclist( _
                    UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            ElseIf UserList(UserIndex).Stats.GLD < n Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "No tienes esa cantidad." & "°" & str(Npclist(UserList( _
                    UserIndex).flags.TargetNpc).char.CharIndex))
            Else

                If RandomNumber(1, 100) <= 47 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + n
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Felicidades! Has ganado " & CStr(n) & _
                        " monedas de oro!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                    
                    Apuestas.Perdidas = Apuestas.Perdidas + n
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - n
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Lo siento, has perdido " & CStr(n) & " monedas de oro." _
                        & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                
                    Apuestas.Ganancias = Apuestas.Ganancias + n
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))

                End If

                Apuestas.Jugadas = Apuestas.Jugadas + 1
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
                
                Call EnviarOro(UserIndex)

            End If

            Exit Sub

    End Select
    
    Select Case UCase$(Left$(rData, 8))

        Case "/RETIRAR" 'RETIRA ORO EN EL BANCO o te saca de la armada

            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z30")
                Exit Sub

            End If
             
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = 5 Then
                
                'Se quiere retirar de la armada
                'If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
                '    If Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 0 Then
                '        Call ExpulsarFaccionReal(UserIndex)
                '        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "º" & _
                '                "Serás bienvenido a las fuerzas imperiales si deseas regresar." & "º" & str(Npclist(UserList( _
                '                UserIndex).flags.TargetNpc).char.CharIndex))
                ''
                '   Else
                '       Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "º" & "¡¡¡Sal de aquí bufón!!!" & "º" & str(Npclist( _
                '               UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                ''
                '                   End If
                '
                '                ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
                '
                '                    If Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 1 Then
                '                        Call ExpulsarFaccionCaos(UserIndex)
                ''                        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "º" & "Ya volverás arrastrandote." & "º" & str(Npclist( _
                '                               UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                '                   Else
                '                       Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "º" & "Sal de aquí maldito criminal" & "º" & str(Npclist( _
                                        UserList(UserIndex).flags.TargetNpc).char.CharIndex))

                '                   End If
                    
                '               ElseIf UserList(UserIndex).Faccion.Nemesis = 1 Then
                '
                '                    If Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 5 Then
                '                        Call ExpulsarFaccionNemesis(UserIndex)
                '                        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "º" & "Ya volverás arrastrandote." & "º" & CStr(Npclist( _
                '                                UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                ''                    Else
                '                       Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "º" & "Sal de aquí maldito" & "º" & CStr(Npclist(UserList( _
                '                               UserIndex).flags.TargetNpc).char.CharIndex))
                '
                '                    End If
                '
                '                ElseIf UserList(UserIndex).Faccion.Templario = 1 Then
                '
                '                    If Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 3 Then
                '                        Call ExpulsarFaccionTemplario(UserIndex)
                ''                        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "º" & "Ya volverás arrastrandote." & "º" & CStr(Npclist( _
                '                               UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                '                   Else
                '                       Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "º" & "Sal de aquí maldito" & "º" & CStr(Npclist(UserList( _
                ''                               UserIndex).flags.TargetNpc).char.CharIndex))
                '
                ''                    End If
                ''
                '               Else
                '                   Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "º" & "¡No perteneces a ninguna fuerza!" & "º" & CStr(Npclist( _
                '                           UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                '
                '                End If

                '                Exit Sub
             
            End If
             
            If Len(rData) = 8 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Debes indicar el monto de cuanto quieres retirar" & FONTTYPE_INFO)
                Exit Sub

            End If
             
            rData = Right$(rData, Len(rData) - 9)

            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> eNPCType.Banquero Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub

            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 10 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
                Exit Sub

            End If

            If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                CloseSocket (UserIndex)
                Exit Sub

            End If

            If val(rData) > 0 And val(rData) <= UserList(UserIndex).Stats.Banco Then
                UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(rData)
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + val(rData)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & _
                    " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList( _
                    UserIndex).flags.TargetNpc).char.CharIndex & FONTTYPE_INFO)

            End If

            Call EnviarOro(val(UserIndex)) 'ak antes habia un senduserstatsbox. lo saque. NicoNZ
            Exit Sub

    End Select
    
    Select Case UCase$(Left$(rData, 11))

        Case "/DEPOSITAR " 'DEPOSITAR ORO EN EL BANCO

            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z30")
                Exit Sub

            End If

            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
                Exit Sub

            End If

            rData = Right$(rData, Len(rData) - 11)

            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> eNPCType.Banquero Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub

            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 10 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z27")
                Exit Sub

            End If

            If CLng(val(rData)) > 0 And CLng(val(rData)) <= UserList(UserIndex).Stats.GLD Then
                UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + val(rData)
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rData)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & _
                    " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList( _
                    UserIndex).flags.TargetNpc).char.CharIndex & FONTTYPE_INFO)

            End If

            Call EnviarOro(val(UserIndex))
            Exit Sub

        Case "/DENUNCIAR "

            If denuncias = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Las denuncias estan desactivadas!" & FONTTYPE_DENUNCIAR)
                Exit Sub

            End If
            
            If UserList(UserIndex).flags.Silenciado = 1 Then
                Exit Sub

            End If
            
            rData = Right$(rData, Len(rData) - 11)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||El PJ " & LCase$(UserList(UserIndex).Name) & " Denuncia: " & rData & FONTTYPE_DENUNCIAR)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu Denuncia ha sido enviada." & FONTTYPE_DENUNCIAR)
            
            Exit Sub
            
        Case "/CERRARCLAN"

            If Not UserList(UserIndex).GuildIndex >= 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No perteneces a ningún clan." & FONTTYPE_GUILD)
                Exit Sub

            End If

            If UCase$(Guilds(UserList(UserIndex).GuildIndex).Fundador) <> UCase$(UserList(UserIndex).Name) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No eres líder del clan." & FONTTYPE_GUILD)
                Exit Sub

            End If

            If Guilds(UserList(UserIndex).GuildIndex).CantidadDeMiembros > 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Debes hechar a todos los miembros del clan para cerrarlo." & FONTTYPE_GUILD)
                Exit Sub

            End If

            'If UserList(UserIndex).flags.YaCerroClan = 1 Then
            'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya has cerrado un clan antes" & FONTTYPE_GUILD)
            'Exit Sub
            'End If

            Call SendData(SendTarget.toall, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " cerró." & FONTTYPE_GUILD)

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
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Ya has fundado un clan, sólo se puede fundar uno por personaje." & FONTTYPE_INFO)
                Exit Sub
            End If

            If modGuilds.PuedeFundarUnClan(UserIndex, tStr) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "SHOWFUN")
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)

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
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El personaje no está online." & FONTTYPE_INFO)

            End If

            Exit Sub

        Case "/PARTYLIDER "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)

            If tInt > 0 Then
                Call mdParty.TransformarEnLider(UserIndex, tInt)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El personaje no está online." & FONTTYPE_INFO)

            End If

            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 14))
            
        Case "/CANCELARPARTY"
            UserList(UserIndex).PartySolicitud = 0
            Exit Sub
        
        Case "/ACEPTARPARTY "
            rData = Right$(rData, Len(rData) - 14)
            tInt = NameIndex(rData)

            If tInt > 0 Then
                Call mdParty.AprobarIngresoAParty(UserIndex, tInt)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El personaje no está online." & FONTTYPE_INFO)

            End If

            Exit Sub

        Case "/MIEMBROSCLAN "
            rData = Trim(Right(rData, Len(rData) - 14))
            Name = Replace(rData, "\", "")
            Name = Replace(rData, "/", "")
    
            If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
                Exit Sub

            End If
            
            tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
            
            For i = 1 To tInt
                tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
                'tstr es la victima
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & tStr & "<" & rData & ">." & FONTTYPE_INFO)
            Next i
        
            Exit Sub

    End Select
    
    Procesado = False
           
End Sub

Public Sub ActGM()
    
    frmMain.Gms.Clear
    
    Dim loopc     As Integer
    Dim UserIndex As Integer
    
    For loopc = 1 To LastUser

        'Tiene nombre? Es GM? Si es Dios o Admin, nosotros lo somos también??
        If (UserList(loopc).Name <> "") And UserList(loopc).flags.Privilegios > PlayerType.User And (UserList(loopc).flags.Privilegios < _
            PlayerType.Dios Or UserList(loopc).flags.Privilegios >= PlayerType.Dios) Then
            frmMain.Gms.AddItem (UserList(loopc).Name)

        End If

    Next loopc
   
End Sub

Public Sub ActUser()
    Dim loopc As Integer
    frmMain.User.Clear

    For loopc = 1 To LastUser

        If Len(UserList(loopc).Name) <> 0 And UserList(loopc).flags.Privilegios <= PlayerType.Consejero Then
            
            frmMain.User.AddItem (UserList(loopc).Name)

        End If

    Next loopc

End Sub

Sub MostrarTimeOnline()

    frmMain.CantOnMin.caption = "Minutos Online: " & OnMin
    frmMain.CantOnHor.caption = "Horas Online: " & OnHor
    frmMain.CantOnDay.caption = "Dias Online: " & OnDay

End Sub

Public Sub RegUser()
    Dim loopc As Integer
    Dim tStr  As String
    Dim Count As Long

    For loopc = 1 To NumUsers

        If UserList(loopc).flags.Privilegios = PlayerType.User Then
             
            tStr = UserList(loopc).Name & "," & tStr
            
            Count = Count + "1"
            
            frmMain.CantUsuarios.caption = "Número de usuarios: " & Count

        End If

    Next loopc
    
    If Len(tStr) = 0 Then
        frmMain.CantUsuarios.caption = "Número de usuarios: 0"

    End If

End Sub

Public Sub RegGM()
    Dim loopc As Integer
    Dim tStr  As String
    Dim Count As Long

    For loopc = 1 To NumUsers

        If UserList(loopc).flags.Privilegios > PlayerType.User Then
             
            tStr = UserList(loopc).Name & "," & tStr
            Count = Count + "1"
            frmMain.CantNumGM.caption = "Número de gms: " & Count

        End If

    Next loopc
    
    If Len(tStr) = 0 Then
        frmMain.CantNumGM.caption = "Número de gms: 0"

    End If

End Sub
