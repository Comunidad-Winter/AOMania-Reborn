Attribute VB_Name = "TCP_HandleData1"

Option Explicit

Public Sub HandleData_1(ByVal UserIndex As Integer, rData As String, ByRef Procesado As Boolean)

    Dim LoopC As Integer
    Dim nPos As WorldPos
    Dim tStr As String
    Dim tInt As Integer
    Dim tLong As Long
    Dim TIndex As Integer
    Dim tName As String
    Dim tMessage As String
    Dim AuxInd As Integer
    Dim Arg1 As String
    Dim Arg2 As String
    Dim Arg3 As String
    Dim Arg4 As String
    Dim Ver As String
    Dim encpass As String
    Dim Pass As String
    Dim Mapa As Integer
    Dim Name As String
    Dim ind
    Dim n As Integer
    Dim wpaux As WorldPos
    Dim mifile As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim DummyInt As Integer
    Dim T() As String
    Dim i As Integer

    Procesado = True    'ver al final del sub

    Select Case UCase$(Left$(rData, 1))

    Case ";"    'Hablar
        rData = Right$(rData, Len(rData) - 1)

        If InStr(rData, "°") Then
            Exit Sub

        End If

        '[Consejeros]
        If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
            Call LogGM(UserList(UserIndex).Name, "Dijo: " & rData)

        End If

        ind = UserList(UserIndex).char.CharIndex

        'piedra libre para todos los compas!
        If UserList(UserIndex).flags.Silenciado = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas Silenciado!" & FONTTYPE_WARNING)
            Exit Sub

        End If

        ' If UserList(UserIndex).flags.Oculto > 0 Then
        '     UserList(UserIndex).flags.Oculto = 0
        '     UserList(UserIndex).Counters.Ocultando = 0
        '
        '                If UserList(UserIndex).flags.Invisible = 0 Then
        '                    Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, "NOVER" & UserList(UserIndex).char.CharIndex & ",0," & UserList( _
                             '                            UserIndex).PartyIndex)
        '
        '                    Call SendData(SendTarget.toindex, UserIndex, 0, "Z11")
        '
        '                End If
        '
        '            End If

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToDeadArea, UserIndex, UserList(UserIndex).pos.Map, "||12632256°" & rData & "°" & CStr(ind))
        Else

            '&H4080FF&
            '&H80FF&
            If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & &H4080FF & "°" & rData & "°" & CStr(ind))
                FrmUserhablan.hUser (now & " Mensaje de " & UserList(UserIndex).Name & ":>" & rData)
                Call addConsole(UserList(UserIndex).Name & ": " & rData, 255, 0, 0, True, False)
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbWhite & "°" & rData & "°" & CStr(ind))
                FrmUserhablan.hUser (now & " Mensaje de " & UserList(UserIndex).Name & ":>" & rData)
                Call addConsole(UserList(UserIndex).Name & ": " & rData, 255, 0, 0, True, False)

            End If

        End If

        Exit Sub

    Case "-"    'Gritar

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub

        End If

        If UserList(UserIndex).flags.Silenciado = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas Silenciado!" & FONTTYPE_WARNING)
            Exit Sub

        End If

        rData = Right$(rData, Len(rData) - 1)

        If InStr(rData, "°") Then
            Exit Sub

        End If

        '[Consejeros]
        If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
            Call LogGM(UserList(UserIndex).Name, "Grito: " & rData)

        End If

        'piedra libre para todos los compas!
        'If UserList(UserIndex).flags.Oculto > 0 Then
        '    UserList(UserIndex).flags.Oculto = 0
        '    UserList(UserIndex).Counters.Ocultando = 0
        '
        '                If UserList(UserIndex).flags.Invisible = 0 Then
        '                    Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, "NOVER" & UserList(UserIndex).char.CharIndex & ",0," & UserList( _
                             ''                            UserIndex).PartyIndex)
        '
        '                   Call SendData(SendTarget.toindex, UserIndex, 0, "Z11")

        '                End If
        '
        '            End If

        ind = UserList(UserIndex).char.CharIndex
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbRed & "°" & rData & "°" & str(ind))
        Exit Sub

    Case "\"    'Susurrar al oido

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub

        End If

        If UserList(UserIndex).flags.Silenciado = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas Silenciado!" & FONTTYPE_WARNING)
            Exit Sub

        End If

        rData = Right$(rData, Len(rData) - 1)
        tName = ReadField(1, rData, 32)

        'A los dioses y admins no vale susurrarles si no sos uno vos mismo (así no pueden ver si están conectados o no)
        If (EsDios(tName) Or EsAdmin(tName)) And UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes susurrarle a los Dioses y Admins." & FONTTYPE_INFO)
            Exit Sub

        End If

        'A los Consejeros y SemiDioses no vale susurrarles si sos un PJ común.
        If UserList(UserIndex).flags.Privilegios = PlayerType.User And (EsSemiDios(tName) Or EsConsejero(tName)) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes susurrarle a los GMs" & FONTTYPE_INFO)
            Exit Sub

        End If

        TIndex = NameIndex(tName)

        If TIndex <> 0 Then
            If Len(rData) <> Len(tName) Then
                tMessage = Right$(rData, Len(rData) - (1 + Len(tName)))
                FrmUserhablan.hPrivado (now & " Mensaje de " & UserList(UserIndex).Name & " a " & tName & ":>" & tMessage)

                If UserList(UserIndex).flags.Privilegios = PlayerType.Dios Then
                    Call LogGM(UserList(UserIndex).Name, "Dice en privado a " & tName & ": " & tMessage)
                ElseIf UserList(UserIndex).flags.Privilegios = PlayerType.SemiDios Then
                    Call LogGM(UserList(UserIndex).Name, "Dice en privado a " & tName & ": " & tMessage)
                ElseIf UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                    Call LogGM(UserList(UserIndex).Name, "Dice en privado a " & tName & ": " & tMessage)
                ElseIf UserList(UserIndex).flags.Privilegios = PlayerType.User Then
                    Call LogUser(UserList(UserIndex).Name, "Dice en privado a " & tName & ": " & tMessage)

                End If

            Else
                tMessage = " "

            End If

            If Not EstaPCarea(UserIndex, TIndex) Then
                Call SendData(SendTarget.ToIndex, TIndex, 0, "||" & UserList(UserIndex).Name & ">" & tMessage & FONTTYPE_CONSEJO)
                Call SendData(SendTarget.ToIndex, UserIndex, UserList(UserIndex).pos.Map, ">" & tMessage & FONTTYPE_CONSEJO)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & UserList(UserIndex).Name & ">" & tMessage & FONTTYPE_WARNING)

                Exit Sub

            End If

            ind = UserList(UserIndex).char.CharIndex
            Call SendData(SendTarget.ToIndex, TIndex, 0, "||" & UserList(UserIndex).Name & ">" & tMessage & FONTTYPE_CONSEJO)
            Call SendData(SendTarget.ToIndex, UserIndex, UserList(UserIndex).pos.Map, ">" & tMessage & FONTTYPE_CONSEJO)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & UserList(UserIndex).Name & ">" & tMessage & FONTTYPE_WARNING)
            Exit Sub

            '[Consejeros]
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                Call LogGM(UserList(UserIndex).Name, "Le dijo a '" & UserList(TIndex).Name & "' " & tMessage)

            End If

            Call SendData(SendTarget.ToIndex, TIndex, 0, "||" & UserList(UserIndex).Name & ">" & vbBlue & "°" & tMessage & "°" & str(ind))
            Call SendData(SendTarget.ToIndex, UserIndex, UserList(UserIndex).pos.Map, "||" & vbBlue & "°" & tMessage & "°" & str(ind))

            '[CDT 17-02-2004]
            If UserList(UserIndex).flags.Privilegios < PlayerType.SemiDios Then
                Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, UserList(UserIndex).pos.Map, "||" & vbYellow & "°" & "a " & _
                                                                                                            UserList(TIndex).Name & "> " & tMessage & "°" & str(ind))

            End If

            '[/CDT]
            Exit Sub

        End If

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z13")
        Exit Sub

    Case "Ñ"    'Moverse
        'Dim dummy As Long
        'Dim TempTick As Long
        'If UserList(UserIndex).flags.TimesWalk >= 30 Then
        'TempTick = GetTickCount And &H7FFFFFFF
        'dummy = (TempTick - UserList(UserIndex).flags.StartWalk)
        'If dummy < 6050 Then
        'If TempTick - UserList(UserIndex).flags.CountSH > 90000 Then
        '    UserList(UserIndex).flags.CountSH = 0
        'End If
        'If Not UserList(UserIndex).flags.CountSH = 0 Then
        '   dummy = 126000 \ dummy
        '   Call LogHackAttemp("Tramposo SH: " & UserList(UserIndex).name & " , " & dummy)
        '    Call SendData(SendTarget.ToAdmins, 0, 0, "||Servidor> " & UserList(UserIndex).name & " ha sido echado por el servidor por posible uso de SH." & FONTTYPE_SERVER)
        '     Call CloseSocket(UserIndex)
        '      Exit Sub
        '   Else
        '        UserList(UserIndex).flags.CountSH = TempTick
        '     End If
        '  End If
        '   UserList(UserIndex).flags.StartWalk = TempTick
        '    UserList(UserIndex).flags.TimesWalk = 0
        ' End If

        'UserList(UserIndex).flags.TimesWalk = UserList(UserIndex).flags.TimesWalk + 1

        rData = Right$(rData, Len(rData) - 1)

        Dim direction As Integer

        direction = ReadField(1, rData, 44)

        If UserList(UserIndex).flags.Meditando Then
            Exit Sub
        End If

        If UserList(UserIndex).flags.Paralizado = 1 Then
            Exit Sub
        End If

        'salida parche
        If UserList(UserIndex).Counters.Saliendo Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z15")
            UserList(UserIndex).Counters.Saliendo = False
            UserList(UserIndex).Counters.Salir = 0

        End If

        'If UserList(UserIndex).flags.Oculto = 1 Then

        'Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, "NOVER" & UserList( _
         UserIndex).char.CharIndex & ",0," & UserList(UserIndex).PartyIndex)

        'Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z11")
        'Call SendData(SendTarget.ToIndex, UserIndex, 0, "INVI0")

        'UserList(UserIndex).flags.Oculto = 0
        ' UserList(UserIndex).Counters.Ocultando = 0

        'End If



        If UserList(UserIndex).flags.Paralizado = 0 Then
            If Not UserList(UserIndex).flags.Descansar And Not UserList(UserIndex).flags.Meditando Then

                Dim tick As Long

                tick = ReadField(2, rData, 44)

                Dim tnow As Long
                tnow = (GetTickCount() And &H7FFFFFFF)

                If (tick > tnow) Then
                    'Call SendData(togms, UserIndex, 0, "||Delay erroneo!" & "´" & FontTypeNames.FONTTYPE_INFO)
                    'CRAW; 18/09/2019 --> Avisar de un delay erroneo. (NO USUAL)
                Else
                    UserList(UserIndex).char.delay = tnow - tick

                    'Call SendData(toIndex, UserIndex, 0, "||Tu Delay es de " & UserList(UserIndex).char.delay & "´" & FONTTYPE_INFO)
                End If

                Call MoveUserChar(UserIndex, direction)
            ElseIf UserList(UserIndex).flags.Descansar Then
                UserList(UserIndex).flags.Descansar = False
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has dejado de descansar." & FONTTYPE_INFO)
                Call MoveUserChar(UserIndex, direction)
            ElseIf UserList(UserIndex).flags.Meditando Then
                UserList(UserIndex).flags.Meditando = False
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "MEDOK")
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z16")
                UserList(UserIndex).char.FX = 0
                UserList(UserIndex).char.loops = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & 0 _
                                                                                         & "," & 0)

            End If

        Else    'paralizado

            '[CDT 17-02-2004] (<- emmmmm ?????)
            If Not UserList(UserIndex).flags.UltimoMensaje = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z17")
                UserList(UserIndex).flags.UltimoMensaje = 1

            End If

            '[/CDT]
        End If

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call Empollando(UserIndex)
        Else
            UserList(UserIndex).flags.EstaEmpo = 0
            UserList(UserIndex).EmpoCont = 0

        End If

        If UserList(UserIndex).pos.Map = MapaCasaAbandonada1 Then
            Call Efecto_CaminoCasaEncantada(UserIndex)
        End If

        Exit Sub

    End Select

    Select Case UCase$(rData)

        'Implementaciones del anti cheat By NicoNZ
    Case "TENGOSH"
        Call SendData(SendTarget.ToAdmins, 0, 0, "||Sistema Anti Cheat 2> " & UserList(UserIndex).Name & _
                                               " ha sido expulsado por el Anti Cheat. Por favor, que algun gm lo siga ya que es muy probable que tenga un programa externo corriendo." _
                                               & FONTTYPE_SERVER)
        Call CloseSocket(UserIndex)
        Exit Sub

    Case "RPU"    'Pedido de actualizacion de la posicion
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).pos.X & "," & UserList(UserIndex).pos.Y)
        Exit Sub

    Case "KC"

        'If not in combat mode, can't attack
        If Not UserList(UserIndex).flags.SeguroCombate Then

            Call SendData(SendTarget.ToIndex, UserIndex, 0, _
                          "||No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate." & FONTTYPE_Motd4)
            Exit Sub

        End If

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub
        End If

        UserList(UserIndex).Counters.TimerAttack = IntervaloAttack

        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Proyectil = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z19")
                Exit Sub
            End If

        End If

        Call UsuarioAtaca(UserIndex)

        Exit Sub

    Case "AG"

        If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then
            'Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, "NOVER" & UserList(UserIndex).char.CharIndex & ",0," & UserList( _
             UserIndex).PartyIndex)

            'Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z11")
            'UserList(UserIndex).flags.Invisible = 0
            'UserList(UserIndex).flags.Oculto = 0
            'UserList(UserIndex).Counters.Ocultando = 0
            'UserList(UserIndex).Counters.Invisibilidad = 0
            'Call SendData(SendTarget.ToIndex, UserIndex, 0, "INVI0")

        End If

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub

        End If

        '[Consejeros]
        If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero And Not UserList(UserIndex).flags.EsRolesMaster Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes tomar ningun objeto. " & FONTTYPE_INFO)
            Exit Sub

        End If

        Call GetObj(UserIndex)
        Exit Sub

    Case "SEG"    'Activa / desactiva el seguro

        If UserList(UserIndex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z21")
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ONONS")
            UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro

        End If

        Exit Sub

    Case "ACTUALIZAR"
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).pos.X & "," & UserList(UserIndex).pos.Y)
        Exit Sub

    Case "GLINFO"
        tStr = SendGuildLeaderInfo(UserIndex)

        If tStr = vbNullString Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "GL" & SendGuildsList(UserIndex))
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "LEADERI" & tStr)

        End If

        Exit Sub

    Case "ATRI"
        Call EnviarAtrib(UserIndex)
        Exit Sub

    Case "FAMA"
        Call EnviarFama(UserIndex)
        Exit Sub

    Case "ESKI"
        Call EnviarSkills(UserIndex)
        Exit Sub

    Case "FEST"    'Mini estadisticas :)
        Call EnviarMiniEstadisticas(UserIndex)
        Exit Sub

        '[Alejo]
    Case "FINCOM"
        'User sale del modo COMERCIO
        UserList(UserIndex).flags.Comerciando = False
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "FINCOMOK")
        Exit Sub

    Case "FINCOA"
        UserList(UserIndex).flags.Comerciando = False
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "FINCOMOA")
        Exit Sub

    Case "FINCOC"
        UserList(UserIndex).flags.Comerciando = False
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "FINCOMOC")
        Exit Sub

    Case "FINCOMUSU"

        'Sale modo comercio Usuario
        If UserList(UserIndex).ComUsu.DestUsu > 0 And UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
            Call SendData(SendTarget.ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & _
                                                                                   " ha dejado de comerciar con vos." & FONTTYPE_TALK)
            Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)

        End If

        Call FinComerciarUsu(UserIndex)
        Exit Sub

        '[KEVIN]---------------------------------------
        '******************************************************
    Case "FINBAN"
        'User sale del modo BANCO
        UserList(UserIndex).flags.Comerciando = False
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "FINBANOK")
        Exit Sub
        '-------------------------------------------------------

        Exit Sub

    Case "COMUSUOK"
        'Aceptar el cambio
        Call AceptarComercioUsu(UserIndex)
        Exit Sub

    Case "COMUSUNO"

        'Rechazar el cambio
        If UserList(UserIndex).ComUsu.DestUsu > 0 Then
            If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                Call SendData(SendTarget.ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & _
                                                                                       " ha rechazado tu oferta." & FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)

            End If

        End If

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has rechazado la oferta del otro usuario." & FONTTYPE_TALK)
        Call FinComerciarUsu(UserIndex)
        Exit Sub
        '[/Alejo]

    End Select

    Select Case UCase$(Left$(rData, 2))

        '    Case "/Z"
        '        Dim Pos As WorldPos, Pos2 As WorldPos
        '        Dim O As Obj
        '
        '        For LoopC = 1 To 100
        '            Pos = UserList(UserIndex).Pos
        '            O.Amount = 1
        '            O.ObjIndex = iORO
        '            'Exit For
        '            Call TirarOro(100000, UserIndex)
        '            'Call Tilelibre(Pos, Pos2)
        '            'If Pos2.x = 0 Or Pos2.y = 0 Then Exit For
        '
        '            'Call MakeObj(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, O, Pos2.Map, Pos2.x, Pos2.y)
        '        Next LoopC
        '
        '        Exit Sub
    Case "OH"    'Tirar item

        If UserList(UserIndex).flags.Navegando = 1 Or UserList(UserIndex).flags.Muerto = 1 Or (UserList(UserIndex).flags.Privilegios = _
                                                                                               PlayerType.Consejero And Not UserList(UserIndex).flags.EsRolesMaster) Then Exit Sub
        '[Consejeros]

        rData = Right$(rData, Len(rData) - 2)
        Arg1 = ReadField(1, rData, 44)
        Arg2 = ReadField(2, rData, 44)

        'If not in combat mode, can't attack
        If Not UserList(UserIndex).flags.SeguroObjetos Then

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes tirar. Tienes el seguro de objetos activados!!!" & _
                                                            FONTTYPE_FIGHT)
            Exit Sub

        End If

        If val(Arg1) = FLAGORO Then
            If val(Arg2) > 100000 Then
                Arg2 = "100000"    'Don't drop too much gold

            End If

            Call TirarOro(val(Arg2), UserIndex)
            Call EnviarOro(UserIndex)
            Exit Sub

        End If

        If val(Arg1) <= MAX_INVENTORY_SLOTS And val(Arg1) > 0 Then
            If UserList(UserIndex).Invent.Object(val(Arg1)).ObjIndex = 0 Then
                Exit Sub

            End If

            Call DropObj(UserIndex, val(Arg1), val(Arg2), UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y)
        Else
            Exit Sub

        End If

        Exit Sub

    Case "VB"    ' Lanzar hechizo

        'If not in combat mode, can't attack
        'If UserList(UserIndex).Stats.MinSta <= 0 Then
        '  Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estás muy cansado para lanzar hechizos." & FONTTYPE_INFO)
        '  Exit Sub
        ' End If

        'If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then
        ' Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, "NOVER" & UserList(UserIndex).char.CharIndex & ",0," & UserList( _
          UserIndex).PartyIndex)
        'Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z11")
        'UserList(UserIndex).flags.Invisible = 0
        'UserList(UserIndex).flags.Oculto = 0
        'UserList(UserIndex).Counters.Ocultando = 0
        'UserList(UserIndex).Counters.Invisibilidad = 0
        'Call SendData(SendTarget.ToIndex, UserIndex, 0, "INVI0")
        'End If

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub
        End If

        rData = Right$(rData, Len(rData) - 2)
        UserList(UserIndex).flags.Hechizo = val(rData)
        Exit Sub

    Case "LC"    'Click izquierdo
        rData = Right$(rData, Len(rData) - 2)
        Arg1 = ReadField(1, rData, 44)
        Arg2 = ReadField(2, rData, 44)

        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        X = CInt(Arg1)
        Y = CInt(Arg2)

        Call LookatTile(UserIndex, UserList(UserIndex).pos.Map, X, Y)

        If UserList(UserIndex).flags.SeleccioneA <> "" Then
            Dim elotroindex As Byte
            elotroindex = NameIndex(UserList(UserIndex).flags.SeleccioneA)

            If Not InMapBounds(UserList(UserIndex).pos.Map, X, Y) Then Exit Sub

            If elotroindex <= 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
                UserList(UserIndex).flags.SeleccioneA = ""
                Exit Sub

            End If

            Call WarpUserChar(elotroindex, UserList(UserIndex).pos.Map, X, Y, True)
            UserList(elotroindex).flags.EstoySelec = 0
            UserList(UserIndex).flags.SeleccioneA = ""

        End If

        Exit Sub

    Case "RC"    'Click derecho
        rData = Right$(rData, Len(rData) - 2)
        Arg1 = ReadField(1, rData, 44)
        Arg2 = ReadField(2, rData, 44)

        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        X = CInt(Arg1)
        Y = CInt(Arg2)
        Call Accion(UserIndex, UserList(UserIndex).pos.Map, X, Y)
        Exit Sub

    Case "UK"

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub
        End If

        rData = Right$(rData, Len(rData) - 2)

        Select Case val(rData)

        Case Robar
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Robar)

        Case Magia
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Magia)

        Case Domar
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Domar)

        Case Ocultarse

            If UserList(UserIndex).flags.Navegando = 1 Then
                If Not UserList(UserIndex).flags.UltimoMensaje = 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes ocultarte si estas navegando." & FONTTYPE_INFO)
                    UserList(UserIndex).flags.UltimoMensaje = 3
                End If
                Exit Sub
            End If

            If UserList(UserIndex).flags.Oculto = 1 Then
                If Not UserList(UserIndex).flags.UltimoMensaje = 2 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z28")
                    UserList(UserIndex).flags.UltimoMensaje = 2
                End If
                Exit Sub
            End If

            Call DoOcultarse(UserIndex)

        End Select

        Exit Sub

    End Select

    Select Case UCase$(Left$(rData, 3))

    Case "SH+"
        rData = Right$(rData, Len(rData) - 3)
        Call SendData(SendTarget.ToAdmins, 0, 0, "||Servidor> Alta sospecha de SH por parte de " & UserList(UserIndex).Name & " (" & rData & _
                                                 ")" & FONTTYPE_SERVER)
        Exit Sub

    Case "UMH"    ' Usa macro de hechizos
        Call SendData(SendTarget.ToAdmins, UserIndex, 0, "||" & UserList(UserIndex).Name & " fue expulsado por Anti-macro de hechizos " & _
                                                         FONTTYPE_VENENO)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, _
                      "ERR Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros" & FONTTYPE_INFO)
        Call CloseSocket(UserIndex)
        Exit Sub

    Case "HDP"
        UserList(UserIndex).flags.Potea = True
        Exit Sub

    Case "USA"
        rData = Right$(rData, Len(rData) - 3)

        If val(rData) <= MAX_INVENTORY_SLOTS And val(rData) > 0 Then
            If UserList(UserIndex).Invent.Object(val(rData)).ObjIndex = 0 Then Exit Sub
        Else
            Exit Sub

        End If

        Call UseInvItem(UserIndex, val(rData))
        Exit Sub

    Case "CNS"    ' Construye herreria
        rData = Right$(rData, Len(rData) - 3)
        X = CInt(rData)

        If X < 1 Then Exit Sub
        If ObjData(X).SkHerreria = 0 Then Exit Sub
        Call HerreroConstruirItem(UserIndex, X)
        Exit Sub

    Case "CNC"    ' Construye carpinteria
        rData = Right$(rData, Len(rData) - 3)
        X = CInt(ReadField(1, rData, 44))

        If X < 1 Or ObjData(X).SkCarpinteria = 0 Then Exit Sub
        Call CarpinteroConstruirItem(UserIndex, X, ReadField(2, rData, 44))
        Exit Sub

    Case "DEN"
        UserList(UserIndex).flags.YaDenuncio = 0
        Exit Sub

    Case "WLC"    'Click izquierdo en modo trabajo
        rData = Right$(rData, Len(rData) - 3)
        Arg1 = ReadField(1, rData, 44)
        Arg2 = ReadField(2, rData, 44)
        Arg3 = ReadField(3, rData, 44)

        If Arg3 = "" Or Arg2 = "" Or Arg1 = "" Then Exit Sub
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Or Not Numeric(Arg3) Then Exit Sub

        X = CInt(Arg1)
        Y = CInt(Arg2)
        tLong = CInt(Arg3)

        If UserList(UserIndex).flags.Muerto = 1 Or UserList(UserIndex).flags.Descansar Or UserList(UserIndex).flags.Meditando Or Not _
           InMapBounds(UserList(UserIndex).pos.Map, X, Y) Then Exit Sub

        If Not InRangoVision(UserIndex, X, Y) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).pos.X & "," & UserList(UserIndex).pos.Y)
            Exit Sub

        End If

        Select Case tLong

        Case Proyectiles
            Dim TU As Integer, tN As Integer

            'Nos aseguramos que este usando un arma de proyectiles
            If Not IntervaloPermiteAtacar(UserIndex, False) Or Not IntervaloPermiteUsarArcos(UserIndex) Then
                Exit Sub

            End If

            DummyInt = 0

            If UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Then
                DummyInt = 1
            ElseIf UserList(UserIndex).Invent.WeaponEqpSlot < 1 Or UserList(UserIndex).Invent.WeaponEqpSlot > MAX_INVENTORY_SLOTS Then
                DummyInt = 1
            ElseIf UserList(UserIndex).Invent.MunicionEqpSlot < 1 Or UserList(UserIndex).Invent.MunicionEqpSlot > MAX_INVENTORY_SLOTS Then
                DummyInt = 1
            ElseIf UserList(UserIndex).Invent.MunicionEqpObjIndex = 0 Then
                DummyInt = 1
            ElseIf ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Proyectil <> 1 Then
                DummyInt = 2
            ElseIf ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex).ObjType <> eOBJType.otFlechas Then
                DummyInt = 1
            ElseIf UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MunicionEqpSlot).Amount < 1 Then
                DummyInt = 1

            End If

            If DummyInt <> 0 Then
                If DummyInt = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes municiones." & FONTTYPE_INFO)

                End If

                Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                Exit Sub

            End If

            DummyInt = 0

            'Quitamos stamina
            If UserList(UserIndex).Stats.MinSta >= 10 Then
                Call QuitarSta(UserIndex, RandomNumber(1, 10))
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas muy cansado para luchar." & FONTTYPE_INFO)
                Exit Sub

            End If

            Call LookatTile(UserIndex, UserList(UserIndex).pos.Map, Arg1, Arg2)

            TU = UserList(UserIndex).flags.TargetUser
            tN = UserList(UserIndex).flags.TargetNpc

            'Sólo permitimos atacar si el otro nos puede atacar también
            If TU > 0 Then
                If Abs(UserList(TU).pos.Y - UserList(UserIndex).pos.Y) > RANGO_VISION_Y Or _
                   Abs(UserList(TU).pos.X - UserList(UserIndex).pos.X) > RANGO_VISION_X Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos para atacar." & FONTTYPE_WARNING)
                    Exit Sub
                End If

            ElseIf tN > 0 Then

                If Abs(Npclist(tN).pos.Y - UserList(UserIndex).pos.Y) > RANGO_VISION_Y Or _
                   Abs(Npclist(tN).pos.X - UserList(UserIndex).pos.X) > RANGO_VISION_X Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos para atacar." & FONTTYPE_WARNING)
                    Exit Sub
                End If

            End If

            If TU > 0 Then

                'Previene pegarse a uno mismo
                If TU = UserIndex Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z22")
                    DummyInt = 1
                    Exit Sub

                End If

            End If

            If DummyInt = 0 Then
                'Saca 1 flecha
                DummyInt = UserList(UserIndex).Invent.MunicionEqpSlot

                If DummyInt < 1 Or DummyInt > MAX_INVENTORY_SLOTS Then Exit Sub

                If UserList(UserIndex).Invent.Object(DummyInt).Amount > 0 Then

                    ' Sistema ahorro de flechas

                    If UserList(UserIndex).flags.EspecialArco = 1 Then
                        Dim FlechaPorc As String
                        FlechaPorc = RandomNumber(1, 100)

                        If UserList(UserIndex).flags.EspecialObjArco = 1 Then

                            If FlechaPorc <= 33 Then
                                Call QuitarUserInvItem(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot, 0)
                                UserList(UserIndex).flags.EspecialArco = 0
                            Else
                                Call QuitarUserInvItem(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot, 1)
                                UserList(UserIndex).flags.EspecialArco = 0

                            End If

                        End If

                        If UserList(UserIndex).flags.EspecialObjArco = 53 Then

                            If FlechaPorc <= 50 Then
                                Call QuitarUserInvItem(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot, 0)
                                UserList(UserIndex).flags.EspecialArco = 0
                            Else
                                Call QuitarUserInvItem(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot, 1)
                                UserList(UserIndex).flags.EspecialArco = 0

                            End If

                        End If

                        If UserList(UserIndex).flags.EspecialObjArco = 54 Then

                            If FlechaPorc <= 75 Then
                                Call QuitarUserInvItem(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot, 0)
                                UserList(UserIndex).flags.EspecialArco = 0
                            Else
                                Call QuitarUserInvItem(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot, 1)
                                UserList(UserIndex).flags.EspecialArco = 0

                            End If

                        End If

                    Else
                        Call QuitarUserInvItem(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot, 1)

                    End If

                    '[OBJETO NORMALES]
                    UserList(UserIndex).Invent.Object(DummyInt).Equipped = 1
                    UserList(UserIndex).Invent.MunicionEqpSlot = DummyInt
                    UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(DummyInt).ObjIndex
                    Call UpdateUserInv(False, UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)

                End If

            Else
                'Call UpdateUserInv(False, UserIndex, DummyInt)
                'UserList(UserIndex).Invent.MunicionEqpSlot = 0
                'UserList(UserIndex).Invent.MunicionEqpObjIndex = 0

            End If

            '-----------------------------------

            If tN > 0 Then
                If Npclist(tN).Attackable <> 0 Then
                    Call UsuarioAtacaNpc(UserIndex, tN)

                End If

            ElseIf TU > 0 Then

                If UserList(UserIndex).flags.Seguro Then
                    If Not Criminal(TU) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Para atacar ciudadanos desactiva el seguro!" & FONTTYPE_FIGHT)
                        Exit Sub

                    End If

                End If

                Call UsuarioAtacaUsuario(UserIndex, TU)

            End If

        Case Magia
            If MapInfo(UserList(UserIndex).pos.Map).MagiaSinEfecto > 0 And Not MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Trigger = 2 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Una fuerza oscura te impide canalizar tu energía." & FONTTYPE_Motd4)
                Exit Sub
            End If

            'If it's outside range log it and exit
            If Abs(UserList(UserIndex).pos.X - X) > RANGO_VISION_X Or Abs(UserList(UserIndex).pos.Y - Y) > RANGO_VISION_Y Then
                Call LogCheating("Ataque fuera de rango de " & UserList(UserIndex).Name & "(" & UserList(UserIndex).pos.Map & _
                                 "/" & UserList(UserIndex).pos.X & "/" & UserList(UserIndex).pos.Y & ") ip: " & UserList(UserIndex).ip & _
                               " a la posicion (" & UserList(UserIndex).pos.Map & "/" & X & "/" & Y & ")")
                Exit Sub
            End If

            Call LookatTile(UserIndex, UserList(UserIndex).pos.Map, X, Y)

            'MmMmMmmmmM
            Dim wp2 As WorldPos
            wp2.Map = UserList(UserIndex).pos.Map
            wp2.X = X
            wp2.Y = Y

            If UserList(UserIndex).flags.Hechizo > 0 Then
                If IntervaloPermiteLanzarSpell(UserIndex) Then
                    Call LanzarHechizo(UserList(UserIndex).flags.Hechizo, UserIndex)
                    'UserList(UserIndex).flags.PuedeLanzarSpell = 0
                    UserList(UserIndex).flags.Hechizo = 0
                End If

            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Primero selecciona el hechizo que queres lanzar y después lanzá!" & FONTTYPE_INFO)

            End If


            'If Distancia(UserList(UserIndex).Pos, wp2) > 10 Then
            If (Abs(UserList(UserIndex).pos.X - wp2.X) > 9 Or Abs(UserList(UserIndex).pos.Y - wp2.Y) > 8) Then
                Dim txt As String
                txt = "Ataque fuera de rango de " & UserList(UserIndex).Name & "(" & UserList(UserIndex).pos.Map & "/" & UserList( _
                      UserIndex).pos.X & "/" & UserList(UserIndex).pos.Y & ") ip: " & UserList(UserIndex).ip & " a la posicion (" & _
                      wp2.Map & "/" & wp2.X & "/" & wp2.Y & ") "

                If UserList(UserIndex).flags.Hechizo > 0 Then
                    txt = txt & ". Hechizo: " & Hechizos(UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)).nombre

                End If

                If MapData(wp2.Map, wp2.X, wp2.Y).UserIndex > 0 Then
                    txt = txt & " hacia el usuario: " & UserList(MapData(wp2.Map, wp2.X, wp2.Y).UserIndex).Name
                ElseIf MapData(wp2.Map, wp2.X, wp2.Y).NpcIndex > 0 Then
                    txt = txt & " hacia el NPC: " & Npclist(MapData(wp2.Map, wp2.X, wp2.Y).NpcIndex).Name

                End If

                Call LogCheating(txt)

            End If

        Case Pesca

            AuxInd = UserList(UserIndex).Invent.HerramientaEqpObjIndex

            If AuxInd = 0 Then Exit Sub

            'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
            If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

            If AuxInd <> CAÑA_PESCA And AuxInd <> RED_PESCA Then
                'Call Cerrar_Usuario(UserIndex)
                ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                Exit Sub

            End If

            'Basado en la idea de Barrin
            'Comentario por Barrin: jah, "basado", caradura ! ^^
            If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Trigger = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes pescar desde donde te encuentras." & FONTTYPE_INFO)
                Exit Sub

            End If

            If HayAgua(UserList(UserIndex).pos.Map, X, Y) Then
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_PESCAR)

                Select Case AuxInd

                Case CAÑA_PESCA

                    Call DoPescar(UserIndex)

                Case RED_PESCA

                    With UserList(UserIndex)
                        wpaux.Map = .pos.Map
                        wpaux.X = X
                        wpaux.Y = Y
                    End With

                    If UserList(UserIndex).flags.Navegando = 0 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Para utilizar la red de pesca es necesario estar navegando." & FONTTYPE_CYAN)
                        Exit Sub
                    End If

                    If Distancia(UserList(UserIndex).pos, wpaux) > 6 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estás demasiado lejos para pescar." & FONTTYPE_INFO)
                        Exit Sub
                    End If

                    Call DoPescarRed(UserIndex)

                End Select

            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay agua donde pescar busca un lago, rio o mar." & FONTTYPE_INFO)

            End If

        Case Robar

            If MapInfo(UserList(UserIndex).pos.Map).Pk Then

                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

                Call LookatTile(UserIndex, UserList(UserIndex).pos.Map, X, Y)

                If UserList(UserIndex).flags.TargetUser > 0 And UserList(UserIndex).flags.TargetUser <> UserIndex Then

                    If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 0 Then
                        wpaux.Map = UserList(UserIndex).pos.Map
                        wpaux.X = val(ReadField(1, rData, 44))
                        wpaux.Y = val(ReadField(2, rData, 44))

                        If Distancia(wpaux, UserList(UserIndex).pos) > 2 Then
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                            Exit Sub
                        End If

                        '17/09/02
                        'No aseguramos que el trigger le permite robar
                        If MapData(UserList(UserList(UserIndex).flags.TargetUser).pos.Map, UserList(UserList( _
                                                                                                    UserIndex).flags.TargetUser).pos.X, UserList(UserList(UserIndex).flags.TargetUser).pos.Y).Trigger = _
                                                                                                    eTrigger.ZONASEGURA Then
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes robar aquí." & FONTTYPE_WARNING)
                            Exit Sub
                        End If

                        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Trigger = _
                           eTrigger.ZONASEGURA Then
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes robar aquí." & FONTTYPE_WARNING)
                            Exit Sub
                        End If

                        Call DoRobar(UserIndex, UserList(UserIndex).flags.TargetUser)

                    End If

                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay nadie para robarle!." & FONTTYPE_WARNING)

                End If

            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No puedes robarle en zonas seguras!." & FONTTYPE_WARNING)

            End If

        Case Sastreria
            If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Deberías equiparte la tijera." & FONTTYPE_INFO)
                Exit Sub
            End If

            AuxInd = MapData(UserList(UserIndex).pos.Map, X, Y).OBJInfo.ObjIndex

            If AuxInd > 0 Then

                If ObjData(AuxInd).ObjType = eOBJType.otOveja Then
                    Call DoOveja(UserIndex)
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ninguna oveja alli." & FONTTYPE_GUERRA)
                End If

            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ninguna oveja alli." & FONTTYPE_GUERRA)
            End If


        Case Recolectar
            'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
            If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Deberías equiparte el hoz de mano." & FONTTYPE_INFO)
                Exit Sub
            End If

            AuxInd = MapData(UserList(UserIndex).pos.Map, X, Y).OBJInfo.ObjIndex

            If AuxInd > 0 Then
                wpaux.Map = UserList(UserIndex).pos.Map
                wpaux.X = X
                wpaux.Y = Y

                If Distancia(wpaux, UserList(UserIndex).pos) > 2 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                    Exit Sub
                End If

                '¿Hay un arbol donde clickeo?
                If ObjData(AuxInd).ObjType = eOBJType.otArboles Then

                    Select Case UserList(UserIndex).pos.Map
                    Case 1, 20, 34, 37, 59, 60, 61, 62, 63, 64, 84, 86, 95, 132, 149
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes coger hierbas en las ciudades." & FONTTYPE_INFO)
                        Exit Sub
                    End Select

                    If MapInfo(UserList(UserIndex).pos.Map).Pk = False And Not UserList(UserIndex).pos.Map = 47 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes coger hierbas en zona seguras." & FONTTYPE_INFO)
                        Exit Sub
                    End If

                    If Distancia(wpaux, UserList(UserIndex).pos) = 0 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ningun arbol ahi." & FONTTYPE_GUERRA)
                        Exit Sub
                    End If

                    Call SendData(SendTarget.ToPCArea, CInt(UserIndex), UserList(UserIndex).pos.Map, "TW60")

                    Call DoTalarHierba(UserIndex)

                End If

            Else

                Select Case UserList(UserIndex).pos.Map
                Case 1, 20, 34, 37, 59, 60, 61, 62, 63, 64, 84, 86, 95, 132, 149
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes coger hierbas en las ciudades." & FONTTYPE_INFO)
                    Exit Sub
                End Select

                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ningun arbol ahi." & FONTTYPE_GUERRA)

            End If


        Case talar

            'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
            If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Deberías equiparte el hacha." & FONTTYPE_INFO)
                Exit Sub
            End If

            AuxInd = MapData(UserList(UserIndex).pos.Map, X, Y).OBJInfo.ObjIndex

            If AuxInd > 0 Then
                wpaux.Map = UserList(UserIndex).pos.Map
                wpaux.X = X
                wpaux.Y = Y

                If Distancia(wpaux, UserList(UserIndex).pos) > 2 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                    Exit Sub
                End If

                '¿Hay un arbol donde clickeo?
                If ObjData(AuxInd).ObjType = eOBJType.otArboles Then

                    Select Case UserList(UserIndex).pos.Map
                    Case 1, 20, 34, 37, 59, 60, 61, 62, 63, 64, 84, 86, 95, 132, 149
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes talar en las ciudades." & FONTTYPE_INFO)
                        Exit Sub
                    End Select

                    If MapInfo(UserList(UserIndex).pos.Map).Pk = False And Not UserList(UserIndex).pos.Map = 47 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes talar en zona seguras." & FONTTYPE_INFO)
                        Exit Sub
                    End If

                    If Distancia(wpaux, UserList(UserIndex).pos) = 0 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ningun arbol ahi." & FONTTYPE_GUERRA)
                        Exit Sub
                    End If

                    Call SendData(SendTarget.ToPCArea, CInt(UserIndex), UserList(UserIndex).pos.Map, "TW" & SND_TALAR)

                    Call DoTalar(UserIndex)

                End If

            Else

                Select Case UserList(UserIndex).pos.Map
                Case 1, 20, 34, 37, 59, 60, 61, 62, 63, 64, 84, 86, 95, 132, 149
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes talar en las ciudades." & FONTTYPE_INFO)
                    Exit Sub
                End Select

                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ningun arbol ahi." & FONTTYPE_GUERRA)

            End If

        Case Mineria

            'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
            If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then Exit Sub

            If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> PIQUETE_MINERO Then
                ' Call Cerrar_Usuario(UserIndex)
                ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                Exit Sub

            End If

            Call LookatTile(UserIndex, UserList(UserIndex).pos.Map, X, Y)

            AuxInd = MapData(UserList(UserIndex).pos.Map, X, Y).OBJInfo.ObjIndex

            If AuxInd > 0 Then
                wpaux.Map = UserList(UserIndex).pos.Map
                wpaux.X = X
                wpaux.Y = Y

                If Distancia(wpaux, UserList(UserIndex).pos) > 2 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                    Exit Sub

                End If

                '¿Hay un yacimiento donde clickeo?
                If ObjData(AuxInd).ObjType = eOBJType.otYacimiento Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_MINERO)
                    Call DoMineria(UserIndex)
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahi no hay ningun yacimiento." & FONTTYPE_INFO)

                End If

            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahi no hay ningun yacimiento." & FONTTYPE_INFO)

            End If

        Case Domar
            'Modificado 25/11/02
            'Optimizado y solucionado el bug de la doma de
            'criaturas hostiles.
            Dim CI As Integer

            Call LookatTile(UserIndex, UserList(UserIndex).pos.Map, X, Y)
            CI = UserList(UserIndex).flags.TargetNpc

            If CI > 0 Then
                If Npclist(CI).flags.Domable > 0 Then
                    wpaux.Map = UserList(UserIndex).pos.Map
                    wpaux.X = X
                    wpaux.Y = Y

                    If Distancia(wpaux, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 2 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                        Exit Sub

                    End If

                    If Npclist(CI).flags.AttackedBy <> "" Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes domar una criatura que está luchando con un jugador." & _
                                                                        FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Call DoDomar(UserIndex, CI)
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes domar a esa criatura." & FONTTYPE_INFO)

                End If

            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡ No hay ninguna criatura alli !" & FONTTYPE_INFO)

            End If

        Case FundirMetal

            'Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
            '   If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

            If UserList(UserIndex).flags.TargetObj > 0 Then
                If ObjData(UserList(UserIndex).flags.TargetObj).ObjType = eOBJType.otFragua Then

                    ''chequeamos que no se zarpe duplicando oro
                    If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex <> UserList( _
                       UserIndex).flags.TargetObjInvIndex Then

                        If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex = 0 Or UserList( _
                           UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount = 0 Then
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes mas minerales" & FONTTYPE_INFO)
                            Exit Sub

                        End If

                        ''FUISTE
                        'Call Ban(UserList(UserIndex).Name, "Sistema anti cheats", "Intento de duplicacion de items")
                        'Call LogCheating(UserList(UserIndex).Name & " intento crear minerales a partir de otros: FlagSlot/usaba/usoconclick/cantidad/IP:" & UserList(UserIndex).flags.TargetObjInvSlot & "/" & UserList(UserIndex).flags.TargetObjInvIndex & "/" & UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex & "/" & UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount & "/" & UserList(UserIndex).ip)
                        'UserList(UserIndex).flags.Ban = 1
                        'Call SendData(SendTarget.ToAll, 0, 0, "||>>>> El sistema anti-cheats baneó a " & UserList(UserIndex).Name & " (intento de duplicación). Ip Logged. " & FONTTYPE_FIGHT)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRHas sido expulsado por el sistema anti cheats. Reconéctate.")
                        Call CloseSocket(UserIndex)
                        Exit Sub

                    End If

                    Call FundirMineral(UserIndex)
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahi no hay ninguna fragua." & FONTTYPE_INFO)

                End If

            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahi no hay ninguna fragua." & FONTTYPE_INFO)

            End If

        Case Herrero
            Call LookatTile(UserIndex, UserList(UserIndex).pos.Map, X, Y)

            If UserList(UserIndex).flags.TargetObj > 0 Then
                If ObjData(UserList(UserIndex).flags.TargetObj).ObjType = eOBJType.otYunque Then
                    Call EnviarArmasMagicasConstruibles(UserIndex)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ABHM")
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahi no hay ningun yunque." & FONTTYPE_INFO)

                End If

            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahi no hay ningun yunque." & FONTTYPE_INFO)

            End If


        Case Herreria
            Call LookatTile(UserIndex, UserList(UserIndex).pos.Map, X, Y)

            If UserList(UserIndex).flags.TargetObj > 0 Then
                If ObjData(UserList(UserIndex).flags.TargetObj).ObjType = eOBJType.otYunque Then
                    Call EnivarArmasConstruibles(UserIndex)
                    Call EnivarArmadurasConstruibles(UserIndex)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "SFH")
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahi no hay ningun yunque." & FONTTYPE_INFO)

                End If

            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahi no hay ningun yunque." & FONTTYPE_INFO)

            End If

        End Select

        'UserList(UserIndex).flags.PuedeTrabajar = 0
        Exit Sub

    Case "CIG"
        rData = Right$(rData, Len(rData) - 3)

        If modGuilds.CrearNuevoClan(rData, UserIndex, tStr) Then

        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)

        End If

        Exit Sub

    Case "DRA"  ' Drag en inventario
        rData = Right$(rData, Len(rData) - 3)
        X = ReadField(1, rData, 44)    ' Old Slot
        Y = ReadField(2, rData, 44)    ' New Slot
        Call moveItem(UserIndex, X, Y)
        Exit Sub

    Case "DRO"  ' Drop en Mapa
        rData = Right$(rData, Len(rData) - 3)
        tInt = ReadField(1, rData, 44)  ' Slot
        X = ReadField(2, rData, 44)     ' Pos X
        Y = ReadField(3, rData, 44)     ' Pos Y
        n = ReadField(4, rData, 44)     ' Cantidad

        Mapa = UserList(UserIndex).pos.Map

        If InMapBounds(Mapa, X, Y) Then

            'Desequipa
            If n > 0 Or n <= MAX_INVENTORY_OBJS Then

                Dim tUser As Integer
                Dim tNpc As Integer
                tUser = MapData(Mapa, X, Y).UserIndex
                tNpc = MapData(Mapa, X, Y).NpcIndex

                If tNpc > 0 Then
                    Call DragToNPC(UserIndex, tNpc, tInt, n)
                ElseIf tUser > 0 Then
                    Call DragToUser(UserIndex, tUser, tInt, n)
                Else
                    Call DragToPos(UserIndex, X, Y, tInt, n)

                End If

            End If

        End If

        Exit Sub

    End Select

    Select Case UCase$(Left$(rData, 4))
      
      Case "HBNF"
         rData = Right$(rData, Len(rData) - 4)
         Call FinalizaHablarQuest(UserIndex, UserList(UserIndex).Quest.Quest)
      Exit Sub
     
     Case "HPGM"
        rData = Right$(rData, Len(rData) - 4)
        CountTC = rData
        Call SendData(SendTarget.ToAll, 0, 0, "HUCT" & rData)
        Call DayChange(rData)
        Exit Sub
     
        'CHOTS | Paquetes de Procesos
    Case "PCGF"
        Dim proceso As String
        rData = Right$(rData, Len(rData) - 4)
        proceso = ReadField(1, rData, 44)
        TIndex = ReadField(2, rData, 44)
        Call SendData(SendTarget.ToIndex, TIndex, 0, "PCGN" & proceso & "," & UserList(UserIndex).Name)
        Exit Sub

    Case "PCWC"
        Dim proseso As String
        rData = Right$(rData, Len(rData) - 4)
        proseso = ReadField(1, rData, 44)
        TIndex = ReadField(2, rData, 44)
        Call SendData(SendTarget.ToIndex, TIndex, 0, "PCSS" & proseso & "," & UserList(UserIndex).Name)
        Exit Sub

    Case "PCCC"
        Dim caption As String
        rData = Right$(rData, Len(rData) - 4)
        caption = ReadField(1, rData, 44)
        TIndex = ReadField(2, rData, 44)
        Call SendData(SendTarget.ToIndex, TIndex, 0, "PCCC" & caption & "," & UserList(UserIndex).Name)
        Exit Sub
        'CHOTS | Paquetes de Procesos

    Case "LEFT"    '[rodra]
        rData = Right$(rData, Len(rData) - 4)
        TIndex = ReadField(1, rData, 32)
        rData = ReadField(2, rData, 32)
        Call SendData(SendTarget.ToIndex, TIndex, 0, "||" & UCase$(UserList(UserIndex).Name) & _
                                                   " : Hola!, se supone que no tengo cliente externo, no? " & FONTTYPE_CONSEJO)
        '[Rodra]
        Exit Sub

    Case "CTMR"
        rData = Right$(rData, Len(rData) - 4)
        TIndex = NameIndex(ReadField(1, rData, 2))

        If TIndex <= 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario se encuentra offline." & FONTTYPE_INFO)
            Exit Sub

        End If

        UserList(TIndex).Respuesta = ReadField(2, rData, 2)
        Call WriteVar(App.Path & "\Charfile\" & UserList(TIndex).Name & ".chr", "INIT", "Respuesta", UserList(TIndex).Respuesta)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Respuesta enviada a " & UserList(TIndex).Name & FONTTYPE_INFO)
        Exit Sub

    Case "INFS"    'Informacion del hechizo
        rData = Right$(rData, Len(rData) - 4)

        If val(rData) > 0 And val(rData) < MAXUSERHECHIZOS + 1 Then
            Dim h As Integer
            h = UserList(UserIndex).Stats.UserHechizos(val(rData))

            If h > 0 And h < NumeroHechizos + 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & FONTTYPE_INFO)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Nombre:" & Hechizos(h).nombre & FONTTYPE_INFO)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Descripcion:" & Hechizos(h).Desc & FONTTYPE_INFO)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Skill requerido: " & Hechizos(h).MinSkill & " de magia." & FONTTYPE_INFO)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Mana necesario: " & Hechizos(h).ManaRequerido & FONTTYPE_INFO)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Stamina necesaria: " & Hechizos(h).StaRequerido & FONTTYPE_INFO)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%" & FONTTYPE_INFO)
            End If

        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Primero selecciona el hechizo.!" & FONTTYPE_INFO)

        End If

        Exit Sub

    Case "EQUI"

        If UserList(UserIndex).flags.Montado = True Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Debes Demontarte para poder equiparte!.!" & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub

        End If

        rData = Right$(rData, Len(rData) - 4)

        If val(rData) <= MAX_INVENTORY_SLOTS And val(rData) > 0 Then
            If UserList(UserIndex).Invent.Object(val(rData)).ObjIndex = 0 Then Exit Sub
        Else
            Exit Sub

        End If

        Call EquiparInvItem(UserIndex, val(rData))
        Exit Sub

    Case "CHEA"    'Cambiar Heading ;-)
        rData = Right$(rData, Len(rData) - 4)

        If val(rData) > 0 And val(rData) < 5 Then
            UserList(UserIndex).char.heading = rData
            '[MaTeO 9]
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                            UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                         UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

            '[/MaTeO 9]
        End If

        Exit Sub

    Case "SKSE"    'Modificar skills
        Dim sumatoria As Integer
        Dim incremento As Integer
        rData = Right$(rData, Len(rData) - 4)

        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NUMSKILLS
            incremento = val(ReadField(i, rData, 44))

            If incremento < 0 Then
                'Call SendData(SendTarget.ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
                UserList(UserIndex).Stats.SkillPts = 0
                Call CloseSocket(UserIndex)
                Exit Sub

            End If

            sumatoria = sumatoria + incremento
        Next i

        If sumatoria > UserList(UserIndex).Stats.SkillPts Then
            'UserList(UserIndex).Flags.AdministrativeBan = 1
            'Call SendData(SendTarget.ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
            Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

        For i = 1 To NUMSKILLS
            incremento = val(ReadField(i, rData, 44))
            UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts - incremento
            UserList(UserIndex).Stats.UserSkills(i) = UserList(UserIndex).Stats.UserSkills(i) + incremento

            If UserList(UserIndex).Stats.UserSkills(i) > 100 Then UserList(UserIndex).Stats.UserSkills(i) = 100
        Next i

        Exit Sub

    Case "ENTR"    'Entrena hombre!

        If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub

        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 3 Then Exit Sub

        rData = Right$(rData, Len(rData) - 4)

        If Npclist(UserList(UserIndex).flags.TargetNpc).Mascotas < MAXMASCOTASENTRENADOR Then
            If val(rData) > 0 And val(rData) < Npclist(UserList(UserIndex).flags.TargetNpc).NroCriaturas + 1 Then
                Dim SpawnedNpc As Integer
                SpawnedNpc = SpawnNpc(Npclist(UserList(UserIndex).flags.TargetNpc).Criaturas(val(rData)).NpcIndex, Npclist(UserList( _
                                                                                                                           UserIndex).flags.TargetNpc).pos, True, False)

                If SpawnedNpc > 0 Then
                    Npclist(SpawnedNpc).MaestroNpc = UserList(UserIndex).flags.TargetNpc
                    Npclist(UserList(UserIndex).flags.TargetNpc).Mascotas = Npclist(UserList(UserIndex).flags.TargetNpc).Mascotas + 1

                End If

            End If

        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbWhite & "°" & _
                                                                                       "No puedo traer mas criaturas, mata las existentes!" & "°" & str(Npclist(UserList( _
                                                                                                                                                                UserIndex).flags.TargetNpc).char.CharIndex))

        End If

        Exit Sub

    Case "COMP"

        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub
        End If

        '¿El target es un NPC valido?
        If UserList(UserIndex).flags.TargetNpc > 0 Then

            '¿El NPC puede comerciar?
            If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & FONTTYPE_TALK & "°" & _
                                                                                           "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                Exit Sub

            End If

        Else
            Exit Sub

        End If

        rData = Right$(rData, Len(rData) - 5)

        'User compra el item del slot rdata
        If UserList(UserIndex).flags.Comerciando = False Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No estas comerciando " & FONTTYPE_INFO)
            Exit Sub
        End If

        'listindex+1, cantidad
        Call NPCVentaItem(UserIndex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)), UserList(UserIndex).flags.TargetNpc)
        Exit Sub

        '[KEVIN]*********************************************************************
        '------------------------------------------------------------------------------------

    Case "COAC"

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub
        End If

        If UserList(UserIndex).flags.TargetNpc > 0 Then
            If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & FONTTYPE_TALK & "°" & _
                                                                                           "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If
        Else
            Exit Sub
        End If

        rData = Right$(rData, Len(rData) - 5)

        If UserList(UserIndex).flags.Comerciando = False Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No estas comerciando " & FONTTYPE_INFO)
            Exit Sub
        End If

        Call NpcVentaCredito(UserIndex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)), UserList(UserIndex).flags.TargetNpc)

        Exit Sub

    Case "COAJ"

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub
        End If

        If UserList(UserIndex).flags.TargetNpc > 0 Then
            If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & FONTTYPE_TALK & "°" & _
                                                                                           "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If
        Else
            Exit Sub
        End If

        rData = Right$(rData, Len(rData) - 5)

        If UserList(UserIndex).flags.Comerciando = False Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No estas comerciando " & FONTTYPE_INFO)
            Exit Sub
        End If

        Call NpcVentaCanjes(UserIndex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)), UserList(UserIndex).flags.TargetNpc)

        Exit Sub

    Case "RETI"

        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub

        End If

        '¿El target es un NPC valido?
        If UserList(UserIndex).flags.TargetNpc > 0 Then

            '¿Es el banquero?
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 4 Then
                Exit Sub

            End If

        Else
            Exit Sub

        End If

        rData = Right(rData, Len(rData) - 5)
        'User retira el item del slot rdata
        Call UserRetiraItem(UserIndex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
        Exit Sub

        '-----------------------------------------------------------------------------------
        '[/KEVIN]****************************************************************************

    Case "VEAJ"
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub
        End If

        rData = Right$(rData, Len(rData) - 5)
        '¿El target es un NPC valido?
        tInt = val(ReadField(1, rData, 44))

        If UserList(UserIndex).flags.TargetNpc > 0 Then
            '¿El NPC puede comerciar?
            If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & FONTTYPE_TALK & "°" & _
                                                                                           "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If
        Else
            Exit Sub
        End If

        Call NPCCompraCanjes(UserIndex, ReadField(1, rData, 44), ReadField(2, rData, 44))

        Exit Sub

    Case "VEND"

        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub
        End If

        rData = Right$(rData, Len(rData) - 5)
        '¿El target es un NPC valido?
        tInt = val(ReadField(1, rData, 44))

        If UserList(UserIndex).flags.TargetNpc > 0 Then

            '¿El NPC puede comerciar?
            If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & FONTTYPE_TALK & "°" & _
                                                                                           "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                Exit Sub

            End If

        Else
            Exit Sub

        End If

        '           rdata = Right$(rdata, Len(rdata) - 5)
        'User compra el item del slot rdata
        Call NPCCompraItem(UserIndex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
        Exit Sub

        '[KEVIN]-------------------------------------------------------------------------
        '****************************************************************************************
    Case "DEPO"

        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub

        End If

        '¿El target es un NPC valido?
        If UserList(UserIndex).flags.TargetNpc > 0 Then

            '¿El NPC puede comerciar?
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> eNPCType.Banquero Then
                Exit Sub

            End If

        Else
            Exit Sub

        End If

        rData = Right(rData, Len(rData) - 5)
        'User deposita el item del slot rdata
        Call UserDepositaItem(UserIndex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
        Exit Sub

        '****************************************************************************************
        '[/KEVIN]---------------------------------------------------------------------------------
    End Select

    Select Case UCase$(Left$(rData, 5))

        '#################### LISTA DE AMIGOS  ######################
    Case "NEWFF"
        rData = Right$(rData, Len(rData) - 5)
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "CANTIDAD", "Cant", 9)
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo0", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo1", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo2", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo3", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo4", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo5", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo6", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo7", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo8", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo9", "(Slot Vacio)")
        Exit Sub

    Case "ADDFF"  ' AGREGAR AMIGO
        Dim PathAmigos As String
        Dim Amigo As String
        Dim numFF As Integer
        Dim Amiguito1 As Integer
        rData = Right$(rData, Len(rData) - 5)
        Amigo = ReadField(2, rData, 64)
        Amiguito1 = ReadField(3, rData, 64)
        PathAmigos = App.Path & "\ListadeAmigos\" & UserList(UserIndex).Name & ".log"
        'If Not FileExist(PathAmigos, vbNormal) Then
        '   Call WriteVar(PathAmigos, "CANTIDAD", "Cant", 1)
        '   Call WriteVar(PathAmigos, "AMIGOS", "Amigo1", Amigo)
        '   Call SendData(SendTarget.toindex, UserIndex, 0, "||Has agregado a " & Amigo & " a tu Lista." & FONTTYPE_INFO)
        'Else
        numFF = GetVar(PathAmigos, "CANTIDAD", "Cant")
        'Call WriteVar(PathAmigos, "CANTIDAD", "Cant", numFF + 1)
        TIndex = NameIndex(Amigo)

        If UserList(TIndex).flags.Privilegios > 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes agregar a administradores." & FONTTYPE_INFO)
            Exit Sub

        End If

        If TIndex <= 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Else
            Call WriteVar(PathAmigos, "AMIGOS", "Amigo" & Amiguito1, Amigo)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has agregado a " & Amigo & " a tu Lista." & FONTTYPE_INFO)

        End If

        Exit Sub

    Case "DELFF"    ' BORRAR AMIGO
        rData = Right$(rData, Len(rData) - 5)

        Amigo = ReadField(2, rData, 64)
        Amiguito1 = ReadField(3, rData, 64)
        PathAmigos = App.Path & "\ListadeAmigos\" & UserList(UserIndex).Name & ".log"

        Call WriteVar(PathAmigos, "AMIGOS", "Amigo" & Amiguito1, "(Slot Vacio)")
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has eliminado a " & Amigo & " de tu Lista." & FONTTYPE_INFO)
        Exit Sub

    Case "LISFF"    ' VER AMIGOS
        rData = Right$(rData, Len(rData) - 5)
        'Dim PathAmigos As String
        Dim Amigos1 As Integer
        Dim FF1 As String
        PathAmigos = App.Path & "/ListadeAmigos/" & rData & ".log"
        Amigos1 = val(GetVar(PathAmigos, "CANTIDAD", "Cant"))
        Dim FFAmigos As Integer

        For FFAmigos = 0 To Amigos1
            FF1 = (GetVar(PathAmigos, "AMIGOS", "Amigo" & FFAmigos))
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "FFLI" & FF1)
        Next
        Exit Sub

    Case "ESTFF"    ' ESTADO AMIGOS
        Dim EstadoFF As Integer
        'Dim Amigo As String
        rData = Right$(rData, Len(rData) - 5)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ESOF")
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ESON")

        End If

        Exit Sub

    Case "DEMSG"

        If UserList(UserIndex).flags.TargetObj > 0 Then
            rData = Right$(rData, Len(rData) - 5)
            Dim f As String, Titu As String, Msg As String, f2 As String
            f = App.Path & "\foros\"
            f = f & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & ".for"
            Titu = ReadField(1, rData, 176)
            Msg = ReadField(2, rData, 176)
            Dim n2 As Integer, loopme As Integer

            If FileExist(f, vbNormal) Then
                Dim num As Integer
                num = val(GetVar(f, "INFO", "CantMSG"))

                If num > MAX_MENSAJES_FORO Then

                    For loopme = 1 To num
                        Kill App.Path & "\foros\" & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & loopme & ".for"
                    Next
                    Kill App.Path & "\foros\" & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & ".for"
                    num = 0

                End If

                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & num + 1 & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, Msg
                Call WriteVar(f, "INFO", "CantMSG", num + 1)
            Else
                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & "1" & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, Msg
                Call WriteVar(f, "INFO", "CantMSG", 1)

            End If

            Close #n2

        End If

        Exit Sub

    End Select

    Select Case UCase$(Left$(rData, 6))

        'CRAW; 18/09/2019
    Case "MARAKO"
        Call getServerDelay(UserIndex)
        Exit Sub

    Case "DESPHE"    'Mover Hechizo de lugar
        rData = Right(rData, Len(rData) - 6)

        If UserList(UserIndex).flags.SeguroHechizos Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Seguro hechizos activado, para desactivar apreta la tecla ''*''.!!" & _
                                                            FONTTYPE_Motd4)
        Else
            Call MoverHechizo(UserIndex, CInt(ReadField(1, rData, 44)), CInt(ReadField(2, rData, 44)))
        End If


        Exit Sub

    Case "DESCOD"    'Informacion del hechizo
        rData = Right$(rData, Len(rData) - 6)
        Call modGuilds.ActualizarCodexYDesc(rData, UserList(UserIndex).GuildIndex)
        Exit Sub
    
    Case "ANTISH"
         Call SendData(ToAdmins, 0, 0, "||AntiSH> El usuario " & UserList(UserIndex).Name & " ha intentado usar SpeedHack." & FONTTYPE_Motd4)
     Exit Sub
     
     Case "ANTICH"
         rData = Right$(rData, Len(rData) - 6)
           Call SendData(ToAdmins, 0, 0, "||Cheats> El usuario " & UserList(UserIndex).Name & " ha intentado usar " & rData & "." & FONTTYPE_Motd4)
     Exit Sub
    
    End Select

    '[Alejo]
    Select Case UCase$(Left$(rData, 7))

    Case "SACSAC1"
            rData = Right(rData, Len(rData) - 7)
            h = FreeFile
            Open App.Path & "\LOGS\CHEATERS.log" For Append Shared As h
            
            Print #h, "########################################################################"
            Print #h, "Usuario: " & UserList(UserIndex).Name
            Print #h, "Fecha: " & Date
            Print #h, "Hora: " & Time
            Select Case val(rData)
                Case 1
                    Print #h, "CHEAT: Se detectaron intervalos parecidos en los clicks en lista de hechizos y boton lanzar, posible macro o mouse gamer"
                    
                Case 2
                    Print #h, "CHEAT: Se detectaron intervalos parecidos en los clicks boton inventario y click dentro del inventario, posible macro o mouse gamer"
                
                Case 3
                    Print #h, "CHEAT: Se detectaron intervalos parecidos en los clicks en inventario y boton hechizos, posible macro o mouse gamer"
                    
                Case 4
                    Print #h, "CHEAT: Se detectaron posiciones iguales en clicks en el boton lanzar, posible macro o mouse gamer"
                
                Case 5
                    Print #h, "CHEAT: Se detectaron posiciones iguales en clicks en el boton hechizos, posible macro o mouse gamer"
                    
                Case 6
                    Print #h, "CHEAT: Se detectaron posiciones iguales en clicks en el boton inventario, posible macro o mouse gamer"
                    
            End Select
            Print #h, "########################################################################"
            Print #h, " "
            Close #h
            
            'UserList(UserIndex).flags.Ban = 1
        
            'Avisamos a los admins
            'Call SendData(SendTarget.ToAdmins, 0, 0, "||Sistema Antichit> " & UserList(UserIndex).Name & " ha sido Echado por uso de " & rData & _
                    FONTTYPE_SERVER)
            'Call CloseSocket(UserIndex)
            Exit Sub

    Case "OFRECER"
        rData = Right$(rData, Len(rData) - 7)
        Arg1 = ReadField(1, rData, Asc(","))
        Arg2 = ReadField(2, rData, Asc(","))

        If val(Arg1) <= 0 Or val(Arg2) <= 0 Then
            Exit Sub

        End If

        If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged = False Then
            'sigue vivo el usuario ?
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        Else

            'esta vivo ?
            If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.Muerto = 1 Then
                Call FinComerciarUsu(UserIndex)
                Exit Sub

            End If

            '//Tiene la cantidad que ofrece ??//'
            If val(Arg1) = FLAGORO Then

                'oro
                If val(Arg2) > UserList(UserIndex).Stats.GLD Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
                    Exit Sub

                End If

            Else

                'inventario
                If val(Arg2) > UserList(UserIndex).Invent.Object(val(Arg1)).Amount Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
                    Exit Sub

                End If

            End If

            If UserList(UserIndex).ComUsu.Objeto > 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes cambiar tu oferta." & FONTTYPE_TALK)
                Exit Sub

            End If

            'No permitimos vender barcos mientras están equipados (no podés desequiparlos y causa errores)
            If UserList(UserIndex).flags.Navegando = 1 Then
                If UserList(UserIndex).Invent.BarcoSlot = val(Arg1) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podés vender tu barco mientras lo estés usando." & FONTTYPE_TALK)
                    Exit Sub

                End If

            End If

            UserList(UserIndex).ComUsu.Objeto = val(Arg1)
            UserList(UserIndex).ComUsu.Cant = val(Arg2)

            If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu <> UserIndex Then
                Call FinComerciarUsu(UserIndex)
                Exit Sub
            Else

                '[CORREGIDO]
                If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = True Then
                    'NO NO NO vos te estas pasando de listo...
                    UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = False
                    Call SendData(SendTarget.ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & _
                                                                                           " ha cambiado su oferta." & FONTTYPE_TALK)

                End If

                '[/CORREGIDO]
                'Es la ofrenda de respuesta :)
                Call EnviarObjetoTransaccion(UserList(UserIndex).ComUsu.DestUsu)

            End If

        End If

        Exit Sub

    End Select

    '[/Alejo]

    Select Case UCase$(Left$(rData, 8))
    
    Case "INIQUEST"
     rData = Right$(rData, Len(rData) - 8)
     
     Call IniciarMisionQuest(UserIndex, rData)
    
    Exit Sub
    
    Case "ENTQUEST"
      rData = Right$(rData, Len(rData) - 8)
      
      Call EntregarMisionQuest(UserIndex)
      
    Exit Sub
    
    Case "ACEPPEAT"    'aceptar paz
        rData = Right$(rData, Len(rData) - 8)
        tInt = modGuilds.r_AceptarPropuestaDePaz(UserIndex, rData, tStr)

        If tInt = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "||Tu clan ha firmado la paz con " & rData & _
                                                                                        FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||Tu clan ha firmado la paz con " & Guilds(UserList(UserIndex).GuildIndex).GuildName & FONTTYPE_GUILD)

            Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "TW" & SONIDOS_GUILD.SND_DECLAREWAR)
            Call SendData(SendTarget.ToGuildMembers, tInt, 0, "TW" & SONIDOS_GUILD.SND_DECLAREWAR)

        End If

        Exit Sub

    Case "RECPALIA"    'rechazar alianza
        rData = Right$(rData, Len(rData) - 8)
        tInt = modGuilds.r_RechazarPropuestaDeAlianza(UserIndex, rData, tStr)

        If tInt = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "||Tu clan rechazado la propuesta de alianza de " & _
                                                                                        rData & FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(UserIndex).Name & _
                                                            " ha rechazado nuestra propuesta de alianza con su clan." & FONTTYPE_GUILD)

        End If

        Exit Sub

    Case "RECPPEAT"    'rechazar propuesta de paz
        rData = Right$(rData, Len(rData) - 8)
        tInt = modGuilds.r_RechazarPropuestaDePaz(UserIndex, rData, tStr)

        If tInt = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "||Tu clan rechazado la propuesta de paz de " & rData & _
                                                                                        FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(UserIndex).Name & _
                                                            " ha rechazado nuestra propuesta de paz con su clan." & FONTTYPE_GUILD)

        End If

        Exit Sub

    Case "ACEPALIA"    'aceptar alianza
        rData = Right$(rData, Len(rData) - 8)
        tInt = modGuilds.r_AceptarPropuestaDeAlianza(UserIndex, rData, tStr)

        If tInt = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "||Tu clan ha firmado la alianza con " & rData & _
                                                                                        FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||Tu clan ha firmado la paz con " & UserList(UserIndex).Name & FONTTYPE_GUILD)

        End If

        Exit Sub

    Case "PEACEOFF"
        'un clan solicita propuesta de paz a otro
        rData = Right$(rData, Len(rData) - 8)
        Arg1 = ReadField(1, rData, Asc(","))
        Arg2 = ReadField(2, rData, Asc(","))

        If modGuilds.r_ClanGeneraPropuesta(UserIndex, Arg1, PAZ, Arg2, Arg3) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Propuesta de paz enviada" & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Arg3 & FONTTYPE_GUILD)

        End If

        Exit Sub

    Case "ALLIEOFF"    'un clan solicita propuesta de alianza a otro
        rData = Right$(rData, Len(rData) - 8)
        Arg1 = ReadField(1, rData, Asc(","))
        Arg2 = ReadField(2, rData, Asc(","))

        If modGuilds.r_ClanGeneraPropuesta(UserIndex, Arg1, ALIADOS, Arg2, Arg3) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Propuesta de alianza enviada" & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Arg3 & FONTTYPE_GUILD)

        End If

        Exit Sub

    Case "ALLIEDET"
        'un clan pide los detalles de una propuesta de ALIANZA
        rData = Right$(rData, Len(rData) - 8)
        tStr = modGuilds.r_VerPropuesta(UserIndex, rData, ALIADOS, Arg1)

        If tStr = vbNullString Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Arg1 & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ALLIEDE" & tStr)

        End If

        Exit Sub

    Case "PEACEDET"    '-"ALLIEDET"
        'un clan pide los detalles de una propuesta de paz
        rData = Right$(rData, Len(rData) - 8)
        tStr = modGuilds.r_VerPropuesta(UserIndex, rData, PAZ, Arg1)

        If tStr = vbNullString Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Arg1 & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "PEACEDE" & tStr)

        End If

        Exit Sub

    Case "ENVCOMEN"
        rData = Trim$(Right$(rData, Len(rData) - 8))

        If rData = vbNullString Then Exit Sub
        tStr = modGuilds.a_DetallesAspirante(UserIndex, rData)

        If tStr = vbNullString Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| El personaje no ha mandado solicitud, o no estás habilitado para verla." & _
                                                            FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "PETICIO" & tStr)

        End If

        Exit Sub

    Case "ENVALPRO"    'enviame la lista de propuestas de alianza
        TIndex = modGuilds.r_CantidadDePropuestas(UserIndex, ALIADOS)
        tStr = "ALLIEPR" & TIndex & ","

        If TIndex > 0 Then
            tStr = tStr & modGuilds.r_ListaDePropuestas(UserIndex, ALIADOS)

        End If

        Call SendData(SendTarget.ToIndex, UserIndex, 0, tStr)
        Exit Sub

    Case "ENVPROPP"    'enviame la lista de propuestas de paz
        TIndex = modGuilds.r_CantidadDePropuestas(UserIndex, PAZ)
        tStr = "PEACEPR" & TIndex & ","

        If TIndex > 0 Then
            tStr = tStr & modGuilds.r_ListaDePropuestas(UserIndex, PAZ)

        End If

        Call SendData(SendTarget.ToIndex, UserIndex, 0, tStr)
        Exit Sub

    Case "DECGUERR"    'declaro la guerra
        rData = Right$(rData, Len(rData) - 8)
        tInt = modGuilds.r_DeclararGuerra(UserIndex, rData, tStr)

        If tInt = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
        Else

            Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "TW" & SONIDOS_GUILD.SND_DECLAREWAR)
            Call SendData(SendTarget.ToGuildMembers, tInt, 0, "TW" & SONIDOS_GUILD.SND_DECLAREWAR)

            'WAR shall be!
            Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "||Tu clan le declaró la guerra a " & rData & "." & _
                                                                                        FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & Guilds(UserList(UserIndex).GuildIndex).GuildName & " le declaró la guerra a tu clan." & _
                                                              FONTTYPE_GUILD)
        End If

        Exit Sub

    Case "NEWWEBSI"
        rData = Right$(rData, Len(rData) - 8)
        Call modGuilds.ActualizarWebSite(UserIndex, rData)
        Exit Sub

    Case "ACEPTARI"
        rData = Right$(rData, Len(rData) - 8)

        If Not modGuilds.a_AceptarAspirante(UserIndex, rData, tStr) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
        Else
            tInt = NameIndex(rData)

            If tInt > 0 Then
                Call modGuilds.m_ConectarMiembroAClan(tInt, UserList(UserIndex).GuildIndex)

            End If

            Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "||" & rData & _
                                                                                      " ha sido aceptado en el clan." & FONTTYPE_GUILD)

            Call modGuilds.NuevoMiembro(rData)

            Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "TW" & SONIDOS_GUILD.SND_ACEPTADOCLAN)

            Call modGuilds.GuildNewMemberItem(tInt)

            Call WarpUserChar(tInt, UserList(tInt).pos.Map, UserList(tInt).pos.X, UserList(tInt).pos.Y)
        End If

        Exit Sub

    Case "RECHAZAR"
        rData = Trim$(Right$(rData, Len(rData) - 8))
        Arg1 = ReadField(1, rData, Asc(","))
        Arg2 = ReadField(2, rData, Asc(","))

        If Not modGuilds.a_RechazarAspirante(UserIndex, Arg1, Arg2, Arg3) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| " & Arg3 & FONTTYPE_GUILD)
        Else
            tInt = NameIndex(Arg1)
            tStr = Arg3 & ": " & Arg2       'el mensaje de rechazo

            If tInt > 0 Then
                Call SendData(SendTarget.ToIndex, tInt, 0, "|| " & tStr & FONTTYPE_GUILD)
            Else
                'hay que grabar en el char su rechazo
                Call modGuilds.a_RechazarAspiranteChar(Arg1, UserList(UserIndex).GuildIndex, Arg2)

            End If

        End If

        Exit Sub

    Case "ECHARCLA"
        'el lider echa de clan a alguien
        rData = Trim$(Right$(rData, Len(rData) - 8))

        If Not Guilds(UserList(UserIndex).GuildIndex).ViewPassword(UserIndex, ReadField(1, rData, 44)) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La Contraseña es incorrecta ó no eres el Lider del clan." & FONTTYPE_FIGHT)
            Exit Sub
        End If

        rData = Right(ReadField(2, rData, 44), Len(ReadField(2, rData, 44)) - 1)

        tInt = modGuilds.m_EcharMiembroDeClan(UserIndex, rData, True)

        If tInt > 0 Then
            Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & rData & " fue expulsado del clan." & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| No puedes expulsar ese personaje del clan." & FONTTYPE_GUILD)

        End If

        Exit Sub

    Case "ACTGNEWS"
        rData = Right$(rData, Len(rData) - 8)
        Call modGuilds.ActualizarNoticias(UserIndex, rData)
        Exit Sub

    Case "1HRINFO<"
        rData = Right$(rData, Len(rData) - 8)

        If Trim$(rData) = vbNullString Then Exit Sub
        tStr = modGuilds.a_DetallesPersonaje(UserIndex, rData, Arg1)

        If tStr = vbNullString Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Arg1 & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "CHRINFO" & tStr)

        End If

        Exit Sub

    Case "ABREELEC"

        If Not modGuilds.v_AbrirElecciones(UserIndex, tStr) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, _
                          "||¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " _
                        & UserList(UserIndex).Name & FONTTYPE_GUILD)

        End If

        Exit Sub

    End Select

    Select Case UCase$(Left$(rData, 9))

    Case "SOLICITUD"
        rData = Right$(rData, Len(rData) - 9)
        Arg1 = ReadField(1, rData, Asc(","))
        Arg2 = ReadField(2, rData, Asc(","))

        If Not modGuilds.a_NuevoAspirante(UserIndex, Arg1, Arg2, tStr) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La solicitud fué recibida por el lider del clan, ahora debes esperar la respuesta." _
                                                          & FONTTYPE_GUILD)

        End If

        Exit Sub

    End Select

    Select Case UCase$(Left$(rData, 11))

    Case "CLANDETAILS"
        rData = Right$(rData, Len(rData) - 11)

        If Trim$(rData) = vbNullString Then Exit Sub
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "CLKNDET" & modGuilds.SendGuildDetails(rData))
        Exit Sub

    End Select

    Call HandleData_4(UserIndex, rData, Procesado)

    Procesado = False

End Sub

Public Sub HandleData_4(ByVal UserIndex As Integer, rData As String, ByRef Procesado As Boolean)

    Dim Rs As Integer

    If UCase$(Left$(rData, 9)) = "VALIDBANK" Then
        rData = UCase$(Right$(rData, Len(rData) - 9))

        If rData = "" Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tienes que poner La respuesta secreta." & FONTTYPE_Motd4)
            Exit Sub
        ElseIf rData = UCase$(UserList(UserIndex).PalabraSecreta) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya puedes Abrir el Banco con normalidad, bienvenido/a " & UserList(UserIndex).Name & "." & FONTTYPE_Motd4)
            UserList(UserIndex).flags.ValidBank = 1
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La Respuesta secreta que nos proporciono, no coincide con la del registro." & FONTTYPE_Motd4)
            Exit Sub
        End If

    End If

    If UCase$(Left$(rData, 7)) = "BANKOBJ" Then
        If UserList(UserIndex).flags.Montado = True Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Bajate de la montura ó Quitatela de tu lado.." & FONTTYPE_TALK)
            Exit Sub
        End If

        Call IniciarDeposito(UserIndex)
    End If

    If UCase$(Left$(rData, 7)) = "BANKDEP" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "BAND" & UserList(UserIndex).Stats.GLD)
    End If

    If UCase$(Left$(rData, 7)) = "BANKRET" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "BANR" & UserList(UserIndex).Stats.Banco)
    End If

    If UCase$(Left$(rData, 7)) = "DEPBANK" Then
        rData = UCase$(Right$(rData, Len(rData) - 7))

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub
        End If

        UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + val(rData)
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rData)

        Call EnviarOro(UserIndex)

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "BANF" & UserList(UserIndex).Stats.Banco & "," & UserList(UserIndex).Stats.GLD)
    End If

    If UCase$(Left$(rData, 7)) = "BANKVOL" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "BANP" & UserList(UserIndex).Stats.Banco & "," & UserList(UserIndex).Stats.GLD & "," & UserList(UserIndex).BancoInvent.NroItems)
    End If

    If UCase$(Left$(rData, 7)) = "RETBANK" Then
        rData = UCase$(Right$(rData, Len(rData) - 7))

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub
        End If

        UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(rData)
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + val(rData)

        Call EnviarOro(UserIndex)

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "BANF" & UserList(UserIndex).Stats.Banco & "," & UserList(UserIndex).Stats.GLD)
    End If

    If UCase$(Left$(rData, 7)) = "ENVHECA" Then
        Call EnviarOlvidoHechizos(UserIndex)
    End If

    If UCase$(Left$(rData, 7)) = "OLVHECA" Then
        rData = UCase$(Right$(rData, Len(rData) - 7))

        Rs = val(ReadField(1, rData, 44))
        rData = ReadField(2, rData, 44)

        Call OlvidaHechizo(UserIndex, Rs, rData)

    End If

    If UCase$(Left$(rData, 7)) = "CHAHEAD" Then
        rData = UCase$(Right$(rData, Len(rData) - 7))

        With UserList(UserIndex)

            If .flags.Muerto = 1 Then Exit Sub

            If OroCirujia > .Stats.GLD Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes suficientes monedas de oro para la cirugía." & FONTTYPE_INFO)
                Exit Sub
            End If

            .char.Head = rData
            .Stats.GLD = .Stats.GLD - OroCirujia
            .OrigChar.Head = rData

            Call ChangeUserChar(SendTarget.ToPCArea, UserIndex, .pos.Map, UserIndex, .char.Body, .char.Head, _
                                .char.heading, .char.WeaponAnim, .char.ShieldAnim, .char.CascoAnim, .char.Alas)

            Call EnviarOro(UserIndex)

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu cirugía se ha realizado correctamente." & FONTTYPE_INFO)

        End With

    End If

    If UCase$(Left$(rData, 7)) = "COMSAST" Then
        rData = UCase$(Right$(rData, Len(rData) - 7))

        If rData > 0 Then
            Call SastreConstruirItem(UserIndex, rData)
        End If

    End If

    If UCase$(Left$(rData, 7)) = "COMHECH" Then
        rData = UCase$(Right$(rData, Len(rData) - 7))

        If rData > 0 Then
            Call HechizeriaConstruirItem(UserIndex, ReadField(1, rData, 44), ReadField(2, rData, 44))
        End If

    End If

    If UCase$(Left$(rData, 7)) = "ACTOBHW" Then
        rData = UCase$(Right$(rData, Len(rData) - 7))

        Call EnviarArmasMagicasConstruibles(UserIndex)

    End If

    If UCase$(Left$(rData, 7)) = "ACTOBHA" Then
        rData = UCase$(Right$(rData, Len(rData) - 7))

        Call EnviarArmadurasMagicasConstruibles(UserIndex)

    End If

    If UCase$(Left$(rData, 7)) = "COMHERM" Then
        rData = UCase$(Right$(rData, Len(rData) - 7))

        If rData > 0 Then
            Call HerreroMagicoConstruirItem(UserIndex, rData)
        End If

    End If

    Procesado = False
End Sub
