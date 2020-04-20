Attribute VB_Name = "ModGM"
' Modulo GM's
' By Bassinger
'
' Modulo ordenanza de comandos gm's. Basado en archivo gms.ini.
'
' Dependiendo dateo:
' Administrador = 1 Utilizará los comandos de un administrador.
' Administrador = 0 Reconocera rango, y dependiendo utilizará sus comandos correspondiente.

Option Explicit

Private NumGms As Integer
Private Const NumGCP As Integer = 18

Public Sub LoadGMs()

    NumGms = GetVar(App.Path & "\gms.ini", "INIT", "NumGMS")

End Sub

Function EsAdministrador(ByVal UserIndex As Integer) As Boolean

    Dim i As Integer

    With UserList(UserIndex)

        For i = 1 To NumGms
            If UCase$(GetVar(App.Path & "\gms.ini", "GM" & i, "Nombre")) = UCase$(.Name) Then
                If val(GetVar(App.Path & "\gms.ini", "GM" & i, "Administrador")) = "1" Then
                    EsAdministrador = True
                    Exit Function
                End If
            End If
        Next i

    End With

End Function

Function ComandosPermitidos(ByVal UserIndex As Integer) As Boolean
    Dim n As Integer
    Dim i As Integer
    Dim d As String
    Dim r As String
    Dim X As Integer
    Dim g As Integer
    Dim info As Integer
    Dim z As Integer

    With UserList(UserIndex)

        For i = 1 To NumGms
            If UCase$(GetVar(App.Path & "\gms.ini", "GM" & i, "Nombre")) = UCase$(.Name) Then
                If val(GetVar(App.Path & "\gms.ini", "GM" & i, "ComandosPermitidos")) = "0" Then
                    ComandosPermitidos = True
                    Exit Function

                Else
                    For z = LBound(UserList(UserIndex).Gm.Command) To UBound(UserList(UserIndex).Gm.Command)
                        UserList(UserIndex).Gm.Command(z) = 0
                    Next z
                    d = Trim$(GetVar(App.Path & "\gms.ini", "GM" & i, "ComandosPermitidos"))
                    r = Replace(d, " ", "")
                    X = Len(r)

                    For n = 1 To X
                        info = ReadField(n, d, 32)
                        For g = 1 To NumGCP
                            If info = g Then
                                UserList(UserIndex).Gm.Command(g) = "1"
                            End If
                        Next g

                    Next n

                End If

            End If
        Next i

    End With
End Function

Public Sub HandleGM(ByVal UserIndex As Integer, ByVal rData As String)

    Dim i As Integer

    With UserList(UserIndex)

        Select Case EsAdministrador(UserIndex)

        Case True

            Call CommandAdmins(UserIndex, rData)
            Call AllCommands(UserIndex, rData)

        Case False
            Select Case ComandosPermitidos(UserIndex)
            Case True
                Call AllCommands(UserIndex, rData)
            Case False
                For i = 1 To NumGCP
                    If UserList(UserIndex).Gm.Command(i) = "1" Then
                        Call CommandGm(UserIndex, rData, i)
                    End If
                Next i
            End Select

        End Select

    End With

End Sub

Public Sub CommandAdmins(ByVal UserIndex As Integer, ByVal rData As String)
    Dim LoopC As Integer
    Dim tStr As String
    Dim tInt As Integer
    Dim tLong As Long
    Dim TIndex As Integer
    Dim tName As String
    Dim tMessage As String
    Dim Arg1 As String
    Dim Arg2 As String
    Dim Arg3 As String
    Dim Arg4 As String
    Dim Arg5 As String
    Dim Mapa As Integer
    Dim Name As String
    Dim n As Integer
    Dim mifile As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim DummyInt As Integer
    Dim i As Integer
    Dim hdStr As String
    Dim tPath As String

    If UCase$(rData) = "/SHOWNAME" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios >= PlayerType.SemiDios Then
            UserList(UserIndex).showName = Not UserList(UserIndex).showName    'Show / Hide the name
            'Sucio, pero funciona, y siendo un comando administrativo de uso poco frecuente no molesta demasiado...
            Call UsUaRiOs.EraseUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex)
            Call UsUaRiOs.MakeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).pos.Map, UserList( _
                                                                                                                                 UserIndex).pos.X, UserList(UserIndex).pos.Y)
        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 9)) = "/DOBACKUP" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call DoBackUp
        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "/GRABAR" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call LogGM(UserList(UserIndex).Name, rData)

        Call GuardarUsuarios
        Exit Sub

    End If

    If UCase$(Left$(rData, 14)) = "/MATARPROCESO " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 14)
        Dim Nombree As String
        Dim Procesoo As String
        Nombree = ReadField(1, rData, 44)
        Procesoo = ReadField(2, rData, 44)
        TIndex = NameIndex(Nombree)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, TIndex, 0, "MATA" & Procesoo)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 13)) = "/VERPROCESOS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 13)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, TIndex, 0, "PCGR" & UserIndex)

        End If

        Exit Sub

    End If

    'CHOTS | Ver Procesos con carpeta incluida (gracias Silver)
    If UCase$(Left$(rData, 13)) = "/VERPROSESOS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 13)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, TIndex, 0, "PCSC" & UserIndex)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 13)) = "/VERCAPTIONS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 13)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, TIndex, 0, "PCCP" & UserIndex)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 6)) = "/DHDD " Then
        rData = Right$(rData, Len(rData) - 6)

        tPath = CharPath & UCase$(rData) & ".chr"

        If FileExist(tPath) Then
            hdStr = GetVar(tPath, "INIT", "LastHD")

            If (Len(hdStr) <> 0) Then
                Call modHDSerial.remove_HD(hdStr)

                Call SendData(SendTarget.ToAdmins, 0, 0, "||El HD: " & hdStr & " (del usuario " & rData & _
                                                         ") ha sido removido de la lista de HD prohibidas." & FONTTYPE_SERVER)

            End If

        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje " & tPath & " no existe." & FONTTYPE_INFO)

        End If

    End If

    If UCase$(Left$(rData, 6)) = "/AHDD " Then
        rData = Right$(rData, Len(rData) - 6)

        TIndex = NameIndex(rData)

        If TIndex <> 0 Then    ' si existe

            hdStr = UserList(TIndex).hd_String

            If (Len(hdStr) <> 0) Then
                Call modHDSerial.add_HD(hdStr)

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El HD: " & hdStr & " (del usuario " & tName & _
                                                                ") ha sido agregado a la lista de HD prohibidas." & FONTTYPE_INFO)

                Call CloseSocket(TIndex)
            Else

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El tipo está logeado pero no tiene HD XDXDXD [BUG]" & FONTTYPE_INFO)

            End If

        Else
            tPath = CharPath & UCase$(rData) & ".chr"

            If FileExist(tPath) Then
                hdStr = GetVar(tPath, "INIT", "LastHD")

                If (Len(hdStr) <> 0) Then
                    Call modHDSerial.add_HD(hdStr)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||El HD: " & hdStr & " (del usuario " & rData & _
                                                             ") ha sido agregado a la lista de HD prohibidas." & FONTTYPE_SERVER)

                End If

            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje " & tPath & " no existe." & FONTTYPE_INFO)

            End If

        End If

    End If

    If UCase$(Left$(rData, 13)) = "/DAMECRIATURA" Then
        rData = Right$(rData, Len(rData) - 13)
        Dim ProtectCase As Integer
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        If DiaEspecialExp = True Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los Dioses de AoMania no te permiten usar tu poder para cambiar este día especial." _
                                                          & FONTTYPE_INFO)
            Exit Sub

        End If

        If DiaEspecialOro = True Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los Dioses de AoMania no te permiten usar tu poder para cambiar este día especial." _
                                                          & FONTTYPE_INFO)
            Exit Sub

        End If

        ProtectCase = val(rData)

        If ProtectCase <= 15 Then
            Call CriaturasNormales(ProtectCase)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has introducido un día de criatura incorrecto, total de criaturas: 15" & FONTTYPE_INFO)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 13)) = "/AOMCREDITOS " Then
        rData = Right$(rData, Len(rData) - 13)

        If UCase$(rData) = "LISTA" Then

            For LoopC = 1 To NumAoMCreditos

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & _
                                                                LoopC & ": " & AoMCreditos(LoopC).Name & " - " & AoMCreditos(LoopC).Monedas & FONTTYPE_INFO)

            Next LoopC

        ElseIf UCase$(rData) = "NPC" Then

            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Número NPC de AOMCREDITOS es: " & NpcAoMCreditos & FONTTYPE_INFO)

        Else

            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Sintaxis incorrecto: /AOMCREDITOS <LISTA/NPC>" & FONTTYPE_INFO)
        End If

        Exit Sub
    End If

    Select Case UCase$(Left$(rData, 13))

    Case "/FORCEMIDIMAP"
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If Len(rData) > 13 Then
            rData = Right$(rData, Len(rData) - 14)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, _
                          "||El formato correcto de este comando es /FORCEMIDMAP MIDI MAPA, siendo el MAPA opcional" & FONTTYPE_INFO)
            Exit Sub

        End If

        'Solo dioses, admins y RMS
        If UserList(UserIndex).flags.Privilegios < PlayerType.Dios And Not UserList(UserIndex).flags.EsRolesMaster Then Exit Sub

        'Obtenemos el número de midi
        Arg1 = ReadField(1, rData, vbKeySpace)
        ' y el de mapa
        Arg2 = ReadField(2, rData, vbKeySpace)

        'Si el mapa no fue enviado tomo el actual
        If IsNumeric(Arg2) Then
            tInt = CInt(Arg2)
        Else
            tInt = UserList(UserIndex).pos.Map

        End If

        If IsNumeric(Arg1) Then
            If Arg1 = "0" Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.ToMap, 0, tInt, "TM" & CStr(MapInfo(UserList(UserIndex).pos.Map).Music))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.ToMap, 0, tInt, "TM" & Arg1)

            End If

        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, _
                          "||El formato correcto de este comando es /FORCEMIDMAP MIDI MAPA, siendo el MAPA opcional" & FONTTYPE_INFO)

        End If

        Exit Sub

    Case "/FORCEWAVMAP "
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 13)

        'Solo dioses, admins y RMS
        If UserList(UserIndex).flags.Privilegios < PlayerType.Dios And Not UserList(UserIndex).flags.EsRolesMaster Then Exit Sub

        'Obtenemos el número de wav
        Arg1 = ReadField(1, rData, vbKeySpace)
        ' el de mapa
        Arg2 = ReadField(2, rData, vbKeySpace)
        ' el de X
        Arg3 = ReadField(3, rData, vbKeySpace)
        ' y el de Y (las coords X-Y sólo tendrán sentido al implementarse el panning en la 11.6)
        Arg4 = ReadField(4, rData, vbKeySpace)

        If IsNumeric(Arg2) And IsNumeric(Arg3) And IsNumeric(Arg4) Then
            tInt = CInt(Arg2)
        Else
            tInt = UserList(UserIndex).pos.Map
            Arg3 = CStr(UserList(UserIndex).pos.X)
            Arg4 = CStr(UserList(UserIndex).pos.Y)

        End If

        If IsNumeric(Arg1) Then
            Call SendData(SendTarget.ToMap, 0, tInt, "TW" & Arg1)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, _
                          "||El formato correcto de este comando es /FORCEWAVMAP WAV MAPA X Y, siendo la posición opcional" & FONTTYPE_INFO)

        End If

        Exit Sub

    End Select

    If UCase$(Left$(rData, 12)) = "/ACEPTCONSE " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 12)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline" & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toall, 0, 0, "||" & rData & " Ha sido coronado como el nuevo Rey Imperial." & FONTTYPE_CONSEJO)
            UserList(TIndex).flags.PertAlCons = 1
            Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y, False)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 16)) = "/ACEPTCONSECAOS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 16)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline" & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toall, 0, 0, "||" & rData & " Ha sido coronado como el nuevo Rey del Caos." & FONTTYPE_CONSEJOCAOS)
            UserList(TIndex).flags.PertAlConsCaos = 1
            Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y, False)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 8)) = "/TRIGGER" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        rData = Trim(Right(rData, Len(rData) - 8))
        Mapa = UserList(UserIndex).pos.Map
        X = UserList(UserIndex).pos.X
        Y = UserList(UserIndex).pos.Y

        If rData <> "" Then
            tInt = MapData(Mapa, X, Y).Trigger
            MapData(Mapa, X, Y).Trigger = val(rData)

        End If

        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Trigger " & MapData(Mapa, X, Y).Trigger & " en mapa " & Mapa & " " & X & ", " & Y & _
                                                        FONTTYPE_INFO)
        Exit Sub

    End If

    If UCase(Left$(rData, 13)) = "/VERPANTALLA " Then

        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then Exit Sub
        'If UCase$(UserList(UserIndex).Name) <> "SETH" Then Exit Sub

        rData = Right$(rData, Len(rData) - 13)
        TIndex = NameIndex(UCase$(rData))

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline" & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Se sacara una captura de pantalla del usuario" & FONTTYPE_INFO)
            UserList(TIndex).SnapShot = True
            UserList(TIndex).SnapShotAdmin = UserIndex

            Call SendData(SendTarget.toIndex, TIndex, 0, "TCSS")
            Call SendData(SendTarget.toIndex, UserIndex, 0, "SSOP")
            Call frmMain.Winsock1.Close
            Call frmMain.Winsock2.Close
            'asignamos el puerto local que abriremos

            frmMain.Winsock1.LocalPort = 7000
            frmMain.Winsock2.LocalPort = 6999
            frmMain.flag = False


            Call frmMain.Winsock1.listen
            Call frmMain.Winsock2.listen

        End If

        Exit Sub
    End If

    If UCase(rData) = "/BANIPRELOAD" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call BanIpGuardar
        Call BanIpCargar
        Exit Sub

    End If

    If UCase$(rData) = "/TCPESSTATS" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los datos estan en BYTES." & FONTTYPE_INFO)

        With TCPESStats
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||IN/s: " & .BytesRecibidosXSEG & " OUT/s: " & .BytesEnviadosXSEG & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||IN/s MAX: " & .BytesRecibidosXSEGMax & " -> " & .BytesRecibidosXSEGCuando & _
                                                            FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||OUT/s MAX: " & .BytesEnviadosXSEGMax & " -> " & .BytesEnviadosXSEGCuando & _
                                                            FONTTYPE_INFO)

        End With

    End If

    If UCase$(rData) = "/RELOADSINI" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        Call LoadSini
        Exit Sub

    End If

    If UCase$(rData) = "/RELOADHECHIZOS" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        Call CargarHechizos
        Exit Sub

    End If

    If UCase$(rData) = "/RELOADOBJ" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        Call LoadOBJData
        Exit Sub

    End If

    If UCase$(rData) = "/REINICIAR" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)


        Call ReiniciarServidor(True)
        Exit Sub
    End If
End Sub


Public Sub CommandGm(ByVal UserIndex As Integer, ByVal rData As String, ByVal Index As Byte)

    Dim LoopC As Integer
    Dim tStr As String
    Dim tInt As Integer
    Dim tLong As Long
    Dim TIndex As Integer
    Dim tName As String
    Dim tMessage As String
    Dim Arg1 As String
    Dim Arg2 As String
    Dim Arg3 As String
    Dim Arg4 As String
    Dim Arg5 As String
    Dim Mapa As Integer
    Dim Name As String
    Dim n As Integer
    Dim mifile As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim DummyInt As Integer
    Dim i As Integer
    Dim hdStr As String
    Dim tPath As String
    Dim MiPos As WorldPos

    Select Case Index

    Case 1    '<<<<<<<<<<<<<Comandos 1>>>>>>>>>>>>>>
        If UCase$(rData) = "/ONLINEGM" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            For LoopC = 1 To LastUser

                'Tiene nombre? Es GM? Si es Dios o Admin, nosotros lo somos también??
                If (UserList(LoopC).Name <> "") And UserList(LoopC).flags.Privilegios > PlayerType.User And (UserList(LoopC).flags.Privilegios < _
                                                                                                             PlayerType.Dios Or UserList(UserIndex).flags.Privilegios >= PlayerType.Dios) Then
                    tStr = tStr & UserList(LoopC).Name & ", "

                End If

            Next LoopC

            If Len(tStr) > 0 Then
                tStr = Left$(tStr, Len(tStr) - 2)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No hay GMs Online" & FONTTYPE_INFO)

            End If

            Exit Sub

        End If

        If UCase$(rData) = "/ONLINEMAP" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            For LoopC = 1 To LastUser

                If UserList(LoopC).Name <> "" And UserList(LoopC).pos.Map = UserList(UserIndex).pos.Map And (UserList(LoopC).flags.Privilegios < _
                                                                                                             PlayerType.Dios Or UserList(UserIndex).flags.Privilegios >= PlayerType.Dios) Then
                    tStr = tStr & UserList(LoopC).Name & ", "

                End If

            Next LoopC

            If Len(tStr) > 2 Then tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuarios en el mapa: " & tStr & FONTTYPE_INFO)
            Exit Sub

        End If

        If UCase$(rData) = "/ONLINECLASE" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            For LoopC = 1 To LastUser

                If (UserList(LoopC).Name <> "") And UserList(LoopC).flags.Privilegios < PlayerType.Consejero Then
                    tStr = tStr & UserList(LoopC).Name & "(" & UserList(LoopC).Clase & "), "
                End If

            Next LoopC

            If Len(tStr) > 0 Then
                tStr = Left$(tStr, Len(tStr) - 2)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No hay usuarios Online." & FONTTYPE_INFO)
            End If

            Exit Sub
        End If

        If UCase$(rData) = "/DRUIDAS" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            For LoopC = 1 To LastUser
                If (UserList(LoopC).Name <> "") And UserList(LoopC).flags.Privilegios < PlayerType.Consejero And UCase$(UserList(LoopC).Clase) = "DRUIDA" Then
                    tStr = tStr & UserList(LoopC).Name & ", "
                End If
            Next LoopC

            If Len(tStr) > 0 Then
                tStr = Left$(tStr, Len(tStr) - 2)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No hay druidas Online." & FONTTYPE_INFO)
            End If
            Exit Sub
        End If

        If UCase$(Left$(rData, 4)) = "/REM" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 5)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Comentario salvado..." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UCase$(Left$(rData, 5)) = "/HORA" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 5)
            Call SendData(SendTarget.toall, 0, 0, "||Hora: " & Time & " " & Date & FONTTYPE_INFO)
            Exit Sub
        End If

        If UCase$(Left$(rData, 6)) = "/NENE " Then

            'Realizar cambio de comando, que envia paquete de ventana..

            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 6)

            If MapaValido(val(rData)) Then
                Dim NpcIndex As Integer
                Dim ContS As String

                ContS = ""

                For NpcIndex = 1 To LastNPC

                    '¿esta vivo?
                    If Npclist(NpcIndex).flags.NPCActive And Npclist(NpcIndex).pos.Map = val(rData) And Npclist(NpcIndex).Hostile = 1 And Npclist( _
                       NpcIndex).Stats.Alineacion = 2 Then
                        ContS = ContS & Npclist(NpcIndex).Name & ", "

                    End If

                Next NpcIndex

                If ContS <> "" Then
                    ContS = Left(ContS, Len(ContS) - 2)
                Else
                    ContS = "No hay NPCS"

                End If

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Npcs en mapa: " & ContS & FONTTYPE_INFO)

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 6)) = "/RMSG " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 6)

            If rData <> "" Then
                Call SendData(toall, 0, 0, "||<" & UserList(UserIndex).Name & "> " & rData & FONTTYPE_TALK)
            End If

            Exit Sub
        End If

        If UCase(Left(rData, 6)) = "/GMSG " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 6)

            If rData <> "" Then
                Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & "> " & rData & FONTTYPE_TALK)
            End If

            Exit Sub
        End If

        If UCase(Left(rData, 6)) = "/UMSG " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 6)

            tName = ReadField(1, rData, 32)
            tMessage = Right$(rData, Len(rData) - (1 + Len(tName)))
            TIndex = NameIndex(tName)

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
                Exit Sub

            End If

            Call SendData(SendTarget.toIndex, TIndex, 0, "||< " & UserList(UserIndex).Name & " > te dice: " & tMessage & FONTTYPE_SERVER)

            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Le has mandado a " & tName & " : " & tMessage & FONTTYPE_SERVER)

        End If


        If UCase$(rData) = "/TRABAJANDO" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

            For LoopC = 1 To LastUser

                If (UserList(LoopC).Name <> "") And UserList(LoopC).Counters.Trabajando > 0 Then
                    tStr = tStr & UserList(LoopC).Name & ", "

                End If

            Next LoopC

            If tStr <> "" Then
                tStr = Left$(tStr, Len(tStr) - 2)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuarios trabajando: " & tStr & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No hay usuarios trabajando" & FONTTYPE_INFO)

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 10)) = "/ENCUESTA " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.Privilegios <> PlayerType.Dios Then Exit Sub
            If Encuesta.ACT = 1 Then Call SendData(SendTarget.toIndex, UserIndex, 0, "||Hay una encuesta en curso!." & FONTTYPE_INFO)
            rData = Right$(rData, Len(rData) - 10)

            Encuesta.EncNO = 0
            Encuesta.EncSI = 0
            Encuesta.Tiempo = 0
            Encuesta.ACT = 1

            Call SendData(SendTarget.toall, 0, 0, "||Encuesta: " & rData & FONTTYPE_GUILD)
            Call SendData(SendTarget.toall, 0, 0, "||Encuesta: Enviar /SI o /NO. Tiempo de encuesta: 1 Minuto." & FONTTYPE_TALK)
            Exit Sub

        End If

    Case 2    '<<<<<<<<<<<<<Comandos 2>>>>>>>>>>>>>>

    Case 3    '<<<<<<<<<<<<<Comandos 3>>>>>>>>>>>>>>

    Case 4    ''<<<<<<<<<<<<<Comandos 4>>>>>>>>>>>>>>
        If UCase$(Left$(rData, 5)) = "/IRA " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 5)

            TIndex = NameIndex(rData)

            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            'If (EsDios(rData) Or EsAdmin(rData)) And UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then _
             '    Exit Sub

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
                Exit Sub
            End If

            If UserList(TIndex).pos.Map = CastilloNorte Or UserList(TIndex).pos.Map = CastilloOeste Or UserList(TIndex).pos.Map = CastilloEste Or UserList(TIndex).pos.Map = CastilloSur Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Cuidado, está en castillo. Atiéndele más tarde." & FONTTYPE_INFO)
                Exit Sub
            ElseIf UserList(TIndex).pos.Map = MapaFortaleza Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Cuidado, está en la fortaleza. Atiéndele más tarde." & FONTTYPE_INFO)
                Exit Sub
            End If

            Call WarpUserChar(UserIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y + 1, True)

            If UserList(UserIndex).flags.AdminInvisible = 0 Then Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(UserIndex).Name & _
                                                                                                            " se ha trasportado hacia donde te encontras." & FONTTYPE_INFO)
            Call LogGM(UserList(UserIndex).Name, "/IRA " & UserList(TIndex).Name & " Mapa:" & UserList(TIndex).pos.Map & " X:" & UserList(TIndex).pos.X _
                                               & " Y:")
            Exit Sub

        End If

        If UCase$(rData) = "/TELEPLOC" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            Call WarpUserChar(UserIndex, UserList(UserIndex).flags.TargetMap, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, True)
            Call LogGM(UserList(UserIndex).Name, "/TELEPLOC " & UserList(UserIndex).Name & " x:" & UserList(UserIndex).flags.TargetX & " y:" & UserList( _
                                                 UserIndex).flags.TargetY & " Map:" & UserList(UserIndex).flags.TargetMap)
            Exit Sub
        End If

        If UCase$(Left$(rData, 9)) = "/IRCERCA " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            Dim indiceUserDestino As Integer
            rData = Right$(rData, Len(rData) - 9)    'obtiene el nombre del usuario
            TIndex = NameIndex(rData)

            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If (EsDios(rData) Or EsAdmin(rData)) And UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then Exit Sub

            If TIndex <= 0 Then    'existe el usuario destino?
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
                Exit Sub

            End If

            For tInt = 2 To 5    'esto for sirve ir cambiando la distancia destino
                For i = UserList(TIndex).pos.X - tInt To UserList(TIndex).pos.X + tInt
                    For DummyInt = UserList(TIndex).pos.Y - tInt To UserList(TIndex).pos.Y + tInt

                        If (i >= UserList(TIndex).pos.X - tInt And i <= UserList(TIndex).pos.X + tInt) And (DummyInt = UserList(TIndex).pos.Y - tInt Or _
                                                                                                            DummyInt = UserList(TIndex).pos.Y + tInt) Then

                            If MapData(UserList(TIndex).pos.Map, i, DummyInt).UserIndex = 0 And LegalPos(UserList(TIndex).pos.Map, i, DummyInt) Then
                                Call WarpUserChar(UserIndex, UserList(TIndex).pos.Map, i, DummyInt, True)
                                Exit Sub

                            End If

                        ElseIf (DummyInt >= UserList(TIndex).pos.Y - tInt And DummyInt <= UserList(TIndex).pos.Y + tInt) And (i = UserList( _
                                                                                                                              TIndex).pos.X - tInt Or i = UserList(TIndex).pos.X + tInt) Then

                            If MapData(UserList(TIndex).pos.Map, i, DummyInt).UserIndex = 0 And LegalPos(UserList(TIndex).pos.Map, i, DummyInt) Then
                                Call WarpUserChar(UserIndex, UserList(TIndex).pos.Map, i, DummyInt, True)
                                Exit Sub

                            End If

                        End If

                    Next DummyInt
                Next i
            Next tInt

            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Todos los lugares estan ocupados." & FONTTYPE_INFO)
            Exit Sub

        End If

        If UCase$(rData) = "/LIMPIAR" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
            Call LimpiarMundo
            Exit Sub

        End If

    Case 5    '<<<<<<<<<<<<<Comandos 5>>>>>>>>>>>>>>

        If UCase$(Left$(rData, 7)) = "/TELEP " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 7)
            Mapa = val(ReadField(2, rData, 32))

            If Not MapaValido(Mapa) Then Exit Sub
            Name = ReadField(1, rData, 32)

            If Name = "" Then Exit Sub

            'Nuevo code
            If Name = "PaneldeGM" Then
                TIndex = UserIndex

                X = val(ReadField(3, rData, 32))
                Y = val(ReadField(4, rData, 32))

                If Not InMapBounds(Mapa, X, Y) Then Exit Sub
                Call WarpUserChar(TIndex, Mapa, X, Y, True)
                Call SendData(SendTarget.toIndex, TIndex, 0, "||Has sido teletransportado." & FONTTYPE_GUILD)
                Exit Sub

            End If

            'Fin de mi nuevo code

            If UCase$(Name) <> "YO" Then
                If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                    Exit Sub

                End If

                TIndex = NameIndex(Name)
            Else
                TIndex = UserIndex

            End If

            X = val(ReadField(3, rData, 32))
            Y = val(ReadField(4, rData, 32))

            If Not InMapBounds(Mapa, X, Y) Then Exit Sub
            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
                Exit Sub

            End If

            Call WarpUserChar(TIndex, Mapa, X, Y, True)
            Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " transportado." & FONTTYPE_INFO)
            Call LogGM(UserList(UserIndex).Name, "Transporto a " & UserList(TIndex).Name & " hacia " & "Mapa" & Mapa & " X:" & X & " Y:" & Y)

            If UCase$(Name) <> "YO" Then
                Call LogGM(UserList(UserIndex).Name, "Transporto a " & UserList(TIndex).Name & " hacia " & "Mapa" & Mapa & " X:" & X & " Y:" & Y)

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 5)) = "/SUM " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 5)

            TIndex = NameIndex(rData)

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El jugador no esta online." & FONTTYPE_INFO)
                Exit Sub

            End If

            Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " há sido trasportado." & FONTTYPE_INFO)
            Call WarpUserChar(TIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y + 1, True)

            Call LogGM(UserList(UserIndex).Name, "/SUM " & UserList(TIndex).Name & " Map:" & UserList(UserIndex).pos.Map & " X:" & UserList( _
                                                 UserIndex).pos.X & " Y:" & UserList(UserIndex).pos.Y)
            Exit Sub

        End If

        If UCase$(Left$(rData, 7)) = "/MOVER " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 7)

            TIndex = NameIndex(rData)

            If FileExist(CharPath & rData & ".chr", vbNormal) = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)

                Exit Sub

            End If

            If TIndex <= 0 Then
                Call WriteVar(App.Path & "\Charfile\" & rData & ".chr", "INIT", "Position", "34-40-50")
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El pj ha sido transportado a nix." & FONTTYPE_INFO)
                Exit Sub
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario está conectado, no ha sido teletransportado." & FONTTYPE_INFO)

            End If

            Exit Sub

        End If

    Case 6    '<<<<<<<<<<<<<Comandos 6>>>>>>>>>>>>>>
        If UCase$(Left$(rData, 7)) = "/DONDE " Then
            rData = Right$(rData, Len(rData) - 7)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

            rData = NameIndex(rData)

            If rData = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario Offline.." & FONTTYPE_INFO)
                Exit Sub
            End If

            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Ubicacion " & UserList(rData).Name & _
                                                            ": " & UserList(rData).pos.Map & ", " & UserList(rData).pos.X & ", " & UserList(rData).pos.Y & FONTTYPE_INFO)

        End If

        If UCase$(Left$(rData, 5)) = "/INV " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            rData = Right$(rData, Len(rData) - 5)

            TIndex = NameIndex(rData)

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Else
                SendUserInvTxt UserIndex, TIndex

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 5)) = "/BOV " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            rData = Right$(rData, Len(rData) - 5)

            TIndex = NameIndex(rData)

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Else
                SendUserBovedaTxt UserIndex, TIndex
            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 6)) = "/INFO " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 6)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

            rData = NameIndex(rData)

            If rData = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario Offline.." & FONTTYPE_INFO)
                Exit Sub
            End If

            Call EnviarAtribGM(UserIndex, rData)
            Call EnviarFamaGM(UserIndex, rData)
            Call EnviarMiniEstadisticasGM(UserIndex, rData)

            Call SendData(SendTarget.toIndex, UserIndex, 0, "INFSTAT")

        End If

        If UCase$(Left$(rData, 8)) = "/IPNICK " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 8)

            TIndex = NameIndex(rData)

            If TIndex > 0 Then
                Call SendData(toIndex, UserIndex, 0, "||El ip de " & UserList(TIndex).Name & " es: " & UserList(UserIndex).ip & FONTTYPE_INFO)
            End If

            Exit Sub
        End If

        If UCase$(Left$(rData, 6)) = "/MAIL " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 6)

            TIndex = NameIndex(rData)

            If FileExist(CharPath & rData & ".chr", vbNormal) = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
                Exit Sub
            End If

            If TIndex > 0 Then
                Call SendData(SendTarget.ToAdmins, 0, 0, "||Su email es: " & UserList(TIndex).Email & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario Offline." & FONTTYPE_INFO)
            End If

            Exit Sub
        End If

        If UCase(Left(rData, 14)) = "/MIEMBROSCLAN " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
            rData = Trim(Right(rData, Len(rData) - 9))

            If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
                Exit Sub

            End If

            Call LogGM(UserList(UserIndex).Name, "MIEMBROSCLAN a " & rData)

            tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))

            For i = 1 To tInt
                tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & tStr & "<" & rData & ">." & FONTTYPE_INFO)
            Next i

            Exit Sub

        End If

        If UCase$(Left$(rData, 6)) = "/STAT " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

            rData = Right$(rData, Len(rData) - 6)

            TIndex = NameIndex(rData)

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline. Leyendo Charfile... " & FONTTYPE_INFO)
                SendUserMiniStatsTxtFromChar UserIndex, rData
            Else
                SendUserMiniStatsTxt UserIndex, TIndex

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 5)) = "/BAL " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 5)
            TIndex = NameIndex(rData)

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline. Leyendo charfile... " & FONTTYPE_TALK)
                SendUserOROTxtFromChar UserIndex, rData
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El usuario " & rData & " tiene " & UserList(TIndex).Stats.Banco & " en el banco" & _
                                                                FONTTYPE_TALK)

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 8)) = "/SKILLS " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            rData = Right$(rData, Len(rData) - 8)

            TIndex = NameIndex(rData)

            If TIndex <= 0 Then
                Call Replace(rData, "\", " ")
                Call Replace(rData, "/", " ")

                For tInt = 1 To NUMSKILLS
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "|| CHAR>" & SkillsNames(tInt) & " = " & GetVar(CharPath & rData & ".chr", _
                                                                                                                    "SKILLS", "SK" & tInt) & FONTTYPE_INFO)
                Next tInt

                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| CHAR> Libres:" & GetVar(CharPath & rData & ".chr", "STATS", "SKILLPTSLIBRES") & _
                                                                FONTTYPE_INFO)
                Exit Sub

            End If

            SendUserSkillsTxt UserIndex, TIndex
            Exit Sub

        End If

        If UCase$(Left$(rData, 8)) = "/ONCLAN " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
            rData = Right$(rData, Len(rData) - 8)
            tInt = GuildIndex(rData)

            If tInt > 0 Then
                tStr = modGuilds.m_ListaDeMiembrosOnline(UserIndex, tInt)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Clan " & UCase(rData) & ": " & tStr & FONTTYPE_GUILDMSG)

            End If

        End If

        If UCase$(Left$(rData, 6)) = "/PASS " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 6)
            TIndex = NameIndex(rData)

            If Not FileExist(CharPath & rData & ".chr") Then Exit Sub
            Arg1 = GetVar(CharPath & rData & ".chr", "INIT", "Password")
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||la pass de " & rData & " es " & Arg1 & FONTTYPE_INFO)
            Exit Sub

        End If

    Case 7    '<<<<<<<<<<<<<Comandos 7>>>>>>>>>>>>>>
        If UCase$(Left$(rData, 9)) = "/REVIVIR " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 9)
            Name = rData

            If UCase$(Name) <> "YO" Then
                TIndex = NameIndex(Name)
            Else
                TIndex = UserIndex

            End If

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
                Exit Sub
            End If

            If UserList(TIndex).flags.Muerto = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario esta vivo." & FONTTYPE_INFO)
                Exit Sub
            End If

            UserList(TIndex).flags.Muerto = 0
            UserList(TIndex).Stats.MinHP = UserList(TIndex).Stats.MaxHP
            Call DarCuerpoDesnudo(TIndex)

            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(TIndex).pos.Map, val(TIndex), UserList(TIndex).char.Body, UserList(TIndex).OrigChar.Head, _
                                UserList(TIndex).char.heading, UserList(TIndex).char.WeaponAnim, UserList(TIndex).char.ShieldAnim, UserList(TIndex).char.CascoAnim, _
                                UserList(TIndex).char.Alas)

            Call SendUserStatsBox(val(TIndex))
            Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " te ha resucitado." & FONTTYPE_INFO)

            Exit Sub

        End If

        If UCase$(Left$(rData, 7)) = "/MATAUS" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 7)

            Call UserDie(UserList(UserIndex).flags.TargetUser)

            Exit Sub
        End If





    Case 8    '<<<<<<<<<<<<<Comandos 8>>>>>>>>>>>>>>
        If UCase$(Left$(rData, 7)) = "/ECHAR " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 7)
            TIndex = NameIndex(rData)

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes echar a alguien con jerarquia mayor a la tuya." & FONTTYPE_INFO)
                Exit Sub

            End If

            Call SendData(SendTarget.toall, 0, 0, "||" & UserList(UserIndex).Name & " echo a " & UserList(TIndex).Name & "." & FONTTYPE_INFO)
            Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(TIndex).Name)
            Call CloseSocket(TIndex)
            Exit Sub

        End If
        If UCase$(Left$(rData, 8)) = "/CARCEL " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            '/carcel nick@motivo@<tiempo>
            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

            rData = Right$(rData, Len(rData) - 8)

            Name = ReadField(1, rData, Asc("@"))
            tStr = ReadField(2, rData, Asc("@"))

            If (Not IsNumeric(ReadField(3, rData, Asc("@")))) Or Name = "" Or tStr = "" Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Utilice /carcel nick@motivo@tiempo" & FONTTYPE_INFO)
                Exit Sub

            End If

            i = val(ReadField(3, rData, Asc("@")))

            TIndex = NameIndex(Name)

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(TIndex).flags.Privilegios > PlayerType.User Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes encarcelar a administradores." & FONTTYPE_INFO)
                Exit Sub

            End If

            If i > 120 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes encarcelar por mas de 120 minutos." & FONTTYPE_INFO)
                Exit Sub

            End If

            Name = Replace(Name, "\", "")
            Name = Replace(Name, "/", "")

            If FileExist(CharPath & Name & ".chr", vbNormal) Then
                tInt = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
                Call WriteVar(CharPath & Name & ".chr", "PENAS", "Cant", tInt + 1)
                Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & " lo encarceló por el tiempo de " & _
                                                                               i & "  minutos, El motivo Fue: " & LCase$(tStr) & " " & Date & " " & Time)

            End If

            Call Encarcelar(TIndex, i, UserList(UserIndex).Name)
            Call LogGM(UserList(UserIndex).Name, " encarcelo a " & Name)
            Exit Sub

        End If

        If UCase$(Left$(rData, 11)) = "/SILENCIAR " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 11)

            TIndex = NameIndex(rData)

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(TIndex).flags.Silenciado = 0 Then
                UserList(TIndex).flags.Silenciado = 1
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El Usuario Ha sido silenciado." & FONTTYPE_INFO)
                Call SendData(SendTarget.toIndex, TIndex, 0, "||Has Sido Silenciado" & FONTTYPE_INFO)
            Else
                UserList(TIndex).flags.Silenciado = 0
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El Usuario Ha sido DesSilenciado." & FONTTYPE_INFO)
                Call LogGM(UserList(UserIndex).Name, "/DESsilenciar " & UserList(TIndex).Name)

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 9)) = "/LIBERAR " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 9)
            TIndex = NameIndex(rData)

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El jugador no esta online." & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(TIndex).Counters.Pena > 0 Then
                UserList(TIndex).Counters.Pena = 0
                Call SendData(SendTarget.toIndex, TIndex, 0, "||El gm te ha liberado." & FONTTYPE_Motd5)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario liberado." & FONTTYPE_INFO)
                Call WarpUserChar(TIndex, 48, 75, 65, False)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no esta en la carcel." & FONTTYPE_INFO)
                Exit Sub

            End If

        End If

    Case 9    '<<<<<<<<<<<<<Comandos 9>>>>>>>>>>>>>>
        If UCase$(Left$(rData, 7)) = "/PERDON" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
            rData = Right$(rData, Len(rData) - 8)
            TIndex = NameIndex(rData)

            Call VolverCiudadano(TIndex)
            Exit Sub

        End If

        If UCase$(Left$(rData, 7)) = "/CONDEN" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
            rData = Right$(rData, Len(rData) - 8)
            TIndex = NameIndex(rData)

            If TIndex > 0 Then Call VolverCriminal(TIndex)
            Exit Sub

        End If

        If UCase$(Left$(rData, 7)) = "/RAJAR " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
            rData = Right$(rData, Len(rData) - 7)
            TIndex = NameIndex(UCase$(rData))

            If TIndex > 0 Then
                Call ResetFacciones(TIndex)

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 14)) = "/CASTIGORETOS " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 14)

            Arg1 = rData
            TIndex = NameIndex(rData)

            If TIndex > 0 Then
                If UserList(TIndex).Stats.PuntosRetos > 0 Then
                    Call SendData(toIndex, UserIndex, 0, "||Los puntos retos del usuario " & UserList(TIndex).Name & " han sido castigados." & FONTTYPE_INFO)
                    UserList(TIndex).Stats.PuntosRetos = 0
                Else
                    Call SendData(toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " no tiene puntos retos." & FONTTYPE_INFO)
                End If
            Else

                If FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) = False Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
                Else
                    If GetVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSRETOS") > 0 Then
                        Call SendData(toIndex, UserIndex, 0, "||Los puntos retos del usuario " & Arg1 & " han sido castigados." & FONTTYPE_INFO)
                        Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSRETOS", "0")
                    Else
                        Call SendData(toIndex, UserIndex, 0, "||El usuario " & Arg1 & " no tiene puntos retos." & FONTTYPE_INFO)
                    End If
                End If

            End If

            Exit Sub
        End If

        If UCase$(Left$(rData, 15)) = "/CASTIGOPUNTOS " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 15)

            Arg1 = rData
            TIndex = NameIndex(rData)

            If TIndex > 0 Then
                If UserList(TIndex).Stats.PuntosDuelos > 0 Then
                    Call SendData(toIndex, UserIndex, 0, "||Los puntos del usuario " & UserList(TIndex).Name & " han sido castigados." & FONTTYPE_INFO)
                    UserList(TIndex).Stats.PuntosDuelos = 0
                Else
                    Call SendData(toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " no tiene puntos." & FONTTYPE_INFO)
                End If
            Else

                If FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) = False Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
                Else
                    If GetVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSDUELOS") > 0 Then
                        Call SendData(toIndex, UserIndex, 0, "||Los puntos del usuario " & Arg1 & " han sido castigados." & FONTTYPE_INFO)
                        Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSDUELOS", "0")
                    Else
                        Call SendData(toIndex, UserIndex, 0, "||El usuario " & Arg1 & " no tiene puntos." & FONTTYPE_INFO)
                    End If
                End If

            End If

            Exit Sub
        End If

        If UCase$(Left$(rData, 15)) = "/CASTIGOTORNEO " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 15)

            Arg1 = rData
            TIndex = NameIndex(rData)

            If TIndex > 0 Then
                If UserList(TIndex).Stats.PuntosTorneo > 0 Then
                    Call SendData(toIndex, UserIndex, 0, "||Los puntos torneo del usuario " & UserList(TIndex).Name & " han sido castigados." & FONTTYPE_INFO)
                    UserList(TIndex).Stats.PuntosTorneo = 0
                Else
                    Call SendData(toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " no tiene puntos torneo." & FONTTYPE_INFO)
                End If
            Else

                If FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) = False Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
                Else
                    If GetVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSTORNEO") > 0 Then
                        Call SendData(toIndex, UserIndex, 0, "||Los puntos torneo del usuario " & Arg1 & " han sido castigados." & FONTTYPE_INFO)
                        Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSTORNEO", "0")
                    Else
                        Call SendData(toIndex, UserIndex, 0, "||El usuario " & Arg1 & " no tiene puntos torneo." & FONTTYPE_INFO)
                    End If
                End If

            End If

            Exit Sub
        End If

        If UCase$(Left(rData, 13)) = "/CASTIGOCLAN " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 13)

            Arg1 = rData
            TIndex = NameIndex(rData)

            If TIndex > 0 Then
                If UserList(TIndex).Clan.PuntosClan > 0 Then
                    Call SendData(toIndex, UserIndex, 0, "||Los puntos clan del usuario " & UserList(TIndex).Name & " han sido castigados." & FONTTYPE_INFO)
                    UserList(TIndex).Clan.PuntosClan = 0
                Else
                    Call SendData(toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " no tiene puntos clan." & FONTTYPE_INFO)
                End If
            Else

                If FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) = False Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
                Else
                    If GetVar(CharPath & UCase$(Arg1) & ".chr", "GUILD", "PUNTOSCLAN") > 0 Then
                        Call SendData(toIndex, UserIndex, 0, "||Los puntos clan del usuario " & Arg1 & " han sido castigados." & FONTTYPE_INFO)
                        Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "GUILD", "PUNTOSCLAN", "0")
                    Else
                        Call SendData(toIndex, UserIndex, 0, "||El usuario " & Arg1 & " no tiene puntos clan." & FONTTYPE_INFO)
                    End If
                End If

            End If

            Exit Sub
        End If

        If UCase$(Left$(rData, 14)) = "/CASTIGOTODOS " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 14)

            Arg1 = rData
            TIndex = NameIndex(rData)

            If TIndex > 0 Then

                If UserList(TIndex).Stats.PuntosRetos > 0 Then
                    UserList(TIndex).Stats.PuntosRetos = 0
                    tInt = tInt + 1
                End If

                If UserList(TIndex).Stats.PuntosDuelos > 0 Then
                    UserList(TIndex).Stats.PuntosDuelos = 0
                    tInt = 1
                End If

                If UserList(TIndex).Stats.PuntosTorneo > 0 Then
                    UserList(TIndex).Stats.PuntosTorneo = 0
                    tInt = 1
                End If

                If UserList(TIndex).Clan.PuntosClan > 0 Then
                    UserList(TIndex).Clan.PuntosClan = 0
                    tInt = tInt + 1
                End If

                If tInt > 0 Then
                    Call SendData(toIndex, UserIndex, 0, "||Todos los puntos del usuario " & UserList(TIndex).Name & " han sido castigados." & FONTTYPE_INFO)
                Else
                    Call SendData(toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " no tiene puntos." & FONTTYPE_INFO)
                End If

            Else

                If FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) = False Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
                Else
                    If GetVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSRETOS") > 0 Then
                        Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSRETOS", "0")
                        tInt = tInt + 1
                    End If
                End If

                If FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) = False Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
                Else
                    If GetVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSDUELOS") > 0 Then
                        Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSDUELOS", "0")
                        tInt = tInt + 1
                    End If
                End If

                If FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) = False Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
                Else
                    If GetVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSTORNEO") > 0 Then
                        Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSTORNEO", "0")
                        tInt = tInt + 1
                    End If
                End If

                If FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) = False Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
                Else
                    If GetVar(CharPath & UCase$(Arg1) & ".chr", "GUILD", "PUNTOSCLAN") > 0 Then
                        Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "GUILD", "PUNTOSCLAN", "0")
                        tInt = tInt + 1
                    End If
                End If

                If tInt > 0 Then
                    Call SendData(toIndex, UserIndex, 0, "||Todos los puntos del usuario " & Arg1 & " han sido castigados." & FONTTYPE_INFO)
                Else
                    Call SendData(toIndex, UserIndex, 0, "||El usuario " & Arg1 & " no tiene puntos." & FONTTYPE_INFO)
                End If

            End If

            Exit Sub
        End If

    Case 10    '<<<<<<<<<<<<<Comandos 10>>>>>>>>>>>>>>
        If UCase$(Left$(rData, 5)) = "/BAN " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
            rData = Right$(rData, Len(rData) - 5)
            tStr = ReadField(2, rData, Asc("@"))    ' NICK
            TIndex = NameIndex(tStr)
            Name = ReadField(1, rData, Asc("@"))    ' MOTIVO

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_TALK)

                If FileExist(CharPath & tStr & ".chr", vbNormal) Then
                    tLong = UserDarPrivilegioLevel(tStr)

                    If tLong > UserList(UserIndex).flags.Privilegios Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estás loco??! No podés banear a alguien de mayor jerarquia que vos!" & _
                                                                        FONTTYPE_INFO)
                        Exit Sub

                    End If

                    If GetVar(CharPath & tStr & ".chr", "FLAGS", "Ban") <> "0" Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje ya ha sido baneado anteriormente." & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Call LogBanFromName(tStr, UserIndex, Name)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||AOMania> El GM & " & UserList(UserIndex).Name & "baneó a " & tStr & "." & FONTTYPE_SERVER)

                    'ponemos el flag de ban a 1
                    Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
                    'ponemos la pena
                    tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
                    Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & _
                                                                                   " Lo Baneó por el siguiente motivo: " & LCase$(Name) & " " & Date & " " & Time)

                    If tLong > 0 Then
                        UserList(UserIndex).flags.Ban = 1
                        Call CloseSocket(UserIndex)
                        Call SendData(SendTarget.ToAdmins, 0, 0, "||" & " El gm " & UserList(UserIndex).Name & _
                                                               " fue baneado por el propio servidor por intentar banear a otro admin." & FONTTYPE_FIGHT)

                    End If

                    Call LogGM(UserList(UserIndex).Name, "BAN a " & tStr)
                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El pj " & tStr & " no existe." & FONTTYPE_INFO)

                End If

            Else

                If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes banear a al alguien de mayor jerarquia." & FONTTYPE_INFO)
                    Exit Sub

                End If

                Call LogBan(TIndex, UserIndex, Name)    'BASS Ha baneado a testban

                Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " Ha baneado a " & UserList(TIndex).Name & _
                                                         FONTTYPE_Motd4)

                'Ponemos el flag de ban a 1
                UserList(TIndex).flags.Ban = 1

                If UserList(TIndex).flags.Privilegios > PlayerType.User Then
                    UserList(UserIndex).flags.Ban = 1
                    Call CloseSocket(UserIndex)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " banned by the server por bannear un Administrador." & _
                                                             FONTTYPE_FIGHT)

                End If

                Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(TIndex).Name)

                'ponemos el flag de ban a 1
                Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
                'ponemos la pena
                tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
                Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
                Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & " Lo Baneó Debido a: " & LCase$( _
                                                                                 Name) & " " & Date & " " & Time)

                Call CloseSocket(TIndex)

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 7)) = "/UNBAN " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
            rData = Right$(rData, Len(rData) - 7)

            rData = Replace(rData, "\", "")
            rData = Replace(rData, "/", "")

            If Not FileExist(CharPath & rData & ".chr", vbNormal) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Charfile inexistente (no use +)" & FONTTYPE_INFO)
                Exit Sub

            End If

            Call UnBan(rData)

            'penas
            i = val(GetVar(CharPath & rData & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & rData & ".chr", "PENAS", "Cant", i + 1)
            Call WriteVar(CharPath & rData & ".chr", "PENAS", "P" & i + 1, LCase$(UserList(UserIndex).Name) & " Lo unbaneó. " & Date & " " & Time)

            Call LogGM(UserList(UserIndex).Name, "/UNBAN a " & rData)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & rData & " desbaneado." & FONTTYPE_INFO)

            Exit Sub

        End If

        If UCase$(Left$(rData, 10)) = "/EJECUTAR " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

            rData = Right$(rData, Len(rData) - 10)
            TIndex = NameIndex(rData)

            If UserList(TIndex).flags.Privilegios > PlayerType.User Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, _
                              "||Osea, yo te dejaria pero es un viaje, mira si se caen altos items anda a saber, mejor qedate ahi y no intentes ejecutar mas gms la re puta qe te pario." _
                            & FONTTYPE_EJECUCION)
                Exit Sub

            End If

            If TIndex > 0 Then

                Call UserDie(TIndex)

                If UserList(TIndex).pos.Map = 1 Then
                    Call TirarTodo(TIndex)

                End If

                Call SendData(SendTarget.toall, 0, 0, "||El GameMaster " & UserList(UserIndex).Name & " ha ejecutado a " & UserList(TIndex).Name & _
                                                      FONTTYPE_EJECUCION)
                Call LogGM(UserList(UserIndex).Name, " ejecuto a " & UserList(TIndex).Name)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No está online" & FONTTYPE_EJECUCION)

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 12)) = "/BORRARPENA " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            '/borrarpena pj pena
            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

            rData = Right$(rData, Len(rData) - 12)

            Name = ReadField(1, rData, Asc("@"))
            tStr = ReadField(2, rData, Asc("@"))

            If Name = "" Or tStr = "" Or Not IsNumeric(tStr) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Utilice /borrarpj Nick@NumeroDePena" & FONTTYPE_INFO)
                Exit Sub

            End If

            Name = Replace(Name, "\", "")
            Name = Replace(Name, "/", "")

            If FileExist(CharPath & Name & ".chr", vbNormal) Then
                rData = GetVar(CharPath & Name & ".chr", "PENAS", "P" & val(tStr))
                Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & val(tStr), LCase$(UserList(UserIndex).Name) & ": <Pena borrada> " & Date & " " & _
                                                                                  Time)

            End If

            Call LogGM(UserList(UserIndex).Name, " borro la pena: " & tStr & "-" & rData & " de " & Name)
            Exit Sub

        End If

        If UCase$(Left$(rData, 7)) = "/PJBAN " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 7)
            TIndex = NameIndex(ReadField(2, rData, Asc("@")))
            Name = ReadField(1, rData, Asc("@"))

            Arg1 = ReadField(2, rData, Asc("@"))
            Arg2 = CharPath & Left$(Arg1, 1) & "\" & Arg1 & ".chr"

            If TIndex <= 0 Then
                If PersonajeExiste(Arg1) Then
                    Dim CANALBAN As Integer
                    CANALBAN = FreeFile    ' obtenemos un canal
                    Open App.Path & "\logs\BAN\" & GetVar(Arg2, "INIT", "LastSerie") & ".dat" For Append As #CANALBAN
                    Print #CANALBAN, "PJ:" & Arg1 & " Fecha:" & Date & " GM:" & UserList(UserIndex).Name & " Razón:" & Name
                    Close #CANALBAN
                    Call SendData(toIndex, UserIndex, 0, "||Ban directo a la ficha de " & Arg1 & "." & "´" & FONTTYPE_INFO)
                    Call WriteVar(CharPath & Left$(Arg1, 1) & "\" & Arg1 & ".chr", "FLAGS", "Ban", 1)

                    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Arg1, "BannedBy", UserList(UserIndex).Name)
                    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Arg1, "Reason", Name)
                    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Arg1, "Fecha", Date)

                Else
                    Call SendData(toIndex, UserIndex, 0, "||Ese Pj no existe." & "´" & FONTTYPE_INFO)
                End If
                Exit Sub
            End If

            Exit Sub
        End If

    Case 11    '<<<<<<<<<<<<<Comandos 11>>>>>>>>>>>>>>
        If UCase$(Left$(rData, 5)) = "/MOD " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            rData = UCase$(Right$(rData, Len(rData) - 5))
            tStr = Replace$(ReadField(1, rData, 32), "+", " ")
            TIndex = NameIndex(tStr)

            If LCase$(tStr) = "yo" Then
                TIndex = UserIndex

            End If

            Arg1 = ReadField(2, rData, 32)
            Arg2 = ReadField(3, rData, 32)
            Arg3 = ReadField(4, rData, 32)
            Arg4 = ReadField(5, rData, 32)

            If UserList(UserIndex).flags.EsRolesMaster Then

                Select Case UserList(UserIndex).flags.Privilegios

                Case PlayerType.Consejero

                    ' Los RMs consejeros sólo se pueden editar su head, body y exp
                    If NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                    If Arg1 <> "BODY" And Arg1 <> "HEAD" And Arg1 <> "LEVEL" Then Exit Sub

                Case PlayerType.SemiDios

                    ' Los RMs sólo se pueden editar su level y el head y body de cualquiera
                    If Arg1 = "EXP" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                    If Arg1 <> "BODY" And Arg1 <> "HEAD" Then Exit Sub

                Case PlayerType.Dios

                    ' Si quiere modificar el level sólo lo puede hacer sobre sí mismo
                    If Arg1 = "NIVEL" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                    If Arg1 = "LEVEL" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                    If Arg1 = "ORO" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub

                    ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                    If Arg1 <> "BODY" And Arg1 <> "HEAD" And Arg1 <> "CIU" And Arg1 <> "CRI" And Arg1 <> "CLASE" And Arg1 <> "SKILLS" And Arg1 <> "RAZA" Then Exit Sub

                End Select

            ElseIf UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then   'Si no es RM debe ser dios para poder usar este comando
                Exit Sub

            End If

            Select Case Arg1

            Case "VIDA"
                Dim MaxVida As Long
                Dim ChangeVida As Long

                MaxVida = "32000"
                ChangeVida = ReadField(3, rData, 32)

                If ChangeVida > MaxVida Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No se ha cambiado el valor, por que has superado la maxima vida. (Max: " & _
                                                                    MaxVida & ")" & FONTTYPE_INFO)
                    Exit Sub

                End If

                If ChangeVida <= MaxVida Then

                    If UserList(TIndex).Stats.MaxHP < ChangeVida Then
                        UserList(TIndex).Stats.MinHP = ChangeVida
                    End If

                    UserList(TIndex).Stats.MinHP = ChangeVida
                    UserList(TIndex).Stats.MaxHP = ChangeVida
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has cambiado la vida maxima del personaje " & tStr & " ahora es: " & _
                                                                    ChangeVida & FONTTYPE_INFO)
                    Call SendData(SendTarget.toIndex, TIndex, 0, "MXVID" & ChangeVida)
                    Call EnviarHP(UserIndex)

                End If

                Exit Sub

            Case "MANA"
                Dim MaxMana As Long
                Dim ChangeMana As Long

                MaxMana = "32000"
                ChangeMana = val(ReadField(3, rData, 32))

                If ChangeMana > MaxMana Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No se ha cambiado el valor, por que has superado la maxima mana. (Max: " & _
                                                                    MaxMana & ")" & FONTTYPE_INFO)
                    Exit Sub

                End If

                If ChangeMana <= MaxMana Then

                    UserList(TIndex).Stats.MinMAN = ChangeMana
                    UserList(TIndex).Stats.MaxMAN = ChangeMana
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has cambiado la mana maxima del personaje " & tStr & " ahora es: " & _
                                                                    ChangeMana & FONTTYPE_INFO)
                    Call SendData(SendTarget.toIndex, TIndex, 0, "MXMAN" & ChangeMana)
                    Call EnviarMn(UserIndex)
                End If

                Exit Sub

            Case "NIVEL"
                Dim MassNivel As Long
                Dim ResultMassNivel As Long
                Dim ExpMAX As Long
                Dim ExpMIN As Long
                Dim ExpLvl As Long
                Dim XN As Long

                If TIndex <= 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline: " & tStr & FONTTYPE_INFO)
                    Exit Sub

                End If

                MassNivel = val(Arg2)
                ExpMAX = UserList(TIndex).Stats.ELU
                ExpMIN = UserList(TIndex).Stats.Exp

                If Not IsNumeric(Arg2) Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El Nivel debe ser númerica." & FONTTYPE_GUILD)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis: /Mod " & UserList(TIndex).Name & " NIVEL 2" & FONTTYPE_GUILD)
                    Exit Sub

                End If

                If ExpMAX = 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " tiene el nivel máximo." & _
                                                                    FONTTYPE_INFO)
                    Exit Sub

                End If

                For XN = 1 To MassNivel

                    If ExpMAX = "0" Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & _
                                                                      " subió de nivel pero llego al nivel máximo." & FONTTYPE_INFO)
                        Exit For

                    End If

                    ExpMAX = UserList(TIndex).Stats.ELU
                    ExpMIN = UserList(TIndex).Stats.Exp

                    ResultMassNivel = ExpMAX - ExpMIN
                    UserList(TIndex).Stats.Exp = UserList(TIndex).Stats.Exp + ResultMassNivel

                    Call EnviarExp(TIndex)
                    Call CheckUserLevel(TIndex)

                Next XN

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " ha subido de nivel." & FONTTYPE_Motd1)

                Exit Sub

            Case "ORO"

                If TIndex <= 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                    Exit Sub
                End If

                If Left$(Arg2, 1) = "-" Then

                    If UserList(TIndex).Stats.GLD = 0 Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no tiene oro!!" & FONTTYPE_INFO)
                        Exit Sub
                    End If

                    UserList(TIndex).Stats.GLD = UserList(TIndex).Stats.GLD - val(mid(Arg2, 2))
                    Call EnviarOro(TIndex)

                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has quitado el oro de " & UserList(TIndex).Name & " con resta de: " & val(Arg2) & " de oro." & FONTTYPE_INFO)
                    Call SendData(SendTarget.toIndex, TIndex, 0, "||El GM " & UserList(UserIndex).Name & " te ha quitado: " & val(Arg2) & " de oro." & FONTTYPE_INFO)
                    Exit Sub

                Else

                    If val(Arg2) > MaxOro Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has superado el limite de maximo oro: " & MaxOro & FONTTYPE_INFO)
                        Exit Sub
                    End If

                    UserList(TIndex).Stats.GLD = UserList(TIndex).Stats.GLD + val(Arg2)
                    Call EnviarOro(TIndex)

                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has aumentado el oro de " & UserList(TIndex).Name & " con suma de: " & val(Arg2) & " de oro." & FONTTYPE_INFO)
                    Call SendData(SendTarget.toIndex, TIndex, 0, "||El GM " & UserList(UserIndex).Name & " te ha dado: " & val(Arg2) & " de oro." & FONTTYPE_INFO)
                End If

            Case "EXP"

                If TIndex <= 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                    Exit Sub

                End If

                If UserList(TIndex).Stats.Exp + val(Arg2) > UserList(TIndex).Stats.ELU Then
                    UserList(TIndex).Stats.Exp = UserList(TIndex).Stats.Exp + val(Arg2)
                    Call CheckUserLevel(TIndex)
                Else
                    UserList(TIndex).Stats.Exp = val(Arg2)

                End If

                Call EnviarExp(TIndex)
                Exit Sub

            Case "BODY"

                If TIndex <= 0 Then
                    Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "INIT", "Body", Arg2)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
                    Exit Sub

                End If


                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(TIndex).pos.Map, TIndex, val(Arg2), UserList(TIndex).char.Head, UserList( _
                                                                                                                                  TIndex).char.heading, UserList(TIndex).char.WeaponAnim, UserList(TIndex).char.ShieldAnim, UserList(TIndex).char.CascoAnim, _
                                    UserList(TIndex).char.Alas)

                Exit Sub

            Case "HEAD"

                If TIndex <= 0 Then
                    Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "INIT", "Head", Arg2)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
                    Exit Sub

                End If

                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(TIndex).pos.Map, TIndex, UserList(TIndex).char.Body, val(Arg2), UserList( _
                                                                                                                                  TIndex).char.heading, UserList(TIndex).char.WeaponAnim, UserList(TIndex).char.ShieldAnim, UserList(TIndex).char.CascoAnim, _
                                    UserList(TIndex).char.Alas)
                Exit Sub

            Case "CRI"

                If TIndex <= 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                    Exit Sub

                End If

                UserList(TIndex).Faccion.CriminalesMatados = val(Arg2)
                Exit Sub

            Case "CIU"

                If TIndex <= 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                    Exit Sub

                End If

                UserList(TIndex).Faccion.CiudadanosMatados = val(Arg2)
                Exit Sub

            Case "CLASE"

                If TIndex <= 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline: " & tStr & FONTTYPE_INFO)
                    Exit Sub
                End If

                If UserList(TIndex).Clase = UCase$(Left$(Arg2, 1)) & UCase$(mid$(Arg2, 2)) Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||La clase de: " & tStr & " no ha cambiado porque ya es: " & UCase$(Left$(Arg2, 1)) & UCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                    Exit Sub
                End If

                If Len(Arg2) > 1 Then
                    UserList(TIndex).Clase = UCase$(Left$(Arg2, 1)) & UCase$(mid$(Arg2, 2))
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Clase cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                Else
                    UserList(TIndex).Clase = UCase$(Arg2)
                End If

            Case "RAZA"

                If TIndex <= 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline: " & tStr & FONTTYPE_INFO)
                    Exit Sub
                End If

                If UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) = UserList(TIndex).Raza Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||La raza de: " & tStr & " ya no ha cambiado porque ya es: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                    Exit Sub
                End If

                Select Case UCase$(Arg2)

                Case "HUMANO"
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                    UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                    Call DarCuerpoDesnudo(TIndex)

                Case "ENANO"
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                    UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                    Call DarCuerpoDesnudo(TIndex)

                Case "HOBBIT"
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                    UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                    Call DarCuerpoDesnudo(TIndex)

                Case "ELFO"
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                    UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                    Call DarCuerpoDesnudo(TIndex)

                Case "ELFO OSCURO"
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                    UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                    Call DarCuerpoDesnudo(TIndex)

                Case "LICANTROPO"
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                    UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                    Call DarCuerpoDesnudo(TIndex)

                Case "GNOMO"
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                    UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                    Call DarCuerpoDesnudo(TIndex)

                Case "ORCO"
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                    UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                    Call DarCuerpoDesnudo(TIndex)

                Case "VAMPIRO"
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                    UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                    Call DarCuerpoDesnudo(TIndex)

                Case "CICLOPE"
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                    UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                    Call DarCuerpoDesnudo(TIndex)

                Case Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||La raza: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & " no existe." & FONTTYPE_INFO)
                End Select

            Case "SKILLS"

                For LoopC = 1 To NUMSKILLS
                    If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg2) Then n = LoopC
                Next LoopC

                If n = 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Skill Inexistente!" & FONTTYPE_INFO)
                    Exit Sub
                End If

                If TIndex = 0 Then
                    Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "Skills", "SK" & n, Arg3)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
                Else
                    UserList(TIndex).Stats.UserSkills(n) = val(Arg3)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has cambiado la skill de " & SkillsNames(n) & " a " & UserList(TIndex).Name & " por: " & val(Arg3) & FONTTYPE_INFO)
                    Call SendData(SendTarget.toIndex, TIndex, 0, "||GM " & UserList(UserIndex).Name & " te ha cambiado el valor de la skill " & SkillsNames(n) & " a: " & val(Arg3) & FONTTYPE_INFO)
                End If

                Exit Sub

            Case "SKILLSLIBRES"
                Dim SLName As String
                Dim SLSkills As Integer
                Dim SLResult As Integer

                If Left(Arg2, 1) = "-" Then

                    If TIndex = 0 Then
                        SLName = ReadField(1, rData, 32)

                        If Not FileExist(CharPath & UCase(SLName) & ".chr") Then
                            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje: " & SLName & " no existe." & FONTTYPE_INFO)
                            Exit Sub
                        End If

                        SLSkills = GetVar(CharPath & UCase$(SLName) & ".chr", "STATS", "SkillptsLibres")

                        SLResult = SLSkills - mid(Arg2, 2)

                        If SLResult < 0 Then
                            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El cambio no se ha podido efectuar:" & SLResult & FONTTYPE_INFO)
                            Exit Sub
                        Else
                            Call WriteVar(CharPath & UCase$(SLName) & ".chr", "STATS", "SkillPtsLibres", SLResult)
                            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El charfile de " & SLName & " ha cambiado su SkillsLibres de " & SLSkills & " a " & SLResult & FONTTYPE_INFO)
                            Exit Sub
                        End If

                    Else
                        SLName = UserList(TIndex).Name
                        SLSkills = UserList(TIndex).Stats.SkillPts

                        SLResult = SLSkills - mid(Arg2, 2)

                        If SLResult < 0 Then
                            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El cambio no se ha podido efectuar:" & SLResult & FONTTYPE_INFO)
                            Exit Sub
                        Else
                            UserList(TIndex).Stats.SkillPts = SLResult
                            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Las skills de " & SLName & " han sido modificadas." & FONTTYPE_INFO)
                            Call EnviarSkills(TIndex)
                            Exit Sub
                        End If
                    End If


                Else    'Parte donde Suma

                    If TIndex = 0 Then

                        SLName = ReadField(1, rData, 32)

                        If Not FileExist(CharPath & UCase(SLName) & ".chr") Then
                            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje: " & SLName & " no existe." & FONTTYPE_INFO)
                            Exit Sub
                        End If

                        SLSkills = GetVar(CharPath & UCase$(SLName) & ".chr", "STATS", "SkillptsLibres")

                        SLResult = SLSkills + Arg2

                        Call WriteVar(CharPath & UCase$(SLName) & ".chr", "STATS", "SkillPtsLibres", SLResult)
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||El charfile de " & SLName & " ha cambiado su SkillsLibres de " & SLSkills & " a " & SLResult & FONTTYPE_INFO)
                        Exit Sub

                    Else
                        SLName = UserList(TIndex).Name
                        SLSkills = UserList(TIndex).Stats.SkillPts
                        SLResult = SLSkills + Arg2

                        UserList(TIndex).Stats.SkillPts = SLResult
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Las skills de " & SLName & " han sido modificadas." & FONTTYPE_INFO)
                        Call EnviarSkills(TIndex)
                        Exit Sub

                    End If

                End If

                Exit Sub

            Case Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Sintaxis incorrecto" & FONTTYPE_GUILD)
                Call SendData(SendTarget.toIndex, UserIndex, 0, _
                              "||Comando: /MOD <Nick/yo> <NIVEL/SKILLS/SKILLSLIBRES/ORO/CIU/CRI/EXP/BODY/HEAD> <VALOR>" & FONTTYPE_GUILD)
                Exit Sub

            End Select

            Exit Sub

        End If

        If UCase$(Left$(rData, 7)) = "/SUBIR " Then

            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            rData = Right$(rData, Len(rData) - 7)

            TIndex = NameIndex(ReadField(1, rData, 32))

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline: " & tStr & FONTTYPE_INFO)
                Exit Sub

            End If

            MassNivel = ReadField(2, rData, 32)
            ExpMAX = UserList(TIndex).Stats.ELU
            ExpMIN = UserList(TIndex).Stats.Exp

            If Not IsNumeric(MassNivel) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El Nivel debe ser númerica." & FONTTYPE_GUILD)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis: /SUBIR " & UserList(TIndex).Name & " 2" & FONTTYPE_GUILD)
                Exit Sub

            End If

            If ExpMAX = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " tiene el nivel máximo." & _
                                                                FONTTYPE_INFO)
                Exit Sub

            End If

            For XN = 1 To MassNivel

                If ExpMAX = "0" Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & _
                                                                  " subió de nivel pero llego al nivel máximo." & FONTTYPE_INFO)
                    Exit For

                End If

                ExpMAX = UserList(TIndex).Stats.ELU
                ExpMIN = UserList(TIndex).Stats.Exp

                ResultMassNivel = ExpMAX - ExpMIN
                UserList(TIndex).Stats.Exp = UserList(TIndex).Stats.Exp + ResultMassNivel

                Call EnviarExp(TIndex)
                Call CheckUserLevel(TIndex)

            Next XN

            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " ha subido de nivel." & FONTTYPE_Motd1)

            Exit Sub
        End If

        If UCase$(Left$(rData, 13)) = "/CAMBIARMAIL " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

            rData = Right$(rData, Len(rData) - 13)
            tStr = ReadField(1, rData, Asc("-"))

            If tStr = "" Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||usar /CAMBIARMAIL <pj>-<nuevomail>" & FONTTYPE_GUILD)
                Exit Sub

            End If

            TIndex = NameIndex(tStr)

            If TIndex > 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario esta online, no se puede si esta online" & FONTTYPE_INFO)
                Exit Sub

            End If

            Arg1 = ReadField(2, rData, Asc("-"))

            If Arg1 = "" Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||usar /CAMBIARMAIL <pj>-<nuevomail>" & FONTTYPE_INFO)
                Exit Sub

            End If

            If Not FileExist(CharPath & tStr & ".chr") Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No existe el charfile " & CharPath & tStr & ".chr" & FONTTYPE_INFO)
            Else
                Call WriteVar(CharPath & tStr & ".chr", "CONTACTO", "Email", Arg1)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Email de " & tStr & " cambiado a: " & Arg1 & FONTTYPE_INFO)

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 13)) = "/CAMBIARNICK " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

            rData = Right$(rData, Len(rData) - 13)
            tStr = ReadField(1, rData, Asc("@"))
            Arg1 = ReadField(2, rData, Asc("@"))

            If tStr = "" Or Arg1 = "" Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usar: /CAMBIARNICK NiCK@NUEVO NICK" & FONTTYPE_INFO)
                Exit Sub

            End If

            TIndex = NameIndex(tStr)

            If TIndex > 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El Pj esta online, debe salir para el cambio" & FONTTYPE_WARNING)
                Exit Sub

            End If

            If FileExist(CharPath & UCase(tStr) & ".chr") = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El pj " & tStr & " es inexistente " & FONTTYPE_INFO)
                Exit Sub

            End If

            Arg2 = GetVar(CharPath & UCase(tStr) & ".chr", "GUILD", "GUILDINDEX")

            If IsNumeric(Arg2) Then
                If CInt(Arg2) > 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El pj " & tStr & _
                                                                  " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido. " & FONTTYPE_INFO)
                    Exit Sub

                End If

            End If

            If FileExist(CharPath & UCase(Arg1) & ".chr") = False Then
                FileCopy CharPath & UCase(tStr) & ".chr", CharPath & UCase(Arg1) & ".chr"
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Transferencia exitosa" & FONTTYPE_INFO)
                Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
                'ponemos la pena
                tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
                Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
                Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & ": BAN POR Cambio de nick a " & _
                                                                                 UCase$(Arg1) & " " & Date & " " & Time)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El nick solicitado ya existe" & FONTTYPE_INFO)
                Exit Sub

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 11)) = "/RAJARCLAN " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
            rData = Right$(rData, Len(rData) - 11)
            tInt = modGuilds.m_EcharMiembroDeClan(UserIndex, rData, False)  'me da el guildindex

            If tInt = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| No pertenece a ningun clan o es fundador." & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Expulsado." & FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "|| " & rData & " ha sido expulsado del clan por los administradores del servidor" & _
                                                                  FONTTYPE_GUILD)

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 8)) = "/QUITAR " Then
            Dim QuitObjeto As Obj
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)


            rData = Right$(rData, Len(rData) - 8)

            TIndex = NameIndex(ReadField(1, rData, 32))
            QuitObjeto.ObjIndex = ReadField(2, rData, 32)
            QuitObjeto.Amount = ReadField(3, rData, 32)

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||ERROR: El usuario no esta conectado." & FONTTYPE_EJECUCION)
            Else
                Call QuitarObjetos(QuitObjeto.ObjIndex, QuitObjeto.Amount, TIndex)
            End If

            Exit Sub
        End If


        If UCase$(Left$(rData, 11)) = "/QUITARBOV " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            rData = Right$(rData, Len(rData) - 11)

            TIndex = NameIndex(ReadField(1, rData, 32))
            QuitObjeto.ObjIndex = ReadField(2, rData, 32)
            QuitObjeto.Amount = ReadField(3, rData, 32)

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||ERROR: El usuario no esta conectado." & FONTTYPE_EJECUCION)
            Else
                Call QuitarObjetosBov(QuitObjeto.ObjIndex, QuitObjeto.Amount, TIndex)
            End If

            Exit Sub
        End If

        If UCase$(Left(rData, 7)) = "/CLAVE " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 7)

            tStr = ReadField(1, rData, 32)
            TIndex = NameIndex(tStr)

            If TIndex > 0 Then
                Call SendData(toIndex, UserIndex, 0, "||El usuario debe desconectarse para realizar el cambio de clave." & FONTTYPE_INFO)
                Exit Sub
            End If

            If FileExist(CharPath & UCase$(tStr) & ".chr", vbNormal) Then

                If UCase$(GetVar(CharPath & tStr & ".chr", "CONTACTO", "Email")) <> UCase$(ReadField(2, rData, 32)) Then
                    Call SendData(toIndex, UserIndex, 0, "||El email no coincide." & FONTTYPE_INFO)
                    Exit Sub
                Else

                    For i = 1 To 5
                        tInt = RandomNumber(65, 90)
                        Arg1 = Arg1 + Chr$(tInt)
                    Next i

                    Call SendData(toIndex, UserIndex, 0, "||  Su email es:" & ReadField(2, rData, 32) & FONTTYPE_INFO)
                    Call SendData(toIndex, UserIndex, 0, "|| La Ultima Ip es:" & GetVar(CharPath & UCase$(tStr) & ".chr", "INIT", "LASTIP") & FONTTYPE_INFO)
                    Call SendData(toIndex, UserIndex, 0, "||La nueva clave es: " & Arg1 & FONTTYPE_INFO)
                    Arg1 = MD5String(Arg1)
                    Call WriteVar(CharPath & UCase$(tStr) & ".chr", "INIT", "PASSWORD", Arg1)
                    Exit Sub
                End If

            Else
                Call SendData(toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
            End If

            Exit Sub
        End If

        If UCase$(rData) = "/ECHARTODOSPJSS" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
            Call EcharPjsNoPrivilegiados
            Exit Sub

        End If

        If UCase$(rData) = "/ACTCOM" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If ComerciarAc = True Then
                ComerciarAc = False
                Call SendData(SendTarget.toall, 0, 0, "||Comercio entre usuarios desactivados!!." & FONTTYPE_CYAN)
            Else
                Call SendData(SendTarget.toall, 0, 0, "||Comercio entre usuarios activados!!." & FONTTYPE_CYAN)
                ComerciarAc = True

            End If

            Exit Sub

        End If

    Case 12    '<<<<<<<<<<<<<Comandos 12>>>>>>>>>>>>>>
        If UCase(Left(rData, 4)) = "/CT " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            '/ct mapa_dest x_dest y_dest
            rData = Right(rData, Len(rData) - 4)
            Mapa = ReadField(1, rData, 32)
            X = ReadField(2, rData, 32)
            Y = ReadField(3, rData, 32)

            If MapaValido(Mapa) = False Or InMapBounds(Mapa, X, Y) = False Then
                Exit Sub

            End If

            If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).OBJInfo.ObjIndex > 0 Then
                Exit Sub

            End If

            If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.Map > 0 Then
                Exit Sub

            End If

            If MapData(Mapa, X, Y).OBJInfo.ObjIndex > 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, Mapa, "||Hay un objeto en el piso en ese lugar" & FONTTYPE_INFO)
                Exit Sub

            End If

            Dim ET As Obj
            ET.Amount = 1
            ET.ObjIndex = 378

            Call MakeObj(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, ET, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList( _
                                                                                                                                       UserIndex).pos.Y - 1)

            MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.Map = Mapa
            MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.X = X
            MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.Y = Y

            Exit Sub

        End If

        'Destruir Teleport
        'toma el ultimo click
        If UCase(Left(rData, 3)) = "/DT" Then
            '/dt

            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            Mapa = UserList(UserIndex).flags.TargetMap
            X = UserList(UserIndex).flags.TargetX
            Y = UserList(UserIndex).flags.TargetY

            If ObjData(MapData(Mapa, X, Y).OBJInfo.ObjIndex).ObjType = eOBJType.otTELEPORT And MapData(Mapa, X, Y).TileExit.Map > 0 Then
                Call EraseObj(SendTarget.ToMap, 0, Mapa, MapData(Mapa, X, Y).OBJInfo.Amount, Mapa, X, Y)
                Call EraseObj(SendTarget.ToMap, 0, MapData(Mapa, X, Y).TileExit.Map, 1, MapData(Mapa, X, Y).TileExit.Map, MapData(Mapa, X, _
                                                                                                                                  Y).TileExit.X, MapData(Mapa, X, Y).TileExit.Y)
                MapData(Mapa, X, Y).TileExit.Map = 0
                MapData(Mapa, X, Y).TileExit.X = 0
                MapData(Mapa, X, Y).TileExit.Y = 0

            End If

            Exit Sub

        End If

        If UCase(rData) = "/BLOQ" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Blocked = 0 Then
                MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Blocked = 1
                Call Bloquear(SendTarget.ToMap, UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, _
                              UserList(UserIndex).pos.Y, 1)
            Else
                MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Blocked = 0
                Call Bloquear(SendTarget.ToMap, UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, _
                              UserList(UserIndex).pos.Y, 0)

            End If

            Exit Sub

        End If


    Case 13    '<<<<<<<<<<<<<Comandos 13>>>>>>>>>>>>>>

        If UCase$(Left(rData, 9)) = "/CREARASEDIO " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 9)
            If Len(ReadField(1, rData, Asc("@"))) = 0 Or Len(ReadField(2, rData, Asc("@"))) = 0 Or Len(ReadField(3, rData, Asc("@"))) = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Formato invalido, el formato deberia ser /CREARASEDIO SLOTS@COSTE@TIEMPO." & FONTTYPE_INFO)
            Else
                Call modAsedio.Iniciar_Asedio(UserIndex, val(ReadField(1, rData, Asc("@"))), val(ReadField(2, rData, Asc("@"))), val(ReadField(3, rData, Asc("@"))))
            End If
        End If
        If UCase$(Left(rData, 13)) = "/CANCELARSEDIO" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            Call modAsedio.CancelAsedio
        End If


        If UCase(Left(rData, 9)) = "/UNBANIP " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

            rData = Right(rData, Len(rData) - 9)

            If BanIpQuita(rData) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||La IP """ & rData & """ se ha quitado de la lista de bans." & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||La IP """ & rData & """ NO se encuentra en la lista de bans." & FONTTYPE_INFO)

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 7)) = "/AVISO " Then

            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            rData = Right$(rData, Len(rData) - 7)

            TIndex = NameIndex(rData)

            If TIndex > 0 Then

                If UserList(TIndex).flags.Privilegios > PlayerType.User Then
                    Call CloseSocket(UserIndex)
                Else
                    Call SendData(SendTarget.toall, 0, 0, "||" & UserList(UserIndex).Name & " se va a sacrificar a " & UserList(TIndex).Name & " en NIX." & FONTTYPE_GUILD)
                    Call WarpUserChar(UserIndex, "34", "73", "58", True)
                    Call WarpUserChar(TIndex, "34", "74", "59", True)
                End If

            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario Offline" & FONTTYPE_INFO)
            End If

            Exit Sub
        End If

        If UCase$(Left$(rData, 11)) = "/SACRIFICA " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 11)

            TIndex = NameIndex(rData)

            If TIndex > 0 Then

                If UserList(TIndex).flags.Privilegios > PlayerType.User Then
                    Call CloseSocket(UserIndex)
                Else
                    Dim ObjSC As Obj
                    ObjSC.Amount = "1"
                    ObjSC.ObjIndex = "1043"

                    Call MeterItemEnInventario(TIndex, ObjSC)


                    Call SendData(SendTarget.toall, 0, 0, "||" & UserList(UserIndex).Name & " sacrificó por macrear a " & UserList(TIndex).Name & FONTTYPE_SERVER)
                    Call TirarTodo(TIndex)
                    UserList(TIndex).flags.Ban = 1
                    Call CloseSocket(TIndex)



                End If

            Else

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario Offline." & FONTTYPE_INFO)

            End If


            Exit Sub
        End If

        If UCase(Left(rData, 9)) = "/BANCLAN " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
            rData = Trim(Right(rData, Len(rData) - 9))

            If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
                Exit Sub

            End If

            Call SendData(SendTarget.toall, 0, 0, "|| " & UserList(UserIndex).Name & " banned al clan " & UCase$(rData) & FONTTYPE_FIGHT)

            Call LogGM(UserList(UserIndex).Name, "BANCLAN a " & rData)

            tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))

            For i = 1 To tInt
                tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
                'tstr es la victima
                Call Ban(tStr, "Administracion del servidor", "Clan Banned")
                TIndex = NameIndex(tStr)

                If TIndex > 0 Then

                    UserList(TIndex).flags.Ban = 1
                    Call CloseSocket(TIndex)

                End If

                Call SendData(SendTarget.toall, 0, 0, "||   " & tStr & "<" & rData & "> ha sido expulsado del servidor." & FONTTYPE_FIGHT)

                Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")

                n = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
                Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", n + 1)
                Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & n + 1, LCase$(UserList(UserIndex).Name) & ": BAN AL CLAN: " & rData & " " & Date _
                                                                            & " " & Time)

            Next i

            Exit Sub

        End If

        If UCase(Left(rData, 7)) = "/BANIP " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            Dim BanIP As String, XNick As Boolean

            rData = Right$(rData, Len(rData) - 7)
            tStr = Replace(ReadField(1, rData, Asc(" ")), "+", " ")

            TIndex = NameIndex(tStr)

            If TIndex <= 0 Then
                XNick = False
                BanIP = tStr
            Else
                XNick = True
                Call LogGM(UserList(UserIndex).Name, "/BANLAIP " & UserList(TIndex).Name & " - " & UserList(TIndex).ip)
                BanIP = UserList(TIndex).ip

            End If

            rData = Right$(rData, Len(rData) - Len(tStr))

            If BanIpBuscar(BanIP) > 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||La IP " & BanIP & " ya se encuentra en la lista de bans." & FONTTYPE_INFO)
                Exit Sub

            End If

            Call BanIpAgrega(BanIP)
            Call SendData(SendTarget.ToAdmins, UserIndex, 0, "||" & UserList(UserIndex).Name & " Baneo la IP " & BanIP & FONTTYPE_FIGHT)

            If XNick = True Then
                Call LogBan(TIndex, UserIndex, "Ban por IP desde Nick por " & rData)

                Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " echo a " & UserList(TIndex).Name & "." & FONTTYPE_FIGHT)
                Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " Banned a " & UserList(TIndex).Name & "." & FONTTYPE_FIGHT)

                UserList(TIndex).flags.Ban = 1

                Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(TIndex).Name)
                Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(TIndex).Name)
                Call CloseSocket(TIndex)

            End If

            Exit Sub

        End If

        If UCase(rData) = "/BANIPLIST" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
            tStr = "||"

            For LoopC = 1 To BanIps.Count
                tStr = tStr & BanIps.Item(LoopC) & ", "
            Next LoopC

            tStr = tStr & FONTTYPE_INFO
            Call SendData(SendTarget.toIndex, UserIndex, 0, tStr)
            Exit Sub

        End If

        If UCase$(Left$(rData, 15)) = "/CHAUTEMPLARIO " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 15)
            Call LogGM(UserList(UserIndex).Name, "ECHO DEL TEMPLARIO A: " & rData)

            TIndex = NameIndex(rData)
            Dim tArmIndex As Integer

            If TIndex > 0 Then
                UserList(TIndex).Faccion.Templario = 0
                UserList(TIndex).Faccion.Reenlistadas = 200
                UserList(TIndex).Faccion.RecibioArmaduraTemplaria = 0

                Call PerderItemsFaccionarios(UserIndex, UserList(TIndex).Faccion.ArmaduraFaccionaria)

                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas TEMPLARIAS y prohibida la reenlistada" & _
                                                                FONTTYPE_INFO)

                Call SendData(SendTarget.toIndex, TIndex, 0, "|| " & UserList(UserIndex).Name & _
                                                           " te ha expulsado en forma definitiva de las fuerzas TEMPLARIAS." & FONTTYPE_FIGHT)
            Else

                If FileExist(CharPath & rData & ".chr") Then
                    Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Templario", 0)
                    Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
                    Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(UserIndex).Name)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas TEMPLARIAS y prohibida la reenlistada" & _
                                                                    FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & ".chr inexistente." & FONTTYPE_INFO)

                End If

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 13)) = "/CHAUNEMESIS " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 13)
            Call LogGM(UserList(UserIndex).Name, "ECHO DEL NEMESIS A: " & rData)

            TIndex = NameIndex(rData)

            If TIndex > 0 Then
                UserList(TIndex).Faccion.Nemesis = 0
                UserList(TIndex).Faccion.Reenlistadas = 200
                UserList(TIndex).Faccion.RecibioArmaduraNemesis = 0

                Call PerderItemsFaccionarios(UserIndex, UserList(TIndex).Faccion.ArmaduraFaccionaria)

                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas NEMESIS y prohibida la reenlistada" & _
                                                                FONTTYPE_INFO)
                Call SendData(SendTarget.toIndex, TIndex, 0, "|| " & UserList(UserIndex).Name & _
                                                           " te ha expulsado en forma definitiva de las fuerzas NEMESIS." & FONTTYPE_FIGHT)
            Else

                If FileExist(CharPath & rData & ".chr") Then
                    Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Nemesis", 0)
                    Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
                    Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(UserIndex).Name)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas NEMESIS y prohibida la reenlistada" & _
                                                                    FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & ".chr inexistente." & FONTTYPE_INFO)

                End If

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 10)) = "/CHAUCAOS " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 10)
            Call LogGM(UserList(UserIndex).Name, "ECHO DEL CAOS A: " & rData)

            TIndex = NameIndex(rData)

            If TIndex > 0 Then
                UserList(TIndex).Faccion.FuerzasCaos = 0
                UserList(TIndex).Faccion.Reenlistadas = 200
                UserList(TIndex).Faccion.RecibioArmaduraCaos = 0

                Call PerderItemsFaccionarios(UserIndex, UserList(TIndex).Faccion.ArmaduraFaccionaria)

                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas del caos y prohibida la reenlistada" & _
                                                                FONTTYPE_INFO)
                Call SendData(SendTarget.toIndex, TIndex, 0, "|| " & UserList(UserIndex).Name & _
                                                           " te ha expulsado en forma definitiva de las fuerzas del caos." & FONTTYPE_FIGHT)
            Else

                If FileExist(CharPath & rData & ".chr") Then
                    Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "EjercitoCaos", 0)
                    Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
                    Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(UserIndex).Name)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas del caos y prohibida la reenlistada" & _
                                                                    FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & ".chr inexistente." & FONTTYPE_INFO)

                End If

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 10)) = "/CHAUREAL " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 10)
            Call LogGM(UserList(UserIndex).Name, "ECHO DE LA REAL A: " & rData)

            rData = Replace(rData, "\", "")
            rData = Replace(rData, "/", "")

            TIndex = NameIndex(rData)

            If TIndex > 0 Then
                UserList(TIndex).Faccion.ArmadaReal = 0
                UserList(TIndex).Faccion.Reenlistadas = 200
                UserList(TIndex).Faccion.RecibioArmaduraReal = 0

                Call PerderItemsFaccionarios(UserIndex, UserList(TIndex).Faccion.ArmaduraFaccionaria)

                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas reales y prohibida la reenlistada" & _
                                                                FONTTYPE_INFO)
                Call SendData(SendTarget.toIndex, TIndex, 0, "|| " & UserList(UserIndex).Name & _
                                                           " te ha expulsado en forma definitiva de las fuerzas reales." & FONTTYPE_FIGHT)
            Else

                If FileExist(CharPath & rData & ".chr") Then
                    Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "EjercitoReal", 0)
                    Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
                    Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(UserIndex).Name)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas reales y prohibida la reenlistada" & _
                                                                    FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & ".chr inexistente." & FONTTYPE_INFO)

                End If

            End If

            Exit Sub

        End If

        If UCase(Left(rData, 8)) = "/LASTIP " Then
            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right(rData, Len(rData) - 8)

            'No se si sea MUY necesario, pero por si las dudas... ;)
            rData = Replace(rData, "\", "")
            rData = Replace(rData, "/", "")

            If FileExist(CharPath & rData & ".chr", vbNormal) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||La ultima IP de """ & rData & """ fue : " & GetVar(CharPath & rData & ".chr", _
                                                                                                                      "INIT", "LastIP") & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Charfile """ & rData & """ inexistente." & FONTTYPE_INFO)

            End If

            Exit Sub

        End If

    Case 14    '<<<<<<<<<<<<<Comandos 14>>>>>>>>>>>>>>
        If UCase$(rData) = "/SEGUIR" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.TargetNpc > 0 Then
                Call DoFollow(UserList(UserIndex).flags.TargetNpc, UserList(UserIndex).Name)

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 3)) = "/CC" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            Call EnviarSpawnList(UserIndex)
            Exit Sub

        End If

        If UCase$(rData) = "/MATA" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 5)

            If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
            Call QuitarNPC(UserList(UserIndex).flags.TargetNpc)
            Exit Sub
        End If

    Case 15    '<<<<<<<<<<<<<Comandos 15>>>>>>>>>>>>>>
        If UCase$(rData) = "/LLUVIA" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            Call SecondaryAmbient
            Exit Sub

        End If

        If UCase$(Left$(rData, 7)) = "/CONTAR" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = "3"

            'If rData <= 0 Or rData >= 61 Then Exit Sub
            If CuentaRegresiva > 0 Then Exit Sub
            Call SendData(SendTarget.toall, 0, 0, "||Empieza en " & rData & "..." & FONTTYPE_GUILD)
            CuentaRegresiva = rData
            Exit Sub

        End If

        If UCase$(Left$(rData, 8)) = "/AOMANIA" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 8)

            Call GuerraBanda.Ban_Comienza("32")

        End If

        If UCase$(rData) = "/NAVE" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

            If UserList(UserIndex).flags.Navegando = 1 Then
                UserList(UserIndex).flags.Navegando = 0
            Else
                UserList(UserIndex).flags.Navegando = 1

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 11)) = "/GUARDAMAPA" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

            Call GrabarMapa(UserList(UserIndex).pos.Map, App.Path & "\WorldBackUp\Mapa" & UserList(UserIndex).pos.Map)
            Exit Sub

        End If

        If UCase$(Left$(rData, 20)) = "/TORNEOSAUTOMATICOS " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 20)

            If (Torneo_Activo And Torneo_Esperando) Then
                Call SendData(toIndex, UserIndex, 0, "||Ya hay un torneo automatico en curso, si quieres cancelarla, usa /CANCELATORNEO" & FONTTYPE_INFO)
                Exit Sub
            End If

            If rData > "6" Then
                Call SendData(toIndex, UserIndex, 0, "||Comando: /TORNEOSAUTOMATICOS <1-6>" & FONTTYPE_INFO)
            Else
                xao = 20
                RondaTorneo = rData
                Call SendData(SendTarget.toall, 0, 0, "||Esta empezando un nuevo torneo 1v1 de " & val(2 ^ RondaTorneo) & " participantes!! para participar pon /PARTICIPAR - (No cae inventario)" & FONTTYPE_GUILD)
                Call torneos_auto(RondaTorneo)
                Exit Sub
            End If

            Exit Sub
        End If

        If UCase$(Left$(rData, 14)) = "/CANCELATORNEO" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 14)

            If (Not Torneo_Activo And Not Torneo_Esperando) Then
                Call SendData(toIndex, UserIndex, 0, "||No hay un torneo automatico en curso!!" & FONTTYPE_INFO)
                Exit Sub
            End If

            Call Rondas_Cancela

            Exit Sub
        End If

    Case 16    '<<<<<<<<<<<<<Comandos 16>>>>>>>>>>>>>>
        If UCase$(rData) = "/INVISIBLE" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            Call DoAdminInvisible(UserIndex)
            Exit Sub

        End If



    Case 17    '<<<<<<<<<<<<<Comandos 17>>>>>>>>>>>>>>
        If UCase$(rData) = "/MASSKILL" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            For Y = UserList(UserIndex).pos.Y - MinYBorder + 1 To UserList(UserIndex).pos.Y + MinYBorder - 1
                For X = UserList(UserIndex).pos.X - MinXBorder + 1 To UserList(UserIndex).pos.X + MinXBorder - 1
                    If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
                       If MapData(UserList(UserIndex).pos.Map, X, Y).NpcIndex Then Call QuitarNPC(MapData(UserList(UserIndex).pos.Map, X, Y).NpcIndex)
                Next
            Next
            Exit Sub
        End If

        If UCase$(Left$(rData, 5)) = "/ACC " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 5)
            Call SpawnNpc(val(rData), UserList(UserIndex).pos, True, False)
            Call LogGM(UserList(UserIndex).Name, " Sumoneo un " & Npclist(IndexNPC).Name & " en mapa " & UserList(UserIndex).pos.Map)
            Exit Sub

        End If

        'Crear criatura con respawn, toma directamente el indice
        If UCase$(Left$(rData, 6)) = "/RACC " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 6)
            Call SpawnNpc(val(rData), UserList(UserIndex).pos, True, True)
            Call LogGM(UserList(UserIndex).Name, " Sumoneo un " & Npclist(IndexNPC).Name & " en mapa " & UserList(UserIndex).pos.Map)
            Exit Sub

        End If

        If UCase$(rData) = "/RESETINV" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
            rData = Right$(rData, Len(rData) - 9)

            If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
            Call ResetNpcInv(UserList(UserIndex).flags.TargetNpc)
            Exit Sub

        End If

        Select Case UCase$(Left$(rData, 8))

        Case "/TALKAS "
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.Privilegios > PlayerType.SemiDios Or UserList(UserIndex).flags.EsRolesMaster Then

                If UserList(UserIndex).flags.TargetNpc > 0 Then
                    tStr = Right$(rData, Len(rData) - 8)

                    Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNpc, Npclist(UserList(UserIndex).flags.TargetNpc).pos.Map, _
                                  "||" & vbWhite & "°" & tStr & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, _
                                  "||Debes seleccionar el NPC por el que quieres hablar antes de usar este comando" & FONTTYPE_INFO)

                End If

            End If

            Exit Sub

        End Select

        If UCase$(rData) = "/MASSEJECUTAR" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.UserLogged Then
                        If Not UserList(LoopC).flags.Privilegios >= 1 Then
                            If UserList(LoopC).pos.Map = UserList(UserIndex).pos.Map Then
                                Call UserDie(LoopC)
                            End If

                        End If

                    End If

                End If

            Next LoopC

            Exit Sub

        End If

        If UCase$(rData) = "/RELOADNPCS" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

            Call CargaNpcsDat

            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Npcs.dat y npcsHostiles.dat recargados." & FONTTYPE_INFO)
            Exit Sub

        End If

        If UCase$(Left$(rData, 6)) = "/RMATA" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 6)

            'Los consejeros no pueden RMATAr a nada en el mapa pretoriano
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero And UserList(UserIndex).pos.Map = MAPA_PRETORIANO Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los consejeros no pueden usar este comando en el mapa pretoriano." & FONTTYPE_INFO)
                Exit Sub

            End If

            TIndex = UserList(UserIndex).flags.TargetNpc

            If TIndex > 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||RMatas (con posible respawn) a: " & Npclist(TIndex).Name & FONTTYPE_INFO)
                Dim MiNPC As npc
                MiNPC = Npclist(TIndex)
                Call QuitarNPC(TIndex)
                Call RespawnNPC(MiNPC)

                'SERES
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Debes hacer click sobre el NPC antes" & FONTTYPE_INFO)

            End If

            Exit Sub

        End If

        If UCase(Left(rData, 10)) = "/INVASION " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 10)

            n = val(ReadField(1, rData, 32))
            Mapa = val(ReadField(2, rData, 32))
            tInt = val(ReadField(3, rData, 32))

            If n = "0" Or Mapa = "0" Or tInt = "0" Then
                Call SendData(toIndex, UserIndex, 0, "||Debes utilizar /INVASION NUMERO NPC MAPA CANTIDAD." & FONTTYPE_GUILD)
                Exit Sub
            End If

            For LoopC = 1 To tInt
                MiPos.Map = Mapa
                MiPos.X = RandomNumber(20, 80)
                MiPos.Y = RandomNumber(20, 80)
                Call SpawnNpc(n, MiPos, True, 0)
            Next LoopC

            Call SendData(toall, 0, 0, "||Una invasión de " & Npclist(IndexNPC).Name & " ha caido sobre el mapa " & Mapa & " con una cantidad de " & tInt & " npcs." & FONTTYPE_GUILD)

            Exit Sub
        End If

    Case 18    '<<<<<<<<<<<<<Comandos 18>>>>>>>>>>>>>>
        If UCase(Left(rData, 3)) = "/CI" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            Dim txt As String
            Dim Cadena() As String
            Dim IdItem As String
            Dim Cantidad As String

            txt = rData

            Cadena = Split(txt, Chr$(32))

            If txt = "/CI" Or UBound(Cadena) < 2 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis incorrecto." & FONTTYPE_GUILD)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis: /CI <ID ITEM> <CANTIDAD>" & FONTTYPE_GUILD)
                Exit Sub

            End If

            If Not IsNumeric(Cadena(1)) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El ID Item debe ser númerica." & FONTTYPE_GUILD)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis: /CI <ID ITEM> <CANTIDAD> Ejemplo: /CI 2 <CANTIDAD>." & FONTTYPE_GUILD)
                Exit Sub

            End If

            If Not IsNumeric(Cadena(2)) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El Cantidad debe ser numérica." & FONTTYPE_GUILD)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis: /CI <ID ITEM> <CANTIDAD> Ejemplo: /CI 2 10." & FONTTYPE_GUILD)
                Exit Sub

            End If

            If Cadena(2) > 1200 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Has superado el tope de cantidad. (Max: 1200)" & FONTTYPE_GUILD)
                Exit Sub

            End If

            IdItem = Cadena(1)
            Cantidad = Cadena(2)

            If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).OBJInfo.ObjIndex > 0 Then
                Exit Sub

            End If

            If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.Map > 0 Then
                Exit Sub

            End If

            If val(IdItem) < 1 Or val(IdItem) > NumObjDatas Then
                Exit Sub

            End If

            'Is the object not null?
            If ObjData(val(IdItem)).Name = "" Then Exit Sub

            Dim Objeto As Obj

            Objeto.Amount = val(Cantidad)
            Objeto.ObjIndex = val(IdItem)

            Call MeterItemEnInventario(UserIndex, Objeto)

            Call LogGM(UserList(UserIndex).Name, "Creo: " & Cantidad & " " & ObjData(Objeto.ObjIndex).Name)

            Exit Sub

        End If

        If UCase$(Left$(rData, 5)) = "/DEST" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 5)
            Call EraseObj(SendTarget.ToMap, UserIndex, UserList(UserIndex).pos.Map, 10000, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, _
                          UserList(UserIndex).pos.Y)
            Exit Sub
        End If

        If UCase$(Left$(rData, 9)) = "/MASSDEST" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            With UserList(UserIndex)

                For Y = UserList(UserIndex).pos.Y - MinYBorder + 1 To UserList(UserIndex).pos.Y + MinYBorder - 1
                    For X = UserList(UserIndex).pos.X - MinXBorder + 1 To UserList(UserIndex).pos.X + MinXBorder - 1
                        If InMapBounds(.pos.Map, X, Y) Then
                            If ObjetosBorrable(MapData(.pos.Map, X, Y).OBJInfo.ObjIndex) Then
                                Call EraseObj(SendTarget.ToMap, 0, .pos.Map, 10000, .pos.Map, X, Y)
                            End If
                        End If
                    Next
                Next

            End With

            Exit Sub
        End If

        If UCase$(Left$(rData, 8)) = "/MASSORO" Then

            With UserList(UserIndex)

                Call LogGM(.Name, "Comando: " & rData)

                For Y = .pos.Y - MinYBorder + 1 To .pos.Y + MinYBorder - 1
                    For X = .pos.X - MinXBorder + 1 To .pos.X + MinXBorder - 1

                        If InMapBounds(.pos.Map, X, Y) Then
                            If MapData(.pos.Map, X, Y).OBJInfo.ObjIndex = iORO Then
                                Call EraseObj(SendTarget.ToMap, 0, .pos.Map, 10000, .pos.Map, X, Y)

                            End If

                        End If

                    Next X
                Next Y

            End With

            Exit Sub

        End If

        If UCase$(rData) = "/LIMPIAROBJS" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            Call LimpiarObjs
        End If

        If UCase$(Left$(rData, 7)) = "/BLOKK " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 7)
            TIndex = NameIndex(rData)

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
                Exit Sub

            End If

            UserList(TIndex).flags.Ban = 1
            Call Ban(UserList(TIndex).Name, UserList(UserIndex).Name, "Bloqueo de Cliente")
            Call WriteVar(CharPath & UCase(UserList(TIndex).Name) & ".chr", "FLAGS", "Ban", "1")
            'ponemos la pena
            tInt = val(GetVar(CharPath & UCase(UserList(TIndex).Name) & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & UCase(UserList(TIndex).Name) & ".chr", "PENAS", "Cant", tInt + 1)
            Call WriteVar(CharPath & UCase(UserList(TIndex).Name) & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & " BAN" & " " & _
                                                                                                     Date & " " & Time)

            Call SendData(SendTarget.toIndex, TIndex, 0, "ABBLOCK")
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Cliente BLOQUEADO =)" & FONTTYPE_INFO)
            Exit Sub

        End If

        If UCase$(Left$(rData, 11)) = "/KICKCONSE " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
            rData = Right$(rData, Len(rData) - 11)
            TIndex = NameIndex(rData)

            If TIndex <= 0 Then
                If FileExist(CharPath & rData & ".chr") Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline, Echando de los consejos" & FONTTYPE_INFO)
                    Call WriteVar(CharPath & UCase(rData) & ".chr", "CONSEJO", "PERTENECE", 0)
                    Call WriteVar(CharPath & UCase(rData) & ".chr", "CONSEJO", "PERTENECECAOS", 0)
                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No se encuentra el charfile " & CharPath & rData & ".chr" & FONTTYPE_INFO)
                    Exit Sub

                End If

            Else

                If UserList(TIndex).flags.PertAlCons > 0 Then
                    Call SendData(SendTarget.toIndex, TIndex, 0, "||Has sido echado en el consejo de banderbill" & FONTTYPE_TALK & ENDC)
                    UserList(TIndex).flags.PertAlCons = 0
                    Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y)
                    Call SendData(SendTarget.toall, 0, 0, "||" & rData & " fue expulsado del consejo de Banderbill" & FONTTYPE_CONSEJO)

                End If

                If UserList(TIndex).flags.PertAlConsCaos > 0 Then
                    Call SendData(SendTarget.toIndex, TIndex, 0, "||Has sido echado en el consejo de la legión oscura" & FONTTYPE_TALK & ENDC)
                    UserList(TIndex).flags.PertAlConsCaos = 0
                    Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y)
                    Call SendData(SendTarget.toall, 0, 0, "||" & rData & " fue expulsado del consejo de la Legión Oscura" & FONTTYPE_CONSEJOCAOS)

                End If

            End If

            Exit Sub

        End If

        If UCase$(Left$(rData, 11)) = "/FORCEMIDI " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 11)

            If Not IsNumeric(rData) Then
                Exit Sub
            Else
                Call SendData(SendTarget.toall, 0, 0, "|| " & UserList(UserIndex).Name & " broadcast musica: " & rData & FONTTYPE_SERVER)
                Call SendData(SendTarget.toall, 0, 0, "TM" & rData)

            End If

        End If

        If UCase$(Left$(rData, 10)) = "/FORCEWAV " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 10)

            If Not IsNumeric(rData) Then
                Exit Sub
            Else
                Call SendData(SendTarget.toall, 0, 0, "TW" & rData)

            End If

        End If

        If UCase$(rData) = "/SOLOGM" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

            If ServerSoloGMs > 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Servidor Válido para todos" & FONTTYPE_INFO)
                ServerSoloGMs = 0
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Servidor Válido solo a administradores." & FONTTYPE_INFO)
                ServerSoloGMs = 1

            End If

            Exit Sub

        End If

        If UCase$(rData) = "/PISO" Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            With UserList(UserIndex)

                For Y = 0 To 100
                    For X = 0 To 100

                        If InMapBounds(.pos.Map, X, Y) Then
                            If MapData(.pos.Map, X, Y).OBJInfo.ObjIndex > 0 Then

                                Call SendData(toIndex, UserIndex, 0, "||(" & X & ", " & Y & ") " & ObjData(MapData(.pos.Map, X, Y).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)

                            End If

                        End If

                    Next X
                Next Y

            End With

            Exit Sub
        End If

        If UCase$(Left$(rData, 5)) = "/MAP " Then
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

            rData = Right(rData, Len(rData) - 5)

            Select Case UCase(ReadField(1, rData, 32))

            Case "PK"
                tStr = ReadField(2, rData, 32)

                If tStr <> "" Then
                    MapInfo(UserList(UserIndex).pos.Map).Pk = IIf(tStr = "0", True, False)
                    Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).pos.Map & ".dat", "Mapa" & UserList(UserIndex).pos.Map, "Pk", _
                                  tStr)

                End If

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Mapa " & UserList(UserIndex).pos.Map & " PK: " & MapInfo(UserList( _
                                                                                                                            UserIndex).pos.Map).Pk & FONTTYPE_INFO)

            Case "BACKUP"
                tStr = ReadField(2, rData, 32)

                If tStr <> "" Then
                    MapInfo(UserList(UserIndex).pos.Map).BackUp = CByte(tStr)
                    Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).pos.Map & ".dat", "Mapa" & UserList(UserIndex).pos.Map, _
                                  "backup", tStr)

                End If

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Mapa " & UserList(UserIndex).pos.Map & " Backup: " & MapInfo(UserList( _
                                                                                                                                UserIndex).pos.Map).BackUp & FONTTYPE_INFO)

            End Select

            Exit Sub

        End If

    End Select    ' final select commandos

End Sub

Public Sub AllCommands(ByVal UserIndex As Integer, ByVal rData As String)

    Dim LoopC As Integer
    Dim tStr As String
    Dim tInt As Integer
    Dim tLong As Long
    Dim TIndex As Integer
    Dim tName As String
    Dim tMessage As String
    Dim Arg1 As String
    Dim Arg2 As String
    Dim Arg3 As String
    Dim Arg4 As String
    Dim Arg5 As String
    Dim Mapa As Integer
    Dim Name As String
    Dim n As Integer
    Dim mifile As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim DummyInt As Integer
    Dim i As Integer
    Dim hdStr As String
    Dim tPath As String
    Dim MiPos As WorldPos


    'Mensaje del servidor
    If UCase$(Left$(rData, 6)) = "/RMSG " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 6)

        If rData <> "" Then
            Call SendData(toall, 0, 0, "||<" & UserList(UserIndex).Name & "> " & rData & FONTTYPE_TALK)
        End If

        Exit Sub
    End If

    If UCase$(Left$(rData, 3)) = "SCS" Then
        If UserList(UserIndex).SnapShot = False Then Exit Sub
        Dim adminUI As Integer
        adminUI = UserList(UserIndex).SnapShotAdmin
        If adminUI <= 0 Then Exit Sub
        'If UCase$(UserList(adminUI).Name) <> "SETH" Then Exit Sub


        rData = Right$(rData, Len(rData) - 3)
        Call SendData(SendTarget.toIndex, adminUI, 0, "SCSR" & rData)
        Exit Sub
    End If

    If UCase$(Left$(rData, 9)) = "/IRCERCA " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Dim indiceUserDestino As Integer
        rData = Right$(rData, Len(rData) - 9)    'obtiene el nombre del usuario
        TIndex = NameIndex(rData)

        'Si es dios o Admins no podemos salvo que nosotros también lo seamos
        If (EsDios(rData) Or EsAdmin(rData)) And UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then Exit Sub

        If TIndex <= 0 Then    'existe el usuario destino?
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub

        End If

        For tInt = 2 To 5    'esto for sirve ir cambiando la distancia destino
            For i = UserList(TIndex).pos.X - tInt To UserList(TIndex).pos.X + tInt
                For DummyInt = UserList(TIndex).pos.Y - tInt To UserList(TIndex).pos.Y + tInt

                    If (i >= UserList(TIndex).pos.X - tInt And i <= UserList(TIndex).pos.X + tInt) And (DummyInt = UserList(TIndex).pos.Y - tInt Or _
                                                                                                        DummyInt = UserList(TIndex).pos.Y + tInt) Then

                        If MapData(UserList(TIndex).pos.Map, i, DummyInt).UserIndex = 0 And LegalPos(UserList(TIndex).pos.Map, i, DummyInt) Then
                            Call WarpUserChar(UserIndex, UserList(TIndex).pos.Map, i, DummyInt, True)
                            Exit Sub

                        End If

                    ElseIf (DummyInt >= UserList(TIndex).pos.Y - tInt And DummyInt <= UserList(TIndex).pos.Y + tInt) And (i = UserList( _
                                                                                                                          TIndex).pos.X - tInt Or i = UserList(TIndex).pos.X + tInt) Then

                        If MapData(UserList(TIndex).pos.Map, i, DummyInt).UserIndex = 0 And LegalPos(UserList(TIndex).pos.Map, i, DummyInt) Then
                            Call WarpUserChar(UserIndex, UserList(TIndex).pos.Map, i, DummyInt, True)
                            Exit Sub

                        End If

                    End If

                Next DummyInt
            Next i
        Next tInt

        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Todos los lugares estan ocupados." & FONTTYPE_INFO)
        Exit Sub

    End If

    If UCase$(Left$(rData, 4)) = "/REM" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 5)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Comentario salvado..." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UCase$(Left$(rData, 5)) = "/HORA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 5)
        Call SendData(SendTarget.toall, 0, 0, "||Hora: " & Time & " " & Date & FONTTYPE_INFO)
        Exit Sub
    End If

    If UCase$(rData) = "/LIMPIAROBJS" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Call LimpiarObjs
    End If

    If UCase$(Left$(rData, 6)) = "/NENE " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 6)

        If MapaValido(val(rData)) Then
            Dim NpcIndex As Integer
            Dim ContS As String

            ContS = ""

            For NpcIndex = 1 To LastNPC

                '¿esta vivo?
                If Npclist(NpcIndex).flags.NPCActive And Npclist(NpcIndex).pos.Map = val(rData) And Npclist(NpcIndex).Hostile = 1 And Npclist( _
                   NpcIndex).Stats.Alineacion = 2 Then
                    ContS = ContS & Npclist(NpcIndex).Name & ", "

                End If

            Next NpcIndex

            If ContS <> "" Then
                ContS = Left(ContS, Len(ContS) - 2)
            Else
                ContS = "No hay NPCS"

            End If

            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Npcs en mapa: " & ContS & FONTTYPE_INFO)

        End If

        Exit Sub

    End If

    If UCase$(rData) = "/TELEPLOC" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Call WarpUserChar(UserIndex, UserList(UserIndex).flags.TargetMap, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, True)
        Call LogGM(UserList(UserIndex).Name, "/TELEPLOC " & UserList(UserIndex).Name & " x:" & UserList(UserIndex).flags.TargetX & " y:" & UserList( _
                                             UserIndex).flags.TargetY & " Map:" & UserList(UserIndex).flags.TargetMap)
        Exit Sub
    End If

    If UCase$(Left$(rData, 6)) = "/MAIL " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 6)

        TIndex = NameIndex(rData)

        If FileExist(CharPath & rData & ".chr", vbNormal) = False Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
            Exit Sub
        End If

        If TIndex > 0 Then
            Call SendData(SendTarget.ToAdmins, 0, 0, "||Su email es: " & UserList(TIndex).Email & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario Offline." & FONTTYPE_INFO)
        End If

        Exit Sub
    End If

    If UCase$(Left$(rData, 7)) = "/MOVER " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 7)

        TIndex = NameIndex(rData)

        If FileExist(CharPath & rData & ".chr", vbNormal) = False Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)

            Exit Sub

        End If

        If TIndex <= 0 Then
            Call WriteVar(App.Path & "\Charfile\" & rData & ".chr", "INIT", "Position", "34-40-50")
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El pj ha sido transportado a nix." & FONTTYPE_INFO)
            Exit Sub
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario está conectado, no ha sido teletransportado." & FONTTYPE_INFO)

        End If

        Exit Sub

    End If

    'Teleportar
    If UCase$(Left$(rData, 7)) = "/TELEP " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 7)
        Mapa = val(ReadField(2, rData, 32))

        If Not MapaValido(Mapa) Then Exit Sub
        Name = ReadField(1, rData, 32)

        If Name = "" Then Exit Sub

        'Nuevo code
        If Name = "PaneldeGM" Then
            TIndex = UserIndex

            X = val(ReadField(3, rData, 32))
            Y = val(ReadField(4, rData, 32))

            If Not InMapBounds(Mapa, X, Y) Then Exit Sub
            Call WarpUserChar(TIndex, Mapa, X, Y, True)
            Call SendData(SendTarget.toIndex, TIndex, 0, "||Has sido teletransportado." & FONTTYPE_GUILD)
            Exit Sub

        End If

        'Fin de mi nuevo code

        If UCase$(Name) <> "YO" Then
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                Exit Sub

            End If

            TIndex = NameIndex(Name)
        Else
            TIndex = UserIndex

        End If

        X = val(ReadField(3, rData, 32))
        Y = val(ReadField(4, rData, 32))

        If Not InMapBounds(Mapa, X, Y) Then Exit Sub
        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub

        End If

        Call WarpUserChar(TIndex, Mapa, X, Y, True)
        Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " transportado." & FONTTYPE_INFO)
        Call LogGM(UserList(UserIndex).Name, "Transporto a " & UserList(TIndex).Name & " hacia " & "Mapa" & Mapa & " X:" & X & " Y:" & Y)

        If UCase$(Name) <> "YO" Then
            Call LogGM(UserList(UserIndex).Name, "Transporto a " & UserList(TIndex).Name & " hacia " & "Mapa" & Mapa & " X:" & X & " Y:" & Y)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 11)) = "/SILENCIAR " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 11)

        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(TIndex).flags.Silenciado = 0 Then
            UserList(TIndex).flags.Silenciado = 1
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El Usuario Ha sido silenciado." & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, TIndex, 0, "||Has Sido Silenciado" & FONTTYPE_INFO)
        Else
            UserList(TIndex).flags.Silenciado = 0
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El Usuario Ha sido DesSilenciado." & FONTTYPE_INFO)
            Call LogGM(UserList(UserIndex).Name, "/DESsilenciar " & UserList(TIndex).Name)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 5)) = "/SUM " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 5)

        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El jugador no esta online." & FONTTYPE_INFO)
            Exit Sub

        End If

        Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " há sido trasportado." & FONTTYPE_INFO)
        Call WarpUserChar(TIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y + 1, True)

        Call LogGM(UserList(UserIndex).Name, "/SUM " & UserList(TIndex).Name & " Map:" & UserList(UserIndex).pos.Map & " X:" & UserList( _
                                             UserIndex).pos.X & " Y:" & UserList(UserIndex).pos.Y)
        Exit Sub

    End If

    If UCase$(Left$(rData, 9)) = "/RESPUES " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 9)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline!!." & FONTTYPE_INFO)
            Exit Sub

        End If

        Call MostrarSop(UserIndex, TIndex, rData)
        SendData SendTarget.toIndex, UserIndex, 0, "INITSOP"
        Exit Sub
    End If

    If UCase$(Left$(rData, 10)) = "/EJECUTAR " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        rData = Right$(rData, Len(rData) - 10)
        TIndex = NameIndex(rData)

        If UserList(TIndex).flags.Privilegios > PlayerType.User Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, _
                          "||Osea, yo te dejaria pero es un viaje, mira si se caen altos items anda a saber, mejor qedate ahi y no intentes ejecutar mas gms la re puta qe te pario." _
                        & FONTTYPE_EJECUCION)
            Exit Sub

        End If

        If TIndex > 0 Then

            Call UserDie(TIndex)

            If UserList(TIndex).pos.Map = 1 Then
                Call TirarTodo(TIndex)

            End If

            Call SendData(SendTarget.toall, 0, 0, "||El GameMaster " & UserList(UserIndex).Name & " ha ejecutado a " & UserList(TIndex).Name & _
                                                  FONTTYPE_EJECUCION)
            Call LogGM(UserList(UserIndex).Name, " ejecuto a " & UserList(TIndex).Name)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No está online" & FONTTYPE_EJECUCION)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 11)) = "/DROPQUEST " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 11)

        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            If Quest.Existe(rData) Then Call Quest.Quitar(rData)
            Exit Sub

        End If

        If UserList(TIndex).flags.Quest = 1 Then
            If Quest.Existe(UserList(TIndex).Name) Then Call Quest.Quitar(UserList(TIndex).Name)
            UserList(TIndex).flags.Quest = 0
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has borrado el QUEST de: " & rData & FONTTYPE_INFO)
            Exit Sub
        Else

            If Quest.Existe(UserList(TIndex).Name) Then Call Quest.Quitar(UserList(TIndex).Name)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has borrado el QUEST de: " & rData & FONTTYPE_INFO)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 9)) = "/DROPSOS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 9)

        Call DropSOS(rData, UserIndex)

        Exit Sub

    End If

    If UCase$(Left$(rData, 8)) = "/DROPGM " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 8)

        Call BorrarGM(rData, UserIndex)

        Exit Sub

    End If

    If UCase$(Left$(rData, 14)) = "/PANELCONSULTA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Call CargarArchivosGM(UserIndex)

        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "/CONTAR" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = "3"

        'If rData <= 0 Or rData >= 61 Then Exit Sub
        If CuentaRegresiva > 0 Then Exit Sub
        Call SendData(SendTarget.toall, 0, 0, "||Empieza en " & rData & "..." & FONTTYPE_GUILD)
        CuentaRegresiva = rData
        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "SOSDONE" Then
        rData = Right$(rData, Len(rData) - 7)
        Call Ayuda.Quitar(rData)
        Exit Sub

    End If

    'IR A
    If UCase$(Left$(rData, 10)) = "/ENCUESTA " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.Privilegios <> PlayerType.Dios Then Exit Sub
        If Encuesta.ACT = 1 Then Call SendData(SendTarget.toIndex, UserIndex, 0, "||Hay una encuesta en curso!." & FONTTYPE_INFO)
        rData = Right$(rData, Len(rData) - 10)

        Encuesta.EncNO = 0
        Encuesta.EncSI = 0
        Encuesta.Tiempo = 0
        Encuesta.ACT = 1

        Call SendData(SendTarget.toall, 0, 0, "||Encuesta: " & rData & FONTTYPE_GUILD)
        Call SendData(SendTarget.toall, 0, 0, "||Encuesta: Enviar /SI o /NO. Tiempo de encuesta: 1 Minuto." & FONTTYPE_TALK)
        Exit Sub

    End If

    'Quitar NPC
    If UCase$(rData) = "/MATA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 5)

        If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
        Call QuitarNPC(UserList(UserIndex).flags.TargetNpc)
        Exit Sub
    End If

    If UCase$(rData) = "/MASSKILL" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        For Y = UserList(UserIndex).pos.Y - MinYBorder + 1 To UserList(UserIndex).pos.Y + MinYBorder - 1
            For X = UserList(UserIndex).pos.X - MinXBorder + 1 To UserList(UserIndex).pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
                   If MapData(UserList(UserIndex).pos.Map, X, Y).NpcIndex Then Call QuitarNPC(MapData(UserList(UserIndex).pos.Map, X, Y).NpcIndex)
            Next
        Next
        Exit Sub
    End If

    'Destruir
    If UCase$(Left$(rData, 5)) = "/DEST" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 5)
        Call EraseObj(SendTarget.ToMap, UserIndex, UserList(UserIndex).pos.Map, 10000, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, _
                      UserList(UserIndex).pos.Y)
        Exit Sub
    End If

    If UCase$(Left$(rData, 9)) = "/MASSDEST" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        With UserList(UserIndex)

            For Y = UserList(UserIndex).pos.Y - MinYBorder + 1 To UserList(UserIndex).pos.Y + MinYBorder - 1
                For X = UserList(UserIndex).pos.X - MinXBorder + 1 To UserList(UserIndex).pos.X + MinXBorder - 1
                    If InMapBounds(.pos.Map, X, Y) Then
                        If ObjetosBorrable(MapData(.pos.Map, X, Y).OBJInfo.ObjIndex) Then
                            Call EraseObj(SendTarget.ToMap, 0, .pos.Map, 10000, .pos.Map, X, Y)
                        End If
                    End If
                Next
            Next

        End With

        Exit Sub
    End If

    If UCase$(Left$(rData, 7)) = "/BLOKK " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 7)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub

        End If

        UserList(TIndex).flags.Ban = 1
        Call Ban(UserList(TIndex).Name, UserList(UserIndex).Name, "Bloqueo de Cliente")
        Call WriteVar(CharPath & UCase(UserList(TIndex).Name) & ".chr", "FLAGS", "Ban", "1")
        'ponemos la pena
        tInt = val(GetVar(CharPath & UCase(UserList(TIndex).Name) & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & UCase(UserList(TIndex).Name) & ".chr", "PENAS", "Cant", tInt + 1)
        Call WriteVar(CharPath & UCase(UserList(TIndex).Name) & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & " BAN" & " " & _
                                                                                                 Date & " " & Time)

        Call SendData(SendTarget.toIndex, TIndex, 0, "ABBLOCK")
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Cliente BLOQUEADO =)" & FONTTYPE_INFO)
        Exit Sub

    End If

    If UCase$(Left$(rData, 5)) = "/IRA " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 5)

        TIndex = NameIndex(rData)

        'Si es dios o Admins no podemos salvo que nosotros también lo seamos
        'If (EsDios(rData) Or EsAdmin(rData)) And UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then _
         '    Exit Sub

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(TIndex).pos.Map = CastilloNorte Or UserList(TIndex).pos.Map = CastilloOeste Or UserList(TIndex).pos.Map = CastilloEste Or UserList(TIndex).pos.Map = CastilloSur Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Cuidado, está en castillo. Atiéndele más tarde." & FONTTYPE_INFO)
            Exit Sub
        ElseIf UserList(TIndex).pos.Map = MapaFortaleza Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Cuidado, está en la fortaleza. Atiéndele más tarde." & FONTTYPE_INFO)
            Exit Sub
        End If

        Call WarpUserChar(UserIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y + 1, True)

        If UserList(UserIndex).flags.AdminInvisible = 0 Then Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(UserIndex).Name & _
                                                                                                        " se ha trasportado hacia donde te encontras." & FONTTYPE_INFO)
        Call LogGM(UserList(UserIndex).Name, "/IRA " & UserList(TIndex).Name & " Mapa:" & UserList(TIndex).pos.Map & " X:" & UserList(TIndex).pos.X _
                                           & " Y:")
        Exit Sub

    End If

    'Haceme invisible vieja!
    If UCase$(rData) = "/INVISIBLE" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Call DoAdminInvisible(UserIndex)
        Exit Sub

    End If

    If UCase$(rData) = "/PANELGM" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call SendData(SendTarget.toIndex, UserIndex, 0, "ABPANEL")
        Exit Sub

    End If

    If UCase$(rData) = "LISTUSU" Then
        ' If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        tStr = "LISTUSU"

        For LoopC = 1 To LastUser

            If (UserList(LoopC).Name <> "") Then
                tStr = tStr & UserList(LoopC).Name & ","

            End If

        Next LoopC

        If Len(tStr) > 7 Then
            tStr = Left$(tStr, Len(tStr) - 1)

        End If

        Call SendData(SendTarget.toIndex, UserIndex, 0, tStr)
        Exit Sub

    End If

    If UCase$(rData) = "LISTQST" Then

        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Dim mm As String

        For n = 1 To Quest.Longitud
            mm = Quest.VerElemento(n)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "LISTQST" & mm)
        Next n

        Exit Sub

    End If

    '[Barrin 30-11-03]
    If UCase$(rData) = "/TRABAJANDO" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        For LoopC = 1 To LastUser

            If (UserList(LoopC).Name <> "") And UserList(LoopC).Counters.Trabajando > 0 Then
                tStr = tStr & UserList(LoopC).Name & ", "

            End If

        Next LoopC

        If tStr <> "" Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuarios trabajando: " & tStr & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No hay usuarios trabajando" & FONTTYPE_INFO)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 6)) = "/INFO " Then
        rData = Right$(rData, Len(rData) - 6)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        rData = NameIndex(rData)

        If rData = 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario Offline.." & FONTTYPE_INFO)
            Exit Sub
        End If

        Call EnviarAtribGM(UserIndex, rData)
        Call EnviarFamaGM(UserIndex, rData)
        Call EnviarMiniEstadisticasGM(UserIndex, rData)

        Call SendData(SendTarget.toIndex, UserIndex, 0, "INFSTAT")

    End If

    If UCase$(Left$(rData, 8)) = "/IPNICK " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 8)

        TIndex = NameIndex(rData)

        If TIndex > 0 Then
            Call SendData(toIndex, UserIndex, 0, "||El ip de " & UserList(TIndex).Name & " es: " & UserList(UserIndex).ip & FONTTYPE_INFO)
        End If

        Exit Sub
    End If

    If UCase$(Left$(rData, 7)) = "/DONDE " Then
        rData = Right$(rData, Len(rData) - 7)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        rData = NameIndex(rData)

        If rData = 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario Offline.." & FONTTYPE_INFO)
            Exit Sub
        End If

        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Ubicacion " & UserList(rData).Name & _
                                                        ": " & UserList(rData).pos.Map & ", " & UserList(rData).pos.X & ", " & UserList(rData).pos.Y & FONTTYPE_INFO)

    End If

    If UCase$(Left$(rData, 8)) = "/CARCEL " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        '/carcel nick@motivo@<tiempo>
        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        rData = Right$(rData, Len(rData) - 8)

        Name = ReadField(1, rData, Asc("@"))
        tStr = ReadField(2, rData, Asc("@"))

        If (Not IsNumeric(ReadField(3, rData, Asc("@")))) Or Name = "" Or tStr = "" Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Utilice /carcel nick@motivo@tiempo" & FONTTYPE_INFO)
            Exit Sub

        End If

        i = val(ReadField(3, rData, Asc("@")))

        TIndex = NameIndex(Name)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(TIndex).flags.Privilegios > PlayerType.User Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes encarcelar a administradores." & FONTTYPE_INFO)
            Exit Sub

        End If

        If i > 120 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes encarcelar por mas de 120 minutos." & FONTTYPE_INFO)
            Exit Sub

        End If

        Name = Replace(Name, "\", "")
        Name = Replace(Name, "/", "")

        If FileExist(CharPath & Name & ".chr", vbNormal) Then
            tInt = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & Name & ".chr", "PENAS", "Cant", tInt + 1)
            Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & " lo encarceló por el tiempo de " & _
                                                                           i & "  minutos, El motivo Fue: " & LCase$(tStr) & " " & Date & " " & Time)

        End If

        Call Encarcelar(TIndex, i, UserList(UserIndex).Name)
        Call LogGM(UserList(UserIndex).Name, " encarcelo a " & Name)
        Exit Sub

    End If

    If UCase$(Left$(rData, 6)) = "/RMATA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 6)

        'Los consejeros no pueden RMATAr a nada en el mapa pretoriano
        If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero And UserList(UserIndex).pos.Map = MAPA_PRETORIANO Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los consejeros no pueden usar este comando en el mapa pretoriano." & FONTTYPE_INFO)
            Exit Sub

        End If

        TIndex = UserList(UserIndex).flags.TargetNpc

        If TIndex > 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||RMatas (con posible respawn) a: " & Npclist(TIndex).Name & FONTTYPE_INFO)
            Dim MiNPC As npc
            MiNPC = Npclist(TIndex)
            Call QuitarNPC(TIndex)
            Call RespawnNPC(MiNPC)

            'SERES
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Debes hacer click sobre el NPC antes" & FONTTYPE_INFO)

        End If

        Exit Sub

    End If

    If UCase(Left(rData, 10)) = "/INVASION " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 10)

        n = val(ReadField(1, rData, 32))
        Mapa = val(ReadField(2, rData, 32))
        tInt = val(ReadField(3, rData, 32))

        If n = "0" Or Mapa = "0" Or tInt = "0" Then
            Call SendData(toIndex, UserIndex, 0, "||Debes utilizar /INVASION NUMERO NPC MAPA CANTIDAD." & FONTTYPE_GUILD)
            Exit Sub
        End If

        For LoopC = 1 To tInt
            MiPos.Map = Mapa
            MiPos.X = RandomNumber(20, 80)
            MiPos.Y = RandomNumber(20, 80)
            Call SpawnNpc(n, MiPos, True, 0)
        Next LoopC

        Call SendData(toall, 0, 0, "||Una invasión de " & Npclist(IndexNPC).Name & " ha caido sobre el mapa " & Mapa & " con una cantidad de " & tInt & " npcs." & FONTTYPE_GUILD)

        Exit Sub
    End If

    If UCase$(Left$(rData, 9)) = "/LIBERAR " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 9)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El jugador no esta online." & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(TIndex).Counters.Pena > 0 Then
            UserList(TIndex).Counters.Pena = 0
            Call SendData(SendTarget.toIndex, TIndex, 0, "||El gm te ha liberado." & FONTTYPE_Motd5)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario liberado." & FONTTYPE_INFO)
            Call WarpUserChar(TIndex, 48, 75, 65, False)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no esta en la carcel." & FONTTYPE_INFO)
            Exit Sub

        End If

    End If

    If UCase$(Left$(rData, 7)) = "/AVISO " Then

        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        rData = Right$(rData, Len(rData) - 7)

        TIndex = NameIndex(rData)

        If TIndex > 0 Then

            If UserList(TIndex).flags.Privilegios > PlayerType.User Then
                Call CloseSocket(UserIndex)
            Else
                Call SendData(SendTarget.toall, 0, 0, "||" & UserList(UserIndex).Name & " se va a sacrificar a " & UserList(TIndex).Name & " en NIX." & FONTTYPE_GUILD)
                Call WarpUserChar(UserIndex, "34", "73", "58", True)
                Call WarpUserChar(TIndex, "34", "74", "59", True)
            End If

        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario Offline" & FONTTYPE_INFO)
        End If

        Exit Sub
    End If

    If UCase$(Left$(rData, 11)) = "/SACRIFICA " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 11)

        TIndex = NameIndex(rData)

        If TIndex > 0 Then

            If UserList(TIndex).flags.Privilegios > PlayerType.User Then
                Call CloseSocket(UserIndex)
            Else
                Dim ObjSC As Obj
                ObjSC.Amount = "1"
                ObjSC.ObjIndex = "1043"

                Call MeterItemEnInventario(TIndex, ObjSC)


                Call SendData(SendTarget.toall, 0, 0, "||" & UserList(UserIndex).Name & " sacrificó por macrear a " & UserList(TIndex).Name & FONTTYPE_SERVER)
                Call TirarTodo(TIndex)
                UserList(TIndex).flags.Ban = 1
                Call CloseSocket(TIndex)



            End If

        Else

            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario Offline." & FONTTYPE_INFO)

        End If


        Exit Sub
    End If

    If UCase$(Left$(rData, 13)) = "/ADVERTENCIA " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        '/carcel nick@motivo
        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        rData = Right$(rData, Len(rData) - 13)

        Name = ReadField(1, rData, Asc("@"))
        tStr = ReadField(2, rData, Asc("@"))

        If Name = "" Or tStr = "" Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Utilice /advertencia nick@motivo" & FONTTYPE_INFO)
            Exit Sub

        End If

        TIndex = NameIndex(Name)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(TIndex).flags.Privilegios > PlayerType.User Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes advertir a administradores." & FONTTYPE_INFO)
            Exit Sub

        End If

        Name = Replace(Name, "\", "")
        Name = Replace(Name, "/", "")

        If FileExist(CharPath & Name & ".chr", vbNormal) Then
            tInt = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & Name & ".chr", "PENAS", "Cant", tInt + 1)
            Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & ": ADVERTENCIA por: " & LCase$( _
                                                                             tStr) & " " & Date & " " & Time)

        End If

        Call LogGM(UserList(UserIndex).Name, " advirtio a " & Name)
        Exit Sub

    End If

    If UCase$(Left$(rData, 5)) = "/MOD " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        rData = UCase$(Right$(rData, Len(rData) - 5))
        tStr = Replace$(ReadField(1, rData, 32), "+", " ")
        TIndex = NameIndex(tStr)

        If LCase$(tStr) = "yo" Then
            TIndex = UserIndex

        End If

        Arg1 = ReadField(2, rData, 32)
        Arg2 = ReadField(3, rData, 32)
        Arg3 = ReadField(4, rData, 32)
        Arg4 = ReadField(5, rData, 32)

        If UserList(UserIndex).flags.EsRolesMaster Then

            Select Case UserList(UserIndex).flags.Privilegios

            Case PlayerType.Consejero

                ' Los RMs consejeros sólo se pueden editar su head, body y exp
                If NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                If Arg1 <> "BODY" And Arg1 <> "HEAD" And Arg1 <> "LEVEL" Then Exit Sub

            Case PlayerType.SemiDios

                ' Los RMs sólo se pueden editar su level y el head y body de cualquiera
                If Arg1 = "EXP" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                If Arg1 <> "BODY" And Arg1 <> "HEAD" Then Exit Sub

            Case PlayerType.Dios

                ' Si quiere modificar el level sólo lo puede hacer sobre sí mismo
                If Arg1 = "NIVEL" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                If Arg1 = "LEVEL" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                If Arg1 = "ORO" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub

                ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                If Arg1 <> "BODY" And Arg1 <> "HEAD" And Arg1 <> "CIU" And Arg1 <> "CRI" And Arg1 <> "CLASE" And Arg1 <> "SKILLS" And Arg1 <> "RAZA" Then Exit Sub

            End Select

        ElseIf UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then   'Si no es RM debe ser dios para poder usar este comando
            Exit Sub

        End If

        Select Case Arg1

        Case "VIDA"
            Dim MaxVida As Long
            Dim ChangeVida As Long

            MaxVida = "32000"
            ChangeVida = ReadField(3, rData, 32)

            If ChangeVida > MaxVida Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No se ha cambiado el valor, por que has superado la maxima vida. (Max: " & _
                                                                MaxVida & ")" & FONTTYPE_INFO)
                Exit Sub

            End If

            If ChangeVida <= MaxVida Then

                If UserList(TIndex).Stats.MaxHP < ChangeVida Then
                    UserList(TIndex).Stats.MinHP = ChangeVida
                End If

                UserList(TIndex).Stats.MinHP = ChangeVida
                UserList(TIndex).Stats.MaxHP = ChangeVida
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has cambiado la vida maxima del personaje " & tStr & " ahora es: " & _
                                                                ChangeVida & FONTTYPE_INFO)
                Call SendData(SendTarget.toIndex, TIndex, 0, "MXVID" & ChangeVida)
                Call EnviarHP(UserIndex)

            End If

            Exit Sub

        Case "MANA"
            Dim MaxMana As Long
            Dim ChangeMana As Long

            MaxMana = "32000"
            ChangeMana = val(ReadField(3, rData, 32))

            If ChangeMana > MaxMana Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No se ha cambiado el valor, por que has superado la maxima mana. (Max: " & _
                                                                MaxMana & ")" & FONTTYPE_INFO)
                Exit Sub

            End If

            If ChangeMana <= MaxMana Then

                UserList(TIndex).Stats.MinMAN = ChangeMana
                UserList(TIndex).Stats.MaxMAN = ChangeMana
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has cambiado la mana maxima del personaje " & tStr & " ahora es: " & _
                                                                ChangeMana & FONTTYPE_INFO)
                Call SendData(SendTarget.toIndex, TIndex, 0, "MXMAN" & ChangeMana)
                Call EnviarMn(UserIndex)
            End If

            Exit Sub

        Case "NIVEL"
            Dim MassNivel As Long
            Dim ResultMassNivel As Long
            Dim ExpMAX As Long
            Dim ExpMIN As Long
            Dim ExpLvl As Long
            Dim XN As Long

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline: " & tStr & FONTTYPE_INFO)
                Exit Sub

            End If

            MassNivel = val(Arg2)
            ExpMAX = UserList(TIndex).Stats.ELU
            ExpMIN = UserList(TIndex).Stats.Exp

            If Not IsNumeric(Arg2) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El Nivel debe ser númerica." & FONTTYPE_GUILD)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis: /Mod " & UserList(TIndex).Name & " NIVEL 2" & FONTTYPE_GUILD)
                Exit Sub

            End If

            If ExpMAX = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " tiene el nivel máximo." & _
                                                                FONTTYPE_INFO)
                Exit Sub

            End If

            For XN = 1 To MassNivel

                If ExpMAX = "0" Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & _
                                                                  " subió de nivel pero llego al nivel máximo." & FONTTYPE_INFO)
                    Exit For

                End If

                ExpMAX = UserList(TIndex).Stats.ELU
                ExpMIN = UserList(TIndex).Stats.Exp

                ResultMassNivel = ExpMAX - ExpMIN
                UserList(TIndex).Stats.Exp = UserList(TIndex).Stats.Exp + ResultMassNivel

                Call EnviarExp(TIndex)
                Call CheckUserLevel(TIndex)

            Next XN

            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " ha subido de nivel." & FONTTYPE_Motd1)

            Exit Sub

        Case "ORO"

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If

            If Left$(Arg2, 1) = "-" Then

                If UserList(TIndex).Stats.GLD = 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no tiene oro!!" & FONTTYPE_INFO)
                    Exit Sub
                End If

                UserList(TIndex).Stats.GLD = UserList(TIndex).Stats.GLD - val(mid(Arg2, 2))
                Call EnviarOro(TIndex)

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has quitado el oro de " & UserList(TIndex).Name & " con resta de: " & val(Arg2) & " de oro." & FONTTYPE_INFO)
                Call SendData(SendTarget.toIndex, TIndex, 0, "||El GM " & UserList(UserIndex).Name & " te ha quitado: " & val(Arg2) & " de oro." & FONTTYPE_INFO)
                Exit Sub

            Else

                If val(Arg2) > MaxOro Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has superado el limite de maximo oro: " & MaxOro & FONTTYPE_INFO)
                    Exit Sub
                End If

                UserList(TIndex).Stats.GLD = UserList(TIndex).Stats.GLD + val(Arg2)
                Call EnviarOro(TIndex)

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has aumentado el oro de " & UserList(TIndex).Name & " con suma de: " & val(Arg2) & " de oro." & FONTTYPE_INFO)
                Call SendData(SendTarget.toIndex, TIndex, 0, "||El GM " & UserList(UserIndex).Name & " te ha dado: " & val(Arg2) & " de oro." & FONTTYPE_INFO)
            End If

        Case "EXP"

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(TIndex).Stats.Exp + val(Arg2) > UserList(TIndex).Stats.ELU Then
                UserList(TIndex).Stats.Exp = UserList(TIndex).Stats.Exp + val(Arg2)
                Call CheckUserLevel(TIndex)
            Else
                UserList(TIndex).Stats.Exp = val(Arg2)

            End If

            Call EnviarExp(TIndex)
            Exit Sub

        Case "BODY"

            If TIndex <= 0 Then
                Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "INIT", "Body", Arg2)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
                Exit Sub

            End If


            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(TIndex).pos.Map, TIndex, val(Arg2), UserList(TIndex).char.Head, UserList( _
                                                                                                                              TIndex).char.heading, UserList(TIndex).char.WeaponAnim, UserList(TIndex).char.ShieldAnim, UserList(TIndex).char.CascoAnim, _
                                UserList(TIndex).char.Alas)

            Exit Sub

        Case "HEAD"

            If TIndex <= 0 Then
                Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "INIT", "Head", Arg2)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
                Exit Sub

            End If

            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(TIndex).pos.Map, TIndex, UserList(TIndex).char.Body, val(Arg2), UserList( _
                                                                                                                              TIndex).char.heading, UserList(TIndex).char.WeaponAnim, UserList(TIndex).char.ShieldAnim, UserList(TIndex).char.CascoAnim, _
                                UserList(TIndex).char.Alas)
            Exit Sub

        Case "CRI"

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                Exit Sub

            End If

            UserList(TIndex).Faccion.CriminalesMatados = val(Arg2)
            Exit Sub

        Case "CIU"

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                Exit Sub

            End If

            UserList(TIndex).Faccion.CiudadanosMatados = val(Arg2)
            Exit Sub

        Case "CLASE"

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline: " & tStr & FONTTYPE_INFO)
                Exit Sub
            End If

            If UserList(TIndex).Clase = UCase$(Left$(Arg2, 1)) & UCase$(mid$(Arg2, 2)) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||La clase de: " & tStr & " no ha cambiado porque ya es: " & UCase$(Left$(Arg2, 1)) & UCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                Exit Sub
            End If

            If Len(Arg2) > 1 Then
                UserList(TIndex).Clase = UCase$(Left$(Arg2, 1)) & UCase$(mid$(Arg2, 2))
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Clase cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
            Else
                UserList(TIndex).Clase = UCase$(Arg2)
            End If

        Case "RAZA"

            If TIndex <= 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline: " & tStr & FONTTYPE_INFO)
                Exit Sub
            End If

            If UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) = UserList(TIndex).Raza Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||La raza de: " & tStr & " ya no ha cambiado porque ya es: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                Exit Sub
            End If

            Select Case UCase$(Arg2)

            Case "HUMANO"
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                Call DarCuerpoDesnudo(TIndex)

            Case "ENANO"
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                Call DarCuerpoDesnudo(TIndex)

            Case "HOBBIT"
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                Call DarCuerpoDesnudo(TIndex)

            Case "ELFO"
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                Call DarCuerpoDesnudo(TIndex)

            Case "ELFO OSCURO"
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                Call DarCuerpoDesnudo(TIndex)

            Case "LICANTROPO"
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                Call DarCuerpoDesnudo(TIndex)

            Case "GNOMO"
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                Call DarCuerpoDesnudo(TIndex)

            Case "ORCO"
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                Call DarCuerpoDesnudo(TIndex)

            Case "VAMPIRO"
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                Call DarCuerpoDesnudo(TIndex)

            Case "CICLOPE"
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                Call DarCuerpoDesnudo(TIndex)

            Case Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||La raza: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & " no existe." & FONTTYPE_INFO)
            End Select

        Case "SKILLS"

            For LoopC = 1 To NUMSKILLS
                If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg2) Then n = LoopC
            Next LoopC

            If n = 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Skill Inexistente!" & FONTTYPE_INFO)
                Exit Sub
            End If

            If TIndex = 0 Then
                Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "Skills", "SK" & n, Arg3)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
            Else
                UserList(TIndex).Stats.UserSkills(n) = val(Arg3)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has cambiado la skill de " & SkillsNames(n) & " a " & UserList(TIndex).Name & " por: " & val(Arg3) & FONTTYPE_INFO)
                Call SendData(SendTarget.toIndex, TIndex, 0, "||GM " & UserList(UserIndex).Name & " te ha cambiado el valor de la skill " & SkillsNames(n) & " a: " & val(Arg3) & FONTTYPE_INFO)
            End If

            Exit Sub

        Case "SKILLSLIBRES"
            Dim SLName As String
            Dim SLSkills As Integer
            Dim SLResult As Integer

            If Left(Arg2, 1) = "-" Then

                If TIndex = 0 Then
                    SLName = ReadField(1, rData, 32)

                    If Not FileExist(CharPath & UCase(SLName) & ".chr") Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje: " & SLName & " no existe." & FONTTYPE_INFO)
                        Exit Sub
                    End If

                    SLSkills = GetVar(CharPath & UCase$(SLName) & ".chr", "STATS", "SkillptsLibres")

                    SLResult = SLSkills - mid(Arg2, 2)

                    If SLResult < 0 Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||El cambio no se ha podido efectuar:" & SLResult & FONTTYPE_INFO)
                        Exit Sub
                    Else
                        Call WriteVar(CharPath & UCase$(SLName) & ".chr", "STATS", "SkillPtsLibres", SLResult)
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||El charfile de " & SLName & " ha cambiado su SkillsLibres de " & SLSkills & " a " & SLResult & FONTTYPE_INFO)
                        Exit Sub
                    End If

                Else
                    SLName = UserList(TIndex).Name
                    SLSkills = UserList(TIndex).Stats.SkillPts

                    SLResult = SLSkills - mid(Arg2, 2)

                    If SLResult < 0 Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||El cambio no se ha podido efectuar:" & SLResult & FONTTYPE_INFO)
                        Exit Sub
                    Else
                        UserList(TIndex).Stats.SkillPts = SLResult
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Las skills de " & SLName & " han sido modificadas." & FONTTYPE_INFO)
                        Call EnviarSkills(TIndex)
                        Exit Sub
                    End If
                End If


            Else    'Parte donde Suma

                If TIndex = 0 Then

                    SLName = ReadField(1, rData, 32)

                    If Not FileExist(CharPath & UCase(SLName) & ".chr") Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje: " & SLName & " no existe." & FONTTYPE_INFO)
                        Exit Sub
                    End If

                    SLSkills = GetVar(CharPath & UCase$(SLName) & ".chr", "STATS", "SkillptsLibres")

                    SLResult = SLSkills + Arg2

                    Call WriteVar(CharPath & UCase$(SLName) & ".chr", "STATS", "SkillPtsLibres", SLResult)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El charfile de " & SLName & " ha cambiado su SkillsLibres de " & SLSkills & " a " & SLResult & FONTTYPE_INFO)
                    Exit Sub

                Else
                    SLName = UserList(TIndex).Name
                    SLSkills = UserList(TIndex).Stats.SkillPts
                    SLResult = SLSkills + Arg2

                    UserList(TIndex).Stats.SkillPts = SLResult
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Las skills de " & SLName & " han sido modificadas." & FONTTYPE_INFO)
                    Call EnviarSkills(TIndex)
                    Exit Sub

                End If

            End If

            Exit Sub

        Case Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Sintaxis incorrecto" & FONTTYPE_GUILD)
            Call SendData(SendTarget.toIndex, UserIndex, 0, _
                          "||Comando: /MOD <Nick/yo> <NIVEL/SKILLS/SKILLSLIBRES/ORO/CIU/CRI/EXP/BODY/HEAD> <VALOR>" & FONTTYPE_GUILD)
            Exit Sub

        End Select

        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "/SUBIR " Then

        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        rData = Right$(rData, Len(rData) - 7)

        TIndex = NameIndex(ReadField(1, rData, 32))

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline: " & tStr & FONTTYPE_INFO)
            Exit Sub

        End If

        MassNivel = ReadField(2, rData, 32)
        ExpMAX = UserList(TIndex).Stats.ELU
        ExpMIN = UserList(TIndex).Stats.Exp

        If Not IsNumeric(MassNivel) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El Nivel debe ser númerica." & FONTTYPE_GUILD)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis: /SUBIR " & UserList(TIndex).Name & " 2" & FONTTYPE_GUILD)
            Exit Sub

        End If

        If ExpMAX = 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " tiene el nivel máximo." & _
                                                            FONTTYPE_INFO)
            Exit Sub

        End If

        For XN = 1 To MassNivel

            If ExpMAX = "0" Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & _
                                                              " subió de nivel pero llego al nivel máximo." & FONTTYPE_INFO)
                Exit For

            End If

            ExpMAX = UserList(TIndex).Stats.ELU
            ExpMIN = UserList(TIndex).Stats.Exp

            ResultMassNivel = ExpMAX - ExpMIN
            UserList(TIndex).Stats.Exp = UserList(TIndex).Stats.Exp + ResultMassNivel

            Call EnviarExp(TIndex)
            Call CheckUserLevel(TIndex)

        Next XN

        Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " ha subido de nivel." & FONTTYPE_Motd1)

        Exit Sub
    End If


    If UCase$(Left$(rData, 6)) = "/INFO " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        rData = Right$(rData, Len(rData) - 6)

        TIndex = NameIndex(rData)

        If TIndex <= 0 Then

            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline, Buscando en Charfile." & FONTTYPE_INFO)
            SendUserStatsTxtOFF UserIndex, rData
        Else

            If UserList(TIndex).flags.Privilegios >= PlayerType.Dios Then Exit Sub
            SendUserStatsTxt UserIndex, TIndex

        End If

        Exit Sub

    End If

    'MINISTATS DEL USER
    If UCase$(Left$(rData, 6)) = "/STAT " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        rData = Right$(rData, Len(rData) - 6)

        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline. Leyendo Charfile... " & FONTTYPE_INFO)
            SendUserMiniStatsTxtFromChar UserIndex, rData
        Else
            SendUserMiniStatsTxt UserIndex, TIndex

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 5)) = "/BAL " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 5)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline. Leyendo charfile... " & FONTTYPE_TALK)
            SendUserOROTxtFromChar UserIndex, rData
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El usuario " & rData & " tiene " & UserList(TIndex).Stats.Banco & " en el banco" & _
                                                            FONTTYPE_TALK)

        End If

        Exit Sub

    End If


    If UCase$(Left$(rData, 8)) = "/QUITAR " Then
        Dim QuitObjeto As Obj
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)


        rData = Right$(rData, Len(rData) - 8)

        TIndex = NameIndex(ReadField(1, rData, 32))
        QuitObjeto.ObjIndex = ReadField(2, rData, 32)
        QuitObjeto.Amount = ReadField(3, rData, 32)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||ERROR: El usuario no esta conectado." & FONTTYPE_EJECUCION)
        Else
            Call QuitarObjetos(QuitObjeto.ObjIndex, QuitObjeto.Amount, TIndex)
        End If

        Exit Sub
    End If

    If UCase$(Left$(rData, 5)) = "/INV " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        rData = Right$(rData, Len(rData) - 5)

        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Else
            SendUserInvTxt UserIndex, TIndex

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 11)) = "/QUITARBOV " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        rData = Right$(rData, Len(rData) - 11)

        TIndex = NameIndex(ReadField(1, rData, 32))
        QuitObjeto.ObjIndex = ReadField(2, rData, 32)
        QuitObjeto.Amount = ReadField(3, rData, 32)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||ERROR: El usuario no esta conectado." & FONTTYPE_EJECUCION)
        Else
            Call QuitarObjetosBov(QuitObjeto.ObjIndex, QuitObjeto.Amount, TIndex)
        End If

        Exit Sub
    End If

    If UCase$(Left(rData, 7)) = "/CLAVE " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 7)

        tStr = ReadField(1, rData, 32)
        TIndex = NameIndex(tStr)

        If TIndex > 0 Then
            Call SendData(toIndex, UserIndex, 0, "||El usuario debe desconectarse para realizar el cambio de clave." & FONTTYPE_INFO)
            Exit Sub
        End If

        If FileExist(CharPath & UCase$(tStr) & ".chr", vbNormal) Then

            If UCase$(GetVar(CharPath & tStr & ".chr", "CONTACTO", "Email")) <> UCase$(ReadField(2, rData, 32)) Then
                Call SendData(toIndex, UserIndex, 0, "||El email no coincide." & FONTTYPE_INFO)
                Exit Sub
            Else

                For i = 1 To 5
                    tInt = RandomNumber(65, 90)
                    Arg1 = Arg1 + Chr$(tInt)
                Next i

                Call SendData(toIndex, UserIndex, 0, "||  Su email es:" & ReadField(2, rData, 32) & FONTTYPE_INFO)
                Call SendData(toIndex, UserIndex, 0, "|| La Ultima Ip es:" & GetVar(CharPath & UCase$(tStr) & ".chr", "INIT", "LASTIP") & FONTTYPE_INFO)
                Call SendData(toIndex, UserIndex, 0, "||La nueva clave es: " & Arg1 & FONTTYPE_INFO)
                Arg1 = MD5String(Arg1)
                Call WriteVar(CharPath & UCase$(tStr) & ".chr", "INIT", "PASSWORD", Arg1)
                Exit Sub
            End If

        Else
            Call SendData(toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
        End If

        Exit Sub
    End If

    If UCase$(Left$(rData, 5)) = "/BOV " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        rData = Right$(rData, Len(rData) - 5)

        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Else
            SendUserBovedaTxt UserIndex, TIndex
        End If

        Exit Sub

    End If

    'SKILLS DEL USER
    If UCase$(Left$(rData, 8)) = "/SKILLS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        rData = Right$(rData, Len(rData) - 8)

        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call Replace(rData, "\", " ")
            Call Replace(rData, "/", " ")

            For tInt = 1 To NUMSKILLS
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| CHAR>" & SkillsNames(tInt) & " = " & GetVar(CharPath & rData & ".chr", _
                                                                                                                "SKILLS", "SK" & tInt) & FONTTYPE_INFO)
            Next tInt

            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| CHAR> Libres:" & GetVar(CharPath & rData & ".chr", "STATS", "SKILLPTSLIBRES") & _
                                                            FONTTYPE_INFO)
            Exit Sub

        End If

        SendUserSkillsTxt UserIndex, TIndex
        Exit Sub

    End If

    If UCase$(Left$(rData, 9)) = "/REVIVIR " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 9)
        Name = rData

        If UCase$(Name) <> "YO" Then
            TIndex = NameIndex(Name)
        Else
            TIndex = UserIndex

        End If

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(TIndex).flags.Muerto = 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario esta vivo." & FONTTYPE_INFO)
            Exit Sub
        End If

        UserList(TIndex).flags.Muerto = 0
        UserList(TIndex).Stats.MinHP = UserList(TIndex).Stats.MaxHP
        Call DarCuerpoDesnudo(TIndex)

        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(TIndex).pos.Map, val(TIndex), UserList(TIndex).char.Body, UserList(TIndex).OrigChar.Head, _
                            UserList(TIndex).char.heading, UserList(TIndex).char.WeaponAnim, UserList(TIndex).char.ShieldAnim, UserList(TIndex).char.CascoAnim, _
                            UserList(TIndex).char.Alas)

        Call SendUserStatsBox(val(TIndex))
        Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " te ha resucitado." & FONTTYPE_INFO)

        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "/MATAUS" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 7)

        Call UserDie(UserList(UserIndex).flags.TargetUser)

        Exit Sub
    End If

    If UCase$(rData) = "/ONLINEGM" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        For LoopC = 1 To LastUser

            'Tiene nombre? Es GM? Si es Dios o Admin, nosotros lo somos también??
            If (UserList(LoopC).Name <> "") And UserList(LoopC).flags.Privilegios > PlayerType.User And (UserList(LoopC).flags.Privilegios < _
                                                                                                         PlayerType.Dios Or UserList(UserIndex).flags.Privilegios >= PlayerType.Dios) Then
                tStr = tStr & UserList(LoopC).Name & ", "

            End If

        Next LoopC

        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No hay GMs Online" & FONTTYPE_INFO)

        End If

        Exit Sub

    End If

    If UCase$(rData) = "/ONLINEMAP" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        For LoopC = 1 To LastUser

            If UserList(LoopC).Name <> "" And UserList(LoopC).pos.Map = UserList(UserIndex).pos.Map And (UserList(LoopC).flags.Privilegios < _
                                                                                                         PlayerType.Dios Or UserList(UserIndex).flags.Privilegios >= PlayerType.Dios) Then
                tStr = tStr & UserList(LoopC).Name & ", "

            End If

        Next LoopC

        If Len(tStr) > 2 Then tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuarios en el mapa: " & tStr & FONTTYPE_INFO)
        Exit Sub

    End If

    If UCase$(rData) = "/ONLINECLASE" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        For LoopC = 1 To LastUser

            If (UserList(LoopC).Name <> "") And UserList(LoopC).flags.Privilegios < PlayerType.Consejero Then
                tStr = tStr & UserList(LoopC).Name & "(" & UserList(LoopC).Clase & "), "
            End If

        Next LoopC

        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No hay usuarios Online." & FONTTYPE_INFO)
        End If

        Exit Sub
    End If

    If UCase$(rData) = "/DRUIDAS" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        For LoopC = 1 To LastUser
            If (UserList(LoopC).Name <> "") And UserList(LoopC).flags.Privilegios < PlayerType.Consejero And UCase$(UserList(LoopC).Clase) = "DRUIDA" Then
                tStr = tStr & UserList(LoopC).Name & "(" & UserList(UserIndex).Clase & "), "
            End If
        Next LoopC

        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No hay druidas Online." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If

    If UCase$(Left$(rData, 7)) = "/PERDON" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 8)
        TIndex = NameIndex(rData)

        Call VolverCiudadano(TIndex)
        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "/ECHAR " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 7)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes echar a alguien con jerarquia mayor a la tuya." & FONTTYPE_INFO)
            Exit Sub

        End If

        Call SendData(SendTarget.toall, 0, 0, "||" & UserList(UserIndex).Name & " echo a " & UserList(TIndex).Name & "." & FONTTYPE_INFO)
        Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(TIndex).Name)
        Call CloseSocket(TIndex)
        Exit Sub

    End If

    If UCase$(Left$(rData, 14)) = "/CASTIGORETOS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 14)

        Arg1 = rData
        TIndex = NameIndex(rData)

        If TIndex > 0 Then
            If UserList(TIndex).Stats.PuntosRetos > 0 Then
                Call SendData(toIndex, UserIndex, 0, "||Los puntos retos del usuario " & UserList(TIndex).Name & " han sido castigados." & FONTTYPE_INFO)
                UserList(TIndex).Stats.PuntosRetos = 0
            Else
                Call SendData(toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " no tiene puntos retos." & FONTTYPE_INFO)
            End If
        Else

            If FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
            Else
                If GetVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSRETOS") > 0 Then
                    Call SendData(toIndex, UserIndex, 0, "||Los puntos retos del usuario " & Arg1 & " han sido castigados." & FONTTYPE_INFO)
                    Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSRETOS", "0")
                Else
                    Call SendData(toIndex, UserIndex, 0, "||El usuario " & Arg1 & " no tiene puntos retos." & FONTTYPE_INFO)
                End If
            End If

        End If

        Exit Sub
    End If

    If UCase$(Left$(rData, 15)) = "/CASTIGOPUNTOS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 15)

        Arg1 = rData
        TIndex = NameIndex(rData)

        If TIndex > 0 Then
            If UserList(TIndex).Stats.PuntosDuelos > 0 Then
                Call SendData(toIndex, UserIndex, 0, "||Los puntos del usuario " & UserList(TIndex).Name & " han sido castigados." & FONTTYPE_INFO)
                UserList(TIndex).Stats.PuntosDuelos = 0
            Else
                Call SendData(toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " no tiene puntos." & FONTTYPE_INFO)
            End If
        Else

            If FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
            Else
                If GetVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSDUELOS") > 0 Then
                    Call SendData(toIndex, UserIndex, 0, "||Los puntos del usuario " & Arg1 & " han sido castigados." & FONTTYPE_INFO)
                    Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSDUELOS", "0")
                Else
                    Call SendData(toIndex, UserIndex, 0, "||El usuario " & Arg1 & " no tiene puntos." & FONTTYPE_INFO)
                End If
            End If

        End If

        Exit Sub
    End If

    If UCase$(Left$(rData, 15)) = "/CASTIGOTORNEO " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 15)

        Arg1 = rData
        TIndex = NameIndex(rData)

        If TIndex > 0 Then
            If UserList(TIndex).Stats.PuntosTorneo > 0 Then
                Call SendData(toIndex, UserIndex, 0, "||Los puntos torneo del usuario " & UserList(TIndex).Name & " han sido castigados." & FONTTYPE_INFO)
                UserList(TIndex).Stats.PuntosTorneo = 0
            Else
                Call SendData(toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " no tiene puntos torneo." & FONTTYPE_INFO)
            End If
        Else

            If FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
            Else
                If GetVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSTORNEO") > 0 Then
                    Call SendData(toIndex, UserIndex, 0, "||Los puntos torneo del usuario " & Arg1 & " han sido castigados." & FONTTYPE_INFO)
                    Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSTORNEO", "0")
                Else
                    Call SendData(toIndex, UserIndex, 0, "||El usuario " & Arg1 & " no tiene puntos torneo." & FONTTYPE_INFO)
                End If
            End If

        End If

        Exit Sub
    End If

    If UCase$(Left(rData, 13)) = "/CASTIGOCLAN " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 13)

        Arg1 = rData
        TIndex = NameIndex(rData)

        If TIndex > 0 Then
            If UserList(TIndex).Clan.PuntosClan > 0 Then
                Call SendData(toIndex, UserIndex, 0, "||Los puntos clan del usuario " & UserList(TIndex).Name & " han sido castigados." & FONTTYPE_INFO)
                UserList(TIndex).Clan.PuntosClan = 0
            Else
                Call SendData(toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " no tiene puntos clan." & FONTTYPE_INFO)
            End If
        Else

            If FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
            Else
                If GetVar(CharPath & UCase$(Arg1) & ".chr", "GUILD", "PUNTOSCLAN") > 0 Then
                    Call SendData(toIndex, UserIndex, 0, "||Los puntos clan del usuario " & Arg1 & " han sido castigados." & FONTTYPE_INFO)
                    Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "GUILD", "PUNTOSCLAN", "0")
                Else
                    Call SendData(toIndex, UserIndex, 0, "||El usuario " & Arg1 & " no tiene puntos clan." & FONTTYPE_INFO)
                End If
            End If

        End If

        Exit Sub
    End If

    If UCase$(Left$(rData, 14)) = "/CASTIGOTODOS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 14)

        Arg1 = rData
        TIndex = NameIndex(rData)

        If TIndex > 0 Then

            If UserList(TIndex).Stats.PuntosRetos > 0 Then
                UserList(TIndex).Stats.PuntosRetos = 0
                tInt = tInt + 1
            End If

            If UserList(TIndex).Stats.PuntosDuelos > 0 Then
                UserList(TIndex).Stats.PuntosDuelos = 0
                tInt = 1
            End If

            If UserList(TIndex).Stats.PuntosTorneo > 0 Then
                UserList(TIndex).Stats.PuntosTorneo = 0
                tInt = 1
            End If

            If UserList(TIndex).Clan.PuntosClan > 0 Then
                UserList(TIndex).Clan.PuntosClan = 0
                tInt = tInt + 1
            End If

            If tInt > 0 Then
                Call SendData(toIndex, UserIndex, 0, "||Todos los puntos del usuario " & UserList(TIndex).Name & " han sido castigados." & FONTTYPE_INFO)
            Else
                Call SendData(toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " no tiene puntos." & FONTTYPE_INFO)
            End If

        Else

            If FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
            Else
                If GetVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSRETOS") > 0 Then
                    Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSRETOS", "0")
                    tInt = tInt + 1
                End If
            End If

            If FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
            Else
                If GetVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSDUELOS") > 0 Then
                    Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSDUELOS", "0")
                    tInt = tInt + 1
                End If
            End If

            If FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
            Else
                If GetVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSTORNEO") > 0 Then
                    Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "STATS", "PUNTOSTORNEO", "0")
                    tInt = tInt + 1
                End If
            End If

            If FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) = False Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
            Else
                If GetVar(CharPath & UCase$(Arg1) & ".chr", "GUILD", "PUNTOSCLAN") > 0 Then
                    Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "GUILD", "PUNTOSCLAN", "0")
                    tInt = tInt + 1
                End If
            End If

            If tInt > 0 Then
                Call SendData(toIndex, UserIndex, 0, "||Todos los puntos del usuario " & Arg1 & " han sido castigados." & FONTTYPE_INFO)
            Else
                Call SendData(toIndex, UserIndex, 0, "||El usuario " & Arg1 & " no tiene puntos." & FONTTYPE_INFO)
            End If

        End If

        Exit Sub
    End If

    If UCase$(Left$(rData, 5)) = "/BAN " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 5)
        tStr = ReadField(2, rData, Asc("@"))    ' NICK
        TIndex = NameIndex(tStr)
        Name = ReadField(1, rData, Asc("@"))    ' MOTIVO

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_TALK)

            If FileExist(CharPath & tStr & ".chr", vbNormal) Then
                tLong = UserDarPrivilegioLevel(tStr)

                If tLong > UserList(UserIndex).flags.Privilegios Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estás loco??! No podés banear a alguien de mayor jerarquia que vos!" & _
                                                                    FONTTYPE_INFO)
                    Exit Sub

                End If

                If GetVar(CharPath & tStr & ".chr", "FLAGS", "Ban") <> "0" Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje ya ha sido baneado anteriormente." & FONTTYPE_INFO)
                    Exit Sub

                End If

                Call LogBanFromName(tStr, UserIndex, Name)
                Call SendData(SendTarget.ToAdmins, 0, 0, "||AOMania> El GM & " & UserList(UserIndex).Name & "baneó a " & tStr & "." & FONTTYPE_SERVER)

                'ponemos el flag de ban a 1
                Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
                'ponemos la pena
                tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
                Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
                Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & _
                                                                               " Lo Baneó por el siguiente motivo: " & LCase$(Name) & " " & Date & " " & Time)

                If tLong > 0 Then
                    UserList(UserIndex).flags.Ban = 1
                    Call CloseSocket(UserIndex)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||" & " El gm " & UserList(UserIndex).Name & _
                                                           " fue baneado por el propio servidor por intentar banear a otro admin." & FONTTYPE_FIGHT)

                End If

                Call LogGM(UserList(UserIndex).Name, "BAN a " & tStr)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El pj " & tStr & " no existe." & FONTTYPE_INFO)

            End If

        Else

            If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes banear a al alguien de mayor jerarquia." & FONTTYPE_INFO)
                Exit Sub

            End If

            Call LogBan(TIndex, UserIndex, Name)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " Ha baneado a " & UserList(TIndex).Name & _
                                                     FONTTYPE_Motd4)

            'Ponemos el flag de ban a 1
            UserList(TIndex).flags.Ban = 1

            If UserList(TIndex).flags.Privilegios > PlayerType.User Then
                UserList(UserIndex).flags.Ban = 1
                Call CloseSocket(UserIndex)
                Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " banned by the server por bannear un Administrador." & _
                                                         FONTTYPE_FIGHT)

            End If

            Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(TIndex).Name)

            'ponemos el flag de ban a 1
            Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
            'ponemos la pena
            tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & " Lo Baneó Debido a: " & LCase$( _
                                                                             Name) & " " & Date & " " & Time)

            Call CloseSocket(TIndex)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "/UNBAN " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 7)

        rData = Replace(rData, "\", "")
        rData = Replace(rData, "/", "")

        If Not FileExist(CharPath & rData & ".chr", vbNormal) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Charfile inexistente (no use +)" & FONTTYPE_INFO)
            Exit Sub

        End If

        Call UnBan(rData)

        'penas
        i = val(GetVar(CharPath & rData & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & rData & ".chr", "PENAS", "Cant", i + 1)
        Call WriteVar(CharPath & rData & ".chr", "PENAS", "P" & i + 1, LCase$(UserList(UserIndex).Name) & " Lo unbaneó. " & Date & " " & Time)

        Call LogGM(UserList(UserIndex).Name, "/UNBAN a " & rData)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & rData & " desbaneado." & FONTTYPE_INFO)

        Exit Sub

    End If

    'SEGUIR
    If UCase$(rData) = "/SEGUIR" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.TargetNpc > 0 Then
            Call DoFollow(UserList(UserIndex).flags.TargetNpc, UserList(UserIndex).Name)

        End If

        Exit Sub

    End If

    If UCase(rData) = "/BLOQ" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Blocked = 0 Then
            MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Blocked = 1
            Call Bloquear(SendTarget.ToMap, UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, _
                          UserList(UserIndex).pos.Y, 1)
        Else
            MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Blocked = 0
            Call Bloquear(SendTarget.ToMap, UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, _
                          UserList(UserIndex).pos.Y, 0)

        End If

        Exit Sub

    End If

    If UCase$(rData) = "/ACTCOM" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If ComerciarAc = True Then
            ComerciarAc = False
            Call SendData(SendTarget.toall, 0, 0, "||Comercio entre usuarios desactivados!!." & FONTTYPE_CYAN)
        Else
            Call SendData(SendTarget.toall, 0, 0, "||Comercio entre usuarios activados!!." & FONTTYPE_CYAN)
            ComerciarAc = True

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 8)) = "/AOMANIA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 8)

        Call GuerraBanda.Ban_Comienza("32")

    End If

    'Crear criatura
    If UCase$(Left$(rData, 3)) = "/CC" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Call EnviarSpawnList(UserIndex)
        Exit Sub

    End If

    'Spawn!!!!! ¿What?
    If UCase$(Left$(rData, 3)) = "SPA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 3)

        If val(rData) > 0 And val(rData) < UBound(SpawnList) + 1 Then Call SpawnNpc(SpawnList(val(rData)).NpcIndex, UserList(UserIndex).pos, True, _
                                                                                    False)
        Call LogGM(UserList(UserIndex).Name, "Sumoneo " & SpawnList(val(rData)).NpcName)

        Exit Sub

    End If

    'Resetea el inventario
    If UCase$(rData) = "/RESETINV" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 9)

        If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
        Call ResetNpcInv(UserList(UserIndex).flags.TargetNpc)
        Exit Sub

    End If

    '/Clean
    If UCase$(rData) = "/LIMPIAR" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call LimpiarMundo
        Exit Sub

    End If

    If UCase$(Left$(rData, 8)) = "/ONCLAN " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 8)
        tInt = GuildIndex(rData)

        If tInt > 0 Then
            tStr = modGuilds.m_ListaDeMiembrosOnline(UserIndex, tInt)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Clan " & UCase(rData) & ": " & tStr & FONTTYPE_GUILDMSG)

        End If

    End If

    'Crear Teleport
    If UCase(Left(rData, 4)) = "/CT " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        '/ct mapa_dest x_dest y_dest
        rData = Right(rData, Len(rData) - 4)
        Mapa = ReadField(1, rData, 32)
        X = ReadField(2, rData, 32)
        Y = ReadField(3, rData, 32)

        If MapaValido(Mapa) = False Or InMapBounds(Mapa, X, Y) = False Then
            Exit Sub

        End If

        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).OBJInfo.ObjIndex > 0 Then
            Exit Sub

        End If

        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.Map > 0 Then
            Exit Sub

        End If

        If MapData(Mapa, X, Y).OBJInfo.ObjIndex > 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, Mapa, "||Hay un objeto en el piso en ese lugar" & FONTTYPE_INFO)
            Exit Sub

        End If

        Dim ET As Obj
        ET.Amount = 1
        ET.ObjIndex = 378

        Call MakeObj(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, ET, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList( _
                                                                                                                                   UserIndex).pos.Y - 1)

        MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.Map = Mapa
        MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.X = X
        MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.Y = Y

        Exit Sub

    End If

    'Destruir Teleport
    'toma el ultimo click
    If UCase(Left(rData, 3)) = "/DT" Then
        '/dt

        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        Mapa = UserList(UserIndex).flags.TargetMap
        X = UserList(UserIndex).flags.TargetX
        Y = UserList(UserIndex).flags.TargetY

        If ObjData(MapData(Mapa, X, Y).OBJInfo.ObjIndex).ObjType = eOBJType.otTELEPORT And MapData(Mapa, X, Y).TileExit.Map > 0 Then
            Call EraseObj(SendTarget.ToMap, 0, Mapa, MapData(Mapa, X, Y).OBJInfo.Amount, Mapa, X, Y)
            Call EraseObj(SendTarget.ToMap, 0, MapData(Mapa, X, Y).TileExit.Map, 1, MapData(Mapa, X, Y).TileExit.Map, MapData(Mapa, X, _
                                                                                                                              Y).TileExit.X, MapData(Mapa, X, Y).TileExit.Y)
            MapData(Mapa, X, Y).TileExit.Map = 0
            MapData(Mapa, X, Y).TileExit.X = 0
            MapData(Mapa, X, Y).TileExit.Y = 0

        End If

        Exit Sub

    End If

    If UCase$(rData) = "/LLUVIA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Call SecondaryAmbient
        Exit Sub

    End If

    Select Case UCase$(Left$(rData, 8))

    Case "/TALKAS "
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.Privilegios > PlayerType.SemiDios Or UserList(UserIndex).flags.EsRolesMaster Then

            If UserList(UserIndex).flags.TargetNpc > 0 Then
                tStr = Right$(rData, Len(rData) - 8)

                Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNpc, Npclist(UserList(UserIndex).flags.TargetNpc).pos.Map, _
                              "||" & vbWhite & "°" & tStr & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, _
                              "||Debes seleccionar el NPC por el que quieres hablar antes de usar este comando" & FONTTYPE_INFO)

            End If

        End If

        Exit Sub

    End Select

    If UCase$(rData) = "/MASSEJECUTAR" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        For LoopC = 1 To LastUser

            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.UserLogged Then
                    If Not UserList(LoopC).flags.Privilegios >= 1 Then
                        If UserList(LoopC).pos.Map = UserList(UserIndex).pos.Map Then
                            Call UserDie(LoopC)
                        End If

                    End If

                End If

            End If

        Next LoopC

        Exit Sub

    End If

    '[yb]
    If UCase$(Left$(rData, 8)) = "/MASSORO" Then

        With UserList(UserIndex)

            Call LogGM(.Name, "Comando: " & rData)

            For Y = .pos.Y - MinYBorder + 1 To .pos.Y + MinYBorder - 1
                For X = .pos.X - MinXBorder + 1 To .pos.X + MinXBorder - 1

                    If InMapBounds(.pos.Map, X, Y) Then
                        If MapData(.pos.Map, X, Y).OBJInfo.ObjIndex = iORO Then
                            Call EraseObj(SendTarget.ToMap, 0, .pos.Map, 10000, .pos.Map, X, Y)

                        End If

                    End If

                Next X
            Next Y

        End With

        Exit Sub

    End If

    If UCase$(Left$(rData, 6)) = "/PASS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 6)
        TIndex = NameIndex(rData)

        If Not FileExist(CharPath & rData & ".chr") Then Exit Sub
        Arg1 = GetVar(CharPath & rData & ".chr", "INIT", "Password")
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||la pass de " & rData & " es " & Arg1 & FONTTYPE_INFO)
        Exit Sub

    End If

    If UCase$(Left$(rData, 11)) = "/KICKCONSE " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 11)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            If FileExist(CharPath & rData & ".chr") Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline, Echando de los consejos" & FONTTYPE_INFO)
                Call WriteVar(CharPath & UCase(rData) & ".chr", "CONSEJO", "PERTENECE", 0)
                Call WriteVar(CharPath & UCase(rData) & ".chr", "CONSEJO", "PERTENECECAOS", 0)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No se encuentra el charfile " & CharPath & rData & ".chr" & FONTTYPE_INFO)
                Exit Sub

            End If

        Else

            If UserList(TIndex).flags.PertAlCons > 0 Then
                Call SendData(SendTarget.toIndex, TIndex, 0, "||Has sido echado en el consejo de banderbill" & FONTTYPE_TALK & ENDC)
                UserList(TIndex).flags.PertAlCons = 0
                Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y)
                Call SendData(SendTarget.toall, 0, 0, "||" & rData & " fue expulsado del consejo de Banderbill" & FONTTYPE_CONSEJO)

            End If

            If UserList(TIndex).flags.PertAlConsCaos > 0 Then
                Call SendData(SendTarget.toIndex, TIndex, 0, "||Has sido echado en el consejo de la legión oscura" & FONTTYPE_TALK & ENDC)
                UserList(TIndex).flags.PertAlConsCaos = 0
                Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y)
                Call SendData(SendTarget.toall, 0, 0, "||" & rData & " fue expulsado del consejo de la Legión Oscura" & FONTTYPE_CONSEJOCAOS)

            End If

        End If

        Exit Sub

    End If



    If UCase(rData) = "/BANIPLIST" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        tStr = "||"

        For LoopC = 1 To BanIps.Count
            tStr = tStr & BanIps.Item(LoopC) & ", "
        Next LoopC

        tStr = tStr & FONTTYPE_INFO
        Call SendData(SendTarget.toIndex, UserIndex, 0, tStr)
        Exit Sub

    End If

    If UCase(Left(rData, 14)) = "/MIEMBROSCLAN " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Trim(Right(rData, Len(rData) - 9))

        If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
            Exit Sub

        End If

        Call LogGM(UserList(UserIndex).Name, "MIEMBROSCLAN a " & rData)

        tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))

        For i = 1 To tInt
            tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)

            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & tStr & "<" & rData & ">." & FONTTYPE_INFO)
        Next i

        Exit Sub

    End If

    If UCase(Left(rData, 9)) = "/BANCLAN " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Trim(Right(rData, Len(rData) - 9))

        If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
            Exit Sub

        End If

        Call SendData(SendTarget.toall, 0, 0, "|| " & UserList(UserIndex).Name & " banned al clan " & UCase$(rData) & FONTTYPE_FIGHT)

        Call LogGM(UserList(UserIndex).Name, "BANCLAN a " & rData)

        tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))

        For i = 1 To tInt
            tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
            'tstr es la victima
            Call Ban(tStr, "Administracion del servidor", "Clan Banned")
            TIndex = NameIndex(tStr)

            If TIndex > 0 Then

                UserList(TIndex).flags.Ban = 1
                Call CloseSocket(TIndex)

            End If

            Call SendData(SendTarget.toall, 0, 0, "||   " & tStr & "<" & rData & "> ha sido expulsado del servidor." & FONTTYPE_FIGHT)

            Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")

            n = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", n + 1)
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & n + 1, LCase$(UserList(UserIndex).Name) & ": BAN AL CLAN: " & rData & " " & Date _
                                                                        & " " & Time)

        Next i

        Exit Sub

    End If

    'Ban x IP
    If UCase(Left(rData, 7)) = "/BANIP " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Dim BanIP As String, XNick As Boolean

        rData = Right$(rData, Len(rData) - 7)
        tStr = Replace(ReadField(1, rData, Asc(" ")), "+", " ")

        TIndex = NameIndex(tStr)

        If TIndex <= 0 Then
            XNick = False
            BanIP = tStr
        Else
            XNick = True
            Call LogGM(UserList(UserIndex).Name, "/BANLAIP " & UserList(TIndex).Name & " - " & UserList(TIndex).ip)
            BanIP = UserList(TIndex).ip

        End If

        rData = Right$(rData, Len(rData) - Len(tStr))

        If BanIpBuscar(BanIP) > 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La IP " & BanIP & " ya se encuentra en la lista de bans." & FONTTYPE_INFO)
            Exit Sub

        End If

        Call BanIpAgrega(BanIP)
        Call SendData(SendTarget.ToAdmins, UserIndex, 0, "||" & UserList(UserIndex).Name & " Baneo la IP " & BanIP & FONTTYPE_FIGHT)

        If XNick = True Then
            Call LogBan(TIndex, UserIndex, "Ban por IP desde Nick por " & rData)

            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " echo a " & UserList(TIndex).Name & "." & FONTTYPE_FIGHT)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " Banned a " & UserList(TIndex).Name & "." & FONTTYPE_FIGHT)

            UserList(TIndex).flags.Ban = 1

            Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(TIndex).Name)
            Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(TIndex).Name)
            Call CloseSocket(TIndex)

        End If

        Exit Sub

    End If

    If UCase(Left(rData, 9)) = "/UNBANIP " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        rData = Right(rData, Len(rData) - 9)

        If BanIpQuita(rData) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La IP """ & rData & """ se ha quitado de la lista de bans." & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La IP """ & rData & """ NO se encuentra en la lista de bans." & FONTTYPE_INFO)

        End If

        Exit Sub

    End If

    If UCase(Left(rData, 6)) = "/GMSG " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 6)

        If rData <> "" Then
            Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & "> " & rData & FONTTYPE_TALK)
        End If

        Exit Sub
    End If

    If UCase(Left(rData, 6)) = "/UMSG " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 6)

        tName = ReadField(1, rData, 32)
        tMessage = Right$(rData, Len(rData) - (1 + Len(tName)))
        TIndex = NameIndex(tName)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub

        End If

        Call SendData(SendTarget.toIndex, TIndex, 0, "||< " & UserList(UserIndex).Name & " > te dice: " & tMessage & FONTTYPE_SERVER)

        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Le has mandado a " & tName & " : " & tMessage & FONTTYPE_SERVER)

    End If

    If UCase(Left(rData, 13)) = "/SEARCHNPCSH " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 13)

        Dim nCH As Long

        CountNpcH = 0

        For nCH = 500 To 724

            Call LeerNpcH(nCH, rData, UserIndex)

        Next nCH

    End If

    If UCase(Left(rData, 12)) = "/SEARCHNPCS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 12)

        Dim nC As Long

        CountNpc = 0

        For nC = 1 To 301

            Call LeerNpc(nC, rData, UserIndex)

        Next nC

    End If

    If UCase(Left(rData, 13)) = "/SEARCHITEMS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 13)

        Dim xci As Long
        Dim CID As Long
        Dim CNameItem As String
        Dim Asi As Long

        Asi = 0

        For xci = 1 To NumObjDatas
            CID = xci
            CNameItem = ObjData(xci).Name

            If rData = "" Then
                Asi = Asi + 1
                Call SendData(SendTarget.toIndex, UserIndex, 0, "VCTS" & CID & "#" & CNameItem)
            Else

                If InStr(LCase(CNameItem), LCase(rData)) Then

                    Asi = Asi + 1
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "VCTS" & CID & "#" & CNameItem)

                End If

            End If

        Next xci

        If Asi = 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "VITS" & Asi)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "VITS" & Asi)

        End If

        Exit Sub

    End If

    'Crear Item
    If UCase(Left(rData, 3)) = "/CI" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Dim txt As String
        Dim Cadena() As String
        Dim IdItem As String
        Dim Cantidad As String

        txt = rData

        Cadena = Split(txt, Chr$(32))

        If txt = "/CI" Or UBound(Cadena) < 2 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis incorrecto." & FONTTYPE_GUILD)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis: /CI <ID ITEM> <CANTIDAD>" & FONTTYPE_GUILD)
            Exit Sub

        End If

        If Not IsNumeric(Cadena(1)) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El ID Item debe ser númerica." & FONTTYPE_GUILD)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis: /CI <ID ITEM> <CANTIDAD> Ejemplo: /CI 2 <CANTIDAD>." & FONTTYPE_GUILD)
            Exit Sub

        End If

        If Not IsNumeric(Cadena(2)) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El Cantidad debe ser numérica." & FONTTYPE_GUILD)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis: /CI <ID ITEM> <CANTIDAD> Ejemplo: /CI 2 10." & FONTTYPE_GUILD)
            Exit Sub

        End If

        If Cadena(2) > 1200 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Has superado el tope de cantidad. (Max: 1200)" & FONTTYPE_GUILD)
            Exit Sub

        End If

        IdItem = Cadena(1)
        Cantidad = Cadena(2)

        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).OBJInfo.ObjIndex > 0 Then
            Exit Sub

        End If

        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.Map > 0 Then
            Exit Sub

        End If

        If val(IdItem) < 1 Or val(IdItem) > NumObjDatas Then
            Exit Sub

        End If

        'Is the object not null?
        If ObjData(val(IdItem)).Name = "" Then Exit Sub

        Dim Objeto As Obj

        Objeto.Amount = val(Cantidad)
        Objeto.ObjIndex = val(IdItem)

        Call MeterItemEnInventario(UserIndex, Objeto)

        Call LogGM(UserList(UserIndex).Name, "Creo: " & Cantidad & " " & ObjData(Objeto.ObjIndex).Name)

        Exit Sub

    End If

    If UCase$(Left$(rData, 15)) = "/CHAUTEMPLARIO " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 15)
        Call LogGM(UserList(UserIndex).Name, "ECHO DEL TEMPLARIO A: " & rData)

        TIndex = NameIndex(rData)
        Dim tArmIndex As Integer

        If TIndex > 0 Then
            UserList(TIndex).Faccion.Templario = 0
            UserList(TIndex).Faccion.Reenlistadas = 200
            UserList(TIndex).Faccion.RecibioArmaduraTemplaria = 0

            Call PerderItemsFaccionarios(UserIndex, UserList(TIndex).Faccion.ArmaduraFaccionaria)

            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas TEMPLARIAS y prohibida la reenlistada" & _
                                                            FONTTYPE_INFO)

            Call SendData(SendTarget.toIndex, TIndex, 0, "|| " & UserList(UserIndex).Name & _
                                                       " te ha expulsado en forma definitiva de las fuerzas TEMPLARIAS." & FONTTYPE_FIGHT)
        Else

            If FileExist(CharPath & rData & ".chr") Then
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Templario", 0)
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(UserIndex).Name)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas TEMPLARIAS y prohibida la reenlistada" & _
                                                                FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & ".chr inexistente." & FONTTYPE_INFO)

            End If

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 13)) = "/CHAUNEMESIS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 13)
        Call LogGM(UserList(UserIndex).Name, "ECHO DEL NEMESIS A: " & rData)

        TIndex = NameIndex(rData)

        If TIndex > 0 Then
            UserList(TIndex).Faccion.Nemesis = 0
            UserList(TIndex).Faccion.Reenlistadas = 200
            UserList(TIndex).Faccion.RecibioArmaduraNemesis = 0

            Call PerderItemsFaccionarios(UserIndex, UserList(TIndex).Faccion.ArmaduraFaccionaria)

            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas NEMESIS y prohibida la reenlistada" & _
                                                            FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, TIndex, 0, "|| " & UserList(UserIndex).Name & _
                                                       " te ha expulsado en forma definitiva de las fuerzas NEMESIS." & FONTTYPE_FIGHT)
        Else

            If FileExist(CharPath & rData & ".chr") Then
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Nemesis", 0)
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(UserIndex).Name)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas NEMESIS y prohibida la reenlistada" & _
                                                                FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & ".chr inexistente." & FONTTYPE_INFO)

            End If

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 10)) = "/CHAUCAOS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 10)
        Call LogGM(UserList(UserIndex).Name, "ECHO DEL CAOS A: " & rData)

        TIndex = NameIndex(rData)

        If TIndex > 0 Then
            UserList(TIndex).Faccion.FuerzasCaos = 0
            UserList(TIndex).Faccion.Reenlistadas = 200
            UserList(TIndex).Faccion.RecibioArmaduraCaos = 0

            Call PerderItemsFaccionarios(UserIndex, UserList(TIndex).Faccion.ArmaduraFaccionaria)

            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas del caos y prohibida la reenlistada" & _
                                                            FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, TIndex, 0, "|| " & UserList(UserIndex).Name & _
                                                       " te ha expulsado en forma definitiva de las fuerzas del caos." & FONTTYPE_FIGHT)
        Else

            If FileExist(CharPath & rData & ".chr") Then
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "EjercitoCaos", 0)
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(UserIndex).Name)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas del caos y prohibida la reenlistada" & _
                                                                FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & ".chr inexistente." & FONTTYPE_INFO)

            End If

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 10)) = "/CHAUREAL " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 10)
        Call LogGM(UserList(UserIndex).Name, "ECHO DE LA REAL A: " & rData)

        rData = Replace(rData, "\", "")
        rData = Replace(rData, "/", "")

        TIndex = NameIndex(rData)

        If TIndex > 0 Then
            UserList(TIndex).Faccion.ArmadaReal = 0
            UserList(TIndex).Faccion.Reenlistadas = 200
            UserList(TIndex).Faccion.RecibioArmaduraReal = 0

            Call PerderItemsFaccionarios(UserIndex, UserList(TIndex).Faccion.ArmaduraFaccionaria)

            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas reales y prohibida la reenlistada" & _
                                                            FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, TIndex, 0, "|| " & UserList(UserIndex).Name & _
                                                       " te ha expulsado en forma definitiva de las fuerzas reales." & FONTTYPE_FIGHT)
        Else

            If FileExist(CharPath & rData & ".chr") Then
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "EjercitoReal", 0)
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(UserIndex).Name)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas reales y prohibida la reenlistada" & _
                                                                FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & ".chr inexistente." & FONTTYPE_INFO)

            End If

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 11)) = "/FORCEMIDI " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 11)

        If Not IsNumeric(rData) Then
            Exit Sub
        Else
            Call SendData(SendTarget.toall, 0, 0, "|| " & UserList(UserIndex).Name & " broadcast musica: " & rData & FONTTYPE_SERVER)
            Call SendData(SendTarget.toall, 0, 0, "TM" & rData)

        End If

    End If

    If UCase$(Left$(rData, 10)) = "/FORCEWAV " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 10)

        If Not IsNumeric(rData) Then
            Exit Sub
        Else
            Call SendData(SendTarget.toall, 0, 0, "TW" & rData)

        End If

    End If

    If UCase$(Left$(rData, 12)) = "/BORRARPENA " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        '/borrarpena pj pena
        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        rData = Right$(rData, Len(rData) - 12)

        Name = ReadField(1, rData, Asc("@"))
        tStr = ReadField(2, rData, Asc("@"))

        If Name = "" Or tStr = "" Or Not IsNumeric(tStr) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Utilice /borrarpj Nick@NumeroDePena" & FONTTYPE_INFO)
            Exit Sub

        End If

        Name = Replace(Name, "\", "")
        Name = Replace(Name, "/", "")

        If FileExist(CharPath & Name & ".chr", vbNormal) Then
            rData = GetVar(CharPath & Name & ".chr", "PENAS", "P" & val(tStr))
            Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & val(tStr), LCase$(UserList(UserIndex).Name) & ": <Pena borrada> " & Date & " " & _
                                                                              Time)

        End If

        Call LogGM(UserList(UserIndex).Name, " borro la pena: " & tStr & "-" & rData & " de " & Name)
        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "/PJBAN " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 7)
        TIndex = NameIndex(ReadField(2, rData, Asc("@")))
        Name = ReadField(1, rData, Asc("@"))

        Arg1 = ReadField(2, rData, Asc("@"))
        Arg2 = CharPath & Left$(Arg1, 1) & "\" & Arg1 & ".chr"

        If TIndex <= 0 Then
            If PersonajeExiste(Arg1) Then
                Dim CANALBAN As Integer
                CANALBAN = FreeFile    ' obtenemos un canal
                Open App.Path & "\logs\BAN\" & GetVar(Arg2, "INIT", "LastSerie") & ".dat" For Append As #CANALBAN
                Print #CANALBAN, "PJ:" & Arg1 & " Fecha:" & Date & " GM:" & UserList(UserIndex).Name & " Razón:" & Name
                Close #CANALBAN
                Call SendData(toIndex, UserIndex, 0, "||Ban directo a la ficha de " & Arg1 & "." & "´" & FONTTYPE_INFO)
                Call WriteVar(CharPath & Left$(Arg1, 1) & "\" & Arg1 & ".chr", "FLAGS", "Ban", 1)

                Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Arg1, "BannedBy", UserList(UserIndex).Name)
                Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Arg1, "Reason", Name)
                Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Arg1, "Fecha", Date)

            Else
                Call SendData(toIndex, UserIndex, 0, "||Ese Pj no existe." & "´" & FONTTYPE_INFO)
            End If
            Exit Sub
        End If

        Exit Sub
    End If

    If UCase(Left(rData, 8)) = "/LASTIP " Then
        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right(rData, Len(rData) - 8)

        'No se si sea MUY necesario, pero por si las dudas... ;)
        rData = Replace(rData, "\", "")
        rData = Replace(rData, "/", "")

        If FileExist(CharPath & rData & ".chr", vbNormal) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La ultima IP de """ & rData & """ fue : " & GetVar(CharPath & rData & ".chr", _
                                                                                                                  "INIT", "LastIP") & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Charfile """ & rData & """ inexistente." & FONTTYPE_INFO)

        End If

        Exit Sub

    End If

    'Quita todos los NPCs del area
    If UCase$(rData) = "/LIMPIAR" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call LimpiarMundo
        Exit Sub

    End If

    'Crear criatura, toma directamente el indice
    If UCase$(Left$(rData, 5)) = "/ACC " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 5)
        Call SpawnNpc(val(rData), UserList(UserIndex).pos, True, False)
        Call LogGM(UserList(UserIndex).Name, " Sumoneo un " & Npclist(IndexNPC).Name & " en mapa " & UserList(UserIndex).pos.Map)
        Exit Sub

    End If

    'Crear criatura con respawn, toma directamente el indice
    If UCase$(Left$(rData, 6)) = "/RACC " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 6)
        Call SpawnNpc(val(rData), UserList(UserIndex).pos, True, True)
        Call LogGM(UserList(UserIndex).Name, " Sumoneo un " & Npclist(IndexNPC).Name & " en mapa " & UserList(UserIndex).pos.Map)
        Exit Sub

    End If

    'Comando para depurar la navegacion
    If UCase$(rData) = "/NAVE" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        If UserList(UserIndex).flags.Navegando = 1 Then
            UserList(UserIndex).flags.Navegando = 0
        Else
            UserList(UserIndex).flags.Navegando = 1

        End If

        Exit Sub

    End If

    If UCase$(rData) = "/SOLOGM" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        If ServerSoloGMs > 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Servidor Válido para todos" & FONTTYPE_INFO)
            ServerSoloGMs = 0
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Servidor Válido solo a administradores." & FONTTYPE_INFO)
            ServerSoloGMs = 1

        End If

        Exit Sub

    End If

    If UCase$(rData) = "/PISO" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        With UserList(UserIndex)

            For Y = 0 To 100
                For X = 0 To 100

                    If InMapBounds(.pos.Map, X, Y) Then
                        If MapData(.pos.Map, X, Y).OBJInfo.ObjIndex > 0 Then

                            Call SendData(toIndex, UserIndex, 0, "||(" & X & ", " & Y & ") " & ObjData(MapData(.pos.Map, X, Y).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)

                        End If

                    End If

                Next X
            Next Y

        End With

        Exit Sub
    End If

    If UCase$(Left$(rData, 7)) = "/CONDEN" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 8)
        TIndex = NameIndex(rData)

        If TIndex > 0 Then Call VolverCriminal(TIndex)
        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "/RAJAR " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 7)
        TIndex = NameIndex(UCase$(rData))

        If TIndex > 0 Then
            Call ResetFacciones(TIndex)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 11)) = "/RAJARCLAN " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 11)
        tInt = modGuilds.m_EcharMiembroDeClan(UserIndex, rData, False)  'me da el guildindex

        If tInt = 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| No pertenece a ningun clan o es fundador." & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Expulsado." & FONTTYPE_INFO)
            Call SendData(SendTarget.ToGuildMembers, tInt, 0, "|| " & rData & " ha sido expulsado del clan por los administradores del servidor" & _
                                                              FONTTYPE_GUILD)

        End If

        Exit Sub

    End If

    'altera email
    If UCase$(Left$(rData, 13)) = "/CAMBIARMAIL " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        rData = Right$(rData, Len(rData) - 13)
        tStr = ReadField(1, rData, Asc("-"))

        If tStr = "" Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||usar /CAMBIARMAIL <pj>-<nuevomail>" & FONTTYPE_GUILD)
            Exit Sub

        End If

        TIndex = NameIndex(tStr)

        If TIndex > 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario esta online, no se puede si esta online" & FONTTYPE_INFO)
            Exit Sub

        End If

        Arg1 = ReadField(2, rData, Asc("-"))

        If Arg1 = "" Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||usar /CAMBIARMAIL <pj>-<nuevomail>" & FONTTYPE_INFO)
            Exit Sub

        End If

        If Not FileExist(CharPath & tStr & ".chr") Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No existe el charfile " & CharPath & tStr & ".chr" & FONTTYPE_INFO)
        Else
            Call WriteVar(CharPath & tStr & ".chr", "CONTACTO", "Email", Arg1)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Email de " & tStr & " cambiado a: " & Arg1 & FONTTYPE_INFO)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 13)) = "/CAMBIARNICK " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        rData = Right$(rData, Len(rData) - 13)
        tStr = ReadField(1, rData, Asc("@"))
        Arg1 = ReadField(2, rData, Asc("@"))

        If tStr = "" Or Arg1 = "" Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usar: /CAMBIARNICK NiCK@NUEVO NICK" & FONTTYPE_INFO)
            Exit Sub

        End If

        TIndex = NameIndex(tStr)

        If TIndex > 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El Pj esta online, debe salir para el cambio" & FONTTYPE_WARNING)
            Exit Sub

        End If

        If FileExist(CharPath & UCase(tStr) & ".chr") = False Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El pj " & tStr & " es inexistente " & FONTTYPE_INFO)
            Exit Sub

        End If

        Arg2 = GetVar(CharPath & UCase(tStr) & ".chr", "GUILD", "GUILDINDEX")

        If IsNumeric(Arg2) Then
            If CInt(Arg2) > 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El pj " & tStr & _
                                                              " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido. " & FONTTYPE_INFO)
                Exit Sub

            End If

        End If

        If FileExist(CharPath & UCase(Arg1) & ".chr") = False Then
            FileCopy CharPath & UCase(tStr) & ".chr", CharPath & UCase(Arg1) & ".chr"
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Transferencia exitosa" & FONTTYPE_INFO)
            Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
            'ponemos la pena
            tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & ": BAN POR Cambio de nick a " & _
                                                                             UCase$(Arg1) & " " & Date & " " & Time)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El nick solicitado ya existe" & FONTTYPE_INFO)
            Exit Sub

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 11)) = "/GUARDAMAPA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        Call GrabarMapa(UserList(UserIndex).pos.Map, App.Path & "\WorldBackUp\Mapa" & UserList(UserIndex).pos.Map)
        Exit Sub

    End If

    If UCase$(Left$(rData, 20)) = "/TORNEOSAUTOMATICOS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 20)

        If (Torneo_Activo And Torneo_Esperando) Then
            Call SendData(toIndex, UserIndex, 0, "||Ya hay un torneo automatico en curso, si quieres cancelarla, usa /CANCELARTORNEO" & FONTTYPE_INFO)
            Exit Sub
        End If

        If rData > "6" Then
            Call SendData(toIndex, UserIndex, 0, "||Comando: /TORNEOSAUTOMATICOS <1-6>" & FONTTYPE_INFO)
        Else
            xao = 20
            RondaTorneo = rData
            Call SendData(SendTarget.toall, 0, 0, "||Esta empezando un nuevo torneo 1v1 de " & val(2 ^ RondaTorneo) & " participantes!! para participar pon /PARTICIPAR - (No cae inventario)" & FONTTYPE_GUILD)
            Call torneos_auto(RondaTorneo)
            Exit Sub
        End If

        Exit Sub
    End If

    If UCase$(Left$(rData, 14)) = "/CANCELATORNEO" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 14)

        If (Not Torneo_Activo And Not Torneo_Esperando) Then
            Call SendData(toIndex, UserIndex, 0, "||No hay un torneo automatico en curso!!" & FONTTYPE_INFO)
            Exit Sub
        End If

        Call Rondas_Cancela

        Exit Sub
    End If

    If UCase$(Left$(rData, 5)) = "/MAP " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        rData = Right(rData, Len(rData) - 5)

        Select Case UCase(ReadField(1, rData, 32))

        Case "PK"
            tStr = ReadField(2, rData, 32)

            If tStr <> "" Then
                MapInfo(UserList(UserIndex).pos.Map).Pk = IIf(tStr = "0", True, False)
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).pos.Map & ".dat", "Mapa" & UserList(UserIndex).pos.Map, "Pk", _
                              tStr)

            End If

            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Mapa " & UserList(UserIndex).pos.Map & " PK: " & MapInfo(UserList( _
                                                                                                                        UserIndex).pos.Map).Pk & FONTTYPE_INFO)

        Case "BACKUP"
            tStr = ReadField(2, rData, 32)

            If tStr <> "" Then
                MapInfo(UserList(UserIndex).pos.Map).BackUp = CByte(tStr)
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).pos.Map & ".dat", "Mapa" & UserList(UserIndex).pos.Map, _
                              "backup", tStr)

            End If

            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Mapa " & UserList(UserIndex).pos.Map & " Backup: " & MapInfo(UserList( _
                                                                                                                            UserIndex).pos.Map).BackUp & FONTTYPE_INFO)

        End Select

        Exit Sub

    End If

    If UCase$(Left$(rData, 11)) = "/BORRAR SOS" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        Call Ayuda.Reset
        Exit Sub

    End If

    If UCase$(Left$(rData, 9)) = "/SHOW INT" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        Call frmMain.mnuMostrar_Click
        Exit Sub

    End If


    If UCase$(rData) = "/ECHARTODOSPJSS" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call EcharPjsNoPrivilegiados
        Exit Sub

    End If

    If UCase$(rData) = "/RELOADNPCS" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        Call CargaNpcsDat

        Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Npcs.dat y npcsHostiles.dat recargados." & FONTTYPE_INFO)
        Exit Sub

    End If

End Sub
