Attribute VB_Name = "Extra"

Option Explicit

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean

    EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie

End Function

Public Function esTemplario(ByVal UserIndex As Integer) As Boolean

    esTemplario = (UserList(UserIndex).Faccion.Templario = 1)

End Function

Public Function esNemesis(ByVal UserIndex As Integer) As Boolean

    esNemesis = (UserList(UserIndex).Faccion.Nemesis = 1)

End Function

Public Function esArmada(ByVal UserIndex As Integer) As Boolean

    esArmada = (UserList(UserIndex).Faccion.ArmadaReal = 1)

End Function

Public Function esCaos(ByVal UserIndex As Integer) As Boolean

    esCaos = (UserList(UserIndex).Faccion.FuerzasCaos = 1)

End Function

Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

    On Error GoTo errhandler
    
    Dim TimerTile As Integer
    
    TimerTile = RandomNumber(1, 30000)

    If UserList(UserIndex).GranPoder = 1 And Not PermiteMapaPoder(UserIndex) Then
        Call mod_GranPoder.QuitarPoder(UserIndex)
    End If

    'Sonido de pajaritos

    If MapInfo(Map).Zona <> Dungeon Then

        Dim SoundPajaro As Integer
        Dim PorcPajaro As Integer

        SoundPajaro = RandomNumber(21, 22)
        PorcPajaro = RandomNumber(1, 30000)

        If PorcPajaro > 29950 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "PJ" & SoundPajaro)
        End If

    End If

    'Sonido al Pasar rayos casa encantada
    If UserList(UserIndex).Pos.Map = MapaCasaAbandonada1 Then

        Dim SoundCasa As Integer
        Dim PorcCasa As Integer

        SoundCasa = RandomNumber(111, 113)
        PorcCasa = RandomNumber(1, 80)

        If UserList(UserIndex).Pos.X = 51 And UserList(UserIndex).Pos.Y = 75 Then
            Call SendData(SendTarget.ToMap, 0, MapaCasaAbandonada1, "TW108")
        End If

        If PorcCasa < 2 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "TW" & SoundCasa)
        End If

        PorcCasa = RandomNumber(1, 30000)

        If PorcCasa < 50 Then
            Call Efecto_CaminoCasaEncantada(UserIndex)
        End If

    End If


    If UserList(UserIndex).flags.Angel Or UserList(UserIndex).flags.Demonio Then
        If UserList(UserIndex).Pos.Map = MapaBan Then
            If Not HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then
                If Not UserList(UserIndex).char.Body = 347 And UserList(UserIndex).flags.Angel Then
                    Call Ban_ReloadTransforma(UserIndex)
                End If
                If Not UserList(UserIndex).char.Body = 348 And UserList(UserIndex).flags.Demonio Then
                    Call Ban_ReloadTransforma(UserIndex)
                End If
            End If
        End If
    End If

    If UserList(UserIndex).flags.Corsarios = True Or UserList(UserIndex).flags.Piratas = True Then
        If UserList(UserIndex).Pos.Map = MapaMedusa Then
            If HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then
                If UserList(UserIndex).flags.Muerto = 0 Then
                    Call Med_ReloadTransforma(UserIndex)
                End If
            ElseIf Not HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then

                Call Med_AguaDestransforma(UserIndex)
            End If
        End If
    End If

    If StatusInvo = True Then
        If MapData(mapainvo, mapainvoX1, mapainvoY1).UserIndex > 0 And MapData(mapainvo, mapainvoX2, mapainvoY2).UserIndex > 0 And MapData( _
           mapainvo, mapainvoX3, mapainvoY3).UserIndex > 0 And MapData(mapainvo, mapainvoX4, mapainvoY4).UserIndex > 0 And MapInfo( _
           mapainvo).criatinv = 0 Then

            Call SendData(SendTarget.ToAll, 0, 0, "||Se ha invocado una criatura en la Sala de Invocaciones." & FONTTYPE_TALK)
            Call SendData(SendTarget.ToMap, 0, "96", "TW160")
            MapInfo(mapainvo).criatinv = 1
            Dim criatura As Integer
            Dim invoca As Integer
            criatura = 661
            invoca = criatura
            Call SpawnNpc(invoca, UserList(MapData(mapainvo, mapainvoX3, mapainvoY3).UserIndex).Pos, True, False)

        End If

    End If

    Dim nPos As WorldPos
    Dim FxFlag As Boolean

    If InMapBounds(Map, X, Y) Then

        If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
            FxFlag = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ObjType = eOBJType.otTELEPORT

        End If

        If MapData(Map, X, Y).TileExit.Map > 0 Then

            'CHOTS | Solo Guerres y Kzas (Restricion telepeort)
            'If MapData(Map, X, Y).TileExit.Map = 69 Then
            '    If UCase(UserList(UserIndex).Clase) = "MAGO" Or UCase(UserList(UserIndex).Clase) = "BARDO" Or UCase(UserList(UserIndex).Clase) = _
                 '            "ASESINO" Or UCase(UserList(UserIndex).Clase) = "CLERIGO" Or UCase(UserList(UserIndex).Clase) = "PALADIN" Then
            '        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Este mapa es exclusivo para Guerreros y Cazadores." & FONTTYPE_INFO)
            '        Call WarpUserChar(UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y + 1)
            '        Exit Sub
            '    End If
            'End If

            If MapData(Map, X, Y).TileExit.Map = 96 Then
                If Not UCase(UserList(UserIndex).Stats.ELV) >= 30 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Necesitas ser lvl 30 para poder ingresar a la sala de invocaciones!." & _
                                                                    FONTTYPE_INFO)
                    Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1)
                    Exit Sub
                End If
            End If

            If MapData(Map, X, Y).TileExit.Map = 98 Or MapData(Map, X, Y).TileExit.Map = 99 Or MapData(Map, X, Y).TileExit.Map = 100 Or MapData(Map, X, Y).TileExit.Map = 101 Or MapData(Map, X, Y).TileExit.Map = 102 Then
                If UserList(UserIndex).NroMacotas > 0 Or UserList(UserIndex).flags.Montado = True Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No se permiten entrar al castillo con mascotas!!." & FONTTYPE_INFO)
                    Exit Sub
                End If
            End If

            If MapData(Map, X, Y).TileExit.Map = MapaCasaAbandonada1 Then
                If (UserList(UserIndex).Stats.GLD < 30000 Or UserList(UserIndex).Invent.ArmourEqpObjIndex = 0 Or EsNewbie(UserIndex)) Or UserList( _
                   UserIndex).Stats.ELV < 30 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, _
                                  "||Los espíritus no te dejan entrar si tienes menos de 30000 Monedas, eres Newbie, eres menor de level 30 o estás Desnudo." _
                                & FONTTYPE_INFO)
                    Exit Sub
                End If
            End If

            'CHOTS | Solo Guerres y Kzas

            '¿Es mapa de newbies?
            If UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).Restringir) = "SI" Then

                '¿El usuario es un newbie?
                If EsNewbie(UserIndex) Then
                    If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua( _
                                                                                                                               UserIndex)) Then

                        If FxFlag Then    '¿FX?
                            Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, _
                                                                                                                                 Y).TileExit.Y, True)
                        Else
                            Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, _
                                                                                                                                 Y).TileExit.Y)

                        End If

                    Else
                        Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)

                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            If FxFlag Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
                            Else
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)

                            End If

                        End If

                    End If

                Else    'No es newbie
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Mapa exclusivo para newbies." & FONTTYPE_INFO)
                    Dim veces As Byte
                    veces = 0
                    Call ClosestStablePos(UserList(UserIndex).Pos, nPos)

                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)

                    End If

                End If

            Else    'No es un mapa de newbies

                If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua( _
                                                                                                                           UserIndex)) Then

                    If FxFlag Then
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, _
                                          True)
                    Else
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)

                    End If

                Else
                    Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)

                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        If FxFlag Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
                        Else
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)

                        End If

                    End If

                End If

            End If

        End If

    End If
    
    If ((UserList(UserIndex).Pos.Map > 14 And UserList(UserIndex).Pos.Map < 19) Or (UserList(UserIndex).Pos.Map > 21 And UserList(UserIndex).Pos.Map < 25)) And UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.Privilegios < 1 Then
      If TimerTile > 29950 Then Call Gusano(UserIndex)
    End If

    Exit Sub

errhandler:
    Call LogError("Error en DotileEvents")

End Sub

Public Sub Gusano(ByVal UserIndex As Integer)

    Dim Daño As Long

    Dim lado As Integer

    Daño = RandomNumber(5, 20)
    lado = RandomNumber(62, 63)
    
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & 121)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & lado & "," & 0)
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - Daño
    Call SendData(ToIndex, UserIndex, 0, "||¡¡ Un Gusano te causa " & Daño & " de daño!!" & FONTTYPE_Motd4)
    
    Call SendUserStatsBox(UserIndex)

    If UserList(UserIndex).Stats.MinHP <= 0 Then Call UserDie(UserIndex)

End Sub

Function InRangoVision(ByVal UserIndex As Integer, X As Integer, Y As Integer) As Boolean

    If X > UserList(UserIndex).Pos.X - MinXBorder And X < UserList(UserIndex).Pos.X + MinXBorder Then
        If Y > UserList(UserIndex).Pos.Y - MinYBorder And Y < UserList(UserIndex).Pos.Y + MinYBorder Then
            InRangoVision = True
            Exit Function

        End If

    End If

    InRangoVision = False

End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, X As Integer, Y As Integer) As Boolean

    If X > Npclist(NpcIndex).Pos.X - MinXBorder And X < Npclist(NpcIndex).Pos.X + MinXBorder Then
        If Y > Npclist(NpcIndex).Pos.Y - MinYBorder And Y < Npclist(NpcIndex).Pos.Y + MinYBorder Then
            InRangoVisionNPC = True
            Exit Function

        End If

    End If

    InRangoVisionNPC = False

End Function

Function ObjetosBorrable(ByVal Obj As Integer)

    Dim tInt As String

    If Obj > 0 Then

        tInt = ObjData(Obj).ObjType

        If tInt <> otArboles And tInt <> otPuertas And tInt <> otCONTENEDORES And tInt <> otCARTELES And tInt <> otFOROS And tInt _
           <> otYacimiento And tInt <> otTELEPORT And tInt <> otYunque And tInt <> otFragua And tInt <> otMANCHAS And _
           tInt <> otOveja Then

            If Not ObjData(Obj).Gm = "1" And Not ObjData(Obj).sagrado = "1" And Not ObjData(Obj).Limpiar = "1" Then
                ObjetosBorrable = True
                Exit Function
            End If

        End If

    End If

    ObjetosBorrable = False

End Function

Function InMapBounds(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        InMapBounds = False
    Else
        InMapBounds = True

    End If

End Function

Sub ClosestLegalPosNpc(Pos As WorldPos, ByRef nPos As WorldPos, navegando As Boolean)
'*****************************************************************
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************
On Error GoTo err
Dim Notfound As Boolean
Dim LoopC As Integer
Dim tx As Integer
Dim ty As Integer
Dim error As Integer

nPos = Pos

error = 1
Do While True 'Not LegalPos(Pos.Map, nPos.X, nPos.Y)
    If LoopC > 100 Then
        Notfound = True
        Exit Do
    End If
    
    
    error = 2
    For ty = Pos.Y - LoopC To Pos.Y + LoopC
         For tx = Pos.X - LoopC To Pos.X + LoopC
            error = 3
            If LegalPosNPC(nPos.Map, tx, ty, navegando) And (MapData(nPos.Map, tx, ty).TileExit.Map = 0) Then
                nPos.X = tx
                nPos.Y = ty
                error = 4
                Exit Sub
                '¿Hay objeto?
                error = 5
'                tx = Pos.X + LoopC
'                ty = Pos.Y + LoopC

            End If

        Next tx
    Next ty

    LoopC = LoopC + 1
    
Loop
error = 6

If Notfound = True Then
    nPos.X = 0
    nPos.Y = 0
End If

Exit Sub

err:
End Sub

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

    Dim Notfound As Boolean
    Dim LoopC As Integer
    Dim tx As Long
    Dim ty As Long

    nPos.Map = Pos.Map

    Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y)

        If LoopC > 12 Then
            Notfound = True
            Exit Do

        End If

        For ty = Pos.Y - LoopC To Pos.Y + LoopC
            For tx = Pos.X - LoopC To Pos.X + LoopC

                If LegalPos(nPos.Map, tx, ty) Then
                    nPos.X = tx
                    nPos.Y = ty
                    '¿Hay objeto?

                    tx = Pos.X + LoopC
                    ty = Pos.Y + LoopC

                End If

            Next tx
        Next ty

        LoopC = LoopC + 1

    Loop

    If Notfound = True Then
        nPos.X = 0
        nPos.Y = 0

    End If

End Sub

Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
'*****************************************************************

    Dim Notfound As Boolean
    Dim LoopC As Integer
    Dim tx As Long
    Dim ty As Long

    nPos.Map = Pos.Map

    Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y)

        If LoopC > 12 Then
            Notfound = True
            Exit Do

        End If

        For ty = Pos.Y - LoopC To Pos.Y + LoopC
            For tx = Pos.X - LoopC To Pos.X + LoopC

                If LegalPos(nPos.Map, tx, ty) And MapData(nPos.Map, tx, ty).TileExit.Map = 0 Then
                    nPos.X = tx
                    nPos.Y = ty
                    '¿Hay objeto?

                    tx = Pos.X + LoopC
                    ty = Pos.Y + LoopC

                End If

            Next tx
        Next ty

        LoopC = LoopC + 1

    Loop

    If Notfound = True Then
        nPos.X = 0
        nPos.Y = 0

    End If

End Sub

Function NameIndex(ByVal Name As String) As Integer
    Dim UserIndex As Integer, i As Integer

    Name = Replace$(Name, "+", " ")

    If Len(Name) = 0 Then
        NameIndex = 0
        Exit Function
    End If

    UserIndex = 1

    If Right$(Name, 1) = "*" Then
        Name = Left$(Name, Len(Name) - 1)
        For i = 1 To LastUser
            If UCase$(UserList(i).Name) = UCase$(Name) Then
                NameIndex = i
                Exit Function
            End If
        Next
    Else
        For i = 1 To LastUser
            If UCase$(Left$(UserList(i).Name, Len(Name))) = UCase$(Name) Then
                NameIndex = i
                Exit Function
            End If
        Next
    End If

End Function
Function IP_Index(ByVal inIP As String) As Integer

    Dim UserIndex As Integer

    '¿Nombre valido?
    If inIP = "" Then
        IP_Index = 0
        Exit Function

    End If

    UserIndex = 1

    Do Until UserList(UserIndex).ip = inIP

        UserIndex = UserIndex + 1

        If UserIndex > MaxUsers Then
            IP_Index = 0
            Exit Function

        End If

    Loop

    IP_Index = UserIndex

    Exit Function

End Function

Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean
    Dim LoopC As Integer

    For LoopC = 1 To MaxUsers

        If UserList(LoopC).flags.UserLogged = True Then
            If UserList(LoopC).ip = UserIP And UserIndex <> LoopC Then
                CheckForSameIP = True
                Exit Function

            End If

        End If

    Next LoopC

    CheckForSameIP = False

End Function

Function CheckForSameName(ByVal UserIndex As Integer, ByVal Name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
    Dim LoopC As Long

    For LoopC = 1 To MaxUsers

        If UserList(LoopC).flags.UserLogged Then

            'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
            'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
            'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
            'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
            'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.

            If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
                CheckForSameName = True
                'Call SendData(SendTarget.ToIndex, LoopC, 0, _
                          "ERRSe ha conectado un usuario con el mismo nombre.")
                Call CloseSocket(LoopC)
                Exit Function

            End If

        End If

    Next LoopC

    CheckForSameName = False

End Function

'Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)
''***************************************************
''Author: Unknown
''Last Modification: -
''Toma una posicion y se mueve hacia donde esta perfilado
''*****************************************************************
''
'   Select Case Head
'
'    Case eHeading.NORTH
'        Pos.Y = Pos.Y - 1
''
'   Case eHeading.SOUTH
'       Pos.Y = Pos.Y + 1
'
'    Case eHeading.EAST
'        Pos.X = Pos.X + 1
''
'   Case eHeading.WEST
'       Pos.X = Pos.X - 1
'
'    End Select
'
'End Sub

Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)

    Dim X       As Integer

    Dim Y       As Integer

    Dim tempVar As Single

    Dim nX      As Integer

    Dim nY      As Integer

    X = Pos.X
    Y = Pos.Y

    If Head = NORTH Then
        nX = X
        nY = Y - 1

    End If

    If Head = SOUTH Then
        nX = X
        nY = Y + 1

    End If

    If Head = EAST Then
        nX = X + 1
        nY = Y

    End If

    If Head = WEST Then
        nX = X - 1
        nY = Y

    End If

    'Devuelve valores
    Pos.X = nX
    Pos.Y = nY

End Sub

Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal Elemental As Boolean = False) As Boolean

'¿Es un mapa valido?
    If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPos = False
    Else

        If Not PuedeAgua Then
            LegalPos = (MapData(Map, X, Y).Blocked <> 1) And (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And (Not _
                                                                                                                                           HayAgua(Map, X, Y))
        Else
            LegalPos = (MapData(Map, X, Y).Blocked <> 1) And (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And (HayAgua( _
                                                                                                                                           Map, X, Y))
           
           If Elemental Then
               
               LegalPos = (MapData(Map, X, Y).Blocked <> 1) And (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And (Not _
                    HayAgua(Map, X, Y))
                    
            End If
           
        End If

    End If

End Function

Function EsElemental(ByVal NpcIndex As Integer) As Boolean
          
          NpcIndex = Npclist(NpcIndex).Numero
          
          Select Case NpcIndex
          
                Case 89
                   EsElemental = True
                   Exit Function
                
                Case 92
                   EsElemental = True
                   Exit Function
                
                Case 93
                   EsElemental = True
                   Exit Function
                
                Case 94
                   EsElemental = True
                   Exit Function
                
                Case 166
                   EsElemental = True
                   Exit Function
                
                Case 242
                   EsElemental = True
                   Exit Function
                
                Case 618
                   EsElemental = True
                   Exit Function
                
                Case 619
                   EsElemental = True
                   Exit Function
                
                Case 620
                   EsElemental = True
                   Exit Function
                
                Case 693
                   EsElemental = True
                   Exit Function
                
                Case 721
                   EsElemental = True
                   Exit Function
                
          End Select
                
          
End Function


Function MoveToLegalPos(ByVal Map As Integer, _
                        ByVal X As Integer, _
                        ByVal Y As Integer, _
                        Optional ByVal PuedeAgua As Boolean = False, _
                        Optional ByVal PuedeTierra As Boolean = True) As Boolean
'***************************************************
'Autor: ZaMa
'Last Modification: 13/07/2009
'Checks if the position is Legal, but considers that if there's a casper, it's a legal movement.
'13/07/2009: ZaMa - Now it's also legal move where an invisible admin is.
'***************************************************

    Dim UserIndex As Integer
    Dim IsDeadChar As Boolean
    Dim IsAdminInvisible As Boolean

    '¿Es un mapa valido?
    If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        MoveToLegalPos = False
    Else

        With MapData(Map, X, Y)
            UserIndex = .UserIndex

            If UserIndex > 0 Then
                IsDeadChar = (UserList(UserIndex).flags.Muerto = 1)
                IsAdminInvisible = (UserList(UserIndex).flags.AdminInvisible = 1)
            Else
                IsDeadChar = False
                IsAdminInvisible = False

            End If

            If PuedeAgua And PuedeTierra Then
                MoveToLegalPos = (.Blocked <> 1) And (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0)
            ElseIf PuedeTierra And Not PuedeAgua Then
                MoveToLegalPos = (.Blocked <> 1) And (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (Not HayAgua(Map, X, _
                                                                                                                                            Y))
            ElseIf PuedeAgua And Not PuedeTierra Then
                MoveToLegalPos = (.Blocked <> 1) And (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (HayAgua(Map, X, Y))
            Else
                MoveToLegalPos = False

            End If

        End With

    End If

End Function

Function LegalPosNPC(ByVal Map As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal AguaValida As Byte) As Boolean

    If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPosNPC = False
    Else

        If AguaValida = 0 Then
            LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And (MapData(Map, X, Y).Trigger <> POSINVALIDA) And Not HayAgua(Map, X, Y)
        Else
            LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And (MapData(Map, X, Y).Trigger <> POSINVALIDA)

        End If
 
    End If

End Function

Sub SendHelp(ByVal Index As Integer)
    Dim NumHelpLines As Integer
    Dim LoopC As Integer

    NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

    For LoopC = 1 To NumHelpLines
        Call SendData(SendTarget.ToIndex, Index, 0, "||" & GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC) & FONTTYPE_INFO)
    Next LoopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

    If Npclist(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & Npclist(NpcIndex).Expresiones(randomi) & _
                                                                                   "°" & Npclist(NpcIndex).char.CharIndex & FONTTYPE_INFO)

    End If

End Sub

Sub LookatTile_AutoAim(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

    Dim myX As Integer, myY As Integer
    Dim Area As Integer

    Call LookatTile(UserIndex, Map, X, Y)

    If UserList(UserIndex).flags.TargetUser <> 0 Or UserList(UserIndex).flags.TargetNpc <> 0 Then Exit Sub

    For Area = 1 To 3
        For myX = (X - Area) To (X + Area)
            For myY = (Y - Area) To (Y + Area)
                Call LookatTile(UserIndex, Map, myX, myY)

                If (UserList(UserIndex).flags.TargetUser <> 0 Or UserList(UserIndex).flags.TargetNpc <> 0) And UserList(UserIndex).flags.TargetUser _
                   <> UserIndex Then Exit Sub

            Next myY
        Next myX
    Next Area

    Call LookatTile(UserIndex, Map, X, Y)

End Sub

Sub LookatTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

'<EhHeader>
    On Error GoTo LookatTile_Err

    '</EhHeader>

    'Responde al click del usuario sobre el mapa
    Dim FoundChar As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex As Integer
    Dim Stat As String
    Dim ObjType As Integer

    With UserList(UserIndex)

        '¿Posicion valida?
        If InMapBounds(Map, X, Y) Then
            .flags.TargetMap = Map
            .flags.TargetX = X
            .flags.TargetY = Y

            '¿Es un obj?
            If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
                'Informa el nombre
                .flags.TargetObjMap = Map
                .flags.TargetObjX = X
                .flags.TargetObjY = Y
                FoundSomething = 1
            ElseIf MapData(Map, X + 1, Y).OBJInfo.ObjIndex > 0 Then

                'Informa el nombre
                If ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).ObjType = eOBJType.otPuertas Then
                    .flags.TargetObjMap = Map
                    .flags.TargetObjX = X + 1
                    .flags.TargetObjY = Y
                    FoundSomething = 1

                End If

            ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex > 0 Then

                If ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).ObjType = eOBJType.otPuertas Then
                    'Informa el nombre
                    .flags.TargetObjMap = Map
                    .flags.TargetObjX = X + 1
                    .flags.TargetObjY = Y + 1
                    FoundSomething = 1

                End If

            ElseIf MapData(Map, X, Y + 1).OBJInfo.ObjIndex > 0 Then

                If ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).ObjType = eOBJType.otPuertas Then
                    'Informa el nombre
                    .flags.TargetObjMap = Map
                    .flags.TargetObjX = X
                    .flags.TargetObjY = Y + 1
                    FoundSomething = 1

                End If

            End If

            If FoundSomething = 1 Then
                .flags.TargetObj = MapData(Map, .flags.TargetObjX, .flags.TargetObjY).OBJInfo.ObjIndex

                If MostrarCantidad(.flags.TargetObj) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & ObjData(.flags.TargetObj).Name & " - " & MapData(.flags.TargetObjMap, _
                                                                                                                            .flags.TargetObjX, .flags.TargetObjY).OBJInfo.Amount & vbNullString & FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & ObjData(.flags.TargetObj).Name & FONTTYPE_INFO)

                End If

            End If

            '¿Es un personaje?
            If Y + 1 <= YMaxMapSize Then
                If MapData(Map, X, Y + 1).UserIndex > 0 Then
                    TempCharIndex = MapData(Map, X, Y + 1).UserIndex

                    If UserList(TempCharIndex).showName Then    ' Es GM y pidió que se oculte su nombre??
                        FoundChar = 1

                    End If

                End If

                If MapData(Map, X, Y + 1).NpcIndex > 0 Then
                    TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
                    FoundChar = 2

                End If

            End If

            '¿Es un personaje?
            If FoundChar = 0 Then
                If MapData(Map, X, Y).UserIndex > 0 Then
                    TempCharIndex = MapData(Map, X, Y).UserIndex

                    If UserList(TempCharIndex).showName Then    ' Es GM y pidió que se oculte su nombre??
                        FoundChar = 1

                    End If

                End If

                If MapData(Map, X, Y).NpcIndex > 0 Then
                    TempCharIndex = MapData(Map, X, Y).NpcIndex
                    FoundChar = 2

                End If

            End If

            'Reaccion al personaje
            If FoundChar = 1 Then    '  ¿Encontro un Usuario?



                If UserList(TempCharIndex).flags.AdminInvisible = 0 Or ((.flags.Privilegios = PlayerType.Dios) Or (.flags.Privilegios = _
                                                                                                                   PlayerType.Admin)) Then

                    If Len(UserList(TempCharIndex).DescRM) = 0 Then

                        If UserList(TempCharIndex).flags.Privilegios = PlayerType.User Then

                            If EsNewbie(TempCharIndex) Then
                                Stat = " <NEWBIE>"

                            End If

                            'Casado?
                            If UserList(TempCharIndex).flags.Casado = 1 Then

                                Select Case UCase$(UserList(TempCharIndex).Genero)

                                Case "HOMBRE"
                                    Stat = Stat & " [Esposo de " & UserList(TempCharIndex).Pareja & "]"

                                Case "MUJER"
                                    Stat = Stat & " [Esposa de " & UserList(TempCharIndex).Pareja & "]"

                                End Select

                            End If

                            If UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                                Stat = Stat & " <Armada del Credo> " & "<" & TituloReal(TempCharIndex) & ">"
                            ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                                Stat = Stat & " <Demonios de Abbadon> " & "<" & TituloCaos(TempCharIndex) & ">"
                            ElseIf UserList(TempCharIndex).Faccion.Nemesis = 1 Then
                                Stat = Stat & " <Caballeros de las Tinieblas> " & "<" & TituloNemesis(TempCharIndex) & ">"
                            ElseIf UserList(TempCharIndex).Faccion.Templario = 1 Then
                                Stat = Stat & " <Orden Templaria> " & "<" & TituloTemplario(TempCharIndex) & ">"
                            End If

                            If UserList(TempCharIndex).GuildIndex > 0 Then
                                If UserList(TempCharIndex).Clan.PuntosClan < 1000 Then
                                    Stat = Stat & " <" & Guilds(UserList(TempCharIndex).GuildIndex).GuildName & " (Soldado)" & ">"
                                ElseIf UserList(TempCharIndex).Clan.PuntosClan < 2000 Then
                                    Stat = Stat & " <" & Guilds(UserList(TempCharIndex).GuildIndex).GuildName & " (Teniente)" & ">"
                                ElseIf UserList(TempCharIndex).Clan.PuntosClan < 3000 Then
                                    Stat = Stat & " <" & Guilds(UserList(TempCharIndex).GuildIndex).GuildName & " (Cápitan)" & ">"
                                ElseIf UserList(TempCharIndex).Clan.PuntosClan < 4000 Then
                                    Stat = Stat & " <" & Guilds(UserList(TempCharIndex).GuildIndex).GuildName & " (General)" & ">"
                                ElseIf UserList(TempCharIndex).Clan.PuntosClan < 5000 Then
                                    Stat = Stat & " <" & Guilds(UserList(TempCharIndex).GuildIndex).GuildName & " (SubLider)" & ">"
                                ElseIf UserList(TempCharIndex).Clan.PuntosClan > 4999 Then
                                    Stat = Stat & " <" & Guilds(UserList(TempCharIndex).GuildIndex).GuildName & " (Lider)" & ">"
                                End If
                            End If

                            If Len(UserList(TempCharIndex).Desc) > 1 Then
                                Stat = "Ves a " & UserList(TempCharIndex).Name & Stat & " " & UserList(TempCharIndex).Desc
                            Else
                                Stat = "Ves a " & UserList(TempCharIndex).Name & Stat

                            End If

                            If UserList(TempCharIndex).flags.PertAlCons > 0 Then
                                Stat = Stat & " [Consejo de la Luz]" & FONTTYPE_CONSEJOVesA
                            ElseIf UserList(TempCharIndex).flags.PertAlConsCaos > 0 Then
                                Stat = Stat & " [Consejo de las Sombras]" & FONTTYPE_CONSEJOCAOSVesA
                            Else

                                If Criminal(TempCharIndex) Then
                                    Stat = Stat & " <CRIMINAL>"
                                Else
                                    Stat = Stat & " <CIUDADANO>"
                                End If

                                If UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                                    Stat = Stat & "~0~0~200~1~0"
                                ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                                    Stat = Stat & "~255~0~0~1~0"
                                ElseIf UserList(TempCharIndex).Faccion.Nemesis = 1 Then
                                    Stat = Stat & "~102~102~102~1~0"
                                ElseIf UserList(TempCharIndex).Faccion.Templario = 1 Then
                                    Stat = Stat & "~255~255~255~1~0"
                                End If

                                If UserList(TempCharIndex).Faccion.ArmadaReal = 0 And UserList(TempCharIndex).Faccion.FuerzasCaos = 0 And _
                                   UserList(TempCharIndex).Faccion.Nemesis = 0 And UserList(TempCharIndex).Faccion.Templario = 0 Then

                                    If Criminal(TempCharIndex) Then
                                        Stat = Stat & "~255~0~0~1~0"
                                    Else
                                        Stat = Stat & "~0~0~200~1~0"
                                    End If

                                End If


                            End If

                        Else

                            Stat = Stat & "Ves a " & UserList(TempCharIndex).Name & " - <Game Master> ~255~128~64~1~0"
                        End If

                    Else
                        Stat = UserList(TempCharIndex).DescRM & " " & FONTTYPE_INFOBOLD

                    End If

                    If Len(Stat) > 0 Then Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Stat)

                    FoundSomething = 1
                    .flags.TargetUser = TempCharIndex
                    .flags.TargetNpc = 0
                    .flags.TargetNpcTipo = eNPCType.Comun

                End If

            End If

            If FoundChar = 2 Then    '¿Encontro un NPC?
                Dim estatus As String
                Dim tNpc As npc

                tNpc = Npclist(TempCharIndex)

                If tNpc.MaestroUser = 0 Then
                    If tNpc.Stats.MinHP < (tNpc.Stats.MaxHP * 0.1) Then
                        estatus = estatus & " (Agonizando)"
                    ElseIf tNpc.Stats.MinHP < (tNpc.Stats.MaxHP * 0.4) Then
                        estatus = estatus & " (Gravemente Herido)"
                    ElseIf tNpc.Stats.MinHP < (tNpc.Stats.MaxHP * 0.65) Then
                        estatus = estatus & " (Bastante herido)"
                    ElseIf tNpc.Stats.MinHP < (tNpc.Stats.MaxHP * 0.9) Then
                        estatus = estatus & " (Apenas lastimado)"
                    Else
                        estatus = estatus & " (Totalmente sano)"
                    End If
                End If

                If Len(tNpc.Desc) > 1 And UserIndex <> Centinela.RevisandoUserIndex Then
                
                    If .Quest.Start = 1 Then
                        
                        If .Quest.ValidNpcDD = 1 Then
                            Call CambiaDescQuest(UserIndex, .Quest.Quest, TempCharIndex)
                        ElseIf .Quest.ValidNpcDescubre = 1 Then
                            Call CambiaDescQuest(UserIndex, .Quest.Quest, TempCharIndex)
                        ElseIf .Quest.NumObjNpc > 0 Then
                            Call CambiaDescQuest(UserIndex, .Quest.Quest, TempCharIndex)
                        ElseIf .Quest.ValidHablarNpc > 0 Then
                             Call CambiaDescQuest(UserIndex, .Quest.Quest, TempCharIndex)
                          Else
                          Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & tNpc.char.CharIndex _
                                                                  & FONTTYPE_INFO)
                         End If
                    
                    Else
                    
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & tNpc.char.CharIndex _
                                                                  & FONTTYPE_INFO)
                                                                  
                     End If
                                                                  
                ElseIf Len(tNpc.Desc) > 1 And UserIndex = Centinela.RevisandoUserIndex Then
                    'Enviamos nuevamente el texto del centinela según quien pregunta
                    Call modCentinela.CentinelaSendClave(UserIndex)
                Else

                    If tNpc.MaestroUser > 0 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| " & tNpc.Name & " es mascota de " & UserList(tNpc.MaestroUser).Name & _
                                                                        estatus & "." & FONTTYPE_INFO)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| " & tNpc.Name & estatus & "." & FONTTYPE_INFO)

                    End If

                End If

                FoundSomething = 1
                .flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
                .flags.TargetNpc = TempCharIndex
                .flags.TargetUser = 0
                .flags.TargetObj = 0

            End If

            If FoundChar = 0 Then
                .flags.TargetNpc = 0
                .flags.TargetNpcTipo = eNPCType.Comun
                .flags.TargetUser = 0

            End If

            '*** NO ENCOTRO NADA ***
            If FoundSomething = 0 Then
                .flags.TargetNpc = 0
                .flags.TargetNpcTipo = eNPCType.Comun
                .flags.TargetUser = 0
                .flags.TargetObj = 0
                .flags.TargetObjMap = 0
                .flags.TargetObjX = 0
                .flags.TargetObjY = 0

            End If

        Else

            If FoundSomething = 0 Then
                .flags.TargetNpc = 0
                .flags.TargetNpcTipo = eNPCType.Comun
                .flags.TargetUser = 0
                .flags.TargetObj = 0
                .flags.TargetObjMap = 0
                .flags.TargetObjX = 0
                .flags.TargetObjY = 0

            End If

        End If

    End With

    Exit Sub

LookatTile_Err:
    LogError err.Description & " in LookatTile " & "at line " & Erl

End Sub

'</EhFooter>

Function FindDirection(Pos As WorldPos, Target As WorldPos) As Byte

Dim X As Integer
Dim Y As Integer

X = Pos.X - Target.X
Y = Pos.Y - Target.Y

If Sgn(X) = -1 And Sgn(Y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If


If Sgn(X) = 1 And Sgn(Y) = 1 Then
    FindDirection = WEST
    Exit Function
End If

If Sgn(X) = 1 And Sgn(Y) = -1 Then
    FindDirection = WEST
    Exit Function
End If

If Sgn(X) = -1 And Sgn(Y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

If Sgn(X) = 0 And Sgn(Y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

If Sgn(X) = 0 And Sgn(Y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

If Sgn(X) = 1 And Sgn(Y) = 0 Then
    FindDirection = WEST
    Exit Function
End If

If Sgn(X) = -1 And Sgn(Y) = 0 Then
    FindDirection = EAST
    Exit Function
End If

If Sgn(X) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If

End Function

Function FindDirectionEAO(a As WorldPos, b As WorldPos, Optional PuedeAgu As Boolean) As Byte

    Dim r As Byte

    'Mejoras:
    'Ahora los NPC puden doblar las esquinas, y pasar a los lados de los arboles, _
     Tambien cuando te persiguen en ves de ir en forma orizontal y despues en vertical, te van sigsagueando.

    'A = NPCPOS
    'B = UserPos

    'Esto es para que el NPC retroceda en caso de no poder seguir adelante, en ese caso se retrocede.

    'Lo que no pueden hacer los Npcs, es rodear cosas, EJ:

    '
    ' *******************
    ' *                 *
    ' *                 *
    ' * B              *
    ' ******     ********
    '   A  <------- El npc se va a quedar loco tratando de pasar de frente en ves de rodearlo.
    '
    '
    'Saluda: <-.Siameze.->


    Dim PV As Integer

    r = RandomNumber(1, 2)

    If a.X > b.X And a.Y > b.Y Then
        FindDirectionEAO = IIf(r = 1, NORTH, WEST)

    ElseIf a.X < b.X And a.Y < b.Y Then
        FindDirectionEAO = IIf(r = 1, SOUTH, EAST)

    ElseIf a.X < b.X And a.Y > b.Y Then
        FindDirectionEAO = IIf(r = 1, NORTH, EAST)

    ElseIf a.X > b.X And a.Y < b.Y Then
        FindDirectionEAO = IIf(r = 1, SOUTH, WEST)

    ElseIf a.X = b.X Then
        FindDirectionEAO = IIf(a.Y < b.Y, SOUTH, NORTH)

    ElseIf a.Y = b.Y Then
        FindDirectionEAO = IIf(a.X < b.X, EAST, WEST)

    Else

        FindDirectionEAO = 0    ' this is imposible!

    End If

    If Distancia(a, b) > 1 Then

        Select Case FindDirectionEAO


        Case NORTH
            If Not LegalPos(a.Map, a.X, a.Y - 1, PuedeAgu) Then

                If a.X > b.X Then
                    FindDirectionEAO = WEST
                ElseIf a.X < b.X Then
                    FindDirectionEAO = EAST
                Else
                    FindDirectionEAO = IIf(r > 1, WEST, EAST)
                End If
                PV = 1

            End If


        Case SOUTH
            If Not LegalPos(a.Map, a.X, a.Y + 1, PuedeAgu) Then

                If a.X > b.X Then
                    FindDirectionEAO = WEST
                ElseIf a.X < b.X Then
                    FindDirectionEAO = EAST
                Else
                    FindDirectionEAO = IIf(r > 1, WEST, EAST)
                End If
                PV = 1

            End If



        Case WEST
            If Not LegalPos(a.Map, a.X - 1, a.Y, PuedeAgu) Then

                If a.Y > b.Y Then
                    FindDirectionEAO = NORTH
                ElseIf a.Y < b.Y Then
                    FindDirectionEAO = SOUTH
                Else
                    FindDirectionEAO = IIf(r > 1, NORTH, SOUTH)
                End If
                PV = 1
            End If

        Case EAST
            If Not LegalPos(a.Map, a.X + 1, a.Y, PuedeAgu) Then
                If a.Y > b.Y Then
                    FindDirectionEAO = NORTH
                ElseIf a.Y < b.Y Then
                    FindDirectionEAO = SOUTH
                Else
                    FindDirectionEAO = IIf(r > 1, NORTH, SOUTH)
                End If
                PV = 1

            End If

        End Select

        If PV = 1 Then

            Select Case FindDirectionEAO
            Case EAST
                If Not LegalPos(a.Map, a.X + 1, a.Y) Then FindDirectionEAO = WEST

            Case WEST
                If Not LegalPos(a.Map, a.X - 1, a.Y) Then FindDirectionEAO = EAST

            Case NORTH
                If Not LegalPos(a.Map, a.X, a.Y - 1) Then FindDirectionEAO = SOUTH

            Case SOUTH
                If Not LegalPos(a.Map, a.X, a.Y + 1) Then FindDirectionEAO = NORTH

            End Select


        End If


    End If

End Function

'[Barrin 30-11-03]
Public Function ItemNoEsDeMapa(ByVal Index As Integer) As Boolean

    ItemNoEsDeMapa = ObjData(Index).ObjType <> eOBJType.otPuertas And ObjData(Index).ObjType <> eOBJType.otFOROS And ObjData(Index).ObjType <> eOBJType.otCARTELES And ObjData(Index).ObjType <> eOBJType.otArboles And ObjData(Index).ObjType <> eOBJType.otYacimiento And ObjData(Index).ObjType <> eOBJType.otTELEPORT

End Function

'[/Barrin 30-11-03]

Public Function MostrarCantidad(ByVal Index As Integer) As Boolean

    MostrarCantidad = ObjData(Index).ObjType <> eOBJType.otPuertas And ObjData(Index).ObjType <> eOBJType.otFOROS And ObjData(Index).ObjType <> _
                      eOBJType.otCARTELES And ObjData(Index).ObjType <> eOBJType.otArboles And ObjData(Index).ObjType <> eOBJType.otYacimiento And ObjData( _
                      Index).ObjType <> eOBJType.otTELEPORT

End Function
