Attribute VB_Name = "Extra"

Option Explicit

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean

    EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie

End Function

Public Function esTemplario(ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 23/01/2007
    '***************************************************

    esTemplario = (UserList(UserIndex).Faccion.Templario = 1)

End Function

Public Function esNemesis(ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 23/01/2007
    '***************************************************

    esNemesis = (UserList(UserIndex).Faccion.Nemesis = 1)

End Function

Public Function esArmada(ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 23/01/2007
    '***************************************************

    esArmada = (UserList(UserIndex).Faccion.ArmadaReal = 1)

End Function

Public Function esCaos(ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 23/01/2007
    '***************************************************

    esCaos = (UserList(UserIndex).Faccion.FuerzasCaos = 1)

End Function

Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

    On Error GoTo errhandler
    
    If UserList(UserIndex).GranPoder = 1 And Not PermiteMapaPoder(UserIndex) Then
        Call mod_GranPoder.QuitarPoder(UserIndex)
    End If
    
    'Sonido de pajaritos
    
    If MapInfo(Map).Zona = "DUNGEON" Then
      
    Else
        Dim SoundPajaro As Integer
        Dim PorcPajaro  As Integer
      
        SoundPajaro = RandomNumber(21, 22)
        PorcPajaro = RandomNumber(1, 1000)
      
        If PorcPajaro < 5 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "PJ" & SoundPajaro)

        End If
      
    End If
        
    'Sonido al Pasar rayos casa encantada
    If UserList(UserIndex).pos.Map = MapaCasaAbandonada1 Then
      
        If UserList(UserIndex).pos.X = 51 And UserList(UserIndex).pos.Y = 75 Then
            Call SendData(SendTarget.toall, 0, 0, "TW108")

        End If
       
    End If
    
    
    If UserList(UserIndex).flags.angel Or UserList(UserIndex).flags.demonio Then
        If UserList(UserIndex).pos.Map = MapaBan Then
            If Not HayAgua(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y) Then
                If Not UserList(UserIndex).char.Body = 347 And UserList(UserIndex).flags.angel Then
                    Call Ban_ReloadTransforma(UserIndex)
                End If
                If Not UserList(UserIndex).char.Body = 348 And UserList(UserIndex).flags.demonio Then
                    Call Ban_ReloadTransforma(UserIndex)
                End If
            End If
        End If
    End If
    
    If UserList(UserIndex).flags.Corsarios = True Or UserList(UserIndex).flags.Piratas = True Then
        If UserList(UserIndex).pos.Map = MapaMedusa Then
            If HayAgua(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y) Then
                If UserList(UserIndex).flags.Muerto = 0 Then
                    Call Med_ReloadTransforma(UserIndex)
                End If
            ElseIf Not HayAgua(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y) Then
        
                Call Med_AguaDestransforma(UserIndex)
            End If
        End If
    End If
  
    If StatusInvo = True Then
        If MapData(mapainvo, mapainvoX1, mapainvoY1).UserIndex > 0 And MapData(mapainvo, mapainvoX2, mapainvoY2).UserIndex > 0 And MapData( _
            mapainvo, mapainvoX3, mapainvoY3).UserIndex > 0 And MapData(mapainvo, mapainvoX4, mapainvoY4).UserIndex > 0 And MapInfo( _
            mapainvo).criatinv = 0 Then

            Call SendData(SendTarget.toall, 0, 0, "||Se ha invocado una criatura en la Sala de Invocaciones." & FONTTYPE_TALK)
            Call SendData(SendTarget.ToMap, 0, "96", "TW160")
            MapInfo(mapainvo).criatinv = 1
            Dim criatura As Integer
            Dim invoca   As Integer
            criatura = 661
            invoca = criatura
            Call SpawnNpc(invoca, UserList(MapData(mapainvo, mapainvoX3, mapainvoY3).UserIndex).pos, True, False)

        End If

    End If

    Dim nPos   As WorldPos
    Dim FxFlag As Boolean

    If InMapBounds(Map, X, Y) Then
    
        If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
            FxFlag = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otTELEPORT

        End If
    
        If MapData(Map, X, Y).TileExit.Map > 0 Then
    
            'CHOTS | Solo Guerres y Kzas
            If MapData(Map, X, Y).TileExit.Map = 69 Then
                If UCase(UserList(UserIndex).Clase) = "MAGO" Or UCase(UserList(UserIndex).Clase) = "BARDO" Or UCase(UserList(UserIndex).Clase) = _
                    "ASESINO" Or UCase(UserList(UserIndex).Clase) = "CLERIGO" Or UCase(UserList(UserIndex).Clase) = "PALADIN" Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Este mapa es exclusivo para Guerreros y Cazadores." & FONTTYPE_INFO)
                    Call WarpUserChar(UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y + 1)
                    Exit Sub

                End If

            End If
    
            If MapData(Map, X, Y).TileExit.Map = 96 Then
                If Not UCase(UserList(UserIndex).Stats.ELV) >= 30 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Necesitas ser lvl 30 para poder ingresar a la sala de invocaciones!." & _
                        FONTTYPE_INFO)
                    Call WarpUserChar(UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y + 1)
                    Exit Sub

                End If

            End If

            If MapData(Map, X, Y).TileExit.Map = MapaCasaAbandonada1 Then

                If (UserList(UserIndex).Stats.GLD < 30000 Or UserList(UserIndex).Invent.ArmourEqpObjIndex = 0 Or EsNewbie(UserIndex)) Or UserList( _
                    UserIndex).Stats.ELV < 30 Then

                    Call SendData(SendTarget.toIndex, UserIndex, 0, _
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

                        If FxFlag Then '¿FX?
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

                Else 'No es newbie
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Mapa exclusivo para newbies." & FONTTYPE_INFO)
                    Dim veces As Byte
                    veces = 0
                    Call ClosestStablePos(UserList(UserIndex).pos, nPos)

                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)

                    End If

                End If

            Else 'No es un mapa de newbies

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

    Exit Sub

errhandler:
    Call LogError("Error en DotileEvents")

End Sub

Function InRangoVision(ByVal UserIndex As Integer, X As Integer, Y As Integer) As Boolean

    If X > UserList(UserIndex).pos.X - MinXBorder And X < UserList(UserIndex).pos.X + MinXBorder Then
        If Y > UserList(UserIndex).pos.Y - MinYBorder And Y < UserList(UserIndex).pos.Y + MinYBorder Then
            InRangoVision = True
            Exit Function

        End If

    End If

    InRangoVision = False

End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, X As Integer, Y As Integer) As Boolean

    If X > Npclist(NpcIndex).pos.X - MinXBorder And X < Npclist(NpcIndex).pos.X + MinXBorder Then
        If Y > Npclist(NpcIndex).pos.Y - MinYBorder And Y < Npclist(NpcIndex).pos.Y + MinYBorder Then
            InRangoVisionNPC = True
            Exit Function

        End If

    End If

    InRangoVisionNPC = False

End Function

Function InMapBounds(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        InMapBounds = False
    Else
        InMapBounds = True

    End If

End Function

Sub ClosestLegalPos(pos As WorldPos, ByRef nPos As WorldPos)
    '*****************************************************************
    'Encuentra la posicion legal mas cercana y la guarda en nPos
    '*****************************************************************

    Dim Notfound As Boolean
    Dim loopc    As Integer
    Dim Tx       As Long
    Dim Ty       As Long

    nPos.Map = pos.Map

    Do While Not LegalPos(pos.Map, nPos.X, nPos.Y)

        If loopc > 12 Then
            Notfound = True
            Exit Do

        End If
    
        For Ty = pos.Y - loopc To pos.Y + loopc
            For Tx = pos.X - loopc To pos.X + loopc
            
                If LegalPos(nPos.Map, Tx, Ty) Then
                    nPos.X = Tx
                    nPos.Y = Ty
                    '¿Hay objeto?
                
                    Tx = pos.X + loopc
                    Ty = pos.Y + loopc
  
                End If
        
            Next Tx
        Next Ty
    
        loopc = loopc + 1
    
    Loop

    If Notfound = True Then
        nPos.X = 0
        nPos.Y = 0

    End If

End Sub

Sub ClosestStablePos(pos As WorldPos, ByRef nPos As WorldPos)
    '*****************************************************************
    'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
    '*****************************************************************

    Dim Notfound As Boolean
    Dim loopc    As Integer
    Dim Tx       As Long
    Dim Ty       As Long

    nPos.Map = pos.Map

    Do While Not LegalPos(pos.Map, nPos.X, nPos.Y)

        If loopc > 12 Then
            Notfound = True
            Exit Do

        End If
    
        For Ty = pos.Y - loopc To pos.Y + loopc
            For Tx = pos.X - loopc To pos.X + loopc
            
                If LegalPos(nPos.Map, Tx, Ty) And MapData(nPos.Map, Tx, Ty).TileExit.Map = 0 Then
                    nPos.X = Tx
                    nPos.Y = Ty
                    '¿Hay objeto?
                
                    Tx = pos.X + loopc
                    Ty = pos.Y + loopc
  
                End If
        
            Next Tx
        Next Ty
    
        loopc = loopc + 1
    
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
    Dim loopc As Integer

    For loopc = 1 To MaxUsers

        If UserList(loopc).flags.UserLogged = True Then
            If UserList(loopc).ip = UserIP And UserIndex <> loopc Then
                CheckForSameIP = True
                Exit Function

            End If

        End If

    Next loopc

    CheckForSameIP = False

End Function

Function CheckForSameName(ByVal UserIndex As Integer, ByVal Name As String) As Boolean
    'Controlo que no existan usuarios con el mismo nombre
    Dim loopc As Long

    For loopc = 1 To MaxUsers

        If UserList(loopc).flags.UserLogged Then
        
            'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
            'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
            'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
            'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
            'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
        
            If UCase$(UserList(loopc).Name) = UCase$(Name) Then
                CheckForSameName = True
                Exit Function

            End If

        End If

    Next loopc

    CheckForSameName = False

End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef pos As WorldPos)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'Toma una posicion y se mueve hacia donde esta perfilado
    '*****************************************************************

    Select Case Head

        Case eHeading.NORTH
            pos.Y = pos.Y - 1
        
        Case eHeading.SOUTH
            pos.Y = pos.Y + 1
        
        Case eHeading.EAST
            pos.X = pos.X + 1
        
        Case eHeading.WEST
            pos.X = pos.X - 1

    End Select

End Sub

'Sub HeadtoPos(ByVal Head As eHeading, ByRef pos As WorldPos)
'    '*****************************************************************
'    'Toma una posicion y se mueve hacia donde esta perfilado
'    '*****************************************************************
'    Dim x       As Integer
'    Dim y       As Integer
'    Dim tempVar As Single
'    Dim nX      As Integer
'    Dim nY      As Integer
'
'    x = pos.x
'    y = pos.y
'
'    If Head = eHeading.NORTH Then
'        nX = x
'        nY = y - 1
'
'    End If
'
'    If Head = eHeading.SOUTH Then
'        nX = x
'        nY = y + 1
'
'    End If
'
'    If Head = eHeading.EAST Then
'        nX = x + 1
'        nY = y
'
'    End If
'
'    If Head = eHeading.WEST Then
'        nX = x - 1
'        nY = y
'
'    End If
'
'    'Devuelve valores
'    pos.x = nX
'    pos.y = nY
'
'End Sub

Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False) As Boolean

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

        End If
   
    End If

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

    Dim UserIndex        As Integer
    Dim IsDeadChar       As Boolean
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

Function LegalPosNPC(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte) As Boolean

    If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPosNPC = False
    Else

        If AguaValida = 0 Then
            LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And ( _
                MapData(Map, X, Y).trigger <> eTrigger.POSINVALIDA) And Not HayAgua(Map, X, Y)
        Else
            LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And ( _
                MapData(Map, X, Y).trigger <> eTrigger.POSINVALIDA)

        End If
 
    End If

End Function

Sub SendHelp(ByVal Index As Integer)
    Dim NumHelpLines As Integer
    Dim loopc        As Integer

    NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

    For loopc = 1 To NumHelpLines
        Call SendData(SendTarget.toIndex, Index, 0, "||" & GetVar(DatPath & "Help.dat", "Help", "Line" & loopc) & FONTTYPE_INFO)
    Next loopc

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

    If Npclist(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbWhite & "°" & Npclist(NpcIndex).Expresiones(randomi) & _
            "°" & Npclist(NpcIndex).char.CharIndex & FONTTYPE_INFO)

    End If

End Sub

Sub LookatTile_AutoAim(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

    Dim myX  As Integer, myY As Integer
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
    Dim FoundChar      As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex  As Integer
    Dim Stat           As String
    Dim OBJType        As Integer
    
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
                If ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                    .flags.TargetObjMap = Map
                    .flags.TargetObjX = X + 1
                    .flags.TargetObjY = Y
                    FoundSomething = 1

                End If

            ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex > 0 Then

                If ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                    'Informa el nombre
                    .flags.TargetObjMap = Map
                    .flags.TargetObjX = X + 1
                    .flags.TargetObjY = Y + 1
                    FoundSomething = 1

                End If

            ElseIf MapData(Map, X, Y + 1).OBJInfo.ObjIndex > 0 Then

                If ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
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
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & ObjData(.flags.TargetObj).Name & " - " & MapData(.flags.TargetObjMap, _
                        .flags.TargetObjX, .flags.TargetObjY).OBJInfo.Amount & vbNullString & FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & ObjData(.flags.TargetObj).Name & FONTTYPE_INFO)

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
            If FoundChar = 1 Then '  ¿Encontro un Usuario?
               
              
              
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
            
                    If Len(Stat) > 0 Then Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & Stat)

                    FoundSomething = 1
                    .flags.TargetUser = TempCharIndex
                    .flags.TargetNpc = 0
                    .flags.TargetNpcTipo = eNPCType.Comun

                End If

            End If

            If FoundChar = 2 Then '¿Encontro un NPC?
                Dim estatus As String
                Dim tNpc    As npc
                
                tNpc = Npclist(TempCharIndex)
                estatus = " (" & tNpc.Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") "

                If tNpc.Stats.MinHP < (tNpc.Stats.MaxHP * 0.05) Then
                    estatus = estatus & "(Agonizando)"
                ElseIf tNpc.Stats.MinHP < (tNpc.Stats.MaxHP * 0.1) Then
                    estatus = estatus & "(Casi muerto)"
                ElseIf tNpc.Stats.MinHP < (tNpc.Stats.MaxHP * 0.25) Then
                    estatus = estatus & "(Gravemente Herido)"
                ElseIf tNpc.Stats.MinHP < (tNpc.Stats.MaxHP * 0.5) Then
                    estatus = estatus & "(Herido)"
                ElseIf tNpc.Stats.MinHP < (tNpc.Stats.MaxHP * 0.75) Then
                    estatus = estatus & "(Levemente herido)"
                ElseIf tNpc.Stats.MinHP < (tNpc.Stats.MaxHP) Then
                    estatus = estatus & "(Sano)"
                Else
                    estatus = estatus & "(Intacto)"

                End If
            
                If Len(tNpc.Desc) > 1 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & tNpc.char.CharIndex _
                        & FONTTYPE_INFO)
                ElseIf TempCharIndex = CentinelaNPCIndex Then
                    'Enviamos nuevamente el texto del centinela según quien pregunta
                    Call modCentinela.CentinelaSendClave(UserIndex)
                Else

                    If tNpc.MaestroUser > 0 Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & tNpc.Name & " es mascota de " & UserList(tNpc.MaestroUser).Name & _
                            estatus & FONTTYPE_INFO)
                    Else
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & tNpc.Name & estatus & "." & FONTTYPE_INFO)

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
    LogError Err.Description & " in LookatTile " & "at line " & Erl

End Sub

'</EhFooter>

Function FindDirection(pos As WorldPos, Target As WorldPos) As eHeading
    '*****************************************************************
    'Devuelve la direccion en la cual el target se encuentra
    'desde pos, 0 si la direc es igual
    '*****************************************************************
    Dim X As Integer
    Dim Y As Integer

    X = pos.X - Target.X
    Y = pos.Y - Target.Y

    'NE
    If Sgn(X) = -1 And Sgn(Y) = 1 Then
        FindDirection = eHeading.NORTH
        Exit Function

    End If

    'NW
    If Sgn(X) = 1 And Sgn(Y) = 1 Then
        FindDirection = eHeading.WEST
        Exit Function

    End If

    'SW
    If Sgn(X) = 1 And Sgn(Y) = -1 Then
        FindDirection = eHeading.WEST
        Exit Function

    End If

    'SE
    If Sgn(X) = -1 And Sgn(Y) = -1 Then
        FindDirection = eHeading.SOUTH
        Exit Function

    End If

    'Sur
    If Sgn(X) = 0 And Sgn(Y) = -1 Then
        FindDirection = eHeading.SOUTH
        Exit Function

    End If

    'norte
    If Sgn(X) = 0 And Sgn(Y) = 1 Then
        FindDirection = eHeading.NORTH
        Exit Function

    End If

    'oeste
    If Sgn(X) = 1 And Sgn(Y) = 0 Then
        FindDirection = eHeading.WEST
        Exit Function

    End If

    'este
    If Sgn(X) = -1 And Sgn(Y) = 0 Then
        FindDirection = eHeading.EAST
        Exit Function

    End If

    'misma
    If Sgn(X) = 0 And Sgn(Y) = 0 Then
        FindDirection = 0
        Exit Function

    End If

End Function


'[Barrin 30-11-03]
Public Function ItemNoEsDeMapa(ByVal Index As Integer) As Boolean

    ItemNoEsDeMapa = ObjData(Index).OBJType <> eOBJType.otPuertas And ObjData(Index).OBJType <> eOBJType.otFOROS And ObjData(Index).OBJType <> _
        eOBJType.otCARTELES And ObjData(Index).OBJType <> eOBJType.otArboles And ObjData(Index).OBJType <> eOBJType.otYacimiento And ObjData( _
        Index).OBJType <> eOBJType.otTELEPORT

End Function

'[/Barrin 30-11-03]

Public Function MostrarCantidad(ByVal Index As Integer) As Boolean

    MostrarCantidad = ObjData(Index).OBJType <> eOBJType.otPuertas And ObjData(Index).OBJType <> eOBJType.otFOROS And ObjData(Index).OBJType <> _
        eOBJType.otCARTELES And ObjData(Index).OBJType <> eOBJType.otArboles And ObjData(Index).OBJType <> eOBJType.otYacimiento And ObjData( _
        Index).OBJType <> eOBJType.otTELEPORT

End Function

