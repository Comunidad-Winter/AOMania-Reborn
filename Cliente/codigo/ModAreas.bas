Attribute VB_Name = "ModAreas"

Option Explicit

'>>>>>>AREAS>>>>>AREAS>>>>>>>>AREAS>>>>>>>AREAS>>>>>>>>>>
Public Type AreaInfo

    AreaPerteneceX As Integer
    AreaPerteneceY As Integer
    
    AreaReciveX As Integer
    AreaReciveY As Integer
    
    MinX As Integer '-!!!
    MinY As Integer '-!!!
    
    AreaID As Long

End Type

Public Type ConnGroup

    CountEntrys As Long
    OptValue As Long
    UserEntrys() As Long

End Type

Public Const USER_NUEVO               As Byte = 255

'Cuidado:
' ¡¡¡LAS AREAS ESTÁN HARDCODEADAS!!!
Private CurDay                        As Byte
Private CurHour                       As Byte

Private AreasInfo(1 To 100, 1 To 100) As Byte
Private PosToArea(1 To 100)           As Byte

Private AreasRecive(12)               As Integer
'Private AreasEnvia(12) As Integer

Public ConnGroups()                   As ConnGroup

Public Sub InitAreas()

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim loopc   As Long
    Dim loopX   As Long
    Dim CurArea As Byte

    ' Setup areas...
    For loopc = 0 To 11
        AreasRecive(loopc) = (2 ^ loopc) Or IIf(loopc <> 0, 2 ^ (loopc - 1), 0) Or IIf(loopc <> 11, 2 ^ (loopc + 1), 0)
        '        AreasEnvia(LoopC) = 2 ^ (LoopC + 1)
    Next loopc
    
    For loopc = 1 To 100
        PosToArea(loopc) = loopc \ 9
    Next loopc
    
    For loopc = 1 To 100
        For loopX = 1 To 100
            'Usamos 121 IDs de area para saber si pasasamos de area "más rápido"
            AreasInfo(loopc, loopX) = (loopc \ 9 + 1) * (loopX \ 9 + 1)
        Next loopX
    Next loopc

    'Setup AutoOptimizacion de areas
    CurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
    CurHour = Fix(Hour(Time) \ 3) 'A ke parte de la hora pertenece
    
    ReDim ConnGroups(1 To NumMaps) As ConnGroup
    
    For loopc = 1 To NumMaps
        ConnGroups(loopc).OptValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & loopc, CurDay & "-" & CurHour))
        
        If ConnGroups(loopc).OptValue = 0 Then ConnGroups(loopc).OptValue = 1
        ReDim ConnGroups(loopc).UserEntrys(1 To ConnGroups(loopc).OptValue) As Long
    Next loopc

End Sub

Public Sub AreasOptimizacion()

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    'Es la función de autooptimizacion.... la idea es no mandar redimensionando arrays grandes todo el tiempo
    '**************************************************************
    Dim loopc      As Long
    Dim tCurDay    As Byte
    Dim tCurHour   As Byte
    Dim EntryValue As Long
    
    If (CurDay <> IIf(Weekday(Date) > 6, 1, 2)) Or (CurHour <> Fix(Hour(Time) \ 3)) Then
        
        tCurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
        tCurHour = Fix(Hour(Time) \ 3) 'A ke parte de la hora pertenece
        
        For loopc = 1 To NumMaps
            EntryValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & loopc, CurDay & "-" & CurHour))
            Call WriteVar(DatPath & "AreasStats.dat", "Mapa" & loopc, CurDay & "-" & CurHour, CInt((EntryValue + ConnGroups(loopc).OptValue) \ 2))
            
            ConnGroups(loopc).OptValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & loopc, tCurDay & "-" & tCurHour))

            If ConnGroups(loopc).OptValue = 0 Then ConnGroups(loopc).OptValue = 1
            If ConnGroups(loopc).OptValue >= MapInfo(loopc).NumUsers Then ReDim Preserve ConnGroups(loopc).UserEntrys(1 To ConnGroups( _
                    loopc).OptValue) As Long
        Next loopc
        
        CurDay = tCurDay
        CurHour = tCurHour

    End If

End Sub

Public Sub CheckUpdateNeededUser(ByVal UserIndex As Integer, ByVal Head As Byte)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    'Es la función clave del sistema de areas... Es llamada al mover un user
    '**************************************************************
    If UserList(UserIndex).AreasInfo.AreaID = AreasInfo(UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y) Then Exit Sub
    
    Dim MinX    As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long
    Dim TempInt As Long, Map As Long
    
    With UserList(UserIndex)
        
        MinX = .AreasInfo.MinX
        MinY = .AreasInfo.MinY
        
        If Head = eHeading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - 9
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.SOUTH Then
            MaxY = MinY + 35
            MinY = MinY + 27
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY - 18)
        
        ElseIf Head = eHeading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - 9
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.EAST Then
            MaxX = MinX + 35
            MinX = MinX + 27
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX - 18)
            .AreasInfo.MinY = CInt(MinY)
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.pos.Y \ 9) - 1) * 9
            MaxY = MinY + 26
            
            MinX = ((.pos.X \ 9) - 1) * 9
            MaxX = MinX + 26
            
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)

        End If
        
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > 100 Then MaxY = 100
        If MaxX > 100 Then MaxX = 100
        
        Map = UserList(UserIndex).pos.Map
        
        'Esto es para ke el cliente elimine lo "fuera de area..."
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "CA" & .pos.X & "," & .pos.Y)
        
        'Actualizamos!!!
        For X = MinX To MaxX
            For Y = MinY To MaxY
                
                '<<< User >>>
                If MapData(Map, X, Y).UserIndex Then
                    
                    TempInt = MapData(Map, X, Y).UserIndex
                    
                    If UserIndex <> TempInt Then
                        Call MakeUserChar(SendTarget.ToIndex, UserIndex, 0, CInt(TempInt), Map, X, Y)
                        Call MakeUserChar(SendTarget.ToIndex, CInt(TempInt), 0, UserIndex, .pos.Map, .pos.X, .pos.Y)
                        
                        'Si el user estaba invisible le avisamos al nuevo cliente de eso
       
                        If UserList(TempInt).flags.Invisible Or UserList(TempInt).flags.Oculto Then
                            Call EnviarDatosASlot(UserIndex, "NOVER" & UserList(TempInt).char.CharIndex & ",1," & UserList(TempInt).PartyIndex & ENDC)

                        End If
             
                        If .flags.Invisible Or .flags.Oculto Then
                            Call EnviarDatosASlot(TempInt, "NOVER" & .char.CharIndex & ",1," & .PartyIndex & ENDC)

                        End If
                       
                    ElseIf Head = USER_NUEVO Then
                        Call MakeUserChar(SendTarget.ToIndex, UserIndex, 0, UserIndex, Map, X, Y)

                    End If
                
                End If
                
                '<<< Npc >>>
                If MapData(Map, X, Y).NpcIndex Then
                    Call MakeNPCChar(SendTarget.ToIndex, UserIndex, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)

                End If
                 
                '<<< Item >>>
                If MapData(Map, X, Y).OBJInfo.ObjIndex Then
                    TempInt = MapData(Map, X, Y).OBJInfo.ObjIndex
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "HO" & ObjData(TempInt).GrhIndex & "," & X & "," & Y)
                    
                    If ObjData(TempInt).OBJType = eOBJType.otPuertas Then
                        If TempInt = ObjData(TempInt).IndexAbierta Then
                          
                            'Desbloquea
                            MapData(Map, X, Y).Blocked = 0
                            MapData(Map, X - 1, Y).Blocked = 0
                    
                            'Bloquea todos los mapas
                            Call Bloquear(SendTarget.ToIndex, UserIndex, 0, CInt(Map), X, Y, 0)
                            Call Bloquear(SendTarget.ToIndex, UserIndex, 0, CInt(Map), X - 1, Y, 0)

                        End If

                        If ObjData(TempInt).Cerrada = 1 Then
                            Call Bloquear(SendTarget.ToIndex, UserIndex, 0, CInt(Map), X, Y, MapData(Map, X, Y).Blocked)
                            Call Bloquear(SendTarget.ToIndex, UserIndex, 0, CInt(Map), X - 1, Y, MapData(Map, X - 1, Y).Blocked)

                        End If

                    End If

                End If
            
            Next Y
        Next X
            
        'Precalculados :P
        TempInt = .pos.X \ 9
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
        
        TempInt = .pos.Y \ 9
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
        .AreasInfo.AreaID = AreasInfo(.pos.X, .pos.Y)

    End With

End Sub

Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    ' Se llama cuando se mueve un Npc
    '**************************************************************
    
    If Npclist(NpcIndex).AreasInfo.AreaID = AreasInfo(Npclist(NpcIndex).pos.X, Npclist(NpcIndex).pos.Y) Then Exit Sub
    
    Dim MinX    As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long
    Dim TempInt As Long
    
    With Npclist(NpcIndex)
        MinX = .AreasInfo.MinX
        MinY = .AreasInfo.MinY
        
        If Head = eHeading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - 9
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.SOUTH Then
            MaxY = MinY + 35
            MinY = MinY + 27
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY - 18)
        
        ElseIf Head = eHeading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - 9
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.EAST Then
            MaxX = MinX + 35
            MinX = MinX + 27
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX - 18)
            .AreasInfo.MinY = CInt(MinY)
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.pos.Y \ 9) - 1) * 9
            MaxY = MinY + 26
            
            MinX = ((.pos.X \ 9) - 1) * 9
            MaxX = MinX + 26
            
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)

        End If
        
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > 100 Then MaxY = 100
        If MaxX > 100 Then MaxX = 100
        
        'Actualizamos!!!
        If MapInfo(.pos.Map).NumUsers <> 0 Then

            For X = MinX To MaxX
                For Y = MinY To MaxY

                    If MapData(.pos.Map, X, Y).UserIndex Then
                        Call MakeNPCChar(SendTarget.ToIndex, MapData(.pos.Map, X, Y).UserIndex, 0, NpcIndex, .pos.Map, .pos.X, .pos.Y)

                    End If

                Next Y
            Next X

        End If
            
        'Precalculados :P
        TempInt = .pos.X \ 9
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
            
        TempInt = .pos.Y \ 9
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
        .AreasInfo.AreaID = AreasInfo(.pos.X, .pos.Y)

    End With

End Sub

Public Sub QuitarUser(ByVal UserIndex As Integer, ByVal Map As Integer)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim TempVal As Long
    Dim loopc   As Long
    
    'Search for the user
    For loopc = 1 To ConnGroups(Map).CountEntrys

        If ConnGroups(Map).UserEntrys(loopc) = UserIndex Then Exit For
    Next loopc
    
    'Char not found
    If loopc > ConnGroups(Map).CountEntrys Then Exit Sub
    
    'Remove from old map
    ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys - 1
    TempVal = ConnGroups(Map).CountEntrys
    
    'Move list back
    For loopc = loopc To TempVal
        ConnGroups(Map).UserEntrys(loopc) = ConnGroups(Map).UserEntrys(loopc + 1)
    Next loopc
    
    If TempVal > ConnGroups(Map).OptValue Then 'Nescesito Redim?
        ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Long

    End If

End Sub

Public Sub AgregarUser(ByVal UserIndex As Integer, ByVal Map As Integer, Optional ByVal EsNuevo As Boolean = True)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim TempVal As Long
    
    If EsNuevo Then
        If Not MapaValido(Map) Then Exit Sub
        'Update map and connection groups data
        ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys + 1
        TempVal = ConnGroups(Map).CountEntrys
        
        If TempVal > ConnGroups(Map).OptValue Then 'Nescesito Redim
            ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Long

        End If
        
        ConnGroups(Map).UserEntrys(TempVal) = UserIndex

    End If

    'Update user
    UserList(UserIndex).AreasInfo.AreaID = 0
    
    UserList(UserIndex).AreasInfo.AreaPerteneceX = 0
    UserList(UserIndex).AreasInfo.AreaPerteneceY = 0
    UserList(UserIndex).AreasInfo.AreaReciveX = 0
    UserList(UserIndex).AreasInfo.AreaReciveY = 0

End Sub

Public Sub ArgegarNpc(ByVal NpcIndex As Integer)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    With Npclist(NpcIndex)
        .AreasInfo.AreaID = 0
        
        .AreasInfo.AreaPerteneceX = 0
        .AreasInfo.AreaPerteneceY = 0
        .AreasInfo.AreaReciveX = 0
        .AreasInfo.AreaReciveY = 0

    End With
    
End Sub

Public Sub SendToUserArea(ByVal UserIndex As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim loopc     As Long
    Dim TempIndex As Integer
    
    Dim Map       As Integer
    Dim AreaX     As Integer
    Dim AreaY     As Integer
    
    Map = UserList(UserIndex).pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

    If Not MapaValido(Map) Then Exit Sub
    
    sdData = sdData & ENDC
    
    For loopc = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(loopc)
        
        If UserList(TempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(TempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(TempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(TempIndex, sdData)

                End If

            End If

        End If

    Next loopc

End Sub

Public Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    ' ESTA SOLO SE USA PARA ENVIAR MPs asi que se puede encriptar desde aca :)
    '**************************************************************
    Dim loopc     As Long
    Dim TempInt   As Integer
    Dim TempIndex As Integer
    
    Dim Map       As Integer
    Dim AreaX     As Integer
    Dim AreaY     As Integer
    
    sdData = sdData & ENDC
    
    Map = UserList(UserIndex).pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

    If Not MapaValido(Map) Then Exit Sub

    For loopc = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(loopc)
            
        TempInt = UserList(TempIndex).AreasInfo.AreaReciveX And AreaX

        If TempInt Then  'Esta en el area?
            TempInt = UserList(TempIndex).AreasInfo.AreaReciveY And AreaY

            If TempInt Then
                If TempIndex <> UserIndex Then
                    If UserList(TempIndex).ConnIDValida Then
                        Call EnviarDatosASlot(TempIndex, sdData)

                    End If

                End If

            End If

        End If

    Next loopc

End Sub

Public Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sdData As String)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim loopc     As Long
    Dim TempInt   As Integer
    Dim TempIndex As Integer
    
    Dim Map       As Integer
    Dim AreaX     As Integer
    Dim AreaY     As Integer
    
    sdData = sdData & ENDC
    
    Map = Npclist(NpcIndex).pos.Map
    AreaX = Npclist(NpcIndex).AreasInfo.AreaPerteneceX
    AreaY = Npclist(NpcIndex).AreasInfo.AreaPerteneceY
        
    If Not MapaValido(Map) Then Exit Sub
    
    For loopc = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(loopc)
        
        TempInt = UserList(TempIndex).AreasInfo.AreaReciveX And AreaX

        If TempInt Then  'Esta en el area?
            TempInt = UserList(TempIndex).AreasInfo.AreaReciveY And AreaY

            If TempInt Then
                If UserList(TempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(TempIndex, sdData)

                End If

            End If

        End If

    Next loopc

End Sub

Public Sub SendToAreaByPos(ByVal Map As Integer, ByVal AreaX As Integer, ByVal AreaY As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim loopc     As Long
    Dim TempInt   As Integer
    Dim TempIndex As Integer
    
    AreaX = 2 ^ (AreaX \ 9)
    AreaY = 2 ^ (AreaY \ 9)
    
    sdData = sdData & ENDC
    
    If Not MapaValido(Map) Then Exit Sub

    For loopc = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(loopc)
            
        TempInt = UserList(TempIndex).AreasInfo.AreaReciveX And AreaX

        If TempInt Then  'Esta en el area?
            TempInt = UserList(TempIndex).AreasInfo.AreaReciveY And AreaY

            If TempInt Then
                If UserList(TempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(TempIndex, sdData)

                End If

            End If

        End If

    Next loopc

End Sub
