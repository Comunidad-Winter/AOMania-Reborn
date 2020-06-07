Attribute VB_Name = "ModAreas"
'************************************************
'*   Sistema de areas refactorizado por WyroX   *
'*       Menos brujeria, mas comentarios.       *
'************************************************

'************************************************
'*                Como funciona?                *
'************************************************
' El mapa se divide en una cuadricula con secto-
' res de igual tamanio. El objetivo es que cada
' usuario reciba solo los paquetes de otros usua-
' rios, NPCs y objetos correspondientes a su area,
' es decir, su cuadricula y las 8 adyacentes.
' Los arrays de usuarios por cada mapa tienen un
' tamanio inicial predeterminado que se calcula
' en base a la cantidad de usuarios que frecuenta
' el mapa.
' El servidor solo se encarga de enviar los nuevos
' elementos del area y el cliente se encarga de
' eliminar los que ya no correspondan.

Option Explicit



'************************************************
'*            Valores modificables              *
'************************************************

' Tamanio del mapa
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

' Tamanio en tiles de la pantalla.
'ADVERTENCIA: TIENEN QUE SER IMPAR!
Public Const XWindow As Byte = 23
Public Const YWindow As Byte = 19

' Cantidad de tiles buffer
' (para que graficos grandes se vean desde fuera de la pantalla)
Private Const TileBufferSize As Byte = 5

' Intervalo en minutos para actualizar el tamanio
' optimo del array de usuarios de los mapa
Private Const TimeOptimization As Byte = 60



'************************************************
'*      Valores calculados automaticamente      *
'************************************************

' Rangos de vision
Public Const RANGO_VISION_X As Byte = XWindow \ 2
Public Const RANGO_VISION_Y As Byte = YWindow \ 2

' Tamanio de las areas
Public Const AREAS_X As Byte = RANGO_VISION_X + TileBufferSize
Public Const AREAS_Y As Byte = RANGO_VISION_Y + TileBufferSize



'************************************************
'*               Otras variables                *
'************************************************

Private AreasIO As clsIniManager

Private FILE_AREAS As String

Private Const USER_NUEVO As Byte = 255

Private CountOptimization As Byte

Public ConnGroups() As ConnGroup

Public Type AreaInfo
    AreaPerteneceX As Integer
    AreaPerteneceY As Integer
End Type

Public Type ConnGroup
    CountEntrys As Long
    OptValue As Long
    UserEntrys() As Long
End Type


'************************************************
'* GenerarAreas: leemos por cada mapa el tamanio*
'* optimo para el array de usuarios o por defec-*
'* to es uno.                                   *
'************************************************
Public Sub GenerarAreas()

    FILE_AREAS = DatPath & "\AreasStats.dat"

    'Si no existe {FILE_AREAS}, crea un archivo vacio con su nombre.
    If Not FileExist(FILE_AREAS, vbNormal) Then
        Open FILE_AREAS For Output As #1
        Close #1
    End If

    Set AreasIO = New clsIniManager
    Call AreasIO.Initialize(FILE_AREAS)

    ReDim ConnGroups(1 To NumMaps) As ConnGroup

    Dim i As Integer
    For i = 1 To NumMaps

        ' Leemos el valor guardado para el tamanio optimo
        ConnGroups(i).OptValue = val(AreasIO.GetValue("Areas", "Mapa" & i))

        ' Si es cero, lo dejamos en uno
        If ConnGroups(i).OptValue = 0 Then ConnGroups(i).OptValue = 1

        ' Redimensionamos
        ReDim ConnGroups(i).UserEntrys(1 To ConnGroups(i).OptValue) As Long

    Next i

    Set AreasIO = Nothing

End Sub


'*************************************************************************************************************
'* AreasOptimizacion: Funcion llamada cada {TimeOptimization} segundos cuyo objetivo es modificar el tamanio *
'                     optimo del array de usuarios segun la cantidad que haya en el mapa en ese momento.     *
'*************************************************************************************************************
Public Sub AreasOptimizacion()

'Objeto donde almacenamos la info. que vamos a escribir en el archivo {FILE_AREAS}
    Set AreasIO = New clsIniManager

    ' Solo modificamos los valores cada (TimeOptimization) segundos.
    CountOptimization = CountOptimization + 1

    If CountOptimization > TimeOptimization Then

        Dim i As Integer
        For i = 1 To NumMaps

            ' Modificamos el valor optimo haciendo un promedio entre el tamanio anterior y la cantidad de usuarios actual
            ConnGroups(i).OptValue = (ConnGroups(i).OptValue + ConnGroups(i).CountEntrys) \ 2

            ' Escribimos en el archivo de areas
            Call AreasIO.ChangeValue("Areas", "Mapa" & i, ConnGroups(i).OptValue)

            ' Si es cero, lo dejamos en uno
            If ConnGroups(i).OptValue = 0 Then ConnGroups(i).OptValue = 1

            ' Redimensionamos si es necesario
            If ConnGroups(i).OptValue > ConnGroups(i).CountEntrys Then
                ReDim Preserve ConnGroups(i).UserEntrys(1 To ConnGroups(i).OptValue) As Long
            End If

        Next i

        CountOptimization = 0

    End If

    'Escribimos la informacion en el objeto en el archivo {FILE_AREAS}
    Call AreasIO.DumpFile(FILE_AREAS)
    Set AreasIO = Nothing

End Sub


'*****************************************************************************************
'* AgregarUser: Agrega el usuario al mapa, enviando los datos correspondientes a su area *
'               y notificando al resto de usuarios.                                      *
'*****************************************************************************************
Public Sub AgregarUser(ByVal UserIndex As Integer, ByVal Map As Integer, Optional ByVal ButIndex As Boolean = False)

    If Not MapaValido(Map) Then Exit Sub

    Dim EsNuevo As Boolean
    EsNuevo = True

    ' Evitamos agregar usuarios repetidos
    Dim i As Integer
    For i = 1 To ConnGroups(Map).CountEntrys
        If ConnGroups(Map).UserEntrys(i) = UserIndex Then
            EsNuevo = False
            Exit For
        End If
    Next i

    ' Si es nuevo en el mapa
    If EsNuevo Then

        ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys + 1

        ' Aumentamos el tamanio del array de ser necesario
        If ConnGroups(Map).CountEntrys > UBound(ConnGroups(Map).UserEntrys) Then
            ReDim Preserve ConnGroups(Map).UserEntrys(1 To ConnGroups(Map).CountEntrys) As Long
        End If

        ConnGroups(Map).UserEntrys(ConnGroups(Map).CountEntrys) = UserIndex
    End If

    With UserList(UserIndex)
        .AreasInfo.AreaPerteneceX = -1
        .AreasInfo.AreaPerteneceY = -1
    End With

    Call CheckUpdateNeededUser(UserIndex, USER_NUEVO, ButIndex)

End Sub

'************************************************************************************
'* AgregarNpc: agrega el npc al mapa, notificando a los usuarios dentro de su area. *
'************************************************************************************
Public Sub AgregarNpc(ByVal NpcIndex As Integer)

    With Npclist(NpcIndex)
        .AreasInfo.AreaPerteneceX = -1
        .AreasInfo.AreaPerteneceY = -1
    End With

    Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)

End Sub

'*****************************************************************************
'* QuitarUser: remueve el usuario del array del mapa en el que se encuentra. *
'*****************************************************************************
Public Sub QuitarUser(ByVal UserIndex As Integer, ByVal Map As Integer)

' Buscamos el index dentro del array
    Dim LoopC As Long
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        If ConnGroups(Map).UserEntrys(LoopC) = UserIndex Then Exit For
    Next LoopC

    ' Si no existe salimos
    If LoopC > ConnGroups(Map).CountEntrys Then Exit Sub

    ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys - 1

    ' Corremos el array para llenar el hueco que dejo
    For LoopC = LoopC To ConnGroups(Map).CountEntrys - 1
        ConnGroups(Map).UserEntrys(LoopC) = ConnGroups(Map).UserEntrys(LoopC + 1)
    Next LoopC

    ' Reducimos el array
    If ConnGroups(Map).CountEntrys >= ConnGroups(Map).OptValue Then
        ReDim Preserve ConnGroups(Map).UserEntrys(1 To ConnGroups(Map).CountEntrys) As Long
    End If

End Sub

'***************************************************************************************************************
'* CheckUpdateNeededUser: Comprueba si es necesario modificar el area del usuario,                             *
'                         de ser asi, le envia todos los datos nuevos y avisa a los demas usuarios de la zona. *
'***************************************************************************************************************
Public Sub CheckUpdateNeededUser(ByVal UserIndex As Integer, ByVal heading As Byte, Optional ByVal ButIndex As Boolean = False, Optional verInvis As Byte = 0)

    With UserList(UserIndex)

        ' Comprobamos si cambio de area
        If .AreasInfo.AreaPerteneceX = .Pos.X \ AREAS_X And _
           .AreasInfo.AreaPerteneceY = .Pos.Y \ AREAS_Y Then _
           Exit Sub

        Dim MinX As Integer, MaxX As Integer, MinY As Integer, MaxY As Integer, X As Integer, Y As Integer, CurUser As Integer, CurObj As Integer, Map As Integer

        ' Calculamos segun la direccion del usuario el area nueva que tenemos que mandarle
        Call CalcularNuevaArea(.Pos.X, .Pos.Y, heading, MinX, MaxX, MinY, MaxY)

        ' Avisamos al cliente para que borre todo lo que esta fuera del area
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "CA" & .Pos.X & "," & .Pos.Y)

        Map = .Pos.Map

        For X = MinX To MaxX
            For Y = MinY To MaxY

                '<<< User >>>
                If MapData(Map, X, Y).UserIndex Then

                    CurUser = MapData(Map, X, Y).UserIndex

                    ' No nos enviamos a nosotros mismos...
                    If UserIndex <> CurUser Then

                        ' No vemos admins invisibles
                        If Not (UserList(CurUser).flags.AdminInvisible = 1) Then
                            Call MakeUserChar(SendTarget.ToIndex, UserIndex, 0, CurUser, Map, X, Y)

                            ' Enviamos la invisibilidad de ser necesario
                            If UserList(CurUser).flags.Invisible Or UserList(CurUser).flags.Oculto Then
                                If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
                                    Call EnviarDatosASlot(UserIndex, "NOVER" & UserList(CurUser).char.CharIndex & ",1," & UserList(CurUser).PartyIndex & ENDC)
                                End If
                            End If
                        End If

                        ' Si no somos un admin invisible
                        If Not (UserList(UserIndex).flags.AdminInvisible = 1) Then
                            ' Enviamos nuestro char al usuario
                            Call MakeUserChar(SendTarget.ToIndex, CurUser, 0, UserIndex, .Pos.Map, .Pos.X, .Pos.Y)

                            ' Enviamos la invisibilidad de ser necesario
                            If UserList(UserIndex).flags.Invisible Or UserList(UserIndex).flags.Oculto Then
                                If UserList(CurUser).flags.Privilegios And PlayerType.User Then
                                    Call EnviarDatosASlot(CurUser, "NOVER" & .char.CharIndex & ",1," & .PartyIndex & ENDC)
                                End If
                            End If
                        End If

                        '... excepto que nos hayamos warpeado al mapa
                    ElseIf heading = USER_NUEVO And Not ButIndex Then
                        Call MakeUserChar(SendTarget.ToIndex, UserIndex, 0, UserIndex, Map, X, Y)
               
                    End If
                  
                End If

                '<<< Npc >>>
                If MapData(Map, X, Y).NpcIndex Then
                    Call MakeNPCChar(SendTarget.ToIndex, UserIndex, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
                End If

                '<<< Item >>>
                If MapData(Map, X, Y).OBJInfo.ObjIndex Then
                    CurObj = MapData(Map, X, Y).OBJInfo.ObjIndex
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "HO" & ObjData(CurObj).GrhIndex & "," & X & "," & Y)

                    If ObjData(CurObj).ObjType = eOBJType.otPuertas Then
                        If CurObj = ObjData(CurObj).IndexAbierta Then

                            'Desbloquea
                            MapData(Map, X, Y).Blocked = 0
                            MapData(Map, X - 1, Y).Blocked = 0

                            Call Bloquear(SendTarget.ToIndex, UserIndex, 0, Map, X, Y, 0)
                            Call Bloquear(SendTarget.ToIndex, UserIndex, 0, Map, X - 1, Y, 0)

                        End If

                        If ObjData(CurObj).Cerrada = 1 Then
                            Call Bloquear(SendTarget.ToIndex, UserIndex, 0, Map, X, Y, MapData(Map, X, Y).Blocked)
                            Call Bloquear(SendTarget.ToIndex, UserIndex, 0, Map, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                        End If

                    End If

                End If

            Next Y
        Next X

        .AreasInfo.AreaPerteneceX = .Pos.X \ AREAS_X
        .AreasInfo.AreaPerteneceY = .Pos.Y \ AREAS_Y

    End With
End Sub


'***************************************************************************************************************
'* CheckUpdateNeededNpc: comprueba si el NPC cambio de area y le avisa a todos los usuarios que sea necesario. *
'***************************************************************************************************************
Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal heading As Byte)
    With Npclist(NpcIndex)

        ' Comprobamos si cambio de area
        If .AreasInfo.AreaPerteneceX = .Pos.X \ AREAS_X And _
           .AreasInfo.AreaPerteneceY = .Pos.Y \ AREAS_Y Then _
           Exit Sub

        Dim MinX As Integer, MaxX As Integer, MinY As Integer, MaxY As Integer, X As Integer, Y As Integer, UserIndex As Long

        ' Calculamos el area nueva segun la direccion del NPC
        Call CalcularNuevaArea(.Pos.X, .Pos.Y, heading, MinX, MaxX, MinY, MaxY)

        ' Si no hay usuarios en el mapa ahorramos tiempo y salimos
        If MapInfo(.Pos.Map).NumUsers <> 0 Then

            For X = MinX To MaxX
                For Y = MinY To MaxY
                    ' Si hay un usuario le enviamos el NPC
                    If MapData(.Pos.Map, X, Y).UserIndex Then _
                       Call MakeNPCChar(SendTarget.ToIndex, MapData(.Pos.Map, X, Y).UserIndex, 0, NpcIndex, .Pos.Map, .Pos.X, .Pos.Y)
                Next Y
            Next X

        End If

        .AreasInfo.AreaPerteneceX = .Pos.X \ AREAS_X
        .AreasInfo.AreaPerteneceY = .Pos.Y \ AREAS_Y

    End With
End Sub

'**************************************************************************************************************************
'* CalcularNuevaArea: segun la posicion actual y la direccion dada, se calcula el area en tiles que debe ser actualizada. *
'**************************************************************************************************************************
Private Sub CalcularNuevaArea(ByVal X As Integer, ByVal Y As Integer, ByVal heading As Byte, ByRef MinX As Integer, ByRef MaxX As Integer, ByRef MinY As Integer, ByRef MaxY As Integer)

    Dim AreaX As Integer, AreaY As Integer
    Dim MinAreaX As Integer, MaxAreaX As Integer, MinAreaY As Integer, MaxAreaY As Integer

    ' Calculamos el area actual al que pertenece
    AreaX = X \ AREAS_X
    AreaY = Y \ AREAS_Y

    ' Calculamos el conjunto de areas nuevas
    Select Case heading
    Case eHeading.NORTH
        ' 3 areas nuevas arriba
        MinAreaX = AreaX - 1
        MinAreaY = AreaY - 1
        MaxAreaX = AreaX + 1
        MaxAreaY = AreaY - 1

    Case eHeading.EAST
        ' 3 areas nuevas a la derecha
        MinAreaX = AreaX + 1
        MinAreaY = AreaY - 1
        MaxAreaX = AreaX + 1
        MaxAreaY = AreaY + 1

    Case eHeading.SOUTH
        ' 3 areas nuevas abajo
        MinAreaX = AreaX - 1
        MinAreaY = AreaY + 1
        MaxAreaX = AreaX + 1
        MaxAreaY = AreaY + 1

    Case eHeading.WEST
        ' 3 areas nuevas a la izquierda
        MinAreaX = AreaX - 1
        MinAreaY = AreaY - 1
        MaxAreaX = AreaX - 1
        MaxAreaY = AreaY + 1

    Case Else    ' Heading = USER_NUEVO (cambio de mapa, recien logueado, etc.)
        ' 9 areas nuevas alrededor del usuario (3x3)
        MinAreaX = AreaX - 1
        MinAreaY = AreaY - 1
        MaxAreaX = AreaX + 1
        MaxAreaY = AreaY + 1
    End Select

    ' Convertimos de areas a tiles
    MinX = MinAreaX * AREAS_X
    MinY = MinAreaY * AREAS_Y
    MaxX = (MaxAreaX + 1) * AREAS_X - 1
    MaxY = (MaxAreaY + 1) * AREAS_Y - 1

    ' Comprobamos que este dentro del mapa
    If MinX < XMinMapSize Then MinX = XMinMapSize
    If MinY < YMinMapSize Then MinY = YMinMapSize
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize

End Sub

'******************************************************************************************
'* EstanMismoArea: devuelve verdadero si los usuarios deben enviarse paquetes mutuamente. *
'******************************************************************************************
Public Function EstanMismoArea(ByVal UserA As Integer, ByVal UserB As Integer) As Boolean
    EstanMismoArea = Abs(UserList(UserA).AreasInfo.AreaPerteneceX - UserList(UserB).AreasInfo.AreaPerteneceX) <= 1 And _
                     Abs(UserList(UserA).AreasInfo.AreaPerteneceY - UserList(UserB).AreasInfo.AreaPerteneceY) <= 1
End Function

'*****************************************************************************************
'* EstanMismoAreaNPC: devuelve verdadero si el usuario esta en el mismo area que el NPC. *
'*****************************************************************************************
Public Function EstanMismoAreaNPC(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
    EstanMismoAreaNPC = Abs(UserList(UserIndex).AreasInfo.AreaPerteneceX - Npclist(NpcIndex).AreasInfo.AreaPerteneceX) <= 1 And _
                        Abs(UserList(UserIndex).AreasInfo.AreaPerteneceY - Npclist(NpcIndex).AreasInfo.AreaPerteneceY) <= 1
End Function

'**********************************************************************************************
'* EstanMismoAreaPos: devuelve verdadero si el usuario esta en el mismo area que la posicion. *
'**********************************************************************************************
Public Function EstanMismoAreaPos(ByVal UserIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    EstanMismoAreaPos = Abs(UserList(UserIndex).AreasInfo.AreaPerteneceX - X \ AREAS_X) <= 1 And _
                        Abs(UserList(UserIndex).AreasInfo.AreaPerteneceY - Y \ AREAS_Y) <= 1
End Function

Public Sub SendToUserArea(ByVal UserIndex As Integer, ByVal sdData As String)

'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim LoopC As Long
    Dim TempIndex As Integer

    Dim Map As Integer

    Map = UserList(UserIndex).Pos.Map
    If Not MapaValido(Map) Then Exit Sub

    sdData = sdData & ENDC

    For LoopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(LoopC)

        If EstanMismoArea(UserIndex, TempIndex) Then
            If UserList(TempIndex).ConnIDValida Then
                Call EnviarDatosASlot(TempIndex, sdData)
            End If
        End If

    Next LoopC

End Sub

Public Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal sdData As String)

'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
' ESTA SOLO SE USA PARA ENVIAR MPs asi que se puede encriptar desde aca :)
'**************************************************************
    Dim LoopC As Long
    Dim TempIndex As Integer

    Dim Map As Integer

    sdData = sdData & ENDC

    Map = UserList(UserIndex).Pos.Map
    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(LoopC)

        If TempIndex <> UserIndex Then
            If EstanMismoArea(UserIndex, TempIndex) Then
                If UserList(TempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(TempIndex, sdData)
                End If
            End If
        End If

    Next LoopC

End Sub

Public Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sdData As String)

'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim LoopC As Long
    Dim TempIndex As Integer

    Dim Map As Integer

    sdData = sdData & ENDC

    Map = Npclist(NpcIndex).Pos.Map
    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(LoopC)

        If EstanMismoAreaNPC(NpcIndex, TempIndex) Then
            If UserList(TempIndex).ConnIDValida Then
                Call EnviarDatosASlot(TempIndex, sdData)
            End If
        End If

    Next LoopC

End Sub

Public Sub SendToAreaByPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal sdData As String)

'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim LoopC As Long
    Dim TempIndex As Integer

    sdData = sdData & ENDC

    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(LoopC)

        If EstanMismoAreaPos(TempIndex, X, Y) Then
            If UserList(TempIndex).ConnIDValida Then
                Call EnviarDatosASlot(TempIndex, sdData)
            End If
        End If

    Next LoopC

End Sub
