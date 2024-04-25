Attribute VB_Name = "Mod_General"
Option Explicit

'Tamaño del mapa
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

Private Type tCabecera 'Cabecera de los con

    Desc As String * 255
    crc As Long
    MagicWord As Long

End Type

Private MiCabecera As tCabecera

Private Type Obj

    ObjIndex As Integer
    Amount As Integer

End Type

Private Type Position

    X As Integer
    Y As Integer

End Type

Private Type WorldPos

    Map As Integer
    X As Integer
    Y As Integer

End Type

Private Enum eTrigger

    NADA = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6

End Enum

'Tile
Public Type MapBlock

    Blocked As Byte
    Graphic(1 To 4) As Long
    UserIndex As Integer
    NpcIndex As Integer
    ObjInfo As Obj
    TileExit As WorldPos
    Trigger As eTrigger

End Type

Public MapData() As MapBlock
Public MapInfo() As Integer

Public Sub CargarMapaOLD(ByVal Map As Long, ByRef MAPFl As String)

    '***************************************************
    'Author: Unknown
    'Last Modification: 10/08/2010
    '10/08/2010 - Pato: Implemento el clsByteBuffer y el clsIniManager para la carga de mapa
    '***************************************************

    On Error GoTo errh

    Dim hFile     As Integer
    Dim X         As Long
    Dim Y         As Long
    Dim ByFlags   As Byte
    Dim MapReader As clsByteBuffer
    Dim InfReader As clsByteBuffer
    Dim Buff()    As Byte
    
    Set MapReader = New clsByteBuffer
    Set InfReader = New clsByteBuffer
    
    hFile = FreeFile

    Open MAPFl & ".map" For Binary As #hFile
    Seek hFile, 1

    ReDim Buff(LOF(hFile) - 1) As Byte
    
    Get #hFile, , Buff
    Close hFile
    
    Call MapReader.initializeReader(Buff)

    'inf
    Open MAPFl & ".inf" For Binary As #hFile
    Seek hFile, 1

    ReDim Buff(LOF(hFile) - 1) As Byte
    
    Get #hFile, , Buff
    Close hFile
    
    Call InfReader.initializeReader(Buff)
    
    'map Header
    MapInfo(Map) = MapReader.getInteger
    
    MiCabecera.Desc = MapReader.getString(Len(MiCabecera.Desc))
    MiCabecera.crc = MapReader.getLong
    MiCabecera.MagicWord = MapReader.getLong
    
    Call MapReader.getDouble

    'inf Header
    Call InfReader.getDouble
    Call InfReader.getInteger

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            With MapData(Map, X, Y)
                '.map file
                ByFlags = MapReader.getByte

                If ByFlags And 1 Then .Blocked = 1

                .Graphic(1) = CLng(MapReader.getInteger)

                'Layer 2 used?
                If ByFlags And 2 Then .Graphic(2) = MapReader.getInteger

                'Layer 3 used?
                If ByFlags And 4 Then .Graphic(3) = MapReader.getInteger

                'Layer 4 used?
                If ByFlags And 8 Then .Graphic(4) = MapReader.getInteger

                'Trigger used?
                If ByFlags And 16 Then .Trigger = MapReader.getInteger

                '.inf file
                ByFlags = InfReader.getByte

                If ByFlags And 1 Then
                    .TileExit.Map = InfReader.getInteger
                    .TileExit.X = InfReader.getInteger
                    .TileExit.Y = InfReader.getInteger

                End If
 
                If ByFlags And 2 Then 'Get and make NPC
                    .NpcIndex = InfReader.getInteger

                End If

                If ByFlags And 4 Then  'Get and make Object
                    .ObjInfo.ObjIndex = InfReader.getInteger
                    .ObjInfo.Amount = InfReader.getInteger

                End If

            End With

        Next X
    Next Y
    
    Set MapReader = Nothing
    Set InfReader = Nothing
    
    Erase Buff
    Exit Sub

errh:
    Call MsgBox("Error cargando mapa: " & Map & " - Pos: " & X & "," & Y & "." & Err.Description)

    Set MapReader = Nothing
    Set InfReader = Nothing

End Sub

Public Sub GrabarMapa(ByVal Map As Long, ByRef MAPFILE As String)

    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2011
    '10/08/2010 - Pato: Implemento el clsByteBuffer para el grabado de mapas
    '28/10/2010:ZaMa - Ahora no se hace backup de los pretorianos.
    '12/01/2011 - Amraphen: Ahora no se hace backup de NPCs prohibidos (Pretorianos, Mascotas, Invocados y Centinela)
    '***************************************************

    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y           As Long
    Dim X           As Long
    Dim ByFlags     As Byte
    Dim LoopC       As Long
    Dim MapWriter   As clsByteBuffer
    Dim InfWriter   As clsByteBuffer
    
    Set MapWriter = New clsByteBuffer
    Set InfWriter = New clsByteBuffer
    
    If FileExist(MAPFILE & ".map", vbNormal) Then
        Kill MAPFILE & ".map"

    End If
    
    If FileExist(MAPFILE & ".inf", vbNormal) Then
        Kill MAPFILE & ".inf"

    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open MAPFILE & ".Map" For Binary As FreeFileMap
    
    Call MapWriter.initializeWriter(FreeFileMap)
    
    'Open .inf file
    FreeFileInf = FreeFile
    Open MAPFILE & ".Inf" For Binary As FreeFileInf
    
    Call InfWriter.initializeWriter(FreeFileInf)
    
    'map Header
    Call MapWriter.putInteger(MapInfo(Map))
        
    Call MapWriter.putString(MiCabecera.Desc, False)
    Call MapWriter.putLong(MiCabecera.crc)
    Call MapWriter.putLong(MiCabecera.MagicWord)
    
    Call MapWriter.putDouble(0)
    
    'inf Header
    Call InfWriter.putDouble(0)
    Call InfWriter.putInteger(0)
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            With MapData(Map, X, Y)
                ByFlags = 0
                
                If .Blocked Then ByFlags = ByFlags Or 1
                If .Graphic(2) Then ByFlags = ByFlags Or 2
                If .Graphic(3) Then ByFlags = ByFlags Or 4
                If .Graphic(4) Then ByFlags = ByFlags Or 8
                If .Trigger Then ByFlags = ByFlags Or 16
                
                Call MapWriter.putByte(ByFlags)
                
                Call MapWriter.putLong(.Graphic(1))
                
                For LoopC = 2 To 4

                    If .Graphic(LoopC) Then Call MapWriter.putLong(.Graphic(LoopC))
                Next LoopC
                
                If .Trigger Then Call MapWriter.putInteger(CInt(.Trigger))
                
                '.inf file
                ByFlags = 0
       
                If .TileExit.Map Then ByFlags = ByFlags Or 1
                If .NpcIndex Then ByFlags = ByFlags Or 2
                If .ObjInfo.ObjIndex Then ByFlags = ByFlags Or 4
                
                Call InfWriter.putByte(ByFlags)
                
                If .TileExit.Map Then
                    Call InfWriter.putInteger(.TileExit.Map)
                    Call InfWriter.putInteger(.TileExit.X)
                    Call InfWriter.putInteger(.TileExit.Y)

                End If
                
                If .NpcIndex Then Call InfWriter.putInteger(.NpcIndex)
                
                If .ObjInfo.ObjIndex Then
                    Call InfWriter.putInteger(.ObjInfo.ObjIndex)
                    Call InfWriter.putInteger(.ObjInfo.Amount)

                End If

            End With

        Next X
    Next Y
    
    Call MapWriter.saveBuffer
    Call InfWriter.saveBuffer
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf
    
    Set MapWriter = Nothing
    Set InfWriter = Nothing
    
End Sub

Function FileExist(ByVal File As String, _
                   Optional FileType As VbFileAttribute = vbNormal) As Boolean
    '*****************************************************************
    'Se fija si existe el archivo
    '*****************************************************************

    FileExist = LenB(Dir$(File, FileType)) <> 0

End Function

