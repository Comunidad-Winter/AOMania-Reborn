Attribute VB_Name = "Mod_TileEngine"
Option Explicit

Private grhCount            As Long
Private Velocidad           As Single

Public Const PI             As Single = 3.14159
Public Const DegreeToRadian As Single = PI / 180
Public Const RadianToDegree As Single = 180 / PI

'Major DX Objects
Public DirectX              As DirectX8
Public DirectD3D            As Direct3D8
Public DirectDevice         As Direct3DDevice8
Public DirectD3D8           As D3DX8
Public base_light           As Long

Private DirectD3Dpp         As D3DPRESENT_PARAMETERS
Private DirectD3Dcaps       As D3DCAPS8

Private Projection          As D3DMATRIX
Private View                As D3DMATRIX

Private MainViewRect        As D3DRECT
Private ConnectRect         As D3DRECT

Private Const FVF           As Long = D3DFVF_XYZ Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE

Private Type D3DXIMAGE_INFO_A

    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long

End Type
        
Private Type tStructureLng

    X As Long
    Y As Long

End Type
        
Private Type CharVA

    X As Integer
    Y As Integer
    w As Integer
    H As Integer
    
    Tx1 As Single
    Tx2 As Single
    Ty1 As Single
    Ty2 As Single

End Type

Private Type VFH

    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte 'The actual factual width of each character
    CharVA(0 To 255) As CharVA

End Type

Private Type CustomFont

    HeaderInfo As VFH            'Holds the header information
    Texture As Direct3DTexture8  'Holds the texture of the text
    RowPitch As Integer          'Number of characters per row
    RowFactor As Single          'Percentage of the texture width each character takes
    ColFactor As Single          'Percentage of the texture height each character takes
    CharHeight As Byte           'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As tStructureLng 'Size of the texture

End Type

Private Texture            As clsSurfaceManager
Private SpriteBatch        As clsBatch
Public cfonts              As CustomFont

Private LastInvRender      As Long

Public White(0 To 3)       As Long
Public Red(0 To 3)         As Long
Public Orange(0 To 3)      As Long
Public Cyan(0 To 3)        As Long
Public Black(0 To 3)       As Long
Public FaintBlack(0 To 3)  As Long
Public Yellow(0 To 3)      As Long
Public Gray(0 To 3)        As Long
Public Transparent(0 To 3) As Long
Public Green(0 To 3)       As Long
Public Magenta(0 To 3) As Long
Public Blue(0 To 3)        As Long
Public ColorSA(0 To 3)     As Long
Public ColorSD(0 To 3)     As Long
Public ColorSD2(0 To 3)    As Long
Public Gris(0 To 3)        As Long
Public Caos(0 To 3) As Long
Public CaosClan(0 To 3) As Long
Public Real(0 To 3) As Long
Public RealClan(0 To 3) As Long
Public VerdeF(0 To 3) As Long
Public Tini(0 To 3) As Long
Public Onlines(0 To 3) As Long
Public Candados(0 To 3) As Long

Public EngineRun           As Boolean

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Dim timerElapsedTime         As Single
Public engineBaseSpeed       As Single
Public ScrollPixelFrame      As Single
Dim timerTicksPerFrame       As Single

'Map sizes in tiles
Public Const XMaxMapSize     As Byte = 100
Public Const XMinMapSize     As Byte = 1
Public Const YMaxMapSize     As Byte = 100
Public Const YMinMapSize     As Byte = 1

Private Const GrhFogata      As Integer = 1521

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Posicion en un mapa
Public Type Position

    X As Single
    Y As Single

End Type

'Posicion en el Mundo
Public Type WorldPos

    Map As Integer
    X As Integer
    Y As Integer

End Type
    
Public Type GrhData

    sX          As Integer
    sY          As Integer
    FileNum     As Long
    pixelWidth  As Integer
    pixelHeight As Integer
    TileWidth   As Single
    TileHeight  As Single
   
    NumFrames   As Integer
    Frames()    As Long
    Speed       As Single

End Type
 
Public Type Grh
    
    GrhIndex     As Long
    FrameCounter As Single
    Speed        As Single
    Started      As Byte
    Loops        As Integer
    
End Type

'Lista de cuerpos
Public Type BodyData

    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position

End Type

'Lista de cabezas
Public Type HeadData

    Head(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Lista de las animaciones de las armas
Public Type WeaponAnimData

    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
    WeaponAttack As Single

End Type

'Lista de las animaciones de los escudos
Public Type ShieldAnimData

    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
    ShieldAttack As Single

End Type

'Lista de cuerpos
Public Type FxData

    fx As Grh
    OffsetX As Long
    OffsetY As Long

End Type

Public Type rStats
      MinHP As Long
      MaxHP As Long
End Type

Public Type Char

    Particle_Count As Integer
    Particle_Group() As Long

    AnimTime As Byte
    Active As Byte
    Heading As Byte ' As E_Heading ?
    pos As WorldPos
    oldPos As WorldPos
    CvcBlue As Byte
    CvcRed As Byte

    BodyNum As Integer
    
    iHead As Integer
    iBody As Integer
    
    Alas As BodyData
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    
    EscudoEqu As Boolean
    UsandoArma As Boolean
    
    fx As Grh
    FxIndex As Integer
    
    Criminal As Byte
    Nombre As String
        
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    Pie As Boolean
    Muerto As Boolean
    Invisible As Boolean
    priv As Byte
    
    PartyIndex As Integer
    
    Stats As rStats
    
    NpcType As Byte
    
    Icono As Byte

End Type

'Info de un objeto
Public Type Obj

    ObjIndex As Integer
    Amount As Integer

End Type

Private Type tSangre
    AlphaB As Byte
    Activo As Boolean
    Estado As Byte
    grhSangre As Grh
    Counter As Long
    osX As Integer
    OSY As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock

    Graphic(1 To 4) As Grh
    charindex As Integer
    ObjGrh As Grh
    
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
    Particle_Group As Integer
    Color(0 To 3) As Long
    Sangre(1 To 5) As tSangre

End Type

'Info de cada mapa
Public Type MapInfo

    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    
End Type

'Particle Groups
Private Type Stream

    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
    Angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    Grh_List() As Long
    ColortInt(0 To 3) As Long
    
    Speed As Single
    life_counter As Long
   
End Type

Private Type Particle

    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    Angle As Single
    Grh As Grh
    alive_counter As Long
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    Rgb_List(0 To 3) As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer

End Type

Private Type Particle_Group

    Active As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    charindex As Long
 
    Frame_Counter As Single
    Frame_Speed As Single
    
    Stream_Type As Byte
 
    Particle_Stream() As Particle
    Particle_Count As Long
    
    Grh_Index_List() As Long
    Grh_Index_Count As Long
    
    alpha_blend As Boolean
    alive_counter As Long
    Never_Die As Boolean
    
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
    Angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    Rgb_List(0 To 3) As Long
    
    'Added by Juan Martín Sotuyo Dodero
    Speed As Single
    life_counter As Long
    
End Type

Private TotalStreams          As Integer
Private StreamData()          As Stream
Private Particle_Group_List() As Particle_Group
Private Particle_Group_Count  As Long
Private Particle_Group_Last   As Long

'Bordes del mapa
Public MinXBorder             As Byte
Public MaxXBorder             As Byte
Public MinYBorder             As Byte
Public MaxYBorder             As Byte

'Status del user
Public CurMap                 As Integer 'Mapa actual
Public userindex              As Integer
Public UserMoving             As Byte
Public UserBody               As Integer
Public UserHead               As Integer
Public UserPos                As WorldPos  'Posicion
Public AddtoUserPos           As WorldPos 'Si se mueve
Public UserCharIndex          As Integer

Public UserMaxAGU             As Integer
Public UserMinAGU             As Integer
Public UserMaxHAM             As Integer
Public UserMinHAM             As Integer

Public FramesPerSec As Integer

Public FPS                    As Long
Public FramesPerSecCounter    As Long
Private FpsLastCheck          As Long

Public ScreenWidth            As Integer
Public ScreenHeight           As Integer

'Tamaño del la vista en Tiles
Public WindowTileWidth        As Integer
Public WindowTileHeight       As Integer

Private HalfWindowTileHeight  As Integer
Private HalfWindowTileWidth   As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize         As Byte

'Tamaño de los tiles en pixels
Public TilePixelHeight        As Integer
Public TilePixelWidth         As Integer

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Totales?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public LastChar               As Integer
Public NumChars               As Integer
Public NumBodies              As Integer
Public NumHeads               As Integer
Public NumFxs                 As Integer
Public NumWeaponAnims         As Integer
Public NumShieldAnims         As Integer

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData()              As GrhData 'Guarda todos los grh
Public BodyData()             As BodyData
Public HeadData()             As HeadData
Public FxData()               As tIndiceFx
Public WeaponAnimData()       As WeaponAnimData
Public ShieldAnimData()       As ShieldAnimData
Public CascoAnimData()        As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData()              As MapBlock ' Mapa
Public MapInfo                As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bSecondaryAmbient      As Integer ' Está el ambiente cambiado?
Public bLluvia()              As Byte    ' Mapas en los que llueve
Public bTecho                 As Boolean ' Hay Techo?
Public CharList(1 To 10000)   As Char

Public Enum PlayLoop

    plNone = 0
    plLluviain = 1
    plLluviaout = 2

End Enum

Public IsPlaying As PlayLoop

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
                
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long

Public ColorAmbiente()   As D3DCOLORVALUE
Private colorActual      As D3DCOLORVALUE
Private colorFinal       As D3DCOLORVALUE
 
Private Fade             As Boolean
Private AmbientLastCheck As Long
 
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef source As Any, ByVal Length As Long)
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pClsid As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, _
    ByVal lSize As Long, _
    ByVal fRunmode As Long, _
    riid As Any, _
    ppvObj As Any) As Long

Public Function ArrayToPicture(inArray() As Byte, offset As Long, Size As Long) As IPicture
    Dim o_hMem        As Long
    Dim o_lpMem       As Long
    Dim aGUID(0 To 3) As Long
    Dim IIStream      As IUnknown
    
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    
    o_hMem = GlobalAlloc(&H2&, Size)

    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)

        If Not o_lpMem = 0& Then
            Call CopyMemory(ByVal o_lpMem, inArray(offset), Size)
            Call GlobalUnlock(o_hMem)

            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), ArrayToPicture)
            End If

        End If

    End If

End Function
 
Public Function PictureFromByteStream(B() As Byte) As IPicture
    Dim LowerBound As Long
    Dim byteCount  As Long
    Dim hMem       As Long
    Dim lpMem      As Long
    Dim IID_IPicture(15)
    Dim istm As stdole.IUnknown

    On Error GoTo Err_Init

    If UBound(B, 1) < 0 Then
        Exit Function

    End If
    
    LowerBound = LBound(B)
    byteCount = (UBound(B) - LowerBound) + 1
    hMem = GlobalAlloc(&H2, byteCount)

    If hMem <> 0 Then
        lpMem = GlobalLock(hMem)

        If lpMem <> 0 Then
            MoveMemory ByVal lpMem, B(LowerBound), byteCount
            Call GlobalUnlock(hMem)

            If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
                If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                    Call OleLoadPicture(ByVal ObjPtr(istm), byteCount, 0, IID_IPicture(0), PictureFromByteStream)

                End If

            End If

        End If

    End If
    
    Exit Function
    
Err_Init:

    If Err.Number = 9 Then
        'Uninitialized array
        MsgBox "You must pass a non-empty byte array to this function!"
    Else
        MsgBox Err.Number & " - " & Err.Description

    End If

End Function

Public Sub Ambient_LoadColor()

    Dim PathDir As String, X As Byte
    Dim i       As Long
    
    Dim Data()  As Byte
    Dim handle  As Integer
    
    If Not Get_File_Data(DirRecursos, "AMBIENTE.TXT", Data, INIT_RESOURCE_FILE) Then Exit Sub

    PathDir = DirRecursos & "AMBIENTE.TXT"
    
    handle = FreeFile
    Open PathDir For Binary Access Write As handle
    Put handle, , Data
    Close handle
    
    X = CByte(GetVar(PathDir, "AMBIENTE", "CANT"))

    ReDim ColorAmbiente(1 To X) As D3DCOLORVALUE
    
    For i = 1 To X
    
        With ColorAmbiente(i)
        
            .r = Val(GetVar(PathDir, i, "R"))
            .G = Val(GetVar(PathDir, i, "G"))
            .B = Val(GetVar(PathDir, i, "B"))
            .a = 255

        End With
    
    Next i
    
    If FileExist(PathDir, vbArchive) Then Call Kill(PathDir)

End Sub
 
Public Sub Ambient_Get(ByRef Color As D3DCOLORVALUE)
 
    Color = colorActual
 
End Sub
  
Public Sub Ambient_Fade()
 
    CalculateRGB colorFinal.r, colorFinal.G, colorFinal.B
    Fade = Not (colorFinal.r = colorActual.r And colorFinal.G = colorActual.G And colorFinal.B = colorActual.B)
 
    Dim X As Long, Y As Long, i As Long
 
    For X = 1 To 100
        For Y = 1 To 100
        
            For i = 0 To 3
            
                MapData(X, Y).Color(i) = D3DColorXRGB(colorActual.r, colorActual.G, colorActual.B)
 
            Next i
        Next Y
    Next X
  
End Sub
 
Public Sub Ambient_SetActual(ByVal r As Byte, ByVal G As Byte, ByVal B As Byte)
 
    colorActual.r = r
    colorActual.G = G
    colorActual.B = B
 
End Sub
 
Public Sub Ambient_SetFinal(ByVal Ambiente As Byte)
 
    With colorFinal
    
        .a = ColorAmbiente(Ambiente).a
        .r = ColorAmbiente(Ambiente).r
        .G = ColorAmbiente(Ambiente).G
        .B = ColorAmbiente(Ambiente).B
        
        Fade = Not (.r = colorActual.r And .G = colorActual.G And .B = colorActual.B)

    End With

End Sub
 
Public Sub CalculateRGB(ByVal r As Byte, ByVal G As Byte, ByVal B As Byte)
 
    With colorActual
 
        If .r < r Then
            .r = .r + 1
        ElseIf .r > r Then
            .r = .r - 1
 
        End If
 
        If .B < B Then
            .B = .B + 1
        ElseIf .B > B Then
            .B = .B - 1
 
        End If
 
        If .G < G Then
            .G = .G + 1
        ElseIf .G > G Then
            .G = .G - 1
 
        End If
 
    End With
 
End Sub
                
Sub CargarParticulas()
    Dim StreamFile As String
    Dim LooPC      As Long
    Dim i          As Long
    Dim GrhListing As String
    Dim TempSet    As String
    Dim ColorSet   As Long
    
    Dim Data()     As Byte
    Dim handle     As Integer
    
    If Not Get_File_Data(DirRecursos, "PARTICLES.INI", Data, INIT_RESOURCE_FILE) Then Exit Sub

    StreamFile = DirRecursos & "PARTICLES.INI"
    
    handle = FreeFile
    Open StreamFile For Binary Access Write As handle
    Put handle, , Data
    Close handle
    
    TotalStreams = Val(GetVar(StreamFile, "INIT", "Total"))
 
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
 
    'fill StreamData array with info from Particles.ini
    For LooPC = 1 To TotalStreams

        With StreamData(LooPC)
            .Name = GetVar(StreamFile, Val(LooPC), "Name")
            .NumOfParticles = GetVar(StreamFile, Val(LooPC), "NumOfParticles")
            .X1 = GetVar(StreamFile, Val(LooPC), "X1")
            .Y1 = GetVar(StreamFile, Val(LooPC), "Y1")
            .X2 = GetVar(StreamFile, Val(LooPC), "X2")
            .Y2 = GetVar(StreamFile, Val(LooPC), "Y2")
            .Angle = GetVar(StreamFile, Val(LooPC), "Angle")
            .vecx1 = GetVar(StreamFile, Val(LooPC), "VecX1")
            .vecx2 = GetVar(StreamFile, Val(LooPC), "VecX2")
            .vecy1 = GetVar(StreamFile, Val(LooPC), "VecY1")
            .vecy2 = GetVar(StreamFile, Val(LooPC), "VecY2")
            .life1 = GetVar(StreamFile, Val(LooPC), "Life1")
            .life2 = GetVar(StreamFile, Val(LooPC), "Life2")
            .friction = GetVar(StreamFile, Val(LooPC), "Friction")
            .spin = GetVar(StreamFile, Val(LooPC), "Spin")
            .spin_speedL = GetVar(StreamFile, Val(LooPC), "Spin_SpeedL")
            .spin_speedH = GetVar(StreamFile, Val(LooPC), "Spin_SpeedH")
            .AlphaBlend = GetVar(StreamFile, Val(LooPC), "AlphaBlend")
            .gravity = GetVar(StreamFile, Val(LooPC), "Gravity")
            .grav_strength = GetVar(StreamFile, Val(LooPC), "Grav_Strength")
            .bounce_strength = GetVar(StreamFile, Val(LooPC), "Bounce_Strength")
            .XMove = GetVar(StreamFile, Val(LooPC), "XMove")
            .YMove = GetVar(StreamFile, Val(LooPC), "YMove")
            .move_x1 = GetVar(StreamFile, Val(LooPC), "move_x1")
            .move_x2 = GetVar(StreamFile, Val(LooPC), "move_x2")
            .move_y1 = GetVar(StreamFile, Val(LooPC), "move_y1")
            .move_y2 = GetVar(StreamFile, Val(LooPC), "move_y2")
            .life_counter = GetVar(StreamFile, Val(LooPC), "life_counter")
            .Speed = Val(GetVar(StreamFile, Val(LooPC), "Speed"))
            .NumGrhs = GetVar(StreamFile, Val(LooPC), "NumGrhs")
        
            ReDim .Grh_List(1 To .NumGrhs) As Long
            GrhListing = GetVar(StreamFile, Val(LooPC), "Grh_List")
        
            For i = 1 To .NumGrhs
                .Grh_List(i) = ReadField(str(i), GrhListing, 44)
            Next i

            For ColorSet = 1 To 4
                TempSet = GetVar(StreamFile, Val(LooPC), "ColorSet" & ColorSet)
                .ColortInt(ColorSet - 1) = D3DColorXRGB(ReadField(1, TempSet, 44), ReadField(2, TempSet, 44), ReadField(3, TempSet, 44))
            Next ColorSet

        End With

    Next LooPC
    
    If FileExist(StreamFile, vbArchive) Then Call Kill(StreamFile)
 
End Sub
 
Public Function Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal Particle_Life As Long = 0) As Long

    If ParticulaInd > 0 And ParticulaInd <= TotalStreams Then

        With StreamData(ParticulaInd)
    
            Particle_Create = Particle_Group_Create(X, Y, .Grh_List, .ColortInt(), .NumOfParticles, ParticulaInd, .AlphaBlend, IIf(Particle_Life = _
                0, .life_counter, Particle_Life), .Speed, , .X1, .Y1, .Angle, .vecx1, .vecx2, .vecy1, .vecy2, .life1, .life2, .friction, _
                .spin_speedL, .gravity, .grav_strength, .bounce_strength, .X2, .Y2, .XMove, .move_x1, .move_x2, .move_y1, .move_y2, .YMove, _
                .spin_speedH, .spin)

        End With

    End If

End Function

Public Function Char_Particle_Create(ByVal ParticulaInd As Long, ByVal charindex As Integer, Optional ByVal Particle_Life As Long = 0) As Long

    If ParticulaInd > 0 And ParticulaInd <= TotalStreams Then

        With StreamData(ParticulaInd)

            Char_Particle_Create = Char_Particle_Group_Create(charindex, .Grh_List, .ColortInt(), .NumOfParticles, ParticulaInd, .AlphaBlend, IIf( _
                Particle_Life = 0, .life_counter, Particle_Life), .Speed, , .X1, .Y1, .Angle, .vecx1, .vecx2, .vecy1, .vecy2, .life1, .life2, _
                .friction, .spin_speedL, .gravity, .grav_strength, .bounce_strength, .X2, .Y2, .XMove, .move_x1, .move_x2, .move_y1, .move_y2, _
                .YMove, .spin_speedH, .spin)

        End With

    End If

End Function
                
Sub CargarCabezas()

    Dim i       As Long, NumHeads As Integer, MisCabezas() As tIndiceCabeza
    Dim Data()  As Byte
    Dim FileBuf As clsByteBuffer
    
    Set FileBuf = New clsByteBuffer
    
    If Not Get_File_Data(DirRecursos, "CABEZAS.IND", Data, INIT_RESOURCE_FILE) Then Exit Sub
    
    FileBuf.initializeReader Data
    
    MiCabecera.desc = FileBuf.getString(Len(MiCabecera.desc))
    MiCabecera.CRC = FileBuf.getLong
    MiCabecera.MagicWord = FileBuf.getLong
    
    'num de cabezas
    NumHeads = FileBuf.getInteger

    'Resize array
    ReDim HeadData(0 To NumHeads) As HeadData
    ReDim MisCabezas(0 To NumHeads) As tIndiceCabeza

    For i = 1 To NumHeads
    
        MisCabezas(i).Head(1) = FileBuf.getLong()
        MisCabezas(i).Head(2) = FileBuf.getLong()
        MisCabezas(i).Head(3) = FileBuf.getLong()
        MisCabezas(i).Head(4) = FileBuf.getLong()
        
        InitGrh HeadData(i).Head(1), MisCabezas(i).Head(1), 0
        InitGrh HeadData(i).Head(2), MisCabezas(i).Head(2), 0
        InitGrh HeadData(i).Head(3), MisCabezas(i).Head(3), 0
        InitGrh HeadData(i).Head(4), MisCabezas(i).Head(4), 0
    Next i

    Set FileBuf = Nothing

End Sub

Sub CargarCascos()

    Dim i       As Long, NumCascos As Integer, MisCabezas() As tIndiceCabeza
    Dim Data()  As Byte
    Dim FileBuf As clsByteBuffer
    
    Set FileBuf = New clsByteBuffer
    
    If Not Get_File_Data(DirRecursos, "CASCOS.IND", Data, INIT_RESOURCE_FILE) Then Exit Sub
    
    FileBuf.initializeReader Data
    
    MiCabecera.desc = FileBuf.getString(Len(MiCabecera.desc))
    MiCabecera.CRC = FileBuf.getLong
    MiCabecera.MagicWord = FileBuf.getLong
    
    'num de cascos
    NumCascos = FileBuf.getInteger

    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim MisCabezas(0 To NumCascos) As tIndiceCabeza

    For i = 1 To NumCascos
        MisCabezas(i).Head(1) = FileBuf.getLong()
        MisCabezas(i).Head(2) = FileBuf.getLong()
        MisCabezas(i).Head(3) = FileBuf.getLong()
        MisCabezas(i).Head(4) = FileBuf.getLong()
        
        InitGrh CascoAnimData(i).Head(1), MisCabezas(i).Head(1), 0
        InitGrh CascoAnimData(i).Head(2), MisCabezas(i).Head(2), 0
        InitGrh CascoAnimData(i).Head(3), MisCabezas(i).Head(3), 0
        InitGrh CascoAnimData(i).Head(4), MisCabezas(i).Head(4), 0
    Next i

    Set FileBuf = Nothing

End Sub

Sub CargarCuerpos()

    Dim i            As Long, NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    Dim Data()       As Byte
    Dim FileBuf      As clsByteBuffer
    
    Set FileBuf = New clsByteBuffer
    
    If Not Get_File_Data(DirRecursos, "PERSONAJES.IND", Data, INIT_RESOURCE_FILE) Then Exit Sub
    
    FileBuf.initializeReader Data
    
    MiCabecera.desc = FileBuf.getString(Len(MiCabecera.desc))
    MiCabecera.CRC = FileBuf.getLong
    MiCabecera.MagicWord = FileBuf.getLong
    
    'num de cascos
    NumCuerpos = FileBuf.getInteger

    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo

    For i = 1 To NumCuerpos
        MisCuerpos(i).Body(1) = FileBuf.getLong
        MisCuerpos(i).Body(2) = FileBuf.getLong
        MisCuerpos(i).Body(3) = FileBuf.getLong
        MisCuerpos(i).Body(4) = FileBuf.getLong
        MisCuerpos(i).HeadOffsetX = FileBuf.getInteger
        MisCuerpos(i).HeadOffsetY = FileBuf.getInteger
        
        InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
        InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
        InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
        InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
        
        BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
        BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        
    Next i

    Set FileBuf = Nothing

End Sub

Sub CargarFxs()

    Dim n       As Integer, i As Integer
    Dim NumFxs  As Integer

    Dim Data()  As Byte
    Dim FileBuf As clsByteBuffer
    
    Set FileBuf = New clsByteBuffer
    
    If Not Get_File_Data(DirRecursos, "FXS.IND", Data, INIT_RESOURCE_FILE) Then Exit Sub
    
    FileBuf.initializeReader Data
    
    MiCabecera.desc = FileBuf.getString(Len(MiCabecera.desc))
    MiCabecera.CRC = FileBuf.getLong
    MiCabecera.MagicWord = FileBuf.getLong

    'num de cabezas
    NumFxs = FileBuf.getInteger

    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    'ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx

    For i = 1 To NumFxs
        'Get #n, , FxData(i)
        FxData(i).Animacion = FileBuf.getLong
        FxData(i).OffsetX = FileBuf.getInteger
        FxData(i).OffsetY = FileBuf.getInteger

    Next i

    Set FileBuf = Nothing
    'Close #n

End Sub

Sub CargarArrayLluvia()

    Dim Read   As clsIniManager, i As Long
    Dim Nu     As Integer, FileName As String

    Dim Data() As Byte
    Dim handle As Integer
    
    If Not Get_File_Data(DirRecursos, "FK.TXT", Data, INIT_RESOURCE_FILE) Then Exit Sub

    FileName = DirRecursos & "FK.TXT"
    
    handle = FreeFile
    Open FileName For Binary Access Write As handle
    Put handle, , Data
    Close handle

    Set Read = New clsIniManager
    Call Read.Initialize(FileName)

    'num de cabezas
    Nu = CInt(Read.GetValue("INIT", "MapCant"))

    'Resize array
    ReDim bLluvia(1 To Nu) As Byte

    For i = 1 To Nu
        bLluvia(i) = CByte(Read.GetValue("INIT", "Map" & i))
    Next i

    Set Read = Nothing
    
    If FileExist(FileName, vbArchive) Then Call Kill(FileName)

End Sub

Public Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef TX As Byte, ByRef TY As Byte)

    '******************************************
    'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
    '******************************************
    TX = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    TY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2

End Sub

Sub MakeChar(ByVal charindex As Integer, _
    ByVal Body As Integer, _
    ByVal Head As Integer, _
    ByVal Heading As Byte, _
    ByVal X As Integer, _
    ByVal Y As Integer, _
    ByVal Arma As Integer, _
    ByVal Escudo As Integer, _
    ByVal Casco As Integer, _
    ByVal Alas As Integer)

    'Apuntamos al ultimo Char
    If charindex > LastChar Then LastChar = charindex

    With CharList(charindex)

        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .Active = 0 Then NumChars = NumChars + 1
    
        .iHead = Head
        .iBody = Body

        If InMapBounds(.oldPos.X, .oldPos.Y) Then
            MapData(.oldPos.X, .oldPos.Y).charindex = 0

        End If

        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
      
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        .Alas = BodyData(Alas)
        
        .Arma.WeaponAttack = 0
        .Escudo.ShieldAttack = 0
        
        .Heading = Heading

        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0

        'Update position
        .pos.X = X
        .pos.Y = Y
        
        .oldPos.X = .pos.X
        .oldPos.Y = .pos.Y

        'Make active
        .Active = 1

    End With

    'Plot on map
    MapData(X, Y).charindex = charindex

End Sub

Sub EraseChar(ByVal charindex As Integer)
 
    '*****************************************************************
    'Erases a character from CharList and map
    '*****************************************************************

    CharList(charindex).Active = 0

    'Update lastchar
    If charindex = LastChar Then

        Do Until CharList(LastChar).Active = 1
            LastChar = LastChar - 1

            If LastChar = 0 Then Exit Do
        Loop

    End If

    If InMapBounds(CharList(charindex).pos.X, CharList(charindex).pos.Y) Then
        MapData(CharList(charindex).pos.X, CharList(charindex).pos.Y).charindex = 0
    End If

    Call Dialogos.RemoveDialog(charindex)
    
    'Clear the selected index
    Dim TempChar As Char
    
    CharList(charindex) = TempChar
    
    'Update NumChars
    NumChars = NumChars - 1

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Long, Optional Started As Byte = 2)
    '*****************************************************************
    'Sets up a grh. MUST be done before rendering
    '*****************************************************************

    If GrhIndex > grhCount Then Exit Sub

    Grh.GrhIndex = GrhIndex

    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0

        End If

    Else
      
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started

    End If
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    
    End If

    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed

End Sub

Sub MoveCharbyHead(ByVal charindex As Integer, ByVal nHeading As E_Heading)
    '*****************************************************************
    'Starts the movement of a character in nHeading direction
    '*****************************************************************
    Dim addx As Integer
    Dim addy As Integer
    Dim X    As Integer
    Dim Y    As Integer
    Dim nX   As Integer
    Dim nY   As Integer

    X = CharList(charindex).pos.X
    Y = CharList(charindex).pos.Y

    'Figure out which way to move
    Select Case nHeading

        Case E_Heading.NORTH
            addy = -1

        Case E_Heading.EAST
            addx = 1

        Case E_Heading.SOUTH
            addy = 1
    
        Case E_Heading.WEST
            addx = -1
        
    End Select

    nX = X + addx
    nY = Y + addy

    If InMapBounds(nX, nY) Then
        MapData(nX, nY).charindex = charindex
        CharList(charindex).pos.X = nX
        CharList(charindex).pos.Y = nY
        MapData(X, Y).charindex = 0

    End If

    CharList(charindex).MoveOffsetX = -1 * (TilePixelWidth * addx)
    CharList(charindex).MoveOffsetY = -1 * (TilePixelHeight * addy)

    CharList(charindex).Moving = 1
    CharList(charindex).Heading = nHeading
 
    CharList(charindex).scrollDirectionX = addx
    CharList(charindex).scrollDirectionY = addy

    If UserEstado <> 1 Then Call DoPasosFx(charindex)

    If charindex <> UserCharIndex Then
        If Not EstaDentroDelArea(nX, nY) Then
            Call EraseChar(charindex)
        End If
    End If

End Sub

Public Sub DoFogataFx()

    If bFogata Then
        bFogata = HayFogata()

        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0

        End If

    Else
        bFogata = HayFogata()

        If bFogata And FogataBufferIndex = 0 Then
            FogataBufferIndex = Audio.PlayWave("fuego.wav", 0, 0, LoopStyle.Enabled)

        End If

    End If
 
End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean

    Dim X As Integer, Y As Integer

    For Y = UserPos.Y - MinYBorder + 1 To UserPos.Y + MinYBorder - 1
        For X = UserPos.X - MinXBorder + 1 To UserPos.X + MinXBorder - 1
            
            If MapData(X, Y).charindex = Index2 Then
                EstaPCarea = True
                Exit Function

            End If
        
        Next X
    Next Y

    EstaPCarea = False

End Function

Sub DoPasosFx(ByVal charindex As Integer)

    Static Pie As Boolean

    If Not UserNavegando Then
        If Not CharList(charindex).Muerto And EstaPCarea(charindex) Then
            CharList(charindex).Pie = Not CharList(charindex).Pie
            
            
            
            'CRAW; 18/03/2020 --> ARREGLO SONIDO 3D
              If MapData(CharList(charindex).pos.X, CharList(charindex).pos.Y).Graphic(1).GrhIndex >= 6000 And MapData(CharList(charindex).pos.X, CharList(charindex).pos.Y).Graphic(1).GrhIndex <= 6303 Then
               If CharList(charindex).Pie Then
                Call Audio.PlayWave(SND_PASOS3, CharList(charindex).pos.X, CharList(charindex).pos.Y)
            Else
                Call Audio.PlayWave(SND_PASOS4, CharList(charindex).pos.X, CharList(charindex).pos.Y)
            End If
            
            Else

            If CharList(charindex).Pie Then
                Call Audio.PlayWave(SND_PASOS1, CharList(charindex).pos.X, CharList(charindex).pos.Y)
            Else
                Call Audio.PlayWave(SND_PASOS2, CharList(charindex).pos.X, CharList(charindex).pos.Y)
            End If
         End If

        End If
        

    Else
        Call Audio.PlayWave(SND_NAVEGANDO, CharList(charindex).pos.X, CharList(charindex).pos.Y)

    End If

End Sub

Sub MoveCharbyPos(ByVal charindex As Integer, ByVal nX As Integer, ByVal nY As Integer)
 
    Dim X        As Integer
    Dim Y        As Integer
    Dim addx     As Integer
    Dim addy     As Integer
    Dim nHeading As E_Heading

    X = CharList(charindex).pos.X
    Y = CharList(charindex).pos.Y

    If InMapBounds(X, Y) Then
        MapData(X, Y).charindex = 0

        addx = nX - X
        addy = nY - Y

        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST

        End If

        If Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST

        End If

        If Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH

        End If

        If Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH

        End If

        MapData(nX, nY).charindex = charindex

        CharList(charindex).pos.X = nX
        CharList(charindex).pos.Y = nY

        CharList(charindex).MoveOffsetX = -1 * (TilePixelWidth * addx)
        CharList(charindex).MoveOffsetY = -1 * (TilePixelHeight * addy)

        CharList(charindex).Moving = 1
        CharList(charindex).Heading = nHeading
 
        CharList(charindex).scrollDirectionX = Sgn(addx)
        CharList(charindex).scrollDirectionY = Sgn(addy)

        'parche para que no medite cuando camina
        Dim fxCh As Integer
        fxCh = CharList(charindex).FxIndex

        If fxCh = FxMeditar.CHICO Or fxCh = FxMeditar.GRANDE Or fxCh = FxMeditar.MEDIANO Or fxCh = FxMeditar.XGRANDE Then
            CharList(charindex).FxIndex = 0
      
        End If

    End If

    If Not EstaPCarea(charindex) Then Call Dialogos.RemoveDialog(charindex)

    If Not EstaDentroDelArea(nX, nY) Then
        Call EraseChar(charindex)
    End If

End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
    '******************************************
    'Starts the screen moving in a direction
    '******************************************
    Dim X  As Integer
    Dim Y  As Integer
    Dim TX As Integer
    Dim TY As Integer

    'Figure out which way to move
    Select Case nHeading

        Case E_Heading.NORTH
            Y = -1

        Case E_Heading.EAST
            X = 1

        Case E_Heading.SOUTH
            Y = 1
    
        Case E_Heading.WEST
            X = -1
        
    End Select

    'Fill temp pos
    TX = UserPos.X + X
    TY = UserPos.Y + Y

    'Check to see if its out of bounds
    If TX < MinXBorder Or TX > MaxXBorder Or TY < MinYBorder Or TY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = TX
        AddtoUserPos.Y = Y
        UserPos.Y = TY
        UserMoving = 1
   
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, UserPos.Y).Trigger = 2 Or MapData(UserPos.X, _
            UserPos.Y).Trigger = 4, True, False)
   
    End If

End Sub

Function HayFogata() As Boolean
    Dim j As Integer, k As Integer

    For j = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6

            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    HayFogata = True
                    Exit Function

                End If

            End If

        Next k
    Next j

End Function
    
Public Function LoadGrhData() As Boolean

    Dim Grh         As Long
    Dim Frame       As Long
    Dim handle      As Integer
    Dim FileVersion As Long
    
    'Open files
    handle = FreeFile()
    
    Dim Data()       As Byte
    Dim TemporalFile As String

    Call Get_File_Data(DirRecursos, "GRAFICOS.IND", Data, INIT_RESOURCE_FILE)
    
    TemporalFile = DirRecursos & "Temporal.AOM"
        
    Open TemporalFile For Binary Access Write As handle
    Put handle, , Data
    Close handle
    
    'Open files
    handle = FreeFile()
    Open TemporalFile For Binary Access Read As handle
    
    Seek handle, 1
    
    Get handle, , FileVersion
    Get handle, , grhCount
    
    ReDim GrhData(0 To grhCount) As GrhData
    
    While Not EOF(handle)

        Get handle, , Grh
        
        If Grh <> 0 Then

            With GrhData(Grh)

                Get handle, , .NumFrames

                If .NumFrames <= 0 Then GoTo ErrorHandler
            
                ReDim .Frames(1 To .NumFrames)
            
                If .NumFrames > 1 Then

                    For Frame = 1 To .NumFrames
                        Get handle, , .Frames(Frame)

                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then

                            GoTo ErrorHandler

                        End If

                    Next Frame
                
                    Get handle, , .Speed
                
                    If .Speed <= 0 Then GoTo ErrorHandler
                
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight

                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth

                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                    .TileWidth = GrhData(.Frames(1)).TileWidth

                    If .TileWidth <= 0 Then GoTo ErrorHandler
                
                    .TileHeight = GrhData(.Frames(1)).TileHeight

                    If .TileHeight <= 0 Then GoTo ErrorHandler
                Else
                    Get handle, , .FileNum

                    If .FileNum <= 0 Then GoTo ErrorHandler
                
                    Get handle, , .sX

                    If .sX < 0 Then GoTo ErrorHandler
                
                    Get handle, , .sY

                    If .sY < 0 Then GoTo ErrorHandler
                
                    Get handle, , .pixelWidth

                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                    Get handle, , .pixelHeight

                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                
                    .Frames(1) = Grh

                End If

            End With

        End If

    Wend
    
    Close handle

    If FileExist(TemporalFile, vbArchive) Then Call Kill(TemporalFile)
    LoadGrhData = True
    Exit Function

ErrorHandler:

    If FileExist(TemporalFile, vbArchive) Then Call Kill(TemporalFile)
    LoadGrhData = False
    Call MsgBox("Error while loading the Grh.dat! Stopped at GRH number: " & Grh)
    Close handle

End Function

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean

    '*****************************************************************
    'Checks to see if a tile position is legal
    '*****************************************************************
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function

    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function

    End If
    
    '¿Hay un personaje?
    If MapData(X, Y).charindex > 0 Then
        Exit Function

    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function

    End If
    
    LegalPos = True

End Function

Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 01/08/2009
    'Checks to see if a tile position is legal, including if there is a casper in the tile
    '10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
    '01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
    '*****************************************************************
    Dim charindex As Integer
    
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function

    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function

    End If
    
    charindex = MapData(X, Y).charindex

    '¿Hay un personaje?
    If charindex > 0 Then
    
        If MapData(UserPos.X, UserPos.Y).Blocked = 1 Then
            Exit Function

        End If
        
        With CharList(charindex)

            ' Si no es casper, no puede pasar
            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else

                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.X, UserPos.Y) Then
                    If Not HayAgua(X, Y) Then Exit Function
                Else

                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(X, Y) Then Exit Function

                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If CharList(UserCharIndex).priv > 0 And CharList(UserCharIndex).priv < 6 Then
                    If CharList(UserCharIndex).Invisible = True Then Exit Function

                End If

            End If

        End With

    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function

    End If
    
    MoveToLegalPos = True

End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
    '*****************************************************************
    'Checks to see if a tile position is in the maps bounds
    '*****************************************************************

    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        InMapBounds = False
        Exit Function

    End If

    InMapBounds = True

End Function

Private Sub DrawGrhtoSurface(ByRef Grh As Grh, _
    ByVal X As Integer, _
    ByVal Y As Integer, _
    ByVal center As Byte, _
    ByVal Animate As Byte, _
    ByRef Color() As Long, _
    Optional ByVal killAtEnd As Byte = 1, _
    Optional ByVal Angle As Single = 0, _
    Optional ByVal AlphaB As Byte = 0)
 
    Dim CurrentGrhIndex As Long
    
    On Error GoTo hError
    
    If Grh.GrhIndex <= 0 Then Exit Sub
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * Velocidad

            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1

                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0

                        If killAtEnd Then Exit Sub

                    End If

                End If

            End If

        End If

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    If CurrentGrhIndex <= 0 Then Exit Sub
    If GrhData(CurrentGrhIndex).FileNum = 0 Then Exit Sub
    
    With GrhData(CurrentGrhIndex)

        'Center Grh over X,Y pos
        If center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * 16) + 16

            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * 32) + 32

            End If
    
        End If
        
        'Draw
        Call Directx_Render_Texture(.FileNum, X, Y, .pixelHeight, .pixelWidth, .sX, .sY, Color(), Angle, AlphaB)

    End With

hError:

    If Err.Number <> 0 Then
        If Err.Number = 9 And Grh.FrameCounter < 1 Then
            Grh.FrameCounter = 1
            Resume
        Else
            LogError "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
                vbCrLf & Err.Description
            End

        End If

    End If

End Sub

Public Sub DrawGrhIndextoSurface(ByVal GrhIndex As Long, _
    ByVal X As Integer, _
    ByVal Y As Integer, _
    ByVal center As Byte, _
    ByRef Color() As Long, _
    Optional ByVal Angle As Single = 0, _
    Optional ByVal AlphaB As Byte = 0)
                                 
    If GrhIndex < 0 Then Exit Sub
    
    With GrhData(GrhIndex)

        'Center Grh over X,Y pos
        If center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * 16) + 16

            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * 32) + 32

            End If

        End If

        'Draw
        Call Directx_Render_Texture(.FileNum, X, Y, .pixelHeight, .pixelWidth, .sX, .sY, Color(), Angle, AlphaB)

    End With

End Sub

Sub DrawGrhtoHdc(ByVal DestHdc As Long, ByVal GrhIndex As Long, ByRef DestRect As RECT)

    Dim FilePath As String
    Dim hDCSrc   As Long
    Dim PrevObj  As Long
    Dim Screen_X As Integer
    Dim Screen_Y As Integer
    
    Screen_X = DestRect.Left
    Screen_Y = DestRect.Top
    
    If GrhIndex <= 0 Then Exit Sub

    With GrhData(GrhIndex)

        If .NumFrames <> 1 Then GrhIndex = .Frames(1)
        
        Dim Data()  As Byte
        Dim BmpData As StdPicture
        
        'get Picture
        If Get_File_Data(DirRecursos, CStr(.FileNum) & ".BMP", Data, GRH_RESOURCE_FILE) Then
            Set BmpData = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        
            hDCSrc = CreateCompatibleDC(DestHdc)
            PrevObj = SelectObject(hDCSrc, BmpData)
       
           Call BitBlt(DestHdc, Screen_X, Screen_Y, .pixelWidth, .pixelHeight, hDCSrc, .sX, .sY, vbSrcCopy)
 
           Call DeleteDC(hDCSrc)
            
            Set BmpData = Nothing

        End If

    End With

End Sub

Private Sub CharRender(ByVal charindex As Integer, ByVal PixelOffSetX As Integer, ByVal PixelOffSetY As Integer, ByRef Light() As Long)

    Dim Moved As Boolean
    Dim PartyIndexTrue As Boolean
    Dim ColorClan(0 To 3) As Long
    Dim Color(0 To 3) As Long
    Dim MismoChar As Boolean
    Dim pos As Integer
    Dim TempChar As Char

    With CharList(charindex)

        If .Moving Then

            If .scrollDirectionX <> 0 Then

                .MoveOffsetX = .MoveOffsetX + ScrollPixelFrame * Sgn(.scrollDirectionX) * timerTicksPerFrame

                If .Body.Walk(.Heading).Speed > 0 Then .Body.Walk(.Heading).Started = 1

                .Alas.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1

                Moved = True
                .AnimTime = 10

                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0

                End If

            End If

            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelFrame * Sgn(.scrollDirectionY) * timerTicksPerFrame

                If .Body.Walk(.Heading).Speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Alas.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1

                Moved = True
                .AnimTime = 10

                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0

                End If

            End If

        End If

        If .Heading = 0 Then .Heading = 3

        If Moved = 0 Then

            If .AnimTime = 0 Then
                .Moving = 0

                .Body.Walk(.Heading).FrameCounter = 1
                .Body.Walk(.Heading).Started = 0

                .Alas.Walk(.Heading).FrameCounter = 1
                .Alas.Walk(.Heading).Started = 0

                If .Arma.WeaponAttack > 0 Then
                    .Arma.WeaponAttack = .Arma.WeaponAttack - 0.2

                    If .Arma.WeaponAttack <= 0 Then
                        .Arma.WeaponWalk(.Heading).Started = 0
                        .Arma.WeaponWalk(.Heading).FrameCounter = 1

                    End If

                Else
                    .Arma.WeaponWalk(.Heading).FrameCounter = 1
                    .Arma.WeaponWalk(.Heading).Started = 0

                End If

                If .Escudo.ShieldAttack > 0 Then
                    .Escudo.ShieldAttack = .Escudo.ShieldAttack - 0.2

                    If .Escudo.ShieldAttack <= 0 Then
                        .Escudo.ShieldWalk(.Heading).Started = 0
                        .Escudo.ShieldWalk(.Heading).FrameCounter = 1

                    End If

                Else
                    .Escudo.ShieldWalk(.Heading).FrameCounter = 1
                    .Escudo.ShieldWalk(.Heading).Started = 0

                End If

            Else
                .AnimTime = .AnimTime - 1

            End If

        End If

        PixelOffSetX = PixelOffSetX + .MoveOffsetX
        PixelOffSetY = PixelOffSetY + .MoveOffsetY

        Velocidad = 0.5
        MismoChar = (UserCharIndex = charindex)

        If Not .Invisible Or MismoClan(charindex) Or MismoChar Or MismaParty(charindex) Or EsGm(UserCharIndex) Then

            If MismoChar Then
                If CharList(UserCharIndex).PartyIndex > 0 And (CharList(UserCharIndex).Invisible = True Or CharList(UserCharIndex).Invisible = False) Then
                    PartyIndexTrue = True
                End If
            Else
                PartyIndexTrue = MismaParty(charindex, True)
            End If


            If .Heading = E_Heading.SOUTH Then
                If .Alas.Walk(.Heading).GrhIndex <> 0 Then
                    Call DrawGrhtoSurface(.Alas.Walk(.Heading), PixelOffSetX + .Body.HeadOffset.X, PixelOffSetY + .Body.HeadOffset.Y + 35, 1, 1, _
                                          White, 0)

                End If

            End If

            Call DrawGrhtoSurface(.Body.Walk(.Heading), PixelOffSetX, PixelOffSetY, 1, 1, White, 0)

            Call DrawGrhtoSurface(.Head.Head(.Heading), PixelOffSetX + .Body.HeadOffset.X, PixelOffSetY + .Body.HeadOffset.Y, 1, 0, White)

            If .Casco.Head(.Heading).GrhIndex <> 0 Then
                Call DrawGrhtoSurface(.Casco.Head(.Heading), PixelOffSetX + .Body.HeadOffset.X, PixelOffSetY + .Body.HeadOffset.Y, 1, 0, White)

            End If

            If .Heading <> E_Heading.SOUTH Then
                If .Alas.Walk(.Heading).GrhIndex <> 0 Then
                    Call DrawGrhtoSurface(.Alas.Walk(.Heading), PixelOffSetX + .Body.HeadOffset.X, PixelOffSetY + .Body.HeadOffset.Y + IIf(.Heading _
                                                                                                                                         = E_Heading.NORTH, 35, 35), 1, 1, White, 0)    'El primer 25, es cuando esta mirando para arriba, el siguiente 20 es cuando esta mirando para izquierda o derecha ?Ta?, anda cambiando el "20"
                End If
            End If

            Dim xx As Integer

            If .Arma.WeaponWalk(.Heading).GrhIndex <> 0 Then
                If .Body.HeadOffset.Y = -69 Then
                    xx = 31
                ElseIf .Body.HeadOffset.Y = -94 Then
                    xx = 59
                ElseIf .Body.HeadOffset.Y = -78 Then
                    xx = 42
                ElseIf .Body.HeadOffset.Y = -75 Then
                    xx = 37
                ElseIf .Body.HeadOffset.Y = -55 Then
                    xx = 21
                ElseIf .Body.HeadOffset.Y = -83 Then
                    xx = 45
                ElseIf .Body.HeadOffset.Y = -65 Then
                    xx = 27
                ElseIf .Body.HeadOffset.Y = -60 Then
                    xx = 22
                ElseIf .Body.HeadOffset.Y = -95 Then
                    xx = 60
                ElseIf .Body.HeadOffset.Y = -48 Then
                    xx = 14
                ElseIf .Body.HeadOffset.Y = -68 Then
                    xx = 30
                ElseIf .Body.HeadOffset.Y = -120 Then
                    xx = 85
                ElseIf .Body.HeadOffset.Y = -72 Then
                    xx = 34
                ElseIf .Body.HeadOffset.Y = -52 Then
                    xx = 18
                ElseIf .Body.HeadOffset.Y = -80 Then
                    xx = 44
                ElseIf .Body.HeadOffset.Y = -88 Then
                    xx = 52
                ElseIf .Body.HeadOffset.Y = -90 Then
                    xx = 54
                ElseIf .Body.HeadOffset.Y = -38 Then
                    xx = 4
                ElseIf .Body.HeadOffset.Y = -50 Then
                    xx = 16
                ElseIf .Body.HeadOffset.Y = -68 Then
                    xx = 30
                Else
                    xx = 0
                End If
                Call DrawGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffSetX, PixelOffSetY - xx, 1, 1, White, 0)
            End If

            If .Escudo.ShieldWalk(.Heading).GrhIndex <> 0 Then
                If .Body.HeadOffset.Y = -69 Then
                    xx = 31
                ElseIf .Body.HeadOffset.Y = -94 Then
                    xx = 59
                ElseIf .Body.HeadOffset.Y = -78 Then
                    xx = 40
                ElseIf .Body.HeadOffset.Y = -75 Then
                    xx = 37
                ElseIf .Body.HeadOffset.Y = -55 Then
                    xx = 21
                ElseIf .Body.HeadOffset.Y = -83 Then
                    xx = 45
                ElseIf .Body.HeadOffset.Y = -65 Then
                    xx = 27
                ElseIf .Body.HeadOffset.Y = -60 Then
                    xx = 22
                ElseIf .Body.HeadOffset.Y = -95 Then
                    xx = 60
                ElseIf .Body.HeadOffset.Y = -48 Then
                    xx = 14
                ElseIf .Body.HeadOffset.Y = -120 Then
                    xx = 85
                ElseIf .Body.HeadOffset.Y = -68 Then
                    xx = 30
                ElseIf .Body.HeadOffset.Y = -72 Then
                    xx = 34
                ElseIf .Body.HeadOffset.Y = -52 Then
                    xx = 18
                ElseIf .Body.HeadOffset.Y = -80 Then
                    xx = 44
                ElseIf .Body.HeadOffset.Y = -88 Then
                    xx = 52
                ElseIf .Body.HeadOffset.Y = -90 Then
                    xx = 54
                ElseIf .Body.HeadOffset.Y = -38 Then
                    xx = 4
                ElseIf .Body.HeadOffset.Y = -50 Then
                    xx = 16
                ElseIf .Body.HeadOffset.Y = -68 Then
                    xx = 30
                Else
                    xx = 0
                End If
                Call DrawGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffSetX, PixelOffSetY - xx, 1, 1, White, 0)
            End If
             
             Call ColoresNick(charindex, PixelOffSetX, PixelOffSetY, PartyIndexTrue)
        
             If PartyIndexTrue Then Call BarraParty(charindex, PixelOffSetX, PixelOffSetY)
            
        End If
        
        If .Icono = 1 Then
           With GrhData("17395")
                    Call Directx_Render_Texture(.FileNum, PixelOffSetX + 12.5, PixelOffSetY - 45, .pixelHeight, .pixelWidth, .sX, .sY, White, 0, 0)
           End With
        End If
        
        If .NpcType = eNPCType.nQuest Then
        
           If ProcesoQuest = 0 Then
            
           With GrhData("17394")
                    Call Directx_Render_Texture(.FileNum, PixelOffSetX + 12.5, PixelOffSetY - 45, .pixelHeight, .pixelWidth, .sX, .sY, White, 0, 0)
           End With
           
           ElseIf ProcesoQuest = 1 Then
           
           With GrhData("17397")
                    Call Directx_Render_Texture(.FileNum, PixelOffSetX + 10, PixelOffSetY - 45, .pixelHeight, .pixelWidth, .sX, .sY, White, 0, 0)
           End With
           
           ElseIf ProcesoQuest = 2 Then
            
            With GrhData("17396")
                    Call Directx_Render_Texture(.FileNum, PixelOffSetX + 10, PixelOffSetY - 45, .pixelHeight, .pixelWidth, .sX, .sY, White, 0, 0)
           End With
           
           End If
            
        End If

        Velocidad = 1

        Dim i As Long

        If .Particle_Count > 0 Then
            For i = 1 To .Particle_Count
                If .Particle_Group(i) > 0 Then
                    Call Particle_Group_Render(.Particle_Group(i), PixelOffSetX, PixelOffSetY)
                End If
            Next i
        End If

        Call Dialogos.UpdateDialogPos(PixelOffSetX + 4 + .Body.HeadOffset.X, PixelOffSetY + .Body.HeadOffset.Y, charindex)

        If .FxIndex <> 0 Then
            Dim Colormeditar(0 To 3) As Long
            Call longToArray(Colormeditar, D3DColorRGBA(255, 255, 255, 120))
            If AoSetup.bTransparencia = 0 Then
                Call DrawGrhtoSurface(.fx, PixelOffSetX + FxData(.FxIndex).OffsetX, PixelOffSetY + FxData(.FxIndex).OffsetY, 1, 1, Colormeditar)
            Else
                Call DrawGrhtoSurface(.fx, PixelOffSetX + FxData(.FxIndex).OffsetX, PixelOffSetY + FxData(.FxIndex).OffsetY, 1, 1, White)
            End If
            If .fx.Started = 0 Then .FxIndex = 0
        End If

    End With

End Sub

Public Sub BarraParty(ByVal charindex As Integer, _
                      ByVal PixelOffSetX As Integer, _
                      ByVal PixelOffSetY As Integer)
                      
    Dim pos As Integer
    Dim sClan As String
                               
    With CharList(charindex)
    
          If Nombres Then
               
              pos = getTagPosition(CharList(charindex).Nombre)
              sClan = mid$(CharList(charindex).Nombre, pos)
              
                  If .Stats.MaxHP > 0 Then
                    If Len(sClan) = 0 Then
                        Draw_Box PixelOffSetX - 32, PixelOffSetY + 42, CLng(((.Stats.MaxHP / 100) / (.Stats.MaxHP / 100)) * 100), 10, Black
                        Draw_Box PixelOffSetX - 32, PixelOffSetY + 42, CLng(((.Stats.MinHP / 100) / (.Stats.MaxHP / 100)) * 100), 10, Red
                    Else
                        Draw_Box PixelOffSetX - 32, PixelOffSetY + 54, CLng(((.Stats.MaxHP / 100) / (.Stats.MaxHP / 100)) * 100), 10, Black
                        Draw_Box PixelOffSetX - 32, PixelOffSetY + 54, CLng(((.Stats.MinHP / 100) / (.Stats.MaxHP / 100)) * 100), 10, Red
                    End If
                  End If
             End If
        
    End With
                       
End Sub

Public Sub ColoresNick(ByVal charindex As Integer, _
                       ByVal PixelOffSetX As Integer, _
                       ByVal PixelOffSetY As Integer, _
                       ByVal PartyIndexTrue As Boolean)
     
    Dim ColorClan(0 To 3) As Long

    Dim Color(0 To 3)     As Long

    Dim pos               As Integer
    
    Dim X                 As Integer
      
    With CharList(charindex)
           
        If Nombres Then
            If Len(.Nombre) <> 0 Then
                pos = getTagPosition(.Nombre)

                Dim lCenter     As Long

                Dim lCenterClan As Long

                '  If InStr(.Nombre, "<") > 0 And InStr(.Nombre, ">") > 0 Then
                                        
                Dim Line        As String

                Line = Left$(.Nombre, pos - 2)
                lCenter = (Len(Line) * 6 / 2) - 15
                                            
                Dim sClan As String

                sClan = mid$(.Nombre, pos)
                lCenterClan = (Len(sClan) * 6 / 2) - 15

                If .Criminal = 1 Then
                    'ColorClan = RGB(255, 0, 0)
                    ColorToArray ColorClan, CaosClan
                ElseIf .Criminal = 2 Then
                    'ColorClan = RGB(0, 255, 0)
                    ColorToArray ColorClan, Caos
                ElseIf .Criminal = 3 Then
                    'ColorClan = RGB(0, 255, 0)
                    ColorToArray ColorClan, White
                ElseIf .Criminal = 4 Then
                    'ColorClan = RGB(0, 255, 0)
                    ColorToArray ColorClan, Real
                ElseIf .Criminal = 5 Then
                    'ColorClan = RGB(150, 150, 150)
                    ColorToArray ColorClan, Tini
                Else
                    'ColorClan = RGB(0, 128, 255)
                    ColorToArray ColorClan, RealClan

                End If
                                        
                Select Case .priv
                                
                    Case 0
                            
                        If .Invisible = False Then
                               
                            If Len(sClan) = 0 Then
                                
                                If .Criminal = 0 Then
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Real)
                                ElseIf .Criminal = 1 Then
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, CaosClan)
                                ElseIf .Criminal = 2 Then
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Caos)
                                ElseIf .Criminal = 3 Then
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, White)
                                ElseIf .Criminal = 4 Then
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Real)
                                ElseIf .Criminal = 5 Then
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Tini)
                                Else
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, RealClan)

                                End If
                                
                            ElseIf Len(sClan) > 0 Then
                                     
                                If .Criminal = 0 Then
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Real)
                                    Call Text_Draw(PixelOffSetX - lCenterClan, PixelOffSetY + 40, sClan, RealClan)
                                     
                                ElseIf .Criminal = 1 Then
                                    longToArray Color, ColoresPJ(50)
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Caos)
                                    Call Text_Draw(PixelOffSetX - lCenterClan, PixelOffSetY + 40, sClan, CaosClan)
                                        
                                ElseIf .Criminal = 2 Then ' Templario
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Caos)
                                    Call Text_Draw(PixelOffSetX - lCenterClan, PixelOffSetY + 40, sClan, CaosClan)
                                        
                                ElseIf .Criminal = 3 Then ' Templario
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, White)
                                    Call Text_Draw(PixelOffSetX - lCenterClan, PixelOffSetY + 40, sClan, RealClan)
                                        
                                ElseIf .Criminal = 4 Then ' Clero
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Real)
                                    Call Text_Draw(PixelOffSetX - lCenterClan, PixelOffSetY + 40, sClan, RealClan)
                                        
                                ElseIf .Criminal = 5 Then ' Namesis
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Tini)
                                    Call Text_Draw(PixelOffSetX - lCenterClan, PixelOffSetY + 40, sClan, CaosClan)

                                End If
                                
                            End If
                            
                            If .CvcBlue = 1 Or .CvcRed = 1 Then

                                If .CvcBlue = 1 Then

                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Green)

                                    Call Text_Draw(PixelOffSetX - lCenterClan, PixelOffSetY + 40, sClan, Green)

                                ElseIf .CvcRed = 1 Then

                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Red)

                                    Call Text_Draw(PixelOffSetX - lCenterClan, PixelOffSetY + 40, sClan, Red)

                                End If

                            End If
                                
                        ElseIf .Invisible = True Then

                            If Len(sClan) = 0 Then
                                Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, VerdeF)
                            ElseIf Len(sClan) > 0 Then

                                If .Criminal = 0 Then
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Real)
                                    Call Text_Draw(PixelOffSetX - lCenterClan, PixelOffSetY + 40, sClan, RealClan)
                                     
                                ElseIf .Criminal = 1 Then
                                    longToArray Color, ColoresPJ(50)
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Caos)
                                    Call Text_Draw(PixelOffSetX - lCenterClan, PixelOffSetY + 40, sClan, VerdeF)
                                        
                                ElseIf .Criminal = 2 Then ' Templario
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Caos)
                                    Call Text_Draw(PixelOffSetX - lCenterClan, PixelOffSetY + 40, sClan, VerdeF)
                                        
                                ElseIf .Criminal = 3 Then ' Templario
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, White)
                                    Call Text_Draw(PixelOffSetX - lCenterClan, PixelOffSetY + 40, sClan, VerdeF)
                                        
                                ElseIf .Criminal = 4 Then ' Clero
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Real)
                                    Call Text_Draw(PixelOffSetX - lCenterClan, PixelOffSetY + 40, sClan, VerdeF)
                                        
                                ElseIf .Criminal = 5 Then ' Namesis
                                    Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Tini)
                                    Call Text_Draw(PixelOffSetX - lCenterClan, PixelOffSetY + 40, sClan, VerdeF)

                                End If
                                         
                            End If

                        End If
                        
                    Case 3  'admin
                        longToArray ColorClan, D3DColorXRGB(255, 128, 64)
                        Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, ColorClan)
                        Call Text_Draw(PixelOffSetX - lCenterClan, PixelOffSetY + 40, sClan, ColorClan)
                                      
                    Case Else 'el resto
                        longToArray Color, ColoresPJ(.priv)
                        Call Text_Draw(PixelOffSetX - lCenter, PixelOffSetY + 30, Line, Color)

                End Select

                '       End If
                               
            End If
         
        End If
      
    End With
       
End Sub


Sub crearsangrepos(ByVal X As Byte, ByVal Y As Byte)
        With MapData(X, Y)
            Dim ii As Long, haySlot As Boolean
            For ii = 1 To 5
                If .Sangre(ii).Activo = False Then
                    .Sangre(ii).Activo = True
                    .Sangre(ii).AlphaB = 0
                    .Sangre(ii).Estado = 0
                    .Sangre(ii).Counter = GetTickCount
                    haySlot = True
                    Exit For
                End If
            Next ii
            
            If haySlot = False Then
                .Sangre(1).Activo = True
                .Sangre(1).AlphaB = 0
                .Sangre(1).Estado = 0
                .Sangre(1).Counter = GetTickCount
            End If
            
        End With
End Sub
Sub CrearSangre(ByVal nChar As Integer, Optional ByVal Npc As Boolean = False)
    If nChar <= 0 Or nChar > 10000 Then Exit Sub
    With CharList(nChar).pos
        Dim ii As Long, haySlot As Boolean
        Dim xDistance As Integer
        xDistance = ((32 * 3) / 2) / 2
        With MapData(.X, .Y)
            For ii = 1 To 5
                If .Sangre(ii).Activo = False Then
                    .Sangre(ii).Activo = True
                    .Sangre(ii).AlphaB = 0
                    .Sangre(ii).Estado = 0
                    .Sangre(ii).Counter = GetTickCount
                    
                    If Npc = True Then
                        
                        .Sangre(ii).osX = RandomNumber(-xDistance, xDistance)
                        .Sangre(ii).OSY = RandomNumber(-xDistance, xDistance)
                    Else
                        .Sangre(ii).osX = 0
                        .Sangre(ii).OSY = 0
                    End If
                    haySlot = True
                    Exit For
                End If
            Next ii
            
            If haySlot = False Then
                .Sangre(1).Activo = True
                .Sangre(1).AlphaB = 0
                .Sangre(1).Estado = 0
                .Sangre(1).Counter = GetTickCount
                If Npc = True Then
                    .Sangre(1).osX = RandomNumber(-xDistance, xDistance)
                    .Sangre(1).OSY = RandomNumber(-xDistance, xDistance)
                Else
                    .Sangre(ii).osX = 0
                    .Sangre(ii).OSY = 0
                End If
            End If
            

        End With
    End With
End Sub


Sub RenderScreen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffSetX As Integer, ByVal PixelOffSetY As Integer)

On Local Error Resume Next

    Dim Y          As Long     'Keeps track of where on map we are
    Dim X          As Long     'Keeps track of where on map we are
    
    Dim screenminY As Integer  'Start Y pos on current screen
    Dim screenmaxY As Integer  'End Y pos on current screen
    Dim screenminX As Integer  'Start X pos on current screen
    Dim screenmaxX As Integer  'End X pos on current screen
    
    Dim minY       As Integer  'Start Y pos on current map
    Dim maxY       As Integer  'End Y pos on current map
    Dim minX       As Integer  'Start X pos on current map
    Dim maxX       As Integer  'End X pos on current map
    
    Dim ScreenX    As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY    As Integer  'Keeps track of where to place tile on screen
    
    Dim minXOffset As Integer
    Dim minYOffset As Integer
    
    Dim iPPX       As Integer 'For centering grhs
    Dim iPPY       As Integer 'For centering grhs
    Dim sColor(3) As Long
    Dim ii As Long
    Static cSangre As Long
    
    'Figure out Ends and Starts of screen
    screenminY = TileY - HalfWindowTileHeight
    screenmaxY = TileY + HalfWindowTileHeight
    screenminX = TileX - HalfWindowTileWidth
    screenmaxX = TileX + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize * 2 ' WyroX: Parche para que no desaparezcan techos y arboles
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize

    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize

    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1

    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1

    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1

    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX

            If InMapBounds(X, Y) Then

                iPPX = (ScreenX - 1) * 32 + PixelOffSetX
                iPPY = (ScreenY - 1) * 32 + PixelOffSetY

                With MapData(X, Y)

                    'Layer 1 **********************************
                    If .Graphic(1).GrhIndex <> 0 Then
                        Call DrawGrhtoSurface(.Graphic(1), iPPX, iPPY, 0, 1, .Color())

                    End If

                    '******************************************
                    'Layer 2 **********************************
                    If .Graphic(2).GrhIndex <> 0 Then
                        Call DrawGrhtoSurface(.Graphic(2), iPPX, iPPY, 1, 1, .Color())

                    End If

                    '******************************************
                End With

            End If
           
            ScreenX = ScreenX + 1
        Next X

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y

    'Draw Transparent Layers  (Layer 2, 3)
    ScreenY = minYOffset - TileBufferSize
    Dim nSangre(4) As Grh, temp_rgb(3) As Long

    
    InitGrh nSangre(0), 17350
    InitGrh nSangre(1), 17351
    InitGrh nSangre(2), 17352
    InitGrh nSangre(3), 17353
    InitGrh nSangre(4), 17354

    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize

        For X = minX To maxX

            If InMapBounds(X, Y) Then
                iPPX = ScreenX * 32 + PixelOffSetX
                iPPY = ScreenY * 32 + PixelOffSetY

                With MapData(X, Y)

                    'Object Layer **********************************
                    If .ObjGrh.GrhIndex <> 0 Then
                        Call DrawGrhtoSurface(.ObjGrh, iPPX, iPPY, 1, 1, .Color())
         
                    End If
                    For ii = 1 To 5
                        If .Sangre(ii).Activo = True Then
                            If .Sangre(ii).Estado = 4 Then
                                temp_rgb(0) = D3DColorRGBA(255, 255, 255, 255 - .Sangre(ii).AlphaB)
                                    temp_rgb(1) = temp_rgb(0)
                                    temp_rgb(2) = temp_rgb(0)
                                    temp_rgb(3) = temp_rgb(0)
                                Call DrawGrhtoSurface(nSangre(4), iPPX + .Sangre(ii).osX, iPPY + .Sangre(ii).OSY, 1, 0, temp_rgb) ', , , .Sangre(ii).AlphaB)
                                
                               
                                    If GetTickCount - .Sangre(ii).Counter > IIf((.Sangre(ii).AlphaB > 0), 1500, 10) Then
                                        
                                        .Sangre(ii).AlphaB = (.Sangre(ii).AlphaB) + 1
                                         ', , , .Sangre(ii).AlphaB)
                                        If .Sangre(ii).AlphaB = 255 Then .Sangre(ii).Activo = False
                                        'Debug.Print .Sangre(ii).AlphaB
                                        cSangre = .Sangre(ii).Counter
                                    End If
                                    
                            Else
                                
                                 Call DrawGrhtoSurface(nSangre(.Sangre(ii).Estado), iPPX + .Sangre(ii).osX, iPPY + .Sangre(ii).OSY, 1, 0, .Color)
                                 If GetTickCount - .Sangre(ii).Counter > 50 Then
                                    .Sangre(ii).Estado = (.Sangre(ii).Estado) + 1
                                    .Sangre(ii).Counter = GetTickCount
                                    'Debug.Print "subio estado"
                                End If
                            End If
                        End If
                    Next ii

                    '***********************************************
                    'Char layer ************************************
                    If .charindex <> 0 Then
                        Call CharRender(.charindex, iPPX, iPPY, White)

                    End If

                    '*************************************************
                    'Layer 3 *****************************************
                    If .Graphic(3).GrhIndex <> 0 Then
                        Call DrawGrhtoSurface(.Graphic(3), iPPX, iPPY, 1, 1, .Color())

                    End If

                    '*************************************************
                End With

            End If

            '************************************************
            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y

    ScreenY = minYOffset - 5

    'Draw blocked tiles and grid
    ScreenY = minYOffset - TileBufferSize

    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize

        For X = minX To maxX

            'Check to see if in bounds
            If InMapBounds(X, Y) Then

                With MapData(X, Y)
                    iPPX = ScreenX * 32 + PixelOffSetX
                    iPPY = ScreenY * 32 + PixelOffSetY

                    If .Particle_Group > 0 Then
                        Call Particle_Group_Render(.Particle_Group, iPPX, iPPY)
                
                    End If

                    If Not bTecho And .Graphic(4).GrhIndex <> 0 Then
                        Call DrawGrhtoSurface(.Graphic(4), iPPX, iPPY, 1, 1, .Color())

                    End If

                End With

            End If

            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y
   
    If bSecondaryAmbient > 0 And bLluvia(UserMap) > 0 Then Call Particle_Group_Render(bSecondaryAmbient, 0, 0)

End Sub

Public Sub RenderSounds()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 4/22/2006
    'Actualiza todos los sonidos del mapa.
    '**************************************************************
    
    If bSecondaryAmbient > 0 And bLluvia(UserMap) = 1 Then
  
        If bTecho Then
            
            If IsPlaying <> PlayLoop.plLluviain Then
                If RainBufferIndex Then Call Audio.StopWave(RainBufferIndex)
                RainBufferIndex = Audio.PlayWave("lluviain.wav", 0, 0, LoopStyle.Enabled)
                IsPlaying = PlayLoop.plLluviain

            End If

        Else

            If IsPlaying <> PlayLoop.plLluviaout Then
                If RainBufferIndex Then Call Audio.StopWave(RainBufferIndex)
                RainBufferIndex = Audio.PlayWave("lluviaout.wav", 0, 0, LoopStyle.Enabled)
                IsPlaying = PlayLoop.plLluviaout

            End If

        End If

    End If
    
    Call DoFogataFx

End Sub

Public Sub SetCharacterFx(ByVal charindex As Integer, ByVal fx As Integer, ByVal Loops As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************
    With CharList(charindex)

        If fx > UBound(FxData()) Then Exit Sub
    
        .FxIndex = fx

        If .FxIndex > 0 Then
            Call InitGrh(.fx, FxData(fx).Animacion)
            .fx.Loops = Loops

        End If

    End With

End Sub

'[END]'
Function InitTileEngine(ByVal setDisplayFormhWnd As Long, _
    ByVal setTilePixelHeight As Integer, _
    ByVal setTilePixelWidth As Integer, _
    ByVal setWindowTileHeight As Integer, _
    ByVal setWindowTileWidth As Integer) As Boolean
                        
    '*****************************************************************
    'InitEngine
    '*****************************************************************
    
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    
    HalfWindowTileHeight = setWindowTileHeight \ 2
    HalfWindowTileWidth = setWindowTileWidth \ 2
    
    TileBufferSize = ModAreas.TilesBuffer

    Call CalcularAreas(HalfWindowTileWidth, HalfWindowTileHeight)
    
    FPS = 144
    FramesPerSecCounter = 120
    engineBaseSpeed = 0.017
    ScrollPixelFrame = 7.65

    MinXBorder = XMinMapSize + (HalfWindowTileWidth)
    MaxXBorder = XMaxMapSize - (HalfWindowTileWidth)
    MinYBorder = YMinMapSize + (HalfWindowTileHeight)
    MaxYBorder = YMaxMapSize - (HalfWindowTileHeight)

    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    
    Call AddtoRichTextBox(frmCargando.status, "Cargando Gráficos....", 255, 0, 0, True, False, True)
    
    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    capturaPath = App.Path & "/imagenes/cpt.bmp"
    Velocidad = 1

    'We are done!
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)

    InitTileEngine = True

End Function

Private Sub ShowNextFrame()

    Static OffsetCounterX As Double
    Static OffsetCounterY As Double
    
    Call DirectDevice.BeginScene

    If UserMoving Then

        If AddtoUserPos.X <> 0 Then
            OffsetCounterX = OffsetCounterX - ScrollPixelFrame * AddtoUserPos.X * timerTicksPerFrame

            If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.X) Then
                OffsetCounterX = 0
                AddtoUserPos.X = 0
                UserMoving = False

            End If

        End If
     
        '****** Move screen Up and Down if needed ******
        If AddtoUserPos.Y <> 0 Then
            OffsetCounterY = OffsetCounterY - ScrollPixelFrame * AddtoUserPos.Y * timerTicksPerFrame

            If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.Y) Then
                OffsetCounterY = 0
                AddtoUserPos.Y = 0
                UserMoving = False

            End If

        End If

    End If

    '****** Update screen ******
    Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
    
    
    If Estadisticas = True Then
    
    If (UserClicado > 0) Then
 
         Call Text_Draw(315 + CharList(UserClicado).pos.X, 108 + CharList(UserClicado).pos.Y, "" & "Nick: " & CharList(UserClicado).Nombre, Orange)
        
         Call Text_Draw(315, 118, "" & "---------------", White)
        
         Call Text_Draw(315, 128, "" & "Kill: " & ClickMatados, Red)
        
         Call Text_Draw(315, 138, "" & "---------------", White)
        
         Call Text_Draw(315, 148, "" & "Clase: " & ClickClase, Green)
        
         Call Text_Draw(315, 158, "" & "---------------", White)
         
    End If
 
  End If
    
    If IsSeguro Then
        Call Text_Draw(0, 13, "S", ColorSD)
        Call Text_Draw(12, 13, "SEGURO", ColorSD2)
    Else
        Call Text_Draw(0, 13, "S", Candados)

    End If

    If IsSeguroClan Then
        Call Text_Draw(0, 26, "W", ColorSD)
        Call Text_Draw(12, 26, "SEGURO CLAN", ColorSD2)
    Else
        Call Text_Draw(0, 26, "W", Candados)

    End If

    If IsSeguroCombate Then
        Call Text_Draw(0, 39, "C", ColorSD)
        Call Text_Draw(12, 39, "SEGURO COMBATE", ColorSD2)
    Else
        Call Text_Draw(0, 39, "C", Candados)

    End If

    If Not IsSeguroHechizos Then
        Call Text_Draw(0, 65, "*", ColorSD)
        Call Text_Draw(12, 65, "SEGURO HECHIZOS", ColorSD2)
    Else
        Call Text_Draw(0, 65, "*", Candados)

    End If

    If IsSeguroObjetos Then
        Call Text_Draw(0, 52, "X", ColorSD)
        Call Text_Draw(12, 52, "SEGURO OBJETOS", ColorSD2)
    Else
        Call Text_Draw(0, 52, "X", Candados)

    End If
    
    If VidaAmarilla > 0 And StatusAmarilla = True Then
        Call Text_Draw(470, 78, "Agilidad:", Gris)

        If Amarilla <= 20 Then
            Call Text_Draw(520, 78, Amarilla, Red)
        ElseIf Amarilla >= 21 And Amarilla <= 34 Then
            Call Text_Draw(520, 78, Amarilla, Orange)
        ElseIf Amarilla >= 35 Then
            Call Text_Draw(520, 78, Amarilla, Green)

            If VidaAmarilla > 20 Then
                Call Text_Draw(520, 78, Amarilla, Green)
            ElseIf VidaAmarilla <= 20 And VidaAmarilla > 10 Then
                Call Text_Draw(520, 78, Amarilla, Orange)
            ElseIf VidaAmarilla <= 10 Then
                Call Text_Draw(520, 78, Amarilla, Red)

            End If

        End If
       
    End If

    If VidaVerde > 0 And StatusVerde = True Then
        Call Text_Draw(0, 78, "Fuerza:", Gris)

        If Verde <= 20 Then
            Call Text_Draw(45, 78, Verde, Red)
        ElseIf Verde >= 21 And Verde <= 34 Then
            Call Text_Draw(45, 78, Verde, Orange)
        ElseIf Verde >= 35 Then
            Call Text_Draw(45, 78, Verde, Green)

            If VidaVerde > 20 Then
                Call Text_Draw(45, 78, Verde, Green)
            ElseIf VidaVerde <= 20 And VidaVerde > 10 Then
                Call Text_Draw(45, 78, Verde, Orange)
            ElseIf VidaVerde <= 10 Then
                Call Text_Draw(45, 78, Verde, Red)

            End If

        End If

    End If

    If CartelInvisibilidad > 0 Then
        Call Text_Draw(200, 13, "TIEMPO INVISIBLE: " & Int(CartelInvisibilidad / 40) & " seg", Cyan)

    End If

    Call Text_Draw(0, 0, NameMap & " (" & UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y & ")", White)

    Call Text_Draw(465, 0, "Hora: " & TimeChange & ":00", White)
    Call Text_Draw(415, 0, "FPS: " & FPS, White)
    Call Text_Draw(465, 16, NameDay, White)
    Call Text_Draw(430, 32, "Hay " & NumUsers & " Usuarios Online.", Onlines)
    
    If Not SeguroCvc = True Then
        Call Text_Draw(0, 91, "P", ColorSD)
        Call Text_Draw(12, 91, "SEGURO CVC", ColorSD2)
    Else
        Call Text_Draw(0, 91, "P", Candados)

    End If
    
    If TiempoAsedio <> 0 And (UserMap = 114 Or UserMap = 115) Then Call Text_Draw(260, 655, "Faltan " & TiempoAsedio & " minutos para que finalize el Asedio.", Cyan)

    Call Dialogos.Render
    Call DibujarCartel
    Call DialogosClanes.Draw
    
    '*******************************
    'Flip the backbuffer to the screen
    Call Directx_EndScene(MainViewRect, 0)

    '*******************************
    If GetTickCount - LastInvRender > 56 Then
        Call Inventario.RenderInv
        LastInvRender = GetTickCount

    End If
    
    If Fade Then Ambient_Fade

End Sub

Public Function GetElapsedTime() As Single
    
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Gets the time that past since the last call
    '**************************************************************
    Dim start_time    As Currency
    Static end_time   As Currency
    Static timer_freq As Currency
    
    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    
    End If
        
    'Get current time
    QueryPerformanceCounter start_time
        
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
        
    'Get next end time
    QueryPerformanceCounter end_time
    
End Function

Private Function MismaParty(ByVal charindex As Integer, Optional ByVal Invisible As Boolean = False) As Boolean

        MismaParty = False
        
        If CharList(charindex).PartyIndex > 0 And CharList(UserCharIndex).PartyIndex > 0 Then
            If CharList(charindex).PartyIndex = CharList(UserCharIndex).PartyIndex Then
                        MismaParty = True
                        Exit Function
            End If
        End If
         
End Function

Private Function isSailing(ByVal userindex As Integer) As Boolean

    isSailing = False

    With CharList(userindex)
        If .Body.Walk(.Heading).GrhIndex >= 10380 And .Body.Walk(.Heading).GrhIndex <= 10383 Then isSailing = True
    End With


End Function
Private Function MismoClan(ByVal userindex As Integer) As Boolean

    Dim pos As Integer
    Dim SuClan As String
    Dim MiClan As String
    
    pos = getTagPosition(CharList(userindex).Nombre)
    SuClan = mid(CharList(userindex).Nombre, pos)
    pos = getTagPosition(CharList(UserCharIndex).Nombre)
    MiClan = mid(CharList(UserCharIndex).Nombre, pos)
    
    MismoClan = False
    
    If Len(SuClan) > 0 And Len(MiClan) > 0 Then
        If SuClan = MiClan Then
          MismoClan = True
        End If
    End If

    'If InStr(CharList(UserIndex).Nombre, "<") > 0 And InStr(CharList(UserCharIndex).Nombre, ">") > 0 Then
    '    If UserClan = mid(CharList(UserIndex).Nombre, InStr(CharList(UserIndex).Nombre, "<")) Then
    '        MismoClan = True
    '    End If
    'End If

End Function

Private Sub InitColours()
        
'    White(0) = D3DColorXRGB(255, 255, 255)
'    White(1) = White(0)
'    White(2) = White(0)
'    White(3) = White(0)
    
    longToArray White, D3DColorXRGB(255, 255, 255)
    
'    Red(0) = D3DColorXRGB(255, 0, 0)
'    Red(1) = Red(0)
'    Red(2) = Red(0)
'    Red(3) = Red(0)
    
    longToArray Red, D3DColorXRGB(255, 0, 0)
    
'    Cyan(0) = D3DColorXRGB(0, 255, 255)
'    Cyan(1) = Cyan(0)
'    Cyan(2) = Cyan(0)
'    Cyan(3) = Cyan(0)
    
    longToArray Cyan, D3DColorXRGB(0, 255, 255)
    
'    Black(0) = D3DColorARGB(255, 0, 0, 0)
'    Black(1) = Black(0)
'    Black(2) = Black(0)
'    Black(3) = Black(0)
    
    longToArray Black, D3DColorARGB(255, 0, 0, 0)
    
'    FaintBlack(0) = D3DColorXRGB(0, 128, 255)
'    FaintBlack(1) = FaintBlack(0)
'    FaintBlack(2) = FaintBlack(0)
'    FaintBlack(3) = FaintBlack(0)
    
    longToArray FaintBlack, D3DColorXRGB(0, 128, 255)
    
    'Orange(0) = D3DColorXRGB(239, 127, 26)
    'Orange(1) = Orange(0)
    'Orange(2) = Orange(0)
    'Orange(3) = Orange(0)
    
    longToArray Orange, D3DColorXRGB(239, 127, 26)
    
    'Yellow(0) = D3DColorXRGB(255, 255, 0)
    'Yellow(1) = Yellow(0)
    'Yellow(2) = Yellow(0)
    'Yellow(3) = Yellow(0)
    
    longToArray Yellow, D3DColorXRGB(255, 255, 0)
    
    'Gray(0) = D3DColorXRGB(150, 150, 150)
    'Gray(1) = Gray(0)
    'Gray(2) = Gray(0)
    'Gray(3) = Gray(0)
    
    longToArray Gray, D3DColorXRGB(150, 150, 150)
    
    Transparent(0) = D3DColorXRGB(255, 255, 255)
    Transparent(1) = Transparent(0)
    Transparent(2) = Transparent(0)
    Transparent(3) = Transparent(0)
    
    'Green(0) = D3DColorXRGB(0, 255, 0)
    'Green(1) = Green(0)
    'Green(2) = Green(0)
    'Green(3) = Green(0)
    
    longToArray Green, D3DColorXRGB(0, 255, 0)
    
    longToArray Magenta, D3DColorXRGB(143, 69, 136)
    
    'Blue(0) = D3DColorXRGB(0, 0, 255)
    'Blue(1) = Blue(0)
    'Blue(2) = Blue(0)
    'Blue(3) = Blue(0)
    
    longToArray Blue, D3DColorXRGB(0, 0, 255)
    
    'ColorSA(0) = D3DColorXRGB(112, 116, 248)
    'ColorSA(1) = ColorSA(0)
    'ColorSA(2) = ColorSA(0)
    'ColorSA(3) = ColorSA(0)
    
    longToArray ColorSA, D3DColorXRGB(152, 152, 246)
    
    'ColorSD(0) = D3DColorXRGB(248, 116, 112)
    'ColorSD(1) = ColorSD(0)
    'ColorSD(2) = ColorSD(0)
    'ColorSD(3) = ColorSD(0)
    
    longToArray ColorSD, D3DColorXRGB(248, 116, 112)
    
    'ColorSD2(0) = D3DColorXRGB(200, 140, 8)
    'ColorSD2(1) = ColorSD2(0)
    'ColorSD2(2) = ColorSD2(0)
    'ColorSD2(3) = ColorSD2(0)
    
    longToArray ColorSD2, D3DColorXRGB(200, 140, 8)
    
    'Gris(0) = D3DColorXRGB(153, 153, 153)
    'Gris(1) = Gris(1)
    'Gris(2) = Gris(2)
    'Gris(3) = Gris(3)
    
    longToArray Gris, D3DColorXRGB(153, 153, 153)
    
    'Caos(0) = D3DColorXRGB(204, 0, 0)
    'Caos(1) = Caos(1)
    'Caos(2) = Caos(2)
    'Caos(3) = Caos(3)
    
    longToArray Caos, D3DColorXRGB(204, 0, 0)
    
'    CaosClan(0) = D3DColorXRGB(255, 0, 0)
'    CaosClan(1) = CaosClan(1)
'    CaosClan(2) = CaosClan(2)
'    CaosClan(3) = CaosClan(3)
    
    longToArray CaosClan, D3DColorXRGB(255, 0, 0)
    
'    Real(0) = D3DColorXRGB(0, 102, 204)
'    Real(1) = Real(1)
'    Real(2) = Real(2)
'    Real(3) = Real(3)
    
    longToArray Real, D3DColorXRGB(0, 102, 204)
    
'    RealClan(0) = D3DColorXRGB(0, 153, 255)
'    RealClan(1) = RealClan(1)
'    RealClan(2) = RealClan(2)
'    RealClan(3) = RealClan(3)
    
    longToArray RealClan, D3DColorXRGB(0, 153, 255)
    
'    VerdeF(0) = D3DColorXRGB(0, 255, 153)
'    VerdeF(1) = VerdeF(1)
'    VerdeF(2) = VerdeF(2)
'    VerdeF(3) = VerdeF(3)
    
    longToArray VerdeF, D3DColorXRGB(0, 255, 153)
    
'    Tini(0) = D3DColorXRGB(153, 153, 153)
'    Tini(1) = Tini(1)
'    Tini(2) = Tini(2)
'    Tini(3) = Tini(3)
    
    longToArray Tini, D3DColorXRGB(153, 153, 153)
    
    longToArray Onlines, D3DColorRGBA(152, 152, 246, 255)
    
    longToArray Candados, D3DColorXRGB(113, 116, 246)
    
End Sub

Public Function Directx_Initialize(ByVal Flags As CONST_D3DCREATEFLAGS) As Boolean

          '<EhHeader>
          On Error GoTo Directx_Initialize_Err

          '</EhHeader>
          
          ScreenWidth = frmMain.MainViewPic.ScaleWidth
          ScreenHeight = frmMain.MainViewPic.ScaleHeight

100       Directx_Initialize = False
    
          ' // Iniciacion del Directx8 //
    
          'Create the DirectX8 object
102       Set DirectX = New DirectX8
    
          'Create the Direct3D object
104       Set DirectD3D = DirectX.Direct3DCreate
    
          'Create helper class
106       Set DirectD3D8 = New D3DX8
    
          Dim DirectD3Ddm As D3DDISPLAYMODE
    
108       DirectD3D.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DirectD3Dcaps
110       DirectD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DirectD3Ddm
        
112       With DirectD3Dpp
114           .Windowed = True
116           .SwapEffect = D3DSWAPEFFECT_DISCARD
    
118           .BackBufferWidth = ScreenWidth
120           .BackBufferHeight = ScreenHeight
122           .BackBufferFormat = DirectD3Ddm.Format 'current display depth
              .hDeviceWindow = frmMain.MainViewPic.hwnd

          End With

124       If Flags = 0 Then
126           Flags = D3DCREATE_SOFTWARE_VERTEXPROCESSING

          End If

          'create device
128       Set DirectDevice = DirectD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.MainViewPic.hwnd, Flags, DirectD3Dpp)

130       Call D3DXMatrixOrthoOffCenterLH(Projection, 0, ScreenWidth, ScreenHeight, 0, -1#, 1#)
132       Call D3DXMatrixIdentity(View)
    
134       Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)
136       Call DirectDevice.SetTransform(D3DTS_VIEW, View)
    
138       Call Directx_RenderStates
    
140       Set Texture = New clsSurfaceManager
142       Set SpriteBatch = New clsBatch
    
144       Call Texture.Initialize(DirectD3D8, DirRecursos, 90)
146       Call SpriteBatch.Initialise(2000)
    
148       Call Directx_Init_FontSettings
150       Call Directx_Init_FontTextures
    
152       Call InitColours
    
154       With MainViewRect
156           .X2 = frmMain.MainViewPic.ScaleWidth
158           .Y2 = frmMain.MainViewPic.ScaleHeight

          End With
        
228       Directx_Initialize = True

          '<EhFooter>
          Exit Function

Directx_Initialize_Err:

          If Err.Number <> 0 Then
              Call LogError(Err.Description & vbCrLf & "in ARGENTUM.Directx_Initialize " & "at line " & Erl)

          End If

          '</EhFooter>
End Function

Public Sub Directx_DeInitialize()

    'Set no Textures to standard stage to avoid memory leak
    If Not DirectDevice Is Nothing Then
        DirectDevice.SetTexture 0, Nothing

    End If

    Set DirectX = Nothing
    Set DirectD3D = Nothing
    Set DirectD3D8 = Nothing
    Set DirectDevice = Nothing
    
    Set SpriteBatch = Nothing
    Set Texture = Nothing
    
    'Clear arrays
    Erase GrhData
    Erase BodyData
    Erase HeadData
    Erase FxData
    Erase WeaponAnimData
    Erase ShieldAnimData
    Erase CascoAnimData
    Erase MapData
    Erase CharList

    Exit Sub

End Sub

Private Sub Directx_RenderStates()
        
    'Set the render states
    With DirectDevice
        .SetVertexShader FVF
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, True
        .SetRenderState D3DRS_POINTSCALE_ENABLE, False
    
    End With
    
End Sub

Public Sub Directx_EndScene(ByRef RECT As D3DRECT, ByVal hwnd As Long)
    
    Call SpriteBatch.Flush
    
    Call DirectDevice.EndScene
    Call DirectDevice.Present(RECT, ByVal 0, hwnd, ByVal 0)

End Sub

Public Sub Text_Draw(ByVal Left As Long, ByVal Top As Long, ByVal Text As String, ByRef Color() As Long)
    
    Dim ColorFondo(0 To 3) As Long
    
    Call longToArray(ColorFondo, D3DColorRGBA(38, 39, 38, 255))

    Directx_Render_Text SpriteBatch, cfonts, Text, Left + 1, Top, ColorFondo
    Directx_Render_Text SpriteBatch, cfonts, Text, Left - 1, Top, ColorFondo
    Directx_Render_Text SpriteBatch, cfonts, Text, Left, Top + 1, ColorFondo
    Directx_Render_Text SpriteBatch, cfonts, Text, Left, Top - 1, ColorFondo
    Directx_Render_Text SpriteBatch, cfonts, Text, Left, Top, Color
    SpriteBatch.Flush

End Sub

'Private Function Es_Emoticon(ByVal ascii As Byte) As Boolean ' GSZAO
'
'    '*****************************************************************
'    'Emoticones by ^[GS]^
'    '*****************************************************************
'    Es_Emoticon = False
'
'    If (ascii = 129 Or ascii = 137 Or ascii = 141 Or ascii = 143 Or ascii = 144 Or ascii = 157 Or ascii = 160) Then
'        Es_Emoticon = True
'
'    End If
'
'End Function '

Private Sub Directx_Render_Text(ByRef Batch As clsBatch, _
    ByRef UseFont As CustomFont, _
    ByVal Text As String, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByRef Color() As Long)

    '*****************************************************************
    'Render text with a custom font
    '*****************************************************************
    Dim TempVA    As CharVA
    Dim tempstr() As String
    Dim Count     As Integer
    Dim ascii()   As Byte
    Dim i         As Long
    Dim j         As Long
    Dim yOffset   As Single
    
    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
    
    tempstr = Split(Text, Chr$(32))
    Text = vbNullString

    For i = 0 To UBound(tempstr)

'        If tempstr(i) = ":)" Or tempstr(i) = "=)" Then
'            tempstr(i) = Chr$(129)
'        ElseIf tempstr(i) = ":@" Or tempstr(i) = "=@" Then
'            tempstr(i) = Chr$(137)
'        ElseIf tempstr(i) = ":(" Or tempstr(i) = "=(" Then
'            tempstr(i) = Chr$(141)
'        ElseIf tempstr(i) = "^^" Or tempstr(i) = "^_^" Then
'            tempstr(i) = Chr$(143)
'        ElseIf tempstr(i) = ":D" Or tempstr(i) = "=D" Or tempstr(i) = ":d" Or tempstr(i) = "=d" Then
'            tempstr(i) = Chr$(144)
'        ElseIf tempstr(i) = "xD" Or tempstr(i) = "XD" Or tempstr(i) = "xd" Or tempstr(i) = "Xd" Then
'            tempstr(i) = Chr$(157)
'        ElseIf tempstr(i) = ":S" Or tempstr(i) = "=S" Or tempstr(i) = ":s" Or tempstr(i) = "=s" Then
'            tempstr(i) = Chr$(160)
'
'        End If

        Text = Text & Chr$(32) & tempstr(i)
    Next
    ' Made by ^[GS]^ for GSZAO
    
    'Get the text into arrays (split by vbCrLf)
    tempstr = Split(Text, vbCrLf)

    Batch.SetAlpha False
    
    'Set the texture
    Batch.SetTexture UseFont.Texture
    
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(tempstr)

        If Len(tempstr(i)) > 0 Then
            yOffset = i * UseFont.CharHeight
            Count = 0
        
            'Convert the characters to the ascii value
            ascii() = StrConv(tempstr(i), vbFromUnicode)
        
            'Loop through the characters
            For j = 1 To Len(tempstr(i))

                CopyMemory TempVA, UseFont.HeaderInfo.CharVA(ascii(j - 1)), 24 'this number represents the size of "CharVA" struct
                
                TempVA.X = X + Count
                TempVA.Y = Y + yOffset
                
                'Set the colors
'                If Es_Emoticon(ascii(j - 1)) Then    ' GSZAO los colores no afectan a los emoticones!
'                    If (ascii(j - 1) <> 157) Then Count = Count + 8   ' Los emoticones tienen tamaño propio (despues hay que cargarlos "correctamente" para evitar hacer esto)
'
'                End If
             
                Batch.Draw TempVA.X, TempVA.Y, TempVA.w, TempVA.H, Color, TempVA.Tx1, TempVA.Ty1, TempVA.Tx2, TempVA.Ty2

                'Shift over the the position to render the next character
                Count = Count + UseFont.HeaderInfo.CharWidth(ascii(j - 1))
                
            Next j
            
        End If

    Next i

End Sub

Public Function Text_GetWidth(ByVal Text As String) As Integer

    '***************************************************
    'Returns the width of text
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetTextWidth
    '***************************************************
    Dim i As Long

    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
    
    'Loop through the text
    For i = 1 To Len(Text)
        
        'Add up the stored character widths
        Text_GetWidth = Text_GetWidth + cfonts.HeaderInfo.CharWidth(Asc(mid$(Text, i, 1)))
        
    Next i

End Function

Private Sub Directx_Init_FontTextures()

    '*****************************************************************
    'Init the custom font textures
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontTextures
    '*****************************************************************

    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    '*** Default font ***
    
    'Set the texture
    Set cfonts.Texture = DirectD3D8.CreateTextureFromFileEx(DirectDevice, DirFont & "texdefault.png", D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, _
        D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFF000000, ByVal 0, ByVal 0)
            
    Dim Surface_Desc As D3DSURFACE_DESC

    cfonts.Texture.GetLevelDesc 0, Surface_Desc
    
    'Store the size of the texture
    cfonts.TextureSize.X = Surface_Desc.Width
    cfonts.TextureSize.Y = Surface_Desc.Height
    

End Sub

Private Sub Directx_Init_FontSettings()

    '*****************************************************************
    'Init the custom font settings
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontSettings
    '*****************************************************************
    Dim FileNum  As Byte
    Dim LoopChar As Long
    Dim Row      As Single
    Dim u        As Single
    Dim v        As Single

    '*** Default font ***

    'Load the header information
    FileNum = FreeFile
    Open DirFont & "texdefault.dat" For Binary As #FileNum
    Get #FileNum, , cfonts.HeaderInfo
    Close #FileNum
    
    'Calculate some common values
    cfonts.CharHeight = cfonts.HeaderInfo.CellHeight - 4
    cfonts.RowPitch = cfonts.HeaderInfo.BitmapWidth \ cfonts.HeaderInfo.CellWidth
    cfonts.ColFactor = cfonts.HeaderInfo.CellWidth / cfonts.HeaderInfo.BitmapWidth
    cfonts.RowFactor = cfonts.HeaderInfo.CellHeight / cfonts.HeaderInfo.BitmapHeight
    
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - cfonts.HeaderInfo.BaseCharOffset) \ cfonts.RowPitch
        u = ((LoopChar - cfonts.HeaderInfo.BaseCharOffset) - (Row * cfonts.RowPitch)) * cfonts.ColFactor
        v = Row * cfonts.RowFactor

        'Set the verticies
        With cfonts.HeaderInfo.CharVA(LoopChar)
            .X = 0
            .Y = 0
            .w = cfonts.HeaderInfo.CellWidth
            .H = cfonts.HeaderInfo.CellHeight
            .Tx1 = u
            .Ty1 = v
            .Tx2 = u + cfonts.ColFactor
            .Ty2 = v + cfonts.RowFactor

        End With
        
    Next LoopChar
    
End Sub

Public Sub Directx_Render_Texture(ByVal fileIndex As Long, _
          ByVal X As Integer, _
          ByVal Y As Integer, _
          ByVal Height As Integer, _
          ByVal Width As Integer, _
          ByVal sX As Integer, _
          ByVal sY As Integer, _
          ByRef Color() As Long, _
          Optional ByVal Angle As Single = 0, _
          Optional ByVal AlphaB As Byte = 0)

          '<EhHeader>
          On Error GoTo Directx_Render_Texture_Err

          '</EhHeader>

          Dim TexSurface As Direct3DTexture8
          Dim TexWidth   As Integer, TexHeight As Integer
 
100       Set TexSurface = Texture.Surface(fileIndex, TexWidth, TexHeight)

102       With SpriteBatch

104           Call .SetAlpha(AlphaB)
    
              '// Seteamos la textura
106           Call .SetTexture(TexSurface)

108           If TexWidth <> 0 And TexHeight <> 0 Then
110               Call .Draw(X, Y, Width, Height, Color, sX / TexWidth, sY / TexHeight, (sX + Width) / TexWidth, (sY + Height) / TexHeight, Angle)
              Else
112               Call .Draw(X, Y, TexWidth, TexHeight, Color, , , , , Angle)

              End If
  
          End With
  
Directx_Render_Texture_Err:

          If Err.Number <> 0 Then
              LogError Err.Description & vbCrLf & "in Directx_Render_Texture " & "at line " & Erl

              '</EhFooter>
          End If

End Sub

Public Sub Directx_Renderer()

    If EngineRun Then
    
        'Check if we have the device
        'If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
    
        Call DirectDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0)

        SpriteBatch.Begin

        'Sólo dibujamos si la ventana no está minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call ShowNextFrame
            
            'Play ambient sounds
            Call RenderSounds
            
            Call CheckKeys

        End If
        
        'Limit FPS to 100 (an easy number higher than monitor's vertical refresh rates)
        While (GetTickCount - FpsLastCheck) \ 6 < FramesPerSecCounter

            Sleep 5
        Wend

        'FPS update
        If FpsLastCheck + 1000 < GetTickCount Then
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            FpsLastCheck = GetTickCount
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1

        End If
    
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
        
        SpriteBatch.Finish

    End If

End Sub

Public Sub Draw_Box(ByVal X As Integer, ByVal Y As Integer, ByVal w As Integer, ByVal H As Integer, ByRef BackgroundColor() As Long)
    
    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.Draw(X, Y, w, H, BackgroundColor)
     
End Sub

Public Sub ColorToArray(ByRef TempArray() As Long, ByRef SetArray() As Long)

    TempArray(0) = SetArray(0)
    TempArray(1) = SetArray(1)
    TempArray(2) = SetArray(2)
    TempArray(3) = SetArray(3)

End Sub

Public Sub longToArray(ByRef TempArray() As Long, ByVal SetArray As Long)

    TempArray(0) = SetArray
    TempArray(1) = SetArray
    TempArray(2) = SetArray
    TempArray(3) = SetArray

End Sub

'##############################################
'############## PARTICULAS ORE ################
'##############################################

Private Function Particle_Group_Next_Open() As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim LooPC As Long
   
    If Particle_Group_Last = 0 Then ' Parche, testear.
        Particle_Group_Next_Open = 1
        Exit Function

    End If
   
    LooPC = 1

    Do Until Particle_Group_List(LooPC).Active = False

        If LooPC = Particle_Group_Last Then
            Particle_Group_Next_Open = Particle_Group_Last + 1
            Exit Function

        End If

        LooPC = LooPC + 1
    Loop
   
    Particle_Group_Next_Open = LooPC
    Exit Function
    
ErrorHandler:
    Particle_Group_Next_Open = 1

End Function

Public Function Particle_Group_Create(ByVal map_x As Integer, _
    ByVal map_y As Integer, _
    ByRef Grh_Index_List() As Long, _
    ByRef Rgb_List() As Long, _
    Optional ByVal Particle_Count As Long = 20, _
    Optional ByVal Stream_Type As Long = 1, _
    Optional ByVal alpha_blend As Boolean, _
    Optional ByVal alive_counter As Long = -1, _
    Optional ByVal Frame_Speed As Single = 0.5, _
    Optional ByVal id As Long, _
    Optional ByVal X1 As Integer, _
    Optional ByVal Y1 As Integer, _
    Optional ByVal Angle As Integer, _
    Optional ByVal vecx1 As Integer, _
    Optional ByVal vecx2 As Integer, _
    Optional ByVal vecy1 As Integer, _
    Optional ByVal vecy2 As Integer, _
    Optional ByVal life1 As Integer, _
    Optional ByVal life2 As Integer, _
    Optional ByVal fric As Integer, _
    Optional ByVal spin_speedL As Single, _
    Optional ByVal gravity As Boolean, _
    Optional ByVal grav_strength As Long, _
    Optional ByVal bounce_strength As Long, _
    Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean) As Long
   
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 12/15/2002
    'Returns the particle_group_index if successful, else 0
    '**************************************************************
    
    Dim ParticleOld As Integer
    
    ParticleOld = Map_Particle_Group_Get(map_x, map_y)

    If ParticleOld > 0 Then Call Particle_Group_Remove(ParticleOld)

    Particle_Group_Create = Particle_Group_Next_Open
        
    Call Particle_Group_Make(Particle_Group_Create, map_x, map_y, Particle_Count, Stream_Type, Grh_Index_List(), Rgb_List(), alpha_blend, _
        alive_counter, Frame_Speed, id, X1, Y1, Angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, _
        bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin)

End Function
 
Public Function Particle_Group_Remove(ByVal Particle_Group_Index As Long) As Boolean

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(Particle_Group_Index) Then
        Call Particle_Group_Destroy(Particle_Group_Index)
        Particle_Group_Remove = True

    End If

End Function
 
Public Function Particle_Group_Remove_All() As Boolean
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '*****************************************************************
    Dim Index As Long

    For Index = 1 To Particle_Group_Last

        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index

        End If

    Next Index
   
    Particle_Group_Remove_All = True

End Function
 
Public Function Particle_Group_Find(ByVal id As Long) As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    'Find the index related to the handle
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim LooPC As Long
   
    LooPC = 1

    Do Until Particle_Group_List(LooPC).id = id

        If LooPC = Particle_Group_Last Then
            Particle_Group_Find = 0
            Exit Function

        End If

        LooPC = LooPC + 1
    Loop
   
    Particle_Group_Find = LooPC
    Exit Function
    
ErrorHandler:
    Particle_Group_Find = 0

End Function
 
Private Sub Particle_Group_Make(ByVal ParticleIndex As Long, _
    ByVal map_x As Integer, _
    ByVal map_y As Integer, _
    ByVal Particle_Count As Long, _
    ByVal Stream_Type As Long, _
    ByRef Grh_Index_List() As Long, _
    ByRef Rgb_List() As Long, _
    Optional ByVal alpha_blend As Boolean, _
    Optional ByVal alive_counter As Long = -1, _
    Optional ByVal Frame_Speed As Single = 0.5, _
    Optional ByVal id As Long, _
    Optional ByVal X1 As Integer, _
    Optional ByVal Y1 As Integer, _
    Optional ByVal Angle As Integer, _
    Optional ByVal vecx1 As Integer, _
    Optional ByVal vecx2 As Integer, _
    Optional ByVal vecy1 As Integer, _
    Optional ByVal vecy2 As Integer, _
    Optional ByVal life1 As Integer, _
    Optional ByVal life2 As Integer, _
    Optional ByVal fric As Integer, _
    Optional ByVal spin_speedL As Single, _
    Optional ByVal gravity As Boolean, _
    Optional ByVal grav_strength As Long, _
    Optional ByVal bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Makes a new particle effect
    '*****************************************************************
    'Update array size
    If ParticleIndex > Particle_Group_Last Then
        Particle_Group_Last = ParticleIndex
        ReDim Preserve Particle_Group_List(1 To Particle_Group_Last) As Particle_Group

    End If

    Particle_Group_Count = Particle_Group_Count + 1
   
    With Particle_Group_List(ParticleIndex)
  
        'Make active
        .Active = True
   
        'Map pos
        If (map_x <> -1) And (map_y <> -1) Then
            .map_x = map_x
            .map_y = map_y
            
            'plot particle group on map
            MapData(map_x, map_y).Particle_Group = ParticleIndex
        
        End If
   
        'Grh list
        ReDim .Grh_Index_List(1 To UBound(Grh_Index_List))
        .Grh_Index_List() = Grh_Index_List()
        .Grh_Index_Count = UBound(Grh_Index_List)
   
        'Sets alive vars
        If alive_counter = -1 Then
            .alive_counter = -1
            .Never_Die = True
        Else
            .alive_counter = alive_counter
            .Never_Die = False

        End If
   
        'alpha blending
        .alpha_blend = alpha_blend
   
        'stream type
        .Stream_Type = Stream_Type
   
        'speed
        .Frame_Speed = Frame_Speed
   
        .X1 = X1
        .Y1 = Y1
        .X2 = X2
        .Y2 = Y2
        .Angle = Angle
        .vecx1 = vecx1
        .vecx2 = vecx2
        .vecy1 = vecy1
        .vecy2 = vecy2
        .life1 = life1
        .life2 = life2
        .fric = fric
        .spin = spin
        .spin_speedL = spin_speedL
        .spin_speedH = spin_speedH
        .gravity = gravity
        .grav_strength = grav_strength
        .bounce_strength = bounce_strength
        .XMove = XMove
        .YMove = YMove
        .move_x1 = move_x1
        .move_x2 = move_x2
        .move_y1 = move_y1
        .move_y2 = move_y2
   
        .Rgb_List(0) = Rgb_List(0)
        .Rgb_List(1) = Rgb_List(1)
        .Rgb_List(2) = Rgb_List(2)
        .Rgb_List(3) = Rgb_List(3)
   
        'create particle stream
        .Particle_Count = Particle_Count
        ReDim .Particle_Stream(1 To Particle_Count)
        
    End With

End Sub
 
Private Sub Particle_Render(ByRef Temp_Particle As Particle, _
    ByVal Screen_X As Integer, _
    ByVal Screen_Y As Integer, _
    ByVal Grh_Index As Long, _
    ByRef Rgb_List() As Long, _
    Optional ByVal alpha_blend As Boolean, _
    Optional ByVal no_move As Boolean, _
    Optional ByVal X1 As Integer, _
    Optional ByVal Y1 As Integer, _
    Optional ByVal Angle As Integer, _
    Optional ByVal vecx1 As Integer, _
    Optional ByVal vecx2 As Integer, _
    Optional ByVal vecy1 As Integer, _
    Optional ByVal vecy2 As Integer, _
    Optional ByVal life1 As Integer, _
    Optional ByVal life2 As Integer, _
    Optional ByVal fric As Integer, _
    Optional ByVal spin_speedL As Single, _
    Optional ByVal gravity As Boolean, _
    Optional ByVal grav_strength As Long, _
    Optional ByVal bounce_strength As Long, _
    Optional ByVal X2 As Integer, _
    Optional ByVal Y2 As Integer, _
    Optional ByVal XMove As Boolean, _
    Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 4/24/2003
    '
    '**************************************************************
 
    If no_move = False Then
        If Temp_Particle.alive_counter = 0 Then
        
            InitGrh Temp_Particle.Grh, Grh_Index
            
            Temp_Particle.X = RandomNumber(X1, X2)
            Temp_Particle.Y = RandomNumber(Y1, Y2)
            Temp_Particle.vector_x = RandomNumber(vecx1, vecx2)
            Temp_Particle.vector_y = RandomNumber(vecy1, vecy2)
            Temp_Particle.Angle = Angle
            Temp_Particle.alive_counter = RandomNumber(life1, life2)
            Temp_Particle.friction = fric
        Else

            'Continue old particle
            'Do gravity
            If gravity = True Then
                Temp_Particle.vector_y = Temp_Particle.vector_y + grav_strength

                If Temp_Particle.Y > 0 Then
                    'bounce
                    Temp_Particle.vector_y = bounce_strength

                End If

            End If

            'Do rotation
            If spin = True Then Temp_Particle.Angle = Temp_Particle.Angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
            
            If Temp_Particle.Angle >= 360 Then
                Temp_Particle.Angle = 0

            End If
                               
            If XMove Then Temp_Particle.vector_x = RandomNumber(move_x1, move_x2)
            If YMove Then Temp_Particle.vector_y = RandomNumber(move_y1, move_y2)

        End If
 
        'Add in vector
        Temp_Particle.X = Temp_Particle.X + (Temp_Particle.vector_x \ Temp_Particle.friction)
        Temp_Particle.Y = Temp_Particle.Y + (Temp_Particle.vector_y \ Temp_Particle.friction)
   
        'decrement counter
        Temp_Particle.alive_counter = Temp_Particle.alive_counter - 1

    End If
    
    'Draw it
    If Temp_Particle.Grh.GrhIndex Then
        DrawGrhtoSurface Temp_Particle.Grh, Temp_Particle.X + Screen_X, Temp_Particle.Y + Screen_Y, 1, 1, Rgb_List(), 1, Temp_Particle.Angle, _
            alpha_blend

    End If

End Sub
 
Private Sub Particle_Group_Render(ByVal ParticleIndex As Long, ByVal Screen_X As Integer, ByVal Screen_Y As Integer)
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 12/15/2002
    'Renders a particle stream at a paticular screen point
    '*****************************************************************
    Dim LooPC            As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move          As Boolean
    
    With Particle_Group_List(ParticleIndex)

        'Set colors
        If UserMinHP = 0 Then
            temp_rgb(0) = D3DColorARGB(.alpha_blend, 255, 255, 255)
            temp_rgb(1) = D3DColorARGB(.alpha_blend, 255, 255, 255)
            temp_rgb(2) = D3DColorARGB(.alpha_blend, 255, 255, 255)
            temp_rgb(3) = D3DColorARGB(.alpha_blend, 255, 255, 255)
        Else
            temp_rgb(0) = .Rgb_List(0)
            temp_rgb(1) = .Rgb_List(1)
            temp_rgb(2) = .Rgb_List(2)
            temp_rgb(3) = .Rgb_List(3)

        End If
       
        If .alive_counter Then
   
            'See if it is time to move a particle
            .Frame_Counter = .Frame_Counter + timerTicksPerFrame

            If .Frame_Counter > .Frame_Speed Then
                .Frame_Counter = 0
                no_move = False
            Else
                no_move = True

            End If
   
            'If it's still alive render all the particles inside
            For LooPC = 1 To .Particle_Count
       
                'Render particle
                Call Particle_Render(.Particle_Stream(LooPC), Screen_X, Screen_Y, .Grh_Index_List(Round(RandomNumber(1, .Grh_Index_Count), 0)), _
                    temp_rgb(), .alpha_blend, no_move, .X1, .Y1, .Angle, .vecx1, .vecx2, .vecy1, .vecy2, .life1, .life2, .fric, .spin_speedL, _
                    .gravity, .grav_strength, .bounce_strength, .X2, .Y2, .XMove, .move_x1, .move_x2, .move_y1, .move_y2, .YMove, .spin_speedH, _
                    .spin)
                           
            Next LooPC
       
            If no_move = False Then

                'Update the group alive counter
                If .Never_Die = False Then
                    .alive_counter = .alive_counter - 1

                End If

            End If
   
        Else
            'If it's dead destroy it
            .Particle_Count = .Particle_Count - 1

            If .Particle_Count <= 0 Then Particle_Group_Destroy ParticleIndex

        End If

    End With

End Sub
 
Public Function Particle_Type_Get(ByVal particle_index As Long) As Long

    '*****************************************************************
    'Author: Juan Martín Sotuyo Dodero ([URL='mailto:juansotuyo@hotmail.com']juansotuyo@hotmail.com[/URL])
    'Last Modify Date: 8/27/2003
    'Returns the stream type of a particle stream
    '*****************************************************************
    If Particle_Group_Check(particle_index) Then
        Particle_Type_Get = Particle_Group_List(particle_index).Stream_Type

    End If

End Function
 
Private Function Particle_Group_Check(ByVal Particle_Group_Index As Long) As Boolean
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    
    'check index
    If Particle_Group_Index > 0 And Particle_Group_Index <= Particle_Group_Last Then
        If Particle_Group_Index <> bSecondaryAmbient Then
            If Particle_Group_List(Particle_Group_Index).Active Then
                Particle_Group_Check = True

            End If

        End If

    End If

End Function
 
Public Function Particle_Group_Map_Pos_Set(ByVal Particle_Group_Index As Long, ByVal map_x As Long, ByVal map_y As Long) As Boolean
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 5/27/2003
    'Returns true if successful, else false
    '**************************************************************
    
    'Make sure it's a legal index
    If Particle_Group_Check(Particle_Group_Index) Then

        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then
            'Move it
            Particle_Group_List(Particle_Group_Index).map_x = map_x
            Particle_Group_List(Particle_Group_Index).map_y = map_y
   
            Particle_Group_Map_Pos_Set = True

        End If

    End If

End Function

Private Sub Particle_Group_Destroy(ByVal ParticleIndex As Long)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    Dim temp As Particle_Group
    Dim i    As Long

    With Particle_Group_List(ParticleIndex)

        If .map_x > 0 And .map_y > 0 Then
            MapData(.map_x, .map_y).Particle_Group = 0
        ElseIf .charindex Then

            If Char_Check(.charindex) Then

                For i = 1 To CharList(.charindex).Particle_Count

                    If CharList(.charindex).Particle_Group(i) = ParticleIndex Then
                        CharList(.charindex).Particle_Group(i) = 0
                        Exit For

                    End If

                Next i

            End If

        End If

    End With

    Particle_Group_List(ParticleIndex) = temp
           
    'Update array size
    If ParticleIndex = Particle_Group_Last Then

        Do Until Particle_Group_List(Particle_Group_Last).Active
            Particle_Group_Last = Particle_Group_Last - 1

            If Particle_Group_Last = 0 Then
                Particle_Group_Count = 0
                Exit Sub

            End If

        Loop
        ReDim Preserve Particle_Group_List(1 To Particle_Group_Last)

    End If

    Particle_Group_Count = Particle_Group_Count - 1

End Sub
 
Private Sub Char_Particle_Group_Make(ByVal Particle_Group_Index As Long, _
    ByVal charindex As Integer, _
    ByVal particle_CharIndex As Integer, _
    ByVal Particle_Count As Long, _
    ByVal Stream_Type As Long, _
    ByRef Grh_Index_List() As Long, _
    ByRef Rgb_List() As Long, _
    Optional ByVal alpha_blend As Boolean, _
    Optional ByVal alive_counter As Long = -1, _
    Optional ByVal Frame_Speed As Single = 0.5, _
    Optional ByVal id As Long, _
    Optional ByVal X1 As Integer, _
    Optional ByVal Y1 As Integer, _
    Optional ByVal Angle As Integer, _
    Optional ByVal vecx1 As Integer, _
    Optional ByVal vecx2 As Integer, _
    Optional ByVal vecy1 As Integer, _
    Optional ByVal vecy2 As Integer, _
    Optional ByVal life1 As Integer, _
    Optional ByVal life2 As Integer, _
    Optional ByVal fric As Integer, _
    Optional ByVal spin_speedL As Single, _
    Optional ByVal gravity As Boolean, _
    Optional ByVal grav_strength As Long, _
    Optional ByVal bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)
                               
    '*****************************************************************
    'Author: Aaron Perkins
    'Modified by: Ryan Cain (Onezero)
    'Last Modify Date: 5/15/2003
    'Makes a new particle effect
    'Modified by Juan Martín Sotuyo Dodero
    '*****************************************************************
    'Update array size
    If Particle_Group_Index > Particle_Group_Last Then
        Particle_Group_Last = Particle_Group_Index
        ReDim Preserve Particle_Group_List(1 To Particle_Group_Last)

    End If

    Particle_Group_Count = Particle_Group_Count + 1
   
    'Make active
    Particle_Group_List(Particle_Group_Index).Active = True
   
    'Char index
    Particle_Group_List(Particle_Group_Index).charindex = charindex
   
    'Grh list
    ReDim Particle_Group_List(Particle_Group_Index).Grh_Index_List(1 To UBound(Grh_Index_List))
    Particle_Group_List(Particle_Group_Index).Grh_Index_List() = Grh_Index_List()
    Particle_Group_List(Particle_Group_Index).Grh_Index_Count = UBound(Grh_Index_List)
   
    'Sets alive vars
    If alive_counter = -1 Then
        Particle_Group_List(Particle_Group_Index).alive_counter = -1
        Particle_Group_List(Particle_Group_Index).Never_Die = True
    Else
        Particle_Group_List(Particle_Group_Index).alive_counter = alive_counter
        Particle_Group_List(Particle_Group_Index).Never_Die = False

    End If
   
    'alpha blending
    Particle_Group_List(Particle_Group_Index).alpha_blend = alpha_blend
   
    'stream type
    Particle_Group_List(Particle_Group_Index).Stream_Type = Stream_Type
   
    'speed
    Particle_Group_List(Particle_Group_Index).Frame_Speed = Frame_Speed
   
    Particle_Group_List(Particle_Group_Index).X1 = X1
    Particle_Group_List(Particle_Group_Index).Y1 = Y1
    Particle_Group_List(Particle_Group_Index).X2 = X2
    Particle_Group_List(Particle_Group_Index).Y2 = Y2
    Particle_Group_List(Particle_Group_Index).Angle = Angle
    Particle_Group_List(Particle_Group_Index).vecx1 = vecx1
    Particle_Group_List(Particle_Group_Index).vecx2 = vecx2
    Particle_Group_List(Particle_Group_Index).vecy1 = vecy1
    Particle_Group_List(Particle_Group_Index).vecy2 = vecy2
    Particle_Group_List(Particle_Group_Index).life1 = life1
    Particle_Group_List(Particle_Group_Index).life2 = life2
    Particle_Group_List(Particle_Group_Index).fric = fric
    Particle_Group_List(Particle_Group_Index).spin = spin
    Particle_Group_List(Particle_Group_Index).spin_speedL = spin_speedL
    Particle_Group_List(Particle_Group_Index).spin_speedH = spin_speedH
    Particle_Group_List(Particle_Group_Index).gravity = gravity
    Particle_Group_List(Particle_Group_Index).grav_strength = grav_strength
    Particle_Group_List(Particle_Group_Index).bounce_strength = bounce_strength
    Particle_Group_List(Particle_Group_Index).XMove = XMove
    Particle_Group_List(Particle_Group_Index).YMove = YMove
    Particle_Group_List(Particle_Group_Index).move_x1 = move_x1
    Particle_Group_List(Particle_Group_Index).move_x2 = move_x2
    Particle_Group_List(Particle_Group_Index).move_y1 = move_y1
    Particle_Group_List(Particle_Group_Index).move_y2 = move_y2
   
    'color
    Particle_Group_List(Particle_Group_Index).Rgb_List(0) = Rgb_List(0)
    Particle_Group_List(Particle_Group_Index).Rgb_List(1) = Rgb_List(1)
    Particle_Group_List(Particle_Group_Index).Rgb_List(2) = Rgb_List(2)
    Particle_Group_List(Particle_Group_Index).Rgb_List(3) = Rgb_List(3)
    
    'handle
    Particle_Group_List(Particle_Group_Index).id = id
   
    'create particle stream
    Particle_Group_List(Particle_Group_Index).Particle_Count = Particle_Count
    ReDim Particle_Group_List(Particle_Group_Index).Particle_Stream(1 To Particle_Count)
   
    'plot particle group on char
    CharList(charindex).Particle_Group(particle_CharIndex) = Particle_Group_Index

End Sub
 
Public Function Char_Particle_Group_Create(ByVal charindex As Integer, _
    ByRef Grh_Index_List() As Long, _
    ByRef Rgb_List() As Long, _
    Optional ByVal Particle_Count As Long = 20, _
    Optional ByVal Stream_Type As Long = 1, _
    Optional ByVal alpha_blend As Boolean, _
    Optional ByVal alive_counter As Long = -1, _
    Optional ByVal Frame_Speed As Single = 0.5, _
    Optional ByVal id As Long, _
    Optional ByVal X1 As Integer, _
    Optional ByVal Y1 As Integer, _
    Optional ByVal Angle As Integer, _
    Optional ByVal vecx1 As Integer, _
    Optional ByVal vecx2 As Integer, _
    Optional ByVal vecy1 As Integer, _
    Optional ByVal vecy2 As Integer, _
    Optional ByVal life1 As Integer, _
    Optional ByVal life2 As Integer, _
    Optional ByVal fric As Integer, _
    Optional ByVal spin_speedL As Single, _
    Optional ByVal gravity As Boolean, _
    Optional ByVal grav_strength As Long, _
    Optional ByVal bounce_strength As Long, _
    Optional ByVal X2 As Integer, _
    Optional ByVal Y2 As Integer, Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean) As Long
    '**************************************************************
    'Author: Augusto José Rando
    '**************************************************************
    Dim char_part_free_index As Integer
   
    'If Char_Particle_Group_Find(CharIndex, stream_type) Then Exit Function ' hay que ver si dejar o sacar esto...
    If Not Char_Check(charindex) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(charindex)
   
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, charindex, char_part_free_index, Particle_Count, Stream_Type, Grh_Index_List(), _
            Rgb_List(), alpha_blend, alive_counter, Frame_Speed, id, X1, Y1, Angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, _
            spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin
       
    End If
 
End Function
 
Private Function Char_Particle_Group_Find(ByVal charindex As Integer, ByVal Stream_Type As Long) As Integer
    '*****************************************************************
    'Author: Augusto José Rando
    'Modified: returns slot or -1
    '*****************************************************************
 
    Dim i As Long
 
    For i = 1 To CharList(charindex).Particle_Count

        If Particle_Group_List(CharList(charindex).Particle_Group(i)).Stream_Type = Stream_Type Then
            Char_Particle_Group_Find = CharList(charindex).Particle_Group(i)
            Exit Function

        End If

    Next i
 
    Char_Particle_Group_Find = -1
 
End Function
 
Private Function Char_Particle_Group_Next_Open(ByVal charindex As Integer) As Integer

    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim LooPC As Long
   
    LooPC = 1

    Do Until CharList(charindex).Particle_Group(LooPC) = 0

        If LooPC = CharList(charindex).Particle_Count Then
            Char_Particle_Group_Next_Open = CharList(charindex).Particle_Count + 1
            CharList(charindex).Particle_Count = Char_Particle_Group_Next_Open
            ReDim Preserve CharList(charindex).Particle_Group(1 To Char_Particle_Group_Next_Open) As Long
            Exit Function

        End If

        LooPC = LooPC + 1
    Loop
   
    Char_Particle_Group_Next_Open = LooPC
 
    Exit Function
 
ErrorHandler:
    CharList(charindex).Particle_Count = 1
    ReDim CharList(charindex).Particle_Group(1 To 1) As Long
    Char_Particle_Group_Next_Open = 1
 
End Function
 
Public Function Char_Particle_Group_Remove(ByVal charindex As Integer, ByVal Stream_Type As Long)
    '**************************************************************
    'Author: Augusto José Rando
    '**************************************************************
    Dim char_part_index As Integer
   
    If Char_Check(charindex) Then
        char_part_index = Char_Particle_Group_Find(charindex, Stream_Type)

        If char_part_index = -1 Then Exit Function
        Call Particle_Group_Remove(char_part_index)

    End If
 
End Function
 
Public Function Char_Particle_Group_Remove_All(ByVal charindex As Integer)
    '**************************************************************
    'Author: Augusto José Rando
    '**************************************************************
    Dim i As Long
   
    If Char_Check(charindex) Then

        For i = 1 To CharList(charindex).Particle_Count

            If CharList(charindex).Particle_Group(i) <> 0 Then
                Call Particle_Group_Remove(CharList(charindex).Particle_Group(i))

            End If

        Next i

    End If
   
End Function
 
Public Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
 
    If InMapBounds(map_x, map_y) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).Particle_Group
    Else
        Map_Particle_Group_Get = 0

    End If

End Function

Private Function Char_Check(ByVal charindex As Integer) As Boolean

    'check CharIndex
    If charindex > 0 And charindex <= LastChar Then
        Char_Check = (CharList(charindex).Active = 1)

    End If

End Function
