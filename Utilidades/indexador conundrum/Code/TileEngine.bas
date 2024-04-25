Attribute VB_Name = "Mod_TileEngine"
Option Explicit

Public grhCount As Long
Public fileVersion As Long

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'    C       O       N       S      T
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'Map sizes in tiles
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

'bltbit constant
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'    T       I      P      O      S
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'Encabezado bmp
Private Type BITMAPFILEHEADER

    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long

End Type

Public Type ErroresGrh

    colores(0 To 9) As Long
    ErrorCritico As Boolean
    EsAnimacion As Boolean

End Type

'Info del encabezado del bmp
Private Type BITMAPINFOHEADER

    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long

End Type

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

'Contiene info acerca de donde se puede encontrar un grh
'tamaño y animacion
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

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    
    GrhIndex     As Long
    FrameCounter As Single
    Speed        As Single
    Started      As Byte
    Loops        As Integer
    
End Type

'Lista de cuerpos
Public Type BodyData

    Walk(1 To 4) As Grh
    HeadOffset As Position

End Type

'Lista de cabezas
Public Type HeadData

    Head(1 To 4) As Grh

End Type

'Lista de las animaciones de las armas
Type WeaponAnimData

    WeaponWalk(1 To 4) As Grh
    '[ANIM ATAK]
    WeaponAttack As Byte

End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData

    ShieldWalk(1 To 4) As Grh

End Type

'Lista de cuerpos
Public Type FxData

    Fx As Grh
    OffsetX As Long
    OffsetY As Long

End Type

'Apariencia del personaje
Public Type Char

    Active As Byte
    Heading As Byte ' As E_Heading ?
    pos As Position
    
    iHead As Integer
    Ibody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Espalda As HeadData
    Botas As HeadData
    
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    Fx As Integer
    FxLoopTimes As Integer
    Criminal As Byte
    
    Nombre As String
    
    Moving As Byte
    MoveOffset As Position
    ServerIndex As Integer
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
    ClanID As Integer
    ClanName As String
    ColorName As Long

End Type

'Info de un objeto
Public Type Obj

    OBJIndex As Integer
    Amount As Integer

End Type

'Tipo de las celdas del mapa
Public Type MapBlock

    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer

End Type

'Info de cada mapa
Public Type MapInfo

    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    
    'ME Only
    Changed As Byte

End Type

Public Type tIndiceCabeza

    Head(1 To 4) As Long

End Type

Public Type tIndiceCuerpo

    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer

End Type

Public Type tIndiceFx

    Animacion As Long
    OffsetX As Integer
    OffsetY As Integer

End Type

Public IniPath              As String
Public MapPath              As String

'Bordes del mapa
Public MinXBorder           As Byte
Public MaxXBorder           As Byte
Public MinYBorder           As Byte
Public MaxYBorder           As Byte

'Status del user
Public CurMap               As Integer 'Mapa actual
Public UserIndex            As Integer
Public MiClanID             As Integer
Public UserMoving           As Byte
Public UserBody             As Integer
Public UserHead             As Integer
Public UserPos              As Position 'Posicion
Public AddtoUserPos         As Position 'Si se mueve
Public UserCharIndex        As Integer

Public UserMaxAGU           As Integer
Public UserMinAGU           As Integer
Public UserMaxHAM           As Integer
Public UserMinHAM           As Integer

Public EngineRun            As Boolean
Public FramesPerSec         As Integer
Public FramesPerSecCounter  As Long

'Tamaño del la vista en Tiles
Public WindowTileWidth      As Integer
Public WindowTileHeight     As Integer

'Offset del desde 0,0 del main view
Public MainViewTop          As Integer
Public MainViewLeft         As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize       As Integer

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd      As Long

'Tamaño de los tiles en pixels
Public TilePixelHeight      As Integer
Public TilePixelWidth       As Integer

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Totales?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public NumBodies            As Integer
Public numfxs               As Integer

Public NumChars             As Integer
Public LastChar             As Integer
Public NumWeaponAnims       As Integer
Public NumShieldAnims       As Integer

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public LastTime             As Long 'Para controlar la velocidad

'[CODE]:MatuX'
Public MainDestRect         As RECT
'[END]'
Public MainViewRect         As RECT
Public BackBufferRect       As RECT

Public MainViewWidth        As Integer
Public MainViewHeight       As Integer

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData()            As GrhData 'Guarda todos los grh
Public BodyData()           As BodyData
Public HeadData()           As HeadData
Public FxData()             As FxData
Public WeaponAnimData()     As WeaponAnimData
Public ShieldAnimData()     As ShieldAnimData
Public CascoAnimData()      As HeadData
Public BotasAnimData()      As HeadData
Public EspaldaAnimData()    As HeadData
Public Grh()                As Grh 'Animaciones publicas
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData()            As MapBlock ' Mapa
Public MapInfo              As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public PuedeVerClan         As Boolean
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Usuarios?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'
'epa ;)
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿API?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'       [CODE 000]: MatuX
'
Public bRain                As Boolean 'está raineando?
Public bTecho               As Boolean 'hay techo?
Public brstTick             As Long

Private RLluvia(7)          As RECT  'RECT de la lluvia
Private iFrameIndex         As Byte  'Frame actual de la LL
Private llTick              As Long  'Contador
Private LTLluvia(4)         As Integer

Public charlist(1 To 10000) As Char

'[CODE 001]:MatuX
Public Enum PlayLoop

    plNone = 0
    plLluviain = 1
    plLluviaout = 2

End Enum

'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Sub CargarArrayLluvia()

    On Error Resume Next

    Dim N  As Integer, i As Integer
    Dim Nu As Integer

    N = FreeFile
    Open App.Path & "\" & CarpetaDeInis & "\fk.ind" For Binary Access Read As #N

    'cabecera
    Get #N, , MiCabecera

    'num de cabezas
    Get #N, , Nu
    
    Nu = 300

    'Resize array
    ReDim bLluvia(1 To Nu) As Byte

    For i = 1 To Nu
        Get #N, , bLluvia(i)
    Next i

    Close #N

End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)

    charlist(CharIndex).Active = 0
    charlist(CharIndex).Criminal = 0
    charlist(CharIndex).Fx = 0
    charlist(CharIndex).FxLoopTimes = 0
    charlist(CharIndex).invisible = False

    #If SeguridadAlkon Then
        Call MI(CualMI).ResetInvisible(CharIndex)
    #End If

    charlist(CharIndex).Moving = 0
    charlist(CharIndex).muerto = False
    charlist(CharIndex).Nombre = ""
    charlist(CharIndex).ClanName = ""
    charlist(CharIndex).ClanID = 0
    charlist(CharIndex).pie = False
    charlist(CharIndex).pos.X = 0
    charlist(CharIndex).pos.Y = 0
    charlist(CharIndex).UsandoArma = False

End Sub

Sub EraseChar(ByVal CharIndex As Integer)

    On Error Resume Next

    '*****************************************************************
    'Erases a character from CharList and map
    '*****************************************************************

    charlist(CharIndex).Active = 0

    'Update lastchar
    If CharIndex = LastChar Then

        Do Until charlist(LastChar).Active = 1
            LastChar = LastChar - 1

            If LastChar = 0 Then Exit Do
        Loop

    End If

    MapData(charlist(CharIndex).pos.X, charlist(CharIndex).pos.Y).CharIndex = 0

    Call ResetCharInfo(CharIndex)

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
        Grh.Loops = -1
    Else
        Grh.Loops = 0
    
    End If

    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed

End Sub

Sub DrawGrhtoHdc(hWnd As Long, _
                 hDC As Long, _
                 FileNum As Long, _
                 sourceRect As RECT, _
                 destRect As RECT)

    If FileNum <= 0 Then Exit Sub

    On Error Resume Next

    SecundaryClipper.SetHWnd hWnd
    
    SurfaceDB.Surface(FileNum).BltToDC hDC, sourceRect, destRect
    
End Sub

Sub LoadGraphics()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero - complete rewrite
    'Last Modify Date: 11/03/2006
    'Initializes the SurfaceDB and sets up the rain rects
    '**************************************************************
    'New surface manager :D
    Call SurfaceDB.Initialize(DirectDraw, True, App.Path & "\" & CarpetaGraficos & "\")
          
    'Set up te rain rects
    RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
    RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
    RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
    RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128

    RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
    RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
    RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
    RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
    
    'We are done!
End Sub

'[END]'
Function InitTileEngine(ByRef setDisplayFormhWnd As Long, _
                        setMainViewTop As Integer, _
                        setMainViewLeft As Integer, _
                        setTilePixelHeight As Integer, _
                        setTilePixelWidth As Integer, _
                        setWindowTileHeight As Integer, _
                        setWindowTileWidth As Integer, _
                        setTileBufferSize As Integer) As Boolean
    '*****************************************************************
    'InitEngine
    '*****************************************************************
    Dim SurfaceDesc As DDSURFACEDESC2
    Dim ddck        As DDCOLORKEY

    IniPath = App.Path & "\" & CarpetaDeInis & "\"

    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder

    'Fill startup variables

    DisplayFormhWnd = setDisplayFormhWnd
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize

    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

    MainViewWidth = (TilePixelWidth * WindowTileWidth)
    MainViewHeight = (TilePixelHeight * WindowTileHeight)

    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

    DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

    'Primary Surface
    ' Fill the surface description structure
    With SurfaceDesc
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE

    End With

    Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

    Set PrimaryClipper = DirectDraw.CreateClipper(0)
    PrimaryClipper.SetHWnd frmMain.hWnd
    PrimarySurface.SetClipper PrimaryClipper

    Set SecundaryClipper = DirectDraw.CreateClipper(0)

    With BackBufferRect
        .Left = 0
        .Top = 0
        .Right = TilePixelWidth * (WindowTileWidth + 2 * TileBufferSize)
        .Bottom = TilePixelHeight * (WindowTileHeight + 2 * TileBufferSize)

    End With

    With SurfaceDesc
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH

        If True Then
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        Else
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY

        End If

        .lHeight = BackBufferRect.Bottom
        .lWidth = BackBufferRect.Right

    End With

    Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)

    ddck.low = 0
    ddck.high = 0
    BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck

    Call LoadGrhData

    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    'Call CargarEspalda
    'Call CargarBotas
    Call CargarFxs

    LTLluvia(0) = 224
    LTLluvia(1) = 352
    LTLluvia(2) = 480
    LTLluvia(3) = 608
    LTLluvia(4) = 736

    Call LoadGraphics

    InitTileEngine = True

End Function

Sub dibujapj(Surface As DirectDrawSurface7, Grh As GrhData)

    On Error Resume Next

    Dim r2          As RECT, auxr As RECT, auxr2 As RECT
    Dim iGrhIndex   As Long
    Dim SurfaceDesc As DDSURFACEDESC2

    If Grh.FileNum <= 0 Then Exit Sub

    SurfaceDB.Surface(Grh.FileNum).GetSurfaceDesc SurfaceDesc

    With r2
        .Left = Grh.sX
        .Top = Grh.sY

        If .Left + Grh.pixelWidth <= SurfaceDesc.lWidth Then
            .Right = .Left + Grh.pixelWidth
        Else
            .Right = SurfaceDesc.lWidth

        End If

        If .Top + Grh.pixelHeight <= SurfaceDesc.lHeight Then
            .Bottom = .Top + Grh.pixelHeight
        Else
            .Bottom = SurfaceDesc.lHeight

        End If
   
    End With

    With auxr
        .Left = 0
        .Top = 0
        .Right = Grh.pixelWidth
        .Bottom = Grh.pixelHeight

    End With

    auxr2 = auxr

    Surface.BltFast 0, 0, SurfaceDB.Surface(Grh.FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Surface.BltToDC frmMain.Visor.hDC, auxr2, auxr

End Sub

Sub dibujapjESpecial(Surface As DirectDrawSurface7, _
                     Grh As GrhData, _
                     ByVal X As Integer, _
                     ByVal Y As Integer)

    On Error Resume Next

    Dim r2          As RECT, auxr As RECT, auxr2 As RECT
    Dim iGrhIndex   As Long
    Dim SurfaceDesc As DDSURFACEDESC2

    If Grh.FileNum <= 0 Then Exit Sub

    SurfaceDB.Surface(Grh.FileNum).GetSurfaceDesc SurfaceDesc

    With r2
        .Left = Grh.sX
        .Top = Grh.sY

        If .Left + Grh.pixelWidth <= SurfaceDesc.lWidth Then
            .Right = .Left + Grh.pixelWidth
        Else
            .Right = SurfaceDesc.lWidth

        End If

        If .Top + Grh.pixelHeight <= SurfaceDesc.lHeight Then
            .Bottom = .Top + Grh.pixelHeight
        Else
            .Bottom = SurfaceDesc.lHeight

        End If

    End With

    With auxr
        .Left = 0
        .Top = 0
        .Right = Grh.pixelWidth
        .Bottom = Grh.pixelHeight

    End With

    auxr2 = auxr

    Surface.BltFast X, Y, SurfaceDB.Surface(Grh.FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

    'Surface.BltToDC frmMain.Visor.hDC, auxr2, auxr
End Sub

Sub dibujaBMP(Surface As DirectDrawSurface7, FileNum As Long)

    On Error Resume Next

    Dim r2            As RECT, auxr As RECT, auxr2 As RECT
    Dim r             As RECT
    Dim iGrhIndex     As Long
    Dim SurfaceDesc   As DDSURFACEDESC2
    Dim ddsd          As DDSURFACEDESC2
    Dim ddck          As DDCOLORKEY
    Dim surfacecuadro As DirectDrawSurface7
    Dim ii            As Long

    If FileNum <= 0 Then Exit Sub

    SurfaceDB.Surface(FileNum).GetSurfaceDesc SurfaceDesc

    With r2
        .Left = 0
        .Top = 0
        .Right = SurfaceDesc.lWidth
        .Bottom = SurfaceDesc.lHeight
   
    End With

    auxr = r2
    auxr2 = auxr

    If DibujarIndexaciones.activo Then
        ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        ddsd.lWidth = DibujarIndexaciones.Ancho
        ddsd.lHeight = DibujarIndexaciones.Alto
        Set surfacecuadro = DirectDraw.CreateSurface(ddsd)
        Call surfacecuadro.BltColorFill(r, vbBlack)
        surfacecuadro.SetForeColor vbGreen
        surfacecuadro.SetFillColor vbBlack
        Call surfacecuadro.DrawBox(0, 0, DibujarIndexaciones.Ancho, DibujarIndexaciones.Alto)
    
        ddck.high = 0
        ddck.low = 0
        Call surfacecuadro.SetColorKey(DDCKEY_SRCBLT, ddck)
        Call surfacecuadro.GetSurfaceDesc(ddsd)

    End If

    Surface.BltFast 0, 0, SurfaceDB.Surface(FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

    If DibujarIndexaciones.activo Then

        For ii = 1 To DibujarIndexaciones.Total
            Surface.BltFast DibujarIndexaciones.Inicios(ii).X, DibujarIndexaciones.Inicios(ii).Y, _
                    surfacecuadro, r, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Next ii

    End If

    Surface.BltToDC frmMain.Visor.hDC, auxr2, auxr

End Sub

Sub dibujarGrh2(ByRef Grh As GrhData)

    On Error Resume Next

    Dim r As RECT

    If DibujarFondo Then
        BackBufferSurface.BltColorFill r, ColorFondo
    Else
        BackBufferSurface.BltColorFill r, 0

    End If

    Call dibujapj(BackBufferSurface, Grh)
    '*************** *********************************** ***************

End Sub

Sub dibujarBMP2(ByRef Grh As Long)

    On Error Resume Next

    Dim r As RECT

    If DibujarFondo Then
        BackBufferSurface.BltColorFill r, ColorFondo
    Else
        BackBufferSurface.BltColorFill r, 0

    End If

    Call dibujaBMP(BackBufferSurface, Grh)
    '*************** *********************************** ***************
    'If dibujarindexacion Then
    
    'End If

End Sub

