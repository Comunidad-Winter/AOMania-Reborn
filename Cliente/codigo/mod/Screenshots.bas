Attribute VB_Name = "Screenshots"
Option Explicit

' ----==== GDIPlus Const ====----
Const GdiPlusVersion                        As Long = 1
Private Const EncoderParameterValueTypeLong As Long = 4
Private Const EncoderQuality                As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"

' ----==== Sonstige Types ====----
Public Enum MimeType

    JPG = 0
    GIF = 1
    PNG = 2
    BMP = 3

End Enum

Private Type PICTDESC

    cbSizeOfStruct As Long
    picType As Long
    hgdiobj As Long
    hPalOrXYExt As Long

End Type

Private Type IID

    data1 As Long
    data2 As Integer
    Data3 As Integer
    Data4(0 To 7)  As Byte

End Type

Private Type GUID

    data1 As Long
    data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte

End Type

' ----==== GDIPlus Types ====----
Private Type GDIPlusStartupInput

    GdiPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long

End Type

Private Type EncoderParameter

    GUID As GUID
    NumberOfValues As Long

    type As Long

    value As Long

End Type

Private Type EncoderParameters

    Count As Long
    Parameter(15) As EncoderParameter

End Type

Private Type ImageCodecInfo

    Clsid As GUID
    FormatID As GUID
    CodecNamePtr As Long
    DllNamePtr As Long
    FormatDescriptionPtr As Long
    FilenameExtensionPtr As Long
    MimeTypePtr As Long
    Flags As Long
    version As Long
    SigCount As Long
    SigSize As Long
    SigPatternPtr As Long
    SigMaskPtr As Long

End Type

' ----==== GDIPlus Enums ====----
Public Enum status 'GDI+ Status

    Ok = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
    ProfileNotFound = 21

End Enum

' ----==== GDI+ API Declarationen ====----
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, _
    ByRef lpInput As GDIPlusStartupInput, _
    Optional ByRef lpOutput As Any) As status

Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As status

Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As Long, ByRef BITMAP As Long) As status

Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, _
    ByVal FileName As Long, _
    ByRef clsidEncoder As GUID, _
    ByRef encoderParams As Any) As status

Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal BITMAP As Long, ByRef hbmReturn As Long, ByVal background As Long) As status

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, ByRef BITMAP As Long) As status

Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (ByRef numEncoders As Long, ByRef Size As Long) As status

Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, ByRef Encoders As Any) As status

Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As status

Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long

Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PICTDESC, riid As IID, ByVal fOwn As Boolean, lplpvObj As Object)

Private Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long

Private Declare Function lstrcpyW Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long

Private retStatus       As status
Private GdipToken       As Long
Private GdipInitialized As Boolean

Private Function StartUpGDIPlus(ByVal GdipVersion As Long) As status
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = GdipVersion
    StartUpGDIPlus = GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)

End Function

Private Function ShutdownGDIPlus() As status
    ShutdownGDIPlus = GdiplusShutdown(GdipToken)

End Function

Private Function Execute(ByVal lReturn As status) As status
    Dim lCurErr As status

    If lReturn = status.Ok Then
        lCurErr = status.Ok
    Else
        lCurErr = lReturn
        
    End If

    Execute = lCurErr

End Function

Public Function Convertir(ByVal Pic As StdPicture, _
    ByVal FileName As String, _
    Optional ByVal Quality As Long = 85, _
    Optional ByVal FileType As MimeType = JPG) As Boolean
    
    Dim retStatus As status
    Dim retval    As Boolean
    Dim lBitmap   As Long
    '// Variable para el MimeType
    Dim mimeT     As String
    
    Iniciar
    
    If GdipInitialized = False Then Exit Function
    ' Erzeugt eine GDI+ Bitmap vom StdPicture Handle -> lBitmap
    retStatus = Execute(GdipCreateBitmapFromHBITMAP(Pic.handle, 0, lBitmap))
    
    If retStatus = Ok Then
        
        Dim PicEncoder As GUID
        Dim tParams    As EncoderParameters
        
        '// Seleccion de casos para el MimeType
        Select Case FileType

            Case JPG
                mimeT = "image/jpeg"

            Case GIF
                mimeT = "image/gif"

            Case PNG
                mimeT = "image/png"

            Case BMP
                mimeT = "image/bmp"

        End Select
        
        '// Ermitteln der CLSID vom mimeType Encoder
        retval = GetEncoderClsid(mimeT, PicEncoder)

        If retval = True Then
              
            If Quality > 100 Then Quality = 100
            If Quality < 0 Then Quality = 0
              
            ' Initialisieren der Encoderparameter
            tParams.Count = 1

            With tParams.Parameter(0) ' Quality
                ' Setzen der Quality GUID
                CLSIDFromString StrPtr(EncoderQuality), .GUID
                .NumberOfValues = 1
                .type = EncoderParameterValueTypeLong
                .value = VarPtr(Quality)

            End With
              
            ' Speichert lBitmap als JPG
            retStatus = Execute(GdipSaveImageToFile(lBitmap, StrPtr(FileName), PicEncoder, tParams))
              
            If retStatus = Ok Then
                Convertir = True
            Else
                Convertir = False

            End If

        Else
            Convertir = False
            MsgBox "Konnte keinen passenden Encoder ermitteln.", vbOKOnly, "Encoder Error"

        End If
        
        ' Lösche lBitmap
        Call Execute(GdipDisposeImage(lBitmap))
        
        Dim ret As Long

        If GdipInitialized = True Then
            ret = Execute(ShutdownGDIPlus)

        End If

    End If

End Function

Private Function GetEncoderClsid(MimeType As String, pClsid As GUID) As Boolean
    
    Dim num               As Long
    Dim Size              As Long
    Dim pImageCodecInfo() As ImageCodecInfo
    Dim j                 As Long
    Dim buffer            As String
    
    Call GdipGetImageEncodersSize(num, Size)

    If (Size = 0) Then
        GetEncoderClsid = False
        Exit Function

    End If
    
    ReDim pImageCodecInfo(0 To Size \ Len(pImageCodecInfo(0)) - 1)
    Call GdipGetImageEncoders(num, Size, pImageCodecInfo(0))
    
    For j = 0 To num - 1
        buffer = Space$(lstrlenW(ByVal pImageCodecInfo(j).MimeTypePtr))
        
        Call lstrcpyW(ByVal StrPtr(buffer), ByVal pImageCodecInfo(j).MimeTypePtr)
              
        If (StrComp(buffer, MimeType, vbTextCompare) = 0) Then
            pClsid = pImageCodecInfo(j).Clsid
            Erase pImageCodecInfo
            GetEncoderClsid = True
            Exit Function

        End If

    Next j
    
    Erase pImageCodecInfo
    GetEncoderClsid = False

End Function

Private Sub Iniciar()
    Dim ret As Long
    ret = Execute(StartUpGDIPlus(1))

    If ret = 0 Then
        GdipInitialized = True
    Else
        MsgBox "El GDI no está inicializado", vbOKOnly, "GDI Error"

    End If

End Sub

