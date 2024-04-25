Attribute VB_Name = "GameIni"

Option Explicit

Public Type tCabecera 'Cabecera de los con

    Desc As String * 255
    CRC As Long
    MagicWord As Long

End Type

Public Type tGameIni

    Puerto As Long
    Musica As Byte
    Fx As Byte
    tip As Byte
    Password As String
    name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer

End Type

Public MiCabecera    As tCabecera
Public Config_Inicio As tGameIni

Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
    Cabecera.Desc = _
            "Argentum Online by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10

End Sub

Sub WriteVar(ByVal file As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal value As String)
    '*****************************************************************
    'Writes a var to a text file
    '*****************************************************************
    writeprivateprofilestring Main, Var, value, file

End Sub

Public Function ExisteBMP(ByVal NumeroF As Long) As Byte

    'Funcion comprueba la existencia del bmp (tanto en archivo como en archivo de recursos)
    If NumeroF < 0 Or NumeroF > grhCount Then
        ExisteBMP = 0
        Exit Function

    End If

    If ResourceFile = 1 Or ResourceFile = 3 Then
        If FileExist(App.Path & "\" & CarpetaGraficos & "\" & Val(NumeroF) & ".bmp", vbNormal) Then
            ExisteBMP = 1 ' existe el bmp
            Exit Function

        End If

    End If

    If ResourceFile = 2 Or ResourceFile = 3 Then
        If Val(NumeroF) > ResourceF.UltimoGrafico Or NumeroF = 0 Then
            ExisteBMP = 0
            Exit Function

        End If

        If ResourceF.graficos(NumeroF).tamaño > 0 Then
            ExisteBMP = 2 ' existe en el archivo de recursos
            Exit Function

        End If

    End If

End Function

Public Sub GetTamañoBMP(ByVal fileIndex As Integer, ByRef Alto As Long, ByRef Ancho As Long, ByRef BitCount _
        As Integer)
    Dim datos    As ArchivoBMP
    Dim filePath As String

    If ExisteBMP(fileIndex) = ResourceFile And ResourceFile = 2 Then
        'Call Decryptdata(fileIndex, datos)
        Ancho = datos.bmpInfo.bmiHeader.biWidth
        Alto = datos.bmpInfo.bmiHeader.biHeight
        BitCount = datos.bmpInfo.bmiHeader.biBitCount
    ElseIf ExisteBMP(fileIndex) = ResourceFile And ResourceFile = 1 Then
        filePath = App.Path & "\" & CarpetaGraficos & "\" & CStr(fileIndex) & ".bmp"
        Call surfaceDimensions(filePath, Alto, Ancho, BitCount)
    ElseIf ResourceFile = 3 Then

        If ExisteBMP(fileIndex) = 1 Then
            filePath = App.Path & "\" & CarpetaGraficos & "\" & CStr(fileIndex) & ".bmp"
            Call surfaceDimensions(filePath, Alto, Ancho, BitCount)
        ElseIf ExisteBMP(fileIndex) = 2 Then
            'Call Decryptdata(fileIndex, datos)
            Ancho = datos.bmpInfo.bmiHeader.biWidth
            Alto = datos.bmpInfo.bmiHeader.biHeight
            BitCount = datos.bmpInfo.bmiHeader.biBitCount

        End If

    End If

End Sub

Public Sub surfaceDimensions(ByVal Archivo As String, _
                             ByRef Height As Long, _
                             ByRef Width As Long, _
                             ByRef BitCount As Integer)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    'Loads the headers of a bmp file to retrieve it's dimensions at rt
    '**************************************************************
    Dim Handle      As Integer
    Dim bmpFileHead As BITMAPFILEHEADER
    Dim bmpInfoHead As BITMAPINFOHEADER
    
    Handle = FreeFile()
    Open Archivo For Binary Access Read Lock Write As Handle
    Get Handle, , bmpFileHead
    Get Handle, , bmpInfoHead
    Close Handle
    
    Height = bmpInfoHead.biHeight
    Width = bmpInfoHead.biWidth
    BitCount = bmpInfoHead.biBitCount

End Sub

Private Function AlignScan(ByVal inWidth As Long, ByVal inDepth As Integer) As Long
    AlignScan = (((inWidth * inDepth) + &H1F) And Not &H1F&) \ &H8

End Function

Public Sub CalcularPosiciones(ByRef DataIndex As BodyData, ByRef Posiciones() As Position)
    Dim i                As Long
    Dim tGrhIndex        As Long
    Dim graficos(1 To 4) As Integer

    For i = 1 To 4

        If DataIndex.Walk(i).GrhIndex <= 0 Then DataIndex.Walk(i).GrhIndex = 1
        If GrhData(DataIndex.Walk(i).GrhIndex).NumFrames > 1 Then
            graficos(i) = GrhData(DataIndex.Walk(i).GrhIndex).Frames(1)
        Else
            graficos(i) = DataIndex.Walk(i).GrhIndex

        End If

    Next i

    For i = 1 To 4
        tGrhIndex = GrhData(DataIndex.Walk(i).GrhIndex).Frames(1)

        If tGrhIndex <= 0 Then Exit Sub
        If i = 1 Then
            Posiciones(i).X = ((GrhData(graficos(2)).pixelWidth + GrhData(graficos(4)).pixelWidth + 4) / 2) _
                    - (GrhData(graficos(1)).pixelWidth / 2)
            Posiciones(i).Y = 0
        ElseIf i = 2 Then
            Posiciones(i).X = GrhData(graficos(4)).pixelWidth + 2
            Posiciones(i).Y = GrhData(graficos(1)).pixelHeight + 2
        ElseIf i = 3 Then
            Posiciones(i).X = ((GrhData(graficos(2)).pixelWidth + GrhData(graficos(4)).pixelWidth + 4) / 2) _
                    - (GrhData(graficos(3)).pixelWidth / 2)
            Posiciones(i).Y = GrhData(graficos(1)).pixelHeight + GrhData(graficos(2)).pixelHeight + 4
        ElseIf i = 4 Then
            Posiciones(i).X = 0
            Posiciones(i).Y = GrhData(graficos(1)).pixelHeight + 2

        End If

    Next i

End Sub
 
Public Function StringRecurso(ByVal Recurso As Integer) As String

    Select Case Recurso

        Case 1
            StringRecurso = "BMP"

        Case 2
            StringRecurso = "ResF"

        Case 3
            StringRecurso "BMP o ResF"

    End Select

End Function

