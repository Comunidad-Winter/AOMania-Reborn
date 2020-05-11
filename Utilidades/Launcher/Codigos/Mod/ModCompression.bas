Attribute VB_Name = "ModCompression"
' LauncherAoM 1.0.0
' By Bassinger [www.AoMania.Net]

Option Explicit

Private Const PASSWORD_RESOURCE_FILE As String = "123456"
Private Password()                   As Byte ' Contraseña
Private PasswordLen                  As Integer

Public Const GRH_SOURCE_FILE_EXT     As String = ".bmp"
Public Const GRH_RESOURCE_FILE       As String = "Graphics.AoM"
Public Const GRH_PATCH_FILE          As String = "Graphics.AoM"
Public Const INT_RESOURCE_FILE       As String = "Interface.AoM"
Public Const INIT_RESOURCE_FILE      As String = "Init.AoM"
Public Const WAV_RESOURCE_FILE       As String = "Wav.AoM"
Public Const MAPAS_RESOURCE_FILE     As String = "Mapas.AoM"
Public Const MIDI_RESOURCE_FILE      As String = "Midi.AoM"
Public Const MINIMAPA_FILE As String = "MiniMapa.AoM"
Public Const ICONOS_FILE As String = "Iconos.AoM"


Public Type FILEHEADER

    lngNumFiles As Long
    lngFileSize As Long
    lngFileVersion As Long

End Type
 
'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER

    lngFileSize As Long
    lngFileStart As Long
    strFileName As String * 16
    lngFileSizeUncompressed As Long
    
End Type
 
Private Enum PatchInstruction

    Delete_File
    Create_File
    Modify_File

End Enum
 
Private Declare Function compress Lib "zlib.dll" (dest As Any, destlen As Any, Src As Any, ByVal srclen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destlen As Any, Src As Any, ByVal srclen As Long) As Long
 
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef source As Any, ByVal byteCount As Long)
 
'BitMaps Strucures
Public Type BITMAPFILEHEADER

    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long

End Type

Public Type BITMAPINFOHEADER

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

Public Type RGBQUAD

    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte

End Type

Public Type BITMAPINFO

    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD

End Type
 
Private Const BI_RGB       As Long = 0
Private Const BI_RLE8      As Long = 1
Private Const BI_RLE4      As Long = 2
Private Const BI_BITFIELDS As Long = 3
Private Const BI_JPG       As Long = 4
Private Const BI_PNG       As Long = 5
 
'To get free bytes in drive
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, _
    FreeBytesToCaller As Currency, _
    bytesTotal As Currency, _
    FreeBytesTotal As Currency) As Long
 
Public Sub InitializeCompression()
    Dim LooPC As Long
    
    PasswordLen = Len(PASSWORD_RESOURCE_FILE) - 1
     
    ReDim Password(0 To PasswordLen) As Byte

    For LooPC = 0 To PasswordLen
        Password(LooPC) = Asc(Mid$(PASSWORD_RESOURCE_FILE, LooPC + 1, 1))
    Next LooPC
            
End Sub
 
Private Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency

    Dim retval As Long
    Dim FB     As Currency
    Dim BT     As Currency
    Dim FBT    As Currency
    
    retval = GetDiskFreeSpace(Left$(DriveName, 2), FB, BT, FBT)
    
    General_Drive_Get_Free_Bytes = FB * 10000

End Function
 
Private Sub Sort_Info_Headers(ByRef InfoHead() As INFOHEADER, ByVal first As Long, ByVal last As Long)

    Dim aux  As INFOHEADER
    Dim min  As Long
    Dim max  As Long
    Dim comp As String
    
    min = first
    max = last
    
    comp = InfoHead((min + max) \ 2).strFileName
    
    Do While min <= max
        Do While InfoHead(min).strFileName < comp And min < last
            min = min + 1
        Loop

        Do While InfoHead(max).strFileName > comp And max > first
            max = max - 1
        Loop

        If min <= max Then
            aux = InfoHead(min)
            InfoHead(min) = InfoHead(max)
            InfoHead(max) = aux
            min = min + 1
            max = max - 1

        End If

    Loop
    
    If first < max Then Call Sort_Info_Headers(InfoHead, first, max)
    If min < last Then Call Sort_Info_Headers(InfoHead, min, last)

End Sub
  
Private Function BinarySearch(ByRef ResourceFile As Integer, _
    ByRef InfoHead As INFOHEADER, _
    ByVal FirstHead As Long, _
    ByVal LastHead As Long, _
    ByVal FileHeaderSize As Long, _
    ByVal InfoHeaderSize As Long) As Boolean

    Dim ReadingHead  As Long
    Dim ReadInfoHead As INFOHEADER
    
    Do Until FirstHead > LastHead
        ReadingHead = (FirstHead + LastHead) \ 2
 
        Get ResourceFile, FileHeaderSize + InfoHeaderSize * (ReadingHead - 1) + 1, ReadInfoHead
 
        If InfoHead.strFileName = ReadInfoHead.strFileName Then
            InfoHead = ReadInfoHead
            BinarySearch = True
            Exit Function
        Else

            If InfoHead.strFileName < ReadInfoHead.strFileName Then
                LastHead = ReadingHead - 1
            Else
                FirstHead = ReadingHead + 1

            End If

        End If

    Loop

End Function
 

 
Private Function Get_InfoHeader(ByRef ResourcePath As String, _
    ByRef FileName As String, _
    ByRef InfoHead As INFOHEADER, _
    ByRef NameFile As String) As Boolean

    Dim ResourceFile     As Integer
    Dim ResourceFilePath As String
    Dim FileHead         As FILEHEADER
    
    On Local Error GoTo ErrHandler
 
    ResourceFilePath = ResourcePath & NameFile
    

    InfoHead.strFileName = UCase$(FileName)
    

    ResourceFile = FreeFile()
 
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
 
    Get ResourceFile, 1, FileHead
        

    If LOF(ResourceFile) <> FileHead.lngFileSize Then
        MsgBox "Archivo de recursos dañado. " & ResourceFilePath, , "Error"
        Close ResourceFile
        Exit Function

    End If
        
   
    If BinarySearch(ResourceFile, InfoHead, 1, FileHead.lngNumFiles, Len(FileHead), Len(InfoHead)) Then
            
        Get_InfoHeader = True

    End If
        
    Close ResourceFile
    Exit Function
 
ErrHandler:
    Close ResourceFile
    
    Call LogLauncher("Error al intentar leer el archivo " & ResourceFilePath & ". Razón: " & Err.Number & " : " & Err.Description)

End Function
 

Private Sub Compress_Data(ByRef Data() As Byte)
 
    Dim Dimensions As Long
    Dim DimBuffer  As Long
    Dim BufTemp()  As Byte
    Dim LooPC      As Long
    
    Dimensions = UBound(Data) + 1
    
    DimBuffer = Dimensions * 1.06
    
    ReDim BufTemp(DimBuffer)
    
    Call compress(BufTemp(0), DimBuffer, Data(0), Dimensions)
    
    Erase Data
    
    ReDim Data(DimBuffer - 1)
    ReDim Preserve BufTemp(DimBuffer - 1)
    
    Data = BufTemp
    
    Erase BufTemp

    If PasswordLen <= UBound(Data) And PasswordLen <> 0 Then

        For LooPC = 0 To PasswordLen
            Data(LooPC) = Data(LooPC) Xor Password(LooPC)
        Next LooPC

    End If

End Sub

 
Private Sub Decompress_Data(ByRef Data() As Byte, ByVal OrigSize As Long)
 
    Dim BufTemp() As Byte
    Dim LooPC     As Long
    
    ReDim BufTemp(OrigSize - 1)
    
    If PasswordLen <= UBound(Data) And PasswordLen <> 0 Then

        For LooPC = 0 To PasswordLen
            Data(LooPC) = Data(LooPC) Xor Password(LooPC)
        Next LooPC

    End If
    
    Call uncompress(BufTemp(0), OrigSize, Data(0), UBound(Data) + 1)
    
    ReDim Data(OrigSize - 1)
    
    Data = BufTemp
    
    Erase BufTemp

End Sub
  
Public Function Compress_Files(ByRef SourcePath As String, _
    ByRef OutputPath As String, _
    ByVal Version As Long, _
    ByRef prgBar As ProgressBar, _
    ByVal NameFile As String, _
    Optional ByRef Extension As String = ".bmp") As Boolean

    Dim SourceFileName As String
    Dim OutputFilePath As String
    Dim SourceFile     As Long
    Dim OutputFile     As Long
    Dim SourceData()   As Byte
    Dim FileHead       As FILEHEADER
    Dim InfoHead()     As INFOHEADER
    Dim LooPC          As Long
 
    'On Local Error GoTo ErrHandler
    OutputFilePath = OutputPath & NameFile
    SourceFileName = Dir(SourcePath & "*" & Extension, vbNormal)
    
  
    While SourceFileName <> ""

        FileHead.lngNumFiles = FileHead.lngNumFiles + 1
        
        ReDim Preserve InfoHead(FileHead.lngNumFiles - 1)
        InfoHead(FileHead.lngNumFiles - 1).strFileName = UCase$(SourceFileName)
        
        'Search new file
        SourceFileName = Dir()
    Wend
    
    If FileHead.lngNumFiles = 0 Then
        MsgBox "No se encontraron archivos de extención " & GRH_SOURCE_FILE_EXT & " en " & SourcePath & ".", , "Error"
        Exit Function

    End If
    
    If Not prgBar Is Nothing Then
        prgBar.max = FileHead.lngNumFiles
        prgBar.Value = 0

    End If
    
    
    If Dir(OutputFilePath, vbNormal) <> "" Then
        Kill OutputFilePath

    End If
    
  
    FileHead.lngFileVersion = Version
    FileHead.lngFileSize = Len(FileHead) + FileHead.lngNumFiles * Len(InfoHead(0))
    
  
    Call Sort_Info_Headers(InfoHead(), 0, FileHead.lngNumFiles - 1)
    
   
    OutputFile = FreeFile()
    Open OutputFilePath For Binary Access Read Write As OutputFile
     
    Seek OutputFile, FileHead.lngFileSize + 1
        
   
    For LooPC = 0 To FileHead.lngNumFiles - 1
            
        SourceFile = FreeFile()
        Open SourcePath & InfoHead(LooPC).strFileName For Binary Access Read Lock Write As SourceFile
                
      
        InfoHead(LooPC).lngFileSizeUncompressed = LOF(SourceFile)
        ReDim SourceData(LOF(SourceFile) - 1)
                
      
        Get SourceFile, , SourceData
                

        Call Compress_Data(SourceData)
                
     
        Put OutputFile, , SourceData
                
        With InfoHead(LooPC)
  
            .lngFileSize = UBound(SourceData) + 1
            .lngFileStart = FileHead.lngFileSize + 1
                    
           
            FileHead.lngFileSize = FileHead.lngFileSize + .lngFileSize

        End With
                
        Erase SourceData
            
        Close SourceFile
        
     
        If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
        DoEvents
    Next LooPC
        
   
    Seek OutputFile, 1
    Put OutputFile, , FileHead
    Put OutputFile, , InfoHead

    Close OutputFile
    
    Erase InfoHead
    Erase SourceData
    
    Compress_Files = True
    Exit Function
 
ErrHandler:
    Erase SourceData
    Erase InfoHead
    Close OutputFile
    
    Call MsgBox("No se pudo crear el archivo binario. Razón: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error")

End Function
  
Public Function Get_File_RawData(ByRef ResourcePath As String, _
    ByRef InfoHead As INFOHEADER, _
    ByRef Data() As Byte, _
    ByRef NameFile As String) As Boolean

    Dim ResourceFilePath As String
    Dim ResourceFile     As Integer
    
    On Local Error GoTo ErrHandler
    ResourceFilePath = ResourcePath & NameFile
    

    ReDim Data(InfoHead.lngFileSize - 1)
    
   
    ResourceFile = FreeFile
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
   
    Get ResourceFile, InfoHead.lngFileStart, Data
   
    Close ResourceFile
    
    Get_File_RawData = True
    Exit Function
 
ErrHandler:
    Close ResourceFile

End Function
 

Public Function Extract_File(ByRef ResourcePath As String, ByRef InfoHead As INFOHEADER, ByRef Data() As Byte, ByRef NameFile As String) As Boolean
 
    On Local Error GoTo ErrHandler
    
    If Get_File_RawData(ResourcePath, InfoHead, Data, NameFile) Then

      
        If InfoHead.lngFileSize < InfoHead.lngFileSizeUncompressed Then
            Call Decompress_Data(Data, InfoHead.lngFileSizeUncompressed)

        End If
        
        Extract_File = True

    End If

    Exit Function
 
ErrHandler:
    Call LogLauncher("Error al intentar decodificar recursos. Razon: " & Err.Number & " : " & Err.Description)

End Function
 

 
Public Function Extract_Files(ByRef ResourcePath As String, _
    ByRef OutputPath As String, _
    ByRef prgBar As ProgressBar, _
    ByRef NameFile As String) As Boolean
  
    Dim LooPC            As Long
    Dim ResourceFile     As Integer
    Dim ResourceFilePath As String
    Dim OutputFile       As Integer
    Dim SourceData()     As Byte
    Dim FileHead         As FILEHEADER
    Dim InfoHead()       As INFOHEADER
    Dim RequiredSpace    As Currency
    
    On Local Error GoTo ErrHandler
    ResourceFilePath = ResourcePath & NameFile
    
   
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
    
    Get ResourceFile, 1, FileHead
        
    
    If LOF(ResourceFile) <> FileHead.lngFileSize Then
        Call MsgBox("Archivo de recursos dañado. " & ResourceFilePath, , "Error")
        Close ResourceFile
        Exit Function

    End If
        
    ReDim InfoHead(FileHead.lngNumFiles - 1)
        
    Get ResourceFile, , InfoHead
        
    For LooPC = 0 To UBound(InfoHead)
            
        RequiredSpace = RequiredSpace + InfoHead(LooPC).lngFileSizeUncompressed
    Next LooPC
        
    If RequiredSpace >= General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
        Erase InfoHead
        Close ResourceFile
        Call MsgBox("No hay suficiente espacio en el disco para extraer los archivos.", , "Error")
        Exit Function

    End If

    Close ResourceFile
    
    If Not prgBar Is Nothing Then
        prgBar.max = FileHead.lngNumFiles
        prgBar.Value = 0

    End If
    
    For LooPC = 0 To UBound(InfoHead)

        If Extract_File(ResourcePath, InfoHead(LooPC), SourceData, NameFile) Then

            If FileExist(OutputPath & InfoHead(LooPC).strFileName, vbNormal) Then
                Call Kill(OutputPath & InfoHead(LooPC).strFileName)

            End If
            
            OutputFile = FreeFile()
            Open OutputPath & InfoHead(LooPC).strFileName For Binary As OutputFile
            Put OutputFile, , SourceData
            Close OutputFile
            
            Erase SourceData
        Else
            Erase SourceData
            Erase InfoHead
            
            Call MsgBox("No se pudo extraer el archivo " & InfoHead(LooPC).strFileName, vbOKOnly, "Error")
            Exit Function

        End If
            
        If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
        DoEvents
    Next LooPC
    
    Erase InfoHead
    Extract_Files = True
    Exit Function
 
ErrHandler:
    Close ResourceFile
    Erase SourceData
    Erase InfoHead
    
    Call MsgBox("No se pudo extraer el archivo binario correctamente. Razon: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error")

End Function
  
Public Function Get_File_Data(ByRef ResourcePath As String, ByRef FileName As String, ByRef Data() As Byte, ByRef NameFile As String) As Boolean

    Dim InfoHead As INFOHEADER
    
    If Get_InfoHeader(ResourcePath, FileName, InfoHead, NameFile) Then
   
        Get_File_Data = Extract_File(ResourcePath, InfoHead, Data, NameFile)
    Else
        Call LogLauncher("No se encontro el recurso " & FileName)

    End If

End Function
  
Public Function Get_Bitmap(ByRef ResourcePath As String, ByRef FileName As String, ByRef bmpInfo As BITMAPINFO, ByRef Data() As Byte) As Boolean

    Dim InfoHead   As INFOHEADER
    Dim rawData()  As Byte
    Dim offBits    As Long
    Dim bitmapSize As Long
    Dim colorCount As Long
    
    If Get_InfoHeader(ResourcePath, FileName, InfoHead, GRH_RESOURCE_FILE) Then

        If Extract_File(ResourcePath, InfoHead, rawData, GRH_RESOURCE_FILE) Then
            Call CopyMemory(offBits, rawData(10), 4)
            Call CopyMemory(bmpInfo.bmiHeader, rawData(14), 40)
            
            With bmpInfo.bmiHeader
                bitmapSize = AlignScan(.biWidth, .biBitCount) * Abs(.biHeight)
                
                If .biBitCount < 24 Or .biCompression = BI_BITFIELDS Or (.biCompression <> BI_RGB And .biBitCount = 32) Then
                    If .biClrUsed < 1 Then
                        colorCount = 2 ^ .biBitCount
                    Else
                        colorCount = .biClrUsed

                    End If
                    
                    If .biBitCount >= 16 Then colorCount = 3
                    
                    Call CopyMemory(bmpInfo.bmiColors(0), rawData(54), colorCount * 4)

                End If

            End With
            
            ReDim Data(bitmapSize - 1) As Byte
            Call CopyMemory(Data(0), rawData(offBits), bitmapSize)
            
            Get_Bitmap = True

        End If

    Else
        Call MsgBox("-No se encontro el recurso " & FileName)

    End If

End Function
 
Private Function Compare_Datas(ByRef Data1() As Byte, ByRef Data2() As Byte) As Boolean

    Dim Length As Long
    Dim act    As Long
    
    Length = UBound(Data1) + 1
    
    If (UBound(Data2) + 1) = Length Then

        While act < Length

            If Data1(act) Xor Data2(act) Then Exit Function
            
            act = act + 1
        Wend
        
        Compare_Datas = True

    End If

End Function
  
Private Function ReadNext_InfoHead(ByRef ResourceFile As Integer, _
    ByRef FileHead As FILEHEADER, _
    ByRef InfoHead As INFOHEADER, _
    ByRef ReadFiles As Long) As Boolean
 
    If ReadFiles < FileHead.lngNumFiles Then
        'Read header
        Get ResourceFile, Len(FileHead) + Len(InfoHead) * ReadFiles + 1, InfoHead
    
        'Update
        ReadNext_InfoHead = True

    End If
    
    ReadFiles = ReadFiles + 1

End Function
 
 
Public Function GetNext_Bitmap(ByRef ResourcePath As String, _
    ByRef ReadFiles As Long, _
    ByRef bmpInfo As BITMAPINFO, _
    ByRef Data() As Byte, _
    ByRef fileIndex As Long) As Boolean

    On Error Resume Next
 
    Dim ResourceFile As Integer
    Dim FileHead     As FILEHEADER
    Dim InfoHead     As INFOHEADER
    Dim FileName     As String
    
    ResourceFile = FreeFile
    Open ResourcePath & GRH_RESOURCE_FILE For Binary Access Read Lock Write As ResourceFile
    Get ResourceFile, 1, FileHead
    
    If ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ReadFiles) Then
        
        Call Get_Bitmap(ResourcePath, InfoHead.strFileName, bmpInfo, Data())
        FileName = Trim$(InfoHead.strFileName)
        fileIndex = CLng(Left$(FileName, Len(FileName) - 4))
        
        GetNext_Bitmap = True

    End If
    
    Close ResourceFile

End Function
  
Public Function Make_Patch(ByRef NewResourcePath As String, _
    ByRef OldResourcePath As String, _
    ByRef OutputPath As String, _
    ByRef prgBar As ProgressBar) As Boolean

    Dim NewResourceFile     As Integer
    Dim NewResourceFilePath As String
    Dim NewFileHead         As FILEHEADER
    Dim NewInfoHead         As INFOHEADER
    Dim NewReadFiles        As Long
    Dim NewReadNext         As Boolean
    
    Dim OldResourceFile     As Integer
    Dim OldResourceFilePath As String
    Dim OldFileHead         As FILEHEADER
    Dim OldInfoHead         As INFOHEADER
    Dim OldReadFiles        As Long
    Dim OldReadNext         As Boolean
    
    Dim OutputFile          As Integer
    Dim OutputFilePath      As String
    Dim Data()              As Byte
    Dim auxData()           As Byte
    Dim Instruction         As Byte
    

    'On Local Error GoTo ErrHandler
 
    NewResourceFilePath = NewResourcePath & GRH_RESOURCE_FILE
    OldResourceFilePath = OldResourcePath & GRH_RESOURCE_FILE
    OutputFilePath = OutputPath & GRH_PATCH_FILE
    

    OldResourceFile = FreeFile
    Open OldResourceFilePath For Binary Access Read Lock Write As OldResourceFile

    Get OldResourceFile, 1, OldFileHead
        
    If LOF(OldResourceFile) <> OldFileHead.lngFileSize Then
        Call MsgBox("Archivo de recursos anterior dañado. " & OldResourceFilePath, , "Error")
        Close OldResourceFile
        Exit Function

    End If
        
    NewResourceFile = FreeFile()
    Open NewResourceFilePath For Binary Access Read Lock Write As NewResourceFile
            
    Get NewResourceFile, 1, NewFileHead
            
    If LOF(NewResourceFile) <> NewFileHead.lngFileSize Then
        Call MsgBox("Archivo de recursos anterior dañado. " & NewResourceFilePath, , "Error")
        Close NewResourceFile
        Close OldResourceFile
        Exit Function

    End If
            
    If Dir(OutputFilePath, vbNormal) <> "" Then Kill OutputFilePath
            
    OutputFile = FreeFile()
    Open OutputFilePath For Binary Access Read Write As OutputFile
                
    If Not prgBar Is Nothing Then
        prgBar.max = OldFileHead.lngNumFiles + NewFileHead.lngNumFiles
        prgBar.Value = 0

    End If
                
    Put OutputFile, , OldFileHead.lngFileVersion
             
    Put OutputFile, , NewFileHead
                
    If ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) And ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, _
        NewReadFiles) Then
                    
        prgBar.Value = prgBar.Value + 2
                    
        Do 'Main loop

            If OldInfoHead.strFileName = NewInfoHead.strFileName Then
 
                Call Get_File_RawData(OldResourcePath, OldInfoHead, auxData, GRH_RESOURCE_FILE)
                            
                Call Get_File_RawData(NewResourcePath, NewInfoHead, Data, GRH_RESOURCE_FILE)
                            
                If Not Compare_Datas(Data, auxData) Then
                
                    Instruction = PatchInstruction.Modify_File
                    Put OutputFile, , Instruction
    
                    Put OutputFile, , NewInfoHead
           
                    Put OutputFile, , Data

                End If
                            
              If Not ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) Then
                    Exit Do

                End If
                            
               
                If Not ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
                   
                    OldReadFiles = OldReadFiles - 1
                    Exit Do

                End If
                            
          
                If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 2
                        
            ElseIf OldInfoHead.strFileName < NewInfoHead.strFileName Then
                            
         
                Instruction = PatchInstruction.Delete_File
                Put OutputFile, , Instruction
                Put OutputFile, , OldInfoHead
                            
       
                If Not ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) Then
              
                    NewReadFiles = NewReadFiles - 1
                    Exit Do

                End If
                            
                
                If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
                        
            Else
                            
         
                Instruction = PatchInstruction.Create_File
                Put OutputFile, , Instruction
                Put OutputFile, , NewInfoHead
        
                Call Get_File_RawData(NewResourcePath, NewInfoHead, Data, GRH_RESOURCE_FILE)
    
                Put OutputFile, , Data
      
                If Not ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
           
                    OldReadFiles = OldReadFiles - 1
                    Exit Do

                End If
                            

                If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1

            End If
                        
            DoEvents
        Loop
                
    Else

        OldReadFiles = 0
        NewReadFiles = 0

    End If
                

    While ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles)

        Instruction = PatchInstruction.Delete_File
        Put OutputFile, , Instruction
        Put OutputFile, , OldInfoHead
                    
        If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
        DoEvents
    Wend
                
    While ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles)

        Instruction = PatchInstruction.Create_File
        Put OutputFile, , Instruction
        Put OutputFile, , NewInfoHead
                    
        Call Get_File_RawData(NewResourcePath, NewInfoHead, Data, GRH_RESOURCE_FILE)
                    
        Put OutputFile, , Data
                    
        If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
        DoEvents
    Wend
            
    Close OutputFile
        
    Close NewResourceFile
    
    Close OldResourceFile
    
    Make_Patch = True
    Exit Function
 
ErrHandler:
    Close OutputFile
    Close NewResourceFile
    Close OldResourceFile
    
    Call MsgBox("No se pudo terminar de crear el parche. Razon: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error")

End Function
 
Public Function Apply_Patch(ByRef ResourcePath As String, ByRef PatchPath As String, ByRef prgBar As ProgressBar) As Boolean

    Dim ResourceFile       As Integer
    Dim ResourceFilePath   As String
    Dim FileHead           As FILEHEADER
    Dim InfoHead           As INFOHEADER
    Dim ResourceReadFiles  As Long
    Dim EOResource         As Boolean
 
    Dim PatchFile          As Integer
    Dim PatchFilePath      As String
    Dim PatchFileHead      As FILEHEADER
    Dim PatchInfoHead      As INFOHEADER
    Dim Instruction        As Byte
    Dim OldResourceVersion As Long
 
    Dim OutputFile         As Integer
    Dim OutputFilePath     As String
    Dim Data()             As Byte
    Dim WrittenFiles       As Long
    Dim DataOutputPos      As Long
 
    On Local Error GoTo ErrHandler
 
    ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
    PatchFilePath = PatchPath & GRH_PATCH_FILE
    OutputFilePath = ResourcePath & GRH_RESOURCE_FILE & "tmp"
    
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        
    Get ResourceFile, , FileHead
        
    If LOF(ResourceFile) <> FileHead.lngFileSize Then
        Call MsgBox("Archivo de recursos anterior dañado. " & ResourceFilePath, , "Error")
        Close ResourceFile
        Exit Function

    End If
        
    PatchFile = FreeFile()
    Open PatchFilePath For Binary Access Read Lock Write As PatchFile
            
    Get PatchFile, , OldResourceVersion
            
    If OldResourceVersion <> FileHead.lngFileVersion Then
        Call MsgBox("Incongruencia en versiones.", , "Error")
        Close ResourceFile
        Close PatchFile
        Exit Function

    End If
            
    Get PatchFile, , PatchFileHead
            
    If FileExist(OutputFilePath, vbNormal) Then Call Kill(OutputFilePath)
            
    OutputFile = FreeFile()
    Open OutputFilePath For Binary Access Read Write As OutputFile
                
    Put OutputFile, , PatchFileHead
                
    If Not prgBar Is Nothing Then
        prgBar.max = PatchFileHead.lngNumFiles
        prgBar.Value = 0

    End If
                
    DataOutputPos = Len(FileHead) + Len(InfoHead) * PatchFileHead.lngNumFiles + 1
                
    While Loc(PatchFile) < LOF(PatchFile)
                    
        Get PatchFile, , Instruction
        
        Get PatchFile, , PatchInfoHead
                    
        Do
            EOResource = Not ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ResourceReadFiles)
                        
            If Not EOResource And InfoHead.strFileName < PatchInfoHead.strFileName Then
 
                Call Get_File_RawData(ResourcePath, InfoHead, Data, GRH_RESOURCE_FILE)
                InfoHead.lngFileStart = DataOutputPos
                            
                Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, InfoHead
                Put OutputFile, DataOutputPos, Data
                            
                DataOutputPos = DataOutputPos + UBound(Data) + 1
                WrittenFiles = WrittenFiles + 1

                If Not prgBar Is Nothing Then prgBar.Value = WrittenFiles
            Else
                Exit Do

            End If

        Loop
                    
        Select Case Instruction

            Case PatchInstruction.Delete_File

                If InfoHead.strFileName <> PatchInfoHead.strFileName Then
                    Err.Description = "Incongruencia en archivos de recurso"
                    GoTo ErrHandler

                End If
                        
                'Create
            Case PatchInstruction.Create_File

                If (InfoHead.strFileName > PatchInfoHead.strFileName) Or EOResource Then
 
                    'Get file data
                    ReDim Data(PatchInfoHead.lngFileSize - 1)
                    Get PatchFile, , Data
                                
                    'Save it
                    Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, PatchInfoHead
                    Put OutputFile, DataOutputPos, Data
                                
                    'Reanalize last Resource InfoHead
                    EOResource = False
                    ResourceReadFiles = ResourceReadFiles - 1
                                
                    'Update
                    DataOutputPos = DataOutputPos + UBound(Data) + 1
                    WrittenFiles = WrittenFiles + 1

                    If Not prgBar Is Nothing Then prgBar.Value = WrittenFiles
                Else
                    Err.Description = "Incongruencia en archivos de recurso"
                    GoTo ErrHandler

                End If
                        
                'Modify
            Case PatchInstruction.Modify_File

                If InfoHead.strFileName = PatchInfoHead.strFileName Then
 
                    'Get file data
                    ReDim Data(PatchInfoHead.lngFileSize - 1)
                    Get PatchFile, , Data
                                
                    'Save it
                    Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, PatchInfoHead
                    Put OutputFile, DataOutputPos, Data
                                
                    'Update
                    DataOutputPos = DataOutputPos + UBound(Data) + 1
                    WrittenFiles = WrittenFiles + 1

                    If Not prgBar Is Nothing Then prgBar.Value = WrittenFiles
                Else
                    Err.Description = "Incongruencia en archivos de recurso"
                    GoTo ErrHandler

                End If

        End Select
                    
        DoEvents
    Wend
                
    While ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ResourceReadFiles)
 
        Call Get_File_RawData(ResourcePath, InfoHead, Data, GRH_RESOURCE_FILE)
        InfoHead.lngFileStart = DataOutputPos
                    
        Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, InfoHead
        Put OutputFile, DataOutputPos, Data
                    
        DataOutputPos = DataOutputPos + UBound(Data) + 1
        WrittenFiles = WrittenFiles + 1

        If Not prgBar Is Nothing Then prgBar.Value = WrittenFiles
        DoEvents
    Wend
            
    Close OutputFile
        
    Close PatchFile
    
    Close ResourceFile
    
    If (PatchFileHead.lngNumFiles = WrittenFiles) Then
    
        Call Kill(ResourceFilePath)
        Name OutputFilePath As ResourceFilePath
        
    Else
        Err.Description = "Falla al procesar parche"
        GoTo ErrHandler

    End If
    
    Apply_Patch = True
    Exit Function
 
ErrHandler:
    Close OutputFile
    Close PatchFile
    Close ResourceFile

    If FileExist(OutputFilePath, vbNormal) Then Call Kill(OutputFilePath)
    
    Call MsgBox("No se pudo parchear. Razon: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error")

End Function
 
Private Function AlignScan(ByVal inWidth As Long, ByVal inDepth As Integer) As Long

    AlignScan = (((inWidth * inDepth) + &H1F) And Not &H1F&) \ &H8

End Function
  
Public Function GetVersion(ByVal ResourceFilePath As String) As Long

    Dim ResourceFile As Integer
    Dim FileHead     As FILEHEADER
    
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
    'Extract the FILEHEADER
    Get ResourceFile, 1, FileHead
        
    Close ResourceFile
    
    GetVersion = FileHead.lngFileVersion

End Function

