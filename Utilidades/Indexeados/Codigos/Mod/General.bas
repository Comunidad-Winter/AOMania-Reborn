Attribute VB_Name = "General"
Option Explicit

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

Sub Main()
     Call InitializeCompression
     Call CargarGRH
     Call LoadObjData
     frmMain.Show
   
End Sub

Public Function CargarGRH() As Boolean

    Dim Grh         As Long
    Dim Frame       As Long
    Dim handle      As Integer
    Dim fileVersion As Long
    Dim i As Integer

    
    Dim Data()       As Byte
    Dim TemporalFile As String

   Call Get_File_Data(DirLibs, "GRAFICOS.IND", Data, INIT_RESOURCE_FILE)
    
    TemporalFile = DirLibs & "GRAFICOS.IND"
handle = FreeFile
    
    Open TemporalFile For Binary Access Write As handle
    Put handle, , Data
    Close handle
     
    Open TemporalFile For Binary Access Read As handle
    
    Seek handle, 1
    
    Get handle, , fileVersion
    Get handle, , grhCount

    'Resize arrays
    ReDim GrhData(0 To 50000) As GrhData
   
    While Not EOF(handle)

        Get handle, , Grh
        
        If Grh <> 0 Then

            With GrhData(Grh)
                'Get number of frames
                Get handle, , .NumFrames
           
                ReDim .Frames(1 To .NumFrames)
            
                If .NumFrames > 1 Then
                   

                    'Read a animation GRH set
                    For Frame = 1 To .NumFrames
                    
                        Get handle, , .Frames(Frame)
            
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                           

                        End If
        
                    Next Frame
    
                    Get handle, , .Speed

                    If .Speed < 0 Then MsgBox Grh & " velocidad <= 0 ", , "advertencia"
        
                    'Compute width and height
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight

                    If .pixelHeight < 0 Then MsgBox Grh & " ancho <= 0 ", , "advertencia"
        
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth

                    If .pixelWidth < 0 Then MsgBox Grh & " ancho <= 0 ", , "advertencia"
        
                    .TileWidth = GrhData(.Frames(1)).TileWidth

                    If .TileWidth < 0 Then MsgBox Grh & " anchoT <= 0 ", , "advertencia"
        
                    .TileHeight = GrhData(.Frames(1)).TileHeight

                    If .TileHeight < 0 Then MsgBox Grh & " altoT <= 0 ", , "advertencia"
    
                Else
                
                    
                    'Read in normal GRH data
                    Get handle, , .FileNum

                    If .FileNum <= 0 Then MsgBox Grh & " tiene bmp = 0 ", , "advertencia"
           
                    Get handle, , .sX

                    If .sX < 0 Then MsgBox Grh & " tiene Sx <= 0 ", , "advertencia"
        
                    Get handle, , .sY

                    If .sY < 0 Then MsgBox Grh & " tiene Sy <= 0 ", , "advertencia"
            
                    Get handle, , .pixelWidth

                    If .pixelWidth <= 0 Then MsgBox Grh & " ancho <= 0 ", , "advertencia"
        
                    Get handle, , .pixelHeight

                    If .pixelHeight <= 0 Then MsgBox Grh & " alto <= 0 ", , "advertencia"
        
                    'Compute width and height
                    .TileWidth = .pixelWidth / 32
                    .TileHeight = .pixelHeight / 32
        
                    .Frames(1) = Grh

                End If

            End With

        End If

    Wend
    '************************************************

    Close handle
    
    If FileExist(TemporalFile, vbArchive) Then Call Kill(TemporalFile)

End Function

Sub WriteVar(ByVal File As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal value As String)
    '*****************************************************************
    'Writes a var to a text file
    '*****************************************************************
    writeprivateprofilestring Main, Var, value, File

End Sub

Function GetVar(File As String, Main As String, Var As String) As String
    '*****************************************************************
    'Gets a Var from a text file
    '*****************************************************************

    Dim l        As Integer
    Dim Char     As String
    Dim sSpaces  As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found

    szReturn = ""

    sSpaces = Space$(5000)  ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish

    getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File

    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function

Function DirLibs()
     DirLibs = App.Path & "\libs\"
End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (LenB(Dir$(File, FileType)) <> 0)

End Function


Public Sub Grafico(ByVal Numero As Integer)

     Dim Data() As Byte
     Dim PicDIB As Picture
     Dim BmpData As StdPicture

    If Get_File_Data(DirLibs, CStr(GrhData(Numero).FileNum) & ".BMP", Data, GRH_RESOURCE_FILE) Then
                Set BmpData = ArrayToPicture(Data(), 0, UBound(Data) + 1)
      Erase Data

      
          With frmMain.Picture1
              
              .Picture = BmpData
              .PaintPicture frmMain.Picture1, 0, 0, , , GrhData(Numero).sX, _
                                    GrhData(Numero).sY, .ScaleWidth, .ScaleHeight
              .Width = GrhData(Numero).pixelWidth
              .Height = GrhData(Numero).pixelHeight
              
         End With
      
      End If
     
End Sub

Function DatPath()
    DatPath = App.Path & "\Dat\"
End Function

