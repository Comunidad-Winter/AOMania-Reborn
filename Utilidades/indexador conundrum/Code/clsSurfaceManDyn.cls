VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurfaceManDyn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsSurfaceManDyn.cls - Inherits from clsSurfaceManager. Is designed to load
'surfaces dynamically without using more than an arbitrary amount of Mb.
'For removale it uses LRU, attempting to just keep in memory those surfaces
'that are actually usefull.
'
' Developed by Maraxus (Juan Mart�n Sotuyo Dodero - juansotuyo@hotmail.com)
' Last Modify Date: 3/06/2006
'**************************************************************

'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

Option Explicit

'Inherit from the surface manager
Implements clsSurfaceManager

Private Const BYTES_PER_MB          As Long = 1048576                        '1Mb = 1024 Kb = 1024 * 1024 bytes = 1048576 bytes
Private Const MIN_MEMORY_TO_USE     As Long = 4 * BYTES_PER_MB          '4 Mb
Private Const DEFAULT_MEMORY_TO_USE As Long = 16 * BYTES_PER_MB     '16 Mb

Private Type SURFACE_ENTRY_DYN

    fileIndex As Long
    lastAccess As Long
    Surface As DirectDrawSurface7

End Type

Private surfaceList()     As SURFACE_ENTRY_DYN
Private surfaceCount      As Long

Private surfaceIndexes()  As Long
Private surfaceIndexCount As Long

Private DirectDraw        As DirectDraw7

Private maxBytesToUse     As Long
Private usedBytes         As Long

Private useVideoMemory    As Boolean

Private GrhPath           As String

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (ByRef dest As Any, _
                                       ByRef source As Any, _
                                       ByVal byteCount As Long)

Private Sub Class_Initialize()
    '**************************************************************
    'Author: Juan Mart�n Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    '
    '**************************************************************
    usedBytes = 0
    surfaceCount = 0
    surfaceIndexCount = 0
    ReDim surfaceList(0) As SURFACE_ENTRY_DYN
    ReDim surfaceIndexes(0) As Long
    maxBytesToUse = MIN_MEMORY_TO_USE

End Sub

Private Sub Class_Terminate()
    '**************************************************************
    'Author: Juan Mart�n Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    'Clean up
    '**************************************************************
    Dim i As Long
    
    'Destroy every surface in memory
    For i = 0 To surfaceCount - 1
        Set surfaceList(i).Surface = Nothing
    Next i
    
    'Destroy the arrays
    Erase surfaceList
    Erase surfaceIndexes

End Sub

Private Sub clsSurfaceManager_Initialize(ByRef DD As DirectDraw7, _
                                         ByVal videoMemory As Boolean, _
                                         ByVal graphicPath As String, _
                                         Optional ByVal maxMemoryUsageInMb As Long = -1)
    '**************************************************************
    'Author: Juan Mart�n Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    'Initializes the manager
    '**************************************************************
    Set DirectDraw = DD
    
    useVideoMemory = videoMemory
    
    GrhPath = graphicPath
    
    If maxMemoryUsageInMb = -1 Then
        maxBytesToUse = DEFAULT_MEMORY_TO_USE   ' 16 Mb by default
    ElseIf maxMemoryUsageInMb * BYTES_PER_MB < MIN_MEMORY_TO_USE Then
        maxBytesToUse = MIN_MEMORY_TO_USE       ' 4 Mb is the minimum allowed
    Else
        maxBytesToUse = maxMemoryUsageInMb * BYTES_PER_MB

    End If

End Sub

Private Property Get clsSurfaceManager_Surface(ByVal fileIndex As Long) As DirectDrawSurface7
    '**************************************************************
    'Author: Juan Mart�n Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    'Retrieves the requested texture
    '**************************************************************
    Dim Index As Long
    
    ' Search the index on the list
    Index = BinarySearch(fileIndex)
    
    If Index < 0 Then
        'Not found, we have to load the file and add it in the position given by the negation of the index
        '(it may be changed by the removal of indexes though, so we let the LoadSurface method notify us)
        Index = LoadSurface(fileIndex, Not Index)

    End If
    
    ' index = 287
    'Return it
    With surfaceList(surfaceIndexes(Index))
        .lastAccess = GetTickCount
        Set clsSurfaceManager_Surface = .Surface

    End With

End Property

Private Function BinarySearch(ByVal fileIndex As Long) As Long
    '**************************************************************
    'Author: Juan Mart�n Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    'Returns the index of the surface in the list, or the negation
    'of the position were it should be if not found (for binary insertion)
    '**************************************************************
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    
    min = 0
    max = surfaceIndexCount - 1
    
    Do While min <= max
        mid = (min + max) \ 2
        
        If surfaceList(surfaceIndexes(mid)).fileIndex < fileIndex Then
            min = mid + 1
        ElseIf surfaceList(surfaceIndexes(mid)).fileIndex > fileIndex Then
            max = mid - 1
        Else
            'We found it
            BinarySearch = mid
            Exit Function

        End If

    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    BinarySearch = Not mid

End Function

Private Function LoadSurface(ByVal fileIndex As Long, ByVal listIndex As Long) As Long

    '**************************************************************
    'Author: Juan Mart�n Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    'Loads the surface named fileIndex + ".bmp" and inserts it to the
    'surface list in the listIndex position
    '**************************************************************
    On Error GoTo ErrHandler

    Dim newSurface As SURFACE_ENTRY_DYN
    Dim ddsd       As DDSURFACEDESC2
    Dim ddck       As DDCOLORKEY
    Dim filePath   As String
    Dim datos      As ArchivoBMP
    Dim lngTemp    As Long
    Dim Encontrado As Boolean
    Dim Existe     As Boolean
    Dim BitCount   As Integer
    Existe = True
    'Set up the surface desc
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    
    If useVideoMemory Then
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY

    End If
    
    'Call Decryptdataex(fileIndex, datos) ' pq es mas lenta? :S
    
    If ResourceFile = 1 And ExisteBMP(fileIndex) = ResourceFile Then
        filePath = GrhPath & CStr(fileIndex) & ".bmp"

        If Not FileExist(filePath, vbNormal) Then Exit Function
        Call surfaceDimensions(filePath, ddsd.lHeight, ddsd.lWidth, BitCount)
    ElseIf ResourceFile = 2 And ExisteBMP(fileIndex) = ResourceFile Then
        'Call Decryptdata(fileIndex, datos)
        ddsd.lWidth = datos.bmpInfo.bmiHeader.biWidth
        ddsd.lHeight = datos.bmpInfo.bmiHeader.biHeight
    ElseIf ResourceFile = 3 And ExisteBMP(fileIndex) > 0 Then

        If ExisteBMP(fileIndex) = 1 Then
            Encontrado = True
            filePath = GrhPath & CStr(fileIndex) & ".bmp"

            If Not FileExist(filePath, vbNormal) Then Exit Function
            Call surfaceDimensions(filePath, ddsd.lHeight, ddsd.lWidth, BitCount)
        ElseIf ExisteBMP(fileIndex) = 2 Then
            'Call Decryptdata(fileIndex, datos)
            ddsd.lWidth = datos.bmpInfo.bmiHeader.biWidth
            ddsd.lHeight = datos.bmpInfo.bmiHeader.biHeight
            Encontrado = False

        End If

    Else
        Existe = False
        ddsd.lWidth = 50
        ddsd.lHeight = 50

    End If

    With newSurface
        .fileIndex = fileIndex
        
        'Set last access time (if we didn't we would reckon this texture as the one lru)
        .lastAccess = GetTickCount
        
        'Load surface
        
        If ResourceFile = 2 And Existe Then
            Set .Surface = DirectDraw.CreateSurface(ddsd)
            lngTemp = .Surface.GetDC
            StretchDIBits lngTemp, 0, 0, datos.bmpInfo.bmiHeader.biWidth, datos.bmpInfo.bmiHeader.biHeight, _
                    0, 0, datos.bmpInfo.bmiHeader.biWidth, datos.bmpInfo.bmiHeader.biHeight, datos.BMPData( _
                    0), datos.bmpInfo, DIB_RGB_COLORS, SRCCOPY
            .Surface.ReleaseDC lngTemp
        ElseIf ResourceFile = 3 And Existe Then

            If Encontrado Then
                Set .Surface = DirectDraw.CreateSurfaceFromFile(filePath, ddsd)
            Else
                Set .Surface = DirectDraw.CreateSurface(ddsd)
                lngTemp = .Surface.GetDC
                StretchDIBits lngTemp, 0, 0, datos.bmpInfo.bmiHeader.biWidth, _
                        datos.bmpInfo.bmiHeader.biHeight, 0, 0, datos.bmpInfo.bmiHeader.biWidth, _
                        datos.bmpInfo.bmiHeader.biHeight, datos.BMPData(0), datos.bmpInfo, DIB_RGB_COLORS, _
                        SRCCOPY
                .Surface.ReleaseDC lngTemp

            End If

        ElseIf ResourceFile = 1 And Existe Then
            Set .Surface = DirectDraw.CreateSurfaceFromFile(filePath, ddsd)
        Else
            Set .Surface = DirectDraw.CreateSurface(ddsd)
            Dim r As RECT
            Call .Surface.BltColorFill(r, vbGreen)

        End If
        
        'Set colorkey
        ddck.high = 0
        ddck.low = 0
        Call .Surface.SetColorKey(DDCKEY_SRCBLT, ddck)
        
        'Retrieve the updated surface desc
        Call .Surface.GetSurfaceDesc(ddsd)

    End With
    
    'Insert surface to the list at the given pos
    Call InsertSurface(newSurface, listIndex)
    
    'Update used bytes
    usedBytes = usedBytes + ddsd.lHeight * ddsd.lPitch
    
    Dim removedFile As Integer

    'Check if we have exceeded our allowed share of memory usage
    Do While usedBytes > maxBytesToUse And surfaceCount > 1
        'Remove a file
        removedFile = RemoveLRU
        
        'If no file could be removed we continue, if the file was previous to our surface we update the index
        If removedFile = 0 Then
            Exit Do
        ElseIf removedFile < listIndex Then
            listIndex = listIndex - 1

        End If

    Loop
    
    'Return the real index in wich it ended after removing any necessary files
    LoadSurface = listIndex
    Exit Function

ErrHandler:

    If Err.Number = DDERR_OUTOFMEMORY Or Err.Number = DDERR_OUTOFVIDEOMEMORY Then
        If surfaceCount Then
            'Remove a surface and try again
            Call RemoveLRU

            Resume Next

        Else
            MsgBox _
                    "No hay memoria disponible! El programa abortar�. Cierra algunos programas e intenta de nuevo"
            End

        End If

    Else
        MsgBox "Un error inesperado ocurri� al intentar cargar el gr�fico " & filePath & ". " & vbCrLf & _
                "El c�digo de error es " & CStr(Err.Number) & vbCrLf & _
                "Copia este mensaje y notifica a los administradores.", vbOKOnly Or vbCritical Or _
                vbExclamation, "Error"
        End

    End If

End Function

Private Sub surfaceDimensions(ByVal Archivo As String, _
                              ByRef Height As Long, _
                              ByRef Width As Long, _
                              ByRef BitCount As Integer)
    '**************************************************************
    'Author: Juan Mart�n Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    'Loads the headers of a bmp file to retrieve it's dimensions at rt
    '**************************************************************
    Dim handle      As Integer
    Dim bmpFileHead As BITMAPFILEHEADER
    Dim bmpInfoHead As BITMAPINFOHEADER
    
    handle = FreeFile()
    Open Archivo For Binary Access Read Lock Write As handle
    Get handle, , bmpFileHead
    Get handle, , bmpInfoHead
    Close handle
    
    Height = bmpInfoHead.biHeight
    Width = bmpInfoHead.biWidth
    BitCount = bmpInfoHead.biBitCount

End Sub

Private Sub InsertSurface(ByRef Surface As SURFACE_ENTRY_DYN, ByVal listIndex As Long)

    '**************************************************************
    'Author: Juan Mart�n Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    'Inserts the given surface in the requested position of the surface list
    '**************************************************************
    On Error GoTo ErrHandler

    Dim i As Long
    
    'Search for an empty spot in the list
    For i = 0 To surfaceCount - 1

        If surfaceList(i).Surface Is Nothing Then Exit For
    Next i
    
    'Enlarge the list if no empty spot was found
    If i = surfaceCount Then
        ReDim Preserve surfaceList(surfaceCount) As SURFACE_ENTRY_DYN
        
        'Increase surface count
        surfaceCount = surfaceCount + 1

    End If
    
    'Insert the new surface
    surfaceList(i) = Surface
    
    'Resize the list
    ReDim Preserve surfaceIndexes(surfaceIndexCount) As Long
    
    'Update the index list
    If surfaceIndexCount > listIndex Then
        'Move back the list - Copying this way is up to 6 times faster than a For
        Dim tempList() As Long
        ReDim tempList(surfaceIndexCount - listIndex) As Long
        
        CopyMemory tempList(0), surfaceIndexes(listIndex), (surfaceIndexCount - listIndex) * 4
        surfaceIndexes(listIndex) = i
        CopyMemory surfaceIndexes(listIndex + 1), tempList(0), (surfaceIndexCount - listIndex) * 4
    Else
        'We are inserting at the bottom of the list
        surfaceIndexes(listIndex) = i

    End If
    
    surfaceIndexCount = surfaceIndexCount + 1
    Exit Sub

ErrHandler:
    MsgBox "Un error irreparable ocurri� al insertar un nuevo gr�fico en la lista." & vbCrLf & _
            "El cliente se cerrar�" & vbCrLf & "Intente usar el cliente no din�mico"
    End

End Sub

Private Function RemoveLRU() As Integer
    '**************************************************************
    'Author: Juan Mart�n Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    'Removes the Least Recently Used surface to make some room for new ones
    '**************************************************************
    Dim LRU  As Long
    Dim i    As Long
    Dim ddsd As DDSURFACEDESC2
    
    'Should never happen, but just in case....
    If surfaceCount = 0 Then Exit Function
    
    'Initialize with the first element of the list
    LRU = 1
    
    'Check out through the whole list for the least recently used
    For i = 1 To surfaceIndexCount - 1

        If surfaceList(surfaceIndexes(LRU)).lastAccess > surfaceList(surfaceIndexes(i)).lastAccess Then
            LRU = i

        End If

    Next i
    
    'Store the index of the surface removed
    RemoveLRU = LRU
    
    'Retrieve the surface desc
    Call surfaceList(surfaceIndexes(LRU)).Surface.GetSurfaceDesc(ddsd)
    
    'Remove it
    Set surfaceList(surfaceIndexes(LRU)).Surface = Nothing
    surfaceList(surfaceIndexes(LRU)).fileIndex = 0
    
    'Move back the list (if necessary)
    If LRU < surfaceIndexCount - 1 Then
        CopyMemory surfaceIndexes(LRU), surfaceIndexes(LRU + 1), (surfaceIndexCount - LRU - 1) * 4

    End If
    
    'Resize the list
    ReDim Preserve surfaceIndexes(surfaceIndexCount - 1) As Long
    
    'Decrease index count
    surfaceIndexCount = surfaceIndexCount - 1
    
    'Update the used bytes
    usedBytes = usedBytes - ddsd.lHeight * ddsd.lPitch

End Function

Private Function FileSize(lngWidth As Long, lngHeight As Long) As Long

    'Return the size of the image portion of the bitmap
    If lngWidth Mod 4 > 0 Then
        FileSize = ((lngWidth \ 4) + 1) * 4 * lngHeight - 1
    Else
        FileSize = lngWidth * lngHeight - 1

    End If

End Function

Private Function IncrPos(ByRef lngOffset As Long, _
                         ByRef lngShift As Long, _
                         ByRef lngIncrement As Long) As Long
    'RETURNS: The length of the bytes to copy, while increasing the offset by the length of bytes copied

    lngOffset = lngOffset + lngShift
    lngShift = lngIncrement
    IncrPos = lngIncrement
    
End Function

Public Sub clsSurfaceManager_BorrarTodo()
    'Vacia la lista de graficos
    Dim i As Long
    
    'Destroy every surface in memory
    For i = 0 To surfaceCount - 1
        Set surfaceList(i).Surface = Nothing
    Next i
    
    'Destroy the arrays
    Erase surfaceList
    Erase surfaceIndexes
    usedBytes = 0
    surfaceCount = 0
    surfaceIndexCount = 0
    ReDim surfaceList(0) As SURFACE_ENTRY_DYN
    ReDim surfaceIndexes(0) As Long

End Sub

Private Function AlignScan(ByVal inWidth As Long, ByVal inDepth As Integer) As Long
    AlignScan = (((inWidth * inDepth) + &H1F) And Not &H1F&) \ &H8

End Function
