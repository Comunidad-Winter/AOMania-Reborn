Attribute VB_Name = "Mod_TileEngine"
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


