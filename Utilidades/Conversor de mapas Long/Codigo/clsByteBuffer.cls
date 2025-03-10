VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsByteBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (ByRef destination As Any, _
                                       ByRef source As Any, _
                                       ByVal length As Long)

Private Data()                      As Byte

Private CurrentPos                  As Long
Private LastPos                     As Long
Private hFile                       As Long

Private Const INVALID_FILE_HANDLER  As Integer = -1
Private Const DEFAULT_MAX_SIZE_FILE As Long = 65535
Private Const BYTE_SIZE             As Byte = 1
Private Const BOOL_SIZE             As Byte = 2
Private Const INTEGER_SIZE          As Byte = 2
Private Const LONG_SIZE             As Byte = 4
Private Const SINGLE_SIZE           As Byte = 4
Private Const DOUBLE_SIZE           As Byte = 8
Private Const STRING_LENGTH_SIZE    As Byte = 2

Public Sub initializeReader(ByRef arrayByte() As Byte)

    LastPos = UBound(arrayByte)
    ReDim Data(LastPos) As Byte

    Call CopyMemory(Data(0), arrayByte(0), LastPos + 1)
    CurrentPos = 0

End Sub

Public Sub initializeWriter(ByVal fileHandler As Integer)

    ReDim Data(DEFAULT_MAX_SIZE_FILE * 20) As Byte

    hFile = fileHandler
    CurrentPos = 0
    LastPos = -1

End Sub
    
Public Sub getBytes(ByRef destination() As Byte, Optional ByVal length As Long = -1)

    If length >= 0 Then
        Call CopyMemory(destination(0), Data(CurrentPos), length)
    Else
        Call CopyMemory(destination(0), Data(0), LastPos + 1)

    End If

End Sub

Public Function getByte() As Byte

    'If getByte = 0 Then Exit Function
    'Call CopyMemory(getByte, data(CurrentPos), BYTE_SIZE)
    getByte = Data(CurrentPos)
    CurrentPos = CurrentPos + BYTE_SIZE

End Function

Public Function getBoolean() As Boolean

    Call CopyMemory(getBoolean, Data(CurrentPos), BOOL_SIZE)
    CurrentPos = CurrentPos + BOOL_SIZE

End Function

Public Function getInteger() As Integer

    Call CopyMemory(getInteger, Data(CurrentPos), INTEGER_SIZE)
    CurrentPos = CurrentPos + INTEGER_SIZE

End Function

Public Function getLong() As Long

    Call CopyMemory(getLong, Data(CurrentPos), LONG_SIZE)
    CurrentPos = CurrentPos + LONG_SIZE

End Function

Public Function getSingle() As Single

    Call CopyMemory(getSingle, Data(CurrentPos), SINGLE_SIZE)
    CurrentPos = CurrentPos + SINGLE_SIZE

End Function

Public Function getDouble() As Double

    Call CopyMemory(getDouble, Data(CurrentPos), DOUBLE_SIZE)
    CurrentPos = CurrentPos + DOUBLE_SIZE

End Function

Public Function getString(Optional ByVal length As Integer = -1) As String

    Dim ret() As Byte

    If length < 0 Then
        length = getInteger()
        getString = getString(length)
    Else

        If length > 0 Then
            ReDim ret(length - 1) As Byte
        
            Call CopyMemory(ret(0), Data(CurrentPos), length)
        
            getString = StrConv(ret, vbUnicode)
            CurrentPos = CurrentPos + length

        End If

    End If

End Function

Public Sub putByte(ByVal value As Byte)

    Data(LastPos + 1) = value
    LastPos = LastPos + BYTE_SIZE

End Sub

Public Sub putBoolean(ByVal value As Boolean)

    Call CopyMemory(Data(LastPos + 1), value, BOOL_SIZE)
    LastPos = LastPos + BOOL_SIZE

End Sub

Public Sub putInteger(ByVal value As Integer)

    Call CopyMemory(Data(LastPos + 1), value, INTEGER_SIZE)
    LastPos = LastPos + INTEGER_SIZE

End Sub

Public Sub putLong(ByVal value As Long)

    Call CopyMemory(Data(LastPos + 1), value, LONG_SIZE)
    LastPos = LastPos + LONG_SIZE

End Sub

Public Sub putSingle(ByVal value As Single)

    Call CopyMemory(Data(LastPos + 1), value, SINGLE_SIZE)
    LastPos = LastPos + SINGLE_SIZE

End Sub

Public Sub putDouble(ByVal value As Double)

    Call CopyMemory(Data(LastPos + 1), value, DOUBLE_SIZE)
    LastPos = LastPos + DOUBLE_SIZE

End Sub

Public Sub putString(ByRef Str As String, Optional ByVal withLength As Boolean = True)

    Dim length As Long

    length = Len(Str)

    If withLength Then
        Call putInteger(length)
        Call putString(Str, False)
    Else

        If length > 0 Then
            Call CopyMemory(Data(LastPos + 1), ByVal StrPtr(StrConv(Str, vbFromUnicode)), length)
    
            LastPos = LastPos + length

        End If

    End If

End Sub

Public Sub getVoid(ByVal length As Integer)

    CurrentPos = CurrentPos + length

End Sub

Public Sub putVoid(ByVal length As Integer)

    LastPos = LastPos + length

End Sub

Public Sub clearData()

    ReDim Data(DEFAULT_MAX_SIZE_FILE) As Byte

    CurrentPos = 0
    LastPos = -1
    hFile = -1

End Sub

Public Function getLastPos() As Long

    getLastPos = LastPos

End Function

Public Function getCurrentPos() As Long

    getCurrentPos = CurrentPos

End Function

Public Function EOF() As Boolean

    EOF = (CurrentPos > UBound(Data))

End Function

Public Sub saveBuffer()

    Dim buf() As Byte

    If hFile > 0 Then
        ReDim buf(LastPos) As Byte
    
        Call CopyMemory(buf(0), Data(0), LastPos + 1)
        Put hFile, , buf

    End If

End Sub

Private Sub Class_Initialize()

    hFile = INVALID_FILE_HANDLER

End Sub

Private Sub Class_Terminate()

    Erase Data()

End Sub

