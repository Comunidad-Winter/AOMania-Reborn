Attribute VB_Name = "modCompressionMaTeO"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef source As Any, ByVal byteCount As Long)

Public Sub ComprimirArchivos(ByRef KeyCript As String)
Dim Archivos() As String * 255
Dim nfile As Integer
Dim LengthArchivos As Long
Dim i As Long
Dim MainPointer As Long
Dim SourceFileName As String


SourceFileName = Dir(App.Path & "\" & KeyCript & "\" & "*.*", vbNormal)
    
' Create list of all files to be compressed
While SourceFileName <> ""
    ReDim Preserve Archivos(LengthArchivos) As String * 255
    Archivos(LengthArchivos) = SourceFileName
    LengthArchivos = LengthArchivos + 1
    Debug.Print SourceFileName & "."
    'Search new file
    SourceFileName = Dir()
Wend

ReDim Pointer(LengthArchivos) As Long

nfile = FreeFile

Dim LengthTotal As Long

LengthTotal = 4 + 259 * LengthArchivos
For i = 0 To LengthArchivos - 1
    LengthTotal = LengthTotal + FileLen(App.Path & "\" & KeyCript & "\" & Replace(Archivos(i), " ", ""))
Next i

Dim data() As Byte
Dim Pointer As Long

ReDim data(LengthTotal) As Byte

Open App.Path & "\" & KeyCript & ".revival" For Binary As #nfile
    Put #nfile, , LengthArchivos
    
    MainPointer = MainPointer + 4 + 259 * LengthArchivos
    
    For i = 0 To LengthArchivos - 1
        Put #nfile, , Archivos(i)
        Put #nfile, , MainPointer
        MainPointer = MainPointer + FileLen(App.Path & "\" & KeyCript & "\" & Replace(Archivos(i), " ", ""))
    Next i
Close #nfile
End Sub
