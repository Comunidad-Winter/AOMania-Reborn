VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Const FRENTE = 1

Private Cola As Collection

Public Sub Reset()

    On Error Resume Next

    Dim i As Integer

    For i = 1 To Me.Longitud
        Cola.Remove FRENTE
    Next i

End Sub

Public Property Get Longitud() As Integer

    Longitud = Cola.Count

End Property

Private Function IndexValido(ByVal i As Integer) As Boolean

    IndexValido = i >= 1 And i <= Me.Longitud

End Function

Private Sub Class_Initialize()

    Set Cola = New Collection

End Sub

Public Function VerElemento(ByVal Index As Integer) As Variant

    On Error Resume Next

    If IndexValido(Index) Then
        VerElemento = Cola(Index)
    Else
        VerElemento = 0

    End If

End Function

Public Sub Push(ByVal info As String, ByVal nombre As String)

    On Error Resume Next

    'Mete elemento en la cola
    Call Cola.Add(UCase$(nombre))

End Sub

Public Function Pop() As Variant

    On Error Resume Next

    'Quita elemento de la cola
    If Cola.Count > 0 Then
        Pop = Cola(FRENTE)
        Call Cola.Remove(FRENTE)
    Else
        Pop = 0

    End If

End Function

Public Function PopByVal() As Variant

    On Error Resume Next

    'Call LogTarea("PopByVal SOS")

    'Quita elemento de la cola
    If Cola.Count > 0 Then
        PopByVal = Cola.Item(1)
    Else
        PopByVal = 0

    End If

End Function

Public Function Existe(ByVal nombre As String) As Boolean

    On Error Resume Next
  
    Dim f                  As Boolean
    Dim v                  As Variant
    Dim i                  As Integer
    Dim NombreEnMayusculas As String
    NombreEnMayusculas = UCase$(nombre)
 
    i = 1
    f = False
 
    For i = 1 To Me.Longitud
        v = Me.VerElemento(i)

        If v = NombreEnMayusculas Then
            Existe = True
            Exit Function

        End If

    Next

End Function

Public Sub Quitar(ByVal nombre As String)

    On Error Resume Next

    Dim v                  As Variant
    Dim i                  As Integer
    Dim NombreEnMayusculas As String
 
    NombreEnMayusculas = UCase$(nombre)
 
    For i = 1 To Me.Longitud
        v = Me.VerElemento(i)

        If v = NombreEnMayusculas Then
            Call Cola.Remove(i)
            Exit Sub

        End If

    Next i
 
End Sub

Public Sub QuitarIndex(ByVal Index As Integer)

    On Error Resume Next

    If IndexValido(Index) Then Call Cola.Remove(Index)

End Sub

Private Sub Class_Terminate()

    'Destruimos el objeto Cola
    Set Cola = Nothing

End Sub
