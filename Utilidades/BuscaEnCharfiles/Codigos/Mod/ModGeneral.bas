Attribute VB_Name = "ModGeneral"
Option Explicit

Function ReadField(ByVal iPos As Long, _
                   ByRef sText As String, _
                   ByVal CharAscii As Long) As String
     
    Dim Read_Field() As String

    Read_Field = Split(sText, ChrW$(CharAscii))
     
    If (iPos - 1) <= UBound(Read_Field()) Then
        'devuelve
        ReadField = (Read_Field(iPos - 1))

    End If
     
End Function

Function DirChar()
   
   DirChar = App.Path & "\Charfile\"
   
End Function
