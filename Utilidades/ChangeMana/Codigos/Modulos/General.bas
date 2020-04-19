Attribute VB_Name = "General"
Option Explicit

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, Value, file
    
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = ""
  
sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
  
  
GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, file
  
GetVar = RTrim$(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Public Sub LogCambios(Name As String, Nivel As Byte, Atributos As Byte, AMana As Integer, NMana As Integer)
   On Error GoTo errhandler
   Dim nfile As Integer
   nfile = FreeFile
   Open App.Path & "\logs\Cambios.log" For Append Shared As #nfile
   Print #nfile, "> Usuario:" & Name & " || Nivel: " & Nivel & " || Atributos: " & Atributos & " || Antiguo Mana: " & AMana & " || Nueva mana " & NMana
   Close #nfile
   Exit Sub
errhandler:
End Sub

Function ReadField(ByVal iPos As Long, ByRef sText As String, ByVal CharAscii As Long) As String
' Mismo que anterior con los parametros formales...
 
    '
    ' @ maTih.-
     
    Dim Read_Field()    As String
 
 
 
    'Creo un array temporal.
    Read_Field = Split(sText, ChrW$(CharAscii))
' Mismo que antes con chrW
     
    If (iPos - 1) <= UBound(Read_Field()) Then
       'devuelve
       ReadField = (Read_Field(iPos - 1))
    End If
     
End Function
