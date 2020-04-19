Attribute VB_Name = "Declaraciones"
Option Explicit


Public Declare Function GetPrivateProfileString _
               Lib "kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                 ByVal lpKeyname As Any, _
                                                 ByVal lpdefault As String, _
                                                 ByVal lpreturnedstring As String, _
                                                 ByVal nsize As Long, _
                                                 ByVal lpfilename As String) As Long
                                                 
                                                 
Public Declare Function writeprivateprofilestring _
               Lib "kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                   ByVal lpKeyname As Any, _
                                                   ByVal lpString As String, _
                                                   ByVal lpfilename As String) As Long
                                                   

Public Type tSetupMods
    bTransparencia    As Byte
    bMusica   As Byte
    bSonido    As Byte
    bResolucion    As Byte
    bEjecutar As Byte
End Type

Public AoSetup As tSetupMods

Public Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: ?/?/?
'Se fija si existe el archivo
'*************************************************
    FileExist = Dir(file, FileType) <> ""
End Function

Function GetVar(ByVal file As String, _
                ByVal Main As String, _
                ByVal Var As String, _
                Optional EmptySpaces As Long = 1024) As String

    Dim sSpaces  As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found
  
    szReturn = ""
  
    sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
  
    GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, file
  
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub WriteVar(ByVal file As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal value As String)
    '*****************************************************************
    'Escribe VAR en un archivo
    '*****************************************************************

    writeprivateprofilestring Main, Var, value, file
    
End Sub

