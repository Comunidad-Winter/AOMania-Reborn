Attribute VB_Name = "Mod_Declaraciones"
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
                                                   
Public NumReg As Long
Public ValReg As Long

    Public Nombre As String
    Public Password As String
    Public Email As String
    Public Clase As String
    Public Raza As String


Function GetVar(ByVal File As String, _
                ByVal Main As String, _
                ByVal Var As String, _
                Optional EmptySpaces As Long = 1024) As String

    Dim sSpaces  As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found
  
    szReturn = ""
  
    sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
  
    GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
  
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub WriteVar(ByVal File As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal value As String)
    '*****************************************************************
    'Escribe VAR en un archivo
    '*****************************************************************

    writeprivateprofilestring Main, Var, value, File
    
End Sub



