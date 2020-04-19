Attribute VB_Name = "Declaraciones"
Option Explicit

Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Type tBrujo
    Inteligencia As Byte
    Name As String
    Nivel As Byte
    Mana As Integer
    
End Type

Public Brujo(1 To 1000) As tBrujo

Public NumBrujos As Integer
