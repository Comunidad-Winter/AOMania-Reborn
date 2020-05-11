Attribute VB_Name = "Declaraciones"
Option Explicit

Public Declare Function writeprivateprofilestring _
               Lib "kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                   ByVal lpKeyname As Any, _
                                                   ByVal lpString As String, _
                                                   ByVal lpfilename As String) As Long

Public grhCount As Long
Public fileVersion As Long

Public TilePixelHeight        As Integer
Public TilePixelWidth         As Integer

Public Declare Function getprivateprofilestring _
               Lib "kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                 ByVal lpKeyname As Any, _
                                                 ByVal lpdefault As String, _
                                                 ByVal lpreturnedstring As String, _
                                                 ByVal nsize As Long, _
                                                 ByVal lpfilename As String) As Long

Public Type GrhData

    sX          As Integer
    sY          As Integer
    FileNum     As Long
    pixelWidth  As Integer
    pixelHeight As Integer
    TileWidth   As Single
    TileHeight  As Single
   
    NumFrames   As Integer
    Frames()    As Long
    Speed       As Single

End Type

Public GrhData()              As GrhData

Public Type tObjData
      
      GrhIndex As Long
      ObjType As Integer
      NumObj As Integer
      
End Type

Public ObjData() As tObjData

Public Enum eObjType
      Arma = 2
      Armadura = 3
      Escudo = 16
      Casco = 17
      Alas = 37

End Enum
