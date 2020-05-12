Attribute VB_Name = "ModLoadCompression"
' LauncherAoM 1.0.0
' By Bassinger [www.AoMania.Net]

Option Explicit

Public Type PicIconos
     Diablo As StdPicture
     Ico_Diablo As Picture
     Mano As StdPicture
     Ico_Mano As Picture
End Type

Public Type PicInterfaces
      BVacia As StdPicture
      BLlena As StdPicture
      
      Vacia As Picture
End Type

Public Interfaces As PicInterfaces
Public Iconos As PicIconos

Sub LoadIconos()
     
     Dim Data() As Byte
     
     With Iconos
            
            If Get_File_Data(DirLibs, "DIABLO.ICO", Data, ICONOS_FILE) Then
                Set .Diablo = ArrayToPicture(Data(), 0, UBound(Data) + 1)
                Set .Ico_Diablo = .Diablo
            End If
            
            Erase Data
            
            If Get_File_Data(DirLibs, "MANO.ICO", Data, ICONOS_FILE) Then
               Set .Mano = ArrayToPicture(Data(), 0, UBound(Data) + 1)
               Set .Ico_Mano = .Mano
            End If
            
            Erase Data
            
     End With
     
End Sub

Sub LoadInterfaces()

     Dim Data() As Byte
     
    With Interfaces
             
            If Get_File_Data(DirLibs, "AU_BARRAVOID.jpg", Data, INT_RESOURCE_FILE) Then
                 Set .BVacia = ArrayToPicture(Data(), 0, UBound(Data) + 1)
            End If
            
            Erase Data
            
            If Get_File_Data(DirLibs, "BLLENANEW.jpg", Data, INT_RESOURCE_FILE) Then
                 Set .BLlena = ArrayToPicture(Data(), 0, UBound(Data) + 1)
            End If
            
            Erase Data
            
     End With
     
End Sub
