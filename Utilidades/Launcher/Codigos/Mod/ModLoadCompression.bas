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
      MBUpdate As StdPicture
      Fondo_Principal As StdPicture
      NoPlay As StdPicture
      Play As StdPicture
      Online As StdPicture
      Offline As StdPicture
      Notice1 As StdPicture
      Notice2 As StdPicture
      Notice3 As StdPicture
End Type

Public Interfaces As PicInterfaces
Public Iconos As PicIconos

Sub LoadInterfaces()
       
       Dim Data() As Byte
       
       With Interfaces
       
           If Get_File_Data(DirLibs, "FLAUNCH.JPG", Data, INT_RESOURCE_FILE) Then
              Set .Fondo_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)

          End If

          Erase Data
          
          If Get_File_Data(DirLibs, "MBUPDATE.JPG", Data, INT_RESOURCE_FILE) Then
              Set .MBUpdate = ArrayToPicture(Data(), 0, UBound(Data) + 1)

          End If

          Erase Data
          
          If Get_File_Data(DirLibs, "BCVACIA.JPG", Data, INT_RESOURCE_FILE) Then
              Set .BVacia = ArrayToPicture(Data(), 0, UBound(Data) + 1)

          End If

          Erase Data
          
          If Get_File_Data(DirLibs, "BCLLENA.JPG", Data, INT_RESOURCE_FILE) Then
              Set .BLlena = ArrayToPicture(Data(), 0, UBound(Data) + 1)

          End If

          Erase Data
          
          If Get_File_Data(DirLibs, "BTNPLAY.JPG", Data, INT_RESOURCE_FILE) Then
              Set .NoPlay = ArrayToPicture(Data(), 0, UBound(Data) + 1)

          End If

          Erase Data
          
          If Get_File_Data(DirLibs, "BTPLAY.JPG", Data, INT_RESOURCE_FILE) Then
              Set .Play = ArrayToPicture(Data(), 0, UBound(Data) + 1)

          End If

          Erase Data
          
          If Get_File_Data(DirLibs, "ONLINE.JPG", Data, INT_RESOURCE_FILE) Then
              Set .Online = ArrayToPicture(Data(), 0, UBound(Data) + 1)

          End If

          Erase Data
          
          If Get_File_Data(DirLibs, "OFFLINE.JPG", Data, INT_RESOURCE_FILE) Then
              Set .Offline = ArrayToPicture(Data(), 0, UBound(Data) + 1)

          End If

          Erase Data
          
          If Get_File_Data(DirLibs, "NOTICE1.JPG", Data, INT_RESOURCE_FILE) Then
              Set .Notice1 = ArrayToPicture(Data(), 0, UBound(Data) + 1)

          End If

          Erase Data
          
        If Get_File_Data(DirLibs, "NOTICE2.JPG", Data, INT_RESOURCE_FILE) Then
              Set .Notice2 = ArrayToPicture(Data(), 0, UBound(Data) + 1)

          End If

          Erase Data
          
        If Get_File_Data(DirLibs, "NOTICE3.JPG", Data, INT_RESOURCE_FILE) Then
              Set .Notice3 = ArrayToPicture(Data(), 0, UBound(Data) + 1)

          End If

          Erase Data
          
       End With
       
End Sub

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
