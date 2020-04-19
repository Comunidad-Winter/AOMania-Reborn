Attribute VB_Name = "Mod_General"
Option Explicit

Public Type PicInterfaces
      BVacia As StdPicture
      BLlena As StdPicture
      
      Vacia As Picture
End Type

Public Type PicIconos
      Mano As StdPicture
      Diablo As StdPicture
      
      Ico_Mano As Picture
      Ico_Diablo As Picture
End Type

Public Interfaces As PicInterfaces
Public Iconos As PicIconos

Sub Main()
  DoEvents
   Call InitializeCompression
   Call LoadInterfaces
   Call LoadIconos
  DoEvents
  
  Call frmUpdate.Show
End Sub

Sub LoadInterfaces()

     Dim Data() As Byte
     
    With Interfaces
             
            If Get_File_Data(DirRecursos, "AU_BARRAVOID.jpg", Data, INT_RESOURCE_FILE) Then
                 Set .BVacia = ArrayToPicture(Data(), 0, UBound(Data) + 1)
            End If
            
            Erase Data
            
            If Get_File_Data(DirRecursos, "BLLENANEW.jpg", Data, INT_RESOURCE_FILE) Then
                 Set .BLlena = ArrayToPicture(Data(), 0, UBound(Data) + 1)
            End If
            
            Erase Data
            
     End With
     
End Sub

Sub LoadIconos()

    Dim Data() As Byte
    
      With Iconos
      
           If Get_File_Data(DirRecursos, "DIABLO.ICO", Data, ICONOS_FILE) Then
                 Set .Diablo = ArrayToPicture(Data(), 0, UBound(Data) + 1)
                 Set .Ico_Diablo = .Diablo
            End If
           
           Erase Data
           
           If Get_File_Data(DirRecursos, "MANO.ICO", Data, ICONOS_FILE) Then
                 Set .Mano = ArrayToPicture(Data(), 0, UBound(Data) + 1)
                 Set .Ico_Mano = .Mano
            End If
           
           Erase Data
           
     End With
End Sub

Sub UnloadAllForms()
    
    Dim mifrm As Form
    
    For Each mifrm In Forms

        Unload mifrm
    Next

End Sub
