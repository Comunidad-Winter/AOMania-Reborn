Attribute VB_Name = "ES"
Option Explicit

Public Sub LoadConfig()
     
     Dim Leer As New clsIniManager
     
     Call Leer.Initialize(DirConf & "Launcher.dat")
     
     With Launcher
          .Play = Val(Leer.GetValue("CONFIG", "Play"))
          .Update = Val(Leer.GetValue("CONFIG", "Update"))
          .Use = Val(Leer.GetValue("CONFIG", "Use"))
     End With
     
     Set Leer = Nothing
     
End Sub

Public Sub SaveConfig()
      
      Dim Leer As New clsIniManager
      
      Call Leer.Initialize(DirConf & "Launcher.dat")
      
      With Launcher
         Call Leer.ChangeValue("CONFIG", "Play", .Play)
         Call Leer.ChangeValue("CONFIG", "Update", .Update)
         Call Leer.ChangeValue("CONFIG", "Use", .Use)
      End With
         
      Call Leer.DumpFile(DirConf & "Launcher.dat")
      
      Set Leer = Nothing
      
End Sub
