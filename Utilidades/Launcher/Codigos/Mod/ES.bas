Attribute VB_Name = "ES"
Option Explicit

Public Sub LoadConfig()
     
     Dim Leer As New clsIniManager
     
     Call Leer.Initialize(DirConf & "Launcher.dat")
     
     With Launcher
          .Play = Val(Leer.GetValue("CONFIG", "Play"))
          .Update = Val(Leer.GetValue("CONFIG", "Update"))
     End With
     
End Sub
