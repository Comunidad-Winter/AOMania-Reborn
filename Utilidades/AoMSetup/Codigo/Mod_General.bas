Attribute VB_Name = "Mod_General"
Option Explicit

Public Sub LeerSetup()
  On Error Resume Next
    If FileExist(App.Path & "\AOM.cfg", vbArchive) Then
        
        Dim handle As Integer
        handle = FreeFile
        
        Open App.Path & "\AOM.cfg" For Binary As handle
            Get handle, , AoSetup
        Close handle
        
   FrmMain.ChkTransparencia.value = AoSetup.bTransparencia
   FrmMain.ChkMusic.value = AoSetup.bMusica
   FrmMain.ChkSonidos.value = AoSetup.bSonido
   FrmMain.ChkPantalla.value = AoSetup.bResolucion
   FrmMain.ChkEjecutar.value = AoSetup.bEjecutar
   End If
End Sub

