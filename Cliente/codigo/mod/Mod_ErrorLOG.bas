Attribute VB_Name = "Mod_ErrorLOG"
Option Explicit

Public Sub LogError(ByVal desc As String)

    Dim nFile As Integer
    nFile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\errores.log" For Append As #nFile
    Print #nFile, desc
    Close #nFile

End Sub

Public Sub LogCustom(ByVal desc As String)

    Dim nFile As Integer
    nFile = FreeFile ' obtenemos un canal
    Open App.Path & "\custom.log" For Append As #nFile
    Print #nFile, Now & " " & desc
    Close #nFile

End Sub

