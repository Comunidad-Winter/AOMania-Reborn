Attribute VB_Name = "Logs"
' LauncherAoM 1.0.0
' By Bassinger [www.AoMania.Net]

Option Explicit

Public Sub LogLauncher(ByVal desc As String)

    Dim nFile As Integer
    nFile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\ErroresLauncher.log" For Append As #nFile
    Print #nFile, desc
    Close #nFile

End Sub
