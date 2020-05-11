Attribute VB_Name = "General"
' LauncherAoM 1.0.0
' By Bassinger [www.AoMania.Net]

Option Explicit

Sub Main()
     
     Call InitializeCompression
     Call LoadIconos
     
     frmMain.Show
      
End Sub

Sub UnloadAllForms()
    
    Dim mifrm As Form
    
    For Each mifrm In Forms

        Unload mifrm
    Next

End Sub

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (LenB(Dir$(File, FileType)) <> 0)

End Function

Function DirLibs()
     DirLibs = App.Path & "\libs\"
End Function

Function DirConf()
      DirConf = App.Path & "\libs\Configuracion\"
End Function
