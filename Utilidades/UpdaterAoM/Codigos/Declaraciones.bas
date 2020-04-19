Attribute VB_Name = "Declaraciones"
Option Explicit

Public Declare Function ShellExecute _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long
                                      
Public Function FileUpdate()
      
      FileUpdate = App.Path & "\Libs\Configuracion\Update.INI"
      
End Function


Public Function DirRecursos() As String

    DirRecursos = App.Path & "\Libs\"

End Function

Function FileExist(ByVal File As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
    '*****************************************************************
    'Se fija si existe el archivo
    '*****************************************************************
    FileExist = Dir$(File, FileType) <> ""

End Function
