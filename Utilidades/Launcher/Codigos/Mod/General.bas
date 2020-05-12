Attribute VB_Name = "General"
' LauncherAoM 1.0.0
' By Bassinger [www.AoMania.Net]

Option Explicit

Private Const FLAG_ICC_FORCE_CONNECTION = &H1

Private Declare Function InternetCheckConnection Lib _
    "wininet.dll" Alias _
    "InternetCheckConnectionA" ( _
        ByVal lpszUrl As String, _
        ByVal dwFlags As Long, _
        ByVal dwReserved As Long) As Long

Sub Main()
     
     Call InitializeCompression
     
     Call LoadIconos
     Call LoadInterfaces
     Call LoadConfig
     
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

Public Function FileUpdate()
      
      FileUpdate = App.Path & "\Libs\Configuracion\Update.INI"
      
End Function

Public Function Url_Path() As String
        Url_Path = "http://argentumania.es/cosas/parches/"
End Function

Function Comprobar_Conexión(Url As String) As Boolean
      
    If LCase(Left(Url, 7)) <> "http://" Then
         
       Url = "http://" & Url
    End If
      
    If InternetCheckConnection(Url, FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
       Comprobar_Conexión = False
    Else
       Comprobar_Conexión = True
    End If
      
End Function
