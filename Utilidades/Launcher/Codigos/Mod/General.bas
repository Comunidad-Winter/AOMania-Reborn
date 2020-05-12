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

Public Function FormatSize(ByVal size As Currency) As String
    Const Kilobyte As Currency = 1024@
    Const HundredK As Currency = 102400@
    Const ThousandK As Currency = 1024000@
    Const Megabyte As Currency = 1048576@
    Const HundredMeg As Currency = 104857600@
    Const ThousandMeg As Currency = 1048576000@
    Const Gigabyte As Currency = 1073741824@
    Const Terabyte As Currency = 1099511627776@
    
    If size < Kilobyte Then
        FormatSize = Int(size) & " bytes"
    ElseIf size < HundredK Then
        FormatSize = Format(size / Kilobyte, "#.0") & " KB"
    ElseIf size < ThousandK Then
        FormatSize = Int(size / Kilobyte) & " KB"
    ElseIf size < HundredMeg Then
        FormatSize = Format(size / Megabyte, "#.0") & " MB"
    ElseIf size < ThousandMeg Then
        FormatSize = Int(size / Megabyte) & " MB"
    ElseIf size < Terabyte Then
        FormatSize = Format(size / Gigabyte, "#.00") & " GB"
    Else
        FormatSize = Format(size / Terabyte, "#.00") & " TB"
    End If
End Function
