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
        
Private Declare Function LoadLibraryRegister Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function CreateThreadForRegister Lib "kernel32" Alias "CreateThread" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetProcAddressRegister Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibraryRegister Lib "kernel32" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Sub Main()
     
     Call InitializeCompression
     
     Call LoadIconos
     Call LoadInterfaces
     Call LoadConfig
     
     DoEvents
     
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
     DirLibs = App.path & "\libs\"
End Function

Function DirConf()
      DirConf = App.path & "\libs\Configuracion\"
End Function

Public Function FileUpdate()
      
      FileUpdate = App.path & "\Libs\Configuracion\Update.INI"
      
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

Public Sub RevDlls()
     
     Call MRRegisterLibrary(DirLibs & "MSVBVM50.DLL", "MSVBVM50")
     Call MRRegisterLibrary(DirLibs & "Captura.ocx", "CAPTURA")
     Call MRRegisterLibrary(DirLibs & "COMCTL32.OCX", "COMCTL32")
     Call MRRegisterLibrary(DirLibs & "CSWSK32.OCX", "CSWSK32")
     Call MRRegisterLibrary(DirLibs & "MSWINSCK.OCX", "MSWINSCK.OCK")
     Call MRRegisterLibrary(DirLibs & "RICHTX32.OCX", "RICHTX32")
     Call MRRegisterLibrary(DirLibs & "vbalProgBar6.ocx", "VBALPROGBAR6")
     Call MRRegisterLibrary(DirLibs & "MSINET.OCX", "MSINET")
     Call MRRegisterLibrary(DirLibs & "ieframe.dll", "IEFRAME")
     Call MRRegisterLibrary(DirLibs & "TABCTL32.OCX", "TABCTLF32")
     Call MRRegisterLibrary(DirLibs & "hook-menu-2.ocx", "HOOK-MENU-2")
     Call MRRegisterLibrary(DirLibs & "dx8vb.dll", "DX8VB")
     
     Launcher.Use = 1
      
End Sub

Public Sub MRRegisterLibrary(path$, ByVal Name As String)
       
      Shell "Regsvr32 /s" & path$, vbNormalFocus
       
      If RegSvr32(path$, False) Then
          frmMain.txtUpdate.Caption = "Se registro componente: " & Name
          Else
          frmMain.txtUpdate.Caption = "Imposible registrar componente: " & Name
      End If

End Sub

Private Function RegSvr32(ByVal FileName As String, bUnReg As Boolean) As Boolean

Dim lLib As Long
Dim lProcAddress As Long
Dim lThreadID As Long
Dim lSuccess As Long
Dim lExitCode As Long
Dim lThread As Long
Dim bAns As Boolean
Dim sPurpose As String

sPurpose = IIf(bUnReg, "DllUnregisterServer", "DllRegisterServer")

If Dir(FileName) = "" Then Exit Function

lLib = LoadLibraryRegister(FileName)
If lLib = 0 Then Exit Function

lProcAddress = GetProcAddressRegister(lLib, sPurpose)

If lProcAddress = 0 Then
   FreeLibraryRegister lLib
   Exit Function
Else
   lThread = CreateThreadForRegister(ByVal 0&, 0&, ByVal lProcAddress, ByVal 0&, 0&, lThread)
   If lThread Then
        lSuccess = (WaitForSingleObject(lThread, 10000) = 0)
        If Not lSuccess Then
           Call GetExitCodeThread(lThread, lExitCode)
           Call ExitThread(lExitCode)
           bAns = False
           Exit Function
        Else
           bAns = True
        End If
        CloseHandle lThread
        FreeLibraryRegister lLib
   End If
End If
    RegSvr32 = bAns
End Function

Public Sub ChangeStatus(ByVal Status As Byte)

        Select Case Status
            
            Case eStatus.Online
                frmMain.PicStatus.Picture = Interfaces.Online
            Exit Sub
            
            Case eStatus.Offline
               frmMain.PicStatus.Picture = Interfaces.Offline
               frmMain.LblExp.Caption = 0
               frmMain.lblOro.Caption = 0
               frmMain.lblUser.Caption = 0
            Exit Sub
            
        End Select
        
End Sub
