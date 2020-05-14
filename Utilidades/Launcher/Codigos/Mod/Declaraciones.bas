Attribute VB_Name = "Declaraciones"
' LauncherAoM 1.0.0
' By Bassinger [www.AoMania.Net]

Option Explicit

Public Declare Function ShellExecute _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long

Public Type tLauncher
       Play As Byte
       Update As Byte
       Use As Byte
End Type

Public Launcher As tLauncher

Public Const IpServidor As String = "192.171.19.138"
Public Const Puerto As Integer = "7669"

Public Enum eStatus
       
       Online = 1
       Offline = 2
       
End Enum
