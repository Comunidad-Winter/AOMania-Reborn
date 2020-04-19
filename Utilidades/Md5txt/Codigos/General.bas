Attribute VB_Name = "General"
Option Explicit

Public fdpc               As Long

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Sub Disco()

    Dim Fso As New Scripting.FileSystemObject
    Dim dr  As Scripting.Drive
    Set dr = Fso.GetDrive("c:")
    fdpc = Abs(dr.SerialNumber)

End Sub
