Attribute VB_Name = "Sounds"
Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long

Public Sub SoundCarga()
     Dim sPath As String
     Dim lRet As Long
     sPath = Chr(34) & App.Path + "\WAV\158.wav" & Chr(34)
     lRet = mciSendString("OPEN " & sPath, 0&, 0, 0)
     lRet = mciSendString("PLAY " & sPath & " FROM 0", 0&, 0, 0)
End Sub
