Attribute VB_Name = "AoDefenderAntiSH"
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public AoDefTime As Long
Public AoDefCount As Integer
Public Sub AoDefAntiShInitialize()
AoDefTime = GetTickCount()
End Sub
Public Function AoDefAntiSh(ByVal FramesPerSec) As Boolean
If GetTickCount - AoDefTime > 350 Or GetTickCount - AoDefTime < 250 Then
        AoDefCount = AoDefCount + 1
    Else
        AoDefCount = 0
    End If
    
    If FramesPerSec < 5 Then
    AoDefCount = AoDefCount + 1
    End If
    
    If AoDefCount > 30 Then
       AoDefAntiSh = True
       Exit Function
    End If

AoDefTime = GetTickCount()
AoDefAntiSh = False
End Function

Public Sub AoDefAntiShOn()
 Call SendData("ANTISH")
MsgBox "Se ha detectado algo inusual en el cliente. Se va a cerrar por seguridad.", vbCritical, "AoMania"
End Sub

Public Function CliEditado()
    Call MsgBox("No se admite editar el cliente en este servidor")
    End
End Function


