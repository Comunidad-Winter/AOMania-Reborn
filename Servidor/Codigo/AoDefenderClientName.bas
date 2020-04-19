Attribute VB_Name = "AoDefenderClientName"
Public AoDefOriginalClientName As String
Public AoDefClientName As String
Public Function AoDefChangeName() As Boolean
If AoDefOriginalClientName <> AoDefClientName Then
AoDefChangeName = True
Exit Function
End If
AoDefChangeName = False
End Function
Public Sub AoDefClientOn()
MsgBox "Se ha detectado cambio de nombre en el ejecutable. No es posible ejecutar el cliente!.", vbCritical, "AoDefender"
End Sub

