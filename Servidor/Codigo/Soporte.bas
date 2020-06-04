Attribute VB_Name = "Soporte"
Public xaoindex As Integer

Public Sub MostrarSop(ByVal UserIndex As Integer, ByVal marika As Integer, ByVal nombre As String)

    xaoindex = NameIndex(nombre)

    If xaoindex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario se encuentra offline." & FONTTYPE_INFO)
        Exit Sub
    Else
        SendData SendTarget.ToIndex, UserIndex, 0, "SOPO" & UserList(marika).Pregunta & Chr$(2) & UserList(marika).Name

    End If

End Sub

Public Sub EnviarResp(ByVal UserIndex As Integer)

    SendData SendTarget.ToIndex, UserIndex, 0, "RESP" & UserList(UserIndex).Respuesta

End Sub

Public Sub ResetSop(ByVal UserIndex As Integer)

    UserList(UserIndex).Pregunta = "Ninguna"
    UserList(UserIndex).Respuesta = "Ninguna"
    UserList(UserIndex).flags.Soporteo = False

End Sub

