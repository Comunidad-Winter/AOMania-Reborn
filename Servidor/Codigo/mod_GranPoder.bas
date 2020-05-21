Attribute VB_Name = "mod_GranPoder"
Option Explicit

Public Type tGranPoder
    Status As Byte
    TipoAura As Byte
    Cantidad As Integer
    UserIndex As Integer
    Timer As Long
    Vida As Long
    Mana As Long
    Agilidad As Long
    Fuerza As Long
End Type

Public Enum hGranPoder
    daño = 1
    Vida = 2
    Mana = 3
    Agilidad = 4
    Fuerza = 5
    Experencia = 6
End Enum

Public GranPoder As tGranPoder

Public Const FX_Poder As Integer = 88
Public Const Sound_Poder As Integer = 147

Public Sub DarGranPoder(ByVal UserIndex As Integer)
    Dim i As Integer
    Dim NewDueño As Boolean

    If NumUsers = 0 Then Exit Sub

    If UserIndex = 0 Then
        Do While NewDueño = False And i < 500
            i = i + 1
            UserIndex = RandomNumber(1, NumUsers)
            If UserList(UserIndex).flags.UserLogged = True And UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
                NewDueño = True
                Exit Do
            End If
        Loop
        If Not NewDueño Then
            UserIndex = 0
            GranPoder.TipoAura = 0
            GranPoder.Cantidad = 0
            GranPoder.Status = 0
            GranPoder.UserIndex = 0
            GranPoder.Vida = 0
            GranPoder.Mana = 0
            GranPoder.Agilidad = 0
            GranPoder.Fuerza = 0
        End If
    End If

    If UserIndex > 0 Then
        If UserList(UserIndex).flags.Muerto <> 0 Then
            GranPoder.Timer = 0
            Exit Sub
        End If
        If Not PermiteMapaPoder(UserIndex) Then
            GranPoder.Timer = 0
            Exit Sub
        End If
        Call SendData(SendTarget.ToAll, UserIndex, 0, "||" & UserList(UserIndex).Name & " poseedor del Aura de los Heroes!!!. En el mapa " & UserList(UserIndex).pos.Map & "." & FONTTYPE_GUERRA)
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & FX_Poder & "," & 1)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "TW" & Sound_Poder)
        GranPoder.Status = 1
        GranPoder.UserIndex = UserIndex
        UserList(UserIndex).GranPoder = 1
        Call DarAuraPoder(UserIndex)
    End If

End Sub

Public Function PermiteMapaPoder(ByVal UserIndex As Integer) As Boolean

    Select Case UserList(UserIndex).pos.Map

    Case 1, 20, 34, 37, 59, 60, 62, 64, 84, 86, 95, 98, 99, 100, 101, 102, 132, 149, 150, _
             154, 159, 160, 161, 162, 163, 164, 192
        PermiteMapaPoder = False
        Exit Function
    End Select

    PermiteMapaPoder = True

End Function

Private Sub DarAuraPoder(ByVal UserIndex As Integer)

    GranPoder.TipoAura = RandomNumber(1, 6)

    Select Case GranPoder.TipoAura
    Case 1    'Daño
        Call SendData(SendTarget.ToAll, UserIndex, 0, "||AURA DE LOS DIOSES ESPECIAL: " & UserList(UserIndex).Name & " INFLIDE DAÑO x2!!!" & FONTTYPE_GUERRA)
    Case 2    'Vida
        GranPoder.Cantidad = RandomNumber(1, 30)
        GranPoder.Vida = UserList(UserIndex).Stats.MaxHP + GranPoder.Cantidad
        UserList(UserIndex).Stats.MinHP = GranPoder.Vida
        UserList(UserIndex).Stats.MaxHP = GranPoder.Vida
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MXVID" & GranPoder.Vida)
        Call EnviarHP(UserIndex)
        Call SendData(SendTarget.ToAll, UserIndex, 0, "||AURA DE LOS DIOSES: " & UserList(UserIndex).Name & " ahora recibe +" & GranPoder.Cantidad & " DE VIDA!!" & FONTTYPE_GUERRA)
    Case 3    'Mana
        If UCase$(UserList(UserIndex).Clase) = "GUERRERO" Or _
           UCase$(UserList(UserIndex).Clase) = "ARQUERO" Or _
           UCase$(UserList(UserIndex).Clase) = "CAZADOR" Then
            GranPoder.Cantidad = RandomNumber(1, 30)
            GranPoder.Vida = UserList(UserIndex).Stats.MaxHP + GranPoder.Cantidad
            UserList(UserIndex).Stats.MinHP = GranPoder.Vida
            UserList(UserIndex).Stats.MaxHP = GranPoder.Vida
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MXVID" & GranPoder.Vida)
            Call EnviarHP(UserIndex)
            Call SendData(SendTarget.ToAll, UserIndex, 0, "||AURA DE LOS DIOSES: " & UserList(UserIndex).Name & " ahora recibe +" & GranPoder.Cantidad & " DE VIDA!!" & FONTTYPE_GUERRA)
            Exit Sub
        End If

        GranPoder.Cantidad = RandomNumber(1, 300)
        GranPoder.Mana = UserList(UserIndex).Stats.MaxMAN + GranPoder.Cantidad
        UserList(UserIndex).Stats.MinMAN = GranPoder.Mana
        UserList(UserIndex).Stats.MaxMAN = GranPoder.Mana
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MXMAN" & GranPoder.Mana)
        Call EnviarMn(UserIndex)
        Call SendData(SendTarget.ToAll, UserIndex, 0, "||AURA DE LOS DIOSES: " & UserList(UserIndex).Name & " ahora recibe +" & GranPoder.Cantidad & " DE MANA!!" & FONTTYPE_GUERRA)
    Case 4    'Agilidad
        GranPoder.Cantidad = RandomNumber(2, 5)
        UserList(UserIndex).flags.EspecialAgilidad = val(UserList(UserIndex).flags.EspecialAgilidad + GranPoder.Cantidad)
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = val(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + GranPoder.Cantidad)
        UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Agilidad) + GranPoder.Cantidad
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MA" & UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call EnviarAmarillas(UserIndex)
        Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
        Call SendData(SendTarget.ToAll, UserIndex, 0, "||AURA DE LOS DIOSES: " & UserList(UserIndex).Name & " ahora recibe +" & GranPoder.Cantidad & " DE AGILIDAD!!" & FONTTYPE_GUERRA)
    Case 5    'Fuerza
        GranPoder.Cantidad = RandomNumber(2, 5)
        UserList(UserIndex).flags.EspecialFuerza = val(UserList(UserIndex).flags.EspecialFuerza + GranPoder.Cantidad)
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = val(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + GranPoder.Cantidad)
        UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Fuerza) + GranPoder.Cantidad
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MF" & UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
        Call EnviarVerdes(UserIndex)
        Call SendData(SendTarget.ToAll, UserIndex, 0, "||AURA DE LOS DIOSES: " & UserList(UserIndex).Name & " ahora recibe +" & GranPoder.Cantidad & " DE FUERZA!!" & FONTTYPE_GUERRA)
    Case 6    'Experencia
        GranPoder.Cantidad = RandomNumber(10, 50)
        Call SendData(SendTarget.ToAll, UserIndex, 0, "||AURA DE LOS DIOSES: " & UserList(UserIndex).Name & " ahora recibe +" & GranPoder.Cantidad & "% EXPERIENCIA!!" & FONTTYPE_GUERRA)
    End Select

End Sub

Public Sub QuitarPoder(ByVal UserIndex As Integer)

    If UserList(UserIndex).GranPoder <> 1 Then Exit Sub

    Select Case GranPoder.TipoAura
    Case 2    ' Vida
        GranPoder.Vida = UserList(UserIndex).Stats.MaxHP - GranPoder.Cantidad
        UserList(UserIndex).Stats.MinHP = GranPoder.Vida
        UserList(UserIndex).Stats.MaxHP = GranPoder.Vida
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MXVID" & GranPoder.Vida)
        Call EnviarHP(UserIndex)
    Case 3    ' Mana
        If UCase$(UserList(UserIndex).Clase) = "GUERRERO" Or _
           UCase$(UserList(UserIndex).Clase) = "ARQUERO" Or _
           UCase$(UserList(UserIndex).Clase) = "CAZADOR" Then
            GranPoder.Vida = UserList(UserIndex).Stats.MaxHP - GranPoder.Cantidad
            UserList(UserIndex).Stats.MinHP = GranPoder.Vida
            UserList(UserIndex).Stats.MaxHP = GranPoder.Vida
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MXVID" & GranPoder.Vida)
            Call EnviarHP(UserIndex)
            GranPoder.TipoAura = 0
            GranPoder.Cantidad = 0
            GranPoder.Status = 0
            GranPoder.UserIndex = 0
            GranPoder.Timer = 0
            GranPoder.Vida = 0
            GranPoder.Mana = 0
            GranPoder.Agilidad = 0
            GranPoder.Fuerza = 0
            UserList(UserIndex).GranPoder = 0

            Call SendData(SendTarget.ToAll, UserIndex, 0, "||Los dioses le quitaron el aura a " & UserList(UserIndex).Name & "." & FONTTYPE_GUERRA)

            Exit Sub
        End If

        GranPoder.Mana = UserList(UserIndex).Stats.MaxMAN - GranPoder.Cantidad
        UserList(UserIndex).Stats.MinMAN = GranPoder.Mana
        UserList(UserIndex).Stats.MaxMAN = GranPoder.Mana
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MXMAN" & GranPoder.Mana)
        Call EnviarMn(UserIndex)
    Case 4    'Agilidad
        UserList(UserIndex).flags.EspecialAgilidad = val(UserList(UserIndex).flags.EspecialAgilidad - GranPoder.Cantidad)
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = val(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) - GranPoder.Cantidad)
        UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Agilidad) - GranPoder.Cantidad
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MA" & UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call EnviarAmarillas(UserIndex)
        Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
    Case 5    'Fuerza
        UserList(UserIndex).flags.EspecialFuerza = val(UserList(UserIndex).flags.EspecialFuerza - GranPoder.Cantidad)
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = val(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) - GranPoder.Cantidad)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MF" & UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Fuerza) - GranPoder.Cantidad
        Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
        Call EnviarVerdes(UserIndex)
    End Select

    GranPoder.TipoAura = 0
    GranPoder.Cantidad = 0
    GranPoder.Status = 0
    GranPoder.UserIndex = 0
    GranPoder.Timer = 0
    GranPoder.Vida = 0
    GranPoder.Mana = 0
    GranPoder.Agilidad = 0
    GranPoder.Fuerza = 0
    UserList(UserIndex).GranPoder = 0

    Call SendData(SendTarget.ToAll, UserIndex, 0, "||Los dioses le quitaron el aura a " & UserList(UserIndex).Name & "." & FONTTYPE_GUERRA)

End Sub

Public Sub DesconectaPoder(ByVal UserIndex As Integer)

    If UserList(UserIndex).GranPoder <> 1 Then Exit Sub

    Select Case GranPoder.TipoAura
    Case 2    'Vida
        GranPoder.Vida = UserList(UserIndex).Stats.MaxHP - GranPoder.Cantidad
        UserList(UserIndex).Stats.MinHP = GranPoder.Vida
        UserList(UserIndex).Stats.MaxHP = GranPoder.Vida
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MXVID" & GranPoder.Vida)
        Call EnviarHP(UserIndex)
    Case 3    'Mana
        If UCase$(UserList(UserIndex).Clase) = "GUERRERO" Or _
           UCase$(UserList(UserIndex).Clase) = "ARQUERO" Or _
           UCase$(UserList(UserIndex).Clase) = "CAZADOR" Then
            GranPoder.Vida = UserList(UserIndex).Stats.MaxHP - GranPoder.Cantidad
            UserList(UserIndex).Stats.MinHP = GranPoder.Vida
            UserList(UserIndex).Stats.MaxHP = GranPoder.Vida
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MXVID" & GranPoder.Vida)
            Call EnviarMn(UserIndex)
            GranPoder.TipoAura = 0
            GranPoder.Cantidad = 0
            GranPoder.Status = 0
            GranPoder.UserIndex = 0
            GranPoder.Timer = 0
            GranPoder.Vida = 0
            GranPoder.Mana = 0
            GranPoder.Agilidad = 0
            GranPoder.Fuerza = 0
            UserList(UserIndex).GranPoder = 0

            Call SendData(SendTarget.ToAll, UserIndex, 0, "||Los dioses le quitaron el aura a " & UserList(UserIndex).Name & "." & FONTTYPE_GUERRA)

            Exit Sub
        End If
        GranPoder.Mana = UserList(UserIndex).Stats.MaxMAN - GranPoder.Cantidad
        UserList(UserIndex).Stats.MinMAN = GranPoder.Mana
        UserList(UserIndex).Stats.MaxMAN = GranPoder.Mana
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MXMAN" & GranPoder.Mana)
    Case 4    'Agilidad
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = val(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) - GranPoder.Cantidad)
        UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Agilidad) - GranPoder.Cantidad
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MA" & UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call EnviarAmarillas(UserIndex)
        Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
    Case 5    'Fuerza
        UserList(UserIndex).flags.EspecialFuerza = val(UserList(UserIndex).flags.EspecialFuerza - GranPoder.Cantidad)
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = val(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) - GranPoder.Cantidad)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MF" & UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Fuerza) - GranPoder.Cantidad
        Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
        Call EnviarVerdes(UserIndex)
    End Select

    GranPoder.TipoAura = 0
    GranPoder.Cantidad = 0
    GranPoder.Status = 0
    GranPoder.UserIndex = 0
    GranPoder.Timer = 0
    GranPoder.Vida = 0
    GranPoder.Mana = 0
    GranPoder.Agilidad = 0
    GranPoder.Fuerza = 0
    UserList(UserIndex).GranPoder = 0

    Call SendData(SendTarget.ToAll, UserIndex, 0, "||" & UserList(UserIndex).Name & " ha abandonado el juego." & FONTTYPE_GUILD)

End Sub

Public Sub UserMataPoder(ByVal UserPoder As Integer, ByVal UserIndex As Integer)

    If UserList(UserPoder).GranPoder <> 1 Then Exit Sub

    Select Case GranPoder.TipoAura
    Case 2    'Vida
        GranPoder.Vida = UserList(UserPoder).Stats.MaxHP - GranPoder.Cantidad
        UserList(UserPoder).Stats.MinHP = GranPoder.Vida
        UserList(UserPoder).Stats.MaxHP = GranPoder.Vida
        Call SendData(SendTarget.ToIndex, UserPoder, 0, "MXVID" & GranPoder.Vida)
        Call EnviarHP(UserIndex)
    Case 3    'Mana
        If UCase$(UserList(UserIndex).Clase) = "GUERRERO" Or _
           UCase$(UserList(UserIndex).Clase) = "ARQUERO" Or _
           UCase$(UserList(UserIndex).Clase) = "CAZADOR" Then
            GranPoder.Vida = UserList(UserIndex).Stats.MaxHP - GranPoder.Cantidad
            UserList(UserIndex).Stats.MinHP = GranPoder.Vida
            UserList(UserIndex).Stats.MaxHP = GranPoder.Vida
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MXVID" & GranPoder.Vida)
            Call EnviarHP(UserIndex)
            GranPoder.TipoAura = 0
            GranPoder.Cantidad = 0
            GranPoder.Status = 0
            GranPoder.UserIndex = 0
            'GranPoder.Timer = 0
            GranPoder.Vida = 0
            GranPoder.Mana = 0
            GranPoder.Agilidad = 0
            GranPoder.Fuerza = 0
            UserList(UserPoder).GranPoder = 0

            Call SendData(SendTarget.ToAll, UserIndex, 0, "||" & UserList(UserPoder).Name & " ha muerto." & FONTTYPE_GUILD)

            If UserList(UserIndex).flags.Privilegios = PlayerType.User Then
                Call DarGranPoder(UserIndex)
            Else
                GranPoder.Timer = 4
            End If
            Exit Sub
        End If
        GranPoder.Mana = UserList(UserPoder).Stats.MaxMAN - GranPoder.Cantidad
        UserList(UserPoder).Stats.MinMAN = GranPoder.Mana
        UserList(UserPoder).Stats.MaxMAN = GranPoder.Mana
        Call SendData(SendTarget.ToIndex, UserPoder, 0, "MXMAN" & GranPoder.Mana)
        Call EnviarMn(UserIndex)
    Case 4    'Agilidad
        UserList(UserPoder).Stats.UserAtributos(eAtributos.Agilidad) = val(UserList(UserPoder).Stats.UserAtributos(eAtributos.Agilidad) - GranPoder.Cantidad)
        UserList(UserPoder).Stats.UserAtributosBackUP(eAtributos.Fuerza) = UserList(UserPoder).Stats.UserAtributosBackUP(eAtributos.Agilidad) - GranPoder.Cantidad
        Call SendData(SendTarget.ToIndex, UserPoder, 0, "MA" & UserList(UserPoder).Stats.UserAtributos(eAtributos.Agilidad))
        Call EnviarAmarillas(UserPoder)
        Call SaveUser(UserPoder, CharPath & UCase$(UserList(UserPoder).Name) & ".chr")
    Case 5    'Fuerza
        UserList(UserPoder).flags.EspecialFuerza = val(UserList(UserPoder).flags.EspecialFuerza - GranPoder.Cantidad)
        UserList(UserPoder).Stats.UserAtributos(eAtributos.Fuerza) = val(UserList(UserPoder).Stats.UserAtributos(eAtributos.Fuerza) - GranPoder.Cantidad)
        Call SendData(SendTarget.ToIndex, UserPoder, 0, "MF" & UserList(UserPoder).Stats.UserAtributos(eAtributos.Fuerza))
        UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Fuerza) = UserList(UserPoder).Stats.UserAtributosBackUP(eAtributos.Fuerza) - GranPoder.Cantidad
        Call SaveUser(UserPoder, CharPath & UCase$(UserList(UserPoder).Name) & ".chr")
        Call EnviarVerdes(UserPoder)
    End Select

    GranPoder.TipoAura = 0
    GranPoder.Cantidad = 0
    GranPoder.Status = 0
    GranPoder.UserIndex = 0
    'GranPoder.Timer = 0
    GranPoder.Vida = 0
    GranPoder.Mana = 0
    GranPoder.Agilidad = 0
    GranPoder.Fuerza = 0
    UserList(UserPoder).GranPoder = 0

    Call SendData(SendTarget.ToAll, UserIndex, 0, "||" & UserList(UserPoder).Name & " ha muerto." & FONTTYPE_GUILD)

    If UserList(UserIndex).flags.Privilegios = PlayerType.User Then
        Call DarGranPoder(UserIndex)
    Else
        GranPoder.Timer = 4
    End If

End Sub

Public Sub MuerePoder(ByVal UserIndex As Integer)

    If UserList(UserIndex).GranPoder <> 1 Then Exit Sub

    Select Case GranPoder.TipoAura
    Case 2    'Vida
        GranPoder.Vida = UserList(UserIndex).Stats.MaxHP - GranPoder.Cantidad
        UserList(UserIndex).Stats.MinHP = 0
        UserList(UserIndex).Stats.MaxHP = GranPoder.Vida
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MXVID" & GranPoder.Vida)
        Call EnviarHP(UserIndex)
    Case 3    'Mana
        If UCase$(UserList(UserIndex).Clase) = "GUERRERO" Or _
           UCase$(UserList(UserIndex).Clase) = "ARQUERO" Or _
           UCase$(UserList(UserIndex).Clase) = "CAZADOR" Then
            GranPoder.Vida = UserList(UserIndex).Stats.MaxHP - GranPoder.Cantidad
            UserList(UserIndex).Stats.MinHP = 0
            UserList(UserIndex).Stats.MaxHP = GranPoder.Vida
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MXVID" & GranPoder.Vida)
            Call EnviarHP(UserIndex)
            GranPoder.TipoAura = 0
            GranPoder.Cantidad = 0
            GranPoder.Status = 0
            GranPoder.UserIndex = 0
            GranPoder.Timer = 0
            GranPoder.Vida = 0
            GranPoder.Mana = 0
            GranPoder.Agilidad = 0
            GranPoder.Fuerza = 0
            UserList(UserIndex).GranPoder = 0

            Call SendData(SendTarget.ToAll, UserIndex, 0, "||" & UserList(UserIndex).Name & " ha muerto." & FONTTYPE_GUILD)

            Exit Sub
        End If
        GranPoder.Mana = UserList(UserIndex).Stats.MaxMAN - GranPoder.Cantidad
        UserList(UserIndex).Stats.MinMAN = 0
        UserList(UserIndex).Stats.MaxMAN = GranPoder.Mana
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MXMAN" & GranPoder.Mana)
        Call EnviarMn(UserIndex)
    Case 4    'Agilidad
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = val(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) - GranPoder.Cantidad)
        UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Agilidad) - GranPoder.Cantidad
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MA" & UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call EnviarAmarillas(UserIndex)
        Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
    Case 5    'Fuerza
        UserList(UserIndex).flags.EspecialFuerza = val(UserList(UserIndex).flags.EspecialFuerza - GranPoder.Cantidad)
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = val(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) - GranPoder.Cantidad)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MF" & UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Fuerza) - GranPoder.Cantidad
        Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
        Call EnviarVerdes(UserIndex)
    End Select

    GranPoder.TipoAura = 0
    GranPoder.Cantidad = 0
    GranPoder.Status = 0
    GranPoder.UserIndex = 0
    GranPoder.Timer = 0
    GranPoder.Vida = 0
    GranPoder.Mana = 0
    GranPoder.Agilidad = 0
    GranPoder.Fuerza = 0
    UserList(UserIndex).GranPoder = 0

    Call SendData(SendTarget.ToAll, UserIndex, 0, "||" & UserList(UserIndex).Name & " ha muerto." & FONTTYPE_GUILD)

End Sub
