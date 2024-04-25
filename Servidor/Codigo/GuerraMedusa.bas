Attribute VB_Name = "GuerraMedusa"
Option Explicit

Public MedAc As Boolean
Public MedEsp As Boolean

Private CantMed As Integer
Public Med_Participantes() As Integer
Public QuitMed_Luchadores() As String
Public CantQuitMed As Integer

Public Corsarios As Integer
Public Piratas As Integer

Public CantidadMedusas As Integer

Public Const MapaMedusa As Integer = 163
Public Const EsperaPirata As Integer = 52
Public Const EsperaPirataY As Integer = 41
Public Const FortaPirata As Integer = 59
Public Const FortaPirataY As Integer = 39
Public Const EsperaCorsario As Integer = 45
Public Const EsperaCorsarioY As Integer = 41
Public Const FortaCorsario As Integer = 42
Public Const FortaCorsarioY As Integer = 39

Public StartPosCorsario As Integer
Public StartPosPirata As Integer

Public Const NpcCorsarios As Integer = 255
Public Const NpcPiratas As Integer = 256

Public Const RecMedOro As Long = 1000000
Public Const RecMedExp As Long = 5000

Public CountFinishPos As Integer
Public CountTwoFinishPos As Integer
Public Const MedFinishMap As Integer = 34
Public Const MedFinishX As Integer = 10
Public Const MedFinishY As Integer = 70

Sub Med_Comienza(Giles As Integer)

    If MedAc = True Then
        Call SendData(SendTarget.ToIndex, 0, 0, "||Ya hay una Guerra de Medusas activa!!" & FONTTYPE_GUERRA)
        Exit Sub
    End If

    If MedEsp = True Then
        Call SendData(SendTarget.ToIndex, 0, 0, "||Ya ha comenzado la Guerra de Medusas" & FONTTYPE_GUERRA)
        Exit Sub
    End If

    CantMed = Giles

    ReDim Med_Participantes(1 To CantMed) As Integer
    ReDim QuitMed_Luchadores(1 To CantMed) As String

    Call Reyes_Medusas

    MedAc = True

    Dim i As Integer

    For i = LBound(Med_Participantes) To UBound(Med_Participantes)
        Med_Participantes(i) = -1
    Next i

End Sub

Sub Med_Entra(UserIndex As Integer)
    Dim i As Integer

    For i = LBound(Med_Participantes) To UBound(Med_Participantes)
        If Med_Participantes(i) = UserIndex Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya eres participante de Guerra de Medusas!!" & FONTTYPE_INFO)
            Exit Sub
        End If
    Next i

    Dim X As Integer

    If CantQuitMed > 0 Then

        For X = LBound(QuitMed_Luchadores) To UBound(QuitMed_Luchadores)
            If readfield2(2, QuitMed_Luchadores(X), 44) = "Corsario" Then

                Med_Participantes(readfield2(1, QuitMed_Luchadores(X), 44)) = UserIndex
                Corsarios = Corsarios + 1

                UserList(Med_Participantes(readfield2(1, QuitMed_Luchadores(X), 44))).flags.bandas = True
                UserList(Med_Participantes(readfield2(1, QuitMed_Luchadores(X), 44))).flags.Angel = True
                Call Med_Transforma(UserIndex)

                Call WarpMedusas(UserIndex, readfield2(3, QuitMed_Luchadores(X), 44))
                CantQuitMed = CantQuitMed - 1
                QuitMed_Luchadores(X) = ""
                Exit Sub
            End If

            If readfield2(2, QuitMed_Luchadores(X), 44) = "Pirata" Then

                Med_Participantes(readfield2(1, QuitMed_Luchadores(X), 44)) = UserIndex
                Piratas = Piratas + 1

                UserList(Med_Participantes(readfield2(1, QuitMed_Luchadores(X), 44))).flags.bandas = True
                UserList(Med_Participantes(readfield2(1, QuitMed_Luchadores(X), 44))).flags.Demonio = True
                Call Med_Transforma(UserIndex)

                Call WarpMedusas(UserIndex, readfield2(3, QuitMed_Luchadores(X), 44))
                CantQuitMed = CantQuitMed - 1
                QuitMed_Luchadores(X) = ""
                Exit Sub
            End If
        Next X

        Exit Sub
    End If

    For i = LBound(Med_Participantes) To UBound(Med_Participantes)

        If Med_Participantes(i) = -1 Then

            Med_Participantes(i) = UserIndex
            UserList(UserIndex).flags.medusas = True
            CantidadMedusas = CantidadMedusas + 1

            If Piratas < Corsarios Then

                UserList(UserIndex).flags.Piratas = True
                Piratas = Piratas + 1
                Call Med_Transforma(UserIndex)
                Call WarpMedusas(UserIndex, Piratas)
                Exit Sub
            Else

                UserList(UserIndex).flags.Corsarios = True
                Corsarios = Corsarios + 1
                Call Med_Transforma(UserIndex)
                Call WarpMedusas(UserIndex, Corsarios)
                Exit Sub
            End If

        End If

    Next i

End Sub

Sub CommandMedusa(UserIndex As Integer)

    If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Medusa Then

        If MedAc = False Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbYellow & "°" & _
                                                            "¡¡La inscripcion para la guerra de medusas empieza cuando queden 10 minutos." & "!!" & "°" & CStr(Npclist(UserList( _
                                                                                                                                                                       UserIndex).flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        If MedEsp = True Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbYellow & "°" & _
                                                            "¡¡La Guerra de Medusa ya ha comenzado." & "!!" & "°" & CStr(Npclist(UserList( _
                                                                                                                                 UserIndex).flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        If UserList(UserIndex).flags.Montado = True Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbYellow & "°" & _
                                                            "No puedes ir a Guerra de Medusa con montura." & "!!" & "°" & CStr(Npclist(UserList( _
                                                                                                                                       UserIndex).flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbYellow & "°" & _
                                                            "Los muertos no pueden entrar en Guerra de Medusa." & "!!" & "°" & CStr(Npclist(UserList( _
                                                                                                                                            UserIndex).flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If


        If UserList(UserIndex).flags.Navegando = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbYellow & "°" & _
                                                            "¡¡Debes estar navegando para participar a la Guerra de Medusas." & "!!" & "°" & CStr(Npclist(UserList( _
                                                                                                                                                          UserIndex).flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        If UserList(UserIndex).Stats.ELV < lvlMedusa Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbYellow & "°" & _
                                                            "Debes ser nivel " & lvlMedusa & "o más para entrar a Guerra de Medusa." & "!!" & "°" & CStr(Npclist(UserList( _
                                                                                                                                                                 UserIndex).flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        Call Med_Entra(UserIndex)

    End If
End Sub

Sub WarpMedusas(UserIndex As Integer, Cantidad As Integer)

    Dim PosX As Byte
    Dim PosY As Byte

    With UserList(UserIndex)

        If .flags.Piratas = True Then
            Select Case Cantidad

            Case 1
                PosX = EsperaPirata
                PosY = EsperaPirataY
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 2
                PosX = EsperaPirata + 1
                PosY = EsperaPirataY
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 3
                PosX = EsperaPirata + 2
                PosY = EsperaPirataY
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 4
                PosX = EsperaPirata + 3
                PosY = EsperaPirataY
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 5
                PosX = EsperaPirata + 4
                PosY = EsperaPirataY
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 6
                PosX = EsperaPirata + 4
                PosY = EsperaPirataY - 1
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 7
                PosX = EsperaPirata + 4
                PosY = EsperaPirataY - 2
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 8
                PosX = EsperaPirata + 4
                PosY = EsperaPirataY - 3
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 9
                PosX = EsperaPirata + 4
                PosY = EsperaPirataY - 4
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 10
                PosX = EsperaPirata + 3
                PosY = EsperaPirataY - 4
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 11
                PosX = EsperaPirata + 2
                PosY = EsperaPirataY - 4
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 12
                PosX = EsperaPirata + 1
                PosY = EsperaPirataY - 4
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 13
                PosX = EsperaPirata
                PosY = EsperaPirataY - 4
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 14
                PosX = EsperaPirata
                PosY = EsperaPirataY - 3
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 15
                PosX = EsperaPirata
                PosY = EsperaPirataY - 2
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 16
                PosX = EsperaPirata
                PosY = EsperaPirataY - 1
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            End Select
        End If

        If .flags.Corsarios = True Then
            Select Case Cantidad

            Case 1
                PosX = EsperaCorsario
                PosY = EsperaCorsarioY
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 2
                PosX = EsperaCorsario + 1
                PosY = EsperaCorsarioY
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 3
                PosX = EsperaCorsario + 2
                PosY = EsperaCorsarioY
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 4
                PosX = EsperaCorsario + 3
                PosY = EsperaCorsarioY
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 5
                PosX = EsperaCorsario + 4
                PosY = EsperaCorsarioY
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 6
                PosX = EsperaCorsario + 4
                PosY = EsperaCorsarioY - 1
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 7
                PosX = EsperaCorsario + 4
                PosY = EsperaCorsarioY - 2
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 8
                PosX = EsperaCorsario + 4
                PosY = EsperaCorsarioY - 3
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 9
                PosX = EsperaCorsario + 4
                PosY = EsperaCorsarioY - 4
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 10
                PosX = EsperaCorsario + 3
                PosY = EsperaCorsarioY - 4
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 11
                PosX = EsperaCorsario + 2
                PosY = EsperaCorsarioY - 4
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 12
                PosX = EsperaCorsario + 1
                PosY = EsperaCorsarioY - 4
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 13
                PosX = EsperaCorsario
                PosY = EsperaCorsarioY - 4
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 14
                PosX = EsperaCorsario
                PosY = EsperaCorsarioY - 3
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 15
                PosX = EsperaCorsario
                PosY = EsperaCorsarioY - 2
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            Case 16
                PosX = EsperaCorsario
                PosY = EsperaCorsarioY - 1
                Call WarpUserChar(UserIndex, _
                                  MapaMedusa, PosX, PosY, True)

            End Select
        End If

    End With
End Sub

Sub Med_Transforma(ByVal UserIndex As Integer)

    On Error GoTo errordm:

    If UserList(UserIndex).flags.Piratas = True Then

        With UserList(UserIndex)
            .CharMimetizado.Body = .char.Body
            .CharMimetizado.Head = .char.Head
            .CharMimetizado.CascoAnim = .char.CascoAnim

            .CharMimetizado.ShieldAnim = .char.ShieldAnim
            .CharMimetizado.WeaponAnim = .char.WeaponAnim

            .CharMimetizado.Alas = .char.Alas

            .flags.Mimetizado = 1

            .char.Body = 695
            .char.Head = 0
            .char.WeaponAnim = 2
            .char.ShieldAnim = 2
            .char.Alas = 0

            Call ChangeUserChar(SendTarget.ToMap, 0, .Pos.Map, UserIndex, .char.Body, .char.Head, _
                                .char.heading, .char.WeaponAnim, .char.ShieldAnim, .char.CascoAnim, .char.Alas)

        End With

    End If

    If UserList(UserIndex).flags.Corsarios = True Then

        With UserList(UserIndex)
            .CharMimetizado.Body = .char.Body
            .CharMimetizado.Head = .char.Head
            .CharMimetizado.CascoAnim = .char.CascoAnim

            .CharMimetizado.ShieldAnim = .char.ShieldAnim
            .CharMimetizado.WeaponAnim = .char.WeaponAnim

            .CharMimetizado.Alas = .char.Alas

            .flags.Mimetizado = 1

            .char.Body = 694
            .char.Head = 0
            .char.WeaponAnim = 2
            .char.ShieldAnim = 2
            .char.Alas = 0

            Call ChangeUserChar(SendTarget.ToMap, 0, .Pos.Map, UserIndex, .char.Body, .char.Head, _
                                .char.heading, .char.WeaponAnim, .char.ShieldAnim, .char.CascoAnim, .char.Alas)

        End With

    End If

errordm:

End Sub

Sub Med_ReloadTransforma(UserIndex)
    With UserList(UserIndex)
        If UserList(UserIndex).flags.Piratas Then
            .flags.Mimetizado = 1

            .char.Body = 695
            .char.Head = 0
            .char.WeaponAnim = 2
            .char.ShieldAnim = 2
            .char.Alas = 0

            Call ChangeUserChar(SendTarget.ToMap, 0, .Pos.Map, UserIndex, .char.Body, .char.Head, _
                                .char.heading, .char.WeaponAnim, .char.ShieldAnim, .char.CascoAnim, .char.Alas)
        End If

        If UserList(UserIndex).flags.Corsarios Then
            .flags.Mimetizado = 1

            .char.Body = 694
            .char.Head = 0
            .char.WeaponAnim = 2
            .char.ShieldAnim = 2
            .char.Alas = 0
            Call ChangeUserChar(SendTarget.ToMap, 0, .Pos.Map, UserIndex, .char.Body, .char.Head, _
                                .char.heading, .char.WeaponAnim, .char.ShieldAnim, .char.CascoAnim, .char.Alas)
        End If
    End With
End Sub

Sub Med_AguaDestransforma(ByVal UserIndex As Integer)
    UserList(UserIndex).Counters.Mimetismo = 0
    UserList(UserIndex).flags.Mimetizado = 0
End Sub

Sub Med_Destransforma(ByVal UserIndex As Integer)

    On Error GoTo errordm

    If UserList(UserIndex).flags.medusas = True Then

        UserList(UserIndex).char.Body = UserList(UserIndex).CharMimetizado.Body
        UserList(UserIndex).char.Head = UserList(UserIndex).CharMimetizado.Head
        UserList(UserIndex).char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
        UserList(UserIndex).char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
        UserList(UserIndex).char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
        UserList(UserIndex).char.Alas = UserList(UserIndex).CharMimetizado.Alas

        UserList(UserIndex).Counters.Mimetismo = 0
        UserList(UserIndex).flags.Mimetizado = 0

        Call ChangeUserChar(SendTarget.ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserIndex, UserList( _
                                                                                                 UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).char.heading, _
                            UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList( _
                                                                                                      UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

    End If

errordm:

End Sub

Sub Med_Desconecta(UserIndex As Integer)

    Dim i As Integer
    Dim Posicion As String
    Dim InfoQuit As String

    If UserList(UserIndex).flags.medusas = True Then

        For i = LBound(Med_Participantes) To UBound(Med_Participantes)
            If Med_Participantes(i) = UserIndex Then
                Med_Participantes(i) = -1
                InfoQuit = i
                Exit For
            End If
        Next i

        If UserList(UserIndex).flags.Piratas = True Then

            Posicion = PositionMedusa(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

            Piratas = Piratas - 1

            InfoQuit = InfoQuit & "," & Posicion

            CantQuitMed = CantQuitMed + 1

        End If

        If UserList(UserIndex).flags.Corsarios = True Then

            Posicion = PositionMedusa(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

            Corsarios = Corsarios - 1

            InfoQuit = InfoQuit & "," & Posicion

            CantQuitMed = CantQuitMed = 1

        End If

        If UserList(UserIndex).flags.Navegando = 0 Then
            UserList(UserIndex).flags.Navegando = 1
        End If

        Call Med_Destransforma(UserIndex)

        Call WarpUserChar(UserIndex, 34, 15, 75, True)


        UserList(UserIndex).flags.medusas = False
        UserList(UserIndex).flags.Corsarios = False
        UserList(UserIndex).flags.Piratas = False

        For i = LBound(QuitMed_Luchadores) To UBound(QuitMed_Luchadores)

            If QuitMed_Luchadores(i) = "" Then
                QuitMed_Luchadores(i) = InfoQuit
                Exit Sub
            End If

        Next i

    End If

End Sub

Function PositionMedusa(UserIndex As Integer, PosX As Integer, PosY As Integer)
    With UserList(UserIndex)

        If .Pos.Map = MapaMedusa Then

            If .flags.Corsarios = True Then

                If PosX = EsperaCorsario And EsperaCorsarioY Then
                    PositionMedusa = "Corsario,1"
                    Exit Function
                ElseIf PosX = EsperaCorsario + 1 And EsperaCorsarioY Then
                    PositionMedusa = "Corsario,2"
                    Exit Function
                ElseIf PosX = EsperaCorsario + 2 And EsperaCorsarioY Then
                    PositionMedusa = "Corsario,3"
                    Exit Function
                ElseIf PosX = EsperaCorsario + 3 And EsperaCorsarioY Then
                    PositionMedusa = "Corsario,4"
                    Exit Function
                ElseIf PosX = EsperaCorsario + 4 And EsperaCorsarioY Then
                    PositionMedusa = "Corsario,5"
                    Exit Function
                ElseIf PosX = EsperaCorsario + 4 And EsperaCorsarioY - 1 Then
                    PositionMedusa = "Corsario,6"
                    Exit Function
                ElseIf PosX = EsperaCorsario + 4 And EsperaCorsarioY - 2 Then
                    PositionMedusa = "Corsario,7"
                    Exit Function
                ElseIf PosX = EsperaCorsario + 4 And EsperaCorsarioY - 3 Then
                    PositionMedusa = "Corsario,8"
                    Exit Function
                ElseIf PosX = EsperaCorsario + 4 And EsperaCorsarioY - 4 Then
                    PositionMedusa = "Corsario,9"
                    Exit Function
                ElseIf PosX = EsperaCorsario + 3 And EsperaCorsarioY - 4 Then
                    PositionMedusa = "Corsario,10"
                    Exit Function
                ElseIf PosX = EsperaCorsario + 2 And EsperaCorsarioY - 4 Then
                    PositionMedusa = "Corsario,11"
                    Exit Function
                ElseIf PosX = EsperaCorsario + 1 And EsperaCorsarioY - 4 Then
                    PositionMedusa = "Corsario,12"
                    Exit Function
                ElseIf PosX = EsperaCorsario And EsperaCorsarioY - 4 Then
                    PositionMedusa = "Corsario,13"
                    Exit Function
                ElseIf PosX = EsperaCorsario And EsperaCorsarioY - 3 Then
                    PositionMedusa = "Corsario,14"
                    Exit Function
                ElseIf PosX = EsperaCorsario And EsperaCorsarioY - 2 Then
                    PositionMedusa = "Corsario,15"
                    Exit Function
                ElseIf PosX = EsperaCorsario And EsperaCorsarioY - 1 Then
                    PositionMedusa = "Corsario,16"
                    Exit Function
                End If

            End If

            If .flags.Piratas = True Then

                If PosX = EsperaPirata And EsperaPirataY Then
                    PositionMedusa = "Pirata,1"
                    Exit Function
                ElseIf PosX = EsperaPirata + 1 And EsperaPirataY Then
                    PositionMedusa = "Pirata,2"
                    Exit Function
                ElseIf PosX = EsperaPirata + 2 And EsperaPirataY Then
                    PositionMedusa = "Pirata,3"
                    Exit Function
                ElseIf PosX = EsperaPirata + 3 And EsperaPirataY Then
                    PositionMedusa = "Pirata,4"
                    Exit Function
                ElseIf PosX = EsperaPirata + 4 And EsperaPirataY Then
                    PositionMedusa = "Pirata,5"
                    Exit Function
                ElseIf PosX = EsperaPirata + 4 And EsperaPirataY - 1 Then
                    PositionMedusa = "Pirata,6"
                    Exit Function
                ElseIf PosX = EsperaPirata + 4 And EsperaPirataY - 2 Then
                    PositionMedusa = "Pirata,7"
                    Exit Function
                ElseIf PosX = EsperaPirata + 4 And EsperaPirataY - 3 Then
                    PositionMedusa = "Pirata,8"
                    Exit Function
                ElseIf PosX = EsperaPirata + 4 And EsperaPirataY - 4 Then
                    PositionMedusa = "Pirata,9"
                    Exit Function
                ElseIf PosX = EsperaPirata + 3 And EsperaPirataY - 4 Then
                    PositionMedusa = "Pirata,10"
                    Exit Function
                ElseIf PosX = EsperaPirata + 2 And EsperaPirataY - 4 Then
                    PositionMedusa = "Pirata,11"
                    Exit Function
                ElseIf PosX = EsperaPirata + 1 And EsperaPirataY - 4 Then
                    PositionMedusa = "Pirata,12"
                    Exit Function
                ElseIf PosX = EsperaPirata And EsperaPirataY - 4 Then
                    PositionMedusa = "Pirata,13"
                    Exit Function
                ElseIf PosX = EsperaPirata And EsperaPirataY - 3 Then
                    PositionMedusa = "Pirata,14"
                    Exit Function
                ElseIf PosX = EsperaPirata And EsperaPirataY - 2 Then
                    PositionMedusa = "Pirata,15"
                    Exit Function
                ElseIf PosX = EsperaPirata And EsperaPirataY - 1 Then
                    PositionMedusa = "Pirata,16"
                    Exit Function
                End If

            End If

        End If

    End With
End Function

Sub Med_Empieza()
    Dim i As Integer

    ReDim Preserve Med_Participantes(1 To CantidadMedusas) As Integer

    Call SendData(SendTarget.ToMap, 0, MapaMedusa, _
                  "||Máten a la medusa del otro bando. GO, GO, GO!!" & FONTTYPE_GUERRA)

    Call SendData(SendTarget.ToMap, 0, MapaMedusa, _
                  "||CORSARIOS: " & val(Corsarios) & " PIRATAS: " & val(Piratas) & FONTTYPE_GUERRA)

    MedEsp = True
    TimerGuerra = 0

    For i = LBound(Med_Participantes) To UBound(Med_Participantes)

        If Med_Participantes(i) <> -1 Then

            If UserList(Med_Participantes(i)).flags.Corsarios = True Then
                StartPosCorsario = StartPosCorsario + 1
                Dim NuevaPos As WorldPos
                Dim FuturePos As WorldPos
                FuturePos.Map = MapaMedusa
                FuturePos.X = FortaCorsario: FuturePos.Y = FortaCorsarioY
                Call MedLegalPos(Med_Participantes(i), StartPosCorsario, FuturePos, NuevaPos)

                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Med_Participantes(i), _
                                                                              NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
            End If

            If UserList(Med_Participantes(i)).flags.Piratas = True Then
                StartPosPirata = StartPosPirata + 1
                Dim NuevaPoss As WorldPos
                Dim FuturePoss As WorldPos
                FuturePoss.Map = MapaMedusa
                FuturePoss.X = FortaPirata: FuturePoss.Y = FortaPirataY
                Call MedLegalPos(Med_Participantes(i), StartPosPirata, FuturePoss, NuevaPoss)

                If NuevaPoss.X <> 0 And NuevaPoss.Y <> 0 Then Call WarpUserChar(Med_Participantes(i), _
                                                                                NuevaPos.Map, NuevaPoss.X, NuevaPoss.Y, True)
            End If

        End If

    Next i
End Sub

Sub MedLegalPos(UserIndex As Integer, StartPos As Integer, Pos As WorldPos, ByRef nPos As WorldPos)
    Dim Tx As Byte
    Dim Ty As Byte

    nPos.Map = Pos.Map

    If UserList(UserIndex).flags.Corsarios = True Then

        Ty = Pos.Y
        Tx = Pos.X - StartPos

        If Not LegalPos(nPos.Map, Tx, Ty) Then
            nPos.X = Tx
            nPos.Y = Ty
        End If
    End If

    If UserList(UserIndex).flags.Piratas = True Then

        Ty = Pos.Y
        Tx = Pos.X + StartPos

        If Not LegalPos(nPos.Map, Tx, Ty) Then
            nPos.X = Tx
            nPos.Y = Ty
        End If
    End If

End Sub

Sub MedFinishPos(UserIndex As Integer, Pos As WorldPos, ByRef nPos As WorldPos)

    Dim Tx As Byte
    Dim Ty As Byte

    nPos.Map = Pos.Map

    If CountFinishPos <= 16 Then

        CountFinishPos = CountFinishPos + 1

        Ty = Pos.Y
        Tx = Pos.X + CountFinishPos

    Else

        CountTwoFinishPos = CountTwoFinishPos + 1

        Ty = Pos.Y + 1
        Tx = Pos.X + CountTwoFinishPos
    End If

    If Not LegalPos(nPos.Map, Tx, Ty) Then
        nPos.X = Tx
        nPos.Y = Ty
    End If

End Sub

Sub Reyes_Medusas()

    Dim Npc3 As Integer
    Dim Npc3Pos As WorldPos
    Npc3 = NpcCorsarios
    Npc3Pos.Map = 163
    Npc3Pos.X = 20
    Npc3Pos.Y = 66

    Dim Npc4 As Integer
    Dim Npc4Pos As WorldPos
    Npc4 = NpcPiratas
    Npc4Pos.Map = 163
    Npc4Pos.X = 81
    Npc4Pos.Y = 66
    Call SpawnNpc(val(Npc3), Npc3Pos, True, False)
    Call SpawnNpc(val(Npc4), Npc4Pos, True, False)

End Sub

Sub RespGuerrasPiratas()

    On Error Resume Next

    Dim i As Integer
    Dim MiNPC As npc
    Dim Npc3 As Integer
    Dim Npc3Pos As WorldPos
    Npc3 = NpcCorsarios
    Npc3Pos.Map = 163
    Npc3Pos.X = 20
    Npc3Pos.Y = 66

    For i = 1 To LastNPC

        'OJO
        If Npclist(i).flags.NPCActive Then

            If Npclist(i).Numero = Npc3 Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)

            End If

            'tildada por sugerencia de yind
            'If Npclist(i).Contadores.TiempoExistencia > 0 Then
            '        Call MuereNpc(i, 0)
            'End If
        End If

    Next i

End Sub

Sub RespGuerrasCorsarios()

    On Error Resume Next

    Dim i As Integer
    Dim MiNPC As npc
    Dim Npc3 As Integer
    Dim Npc3Pos As WorldPos
    Npc3 = NpcPiratas
    Npc3Pos.Map = 163
    Npc3Pos.X = 81
    Npc3Pos.Y = 66

    For i = 1 To LastNPC

        'OJO
        If Npclist(i).flags.NPCActive Then

            If Npclist(i).Numero = Npc3 Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)

            End If

            'tildada por sugerencia de yind
            'If Npclist(i).Contadores.TiempoExistencia > 0 Then
            '        Call MuereNpc(i, 0)
            'End If
        End If

    Next i

End Sub


Sub Med_Corsarios()
    Dim i As Integer

    For i = LBound(Med_Participantes) To UBound(Med_Participantes)

        If Med_Participantes(i) <> -1 Then

            If UserList(Med_Participantes(i)).flags.Corsarios = True Then

                UserList(Med_Participantes(i)).Stats.Exp = UserList(Med_Participantes(i)).Stats.Exp + RecMedExp
                Call CheckUserLevel(Med_Participantes(i))
                Call EnviarExp(Med_Participantes(i))

                Call SendData(SendTarget.ToIndex, Med_Participantes(i), 0, "||Has recibido " & RecMedExp & " de Experencia." & FONTTYPE_FIGHT)

                UserList(Med_Participantes(i)).Stats.GLD = UserList(Med_Participantes(i)).Stats.GLD + RecMedOro
                Call SendUserStatsBox(Med_Participantes(i))

                Call SendData(SendTarget.ToIndex, Med_Participantes(i), 0, "||Has recibido " & RecMedOro & " de Oro." & FONTTYPE_FIGHT)

            End If

            If UserList(Med_Participantes(i)).flags.medusas = True Then

                Dim NuevaPos As WorldPos
                Dim FuturePos As WorldPos

                FuturePos.Map = MedFinishMap
                FuturePos.X = MedFinishX: FuturePos.Y = MedFinishY
                Call MedFinishPos(Med_Participantes(i), FuturePos, NuevaPos)

                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Med_Participantes(i), _
                                                                              NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)

            End If

        End If

        Call Med_Destransforma(Med_Participantes(i))

        UserList(Med_Participantes(i)).flags.medusas = False
        UserList(Med_Participantes(i)).flags.Corsarios = False
        UserList(Med_Participantes(i)).flags.Piratas = False

        MedAc = False
        MedEsp = False
        CantQuitMed = 0
        CantidadMedusas = 0
        StartPosCorsario = 0
        StartPosPirata = 0
        Corsarios = 0
        Piratas = 0
        CountFinishPos = 0
        CountTwoFinishPos = 0
        TimerGuerra = 0
        StatusGuerra = "Banda"

    Next i


End Sub


Sub Med_Piratas()
    Dim i As Integer

    For i = LBound(Med_Participantes) To UBound(Med_Participantes)

        If Med_Participantes(i) <> -1 Then

            If UserList(Med_Participantes(i)).flags.Piratas = True Then

                UserList(Med_Participantes(i)).Stats.Exp = UserList(Med_Participantes(i)).Stats.Exp + RecMedExp
                Call CheckUserLevel(Med_Participantes(i))
                Call EnviarExp(Med_Participantes(i))

                Call SendData(SendTarget.ToIndex, Med_Participantes(i), 0, "||Has recibido " & RecMedExp & " de Experencia." & FONTTYPE_FIGHT)

                UserList(Med_Participantes(i)).Stats.GLD = UserList(Med_Participantes(i)).Stats.GLD + RecMedOro
                Call SendUserStatsBox(Med_Participantes(i))

                Call SendData(SendTarget.ToIndex, Med_Participantes(i), 0, "||Has recibido " & RecMedOro & " de Oro." & FONTTYPE_FIGHT)

            End If

            If UserList(Med_Participantes(i)).flags.medusas = True Then

                Dim NuevaPos As WorldPos
                Dim FuturePos As WorldPos

                FuturePos.Map = MedFinishMap
                FuturePos.X = MedFinishX: FuturePos.Y = MedFinishY
                Call MedFinishPos(Med_Participantes(i), FuturePos, NuevaPos)

                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Med_Participantes(i), _
                                                                              NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)

            End If

        End If

        Call Med_Destransforma(Med_Participantes(i))

        UserList(Med_Participantes(i)).flags.medusas = False
        UserList(Med_Participantes(i)).flags.Corsarios = False
        UserList(Med_Participantes(i)).flags.Piratas = False

        MedAc = False
        MedEsp = False
        CantQuitMed = 0
        CantidadMedusas = 0
        StartPosCorsario = 0
        StartPosPirata = 0
        Corsarios = 0
        Piratas = 0
        CountFinishPos = 0
        CountTwoFinishPos = 0
        TimerGuerra = 0
        StatusGuerra = "Banda"

    Next i

End Sub

Sub Med_Cancela()
    Dim i As Integer

    For i = LBound(Med_Participantes) To UBound(Med_Participantes)

        If Med_Participantes(i) <> -1 Then

            If UserList(Med_Participantes(i)).flags.medusas = True Then

                Dim NuevaPos As WorldPos
                Dim FuturePos As WorldPos

                FuturePos.Map = MedFinishMap
                FuturePos.X = MedFinishX: FuturePos.Y = MedFinishY
                Call MedFinishPos(Med_Participantes(i), FuturePos, NuevaPos)

                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Med_Participantes(i), _
                                                                              NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)

            End If

        End If

        Call Med_Destransforma(Med_Participantes(i))

        UserList(Med_Participantes(i)).flags.medusas = False
        UserList(Med_Participantes(i)).flags.Corsarios = False
        UserList(Med_Participantes(i)).flags.Piratas = False


        MedAc = "False"
        MedEsp = "False"
        CantQuitMed = 0
        CantidadMedusas = 0
        StartPosCorsario = 0
        StartPosPirata = 0
        Corsarios = 0
        Piratas = 0
        CountFinishPos = 0
        CountTwoFinishPos = 0
        TimerGuerra = 0
        StatusGuerra = "Banda"
        Call RespGuerrasPiratas
        Call RespGuerrasCorsarios

    Next i


End Sub

Sub PasaTimeMed()

    Dim HpCorsarios As Integer
    Dim HpPiratas As Integer
    Dim i As Integer

    For i = 1 To NumNPCs

        If Npclist(i).Pos.Map = MapaMedusa Then

            If Npclist(i).Numero = NpcCorsarios Then

                HpCorsarios = Npclist(i).Stats.MinHP

            End If

            If Npclist(i).Numero = NpcPiratas Then

                HpPiratas = Npclist(i).Stats.MinHP

            End If

        End If

    Next i

    If HpCorsarios > HpPiratas Then
        Call SendData(ToAll, 0, 0, "||Corsarios ganaron la batalla de medusas, reciben experiencia como premio!!!" _
                                 & FONTTYPE_GUERRA)
        Call RespGuerrasPiratas
        Call RespGuerrasCorsarios
        Call Med_Corsarios
        Exit Sub
    End If

    If HpPiratas > HpCorsarios Then
        Call SendData(ToAll, 0, 0, "||Piratas ganaron la batalla de medusas, reciben experiencia como premio!!!" _
                                 & FONTTYPE_GUERRA)
        Call RespGuerrasPiratas
        Call RespGuerrasCorsarios
        Call Med_Piratas
        Exit Sub
    End If

    Call SendData(ToAll, 0, 0, "||Corsarios y Piratas empataron la batalla de medusas." _
                             & FONTTYPE_GUERRA)

    Call Med_Cancela
End Sub
