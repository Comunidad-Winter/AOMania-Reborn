Attribute VB_Name = "ModFacciones"
Option Explicit

Private Const SegundoRango As Byte = 5
Private Const TercerRango As Byte = 10

Public Const ExpAlUnirse As Long = 50000
Public Const ExpX100 As Integer = 5000

Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        If .Faccion.ArmadaReal = 1 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & _
                                                            "Ya perteneces a la Armada del Credo!!! Ve a combatir otras amardas!!!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        If .Faccion.FuerzasCaos = 1 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Maldito insolente!!! vete de aqui, ya perceneces a otra armada!!!" & _
                                                            "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        If .Faccion.Templario = 1 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Maldito insolente!!! vete de aqui, ya perceneces a otra armada!!!" & _
                                                            "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub

        End If

        If .Faccion.Nemesis = 1 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Maldito insolente!!! vete de aqui, ya perceneces a otra armada!!!" & _
                                                            "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub

        End If


        If .Stats.ELV < 25 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & _
                                                            "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If


        .Faccion.ArmadaReal = "1"
        .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
        .Faccion.RecompensasReal = 0
        .Faccion.FEnlistado = Now
'        .Reputacion.NobleRep = 1000000
'        .Reputacion.AsesinoRep = 0
'        .Reputacion.BandidoRep = 0
'        .Reputacion.BurguesRep = 0
'        .Reputacion.PlebeRep = 0

        Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

        Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & _
                                                        "Bienvenido a las Armada del Credo!!!, aqui tienes tu ropaje de 1ª Jerarquia. Por cada 2 niveles que subas te dare una recompensa, buena suerte soldado!" _
                                                      & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))

        If .Faccion.RecibioArmaduraReal = 0 Then
            Dim Tipo As Byte    '1 tunica ' 2 armadura

            If UCase$(.Clase) = "MAGO" Or UCase$(.Clase) = "BRUJO" Or UCase$(.Clase) = "DRUIDA" Then
                Tipo = 1
            Else
                Tipo = 2

            End If

            Dim MiObj As Obj
            MiObj.Amount = 1

            Select Case UCase$(.Raza)

            Case "HOBBIT"    ' Poner RAZA en mayuscula

                Select Case UCase$(.Genero)

                Case "HOMBRE"

                    If Tipo = 1 Then
                        MiObj.ObjIndex = 1088
                    Else
                        MiObj.ObjIndex = 1591

                    End If

                Case "MUJER"

                    If Tipo = 1 Then
                        MiObj.ObjIndex = 1088
                    Else
                        MiObj.ObjIndex = 1596

                    End If

                End Select

            Case "ENANO", "GNOMO"    ' Poner RAZA en mayuscula

                Select Case UCase$(.Genero)

                Case "HOMBRE"

                    If Tipo = 1 Then
                        MiObj.ObjIndex = 785
                    Else
                        MiObj.ObjIndex = 492

                    End If

                Case "MUJER"

                    If Tipo = 1 Then
                        MiObj.ObjIndex = 1599
                    Else
                        MiObj.ObjIndex = 1595

                    End If

                End Select

            Case Else    ' Todo el resto de razas.

                Select Case UCase$(.Genero)

                Case "HOMBRE"

                    If Tipo = 1 Then
                        MiObj.ObjIndex = 517
                    Else
                        MiObj.ObjIndex = 370

                    End If

                Case "MUJER"

                    If Tipo = 1 Then
                        MiObj.ObjIndex = 516
                    Else
                        MiObj.ObjIndex = 557

                    End If

                End Select

            End Select

            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)
            End If

            .Faccion.RecibioArmaduraReal = 1
            .Faccion.ArmaduraFaccionaria = MiObj.ObjIndex

        End If

        If .Faccion.RecibioExpInicialReal = 0 Then
            .Stats.Exp = .Stats.Exp + ExpAlUnirse

            If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT)
            .Faccion.RecibioExpInicialReal = 1
            Call CheckUserLevel(UserIndex)

        End If

        Call LogEjercitoReal(.Name)

    End With

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)

    Dim Matados As Integer

    With UserList(UserIndex)

        If .Faccion.RecompensasReal = 10 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ya tienes el mayor rango posible, no puedes subir más!!!" & "°" & CStr(Npclist( _
                                                                                                                                                          .flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If


        Select Case .Faccion.RecompensasReal

        Case 0
            If .Stats.ELV >= 27 Then
                .Faccion.RecompensasReal = .Faccion.RecompensasReal + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                            .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 1
            If .Stats.ELV >= 29 Then
                .Faccion.RecompensasReal = .Faccion.RecompensasReal + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                            .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 2
            If .Stats.ELV >= 31 Then
                .Faccion.RecompensasReal = .Faccion.RecompensasReal + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                            .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 3
            If .Stats.ELV >= 33 Then
                .Faccion.RecompensasReal = .Faccion.RecompensasReal + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                            .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 4
            If .Stats.ELV >= 35 Then
                .Faccion.RecompensasReal = .Faccion.RecompensasReal + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                            .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 5
            If .Stats.ELV >= 37 Then
                .Faccion.RecompensasReal = .Faccion.RecompensasReal + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                            .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 6
            If .Stats.ELV >= 39 Then
                .Faccion.RecompensasReal = .Faccion.RecompensasReal + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                            .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 7
            If .Stats.ELV >= 41 Then
                .Faccion.RecompensasReal = .Faccion.RecompensasReal + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                            .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 8
            If .Stats.ELV >= 43 Then
                .Faccion.RecompensasReal = .Faccion.RecompensasReal + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                            .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 9
            If .Stats.ELV >= 45 Then
                .Faccion.RecompensasReal = .Faccion.RecompensasReal + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                            .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        End Select


        Dim Tipo As Byte    '1 tunica ' 2 armadura

        If UCase$(.Clase) = "MAGO" Or UCase$(.Clase) = "BRUJO" Or UCase$(.Clase) = "DRUIDA" Then
            Tipo = 1
        Else
            Tipo = 2

        End If


        If (.Faccion.RecompensasReal + 1) = SegundoRango Or (.Faccion.RecompensasReal + 1) = TercerRango Then
            Dim MiObj As Obj
            MiObj.Amount = 1

            If (.Faccion.RecompensasReal + 1) = SegundoRango Then

                If .Stats.ELV < 35 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & _
                                                                    "No tienes el suficiente nivel para tu recompensa vuelve cuando seas nivel 35 y tengas " & CStr( _
                                                                    .Faccion.NextRecompensas) & " matados o más!!!." & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
                    Exit Sub
                End If

                Select Case UCase$(.Raza)

                Case "HOBBIT"    ' Poner RAZA en mayuscula

                    Select Case UCase$(.Genero)

                    Case "HOMBRE"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1089
                        Else
                            MiObj.ObjIndex = 1590

                        End If

                    Case "MUJER"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1089
                        Else
                            MiObj.ObjIndex = 1613

                        End If

                    End Select

                Case "ENANO", "GNOMO"    ' Poner RAZA en mayuscula

                    Select Case UCase$(.Genero)

                    Case "HOMBRE"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 799
                        Else
                            MiObj.ObjIndex = 580

                        End If

                    Case "MUJER"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1598
                        Else
                            MiObj.ObjIndex = 1614

                        End If

                    End Select

                Case Else    ' Todo el resto de razas.

                    Select Case UCase$(.Genero)

                    Case "HOMBRE"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 491
                        Else
                            MiObj.ObjIndex = 783

                        End If

                    Case "MUJER"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1601
                        Else
                            MiObj.ObjIndex = 1615

                        End If

                    End Select

                End Select

            Else

                If .Stats.ELV < 45 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & _
                                                                    "No tienes el suficiente nivel para tu recompensa vuelve cuando seas nivel 45 y tengas " & CStr( _
                                                                    .Faccion.NextRecompensas) & " matados o más!!!." & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
                    Exit Sub
                End If

                Select Case UCase$(.Raza)

                Case "HOBBIT"    ' Poner RAZA en mayuscula

                    Select Case UCase$(.Genero)

                    Case "HOMBRE"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1090
                        Else
                            MiObj.ObjIndex = 1589

                        End If

                    Case "MUJER"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1090
                        Else
                            MiObj.ObjIndex = 1593

                        End If

                    End Select

                Case "ENANO", "GNOMO"    ' Poner RAZA en mayuscula

                    Select Case UCase$(.Genero)

                    Case "HOMBRE"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 798
                        Else
                            MiObj.ObjIndex = 522

                        End If

                    Case "MUJER"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1597
                        Else
                            MiObj.ObjIndex = 1594

                        End If

                    End Select

                Case Else    ' Todo el resto de razas.

                    Select Case UCase$(.Genero)

                    Case "HOMBRE"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 581
                        Else
                            MiObj.ObjIndex = 789

                        End If

                    Case "MUJER"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1600
                        Else
                            MiObj.ObjIndex = 1270

                        End If

                    End Select

                End Select

            End If

            Call PerderItemsFaccionarios(UserIndex, .Faccion.ArmaduraFaccionaria)

            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If

            .Faccion.ArmaduraFaccionaria = MiObj.ObjIndex

        End If

        Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Has subido al rango " & .Faccion.RecompensasReal & " en nuestra tropas!!! " & "°" & CStr(Npclist( _
                                                                                                                                                                        .flags.TargetNpc).char.CharIndex))

        .Stats.Exp = .Stats.Exp + ExpX100
        If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT)
        Call CheckUserLevel(UserIndex)


    End With

End Sub

'Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)

'    With UserList(UserIndex).Faccion
'        .ArmadaReal = 0
'        .RecibioArmaduraReal = 0
'        .NextRecompensas = 0
'        .RecompensasReal = 0
'        Call PerderItemsFaccionarios(UserIndex, .ArmaduraFaccionaria)
'        Call SendData(SendTarget.toindex, UserIndex, 0, "||Has sido expulsado de las tropas reales!!!." & FONTTYPE_FIGHT)
'    End With
'End Sub

'Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer)

'    With UserList(UserIndex).Faccion
'        .FuerzasCaos = 0
'        .RecibioArmaduraCaos = 0
'        .NextRecompensas = 0
'        .RecompensasCaos = 0
'
'        Call PerderItemsFaccionarios(UserIndex, .ArmaduraFaccionaria)
'        Call SendData(SendTarget.toindex, UserIndex, 0, "||Has sido expulsado de la legión oscura!!!!." & FONTTYPE_FIGHT)
'    End With
'End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String

    If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then

        Select Case UserList(UserIndex).Faccion.RecompensasReal

        Case 0
            TituloReal = "Soldado del Clero"

        Case 1
            TituloReal = "Sargento del Clero"

        Case 2
            TituloReal = "Teniente del Clero"

        Case 3
            TituloReal = "Capitan del Clero"

        Case 4
            TituloReal = "Coronel del Clero"

        Case 5
            TituloReal = "General del Clero"

        Case 6
            TituloReal = "Consagrado del Clero"

        Case 7
            TituloReal = "Diácono del Clero"

        Case 8
            TituloReal = "Obispo del Clero"

        Case 9
            TituloReal = "Cardenal del Clero"

        Case 10
            TituloReal = "Papa Imperial"

        Case Else
            TituloReal = "CONTACTAR UN ADMINISTRADOR TITULO INEXISTENTE"

        End Select

    Else

        Select Case UserList(UserIndex).Faccion.RecompensasReal

        Case 0
            TituloReal = "Soldada del Clero"

        Case 1
            TituloReal = "Sargenta del Clero"

        Case 2
            TituloReal = "Teniente del Clero"

        Case 3
            TituloReal = "Capitana del Clero"

        Case 4
            TituloReal = "Coronel del Clero"

        Case 5
            TituloReal = "General del Clero"

        Case 6
            TituloReal = "Consagrada del Clero"

        Case 7
            TituloReal = "Diaconisa del Clero"

        Case 8
            TituloReal = "Obispa del Clero"

        Case 9
            TituloReal = "Cardenala del Clero"

        Case 10
            TituloReal = "Mama Imperial"

        Case Else
            TituloReal = "CONTACTAR UN ADMINISTRADOR TITULO INEXISTENTE"

        End Select

    End If

End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        'If Not Criminal(UserIndex) Then
        '    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Largate de aqui, bufon!!!!" & "°" & CStr(Npclist( _
             '            .flags.TargetNpc).char.CharIndex))
        '    Exit Sub
        'End If

        If .Faccion.FuerzasCaos = 1 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ya perteneces a los Demonios de Abbadon!!! Ve a combatir otras amardas!!!" & _
                                                            "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        If .Faccion.ArmadaReal = 1 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Maldito insolente!!! vete de aqui, ya perceneces a otra armada!!!" & _
                                                            "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        If .Faccion.Templario = 1 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Maldito insolente!!! vete de aqui, ya perceneces a otra armada!!!" & _
                                                            "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        If .Faccion.Nemesis = 1 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Maldito insolente!!! vete de aqui, ya perceneces a otra armada!!!" & _
                                                            "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        '[Barrin 17-12-03] Si era miembro de la Armada Real no se puede enlistar
        'If .Faccion.RecibioExpInicialReal = 1 Or .Faccion.RecibioExpInicialNemesis = 1 Or .Faccion.RecibioExpInicialTemplaria = 1 Then    'Tomamos el valor de ahí: ¿Recibio la experiencia para entrar?
        '    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "No permitiré que ningún desertor de otra faccion aqui" & "°" & _
             '            CStr(Npclist(.flags.TargetNpc).char.CharIndex))
        '    Exit Sub
        'End If
        '[/Barrin]

        'If Not Criminal(UserIndex) Then
        '    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ja ja ja tu no eres bienvenido aqui!!!" & "°" & str(Npclist( _
             '            .flags.TargetNpc).char.CharIndex))
        '    Exit Sub
        'End If

        'If .Faccion.CiudadanosMatados < 10 Then
        '    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & _
             '            "Para unirte a nuestras fuerzas debes matar al menos 10 ciudadanos, solo has matado " & .Faccion.CiudadanosMatados & "°" & str( _
             '            Npclist(.flags.TargetNpc).char.CharIndex))
        '    Exit Sub
        'End If

        If .Stats.ELV < 25 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & _
                                                            "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        'If .Faccion.Reenlistadas > 4 Then
        '    If .Faccion.Reenlistadas = 200 Then
        '        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & _
                 '                "Has sido expulsado de las fuerzas oscuras y durante tu rebeldía has atacado a mi ejército. Vete de aquí!" & "°" & CStr( _
                 '                Npclist(.flags.TargetNpc).char.CharIndex))
        '    Else
        '        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & _
                 '                "Has sido expulsado de las fuerzas oscuras demasiadas veces!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
        ' End If
        '   Exit Sub
        ' End If

        .Faccion.FuerzasCaos = 1
        .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
        .Faccion.RecompensasCaos = 0
        .Faccion.FEnlistado = Now
'        .Reputacion.BandidoRep = 1000000
'        .Reputacion.NobleRep = 0
'        .Reputacion.AsesinoRep = 0
'        .Reputacion.BurguesRep = 0
'        .Reputacion.PlebeRep = 0

        Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

        Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & _
                                                        "Bienvenido a los Demonios de Abbadon!!!, aqui tienes tu ropaje de 1ª Jerarquia. Por cada 2 niveles que subas te dare una recompensa, buena suerte soldado!" _
                                                      & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))

        If .Faccion.RecibioArmaduraCaos = 0 Then

            Dim Tipo As Byte    '1 tunica ' 2 armadura

            If UCase$(.Clase) = "MAGO" Or UCase$(.Clase) = "BRUJO" Or UCase$(.Clase) = "DRUIDA" Then
                Tipo = 1
            Else
                Tipo = 2

            End If

            Dim MiObj As Obj
            MiObj.Amount = 1

            Select Case UCase$(.Raza)

            Case "HOBBIT"    ' Poner RAZA en mayuscula

                Select Case UCase$(.Genero)

                Case "HOMBRE"

                    If Tipo = 1 Then
                        MiObj.ObjIndex = 1085
                    Else
                        MiObj.ObjIndex = 1616

                    End If

                Case "MUJER"

                    If Tipo = 1 Then
                        MiObj.ObjIndex = 1085
                    Else
                        MiObj.ObjIndex = 1617

                    End If

                End Select

            Case "ENANO", "GNOMO"    ' Poner RAZA en mayuscula

                Select Case UCase$(.Genero)

                Case "HOMBRE"

                    If Tipo = 1 Then
                        MiObj.ObjIndex = 575
                    Else
                        MiObj.ObjIndex = 786

                    End If

                Case "MUJER"

                    If Tipo = 1 Then
                        MiObj.ObjIndex = 1605
                    Else
                        MiObj.ObjIndex = 1618

                    End If

                End Select

            Case Else    ' Todo el resto de razas.

                Select Case UCase$(.Genero)

                Case "HOMBRE"

                    If Tipo = 1 Then
                        MiObj.ObjIndex = 566
                    Else
                        MiObj.ObjIndex = 572

                    End If

                Case "MUJER"

                    If Tipo = 1 Then
                        MiObj.ObjIndex = 509
                    Else
                        MiObj.ObjIndex = 498

                    End If

                End Select

            End Select

            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If

            .Faccion.ArmaduraFaccionaria = MiObj.ObjIndex
            .Faccion.RecibioArmaduraCaos = 1

        End If

        If .Faccion.RecibioExpInicialCaos = 0 Then
            .Stats.Exp = .Stats.Exp + ExpAlUnirse

            If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT)
            .Faccion.RecibioExpInicialCaos = 1
            Call CheckUserLevel(UserIndex)

        End If

        Call LogEjercitoCaos(.Name)

    End With

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)

    Dim Matados As Integer

    With UserList(UserIndex)

        If .Faccion.RecompensasCaos = 10 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ya te di la ultima recompensa!!!" & "°" & CStr(Npclist( _
                                                                                                                                 .flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        Select Case .Faccion.RecompensasCaos

        Case 0
            If .Stats.ELV >= 27 Then
                .Faccion.RecompensasCaos = .Faccion.RecompensasCaos + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                           .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 1
            If .Stats.ELV >= 29 Then
                .Faccion.RecompensasCaos = .Faccion.RecompensasCaos + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                           .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 2
            If .Stats.ELV >= 31 Then
                .Faccion.RecompensasCaos = .Faccion.RecompensasCaos + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                           .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 3
            If .Stats.ELV >= 33 Then
                .Faccion.RecompensasCaos = .Faccion.RecompensasCaos + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                           .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 4
            If .Stats.ELV >= 35 Then
                .Faccion.RecompensasCaos = .Faccion.RecompensasCaos + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                           .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 5
            If .Stats.ELV >= 37 Then
                .Faccion.RecompensasCaos = .Faccion.RecompensasCaos + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                           .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 6
            If .Stats.ELV >= 39 Then
                .Faccion.RecompensasCaos = .Faccion.RecompensasCaos + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                           .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 7
            If .Stats.ELV >= 41 Then
                .Faccion.RecompensasCaos = .Faccion.RecompensasCaos + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                           .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 8
            If .Stats.ELV >= 43 Then
                .Faccion.RecompensasCaos = .Faccion.RecompensasCaos + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                           .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        Case 9
            If .Stats.ELV >= 45 Then
                .Faccion.RecompensasCaos = .Faccion.RecompensasCaos + 1
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                                                                                                                                                                           .flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If

        End Select


        Dim Tipo As Byte    '1 tunica ' 2 armadura

        If UCase$(.Clase) = "MAGO" Or UCase$(.Clase) = "BRUJO" Or UCase$(.Clase) = "DRUIDA" Then
            Tipo = 1
        Else
            Tipo = 2

        End If

        If (.Faccion.RecompensasCaos + 1) = SegundoRango Or (.Faccion.RecompensasCaos + 1) = TercerRango Then
            Dim MiObj As Obj
            MiObj.Amount = 1

            If (.Faccion.RecompensasCaos + 1) = SegundoRango Then    ' 2Da Jerarquia

                If .Stats.ELV < 35 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & _
                                                                    "No tienes el suficiente nivel para tu recompensa vuelve cuando seas nivel 35 y tengas " & .Faccion.NextRecompensas _
                                                                  & " matados o más!!!." & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))

                    Exit Sub

                End If

                Select Case UCase$(.Raza)

                Case "HOBBIT"    ' Poner RAZA en mayuscula

                    Select Case UCase$(.Genero)

                    Case "HOMBRE"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1086
                        Else
                            MiObj.ObjIndex = 1610

                        End If

                    Case "MUJER"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1086
                        Else
                            MiObj.ObjIndex = 1611

                        End If

                    End Select

                Case "ENANO", "GNOMO"    ' Poner RAZA en mayuscula

                    Select Case UCase$(.Genero)

                    Case "HOMBRE"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1078
                        Else
                            MiObj.ObjIndex = 574

                        End If

                    Case "MUJER"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1608
                        Else
                            MiObj.ObjIndex = 1612

                        End If

                    End Select

                Case Else    ' Todo el resto de razas.

                    Select Case UCase$(.Genero)

                    Case "HOMBRE"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 369
                        Else
                            MiObj.ObjIndex = 523

                        End If

                    Case "MUJER"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1604
                        Else
                            MiObj.ObjIndex = 494

                        End If

                    End Select

                End Select

            Else

                If .Stats.ELV < 45 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & _
                                                                    "No tienes el suficiente nivel para tu recompensa vuelve cuando seas nivel 45 y tengas " & .Faccion.RecompensasCaos _
                                                                  & " matados o más!!!." & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))

                    Exit Sub

                End If

                Select Case UCase$(.Raza)

                Case "HOBBIT"    ' Poner RAZA en mayuscula

                    Select Case UCase$(.Genero)

                    Case "HOMBRE"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1087
                        Else
                            MiObj.ObjIndex = 1606

                        End If

                    Case "MUJER"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1087
                        Else
                            MiObj.ObjIndex = 1607

                        End If

                    End Select

                Case "ENANO", "GNOMO"    ' Poner RAZA en mayuscula

                    Select Case UCase$(.Genero)

                    Case "HOMBRE"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1077
                        Else
                            MiObj.ObjIndex = 794

                        End If

                    Case "MUJER"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1603
                        Else
                            MiObj.ObjIndex = 794

                        End If

                    End Select

                Case Else    ' Todo el resto de razas.

                    Select Case UCase$(.Genero)

                    Case "HOMBRE"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 792
                        Else
                            MiObj.ObjIndex = 795

                        End If

                    Case "MUJER"

                        If Tipo = 1 Then
                            MiObj.ObjIndex = 1055
                        Else
                            MiObj.ObjIndex = 1609

                        End If

                    End Select

                End Select

            End If

            Call PerderItemsFaccionarios(UserIndex, .Faccion.ArmaduraFaccionaria)

            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If

            .Faccion.ArmaduraFaccionaria = MiObj.ObjIndex

        End If

        Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Has subido al rango " & .Faccion.RecompensasCaos & " en nuestra tropas!!! " & "°" & CStr(Npclist( _
                                                                                                                                                                       .flags.TargetNpc).char.CharIndex))

        .Stats.Exp = .Stats.Exp + ExpX100
        If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT)
        Call CheckUserLevel(UserIndex)

    End With

End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String

    If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then

        Select Case UserList(UserIndex).Faccion.RecompensasCaos

        Case 0
            TituloCaos = "Soldado del abbadon"

        Case 1
            TituloCaos = "Sargento del abbadon"

        Case 2
            TituloCaos = "Teniente del abbadon"

        Case 3
            TituloCaos = "Capitán del abbadon"

        Case 4
            TituloCaos = "Coronel del abbadon"

        Case 5
            TituloCaos = "General de abbadon"

        Case 6
            TituloCaos = "Consejero de abbadon"

        Case 7
            TituloCaos = "Ejecutor de abbadon"

        Case 8
            TituloCaos = "Príncipe de inframundo"

        Case 9
            TituloCaos = "Rey del inframundo"

        Case 10
            TituloCaos = "Dios demonio"

        Case Else
            TituloCaos = "CONTACTAR UN ADMINISTRADOR TITULO INEXISTENTE"

        End Select

    Else

        Select Case UserList(UserIndex).Faccion.RecompensasCaos

        Case 0
            TituloCaos = "Soldada del abbadon"

        Case 1
            TituloCaos = "Sargenta del abbadon"

        Case 2
            TituloCaos = "Teniente del abbadon"

        Case 3
            TituloCaos = "Capitana del abbadon"

        Case 4
            TituloCaos = "Coronel del abbadon"

        Case 5
            TituloCaos = "General de abbadon"

        Case 6
            TituloCaos = "Consejera de abbadon"

        Case 7
            TituloCaos = "Ejecutora de abbadon"

        Case 8
            TituloCaos = "Príncesa de inframundo"

        Case 9
            TituloCaos = "Reina del inframundo"

        Case 10
            TituloCaos = "Diosa demonio"

        Case Else
            TituloCaos = "CONTACTAR UN ADMINISTRADOR TITULO INEXISTENTE"

        End Select

    End If

End Function

Public Sub PerderItemsFaccionarios(ByVal UserIndex As Integer, ByVal ArmIndex As Integer)

    Dim i As Long
    Dim ItemIndex As Integer

    With UserList(UserIndex)

        For i = 1 To MAX_INVENTORY_SLOTS
            ItemIndex = .Invent.Object(i).ObjIndex

            If ItemIndex > 0 And ItemIndex = ArmIndex Then

                Call QuitarUserInvItem(UserIndex, i, .Invent.Object(i).Amount)
                Call UpdateUserInv(False, UserIndex, i)

                Exit For

            End If

        Next i

        .Faccion.ArmaduraFaccionaria = 0

    End With

End Sub

Public Sub CambiarBarcoClero(ByVal Tipo As Integer, ByVal UserIndex As Integer)

    Dim Objeto As Obj

    Select Case Tipo

    Case 1
        If Not TieneObjetos(1983, 1, UserIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        Else
            Call QuitarObjetos(1983, 1, UserIndex)
            Objeto.Amount = 1
            Objeto.ObjIndex = 1117
            Call MeterItemEnInventario(UserIndex, Objeto)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ahí tienes." & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If

    Case 2
        If Not TieneObjetos(475, 1, UserIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        Else
            Call QuitarObjetos(475, 1, UserIndex)
            Objeto.Amount = 1
            Objeto.ObjIndex = 1118
            Call MeterItemEnInventario(UserIndex, Objeto)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ahí tienes." & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If

    Case 3
        If Not TieneObjetos(476, 1, UserIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        Else
            Call QuitarObjetos(476, 1, UserIndex)
            Objeto.Amount = 1
            Objeto.ObjIndex = 1119
            Call MeterItemEnInventario(UserIndex, Objeto)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ahí tienes." & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If

    Case 4
        If Not TieneObjetos(1117, 1, UserIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        Else
            Call QuitarObjetos(1117, 1, UserIndex)
            Objeto.Amount = 1
            Objeto.ObjIndex = 1983
            Call MeterItemEnInventario(UserIndex, Objeto)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ahí tienes." & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If

    Case 5
        If Not TieneObjetos(1118, 1, UserIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        Else
            Call QuitarObjetos(1118, 1, UserIndex)
            Objeto.Amount = 1
            Objeto.ObjIndex = 475
            Call MeterItemEnInventario(UserIndex, Objeto)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ahí tienes." & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If

    Case 6
        If Not TieneObjetos(1119, 1, UserIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        Else
            Call QuitarObjetos(1119, 1, UserIndex)
            Objeto.Amount = 1
            Objeto.ObjIndex = 476
            Call MeterItemEnInventario(UserIndex, Objeto)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ahí tienes." & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If

    End Select

End Sub

Public Sub CambiarBarcoAbbadon(ByVal Tipo As Integer, ByVal UserIndex As Integer)

    Dim Objeto As Obj

    Select Case Tipo

    Case 1
        If Not TieneObjetos(1983, 1, UserIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        Else
            Call QuitarObjetos(1983, 1, UserIndex)
            Objeto.Amount = 1
            Objeto.ObjIndex = 1120
            Call MeterItemEnInventario(UserIndex, Objeto)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ahí tienes." & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If

    Case 2
        If Not TieneObjetos(475, 1, UserIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        Else
            Call QuitarObjetos(475, 1, UserIndex)
            Objeto.Amount = 1
            Objeto.ObjIndex = 1121
            Call MeterItemEnInventario(UserIndex, Objeto)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ahí tienes." & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If

    Case 3
        If Not TieneObjetos(476, 1, UserIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        Else
            Call QuitarObjetos(476, 1, UserIndex)
            Objeto.Amount = 1
            Objeto.ObjIndex = 1122
            Call MeterItemEnInventario(UserIndex, Objeto)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ahí tienes." & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If

    Case 4
        If Not TieneObjetos(1120, 1, UserIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        Else
            Call QuitarObjetos(1120, 1, UserIndex)
            Objeto.Amount = 1
            Objeto.ObjIndex = 1983
            Call MeterItemEnInventario(UserIndex, Objeto)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ahí tienes." & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If

    Case 5
        If Not TieneObjetos(1121, 1, UserIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        Else
            Call QuitarObjetos(1121, 1, UserIndex)
            Objeto.Amount = 1
            Objeto.ObjIndex = 475
            Call MeterItemEnInventario(UserIndex, Objeto)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ahí tienes." & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If

    Case 6
        If Not TieneObjetos(1122, 1, UserIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        Else
            Call QuitarObjetos(1122, 1, UserIndex)
            Objeto.Amount = 1
            Objeto.ObjIndex = 476
            Call MeterItemEnInventario(UserIndex, Objeto)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ahí tienes." & "°" & CStr( _
                                                            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If

    End Select

End Sub

