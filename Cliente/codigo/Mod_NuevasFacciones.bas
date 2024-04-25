Attribute VB_Name = "Mod_NuevasFacciones"
Option Explicit

Private Const SegundoRango As Byte = 5
Private Const TercerRango  As Byte = 10

Public Sub EnlistarTemplarios(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        'If Criminal(UserIndex) Then
        '    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Largate de aqui!!!!" & "°" & CStr(Npclist( _
        '            .flags.TargetNpc).char.CharIndex))
        '    Exit Sub
        'End If
        
        If .Faccion.Templario = 1 Then
              Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & _
                    "Ya perteneces a la Orden Templaria!!! Ve a combatir otras amardas!!!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If
        
        If .Faccion.FuerzasCaos = 1 Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Maldito insolente!!! vete de aqui, ya perceneces a otra armada!!!" & _
                    "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        If .Faccion.ArmadaReal = 1 Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Maldito insolente!!! vete de aqui, ya perceneces a otra armada!!!" & _
                    "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If
        
        If .Faccion.Nemesis = 1 Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Maldito insolente!!! vete de aqui, ya perceneces a otra armada!!!" & _
                    "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If
        

        If .Stats.ELV < 25 Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & _
                    "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        'If .Faccion.Reenlistadas > 4 Then
        '    If .Faccion.Reenlistadas = 200 Then
        '        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & _
        '                "Has sido expulsado de otra faccion, no te permitire entrar en la mia!" & "°" & CStr(Npclist( _
        '                .flags.TargetNpc).char.CharIndex))
        '    Else
        '        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Has sido expulsado demasias veces!" & "°" & CStr(Npclist( _
        '                .flags.TargetNpc).char.CharIndex))
        '            End If
        '            Exit Sub
        '        End If
        
        .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
        .Faccion.Templario = 1
        .Faccion.RecompensasTemplaria = 0
        .Faccion.FEnlistado = now
        .Reputacion.NobleRep = 1000000
        .Reputacion.AsesinoRep = 0
        .Reputacion.BandidoRep = 0
        .Reputacion.BurguesRep = 0
        .Reputacion.PlebeRep = 0
        
        Call WarpUserChar(UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y)
        
        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & _
                "Bienvenido a la Orden Templaria!!!, aqui tienes tu ropaje de 1ª Jerarquia. Por cada 2 niveles que subas te dare una recompensa, buena suerte soldado!" _
                & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
        
        If .Faccion.RecibioArmaduraTemplaria = 0 Then
        
            Dim Tipo As Byte '1 tunica ' 2 armadura
            
            If UCase$(.Clase) = "MAGO" Or UCase$(.Clase) = "BRUJO" Or UCase$(.Clase) = "DRUIDA" Then
                Tipo = 1
            Else
                Tipo = 2
            End If
            
            Dim MiObj As Obj
            MiObj.Amount = 1
             
            Select Case UCase$(.Raza)
                
                Case "HOBBIT" ' Poner RAZA en mayuscula
                            
                    Select Case UCase$(.Genero)
                    
                        Case "HOMBRE"
                                    
                            If Tipo = 1 Then
                                MiObj.ObjIndex = 1335
                            Else
                                MiObj.ObjIndex = 1326

                            End If

                        Case "MUJER"
                                    
                            If Tipo = 1 Then
                                MiObj.ObjIndex = 1335
                            Else
                                MiObj.ObjIndex = 1347

                            End If
                                        
                    End Select
                                
                Case "ENANO", "GNOMO" ' Poner RAZA en mayuscula
                            
                    Select Case UCase$(.Genero)
                    
                        Case "HOMBRE"
                                    
                            If Tipo = 1 Then
                                MiObj.ObjIndex = 1338
                            Else
                                MiObj.ObjIndex = 1323

                            End If

                        Case "MUJER"
                                    
                            If Tipo = 1 Then
                                MiObj.ObjIndex = 1341
                            Else
                                MiObj.ObjIndex = 1344

                            End If
                                        
                    End Select
                    
                Case Else ' Todo el resto de razas.
                            
                    Select Case UCase$(.Genero)
                    
                        Case "HOMBRE"
                                    
                            If Tipo = 1 Then
                                MiObj.ObjIndex = 1329
                            Else
                                MiObj.ObjIndex = 1318

                            End If

                        Case "MUJER"
                                    
                            If Tipo = 1 Then
                                MiObj.ObjIndex = 1332
                            Else
                                MiObj.ObjIndex = 1317

                            End If

                    End Select
                    
            End Select

            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.pos, MiObj)

            End If
            
            .Faccion.RecibioArmaduraTemplaria = 1
            .Faccion.ArmaduraFaccionaria = MiObj.ObjIndex

        End If

        If .Faccion.RecibioExpInicialTemplaria = 0 Then
            .Stats.Exp = .Stats.Exp + ExpAlUnirse

            If .Stats.Exp > MAXEXP Then
                .Stats.Exp = MAXEXP

            End If

            Call SendData(SendTarget.toindex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT)
            .Faccion.RecibioExpInicialTemplaria = 1
            Call CheckUserLevel(UserIndex)

        End If

    End With

End Sub

'Public Sub ExpulsarFaccionTemplario(ByVal UserIndex As Integer)
'
'    With UserList(UserIndex).Faccion
'        .Templario = 0
'        .NextRecompensas = 0
'        .RecibioArmaduraTemplaria = 0
'        .RecompensasTemplaria = 0
'
'        Call PerderItemsFaccionarios(UserIndex, .ArmaduraFaccionaria)
'
'        'Call PerderItemsFaccionarios(UserIndex)
'        Call SendData(SendTarget.toindex, UserIndex, 0, "||Has sido expulsado de las tropas TEMPLARIAS!!!." & FONTTYPE_FIGHT)
'
'    End With
'
'End Sub

Public Sub RecompensaTemplario(ByVal UserIndex As Integer)

    Dim Nivel As Byte

    With UserList(UserIndex)

        If .Faccion.RecompensasTemplaria = 10 Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ya te di la ultima recompensa!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        Select Case .Faccion.RecompensasTemplaria
           
           Case 0
           If .Stats.ELV >= 27 Then
               .Faccion.RecompensasTemplaria = .Faccion.RecompensasTemplaria + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 1
           If .Stats.ELV >= 29 Then
                .Faccion.RecompensasTemplaria = .Faccion.RecompensasTemplaria + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 2
            If .Stats.ELV >= 31 Then
                .Faccion.RecompensasTemplaria = .Faccion.RecompensasTemplaria + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 3
            If .Stats.ELV >= 33 Then
                .Faccion.RecompensasTemplaria = .Faccion.RecompensasTemplaria + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 4
            If .Stats.ELV >= 35 Then
               .Faccion.RecompensasTemplaria = .Faccion.RecompensasTemplaria + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 5
            If .Stats.ELV >= 37 Then
               .Faccion.RecompensasTemplaria = .Faccion.RecompensasTemplaria + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 6
            If .Stats.ELV >= 39 Then
                .Faccion.RecompensasTemplaria = .Faccion.RecompensasTemplaria + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 7
            If .Stats.ELV >= 41 Then
                .Faccion.RecompensasTemplaria = .Faccion.RecompensasTemplaria + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 8
            If .Stats.ELV >= 43 Then
                .Faccion.RecompensasTemplaria = .Faccion.RecompensasTemplaria + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 9
            If .Stats.ELV >= 45 Then
                .Faccion.RecompensasTemplaria = .Faccion.RecompensasTemplaria + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
       End Select

     Dim Tipo As Byte '1 tunica ' 2 armadura
            
            If UCase$(.Clase) = "MAGO" Or UCase$(.Clase) = "BRUJO" Or UCase$(.Clase) = "DRUIDA" Then
                Tipo = 1
            Else
                Tipo = 2

            End If
            
             If (.Faccion.RecompensasTemplaria + 1) = SegundoRango Or (.Faccion.RecompensasTemplaria + 1) = TercerRango Then
                Dim MiObj As Obj
                MiObj.Amount = 1
                
                 If (.Faccion.RecompensasTemplaria + 1) = SegundoRango Then
                 
                 If .Stats.ELV < 35 Then
                        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & _
                                "No tienes el suficiente nivel para tu recompensa vuelve cuando seas nivel 35!!!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
                        Exit Sub
                    End If
                 
                    Select Case UCase$(.Raza)
                
                        Case "HOBBIT" ' Poner RAZA en mayuscula
                            
                            Select Case UCase$(.Genero)
                       
                                Case "HOMBRE"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1336
                                    Else
                                        MiObj.ObjIndex = 1327

                                    End If

                                Case "MUJER"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1336
                                    Else
                                        MiObj.ObjIndex = 1348

                                    End If
                                        
                            End Select
                                
                        Case "ENANO", "GNOMO" ' Poner RAZA en mayuscula
                            
                            Select Case UCase$(.Genero)
                    
                                Case "HOMBRE"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1339
                                    Else
                                        MiObj.ObjIndex = 1324

                                    End If

                                Case "MUJER"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1342
                                    Else
                                        MiObj.ObjIndex = 1345

                                    End If
                                        
                            End Select
                    
                        Case Else ' Todo el resto de razas.
                            
                            Select Case UCase$(.Genero)
                    
                                Case "HOMBRE"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1330
                                    Else
                                        MiObj.ObjIndex = 1319

                                    End If

                                Case "MUJER"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1331
                                    Else
                                        MiObj.ObjIndex = 1320

                                    End If

                            End Select
                    
                    End Select
                
                Else ' 3Ra Jerarquia
                   
                    If .Stats.ELV < 45 Then
                        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & _
                                "No tienes el suficiente nivel para tu recompensa vuelve cuando seas nivel 45!!!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
                        Exit Sub
                    End If
                   
                    Select Case UCase$(.Raza)
                
                        Case "HOBBIT" ' Poner RAZA en mayuscula
                            
                            Select Case UCase$(.Genero)
                    
                                Case "HOMBRE"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1337
                                    Else
                                        MiObj.ObjIndex = 1328

                                    End If

                                Case "MUJER"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1337
                                    Else
                                        MiObj.ObjIndex = 1349

                                    End If
                                        
                            End Select
                                
                        Case "ENANO", "GNOMO" ' Poner RAZA en mayuscula
                            
                            Select Case UCase$(.Genero)
                    
                                Case "HOMBRE"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1340
                                    Else
                                        MiObj.ObjIndex = 1571

                                    End If

                                Case "MUJER"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1343
                                    Else
                                        MiObj.ObjIndex = 1346

                                    End If
                                        
                            End Select
                    
                        Case Else ' Todo el resto de razas.
                            
                            Select Case UCase$(.Genero)
                    
                                Case "HOMBRE"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1333
                                    Else
                                        MiObj.ObjIndex = 1322

                                    End If

                                Case "MUJER"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1334
                                    Else
                                        MiObj.ObjIndex = 1321

                                    End If

                            End Select
                    
                    End Select
                        
                End If
                
               Call PerderItemsFaccionarios(UserIndex, .Faccion.ArmaduraFaccionaria)
                
                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(.pos, MiObj)
                End If
                
                .Faccion.ArmaduraFaccionaria = MiObj.ObjIndex

            End If

        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Has subido al rango " & .Faccion.RecompensasTemplaria & " en nuestra tropas!!! " & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
                    
            .Stats.Exp = .Stats.Exp + ExpX100
            If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
            Call SendData(SendTarget.toindex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT)
            Call CheckUserLevel(UserIndex)


    End With

End Sub

Public Function TituloTemplario(ByVal UserIndex As Integer) As String

    Dim tStr As String
    
    If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
   
    Select Case UserList(UserIndex).Faccion.RecompensasTemplaria
       
       Case 0
           tStr = "Soldado del temple"
       
        Case 1
            tStr = "Sargento del temple"

        Case 2
            tStr = "Teniente del temple"

        Case 3
            tStr = "Capitán del temple"

        Case 4
            tStr = "Coronel del temple"

        Case 5
            tStr = "General del temple"

        Case 6
            tStr = "Sirviente del temple"

        Case 7
            tStr = "Escudero del temple"

        Case 8
            tStr = "Comendador del temple"

        Case 9
            tStr = "Guerrero templario"

        Case 10
            tStr = "Maestre supremo"

        Case Else ' Este es igual al ultimo rango
            tStr = "CONTACTAR UN ADMINISTRADOR TITULO INEXISTENTE"

    End Select
    
    Else
       
       Select Case UserList(UserIndex).Faccion.RecompensasTemplaria
       
       Case 0
           tStr = "Soldada del temple"
       
        Case 1
            tStr = "Sargenta del temple"

        Case 2
            tStr = "Teniente del temple"

        Case 3
            tStr = "Capitana del temple"

        Case 4
            tStr = "Coronel del temple"

        Case 5
            tStr = "General del temple"

        Case 6
            tStr = "Sirvienta del temple"

        Case 7
            tStr = "Escudera del temple"

        Case 8
            tStr = "Comendadora del temple"

        Case 9
            tStr = "Guerrera templario"

        Case 10
            tStr = "Maestre supremo"

        Case Else ' Este es igual al ultimo rango
            tStr = "CONTACTAR UN ADMINISTRADOR TITULO INEXISTENTE"

    End Select
       
    End If

    TituloTemplario = tStr

End Function

Public Sub EnlistarNemesis(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        'If Not Criminal(UserIndex) Then
        '    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Largate de aqui!!!!" & "°" & CStr(Npclist( _
        '            .flags.TargetNpc).char.CharIndex))
        '    Exit Sub
        'End If
        
        If .Faccion.Nemesis = 1 Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & _
                    "Ya perteneces a los Caballeros de la Tiniebla!!! Ve a combatir otras amardas!!!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        If .Faccion.FuerzasCaos = 1 Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Maldito insolente!!! vete de aqui, ya perceneces a otra armada!!!" & _
                    "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        If .Faccion.ArmadaReal = 1 Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Maldito insolente!!! vete de aqui, ya perceneces a otra armada!!!" & _
                    "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If
        
        If .Faccion.Templario = 1 Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Maldito insolente!!! vete de aqui, ya perceneces a otra armada!!!" & _
                    "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If
        
        'If .Faccion.RecibioExpInicialReal = 1 Or .Faccion.RecibioExpInicialTemplaria = 1 Or .Faccion.RecibioExpInicialCaos = 1 Then
        '    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "No permitiré que ningún desertor de otra faccion" & "°" & str( _
        '            Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        '    Exit Sub
        'End If

        If .Stats.ELV < 25 Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & _
                    "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If

        'If .Faccion.Reenlistadas > 4 Then
        '    If .Faccion.Reenlistadas = 200 Then
        '        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & _
        '                "Has sido expulsado de otra faccion, no te permitire entrar en la mia!" & "°" & CStr(Npclist( _
        '                .flags.TargetNpc).char.CharIndex))
        '    Else
        '        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Has sido expulsado demasias veces!" & "°" & CStr(Npclist( _
        '                .flags.TargetNpc).char.CharIndex))
        '    End If
        '    Exit Sub
        'End If
        
        .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
        .Faccion.Nemesis = 1
        .Faccion.RecompensasNemesis = 0
        .Faccion.FEnlistado = now
        .Reputacion.BandidoRep = 1000000
        .Reputacion.NobleRep = 0
        .Reputacion.AsesinoRep = 0
        .Reputacion.BurguesRep = 0
        .Reputacion.PlebeRep = 0
        
        Call WarpUserChar(UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y)
        
        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & _
                "Bienvenido a los Caballeros de la Tiniebla!!!, aqui tienes tu ropaje de 1ª Jerarquia. Por cada 2 niveles que subas te dare una recompensa, buena suerte soldado!" _
                & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
                
        If .Faccion.RecibioArmaduraNemesis = 0 Then
   
            Dim Tipo As Byte '1 tunica ' 2 armadura
            
            If UCase$(.Clase) = "MAGO" Or UCase$(.Clase) = "BRUJO" Or UCase$(.Clase) = "DRUIDA" Then
                Tipo = 1
            Else
                Tipo = 2

            End If
            
            Dim MiObj As Obj
            MiObj.Amount = 1
             
            Select Case UCase$(.Raza)
                
                Case "HOBBIT" ' Poner RAZA en mayuscula
                            
                    Select Case UCase$(.Genero)
                    
                        Case "HOMBRE"
                                    
                            If Tipo = 1 Then
                                MiObj.ObjIndex = 1562
                            Else
                                MiObj.ObjIndex = 1572

                            End If

                        Case "MUJER"
                                    
                            If Tipo = 1 Then
                                MiObj.ObjIndex = 1562
                            Else
                                MiObj.ObjIndex = 1555

                            End If
                                        
                    End Select
                                
                Case "ENANO", "GNOMO" ' Poner RAZA en mayuscula
                            
                    Select Case UCase$(.Genero)
                    
                        Case "HOMBRE"
                                    
                            If Tipo = 1 Then
                                MiObj.ObjIndex = 1565
                            Else
                                MiObj.ObjIndex = 1576

                            End If

                        Case "MUJER"
                                    
                            If Tipo = 1 Then
                                MiObj.ObjIndex = 1566
                            Else
                                MiObj.ObjIndex = 1554

                            End If
                                        
                    End Select
                    
                Case Else ' Todo el resto de razas.
                            
                    Select Case UCase$(.Genero)
                    
                        Case "HOMBRE"
                                    
                            If Tipo = 1 Then
                                MiObj.ObjIndex = 1560
                            Else
                                MiObj.ObjIndex = 1578

                            End If

                        Case "MUJER"
                                    
                            If Tipo = 1 Then
                                MiObj.ObjIndex = 1556
                            Else
                                MiObj.ObjIndex = 1547

                            End If

                    End Select
                    
            End Select
   
            .Faccion.RecibioArmaduraNemesis = 1
            .Faccion.ArmaduraFaccionaria = MiObj.ObjIndex

            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.pos, MiObj)

            End If

        End If

        If .Faccion.RecibioExpInicialNemesis = 0 Then
            .Stats.Exp = .Stats.Exp + ExpAlUnirse

            If .Stats.Exp > MAXEXP Then
                .Stats.Exp = MAXEXP

            End If

            Call SendData(SendTarget.toindex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT)
            .Faccion.RecibioExpInicialNemesis = 1
            Call CheckUserLevel(UserIndex)

        End If

    End With

End Sub

'Public Sub ExpulsarFaccionNemesis(ByVal UserIndex As Integer)
'
'    With UserList(UserIndex).Faccion
'        .Nemesis = 0
'        .RecibioArmaduraNemesis = 0
 '       .NextRecompensas = 0
 '       .RecompensasNemesis = 0
'
'        Call PerderItemsFaccionarios(UserIndex, .ArmaduraFaccionaria)
'
'        Call SendData(SendTarget.toindex, UserIndex, 0, "||Has sido expulsado de las tropas NEMESIS!!!." & FONTTYPE_FIGHT)
'
'    End With
'
'End Sub

Public Sub RecompensaNemesis(ByVal UserIndex As Integer)

    Dim Nivel As Byte

    With UserList(UserIndex)

        If .Faccion.RecompensasNemesis = 10 Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ya te di la ultima recompensa!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
            Exit Sub

        End If

        Select Case .Faccion.RecompensasNemesis
           
           Case 0
           If .Stats.ELV >= 27 Then
                .Faccion.RecompensasNemesis = .Faccion.RecompensasNemesis + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 1
           If .Stats.ELV >= 29 Then
                .Faccion.RecompensasNemesis = .Faccion.RecompensasNemesis + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 2
            If .Stats.ELV >= 31 Then
                .Faccion.RecompensasNemesis = .Faccion.RecompensasNemesis + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 3
            If .Stats.ELV >= 33 Then
               .Faccion.RecompensasNemesis = .Faccion.RecompensasNemesis + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 4
            If .Stats.ELV >= 35 Then
                .Faccion.RecompensasNemesis = .Faccion.RecompensasNemesis + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 5
            If .Stats.ELV >= 37 Then
               .Faccion.RecompensasNemesis = .Faccion.RecompensasNemesis + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 6
            If .Stats.ELV >= 39 Then
               .Faccion.RecompensasNemesis = .Faccion.RecompensasNemesis + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 7
            If .Stats.ELV >= 41 Then
                .Faccion.RecompensasNemesis = .Faccion.RecompensasNemesis + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 8
            If .Stats.ELV >= 43 Then
               .Faccion.RecompensasNemesis = .Faccion.RecompensasNemesis + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
           Case 9
            If .Stats.ELV >= 45 Then
               .Faccion.RecompensasNemesis = .Faccion.RecompensasNemesis + 1
              Else
             Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ya has recibido tu recompensa, sube mas niveles para subir de rango!!!" & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
              Exit Sub
           End If
           
       End Select
            
            Dim Tipo As Byte '1 tunica ' 2 armadura
            
            If UCase$(.Clase) = "MAGO" Or UCase$(.Clase) = "BRUJO" Or UCase$(.Clase) = "DRUIDA" Then
                Tipo = 1
            Else
                Tipo = 2

            End If
            
             If (.Faccion.RecompensasNemesis + 1) = SegundoRango Or (.Faccion.RecompensasNemesis + 1) = TercerRango Then
                Dim MiObj As Obj
                MiObj.Amount = 1
                
                 If (.Faccion.RecompensasNemesis + 1) = SegundoRango Then
                 
                 If .Stats.ELV < 35 Then
                        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & _
                                "No tienes el suficiente nivel para tu recompensa vuelve cuando seas nivel 35!!!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
                        Exit Sub
                    End If
                 
                    Select Case UCase$(.Raza)
                
                        Case "HOBBIT" ' Poner RAZA en mayuscula
                            
                            Select Case UCase$(.Genero)
                       
                                Case "HOMBRE"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1557
                                    Else
                                        MiObj.ObjIndex = 1573

                                    End If

                                Case "MUJER"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1557
                                    Else
                                        MiObj.ObjIndex = 1550

                                    End If
                                        
                            End Select
                                
                        Case "ENANO", "GNOMO" ' Poner RAZA en mayuscula
                            
                            Select Case UCase$(.Genero)
                    
                                Case "HOMBRE"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1563
                                    Else
                                        MiObj.ObjIndex = 1577

                                    End If

                                Case "MUJER"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1567
                                    Else
                                        MiObj.ObjIndex = 1549

                                    End If
                                        
                            End Select
                    
                        Case Else ' Todo el resto de razas.
                            
                            Select Case UCase$(.Genero)
                    
                                Case "HOMBRE"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1558
                                    Else
                                        MiObj.ObjIndex = 1575

                                    End If

                                Case "MUJER"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1568
                                    Else
                                        MiObj.ObjIndex = 1548

                                    End If

                            End Select
                    
                    End Select
                
                Else ' 3Ra Jerarquia
                   
                    If .Stats.ELV < 45 Then
                        Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & _
                                "No tienes el suficiente nivel para tu recompensa vuelve cuando seas nivel 45!!!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
                        Exit Sub
                    End If
                   
                    Select Case UCase$(.Raza)
                
                        Case "HOBBIT" ' Poner RAZA en mayuscula
                            
                            Select Case UCase$(.Genero)
                    
                                Case "HOMBRE"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1561
                                    Else
                                        MiObj.ObjIndex = 1574

                                    End If

                                Case "MUJER"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1561
                                    Else
                                        MiObj.ObjIndex = 1553

                                    End If
                                        
                            End Select
                                
                        Case "ENANO", "GNOMO" ' Poner RAZA en mayuscula
                            
                            Select Case UCase$(.Genero)
                    
                                Case "HOMBRE"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1564
                                    Else
                                        MiObj.ObjIndex = 1571

                                    End If

                                Case "MUJER"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1569
                                    Else
                                        MiObj.ObjIndex = 1552

                                    End If
                                        
                            End Select
                    
                        Case Else ' Todo el resto de razas.
                            
                            Select Case UCase$(.Genero)
                    
                                Case "HOMBRE"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1559
                                    Else
                                        MiObj.ObjIndex = 1579

                                    End If

                                Case "MUJER"
                                    
                                    If Tipo = 1 Then
                                        MiObj.ObjIndex = 1570
                                    Else
                                        MiObj.ObjIndex = 1551

                                    End If

                            End Select
                    
                    End Select
                        
                End If
                
                Call PerderItemsFaccionarios(UserIndex, .Faccion.ArmaduraFaccionaria)
                
                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(.pos, MiObj)

                End If
                
                .Faccion.ArmaduraFaccionaria = MiObj.ObjIndex

            End If

    
            Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Has subido al rango " & .Faccion.RecompensasNemesis & " en nuestra tropas!!! " & "°" & CStr(Npclist( _
                    .flags.TargetNpc).char.CharIndex))
                    
            .Stats.Exp = .Stats.Exp + ExpX100
            If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
            Call SendData(SendTarget.toindex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT)
            Call CheckUserLevel(UserIndex)
    End With

End Sub

Public Function TituloNemesis(ByVal UserIndex As Integer) As String

    Dim tStr As String
    
    If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then

    Select Case UserList(UserIndex).Faccion.RecompensasNemesis
    
        Case 0
           tStr = "Soldado de la tiniebla"

        Case 1
            tStr = "Sargento de la tiniebla"

        Case 2
            tStr = "Teniente de la tiniebla"

        Case 3
            tStr = "Capitán de la teniebla"

        Case 4
            tStr = "Coronel de la teniebla"

        Case 5
            tStr = "General de la tiniebla"

        Case 6
            tStr = "Acolito de la tiniebla"

        Case 7
            tStr = "Protector de la oscuridad"

        Case 8
            tStr = "Asesino de la tiniebla"

        Case 9
            tStr = "Carcelero de la tiniebla"

        Case 10
            tStr = "Caudillo de la oscuridad"
           
        Case Else ' Este es igual al ultimo rango
            tStr = "CONTACTAR UN ADMINISTRADOR TITULO INEXISTENTE"

    End Select
    
    Else
         
         Select Case UserList(UserIndex).Faccion.RecompensasNemesis
    
        Case 0
           tStr = "Soldada de la tiniebla"

        Case 1
            tStr = "Sargenta de la tiniebla"

        Case 2
            tStr = "Teniente de la tiniebla"

        Case 3
            tStr = "Capitana de la teniebla"

        Case 4
            tStr = "Coronel de la teniebla"

        Case 5
            tStr = "General de la tiniebla"

        Case 6
            tStr = "Acolita de la tiniebla"

        Case 7
            tStr = "Protectora de la oscuridad"

        Case 8
            tStr = "Asesina de la tiniebla"

        Case 9
            tStr = "Carcelera de la tiniebla"

        Case 10
            tStr = "Caudilla de la oscuridad"
           
        Case Else ' Este es igual al ultimo rango
            tStr = "CONTACTAR UN ADMINISTRADOR TITULO INEXISTENTE"

    End Select
         
    End If

    TituloNemesis = tStr

End Function

Public Sub CambiarBarcoTiniebla(ByVal Tipo As Integer, ByVal UserIndex As Integer)

     Dim Objeto As Obj
    
    Select Case Tipo
       
       Case 1
         If Not TieneObjetos(1983, 1, UserIndex) Then
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
           Call QuitarObjetos(1983, 1, UserIndex)
           Objeto.Amount = 1
           Objeto.ObjIndex = 1580
           Call MeterItemEnInventario(UserIndex, Objeto)
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ahí tienes." & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If
        
        Case 2
         If Not TieneObjetos(475, 1, UserIndex) Then
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
           Call QuitarObjetos(475, 1, UserIndex)
           Objeto.Amount = 1
           Objeto.ObjIndex = 1581
           Call MeterItemEnInventario(UserIndex, Objeto)
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ahí tienes." & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If
        
         Case 3
         If Not TieneObjetos(476, 1, UserIndex) Then
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
           Call QuitarObjetos(476, 1, UserIndex)
           Objeto.Amount = 1
           Objeto.ObjIndex = 1582
           Call MeterItemEnInventario(UserIndex, Objeto)
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ahí tienes." & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If
                                
      Case 4
         If Not TieneObjetos(1580, 1, UserIndex) Then
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
           Call QuitarObjetos(1580, 1, UserIndex)
           Objeto.Amount = 1
           Objeto.ObjIndex = 1983
           Call MeterItemEnInventario(UserIndex, Objeto)
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ahí tienes." & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
         End If
         
         Case 5
         If Not TieneObjetos(1581, 1, UserIndex) Then
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
           Call QuitarObjetos(1581, 1, UserIndex)
           Objeto.Amount = 1
           Objeto.ObjIndex = 475
           Call MeterItemEnInventario(UserIndex, Objeto)
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ahí tienes." & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If
        
        Case 6
         If Not TieneObjetos(1582, 1, UserIndex) Then
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
           Call QuitarObjetos(1582, 1, UserIndex)
           Objeto.Amount = 1
           Objeto.ObjIndex = 476
           Call MeterItemEnInventario(UserIndex, Objeto)
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & "&H808080" & "°" & "Ahí tienes." & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If
    
    End Select
    
End Sub

Public Sub CambiarBarcoTemplario(ByVal Tipo As Integer, ByVal UserIndex As Integer)

     Dim Objeto As Obj
    
    Select Case Tipo
       
       Case 1
         If Not TieneObjetos(1983, 1, UserIndex) Then
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
           Call QuitarObjetos(1983, 1, UserIndex)
           Objeto.Amount = 1
           Objeto.ObjIndex = 1350
           Call MeterItemEnInventario(UserIndex, Objeto)
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ahí tienes." & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If
        
        Case 2
         If Not TieneObjetos(475, 1, UserIndex) Then
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
           Call QuitarObjetos(475, 1, UserIndex)
           Objeto.Amount = 1
           Objeto.ObjIndex = 1351
           Call MeterItemEnInventario(UserIndex, Objeto)
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ahí tienes." & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If
        
         Case 3
         If Not TieneObjetos(476, 1, UserIndex) Then
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
           Call QuitarObjetos(476, 1, UserIndex)
           Objeto.Amount = 1
           Objeto.ObjIndex = 1352
           Call MeterItemEnInventario(UserIndex, Objeto)
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ahí tienes." & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If
                                
      Case 4
         If Not TieneObjetos(1350, 1, UserIndex) Then
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
           Call QuitarObjetos(1350, 1, UserIndex)
           Objeto.Amount = 1
           Objeto.ObjIndex = 1983
           Call MeterItemEnInventario(UserIndex, Objeto)
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ahí tienes." & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
         End If
         
         Case 5
         If Not TieneObjetos(1351, 1, UserIndex) Then
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
           Call QuitarObjetos(1351, 1, UserIndex)
           Objeto.Amount = 1
           Objeto.ObjIndex = 475
           Call MeterItemEnInventario(UserIndex, Objeto)
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ahí tienes." & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If
        
        Case 6
         If Not TieneObjetos(1352, 1, UserIndex) Then
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
           Call QuitarObjetos(1352, 1, UserIndex)
           Objeto.Amount = 1
           Objeto.ObjIndex = 476
           Call MeterItemEnInventario(UserIndex, Objeto)
           Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbWhite & "°" & "Ahí tienes." & "°" & CStr( _
                                Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If
    
    End Select
    
End Sub
