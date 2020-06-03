Attribute VB_Name = "Mod_Facciones"
' MOD FACCIONES!
' By Bassinger
'
'Sistema pensado por DUNCAN con eventos de ciudades.
'Comienzo de enlistamientos en: Real y Nemesis. (La armada del Credo/Los caballeros de la tinieblas)
'Facciones elites (Opcionales): Templario y Nemesis. (La orden templaria/Los demonios de abaddon)

Option Explicit

Private Const SegundoRango As Byte = 5
Private Const TercerRango  As Byte = 10
Public Const ExpAlUnirse = 10000
Public Const ExpX100 = 10000
Public MAX_ARMADURAS_ARMADA As Integer
Public Armaduras_Armada(1000) As Integer

'#####ARMADURAS&TUNICAS ARMADAS DEL CREDO

Public ArmaduraPaladinClero As Integer
Public ArmaduraClerigoClero As Integer
Public ArmaduraEnanoClero As Integer
Public ArmaduraEnanoCleroMujer As Integer
Public ArmaduraCleroMujer As Integer
Public ArmaduraCleroHobbit As Integer
Public ArmaduraCleroHobbitMujer As Integer

Public TunicaMagoClero As Integer
Public TunicaMagoCleroEnano As Integer
Public TunicaMagoCleroEnanoMujer As Integer
Public TunicaMagoCleroHobbit As Integer
Public TunicaMagoCleroHobbitMujer As Integer
Public TunicaMagoCleroMujer As Integer

Public ArmaduraPaladinClero2 As Integer
Public ArmaduraClerigoClero2 As Integer
Public ArmaduraEnanoClero2 As Integer
Public ArmaduraEnanoCleroMujer2 As Integer
Public ArmaduraCleroMujer2 As Integer
Public ArmaduraCleroHobbit2 As Integer
Public ArmaduraCleroHobbitMujer2 As Integer

Public TunicaMagoClero2 As Integer
Public TunicaMagoCleroEnano2 As Integer
Public TunicaMagoCleroEnanoMujer2 As Integer
Public TunicaMagoCleroHobbit2 As Integer
Public TunicaMagoCleroHobbitMujer2 As Integer
Public TunicaMagoCleroMujer2 As Integer

Public ArmaduraPaladinClero3 As Integer
Public ArmaduraClerigoClero3 As Integer
Public ArmaduraEnanoClero3 As Integer
Public ArmaduraEnanoCleroMujer3 As Integer
Public ArmaduraCleroMujer3 As Integer
Public ArmaduraCleroHobbit3 As Integer
Public ArmaduraCleroHobbitMujer3 As Integer

Public TunicaMagoClero3 As Integer
Public TunicaMagoCleroEnano3 As Integer
Public TunicaMagoCleroEnanoMujer3 As Integer
Public TunicaMagoCleroHobbit3 As Integer
Public TunicaMagoCleroHobbitMujer3 As Integer
Public TunicaMagoCleroMujer3 As Integer

'#####ARMADURAS&TUNICAS ARMADAS DE LA TINIEBLA

Public ArmaduraPaladinTiniebla As Integer
Public ArmaduraEnanoTiniebla As Integer
Public ArmaduraEnanoTinieblaMujer As Integer
Public ArmaduraTinieblaMujer As Integer
Public ArmaduraTinieblaHobbit As Integer
Public ArmaduraTinieblaHobbitMujer As Integer

Public TunicaMagoTiniebla As Integer
Public TunicaMagoTinieblaEnano As Integer
Public TunicaMagoTinieblaEnanoMujer As Integer
Public TunicaMagoTinieblaHobbit As Integer
Public TunicaMagoTinieblaMujer As Integer

Public ArmaduraPaladinTiniebla2 As Integer
Public ArmaduraEnanoTiniebla2 As Integer
Public ArmaduraEnanoTinieblaMujer2 As Integer
Public ArmaduraTinieblaMujer2 As Integer
Public ArmaduraTinieblaHobbit2 As Integer
Public ArmaduraTinieblaHobbitMujer2 As Integer

Public TunicaMagoTiniebla2 As Integer
Public TunicaMagoTinieblaEnano2 As Integer
Public TunicaMagoTinieblaEnanoMujer2 As Integer
Public TunicaMagoTinieblaHobbit2 As Integer
Public TunicaMagoTinieblaMujer2 As Integer
Public TunicaMagoTinieblaMujerHobbit2 As Integer

Public ArmaduraPaladinTiniebla3 As Integer
Public ArmaduraEnanoTiniebla3 As Integer
Public ArmaduraEnanoTinieblaMujer3 As Integer
Public ArmaduraTinieblaMujer3 As Integer
Public ArmaduraTinieblaHobbit3 As Integer
Public ArmaduraTinieblaHobbitMujer3 As Integer

Public TunicaMagoTiniebla3 As Integer
Public TunicaMagoTinieblaEnano3 As Integer
Public TunicaMagoTinieblaEnanoMujer3 As Integer
Public TunicaMagoTinieblaHobbit3 As Integer
Public TunicaMagoTinieblaMujer3 As Integer
Public TunicaMagoTinieblaMujerHobbit3 As Integer

'#####ARMADURAS&TUNICAS ARMADAS DE TEMPLARIO

Public ArmaduraPaladinTemplario As Integer
Public ArmaduraEnanoTemplario As Integer
Public ArmaduraEnanoTemplarioMujer As Integer
Public ArmaduraTemplarioMujer As Integer
Public ArmaduraTemplarioHobbit As Integer
Public ArmaduraTemplarioHobbitMujer As Integer

Public TunicaMagoTemplario As Integer
Public TunicaMagoTemplarioEnano As Integer
Public TunicaMagoTemplarioEnanoMujer As Integer
Public TunicaMagoTemplarioHobbit As Integer
Public TunicaMagoTemplarioMujer As Integer

Public ArmaduraPaladinTemplario2 As Integer
Public ArmaduraEnanoTemplario2 As Integer
Public ArmaduraEnanoTemplarioMujer2 As Integer
Public ArmaduraTemplarioMujer2 As Integer
Public ArmaduraTemplarioHobbit2 As Integer
Public ArmaduraTemplarioHobbitMujer2 As Integer

Public TunicaMagoTemplario2 As Integer
Public TunicaMagoTemplarioEnano2 As Integer
Public TunicaMagoTemplarioEnanoMujer2 As Integer
Public TunicaMagoTemplarioHobbit2 As Integer
Public TunicaMagoTemplarioMujer2 As Integer
Public TunicaMagoTemplarioMujerHobbit2 As Integer

Public ArmaduraPaladinTemplario3 As Integer
Public ArmaduraEnanoTemplario3 As Integer
Public ArmaduraEnanoTemplarioMujer3 As Integer
Public ArmaduraTemplarioMujer3 As Integer
Public ArmaduraTemplarioHobbit3 As Integer
Public ArmaduraTemplarioHobbitMujer3 As Integer

Public TunicaMagoTemplario3 As Integer
Public TunicaMagoTemplarioEnano3 As Integer
Public TunicaMagoTemplarioEnanoMujer3 As Integer
Public TunicaMagoTemplarioHobbit3 As Integer
Public TunicaMagoTemplarioMujer3 As Integer
Public TunicaMagoTemplarioMujerHobbit3 As Integer

'#####ARMADURAS&TUNICAS ARMADAS DEL ABADDON

Public ArmaduraPaladinAbaddon As Integer
Public ArmaduraEnanoAbaddon As Integer
Public ArmaduraEnanoAbaddonMujer As Integer
Public ArmaduraAbaddonMujer As Integer
Public ArmaduraGnomoAbaddon As Integer
Public ArmaduraAbaddonHobbitMujer As Integer
Public ArmaduraPaladinAbaddonHobbit As Integer

Public TunicaMagoAbaddon As Integer
Public TunicaMagoAbaddonEnano As Integer
Public TunicaMagoAbaddonEnanoMujer As Integer
Public TunicaMagoAbaddonHobbit As Integer
Public TunicaMagoAbaddonHobbitMujer As Integer
Public TunicaMagoAbaddonMujer As Integer

Public ArmaduraPaladinAbaddon2 As Integer
Public ArmaduraEnanoAbaddon2 As Integer
Public ArmaduraEnanoAbaddonMujer2 As Integer
Public ArmaduraAbaddonMujer2 As Integer
Public ArmaduraAbaddonHobbit2 As Integer
Public ArmaduraAbaddonHobbitMujer2 As Integer
Public ArmaduraGnomoAbaddon2 As Integer

Public TunicaMagoAbaddon2 As Integer
Public TunicaMagoAbaddonEnano2 As Integer
Public TunicaMagoAbaddonEnanoMujer2 As Integer
Public TunicaMagoAbaddonHobbit2 As Integer
Public TunicaMagoAbaddonHobbitMujer2 As Integer
Public TunicaMagoAbaddonMujer2 As Integer

Public ArmaduraPaladinAbaddon3 As Integer
Public ArmaduraEnanoAbaddon3 As Integer
Public ArmaduraEnanoAbaddonMujer3 As Integer
Public ArmaduraAbaddonMujer3 As Integer
Public ArmaduraAbaddonHobbit3 As Integer
Public ArmaduraAbaddonHobbitMujer3 As Integer
Public ArmaduraGnomoAbaddon3 As Integer

Public TunicaMagoAbaddon3 As Integer
Public TunicaMagoAbaddonEnano3 As Integer
Public TunicaMagoAbaddonEnanoMujer3 As Integer
Public TunicaMagoAbaddonHobbit3 As Integer
Public TunicaMagoAbaddonHobbitMujer3 As Integer
Public TunicaMagoAbaddonMujer3 As Integer

Public Sub EnlistarArmadaClero(ByVal UserIndex As Integer)
     
    With UserList(UserIndex)
             
        If .Faccion.ArmadaReal = 1 Or .Faccion.Templario = 1 Then

            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "¡Ya perteneces a la armada del Clero, ve a combatir contra los enemigos!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub

        End If
             
        If .Faccion.FuerzasCaos = 1 Or .Faccion.Nemesis = 1 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Maldito insolente!!! vete de aqui, ya perceneces a otra armada!!!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If
             
        If .Stats.ELV < 14 Then

            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 14!!!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub

        End If
             
        .Faccion.ArmadaReal = 1
        .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
             
        If .Faccion.RecibioExpInicialReal = 0 Then

            Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpAlUnirse, MAXEXP)
            Call SendData(toIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_Motd4)
            .Faccion.RecibioExpInicialReal = 1
            Call CheckUserLevel(UserIndex)

        End If
        
        Call WarpUserChar(UserIndex, 59, 50, 41, True)
             
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Bienvenido a las Armada del Credo!!!, aqui tienes tu ropaje de 1ª Jerarquia. Por cada 2 niveles que subas te dare una recompensa, buena suerte soldado!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
             
    End With
     
End Sub

Public Sub RecompensaArmadaClero(ByVal UserIndex As Integer)
     
     With UserList(UserIndex)
     
        If .Faccion.Templario = 1 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Debes pedirle la recompensa a la Orden Templaria!" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If
        
        If .Faccion.ArmadaReal = 0 And .Faccion.Templario = 0 Then
             Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "No perteneces a la Armada del Credo, vete de aquí o te ahogaras en tu insolencia!!" & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
             Exit Sub
        End If
        
        If .Stats.ELV < 25 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Para recibir la recompensa debes ser al menos de nivel 25" & "°" & CStr(Npclist(.flags.TargetNpc).char.CharIndex))
            Exit Sub
        End If
        
        
     
     End With
     
End Sub

Public Sub CambiarBarcoTemplario(ByVal Tipo As Integer, ByVal UserIndex As Integer)

    Dim Objeto As Obj
    
    Select Case Tipo
       
        Case 1
            If Not TieneObjetos(1983, 1, UserIndex) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(1983, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 1350
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ahí tienes." & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            End If
        
        Case 2
            If Not TieneObjetos(475, 1, UserIndex) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(475, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 1351
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ahí tienes." & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            End If
        
        Case 3
            If Not TieneObjetos(476, 1, UserIndex) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(476, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 1352
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ahí tienes." & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            End If
                                
        Case 4
            If Not TieneObjetos(1350, 1, UserIndex) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(1350, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 1983
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ahí tienes." & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            End If
         
        Case 5
            If Not TieneObjetos(1351, 1, UserIndex) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(1351, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 475
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ahí tienes." & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            End If
        
        Case 6
            If Not TieneObjetos(1352, 1, UserIndex) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(1352, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 476
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ahí tienes." & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            End If
    
    End Select
    
End Sub

Public Sub CambiarBarcoTiniebla(ByVal Tipo As Integer, ByVal UserIndex As Integer)

    Dim Objeto As Obj
    
    Select Case Tipo
       
        Case 1
            If Not TieneObjetos(1983, 1, UserIndex) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & "&H808080" & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(1983, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 1580
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & "&H808080" & "°" & "Ahí tienes." & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            End If
        
        Case 2
            If Not TieneObjetos(475, 1, UserIndex) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & "&H808080" & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(475, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 1581
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & "&H808080" & "°" & "Ahí tienes." & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            End If
        
        Case 3
            If Not TieneObjetos(476, 1, UserIndex) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & "&H808080" & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(476, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 1582
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & "&H808080" & "°" & "Ahí tienes." & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            End If
                                
        Case 4
            If Not TieneObjetos(1580, 1, UserIndex) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & "&H808080" & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(1580, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 1983
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & "&H808080" & "°" & "Ahí tienes." & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            End If
         
        Case 5
            If Not TieneObjetos(1581, 1, UserIndex) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & "&H808080" & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(1581, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 475
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & "&H808080" & "°" & "Ahí tienes." & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            End If
        
        Case 6
            If Not TieneObjetos(1582, 1, UserIndex) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & "&H808080" & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(1582, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 476
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & "&H808080" & "°" & "Ahí tienes." & "°" & CStr( _
                    Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            End If
    
    End Select
    
End Sub


Public Sub CambiarBarcoClero(ByVal Tipo As Integer, ByVal UserIndex As Integer)

    Dim Objeto As Obj

    Select Case Tipo

        Case 1

            If Not TieneObjetos(1983, 1, UserIndex) Then

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(1983, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 1117
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ahí tienes." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))

            End If

        Case 2

            If Not TieneObjetos(475, 1, UserIndex) Then

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(475, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 1118
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ahí tienes." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))

            End If

        Case 3

            If Not TieneObjetos(476, 1, UserIndex) Then

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(476, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 1119
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ahí tienes." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))

            End If

        Case 4

            If Not TieneObjetos(1117, 1, UserIndex) Then

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(1117, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 1983
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ahí tienes." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))

            End If

        Case 5

            If Not TieneObjetos(1118, 1, UserIndex) Then

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(1118, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 475
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ahí tienes." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))

            End If

        Case 6

            If Not TieneObjetos(1119, 1, UserIndex) Then

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(1119, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 476
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbBlue & "°" & "Ahí tienes." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))

            End If

    End Select

End Sub

Public Sub CambiarBarcoAbbadon(ByVal Tipo As Integer, ByVal UserIndex As Integer)

    Dim Objeto As Obj

    Select Case Tipo

        Case 1

            If Not TieneObjetos(1983, 1, UserIndex) Then

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(1983, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 1120
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ahí tienes." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))

            End If

        Case 2

            If Not TieneObjetos(475, 1, UserIndex) Then

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(475, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 1121
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ahí tienes." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))

            End If

        Case 3

            If Not TieneObjetos(476, 1, UserIndex) Then

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(476, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 1122
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ahí tienes." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))

            End If

        Case 4

            If Not TieneObjetos(1120, 1, UserIndex) Then

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(1120, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 1983
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ahí tienes." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))

            End If

        Case 5

            If Not TieneObjetos(1121, 1, UserIndex) Then

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(1121, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 475
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ahí tienes." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))

            End If

        Case 6

            If Not TieneObjetos(1122, 1, UserIndex) Then

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Se puede saber donde esta el barco? :P" & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
            Else
                Call QuitarObjetos(1122, 1, UserIndex)
                Objeto.Amount = 1
                Objeto.ObjIndex = 476
                Call MeterItemEnInventario(UserIndex, Objeto)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & vbRed & "°" & "Ahí tienes." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))

            End If

    End Select

End Sub


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

Public Function TituloNemesis(ByVal UserIndex As Integer) As String

    Dim tStr As String

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

    TituloNemesis = tStr

End Function

Public Function TituloTemplario(ByVal UserIndex As Integer) As String

    Dim tStr As String
   
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

    TituloTemplario = tStr

End Function

Public Function TieneFaccion(ByVal UserIndex As Integer) As Boolean
     
     With UserList(UserIndex)
          
          If .Faccion.ArmadaReal = 1 Or .Faccion.FuerzasCaos = 1 Or .Faccion.Templario = 1 Or .Faccion.Nemesis = 1 Then
              TieneFaccion = True
              Exit Function
          End If
          
     End With
     
End Function

Function UseRangeFragata(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean

    If ObjIndex = 1117 And UserList(UserIndex).Stats.ELV < 25 Then
        UseRangeFragata = False
        Exit Function

    End If

    If ObjIndex = 1118 And UserList(UserIndex).Faccion.RecompensasReal < SegundoRango Then
        UseRangeFragata = False
        Exit Function

    End If

    If ObjIndex = 1119 And UserList(UserIndex).Faccion.RecompensasReal < TercerRango Then
        UseRangeFragata = False
        Exit Function

    End If

    If ObjIndex = 1120 And UserList(UserIndex).Stats.ELV < 25 Then
        UseRangeFragata = False
        Exit Function

    End If

    If ObjIndex = 1121 And UserList(UserIndex).Faccion.RecompensasCaos < SegundoRango Then
        UseRangeFragata = False
        Exit Function

    End If

    If ObjIndex = 1122 And UserList(UserIndex).Faccion.RecompensasCaos < TercerRango Then
        UseRangeFragata = False
        Exit Function

    End If

    If ObjIndex = 1350 And UserList(UserIndex).Stats.ELV < 25 Then
        UseRangeFragata = False
        Exit Function

    End If

    If ObjIndex = 1351 And UserList(UserIndex).Faccion.RecompensasTemplaria < SegundoRango Then
        UseRangeFragata = False
        Exit Function

    End If

    If ObjIndex = 1352 And UserList(UserIndex).Faccion.RecompensasTemplaria < TercerRango Then
        UseRangeFragata = False
        Exit Function

    End If

    If ObjIndex = 1580 And UserList(UserIndex).Stats.ELV < 25 Then
        UseRangeFragata = False
        Exit Function

    End If

    If ObjIndex = 1581 And UserList(UserIndex).Faccion.RecompensasNemesis < SegundoRango Then
        UseRangeFragata = False
        Exit Function

    End If

    If ObjIndex = 1582 And UserList(UserIndex).Faccion.RecompensasNemesis < TercerRango Then
        UseRangeFragata = False
        Exit Function

    End If

    UseRangeFragata = True

End Function

Public Function RangoFaccion(ByVal UserIndex As Integer) As Integer
      
      With UserList(UserIndex)
           
           If .Faccion.ArmadaReal = 1 Then
               RangoFaccion = .Faccion.RecompensasReal
               Exit Function
           ElseIf .Faccion.FuerzasCaos = 1 Then
                RangoFaccion = .Faccion.RecompensasCaos
                Exit Function
           ElseIf .Faccion.Nemesis = 1 Then
                RangoFaccion = .Faccion.RecompensasNemesis
                Exit Function
            ElseIf .Faccion.Templario = 1 Then
                 RangoFaccion = .Faccion.RecompensasTemplaria
                 Exit Function
            End If
           
      End With
      
      RangoFaccion = 0
      
End Function
