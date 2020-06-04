Attribute VB_Name = "mod_CasaAbandonada"
Option Explicit

Private Const FxAlgo             As Byte = 1
Private Const WavCasa            As Byte = 113
Private Const TriggerMuerteCasa  As Byte = 11
'El trigger de la habitación es 1, funca, pero cuando sale de la habitación
'te crashea el juego. Dejo otra numeración para que de momento no afecte.

Public Const MapaCasaAbandonada1 As Integer = 85
Public Const MapaCasaAbandonada2 As Integer = 85

Private Const MapaFuera1         As Byte = 139
Private Const MapaXFuera1        As Byte = 50
Private Const MapaYFuera1        As Byte = 44

Private Const MapaFuera2         As Byte = 139
Private Const MapaXFuera2        As Byte = 69
Private Const MapaYFuera2        As Byte = 45

Public Sub Efecto_CaminoCasaEncantada(ByVal UserIndex As Integer)

    With UserList(UserIndex)
    
        If EsGmChar(.Name) Then Exit Sub
        
        If MapData(.pos.Map, .pos.X, .pos.Y).trigger = TriggerMuerteCasa Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La Habitación de Sangre te ha matado." & FONTTYPE_TALK)
            Call UserDie(UserIndex)

        End If

        Select Case RandomNumber(1, 30000)

            Case 150
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los Espiritus de la Casa te hacen perder el inventario." & FONTTYPE_TALK)
                
                Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, WavCasa)
                
                Call TirarTodosLosItemsNoNewbies(UserIndex)
    
                Call SendData(SendTarget.ToPCArea, UserIndex, .pos.Map, "CFX" & .char.CharIndex & "," & FxAlgo & ",0")
                Exit Sub
                
            Case 150 To 160

                If .pos.Map = MapaCasaAbandonada1 Then
                    If UserList(UserIndex).Stats.GLD > 10000 Then
                        Call TirarOro(10000, UserIndex)
                    Else
                    
                        Exit Sub

                    End If
                    
                Else

                    If UserList(UserIndex).Stats.GLD > 30000 Then
                        Call TirarOro(30000, UserIndex)
                    Else
                    
                        Exit Sub

                    End If

                End If
                                        
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los Espiritus de la Casa te hacen perder Oro." & FONTTYPE_TALK)
                
                Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, WavCasa)

                Call EnviarOro(UserIndex)
               
                Exit Sub

            Case 161 To 171
    
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los Espiritus de la Casa te teleportan fuera de ella." & FONTTYPE_TALK)
                    
                Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, WavCasa)

                If .pos.Map = MapaCasaAbandonada1 Then
                    Call WarpUserChar(UserIndex, MapaFuera1, MapaXFuera1, MapaYFuera1, True)
                Else
                    Call WarpUserChar(UserIndex, MapaFuera2, MapaXFuera2, MapaYFuera2, True)

                End If

            Case 171 To 181
          
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los Espiritus de la Casa te han Paralizado." & FONTTYPE_TALK)
                    
                Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, WavCasa)
              
                Call SendData(SendTarget.ToPCArea, UserIndex, .pos.Map, "CFX" & .char.CharIndex & "," & FxAlgo & ",0")
                  
                .flags.Paralizado = 1
                .Counters.Paralisis = IntervaloParalizado
         
                Call SendData(SendTarget.toIndex, UserIndex, 0, "PARADOW")
                Call SendData(SendTarget.toIndex, UserIndex, 0, "PU" & .pos.X & "," & .pos.Y)
             
        End Select

    End With

End Sub

Public Sub Efecto_AccionCasaEncantada(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    With UserList(UserIndex)

        If EsGmChar(.Name) Then Exit Sub
        If NpcIndex = 0 Then Exit Sub
        
        Select Case RandomNumber(1, 3000)

            Case 1 To 20
     
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los Espiritus de la Casa te han curado el Npc." & FONTTYPE_TALK)
         
                Call SendData(SendTarget.ToPCArea, UserIndex, .pos.Map, "CFX" & Npclist(NpcIndex).char.CharIndex & "," & FXIDs.FXWARP & ",0")
         
                Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, WavCasa)
                Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
                Exit Sub

            Case 21 To 45
                Dim MiPos As WorldPos
                MiPos.Map = .pos.Map
                MiPos.X = .pos.X - 1
                MiPos.Y = .pos.Y

                Call SpawnNpc(550, MiPos, True, False)

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los Espiritus de la Casa te han invocado una Bruja." & FONTTYPE_TALK)
                    
                Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, WavCasa)
                Exit Sub

        End Select

    End With

End Sub

