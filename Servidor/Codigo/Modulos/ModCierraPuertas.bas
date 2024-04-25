Attribute VB_Name = "ModCierraPuertas"
Option Explicit

Public Type tNpcPuertas
      NpcIndex As Integer
      OrigPosX As Byte
      OrigPosY As Byte
      heading As Byte
End Type

Public Type tCierraPuertas
     
     npc As tNpcPuertas
     Map As Integer
     PosX As Byte
     PosY As Byte
     Active As Byte
     
End Type

Public CierraPuertas(1 To 302) As tCierraPuertas

Const NumPuertas As Integer = 302

Public Sub CargaPuertas()
       
    Dim i              As Integer

    Dim LoopC          As Integer
    
    Dim PuertaPos      As WorldPos

    Dim NpcPos         As WorldPos

    Dim X              As Integer

    Dim Y              As Integer
    
    Dim Distanciadores As Integer
    
    Dim Cant           As Integer ' Variable de pruebas.
       
    'Primero buscamos todas las puertas de mapas.
    For i = 1 To NumMaps
               
        For X = 1 To 100
            For Y = 1 To 100
                
                If MapData(i, X, Y).OBJInfo.ObjIndex > 0 Then
                    If ObjData(MapData(i, X, Y).OBJInfo.ObjIndex).ObjType = eOBJType.otPuertas Then
                        
                        Cant = Cant + 1
                        CierraPuertas(Cant).Map = i
                        CierraPuertas(Cant).PosX = X
                        CierraPuertas(Cant).PosY = Y

                        'Call WriteVar(App.Path & "\Puertas.dat", "Puertas", "Puertas" & Cant, "Puerta encontrada mapa: " & i & " Pos X: " & X & " Pox Y:" & Y)
                        'Debug.Print "Puerta encontrada mapa: " & i & " Pos X: " & X & " Pox Y:" & Y
                        'Cant = Cant + 1
                    End If

                End If
                
            Next Y
        Next X
            
    Next i

    Cant = 0
    'Buscamos NPC's de casa posicion de puertas encontradas.
         
    For i = 1 To NumMaps

        For X = 1 To 100
            For Y = 1 To 100

                For LoopC = 1 To 302

                    PuertaPos.Map = CierraPuertas(LoopC).Map
                    PuertaPos.X = CierraPuertas(LoopC).PosX
                    PuertaPos.Y = CierraPuertas(LoopC).PosY

                    If MapData(i, X, Y).NpcIndex > 0 Then

                        If Npclist(MapData(i, X, Y).NpcIndex).Comercia = 1 Then

                            NpcPos.Map = i
                            NpcPos.X = X
                            NpcPos.Y = Y

                            If NpcPos.Map = PuertaPos.Map Then

                                If Distancia(NpcPos, PuertaPos) <= 5 Then
                                    If CierraPuertas(LoopC).Active = 0 Then
                                        CierraPuertas(LoopC).npc.NpcIndex = MapData(i, X, Y).NpcIndex
                                        CierraPuertas(LoopC).npc.OrigPosX = X
                                        CierraPuertas(LoopC).npc.OrigPosY = Y
                                        CierraPuertas(LoopC).npc.heading = Npclist(MapData(i, X, Y).NpcIndex).char.heading
                                        CierraPuertas(LoopC).Active = 1
                                    ElseIf CierraPuertas(LoopC).Active = 1 Then

                                        If X > CierraPuertas(LoopC).npc.OrigPosX Then
                                            CierraPuertas(LoopC).npc.NpcIndex = MapData(i, X, Y).NpcIndex
                                            CierraPuertas(LoopC).npc.OrigPosX = X
                                            CierraPuertas(LoopC).npc.OrigPosY = Y
                                            CierraPuertas(LoopC).npc.heading = Npclist(MapData(i, X, Y).NpcIndex).char.heading
                                            Debug.Print "Un cambio"

                                        End If

                                    End If

                                End If

                            End If

                        End If

                    End If

                Next LoopC
            Next Y
        Next X

    Next i
    
    Debug.Print "NPC Comerciantes encontrados! " & Cant
    
End Sub

Public Sub RevisaPuertasAbierta()
       
    Dim X  As Byte

    Dim Y  As Byte
       
    Dim UI As Integer

    Dim i  As Integer
       
    For i = 1 To 302
        For X = 1 To 100
            For Y = 1 To 100
            
                If CierraPuertas(i).Map = 0 Then
                    Exit Sub

                End If
            
                UI = MapData(CierraPuertas(i).Map, X, Y).UserIndex
            
                If UI > 0 Then
              
                    If CierraPuertas(i).Active = 1 Then
                           
                           If ObjData(MapData(CierraPuertas(i).Map, CierraPuertas(i).PosX, CierraPuertas(i).PosY).OBJInfo.ObjIndex).Cerrada = 0 Then
                                If Npclist(CierraPuertas(i).npc.NpcIndex).flags.CerrarPuertas = 0 Then
                                    Call SendData(ToPCArea, UI, CierraPuertas(i).Map, "||" & vbWhite & "°" & MsjPuerta(RandomNumber(1, 9)) & "°" & CStr(Npclist(CierraPuertas(i).npc.NpcIndex).char.CharIndex))
                                    Npclist(CierraPuertas(i).npc.NpcIndex).flags.CerrarPuertas = 1
                                    Npclist(CierraPuertas(i).npc.NpcIndex).flags.StartPuerta = GetTickCount()
                                'ElseIf Npclist(CierraPuertas(i).npc.NpcIndex).flags.CerrarPuertas = 1 Then
                                
                                End If
                           End If
                           
                    End If
              
                End If
           
            Next Y
        Next X
    Next i
       
End Sub

Public Sub IrCerrarPuerta()

    Dim i         As Long

    Dim X         As Byte

    Dim Y         As Byte
    
    Dim tHeading  As Byte
    
    Dim NpcPos    As WorldPos

    Dim PuertaPos As WorldPos
    
    Dim UI        As Integer
      
    For i = 1 To 302
        For X = 1 To 100
            For Y = 1 To 100
            
                If CierraPuertas(i).Map = 0 Then
                    Exit Sub

                End If
                 
                UI = MapData(CierraPuertas(i).Map, X, Y).UserIndex
                    
                If UI > 0 Then
                    If CierraPuertas(i).Active = 1 Then
                        If Npclist(CierraPuertas(i).npc.NpcIndex).flags.CerrarPuertas = 1 Then
                            Npclist(CierraPuertas(i).npc.NpcIndex).flags.StartPuerta = Npclist(CierraPuertas(i).npc.NpcIndex).flags.StartPuerta - 100000
                            
                            If Npclist(CierraPuertas(i).npc.NpcIndex).flags.StartPuerta <= 10 Then
                             
                                NpcPos.Map = CierraPuertas(i).Map
                                NpcPos.X = CierraPuertas(i).npc.OrigPosX
                                NpcPos.Y = CierraPuertas(i).npc.OrigPosY
                             
                                PuertaPos.Map = CierraPuertas(i).Map
                                PuertaPos.X = CierraPuertas(i).PosX
                                PuertaPos.Y = CierraPuertas(i).PosY
                             
                                If ObjData(MapData(CierraPuertas(i).Map, CierraPuertas(i).PosX, CierraPuertas(i).PosY).OBJInfo.ObjIndex).Cerrada = 0 Then
                                    If Distancia(Npclist(CierraPuertas(i).npc.NpcIndex).pos, PuertaPos) = 1 Then
                                
                                    Else
                                        tHeading = FindDirectionEAO(Npclist(CierraPuertas(i).npc.NpcIndex).pos, PuertaPos)
                                        Call MoveNPCChar(CierraPuertas(i).npc.NpcIndex, tHeading)

                                    End If
                              
                                ElseIf ObjData(MapData(CierraPuertas(i).Map, CierraPuertas(i).PosX, CierraPuertas(i).PosY).OBJInfo.ObjIndex).Cerrada = 1 Then

                                End If
                            
                            End If
                        
                        End If

                    End If
                
                End If

            Next Y
        Next X
    Next i
      
End Sub

Public Function MsjPuerta(ByVal Item As Byte) As String
    
    Select Case Item
        
        Case 1
        MsjPuerta = "HAY QUE CERRAR LA PUERTAS!!!"
        Exit Function
        
        Case 2
        MsjPuerta = "En mi negocio quiero las puertas cerradas, a ver si os enteráis."
        Exit Function
        
        Case 3
        MsjPuerta = "Está bien. Ya cierro la puerta yo. ¬¬"
        Exit Function
        
        Case 4
        MsjPuerta = "Qué frío! Siempre se dejan la puerta abierta!"
        Exit Function
        
        Case 5
        MsjPuerta = "Mal educados! Hay que cerrar la puerta al salir!"
        Exit Function
        
        Case 6
        MsjPuerta = "Haced el favor de cerrar la puerta. Gracias."
        Exit Function
        
        Case 7
        MsjPuerta = "No tenéis frío?"
        Exit Function
        
        Case 8
        MsjPuerta = "Otra vez se olvidaron la puerta! Dónde tendrán la cabeza!"
        Exit Function
        
        Case 9
         MsjPuerta = "No tienen remedio. Otra vez la puerta."
         Exit Function
        
    End Select
    
End Function
