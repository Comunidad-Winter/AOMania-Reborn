Attribute VB_Name = "mod_CasaAbandonada"
Option Explicit

Public Const MapaCasaAbandonada1 As Integer = 85

Private Const MapaFuera1         As Byte = 139
Private Const MapaXFuera1        As Byte = 50
Private Const MapaYFuera1        As Byte = 44

Public Const NpcBruja As Integer = 550
Public Const NpcThorn As Integer = 584

Private Const IntervaloCerdo As Integer = 2400 ' 2 minutos
Public Const IntervaloRenaceThorn As Integer = 3600 ' 1 hora

Public NpcThornVive As Boolean

Public Sub LoadCasaEncantada()
                Dim MiPos As WorldPos
                MiPos.Map = MapaCasaAbandonada1
                MiPos.X = RandomNumber(10, 91)
                MiPos.Y = RandomNumber(13, 83)

                Call SpawnNpc(NpcThorn, MiPos, True, False)
                NpcThornVive = True
End Sub

Public Sub RenaceThorn()
              
                Dim MiPos As WorldPos
                MiPos.Map = MapaCasaAbandonada1
                MiPos.X = RandomNumber(10, 91)
                MiPos.Y = RandomNumber(13, 83)

                Call SpawnNpc(NpcThorn, MiPos, True, False)
                NpcThornVive = True
                
                Call SendData(SendTarget.toall, 0, 0, "||Reaparecio Thorn en la Casa Encantada." & FONTTYPE_GUILD)
                Call SendData(SendTarget.toall, 0, 0, "TW3")
                
End Sub


Public Sub Efecto_CaminoCasaEncantada(ByVal UserIndex As Integer)

    With UserList(UserIndex)
    
        If EsGmChar(.Name) Then Exit Sub
        
        If MapData(MapaCasaAbandonada1, .pos.X, .pos.Y).Graphic(2) >= 260 And MapData(MapaCasaAbandonada1, .pos.X, .pos.Y).Graphic(2) <= 265 And UserList(UserIndex).flags.Muerto = 0 Then
             Call SendData(SendTarget.toindex, UserIndex, 0, "||Los espiritus te han provocado una combustión espontánea." & FONTTYPE_TALKMSG)
             Call SendData(SendTarget.toindex, UserIndex, 0, "||Los espiritus te han matado." & FONTTYPE_TALKMSG)
             Call UserDie(UserIndex)
        ElseIf MapData(MapaCasaAbandonada1, .pos.X, .pos.Y).Graphic(2) = 283 Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||Los espiritus te han provocado una combustión espontánea." & FONTTYPE_TALKMSG)
            Call SendData(SendTarget.toindex, UserIndex, 0, "||Los espiritus te han matado." & FONTTYPE_TALKMSG)
            Call UserDie(UserIndex)
        End If
        
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub

        Select Case RandomNumber(1, 3000)
        
           Case 140 To 149
    
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Los espiritus de la casa te han echado." & FONTTYPE_TALKMSG)
                    
                Call WarpUserChar(UserIndex, MapaFuera1, MapaXFuera1, MapaYFuera1, True)

            Case 150
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Los espiritus de la casa te tiran objetos al suelo." & FONTTYPE_TALKMSG)
                  
                Call TirarTodosLosItemsNoNewbies(UserIndex)
    
                Exit Sub
                
            Case 150 To 160

                    If UserList(UserIndex).Stats.GLD > 30000 Then
                        Call TirarOro(30000, UserIndex)
                          Else
                        Exit Sub
                    End If
                                        
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Los espiritus de la casa te tiran oro al suelo." & FONTTYPE_TALKMSG)
                
                Call EnviarOro(UserIndex)
               
                Exit Sub

            Case 171 To 181
          
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Los espiritus te han paralizado." & FONTTYPE_TALKMSG)
                    
                .flags.Paralizado = 1
                .Counters.Paralisis = IntervaloParalizado
         
                Call SendData(SendTarget.toindex, UserIndex, 0, "PARADOW")
                Call SendData(SendTarget.toindex, UserIndex, 0, "PU" & .pos.X & "," & .pos.Y)
             
        End Select

    End With

End Sub

Public Sub Efecto_AccionCasaEncantada(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
  
  
    With UserList(UserIndex)

        If EsGmChar(.Name) Then Exit Sub
        If NpcIndex = 0 Then Exit Sub
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        
        Select Case RandomNumber(1, 3000)

            Case 1 To 20
     
                Call SendData(SendTarget.toindex, UserIndex, 0, "||Los Espíritus de la casa han Sanado al " & Npclist(NpcIndex).Name & FONTTYPE_TALK)
         
                Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
                
                Exit Sub

            Case 21 To 45
                Dim MiPos As WorldPos
                MiPos.Map = .pos.Map
                MiPos.X = .pos.X - 1
                MiPos.Y = .pos.Y

                Call SpawnNpc(NpcBruja, MiPos, True, False)

                Call SendData(SendTarget.toindex, UserIndex, 0, "||Los Espíritus de la casa te han invocado una Bruja." & FONTTYPE_TALK)
                    
               Exit Sub
               
            Case 50 To 3000
                
            UserList(UserIndex).CharMimetizado.Body = .char.Body
            UserList(UserIndex).CharMimetizado.Head = .char.Head
            UserList(UserIndex).CharMimetizado.CascoAnim = .char.CascoAnim
      
            UserList(UserIndex).CharMimetizado.ShieldAnim = .char.ShieldAnim
            UserList(UserIndex).CharMimetizado.WeaponAnim = .char.WeaponAnim
            
            UserList(UserIndex).CharMimetizado.Alas = .char.Alas
        
            UserList(UserIndex).flags.Mimetizado = 1
           
            UserList(UserIndex).char.Body = 6
            UserList(UserIndex).char.Head = 0
            UserList(UserIndex).char.WeaponAnim = 2
            UserList(UserIndex).char.ShieldAnim = 2
            UserList(UserIndex).char.Alas = 0
       
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, _
                    UserList(UserIndex).char.Heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)
            
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList( _
                UserIndex).char.CharIndex & "," & 1 & "," & 1)
            
            Call SendData(SendTarget.toindex, UserIndex, 0, "||Los espiritus de la casa te han transformado en cerdo." & FONTTYPE_TALKMSG)
            
            UserList(UserIndex).Counters.Cerdo = IntervaloCerdo
                
              Exit Sub

        End Select

    End With
   

End Sub

Public Sub EfectoCerdo(ByVal UserIndex As Integer)
        
        If UserList(UserIndex).Counters.Cerdo > 0 Then
            UserList(UserIndex).Counters.Cerdo = UserList(UserIndex).Counters.Cerdo - 1
        End If
       
       If UserList(UserIndex).Counters.Cerdo = 0 Then
            UserList(UserIndex).char.Body = UserList(UserIndex).CharMimetizado.Body
            UserList(UserIndex).char.Head = UserList(UserIndex).CharMimetizado.Head
            UserList(UserIndex).char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
            UserList(UserIndex).char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
            UserList(UserIndex).char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
             UserList(UserIndex).char.Alas = UserList(UserIndex).CharMimetizado.Alas
        
             UserList(UserIndex).Counters.Mimetismo = 0
             UserList(UserIndex).flags.Mimetizado = 0
       
        Call ChangeUserChar(SendTarget.ToMap, UserIndex, UserList(UserIndex).pos.Map, UserIndex, _
        UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).char.Heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, _
        UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)
        
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & _
                UserList(UserIndex).char.CharIndex & "," & 1 & "," & 1)
        End If
End Sub
