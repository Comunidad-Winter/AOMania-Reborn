Attribute VB_Name = "TCP_HandleData3"
Option Explicit

Public Sub HandleData_3(ByVal UserIndex As Integer, rData As String, ByRef Procesado As Boolean)

    Dim n As Integer
    Dim LoopC As Integer
    Dim Rs As Integer
    
    If UCase$(Left$(rData, 9)) = "/VERPARTY" Then
           Call VerParty(UserIndex)
         Exit Sub
    End If
    
    If UCase$(Left$(rData, 14)) = "/RESETEAARMADA" Then
    
         If UserList(UserIndex).flags.Muerto = 1 Then
             Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estás muerto!!" & FONTTYPE_INFO)
             Exit Sub
        End If
       
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.armada Then
           
            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes acercarte más." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If Not TieneArmada(UserIndex) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbYellow & "°" & _
                    "No perteneces a ninguna armada!!!" & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If
            
             If PiedraFaccion(UserIndex) > 0 Then
              n = PiedraFaccion(UserIndex)
              Else
              Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbYellow & "°" & _
                    "No tienes la piedra vuelve cuando la tengas." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                Exit Sub
             End If
            
            Rs = RandomNumber(1, 10)
            
            If Rs < 5 Then
            Call QuitarObjetos(n, 1, UserIndex)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbYellow & "°" & _
                    "El usuario " & UserList(UserIndex).Name & " ha fallado Cambio de armada." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                Exit Sub
            End If
            
        If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
            Call PerderItemsFaccionarios(UserIndex, UserList(UserIndex).Faccion.ArmaduraFaccionaria)
            Call QuitarObjetos(n, 1, UserIndex)
            UserList(UserIndex).Faccion.ArmadaReal = 0
            UserList(UserIndex).Faccion.FEnlistado = 0
            UserList(UserIndex).Faccion.NextRecompensas = 0
            UserList(UserIndex).Faccion.RecibioArmaduraReal = 0
            UserList(UserIndex).Faccion.RecibioExpInicialReal = 0
            UserList(UserIndex).Faccion.RecompensasReal = 0
        ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
            Call PerderItemsFaccionarios(UserIndex, UserList(UserIndex).Faccion.ArmaduraFaccionaria)
            Call QuitarObjetos(n, 1, UserIndex)
            UserList(UserIndex).Faccion.FuerzasCaos = 0
            UserList(UserIndex).Faccion.FEnlistado = 0
            UserList(UserIndex).Faccion.NextRecompensas = 0
            UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0
            UserList(UserIndex).Faccion.RecibioExpInicialCaos = 0
            UserList(UserIndex).Faccion.RecompensasCaos = 0
        ElseIf UserList(UserIndex).Faccion.Nemesis = 1 Then
            Call PerderItemsFaccionarios(UserIndex, UserList(UserIndex).Faccion.ArmaduraFaccionaria)
            Call QuitarObjetos(n, 1, UserIndex)
            UserList(UserIndex).Faccion.Nemesis = 0
            UserList(UserIndex).Faccion.FEnlistado = 0
            UserList(UserIndex).Faccion.NextRecompensas = 0
            UserList(UserIndex).Faccion.RecibioArmaduraNemesis = 0
            UserList(UserIndex).Faccion.RecibioExpInicialNemesis = 0
            UserList(UserIndex).Faccion.RecompensasNemesis = 0
        ElseIf UserList(UserIndex).Faccion.Templario = 1 Then
        Call PerderItemsFaccionarios(UserIndex, UserList(UserIndex).Faccion.ArmaduraFaccionaria)
            Call QuitarObjetos(n, 1, UserIndex)
            UserList(UserIndex).Faccion.Templario = 0
            UserList(UserIndex).Faccion.FEnlistado = 0
            UserList(UserIndex).Faccion.NextRecompensas = 0
            UserList(UserIndex).Faccion.RecibioArmaduraTemplaria = 0
            UserList(UserIndex).Faccion.RecibioExpInicialTemplaria = 0
            UserList(UserIndex).Faccion.RecompensasTemplaria = 0
        End If
        
        UserList(UserIndex).Reputacion.AsesinoRep = 0
        UserList(UserIndex).Reputacion.BandidoRep = 0
        UserList(UserIndex).Reputacion.BurguesRep = 0
        UserList(UserIndex).Reputacion.NobleRep = 0
        UserList(UserIndex).Reputacion.PlebeRep = 0
        
        Call WarpUserChar(UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y)
        
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbYellow & "°" & _
                    "El usuario " & UserList(UserIndex).Name & " se ha salido de la armada poniendolo todo a 0." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
        End If
    
    End If
    
    If UCase$(Left$(rData, 6)) = "/PARTY" Then
        
        If UserList(UserIndex).flags.TargetUser = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Antes debes hacer click a un usuario." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(UserList(UserList(UserIndex).flags.TargetUser).pos, UserList(UserIndex).pos) > mdParty.MAXDISTANCIAINGRESOPARTY Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estás demasiado lejos de tu compañero." & FONTTYPE_INFO)
           Exit Sub
        End If
        
        If UserList(UserIndex).Stats.ELV < mdParty.MINPARTYLEVEL Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes ser por lo menos nivel 13 para poder organizar una party." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserList(UserIndex).flags.TargetUser).Stats.ELV < mdParty.MINACLEVEL Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El otro usuario debe ser por lo menos nivel " & mdParty.MINACLEVEL & " para poder entrar en una party." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡No puedes!!" & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserList(UserIndex).flags.TargetUser).flags.Privilegios > PlayerType.User Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Los GMs no pueden participar en party's" & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Abs(val(UserList(UserIndex).Stats.ELV - UserList(UserList(UserIndex).flags.TargetUser).Stats.ELV)) > mdParty.MAXPARTYDELTALEVEL Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes meter a " & UserList(UserList(UserIndex).flags.TargetUser).Name & " en tu party porque os lleváis más de " & mdParty.MAXPARTYDELTALEVEL & " niveles de diferencia." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.TargetUser > 0 Then
            
            rData = UserList(UserIndex).flags.TargetUser
            
            If UserList(UserIndex).PartyIndex = 0 Then
                  If Not mdParty.PuedeCrearParty(UserIndex) Then Exit Sub
                  Call mdParty.CrearParty(UserIndex)
            End If
                  
            Call mdParty.EnviarParty(UserIndex, rData)
            
          End If
    End If
    
    If UCase$(Left$(rData, 8)) = "/ACEPTAR" Then
           Call mdParty.AceptarParty(UserIndex)
    End If
    
     If UCase$(Left$(rData, 9)) = "/CANCELAR" Then
           Call mdParty.CancelarParty(UserIndex)
    End If
    
   If UCase$(Left$(rData, 6)) = "/ANGEL" Then
       
       With UserList(UserIndex)
         
            If .Metamorfosis.Angel = 1 Then Exit Sub
            If .Metamorfosis.Demonio = 1 Then Exit Sub
            If .flags.Mimetizado = 1 Then Exit Sub
            If .flags.Meditando = 1 Then Exit Sub
            If .Stats.ELV < STAT_MAXELV Then Exit Sub
            If Criminal(UserIndex) Then Exit Sub
            If .Faccion.FuerzasCaos = 1 Or .Faccion.Nemesis = 1 Then Exit Sub
            If .flags.Navegando = 1 Then Exit Sub
            If .pos.Map = 48 Then Exit Sub
            
            If .Stats.MinSta < .Stats.MaxSta Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡No tienes suficiente energía!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .flags.Invisible = 1 Or .flags.Oculto = 1 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Si estás invisible no puedes transformarte." & FONTTYPE_INFO)
               Exit Sub
            End If
            
            If ZonaDuelos(.pos.Map, .pos.X, .pos.Y) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en duelos no puedes transformarte." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .pos.Map = CastilloNorte Or .pos.Map = CastilloOeste Or .pos.Map = CastilloEste Or .pos.Map = CastilloSur Or .pos.Map = MapaFortaleza Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes transformarte." & FONTTYPE_INFO)
                Exit Sub
            ElseIf .pos.Map = mapainvo Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes convertirte estando en la sala de invocaciones." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            .CharMimetizado.Body = .char.Body
            .CharMimetizado.Head = .char.Head
            .CharMimetizado.CascoAnim = .char.CascoAnim
      
            .CharMimetizado.ShieldAnim = .char.ShieldAnim
            .CharMimetizado.WeaponAnim = .char.WeaponAnim
            
            .CharMimetizado.Alas = .char.Alas
        
            .flags.Mimetizado = 1
           
            .char.Body = 347
            .char.Head = 0
            .char.WeaponAnim = 2
            .char.ShieldAnim = 2
            .char.Alas = 0
       
            Call ChangeUserChar(SendTarget.ToMap, 0, .pos.Map, UserIndex, .char.Body, .char.Head, _
                    .char.heading, .char.WeaponAnim, .char.ShieldAnim, .char.CascoAnim, .char.Alas)
            
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList( _
                UserIndex).char.CharIndex & "," & 1 & "," & 1)
                
       End With
       
    End If
    
    If UCase$(Left$(rData, 8)) = "/DEMONIO" Then
       
       With UserList(UserIndex)
         
            If .Metamorfosis.Angel = 1 Then Exit Sub
            If .Metamorfosis.Demonio = 1 Then Exit Sub
            If .flags.Mimetizado = 1 Then Exit Sub
            If .flags.Meditando = 1 Then Exit Sub
            If .Stats.ELV < STAT_MAXELV Then Exit Sub
            If Not Criminal(UserIndex) Then Exit Sub
            If .Faccion.ArmadaReal = 1 Or .Faccion.Templario = 1 Then Exit Sub
            If .flags.Navegando = 1 Then Exit Sub
            If .pos.Map = 48 Then Exit Sub
            
            If .Stats.MinSta < .Stats.MaxSta Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡No tienes suficiente energía!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .flags.Invisible = 1 Or .flags.Oculto = 1 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Si estás invisible no puedes transformarte." & FONTTYPE_INFO)
               Exit Sub
            End If
            
            If ZonaDuelos(.pos.Map, .pos.X, .pos.Y) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en duelos no puedes transformarte." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .pos.Map = CastilloNorte Or .pos.Map = CastilloOeste Or .pos.Map = CastilloEste Or .pos.Map = CastilloSur Or .pos.Map = MapaFortaleza Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes transformarte." & FONTTYPE_INFO)
                Exit Sub
            ElseIf .pos.Map = mapainvo Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes convertirte estando en la sala de invocaciones." & FONTTYPE_INFO)
                Exit Sub
            End If
            
           .CharMimetizado.Body = .char.Body
            .CharMimetizado.Head = .char.Head
            .CharMimetizado.CascoAnim = .char.CascoAnim
      
            .CharMimetizado.ShieldAnim = .char.ShieldAnim
            .CharMimetizado.WeaponAnim = .char.WeaponAnim
            .CharMimetizado.Alas = .char.Alas
        
            .flags.Mimetizado = 1
            
            .char.Body = 348
            .char.Head = 0
            .char.WeaponAnim = 2
            .char.ShieldAnim = 2
            .char.Alas = 0
        
            Call ChangeUserChar(SendTarget.ToMap, 0, .pos.Map, UserIndex, .char.Body, .char.Head, _
                    .char.heading, .char.WeaponAnim, .char.ShieldAnim, .char.CascoAnim, .char.Alas)
            
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList( _
                UserIndex).char.CharIndex & "," & 1 & "," & 1)
                
       End With
       
    End If
    
     If UCase$(Left$(rData, 8)) = "/OLVIDAR" Then
          If Not ValidMap(UserList(UserIndex).pos.Map) = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tienes que seleccionar un personaje, hace click izquierdo sobre el." & _
                        FONTTYPE_INFO)
                Exit Sub
         End If
         
         If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.OlvidarHechizo Then
         
             If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 5 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                    Exit Sub
                End If
                
                If OroHechizo > UserList(UserIndex).Stats.GLD Then
                   Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes oro suficiente." & FONTTYPE_TALKMSG)
                   Exit Sub
                End If
          
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "HECA" & FONTTYPE_INFO)
         End If
         
     End If
    
     
     If UCase$(Left$(rData, 8)) = "/DRAGON " Then
         rData = Right$(rData, Len(rData) - 8)
           
           Dim Cabeza As Integer
         
          If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.CambiaCabeza Then
         
             If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 5 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                    Exit Sub
             End If
             
             Select Case UCase(rData)
                  Case "ROJA"
                     Cabeza = CabezaDragon.Roja
                  Case "NEGRA"
                      Cabeza = CabezaDragon.negra
                  Case "VERDE"
                      Cabeza = CabezaDragon.Verde
                  Case "LILA"
                      Cabeza = CabezaDragon.lila
                  Case "BLANCA"
                      Cabeza = CabezaDragon.Blanca
                  Case "NARANJA"
                      Cabeza = CabezaDragon.naranja
                  Case "AZUL"
                       Cabeza = CabezaDragon.Azul
                  Case Else
                       Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbCyan & "°" & "No conozco esa armadura que dices." & "°" _
                            & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                       Exit Sub
             End Select
             
             If Not TieneObjetos(Cabeza, 10, UserIndex) Then
                 Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & "&H4580FF" & "°" & "¿¿Dónde están esas 10 cabezas??" & "°" _
                            & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                 Exit Sub
             End If
             
             Call QuitarObjetos(Cabeza, 10, UserIndex)
             Call DarCabezaDragon(UserIndex, rData)
                
        End If
     
     End If
    
    Procesado = False
End Sub


Function PiedraFaccion(ByVal UserIndex As Integer) As Integer
      
      If TieneObjetos(1002, 1, UserIndex) Then
          PiedraFaccion = 1002
          Exit Function
      End If
      
      If TieneObjetos(1204, 1, UserIndex) Then
         PiedraFaccion = 1204
         Exit Function
      End If

     PiedraFaccion = 0
        
End Function

Function TieneArmada(ByVal UserIndex As Integer) As Boolean

    If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
       TieneArmada = True
       Exit Function
    ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
       TieneArmada = True
       Exit Function
    ElseIf UserList(UserIndex).Faccion.Nemesis = 1 Then
        TieneArmada = True
        Exit Function
    ElseIf UserList(UserIndex).Faccion.Templario = 1 Then
        TieneArmada = True
        Exit Function
    End If
    
    TieneArmada = False
    
End Function

