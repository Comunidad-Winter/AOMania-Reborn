Attribute VB_Name = "GuerraBanda"
Option Explicit

Public StatusGuerra As String

Public Const MaxParticipantesGuerra As Integer = 32
Private Const MaxDemonios As Integer = MaxParticipantesGuerra / 2
Private Const MaxAngeles As Integer = MaxParticipantesGuerra / 2

Public TimerGuerra As Long

Public InscripcionBanda As Boolean

Public Const MapaBan     As Integer = 162
Private Const FortaDemon  As Integer = 17
Private Const FortaDemony As Integer = 65
Private Const FortaAngel  As Integer = 81
Private Const FortaAngely As Integer = 32
Private Const EsperaDemonio = 11
Private Const EsperaDemonioy = 89
Private Const EsperaAngel = 87
Private Const EsperaAngely = 9

Private Ban_Luchadores() As Integer
Private Quit_Luchadores() As String
Private CantQuitBan As Integer

Private Demonios         As Integer
Private Angeles          As Integer
Private Cantban           As Integer

Public BanAc As Boolean
Public BanEsp As Boolean
Public CantidadGuerra As Integer

Public Const RecBanOro As Long = 1000000
Public Const RecBanExp As Long = 5000

'Este sub debe ir en el sub main() del modulo general.bas.
Sub LoadGuerras()
  StatusGuerra = "Banda"
End Sub


'Este timer hace funcionamiento de la guerra de banda y del modulo, guerra de medusas.

Sub Timer_GuerradeBanda()
   
    TimerGuerra = TimerGuerra + 1
    
    If StatusGuerra = "Banda" Then
    
    If BanEsp = True Then
        Select Case TimerGuerra
            
            Case 1
            Call PasaTimeBan
                
        End Select
    Else

    Select Case TimerGuerra
      
        Case 1
           Call SendData(SendTarget.toall, 0, 0, "||La proxima Guerra de Banda se jugara dentro de 59 minutos." _
                    & FONTTYPE_GUERRA)
       
        Case 50
            Call SendData(SendTarget.toall, 0, 0, _
                                      "||Quedan 10  minutos para la proxima guerra de bandas (no se pierde inventario)" & FONTTYPE_GUERRA)
            Call Ban_Comienza(32)
            Call SendData(SendTarget.toall, 0, 0, "TW48")
             
        Case 55
            Call SendData(SendTarget.toall, 0, 0, "||Quedan 5 minutos para la proxima guerra de bandas." _
                                      & FONTTYPE_GUERRA)

        Case 56
            Call SendData(SendTarget.toall, 0, 0, "||Quedan 4 minutos para la proxima guerra de bandas." _
                                      & FONTTYPE_GUERRA)
        Case 57
            Call SendData(SendTarget.toall, 0, 0, "||Quedan 3 minutos para la proxima guerra de bandas." _
                                      & FONTTYPE_GUERRA)
        
        Case 58
            Call SendData(SendTarget.toall, 0, 0, "||Quedan 2 minutos para la proxima guerra de bandas." _
                                      & FONTTYPE_GUERRA)
        
        Case 59
            Call SendData(SendTarget.toall, 0, 0, "||Quedan 1 minutos para la proxima guerra de bandas." _
                                      & FONTTYPE_GUERRA)

        Case 60
            Call SendData(SendTarget.toall, 0, 0, _
                                      "||Se cerraron las inscripciones para la guerra de bandas..." _
                                      & FONTTYPE_GUERRA)
           Call SendData(SendTarget.ToMap, 0, MapaBan, _
                                     "||Queda 1 minuto para la guerra de bandas, prepárense..." _
                                     & FONTTYPE_GUERRA)
                                     
           Call SendData(SendTarget.toall, 0, 0, "TW48")
           
            BanAc = False

        Case 61

            If BanAc = True Then
                If CantidadGuerra < 4 Then
                   Call SendData(SendTarget.toall, 0, 0, _
                           "||Se canceló la guerra entre bandas por falta de participantes." & _
                           FONTTYPE_GUERRA)
                    Call Banauto_Cancela
                    bandasqls = 1
                Else
                        Call Banauto_Empieza
                        BanAc = True
                        bandasqls = 1
                End If

            End If
            
        End Select
        
        End If
   
   End If
   
   If StatusGuerra = "Medusa" Then
        
        
        If MedEsp = True Then
           Select Case TimerGuerra
           
                Case 10
                 Call PasaTimeMed
                  
           End Select
                 
        Else
        
        Select Case TimerGuerra
           
           Case 1
           Call SendData(SendTarget.toall, 0, 0, "||La proxima Guerra de Medusas se jugara dentro de 59 minutos." & FONTTYPE_GUERRA)
                               
           Case 50
           Call SendData(SendTarget.toall, 0, 0, "||Quedan 10 minutos para la próxima batalla de las medusas (no se pierde inventario)." & FONTTYPE_GUERRA)
           Call Med_Comienza(32)
           Call SendData(SendTarget.toall, 0, 0, "TW48")
           
           Case 55
           Call SendData(SendTarget.toall, 0, 0, "||Quedan 5 minutos para la proxima batalla de medusas (no se pierde inventario)" & FONTTYPE_GUERRA)
           
           Case 56
           Call SendData(SendTarget.toall, 0, 0, "||Quedan 4 minutos para la proxima batalla de medusas (no se pierde inventario)" & FONTTYPE_GUERRA)
           
           Case 57
           Call SendData(SendTarget.toall, 0, 0, "||Quedan 3 minutos para la proxima batalla de medusas (no se pierde inventario)" & FONTTYPE_GUERRA)
           
           Case 58
           Call SendData(SendTarget.toall, 0, 0, "||Quedan 2 minutos para la proxima batalla de medusas (no se pierde inventario)" & FONTTYPE_GUERRA)
           
           Case 59
           Call SendData(SendTarget.toall, 0, 0, "||Quedan 1 minutos para la proxima batalla de medusas (no se pierde inventario)" & FONTTYPE_GUERRA)
           
           Case 60
           Call SendData(SendTarget.toall, 0, 0, "||Se cerraron las inscripciones para la batalla de las medusas.." & FONTTYPE_GUERRA)
           Call SendData(SendTarget.ToMap, 0, MapaMedusa, "||Queda 1 minuto para la batalla de las medusas, prepárense..." & FONTTYPE_GUERRA)
           Call SendData(SendTarget.toall, 0, 0, "TW48")
           MedAc = False
           
           Case 61
            If CantidadMedusas >= 4 Then
              Call Med_Empieza
              MedAc = True
            Else
              Call SendData(SendTarget.toall, 0, 0, "||Se canceló la batalla de las medusas por falta de participantes." & FONTTYPE_GUERRA)
              Call Med_Cancela
           End If
           
            End Select
        End If
        
   End If
End Sub

Sub CommandGuerra(UserIndex As Integer)

                If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> eNPCType.Banda Then
                     Call SendData(SendTarget.toindex, UserIndex, 0, "Z27")
                    Exit Sub
                End If
    
         If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Banda Then
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > _
                        10 Then
                    Call SendData(SendTarget.toindex, UserIndex, 0, "Z27")
                    Exit Sub

                End If
                
                If BanAc = False Then
                  Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbYellow & "°" & _
                            "¡¡La inscripcion para la guerra de bandas empieza cuando queden 10 minutos." & "!!" & "°" & CStr(Npclist(UserList( _
                            UserIndex).flags.TargetNpc).char.CharIndex))
                 Exit Sub
                End If
                
                If BanEsp = True Then
                   Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbYellow & "°" & _
                            "¡¡La Guerra de Banda ya ha comenzado, te quedaste fuera." & "!!" & "°" & CStr(Npclist(UserList( _
                            UserIndex).flags.TargetNpc).char.CharIndex))
                    Exit Sub
                End If
                
          '  If UserList(UserIndex).flags.Invisible = 1 Then
        '
        '                   Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbCyan & "°" & _
        '                    "¡¡No puedes entrar a Guerra de Banda en invisibilidad" & "!!" & "°" & CStr(Npclist(UserList( _
        '                    UserIndex).flags.TargetNpc).char.CharIndex))
        '        Exit Sub
'
'            End If
      
'            If UserList(UserIndex).flags.Oculto = 1 Then
'                Call SendData(SendTarget.ToIndex, UserIndex, 0, _
'                        "||No puedes ir a eventos estando invisible!." & FONTTYPE_WARNING)
'                Exit Sub
'
'            End If

            If UserList(UserIndex).flags.Montado = True Then
                 Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbYellow & "°" & _
                            "No puedes ir a Guerra de Banda con montura." & "!!" & "°" & CStr(Npclist(UserList( _
                            UserIndex).flags.TargetNpc).char.CharIndex))
                        
                Exit Sub

            End If

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbYellow & "°" & _
                            "Los muertos no pueden entrar en Guerra de Banda." & "!!" & "°" & CStr(Npclist(UserList( _
                            UserIndex).flags.TargetNpc).char.CharIndex))
                            Exit Sub
            End If

            If UserList(UserIndex).flags.EstaDueleando1 = True Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbYellow & "°" & _
                            "No puedes ir a Guerra de Banda estando en duelos." & "!!" & "°" & CStr(Npclist(UserList( _
                            UserIndex).flags.TargetNpc).char.CharIndex))
                            Exit Sub
            End If

            If UserIndex = Team.Pj1 Or UserIndex = Team.Pj2 Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbYellow & "°" & _
                            "No puedes participar en eventos si esperas retos." & "!!" & "°" & CStr(Npclist(UserList( _
                            UserIndex).flags.TargetNpc).char.CharIndex))
                Exit Sub

            End If

            If UserList(UserIndex).Stats.ELV < lvlGuerra Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||" & vbYellow & "°" & _
                            "Debes ser nivel " & lvlGuerra & "o más para entrar a Guerra de Banda." & "!!" & "°" & CStr(Npclist(UserList( _
                            UserIndex).flags.TargetNpc).char.CharIndex))
                Exit Sub
        End If
        
          Call Ban_Entra(UserIndex)
            End If
    
End Sub

Sub Ban_Entra(UserIndex As Integer)

    On Error GoTo errordm:

    Dim i As Integer
 
    For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)

        If (Ban_Luchadores(i) = UserIndex) Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||Ya eres participante de Guerra de Banda!!" & FONTTYPE_INFO)
            Exit Sub

        End If

    Next i
    
    Dim X As Integer
    
    If CantQuitBan > 0 Then
        
        For X = LBound(Quit_Luchadores) To UBound(Quit_Luchadores)
               If ReadField(2, Quit_Luchadores(X), 44) = "Angel" Then
                   
                   Ban_Luchadores(ReadField(1, Quit_Luchadores(X), 44)) = UserIndex
                   Angeles = Angeles + 1
                   
                   UserList(Ban_Luchadores(ReadField(1, Quit_Luchadores(X), 44))).flags.bandas = True
                   UserList(Ban_Luchadores(ReadField(1, Quit_Luchadores(X), 44))).flags.Angel = True
                   Call Transforma(UserIndex)
                   
                Call WarpGuerra(UserIndex, ReadField(3, Quit_Luchadores(X), 44))
                   CantQuitBan = CantQuitBan - 1
                   Quit_Luchadores(X) = ""
                   Exit Sub
               End If
               
               If ReadField(2, Quit_Luchadores(X), 44) = "Demonio" Then
                   
                   Ban_Luchadores(ReadField(1, Quit_Luchadores(X), 44)) = UserIndex
                   Demonios = Demonios + 1
                   
                   UserList(Ban_Luchadores(ReadField(1, Quit_Luchadores(X), 44))).flags.bandas = True
                   UserList(Ban_Luchadores(ReadField(1, Quit_Luchadores(X), 44))).flags.Demonio = True
                   Call Transforma(UserIndex)
                   
                Call WarpGuerra(UserIndex, ReadField(3, Quit_Luchadores(X), 44))
                   CantQuitBan = CantQuitBan - 1
                   Quit_Luchadores(X) = ""
                   Exit Sub
               End If
        Next X
      
      Exit Sub
    End If

    For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)

        If (Ban_Luchadores(i) = -1) Then
            Ban_Luchadores(i) = UserIndex
            UserList(Ban_Luchadores(i)).flags.bandas = True
            CantidadGuerra = CantidadGuerra + 1

            If Demonios < Angeles Then
                ' lo hago q es demonio
                UserList(Ban_Luchadores(i)).flags.Demonio = True
                Demonios = Demonios + 1
                ' convertir en demonio
                Call Transforma(Ban_Luchadores(i))

               Call WarpGuerra(UserIndex, Demonios)
            Else
                     
                UserList(Ban_Luchadores(i)).flags.Angel = True
                Angeles = Angeles + 1
                ' convertir en angel
                Call Transforma(Ban_Luchadores(i))
                        
              Call WarpGuerra(UserIndex, Angeles)
                 
            End If
                 
            Call SendData(SendTarget.toindex, UserIndex, 0, "||Estas dentro de la Guerra!" & FONTTYPE_INFO)
                
            ' Call SendData(SendTarget.toall, 0, 0, "||Guerra AOMania: Entra el participante " & UserList(userindex).name & FONTTYPE_INFO)
                
            If (i = UBound(Ban_Luchadores)) Then
                    
                BanEsp = False
                Call Banauto_Empieza

            End If
              
            Exit Sub

        End If

    Next i

errordm:

End Sub

Sub Banauto_Empieza()

    On Error GoTo errordm
    
    ReDim Preserve Ban_Luchadores(1 To CantidadGuerra) As Integer
 
    Call SendData(SendTarget.ToMap, 0, MapaBan, "||Maten al Rey del otro bando GO, GO, GO ...." & FONTTYPE_GUERRA)
    Call SendData(SendTarget.ToMap, 0, MapaBan, "||Demonios: " & val(Demonios) & " Angeles: " & val(Angeles) & FONTTYPE_GUERRA)
    
    BanEsp = True
    TimerGuerra = 0
    
   Call Reyes_Bandas
    Dim i As Integer

    For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)

        If (Ban_Luchadores(i) <> -1) Then
            If UserList(Ban_Luchadores(i)).flags.Demonio = True Then
                Dim NuevaPos  As WorldPos
                Dim FuturePos As WorldPos
                FuturePos.Map = MapaBan
                FuturePos.X = FortaDemon: FuturePos.Y = FortaDemony
                Call ClosestLegalPos(FuturePos, NuevaPos)

                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), _
                        NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                    

            End If
                    
            If UserList(Ban_Luchadores(i)).flags.Angel = True Then
                Dim NuevaPoss  As WorldPos
                Dim FuturePoss As WorldPos
                FuturePoss.Map = MapaBan
                FuturePoss.X = FortaAngel: FuturePoss.Y = FortaAngely
                Call ClosestLegalPos(FuturePoss, NuevaPoss)

                If NuevaPoss.X <> 0 And NuevaPoss.Y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), _
                        NuevaPoss.Map, NuevaPoss.X, NuevaPoss.Y, True)

            End If

        End If

    Next i

errordm:

End Sub

Sub Destransforma(ByVal UserIndex As Integer)

    On Error GoTo errordm

    If UserList(UserIndex).flags.bandas = True Then

        UserList(UserIndex).char.Body = UserList(UserIndex).CharMimetizado.Body
        UserList(UserIndex).char.Head = UserList(UserIndex).CharMimetizado.Head
        UserList(UserIndex).char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
        UserList(UserIndex).char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
        UserList(UserIndex).char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
        UserList(UserIndex).char.Alas = UserList(UserIndex).CharMimetizado.Alas
        
        UserList(UserIndex).Counters.Mimetismo = 0
        UserList(UserIndex).flags.Mimetizado = 0
      
        Call ChangeUserChar(SendTarget.ToMap, UserIndex, UserList(UserIndex).pos.Map, UserIndex, UserList( _
                UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).char.Heading, _
                UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList( _
                UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)
           
    End If

errordm:

End Sub

Sub Transforma(ByVal UserIndex As Integer)

    On Error GoTo errordm:

    If UserList(UserIndex).flags.Demonio = True Then

        With UserList(UserIndex)
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
                    .char.Heading, .char.WeaponAnim, .char.ShieldAnim, .char.CascoAnim, .char.Alas)

        End With

    End If
    
    If UserList(UserIndex).flags.Angel = True Then

        With UserList(UserIndex)
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
                    .char.Heading, .char.WeaponAnim, .char.ShieldAnim, .char.CascoAnim, .char.Alas)
        
        End With

    End If

errordm:

End Sub

Sub Ban_ReloadTransforma(UserIndex)
  With UserList(UserIndex)
   If UserList(UserIndex).flags.Demonio Then
     .flags.Mimetizado = 1
       
            .char.Body = 348
            .char.Head = 0
            .char.WeaponAnim = 2
            .char.ShieldAnim = 2
            .char.Alas = 0
            
     Call ChangeUserChar(SendTarget.ToMap, 0, .pos.Map, UserIndex, .char.Body, .char.Head, _
                    .char.Heading, .char.WeaponAnim, .char.ShieldAnim, .char.CascoAnim, .char.Alas)
   End If
   
   If UserList(UserIndex).flags.Angel Then
   .flags.Mimetizado = 1
     
            .char.Body = 347
            .char.Head = 0
            .char.WeaponAnim = 2
            .char.ShieldAnim = 2
            .char.Alas = 0
            
            Call ChangeUserChar(SendTarget.ToMap, 0, .pos.Map, UserIndex, .char.Body, .char.Head, _
                    .char.Heading, .char.WeaponAnim, .char.ShieldAnim, .char.CascoAnim, .char.Alas)
   End If
  End With
End Sub

Sub Ban_Comienza(ByVal Giles As Integer)

    On Error GoTo errordm

    If BanAc = True Then
        Call SendData(SendTarget.toindex, 0, 0, "||Ya hay una Guerra de Banda!!" & FONTTYPE_INFO)
        Exit Sub

    End If

    If BanEsp = True Then
        Call SendData(SendTarget.toindex, 0, 0, "||La Guerra de Banda ya ha comenzado!" & FONTTYPE_INFO)
        Exit Sub

    End If

    Cantban = Giles
    
    BanAc = True

    ReDim Ban_Luchadores(1 To Cantban) As Integer
    ReDim Quit_Luchadores(1 To Cantban) As String
    Dim i As Integer

    For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)
        Ban_Luchadores(i) = -1
    Next i

errordm:

End Sub

Sub WarpGuerra(UserIndex As Integer, Cantidad As Integer)

    Dim PosX As Byte
    Dim PosY As Byte
       
       With UserList(UserIndex)
       
         If .flags.Angel = True Then
            
            If Cantidad = 1 Then
              PosX = EsperaAngel
              PosY = EsperaAngely
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
              ElseIf Cantidad = 2 Then
               PosX = EsperaAngel + 1
              PosY = EsperaAngely
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
              ElseIf Cantidad = 3 Then
               PosX = EsperaAngel + 2
              PosY = EsperaAngely
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
            ElseIf Cantidad = 4 Then
               PosX = EsperaAngel + 3
              PosY = EsperaAngely
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
            ElseIf Cantidad = 5 Then
               PosX = EsperaAngel
               PosY = EsperaAngely + 1
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
            ElseIf Cantidad = 6 Then
               PosX = EsperaAngel + 1
              PosY = EsperaAngely + 1
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
            ElseIf Cantidad = 7 Then
               PosX = EsperaAngel + 2
              PosY = EsperaAngely + 1
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
            ElseIf Cantidad = 8 Then
               PosX = EsperaAngel + 3
              PosY = EsperaAngely + 1
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
            ElseIf Cantidad = 9 Then
               PosX = EsperaAngel
              PosY = EsperaAngely + 2
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
            ElseIf Cantidad = 10 Then
               PosX = EsperaAngel + 1
              PosY = EsperaAngely + 2
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
           
           ElseIf Cantidad = 11 Then
               PosX = EsperaAngel + 2
              PosY = EsperaAngely + 2
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
        
        ElseIf Cantidad = 12 Then
               PosX = EsperaAngel + 3
              PosY = EsperaAngely + 2
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
        
        ElseIf Cantidad = 13 Then
               PosX = EsperaAngel
              PosY = EsperaAngely + 3
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
        
        ElseIf Cantidad = 14 Then
               PosX = EsperaAngel + 1
              PosY = EsperaAngely + 3
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)

       
       ElseIf Cantidad = 15 Then
               PosX = EsperaAngel + 2
              PosY = EsperaAngely + 3
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
       
       ElseIf Cantidad = 16 Then
               PosX = EsperaAngel + 3
              PosY = EsperaAngely + 3
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
       End If
       End If
       
       
       If .flags.Demonio = True Then
            
            If Cantidad = 1 Then
              PosX = EsperaDemonio
              PosY = EsperaDemonioy
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
              ElseIf Cantidad = 2 Then
               PosX = EsperaDemonio + 1
              PosY = EsperaDemonioy
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
              ElseIf Cantidad = 3 Then
               PosX = EsperaDemonio + 2
              PosY = EsperaDemonioy
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
            ElseIf Cantidad = 4 Then
               PosX = EsperaDemonio + 3
              PosY = EsperaDemonioy
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
            ElseIf Cantidad = 5 Then
               PosX = EsperaDemonio
               PosY = EsperaDemonioy + 1
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
            ElseIf Cantidad = 6 Then
               PosX = EsperaDemonio + 1
              PosY = EsperaDemonioy + 1
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
            ElseIf Cantidad = 7 Then
               PosX = EsperaDemonio + 2
              PosY = EsperaDemonioy + 1
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
            ElseIf Cantidad = 8 Then
               PosX = EsperaDemonio + 3
              PosY = EsperaDemonioy + 1
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
            ElseIf Cantidad = 9 Then
               PosX = EsperaDemonio
              PosY = EsperaDemonioy + 2
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
                        
            ElseIf Cantidad = 10 Then
               PosX = EsperaDemonio + 1
              PosY = EsperaDemonioy + 2
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
           
           ElseIf Cantidad = 11 Then
               PosX = EsperaDemonio + 2
              PosY = EsperaDemonioy + 2
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
        
        ElseIf Cantidad = 12 Then
               PosX = EsperaDemonio + 3
              PosY = EsperaDemonioy + 2
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
        
        ElseIf Cantidad = 13 Then
               PosX = EsperaDemonio
              PosY = EsperaDemonioy + 3
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
        
        ElseIf Cantidad = 14 Then
               PosX = EsperaDemonio + 1
              PosY = EsperaDemonioy + 3
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)

       
       ElseIf Cantidad = 15 Then
               PosX = EsperaDemonio + 2
              PosY = EsperaDemonioy + 3
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
       
       ElseIf Cantidad = 16 Then
               PosX = EsperaDemonio + 3
              PosY = EsperaDemonioy + 3
              Call WarpUserChar(UserIndex, _
                        MapaBan, PosX, PosY, True)
       End If
       End If
       
       End With
       
       
End Sub

Sub Ban_Desconecta(ByVal UserIndex As Integer)

    On Error GoTo errordm
    
    Dim i As Integer
    Dim Posicion As String
    Dim InfoQuit As String

    
    
    For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)
          
          If Ban_Luchadores(i) = UserIndex Then
              Ban_Luchadores(i) = -1
              InfoQuit = i
          End If
    
    Next i

    If UserList(UserIndex).flags.bandas = True Then
    
    
        If UserList(UserIndex).flags.Demonio = True Then
            
            Posicion = PositionGuerra(UserIndex, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y)
            
            InfoQuit = InfoQuit & "," & Posicion
            
            Demonios = Demonios - 1
            CantQuitBan = CantQuitBan + 1

        End If

        If UserList(UserIndex).flags.Angel = True Then
            Posicion = PositionGuerra(UserIndex, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y)
            
            InfoQuit = InfoQuit & "," & Posicion
            
            Angeles = Angeles - 1
            CantQuitBan = CantQuitBan + 1

        End If

        Call Destransforma(UserIndex)
        UserList(UserIndex).flags.bandas = False
        UserList(UserIndex).flags.Demonio = False
        UserList(UserIndex).flags.Angel = False

        Call WarpUserChar(UserIndex, 34, 30, 50, True)
        
        
        For i = LBound(Quit_Luchadores) To UBound(Quit_Luchadores)
              
              If Quit_Luchadores(i) = "" Then
                   Quit_Luchadores(i) = InfoQuit
                   Exit Sub
              End If
             
        Next i

    End If

errordm:

End Sub

Function PositionGuerra(UserIndex As Integer, PosX As Integer, PosY As Integer)
   
   With UserList(UserIndex)
   
       If .pos.Map = MapaBan Then
          
          If .flags.Angel = True Then
          
           If PosX = EsperaAngel And PosY = EsperaAngely Then
               PositionGuerra = "Angel,1"
               Exit Function
              ElseIf PosX = EsperaAngel + 1 And PosY = EsperaAngely Then
              PositionGuerra = "Angel,2"
              Exit Function
              ElseIf PosX = EsperaAngel + 2 And PosY = EsperaAngely Then
              PositionGuerra = "Angel,3"
              Exit Function
              ElseIf PosX = EsperaAngel + 3 And PosY = EsperaAngely Then
              PositionGuerra = "Angel,4"
              Exit Function
              ElseIf PosX = EsperaAngel And PosY = EsperaAngely + 1 Then
              PositionGuerra = "Angel,5"
              Exit Function
              ElseIf PosX = EsperaAngel + 1 And PosY = EsperaAngely + 1 Then
              PositionGuerra = "Angel,6"
              Exit Function
              ElseIf PosX = EsperaAngel + 2 And PosY = EsperaAngely + 1 Then
              PositionGuerra = "Angel,7"
              Exit Function
              ElseIf PosX = EsperaAngel + 3 And PosY = EsperaAngely + 1 Then
              PositionGuerra = "Angel,8"
              Exit Function
              ElseIf PosX = EsperaAngel And PosY = EsperaAngely + 2 Then
              PositionGuerra = "Angel,9"
              Exit Function
              ElseIf PosX = EsperaAngel + 1 And PosY = EsperaAngely + 2 Then
              PositionGuerra = "Angel,10"
              Exit Function
              ElseIf PosX = EsperaAngel + 2 And PosY = EsperaAngely + 2 Then
              PositionGuerra = "Angel,11"
              Exit Function
              ElseIf PosX = EsperaAngel + 3 And PosY = EsperaAngely + 2 Then
              PositionGuerra = "Angel,12"
              Exit Function
              ElseIf PosX = EsperaAngel And PosY = EsperaAngely + 3 Then
              PositionGuerra = "Angel,13"
              Exit Function
              ElseIf PosX = EsperaAngel + 1 And PosY = EsperaAngely + 3 Then
              PositionGuerra = "Angel,14"
              Exit Function
              ElseIf PosX = EsperaAngel + 2 And PosY = EsperaAngely + 3 Then
              PositionGuerra = "Angel,15"
              Exit Function
              ElseIf PosX = EsperaAngel + 3 And PosY = EsperaAngely + 3 Then
              PositionGuerra = "Angel,16"
              Exit Function
          End If
          
          
           End If
           
          If .flags.Demonio = True Then
          If PosX = EsperaDemonio And PosY = EsperaDemonioy Then
               PositionGuerra = "Demonio,1"
               Exit Function
              ElseIf PosX = EsperaDemonio + 1 And PosY = EsperaDemonioy Then
              PositionGuerra = "Demonio,2"
              Exit Function
              ElseIf PosX = EsperaDemonio + 2 And PosY = EsperaDemonioy Then
              PositionGuerra = "Demonio,3"
              Exit Function
              ElseIf PosX = EsperaDemonio + 3 And PosY = EsperaDemonioy Then
              PositionGuerra = "Demonio,4"
              Exit Function
              ElseIf PosX = EsperaDemonio And PosY = EsperaDemonioy + 1 Then
              PositionGuerra = "Demonio,5"
              Exit Function
              ElseIf PosX = EsperaDemonio + 1 And PosY = EsperaDemonioy + 1 Then
              PositionGuerra = "Demonio,6"
              Exit Function
              ElseIf PosX = EsperaDemonio + 2 And PosY = EsperaDemonioy + 1 Then
              PositionGuerra = "Demonio,7"
              Exit Function
              ElseIf PosX = EsperaDemonio + 3 And PosY = EsperaDemonioy + 1 Then
              PositionGuerra = "Demonio,8"
              Exit Function
              ElseIf PosX = EsperaDemonio And PosY = EsperaDemonioy + 2 Then
              PositionGuerra = "Demonio,9"
              Exit Function
              ElseIf PosX = EsperaDemonio + 1 And PosY = EsperaDemonioy + 2 Then
              PositionGuerra = "Demonio,10"
              Exit Function
              ElseIf PosX = EsperaDemonio + 2 And PosY = EsperaDemonioy + 2 Then
              PositionGuerra = "Demonio,11"
              Exit Function
              ElseIf PosX = EsperaDemonio + 3 And PosY = EsperaDemonioy + 2 Then
              PositionGuerra = "Demonio,12"
              Exit Function
              ElseIf PosX = EsperaDemonio And PosY = EsperaDemonioy + 3 Then
              PositionGuerra = "Demonio,13"
              Exit Function
              ElseIf PosX = EsperaDemonio + 1 And PosY = EsperaDemonioy + 3 Then
              PositionGuerra = "Demonio,14"
              Exit Function
              ElseIf PosX = EsperaDemonio + 2 And PosY = EsperaDemonioy + 3 Then
              PositionGuerra = "Demonio,15"
              Exit Function
              ElseIf PosX = EsperaDemonio + 3 And PosY = EsperaDemonioy + 3 Then
              PositionGuerra = "Demonio,16"
              Exit Function
          End If
          End If
       End If
   
   
   End With
   
   
End Function

''MIRAR AQUI ABAJO
Sub Ban_Muere(ByVal UserIndex As Integer)

    On Error GoTo errord

    If UserList(UserIndex).flags.bandas = True Then
        If UserList(UserIndex).flags.Demonio = True Then
 
            Dim NuevaPosDemon  As WorldPos
            Dim FuturePosDemon As WorldPos
            FuturePosDemon.Map = MapaBan
            FuturePosDemon.X = FortaDemon: FuturePosDemon.Y = FortaDemony
            Call ClosestLegalPos(FuturePosDemon, NuevaPosDemon)

            If NuevaPosDemon.X <> 0 And NuevaPosDemon.Y <> 0 Then Call WarpUserChar(UserIndex, _
                    NuevaPosDemon.Map, NuevaPosDemon.X, NuevaPosDemon.Y, True)

        End If
                    
        If UserList(UserIndex).flags.Angel = True Then
            Dim NuevaPosAngel  As WorldPos
            Dim FuturePosAngel As WorldPos
            FuturePosAngel.Map = MapaBan
            FuturePosAngel.X = FortaAngel: FuturePosAngel.Y = FortaAngely
            Call ClosestLegalPos(FuturePosAngel, NuevaPosAngel)

            If NuevaPosAngel.X <> 0 And NuevaPosAngel.Y <> 0 Then Call WarpUserChar(UserIndex, _
                    NuevaPosAngel.Map, NuevaPosAngel.X, NuevaPosAngel.Y, True)

        End If

    End If

errord:

End Sub


Sub Ban_Cancela()

    On Error GoTo errordm

    If BanAc = False And BanEsp = False Then
        Exit Sub

    End If

    BanEsp = False
    BanAc = False
    
    If CantidadGuerra <> 0 Then
   
        ReDim Preserve Ban_Luchadores(1 To CantidadGuerra) As Integer
    
        Call SendData(SendTarget.toall, 0, 0, _
                "||Se canceló la guerra entre bandas por Game Master" & FONTTYPE_GUERRA)
            
        Dim i As Integer

        For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)

            If (Ban_Luchadores(i) <> -1) Then
                Dim NuevaPos  As WorldPos
                Dim FuturePos As WorldPos
                FuturePos.Map = 34
                FuturePos.X = 30: FuturePos.Y = 50
                Call ClosestLegalPos(FuturePos, NuevaPos)
                    
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), _
                        NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                Call Destransforma(Ban_Luchadores(i))
                UserList(Ban_Luchadores(i)).flags.bandas = False
                UserList(Ban_Luchadores(i)).flags.Demonio = False
                UserList(Ban_Luchadores(i)).flags.Angel = False
                Demonios = 0
                Angeles = 0
                CantidadGuerra = 0
                TimerGuerra = 0
                StatusGuerra = "Medusa"
                Call RespGuerrasDemonio
                Call RespGuerrasAngeles
                 
            End If

        Next i

    Else
        Call SendData(SendTarget.toall, 0, 0, _
                "||Se canceló la guerra entre bandas por Game Master" & FONTTYPE_GUERRA)

    End If

errordm:

End Sub

Sub Banauto_Cancela()

    On Error GoTo errordmm

    If BanAc = False And BanEsp = False Then
        Exit Sub

    End If

    BanEsp = False
    BanAc = False
    
    If CantidadGuerra <> 0 Then
    
        ReDim Preserve Ban_Luchadores(1 To CantidadGuerra) As Integer
        
        Dim i As Integer

        For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)

            If (Ban_Luchadores(i) <> -1) Then
                Dim NuevaPos  As WorldPos
                Dim FuturePos As WorldPos
                FuturePos.Map = 34
                FuturePos.X = 30: FuturePos.Y = 50
                Call ClosestLegalPos(FuturePos, NuevaPos)

                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), _
                        NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                Call Destransforma(Ban_Luchadores(i))
                UserList(Ban_Luchadores(i)).flags.bandas = False
                UserList(Ban_Luchadores(i)).flags.Demonio = False
                UserList(Ban_Luchadores(i)).flags.Angel = False
                InscripcionBanda = False
                Demonios = 0
                Angeles = 0
                CantidadGuerra = 0
                TimerGuerra = 0
                StatusGuerra = "Medusa"
                Call RespGuerrasDemonio
                Call RespGuerrasAngeles
                   
            End If

        Next i

    Else
        Call SendData(SendTarget.toall, 0, 0, _
                "||Se canceló la guerra entre bandas por falta de participantes." & _
                FONTTYPE_GUERRA)

    End If

errordmm:

End Sub

Sub Reyes_Bandas()
    'NPCs antiguos 940 y 941.
    On Error GoTo errordm:

    Dim Npc3    As Integer
    Dim Npc3Pos As WorldPos
    Npc3 = 254
    Npc3Pos.Map = 162
    Npc3Pos.X = 83
    Npc3Pos.Y = 66

    Dim Npc4    As Integer
    Dim Npc4Pos As WorldPos
    Npc4 = 253
    Npc4Pos.Map = 162
    Npc4Pos.X = 18
    Npc4Pos.Y = 36
    Call SpawnNpc(val(Npc3), Npc3Pos, True, False)
    Call SpawnNpc(val(Npc4), Npc4Pos, True, False)
errordm:

End Sub


Sub Ban_Demonios()

    On Error GoTo errordm

    Dim i As Integer

    For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)

        If UserList(Ban_Luchadores(i)).flags.bandas = True Then
    
            If (Ban_Luchadores(i) <> -1) Then
                Dim NuevaPos  As WorldPos
                Dim FuturePos As WorldPos
                FuturePos.Map = 34
                FuturePos.X = 30: FuturePos.Y = 50
                Call ClosestLegalPos(FuturePos, NuevaPos)

                If UserList(Ban_Luchadores(i)).flags.Demonio = True Then
                
                   UserList(Ban_Luchadores(i)).Stats.Exp = UserList(Ban_Luchadores(i)).Stats.Exp + RecBanExp
                   Call CheckUserLevel(Ban_Luchadores(i))
                   Call EnviarExp(Ban_Luchadores(i))
                   Call SendData(SendTarget.toindex, Ban_Luchadores(i), 0, "||Has recibido " & RecBanExp & " de Experencia." & FONTTYPE_FIGHT)
                                
                    UserList(Ban_Luchadores(i)).Stats.GLD = UserList(Ban_Luchadores(i)).Stats.GLD + RecBanOro
                    Call SendUserStatsBox(Ban_Luchadores(i))
                    Call SendData(SendTarget.toindex, Ban_Luchadores(i), 0, "||Has recibido " & RecBanOro & " de Oro." & FONTTYPE_FIGHT)
                

                End If

                If UserList(Ban_Luchadores(i)).flags.bandas = True Then
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), _
                            NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)

                End If

                Call Destransforma(Ban_Luchadores(i))
                UserList(Ban_Luchadores(i)).flags.bandas = False
                UserList(Ban_Luchadores(i)).flags.Demonio = False
                UserList(Ban_Luchadores(i)).flags.Angel = False
                    
                BanAc = False
                BanEsp = False
                InscripcionBanda = False
                Demonios = 0
                Angeles = 0
                CantidadGuerra = 0
                TimerGuerra = 0
                StatusGuerra = "Medusa"

            End If
          
        End If

    Next i

errordm:

End Sub

Sub Ban_Angeles()

    On Error GoTo errordm

    Dim i As Integer

    For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)

        If UserList(Ban_Luchadores(i)).flags.bandas = True Then
  
            If (Ban_Luchadores(i) <> -1) Then
                Dim NuevaPos  As WorldPos
                Dim FuturePos As WorldPos
                FuturePos.Map = 34
                FuturePos.X = 30: FuturePos.Y = 50
                Call ClosestLegalPos(FuturePos, NuevaPos)

                If UserList(Ban_Luchadores(i)).flags.Angel = True Then
                
                   UserList(Ban_Luchadores(i)).Stats.Exp = UserList(Ban_Luchadores(i)).Stats.Exp + RecBanExp
                   Call CheckUserLevel(Ban_Luchadores(i))
                   Call EnviarExp(Ban_Luchadores(i))
                    Call SendData(SendTarget.toindex, Ban_Luchadores(i), 0, "||Has recibido " & RecBanExp & " de Experencia." & FONTTYPE_FIGHT)
                
                    
                    UserList(Ban_Luchadores(i)).Stats.GLD = UserList(Ban_Luchadores(i)).Stats.GLD + RecBanOro
                    Call SendUserStatsBox(Ban_Luchadores(i))
                    Call SendData(SendTarget.toindex, Ban_Luchadores(i), 0, "||Has recibido " & RecBanOro & " de Oro." & FONTTYPE_FIGHT)
                

                End If

                If UserList(Ban_Luchadores(i)).flags.bandas = True Then
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), _
                            NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)

                End If

                Call Destransforma(Ban_Luchadores(i))
                UserList(Ban_Luchadores(i)).flags.bandas = False
                UserList(Ban_Luchadores(i)).flags.Demonio = False
                UserList(Ban_Luchadores(i)).flags.Angel = False
                   
                BanAc = False
                BanEsp = False
                InscripcionBanda = False
                Demonios = 0
                Angeles = 0
                CantidadGuerra = 0
                TimerGuerra = 0
                StatusGuerra = "Medusa"
            End If
         
        End If

    Next i

errordm:

End Sub

Sub PasaTimeBan()
    Dim HpAngeles As Integer
    Dim HpDemonios As Integer
    Dim i As Integer
    
    For i = 1 To NumNPCs
     
      If Npclist(i).pos.Map = MapaBan Then
           
           If Npclist(i).Numero = 253 Then
              HpDemonios = Npclist(i).Stats.MinHP
           End If
           
           If Npclist(i).Numero = 254 Then
               HpAngeles = Npclist(i).Stats.MinHP
           End If
      
      End If
      
    Next i
      
      If HpAngeles > HpDemonios Then
        Call SendData(toall, 0, 0, "||Angeles ganaron la guerra entre bandas, reciben como premio experiencia!!!" & _
        FONTTYPE_GUERRA)
        Call RespGuerrasAngeles
        Call RespGuerrasDemonio
        Call Ban_Angeles
        Exit Sub
      End If
      
      If HpDemonios > HpAngeles Then
        Call SendData(toall, 0, 0, "||Demonios ganaron la guerra entre bandas, reciben como premio experiencia!!!" & _
        FONTTYPE_GUERRA)
        Call RespGuerrasAngeles
        Call RespGuerrasDemonio
        Call Ban_Demonios
        Exit Sub
      End If
     
     Call SendData(toall, 0, 0, "||Angeles y Demonios empataron la guerra de banda." & FONTTYPE_GUERRA)
     Call Banauto_Cancela
End Sub
