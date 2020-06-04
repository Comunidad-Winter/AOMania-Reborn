Attribute VB_Name = "GuerraAOMania"
Option Explicit
Public Totalql            As Integer
Private Cantban           As Integer
Private Const MapaBan     As Integer = 162
Private Const FortaDemon  As Integer = 83
Private Const FortaDemony As Integer = 66
Private Const FortaAngel  As Integer = 18
Private Const FortaAngely As Integer = 36
Private Const EsperaDemonio = 11
Private Const EsperaDemonioy = 89
Private Const EsperaAngel = 87
Private Const EsperaAngely = 9
Public Banac             As Boolean
Public Banesp            As Boolean
Public Bancantidad       As Integer
Private Demonios         As Integer
Private Angeles          As Integer
Public CantidadGuerra    As Integer
Private Ban_Luchadores() As Integer
Public Timer_Banda       As Long

Sub Ban_Entra(ByVal UserIndex)

    On Error GoTo errordm:

    Dim i As Integer

    If Banac = False Then
        Call SendData(SendTarget.ToIndex, 0, 0, "||No hay ninguna Guerra AOMania!" & FONTTYPE_INFO)
        Exit Sub

    End If
 
    If Banesp = False Then
        Call SendData(SendTarget.ToIndex, 0, 0, "||La Guerra AOMania ya ha comenzado, te quedaste fuera!" & FONTTYPE_INFO)
        Exit Sub

    End If
 
    For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)

        If (Ban_Luchadores(i) = UserIndex) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya estas dentro!" & FONTTYPE_WARNING)
            Exit Sub

        End If

    Next i

    For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)

        If (Ban_Luchadores(i) = -1) Then
            Ban_Luchadores(i) = UserIndex
            UserList(Ban_Luchadores(i)).flags.bandas = True
            CantidadGuerra = CantidadGuerra + 1

            If Demonios <= Angeles Then
                ' lo hago q es demonio
                UserList(Ban_Luchadores(i)).flags.demonio = True
                Demonios = Demonios + 1
                ' convertir en demonio
                Call Transforma(Ban_Luchadores(i))
                ' lo teleporto donde los demonios
                Dim NuevaPos  As WorldPos
                Dim FuturePos As WorldPos
                FuturePos.Map = MapaBan
                FuturePos.X = EsperaDemonio
                FuturePos.Y = EsperaDemonioy
                Call ClosestLegalPos(FuturePos, NuevaPos)

                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
           
            Else
                     
                UserList(Ban_Luchadores(i)).flags.angel = True
                Angeles = Angeles + 1
                ' convertir en angel
                Call Transforma(Ban_Luchadores(i))
                ' lo teleporto donde los angeles
                Dim NuevaPoss  As WorldPos
                Dim FuturePoss As WorldPos
                FuturePoss.Map = MapaBan
                FuturePoss.X = EsperaAngel
                FuturePoss.Y = EsperaAngely
                Call ClosestLegalPos(FuturePoss, NuevaPoss)

                If NuevaPoss.X <> 0 And NuevaPoss.Y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), NuevaPoss.Map, NuevaPoss.X, NuevaPoss.Y, True)
                 
            End If
                 
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas dentro de la Guerra!" & FONTTYPE_INFO)
                
            ' Call SendData(SendTarget.toall, 0, 0, "||Guerra AOMania: Entra el participante " & UserList(userindex).name & FONTTYPE_INFO)
                
            If (i = UBound(Ban_Luchadores)) Then
                    
                Banesp = False
                Call Banauto_Empieza

            End If
              
            Exit Sub

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
        
        UserList(UserIndex).Counters.Mimetismo = 0
        UserList(UserIndex).flags.Mimetizado = 0
        '[MaTeO 9]
        Call ChangeUserChar(SendTarget.tomap, UserIndex, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                UserIndex).char.Head, UserList(UserIndex).char.Heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, _
                UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)
        '[/MaTeO 9]
     
    End If

errordm:

End Sub

Sub Transforma(ByVal UserIndex As Integer)

    On Error GoTo errordm:

    If UserList(UserIndex).flags.demonio = True Then

        With UserList(UserIndex)
            .CharMimetizado.Body = .char.Body
            .CharMimetizado.Head = .char.Head
            .CharMimetizado.CascoAnim = .char.CascoAnim
      
            .CharMimetizado.ShieldAnim = .char.ShieldAnim
            .CharMimetizado.WeaponAnim = .char.WeaponAnim
        
            .flags.Mimetizado = 1
        
            'ahora pongo local el del enemigo
            .char.Body = 292

            '[MaTeO 9]
            Call ChangeUserChar(SendTarget.tomap, 0, .pos.Map, UserIndex, .char.Body, .char.Head, .char.Heading, .char.WeaponAnim, _
                    .char.ShieldAnim, .char.CascoAnim, .char.Alas)
            '[/MaTeO 9]

        End With

    End If
    
    If UserList(UserIndex).flags.angel = True Then

        With UserList(UserIndex)
            .CharMimetizado.Body = .char.Body
            .CharMimetizado.Head = .char.Head
            .CharMimetizado.CascoAnim = .char.CascoAnim
      
            .CharMimetizado.ShieldAnim = .char.ShieldAnim
            .CharMimetizado.WeaponAnim = .char.WeaponAnim
        
            .flags.Mimetizado = 1
        
            'ahora pongo local el del enemigo
            .char.Body = 291
       
            '[MaTeO 9]
            Call ChangeUserChar(SendTarget.tomap, 0, .pos.Map, UserIndex, .char.Body, .char.Head, .char.Heading, .char.WeaponAnim, _
                    .char.ShieldAnim, .char.CascoAnim, .char.Alas)
            '[/MaTeO 9]
        
        End With

    End If

errordm:

End Sub

Sub Ban_Comienza(ByVal giles As Integer)

    On Error GoTo errordm

    If Banac = True Then
        Call SendData(SendTarget.ToIndex, 0, 0, "||Ya hay una Guerra AOMania!!" & FONTTYPE_INFO)
        Exit Sub

    End If

    If Banesp = True Then
        Call SendData(SendTarget.ToIndex, 0, 0, "||La Guerra AOMania ya ha comenzado!" & FONTTYPE_INFO)
        Exit Sub

    End If

    Cantban = giles

    Call SendData(SendTarget.toall, 0, 0, "||Se cerraron las inscripciones para la guerra de bandas..." & FONTTYPE_GUERRA)
    Call SendData(SendTarget.toall, 0, 0, "TW48")
    
    Call SendData(SendTarget.tomap, 0, MapaBan, "||Queda 1 minuto para la guerra de bandas, prepárense..." & FONTTYPE_GUERRA)
    
    Banac = True
    Banesp = True
    ReDim Ban_Luchadores(1 To Cantban) As Integer
    Dim i As Integer

    For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)
        Ban_Luchadores(i) = -1
    Next i

errordm:

End Sub

Sub Ban_Muere(ByVal UserIndex As Integer)

    On Error GoTo errord

    If UserList(UserIndex).flags.bandas = True Then
        If UserList(UserIndex).flags.demonio = True Then
 
            Dim NuevaPosDemon  As WorldPos
            Dim FuturePosDemon As WorldPos
            FuturePosDemon.Map = MapaBan
            FuturePosDemon.X = FortaDemon
            FuturePosDemon.Y = FortaDemony
            Call ClosestLegalPos(FuturePosDemon, NuevaPosDemon)

            If NuevaPosDemon.X <> 0 And NuevaPosDemon.Y <> 0 Then Call WarpUserChar(UserIndex, NuevaPosDemon.Map, NuevaPosDemon.X, NuevaPosDemon.Y, _
                    True)

        End If
                    
        If UserList(UserIndex).flags.angel = True Then
            Dim NuevaPosAngel  As WorldPos
            Dim FuturePosAngel As WorldPos
            FuturePosAngel.Map = MapaBan
            FuturePosAngel.X = FortaAngel
            FuturePosAngel.Y = FortaAngely
            Call ClosestLegalPos(FuturePosAngel, NuevaPosAngel)

            If NuevaPosAngel.X <> 0 And NuevaPosAngel.Y <> 0 Then Call WarpUserChar(UserIndex, NuevaPosAngel.Map, NuevaPosAngel.X, NuevaPosAngel.Y, _
                    True)

        End If

    End If

errord:

End Sub

Sub Ban_Desconecta(ByVal UserIndex As Integer)

    On Error GoTo errordm

    If UserList(UserIndex).flags.bandas = True Then
        If UserList(UserIndex).flags.demonio = True Then
            Demonios = Demonios - 1

        End If

        If UserList(UserIndex).flags.angel = True Then
            Angeles = Angeles - 1

        End If

        Call Destransforma(UserIndex)
        UserList(UserIndex).flags.bandas = False
        UserList(UserIndex).flags.demonio = False
        UserList(UserIndex).flags.angel = False

        Call WarpUserChar(UserIndex, 34, 30, 50, True)

    End If

errordm:

End Sub

Sub Ban_Cancela()

    On Error GoTo errordm

    If Banac = False And Banesp = False Then
        Exit Sub

    End If

    Banesp = False
    Banac = False
    
    If CantidadGuerra <> 0 Then
   
        ReDim Preserve Ban_Luchadores(1 To CantidadGuerra) As Integer
    
        Call SendData(SendTarget.toall, 0, 0, "||Se canceló la guerra entre bandas por Game Master" & FONTTYPE_GUERRA)
            
        Dim i As Integer

        For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)

            If (Ban_Luchadores(i) <> -1) Then
                Dim NuevaPos  As WorldPos
                Dim FuturePos As WorldPos
                FuturePos.Map = 1
                FuturePos.X = 50
                FuturePos.Y = 50
                Call ClosestLegalPos(FuturePos, NuevaPos)
                    
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                Call Destransforma(Ban_Luchadores(i))
                UserList(Ban_Luchadores(i)).flags.bandas = False
                UserList(Ban_Luchadores(i)).flags.demonio = False
                UserList(Ban_Luchadores(i)).flags.angel = False
                Demonios = 0
                Angeles = 0
                CantidadGuerra = 0
                Call RespGuerrasDemonio
                Call RespGuerrasAngeles
                 
            End If

        Next i

    Else
        Call SendData(SendTarget.toall, 0, 0, "||Se canceló la guerra entre bandas por Game Master" & FONTTYPE_GUERRA)

    End If

errordm:

End Sub

Sub Banauto_Cancela()

    On Error GoTo errordmm

    If Banac = False And Banesp = False Then
        Exit Sub

    End If

    Banesp = False
    Banac = False
    
    If CantidadGuerra <> 0 Then
    
        ReDim Preserve Ban_Luchadores(1 To CantidadGuerra) As Integer
        
        Call SendData(SendTarget.toall, 0, 0, "||Se canceló la guerra entre bandas por falta de participantes." & FONTTYPE_GUERRA)
        
        Dim i As Integer

        For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)

            If (Ban_Luchadores(i) <> -1) Then
                Dim NuevaPos  As WorldPos
                Dim FuturePos As WorldPos
                FuturePos.Map = 1
                FuturePos.X = 50
                FuturePos.Y = 50
                Call ClosestLegalPos(FuturePos, NuevaPos)

                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                Call Destransforma(Ban_Luchadores(i))
                UserList(Ban_Luchadores(i)).flags.bandas = False
                UserList(Ban_Luchadores(i)).flags.demonio = False
                UserList(Ban_Luchadores(i)).flags.angel = False
                Demonios = 0
                Angeles = 0
                CantidadGuerra = 0
                Call RespGuerrasDemonio
                Call RespGuerrasAngeles
                    
            End If

        Next i

    Else
        Call SendData(SendTarget.toall, 0, 0, "||Se canceló la guerra entre bandas por falta de participantes." & FONTTYPE_GUERRA)

    End If

errordmm:

End Sub

Sub Reyes_Bandas()

    On Error GoTo errordm:

    Dim Npc3    As Integer
    Dim Npc3Pos As WorldPos
    Npc3 = 940
    Npc3Pos.Map = 66
    Npc3Pos.X = 77
    Npc3Pos.Y = 23

    Dim Npc4    As Integer
    Dim Npc4Pos As WorldPos
    Npc4 = 941
    Npc4Pos.Map = 66
    Npc4Pos.X = 77
    Npc4Pos.Y = 77
    Call SpawnNpc(val(Npc3), Npc3Pos, True, False)
    Call SpawnNpc(val(Npc4), Npc4Pos, True, False)
errordm:

End Sub

Sub Banauto_Empieza()

    On Error GoTo errordm

    Banesp = False
    
    ReDim Preserve Ban_Luchadores(1 To CantidadGuerra) As Integer
 
    Call SendData(SendTarget.tomap, 0, MapaBan, "||Maten al Rey del otro bando GO, GO, GO ...." & FONTTYPE_GUERRA)
    Call SendData(SendTarget.tomap, 0, MapaBan, "||Demonios: 2 Angeles: 2" & FONTTYPE_GUERRA)
    
    Call Reyes_Bandas
    Dim i As Integer

    For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)

        If (Ban_Luchadores(i) <> -1) Then
            If UserList(Ban_Luchadores(i)).flags.demonio = True Then
                Dim NuevaPos  As WorldPos
                Dim FuturePos As WorldPos
                FuturePos.Map = MapaBan
                FuturePos.X = FortaDemon
                FuturePos.Y = FortaDemony
                Call ClosestLegalPos(FuturePos, NuevaPos)

                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)

            End If
                    
            If UserList(Ban_Luchadores(i)).flags.angel = True Then
                Dim NuevaPoss  As WorldPos
                Dim FuturePoss As WorldPos
                FuturePoss.Map = MapaBan
                FuturePoss.X = FortaAngel
                FuturePoss.Y = FortaAngely
                Call ClosestLegalPos(FuturePoss, NuevaPoss)

                If NuevaPoss.X <> 0 And NuevaPoss.Y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), NuevaPoss.Map, NuevaPoss.X, NuevaPoss.Y, True)

            End If

        End If

    Next i

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
                FuturePos.Map = 1
                FuturePos.X = 50
                FuturePos.Y = 50
                Call ClosestLegalPos(FuturePos, NuevaPos)

                If UserList(Ban_Luchadores(i)).flags.demonio = True Then
                    UserList(Ban_Luchadores(i)).Stats.GLD = UserList(Ban_Luchadores(i)).Stats.GLD + 1000000
                    Call SendUserStatsBox(Ban_Luchadores(i))

                End If

                If UserList(Ban_Luchadores(i)).flags.bandas = True Then
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)

                End If

                Call Destransforma(Ban_Luchadores(i))
                UserList(Ban_Luchadores(i)).flags.bandas = False
                UserList(Ban_Luchadores(i)).flags.demonio = False
                UserList(Ban_Luchadores(i)).flags.angel = False
                    
                Banac = False
                Banesp = False
                Demonios = 0
                Angeles = 0
                CantidadGuerra = 0

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
                FuturePos.Map = 1
                FuturePos.X = 50
                FuturePos.Y = 50
                Call ClosestLegalPos(FuturePos, NuevaPos)

                If UserList(Ban_Luchadores(i)).flags.angel = True Then
                    UserList(Ban_Luchadores(i)).Stats.GLD = UserList(Ban_Luchadores(i)).Stats.GLD + 1000000
                    Call SendUserStatsBox(Ban_Luchadores(i))

                End If

                If UserList(Ban_Luchadores(i)).flags.bandas = True Then
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)

                End If

                Call Destransforma(Ban_Luchadores(i))
                UserList(Ban_Luchadores(i)).flags.bandas = False
                UserList(Ban_Luchadores(i)).flags.demonio = False
                UserList(Ban_Luchadores(i)).flags.angel = False
                   
                Banac = False
                Banesp = False
                Demonios = 0
                Angeles = 0
                CantidadGuerra = 0

            End If
         
        End If

    Next i

errordm:

End Sub

Sub Timer_GuerradeBanda()
   
    Timer_Banda = Timer_Banda + 1

    Select Case Timer_Banda
    
        Case 2
            Call Ban_Comienza(32)

        Case 50
            Call SendData(SendTarget.toall, 0, 0, "||Quedan 10  minutos para la proxima guerra de bandas (no se pierde inventario)" & FONTTYPE_GUILD)

        Case 55
            Call SendData(SendTarget.toall, 0, 0, "||Quedan 5 minutos para la proxima guerra de bandas." & FONTTYPE_GUILD)

        Case 56
            Call SendData(SendTarget.toall, 0, 0, "||Quedan 4 minutos para la proxima guerra de bandas." & FONTTYPE_GUILD)

        Case 57
            Call SendData(SendTarget.toall, 0, 0, "||Quedan 3 minutos para la proxima guerra de bandas." & FONTTYPE_GUILD)
        
        Case 58
            Call SendData(SendTarget.toall, 0, 0, "||Quedan 2 minutos para la proxima guerra de bandas." & FONTTYPE_GUILD)
        
        Case 59
            Call SendData(SendTarget.toall, 0, 0, "||Quedan 1 minutos para la proxima guerra de bandas." & FONTTYPE_GUILD)

        Case 60
            Call Ban_Comienza(32)

        Case 49

            If Banac = True Then
                If CantidadGuerra < 6 Then
                    Call Banauto_Cancela
                    bandasqls = 1
                Else

                    If Banesp = True Then
                        Call Banauto_Empieza
                        bandasqls = 1
                    Else
                        bandasqls = 1

                    End If

                End If

            End If

    End Select
   
End Sub
