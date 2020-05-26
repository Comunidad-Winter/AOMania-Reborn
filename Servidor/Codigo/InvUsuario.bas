Attribute VB_Name = "InvUsuario"

Option Explicit

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean

    On Error Resume Next

    Dim i        As Integer

    Dim ObjIndex As Integer

    For i = 1 To MAX_INVENTORY_SLOTS
        ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex

        If ObjIndex > 0 Then
            If (ObjData(ObjIndex).ObjType <> eOBJType.otLlaves) And ObjData(ObjIndex).Cae <> 1 And ObjData(ObjIndex).Templ <> 1 And ObjData(ObjIndex).Nemes <> 1 And ObjData(ObjIndex).Real <> 1 And (ObjData(ObjIndex).ObjType <> eOBJType.otMontura) And ObjData(ObjIndex).Caos <> 1 And ObjData(ObjIndex).ObjType <> eOBJType.otalas And ObjData(ObjIndex).NoRobable <> 1 Then
                
                If UserList(UserIndex).Invent.EscudoEqpObjIndex <> ObjIndex And UserList(UserIndex).Invent.AlaEqpObjIndex <> ObjIndex And UserList(UserIndex).Invent.AmuletoEqpObjIndex <> ObjIndex And UserList(UserIndex).Invent.ArmourEqpObjIndex <> ObjIndex And UserList(UserIndex).Invent.BarcoObjIndex <> ObjIndex And UserList(UserIndex).Invent.CascoEqpObjIndex <> ObjIndex And UserList(UserIndex).Invent.HerramientaEqpObjIndex <> ObjIndex And UserList(UserIndex).Invent.WeaponEqpObjIndex <> ObjIndex Then
                   
                    TieneObjetosRobables = True
                    Exit Function
                
                End If

            End If
    
        End If

    Next i

End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean

    On Error GoTo manejador

    'Call LogTarea("ClasePuedeUsarItem")

    Dim flag As Boolean


    Dim i As Integer

    For i = 1 To NUMCLASES

        If ObjData(ObjIndex).ClaseProhibida(i) <> "" Then

            If ObjData(ObjIndex).ClaseProhibida(i) = UCase$(UserList(UserIndex).Clase) Then
                ClasePuedeUsarItem = False
                Exit Function
            End If

        Else

        End If

    Next i

    ClasePuedeUsarItem = True

    Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")

End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
    Dim j As Integer

    For j = 1 To MAX_INVENTORY_SLOTS

        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then

            If ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
            Call UpdateUserInv(False, UserIndex, j)

        End If

    Next j

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)

    Dim j As Integer

    For j = 1 To MAX_INVENTORY_SLOTS
        UserList(UserIndex).Invent.Object(j).ObjIndex = 0
        UserList(UserIndex).Invent.Object(j).Amount = 0
        UserList(UserIndex).Invent.Object(j).Equipped = 0

    Next

    UserList(UserIndex).Invent.NroItems = 0

    UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
    UserList(UserIndex).Invent.ArmourEqpSlot = 0

    UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
    UserList(UserIndex).Invent.WeaponEqpSlot = 0

    UserList(UserIndex).Invent.CascoEqpObjIndex = 0
    UserList(UserIndex).Invent.CascoEqpSlot = 0

    UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
    UserList(UserIndex).Invent.EscudoEqpSlot = 0

    UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
    UserList(UserIndex).Invent.HerramientaEqpSlot = 0

    UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
    UserList(UserIndex).Invent.MunicionEqpSlot = 0

    UserList(UserIndex).Invent.BarcoObjIndex = 0
    UserList(UserIndex).Invent.BarcoSlot = 0

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)

    On Error GoTo errhandler

    'If UserList(UserIndex).flags.Privilegios = 1 Or UserList(UserIndex).flags.Privilegios = 2 Then Exit Sub
    If Cantidad > 100000 Or Cantidad < 1000 Then Exit Sub

    'SI EL NPC TIENE ORO LO TIRAMOS
    If (Cantidad > 0) And (Cantidad <= UserList(UserIndex).Stats.GLD) Then

        Dim i     As Byte

        Dim MiObj As Obj
        
        'info debug
        Dim loops As Integer

        'Seguridad Alkon
        If Cantidad > 49999 Then

            Dim j        As Integer

            Dim k        As Integer

            Dim m        As Integer

            Dim Cercanos As String

            For j = UserList(UserIndex).pos.X - 5 To UserList(UserIndex).pos.X + 5
                For k = UserList(UserIndex).pos.Y - 5 To UserList(UserIndex).pos.Y + 5

                    If LegalPos(m, j, k, True) Then
                        If MapData(m, j, k).UserIndex > 0 Then
                            Cercanos = Cercanos & UserList(MapData(m, j, k).UserIndex).Name & ","

                        End If

                    End If

                Next k
            Next j

            Call LogDesarrollo(UserList(UserIndex).Name & " tira oro. Cercanos: " & Cercanos)

        End If

        '/Seguridad
        Do While (Cantidad > 0) And (UserList(UserIndex).Stats.GLD > 0)
            
            If Cantidad > MAX_INVENTORY_OBJS And UserList(UserIndex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.Amount = MAX_INVENTORY_OBJS
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.Amount
            Else
                MiObj.Amount = Cantidad
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Cantidad
                Cantidad = Cantidad - MiObj.Amount

            End If

            MiObj.ObjIndex = iORO
            
            If UserList(UserIndex).flags.Privilegios > 0 Then Call LogGM(UserList(UserIndex).Name, "Tiro cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
            
            Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)

            '[cristiaen]
            'ARREGLO DE BUG DE CLONAR OBJETOS
            Dim UserFile As String

            UserFile = CharPath & UCase$(UserList(UserIndex).Name) & ".chr"
            Call WriteVar(UserFile, "STATS", "GLD", str(UserList(UserIndex).Stats.GLD))
            '[/cristiaen]
            'info debug
            loops = loops + 1

            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub

            End If
            
        Loop
    
    End If

    Exit Sub

errhandler:

End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)

    Dim MiObj As Obj

    'Desequipa
    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

    If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)

    'Quita un objeto
    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - Cantidad

    '¿Quedan mas?
    If UserList(UserIndex).Invent.Object(Slot).Amount <= 0 Then
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
        UserList(UserIndex).Invent.Object(Slot).Amount = 0

    End If

End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
On Error Resume Next
Dim NullObj As UserOBJ
Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Invent.Object(Slot).ObjIndex > 0 Then
        Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).Invent.Object(Slot))
    Else
        Call ChangeUserInv(UserIndex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_INVENTORY_SLOTS

        'Actualiza el inventario
        If UserList(UserIndex).Invent.Object(LoopC).ObjIndex > 0 Then
            Call ChangeUserInv(UserIndex, LoopC, UserList(UserIndex).Invent.Object(LoopC))
        Else
            
            Call ChangeUserInv(UserIndex, LoopC, NullObj)
            
        End If

    Next LoopC

End If

End Sub

Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

    Dim Obj As Obj

    If num > 0 Then

        If num > UserList(UserIndex).Invent.Object(Slot).Amount Then num = UserList(UserIndex).Invent.Object(Slot).Amount

        'Check objeto en el suelo
        If MapData(UserList(UserIndex).pos.Map, X, Y).OBJInfo.ObjIndex = 0 Or MapData(UserList(UserIndex).pos.Map, X, Y).OBJInfo.ObjIndex = _
           UserList(UserIndex).Invent.Object(Slot).ObjIndex Then

            Obj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex

            If UserList(UserIndex).flags.Privilegios <= PlayerType.Consejero Then
                If ObjData(Obj.ObjIndex).Newbie = 1 And EsNewbie(UserIndex) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes tirar el objeto." & FONTTYPE_INFO)
                    Exit Sub
                End If

                If ObjData(Obj.ObjIndex).Caos = 1 Or ObjData(Obj.ObjIndex).Real = 1 Or ObjData(Obj.ObjIndex).Templ = 1 Or ObjData(Obj.ObjIndex).Nemes = 1 Then

                    If ObjData(Obj.ObjIndex).ObjType = eOBJType.otArmadura Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes tirar el objeto." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
            End If

            If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)

            If num + MapData(UserList(UserIndex).pos.Map, X, Y).OBJInfo.Amount > MAX_INVENTORY_OBJS Then
                num = MAX_INVENTORY_OBJS - MapData(UserList(UserIndex).pos.Map, X, Y).OBJInfo.Amount

            End If

            Obj.Amount = num

            Call MakeObj(SendTarget.ToMap, 0, Map, Obj, Map, X, Y)
            Call QuitarUserInvItem(UserIndex, Slot, num)
            Call UpdateUserInv(False, UserIndex, Slot)


            If ObjData(Obj.ObjIndex).ObjType = eOBJType.otMontura And UserList(UserIndex).flags.Montado = True Then
                UserList(UserIndex).char.Body = UserList(UserIndex).flags.NumeroMont

                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                                UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                             UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

                UserList(UserIndex).flags.Montado = False

            End If

            If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                Call LogGM(UserList(UserIndex).Name, " tiró " & num & " " & ObjData(Obj.ObjIndex).Name)
            End If

            If MapData(UserList(UserIndex).pos.Map, X, Y).OBJInfo.ObjIndex = 1 Or MapData(UserList(UserIndex).pos.Map, X, Y).OBJInfo.ObjIndex = _
               UserList(UserIndex).Invent.Object(Slot).ObjIndex Then

            End If

        Else

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay espacio en el piso." & FONTTYPE_INFO)


        End If


    End If

End Sub

Sub EraseObj(ByVal sndRoute As Byte, _
             ByVal sndIndex As Integer, _
             ByVal sndMap As Integer, _
             ByVal num As Integer, _
             ByVal Map As Byte, _
             ByVal X As Integer, _
             ByVal Y As Integer)

    MapData(Map, X, Y).OBJInfo.Amount = MapData(Map, X, Y).OBJInfo.Amount - num

    If MapData(Map, X, Y).OBJInfo.Amount <= 0 Then
        MapData(Map, X, Y).OBJInfo.ObjIndex = 0
        MapData(Map, X, Y).OBJInfo.Amount = 0

        If sndRoute = SendTarget.ToMap Then
            Call SendToAreaByPos(Map, X, Y, "BO" & X & "," & Y)
        Else
            Call SendData(sndRoute, sndIndex, sndMap, "BO" & X & "," & Y)

        End If

    End If

End Sub

Sub MakeObj(ByVal sndRoute As Byte, _
            ByVal sndIndex As Integer, _
            ByVal sndMap As Integer, _
            Obj As Obj, _
            Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer)

    On Error Resume Next

    'Crea un Objeto

    If MapData(Map, X, Y).OBJInfo.ObjIndex = Obj.ObjIndex Then
        MapData(Map, X, Y).OBJInfo.Amount = MapData(Map, X, Y).OBJInfo.Amount + Obj.Amount
        Obj.Amount = MapData(Map, X, Y).OBJInfo.Amount
        MapData(Map, X, Y).OBJInfo = Obj
        Call SendData(sndRoute, sndIndex, sndMap, "HO" & ObjData(Obj.ObjIndex).GrhIndex & "," & X & "," & Y)
    Else
        MapData(Map, X, Y).OBJInfo = Obj
        Call SendData(sndRoute, sndIndex, sndMap, "HO" & ObjData(Obj.ObjIndex).GrhIndex & "," & X & "," & Y)
    End If

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj) As Boolean

    On Error GoTo errhandler

    'Call LogTarea("MeterItemEnInventario")

    Dim X As Integer
    Dim Y As Integer
    Dim Slot As Byte
    
    '¿el user ya tiene un objeto del mismo tipo?
    Slot = 1

    Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= _
       MAX_INVENTORY_OBJS
        Slot = Slot + 1

        If Slot > MAX_INVENTORY_SLOTS Then
            Exit Do

        End If

    Loop

    'Sino busca un slot vacio
    If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1

        Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z24")
                MeterItemEnInventario = False
                Exit Function

            End If

        Loop
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1

    End If

    'Mete el objeto
    If UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
        'Menor que MAX_INV_OBJS
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
        UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount
    Else
        UserList(UserIndex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS

    End If

    MeterItemEnInventario = True

    Call UpdateUserInv(False, UserIndex, Slot)
    
    If UserList(UserIndex).Quest.Start = 1 Then
        If UserList(UserIndex).Quest.NumObj > 0 Then
            Call BuscaObjQuest(UserIndex, MiObj.ObjIndex, MiObj.Amount, UserList(UserIndex).Quest.Quest)
        ElseIf UserList(UserIndex).Quest.NumObjNpc > 0 Then
             Call BuscaObjNpcQuest(UserIndex, MiObj.ObjIndex, MiObj.Amount, UserList(UserIndex).Quest.Quest)
        End If
    End If
    
    Exit Function
errhandler:

End Function

Sub GetObj(ByVal UserIndex As Integer)

    Dim Obj As ObjData
    Dim MiObj As Obj

    '¿Hay algun obj?
    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).OBJInfo.ObjIndex > 0 Then

        '¿Esta permitido agarrar este obj?
        If ObjData(MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).OBJInfo.ObjIndex).Agarrable <> 1 Then
            Dim X As Integer
            Dim Y As Integer
            Dim Slot As Byte

            X = UserList(UserIndex).pos.X
            Y = UserList(UserIndex).pos.Y
            Obj = ObjData(MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).OBJInfo.ObjIndex)
            MiObj.Amount = MapData(UserList(UserIndex).pos.Map, X, Y).OBJInfo.Amount
            MiObj.ObjIndex = MapData(UserList(UserIndex).pos.Map, X, Y).OBJInfo.ObjIndex

            If Obj.ObjType = otGuita Then
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MiObj.Amount
                Call EraseObj(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, MapData(UserList(UserIndex).pos.Map, X, Y).OBJInfo.Amount, UserList( _
                                                                                                                                           UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y)
                Call SendUserStatsBox(UserIndex)
                Exit Sub

            End If

            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z24")
            Else
                'Quitamos el objeto
                Call EraseObj(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, MapData(UserList(UserIndex).pos.Map, X, Y).OBJInfo.Amount, UserList( _
                                                                                                                                           UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y)

                If UserList(UserIndex).flags.Privilegios > PlayerType.User Then Call LogGM(UserList(UserIndex).Name, "Agarro:" & MiObj.Amount & _
                                                                                                                   " Objeto:" & ObjData(MiObj.ObjIndex).Name)

            End If

        End If

    Else

    End If

End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
'Desequipa el item slot del inventario
    Dim Obj As ObjData

    If (Slot < LBound(UserList(UserIndex).Invent.Object)) Or (Slot > UBound(UserList(UserIndex).Invent.Object)) Then
        Exit Sub
    ElseIf UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0 Then
        Exit Sub

    End If

    Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)

    Select Case Obj.ObjType

    Case eOBJType.otAmuletoDefensa
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.AmuletoEqpObjIndex = 0
        UserList(UserIndex).Invent.AmuletoEqpSlot = 0

    Case eOBJType.otWeapon

        If EspadaSagrada.EspadaSagrada(UserList(UserIndex).Invent.WeaponEqpObjIndex) Then
            Call DeleteSagradaHit(UserIndex)

        End If

        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
        UserList(UserIndex).Invent.WeaponEqpSlot = 0

        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).char.WeaponAnim = NingunArma
            '[MaTeO 9]
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                            UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                         UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

            '[/MaTeO 9]
        End If

        If Obj.ObjetoEspecial > 0 Then
            Call QuitarObjetoEspecial(UserIndex, Obj.ObjetoEspecial)

        End If

    Case eOBJType.otFlechas
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
        UserList(UserIndex).Invent.MunicionEqpSlot = 0

        If Obj.ObjetoEspecial > 0 Then
            Call QuitarObjetoEspecial(UserIndex, Obj.ObjetoEspecial)

        End If

    Case eOBJType.otherramientas
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
        UserList(UserIndex).Invent.HerramientaEqpSlot = 0

    Case eOBJType.otArmadura
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
        UserList(UserIndex).Invent.ArmourEqpSlot = 0
        Call DarCuerpoDesnudo(UserIndex, UserList(UserIndex).flags.Mimetizado = 1)
        '[MaTeO 9]
        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                        UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                     UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

        '[/MaTeO 9]
        If Obj.ObjetoEspecial > 0 Then
            Call QuitarObjetoEspecial(UserIndex, Obj.ObjetoEspecial)

        End If

    Case eOBJType.otCASCO
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.CascoEqpObjIndex = 0
        UserList(UserIndex).Invent.CascoEqpSlot = 0

        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).char.CascoAnim = NingunCasco
            '[MaTeO 9]
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                            UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                         UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

            '[/MaTeO 9]
        End If

        If Obj.ObjetoEspecial > 0 Then
            Call QuitarObjetoEspecial(UserIndex, Obj.ObjetoEspecial)

        End If

    Case eOBJType.otESCUDO
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
        UserList(UserIndex).Invent.EscudoEqpSlot = 0

        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).char.ShieldAnim = NingunEscudo
            '[MaTeO 9]
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                            UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                         UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

            '[/MaTeO 9]
        End If

        If Obj.ObjetoEspecial > 0 Then
            Call QuitarObjetoEspecial(UserIndex, Obj.ObjetoEspecial)

        End If

        '[MaTeO 9]
    Case eOBJType.otalas
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.AlaEqpObjIndex = 0
        UserList(UserIndex).Invent.AlaEqpSlot = 0

        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).char.Alas = NingunAlas
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                            UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                         UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

        End If

        '[/MaTeO 9]

    End Select

    Call EnviarSta(UserIndex)
    Call UpdateUserInv(False, UserIndex, Slot)
    Call SendUserHitBox(UserIndex)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean

    On Error GoTo errhandler

    If ObjData(ObjIndex).Mujer = 1 Then
        SexoPuedeUsarItem = UCase$(UserList(UserIndex).Genero) <> "HOMBRE"
    ElseIf ObjData(ObjIndex).Hombre = 1 Then
        SexoPuedeUsarItem = UCase$(UserList(UserIndex).Genero) <> "MUJER"
    Else
        SexoPuedeUsarItem = True

    End If

    Exit Function
errhandler:
    Call LogError("SexoPuedeUsarItem")

End Function

Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error Resume Next
If ObjData(ObjIndex).Real = 1 Then
      FaccionPuedeUsarItem = (UserList(UserIndex).Faccion.ArmadaReal = 1)
      Exit Function
      
ElseIf ObjData(ObjIndex).Caos = 1 Then
    FaccionPuedeUsarItem = (UserList(UserIndex).Faccion.FuerzasCaos = 1)
    Exit Function
ElseIf ObjData(ObjIndex).Templ = 1 Then
   FaccionPuedeUsarItem = (UserList(UserIndex).Faccion.Templario = 1)
   Exit Function
ElseIf ObjData(ObjIndex).Nemes = 1 Then
    FaccionPuedeUsarItem = (UserList(UserIndex).Faccion.Nemesis = 1)
    Exit Function
 Else
    FaccionPuedeUsarItem = True
End If

End Function


Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

    On Error GoTo errhandler

    'Equipa un item del inventario
    Dim Obj As ObjData
    Dim ObjIndex As Integer

    ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    Obj = ObjData(ObjIndex)

    If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo los newbies pueden usar este objeto." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UserList(UserIndex).flags.Privilegios = PlayerType.User Then

        If Obj.Gm = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Este item solo puede ser utilizado por un GameMaster." & FONTTYPE_INFO)
            Exit Sub
        End If

        If Obj.RazaEnana > 0 Then
            Select Case UCase$(UserList(UserIndex).Raza)
            Case "ENANO"
            Case "GNOMO"
            Case Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clase o raza no puede usar este objeto." & FONTTYPE_INFO)
                Exit Sub
            End Select
        End If

        If Obj.RazaHobbit > 0 Then
            If UCase$(UserList(UserIndex).Raza) <> "HOBBIT" Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clase o raza no puede usar este objeto." & FONTTYPE_INFO)
                Exit Sub
            End If
        End If

        If Obj.RazaVampiro > 0 Then
            If UCase$(UserList(UserIndex).Raza) <> "VAMPIRO" Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clase o raza no puede usar este objeto." & FONTTYPE_INFO)
                Exit Sub
            End If
        End If

        If Obj.RazaOrco > 0 Then
            If UCase$(UserList(UserIndex).Raza) <> "ORCO" Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clase o raza no puede usar este objeto." & FONTTYPE_INFO)
                Exit Sub
            End If
        End If

        If Obj.RazaEnana = 0 And Obj.RazaOrco = 0 And Obj.RazaVampiro = 0 And Obj.RazaHobbit = 0 Then
            Select Case Obj.ObjType

            Case eOBJType.otArmadura
                If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Or UCase$(UserList(UserIndex).Raza) = "HOBBIT" Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clase o raza no puede usar este objeto." & FONTTYPE_INFO)
                    Exit Sub
                End If
            End Select
        End If

        If ClasePuedeUsarItem(UserIndex, ObjIndex) = False Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clase o raza no puede usar este objeto." & FONTTYPE_INFO)
            Exit Sub
        End If

        If SexoPuedeUsarItem(UserIndex, ObjIndex) = False Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu sexo no puede usar este objeto." & FONTTYPE_INFO)
            Exit Sub
        End If


        If ObjData(ObjIndex).Real = 1 Or ObjData(ObjIndex).Caos = 1 Or ObjData(ObjIndex).Nemes = 1 Or ObjData(ObjIndex).Templ = 1 Then
            If Not FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ese item es para miembros de otra facción." & FONTTYPE_INFO)
                Exit Sub
            End If
        End If
    End If

    If Obj.Nivel > 0 Then
        If UserList(UserIndex).Stats.ELV < Obj.Nivel Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes utilizar este item. Para hacerlo debes ser nivel " & Obj.Nivel & " o superior." & FONTTYPE_INFO)
            Exit Sub
        End If
    End If

    Select Case Obj.ObjType

        '[MaTeO 9]
    Case eOBJType.otalas

        If ClasePuedeUsarItem(UserIndex, ObjIndex) Or FaccionPuedeUsarItem(UserIndex, ObjIndex) Or UserList(UserIndex).flags.Privilegios >= _
           PlayerType.Dios Then

            'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                'Quitamos del inv el item
                Call Desequipar(UserIndex, Slot)

                'Animacion por defecto
                If UserList(UserIndex).flags.Mimetizado = 1 Then
                    UserList(UserIndex).CharMimetizado.WeaponAnim = NingunAlas
                Else
                    UserList(UserIndex).char.WeaponAnim = NingunAlas
                    '[MaTeO 9]
                    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                                    UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                                 UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

                    '[/MaTeO 9]
                End If

                Exit Sub

            End If

            'Quitamos el elemento anterior
            If UserList(UserIndex).Invent.AlaEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.AlaEqpSlot)

            End If

            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.AlaEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            UserList(UserIndex).Invent.AlaEqpSlot = Slot

            If UserList(UserIndex).flags.Mimetizado = 1 Then
                UserList(UserIndex).CharMimetizado.Alas = Obj.Ropaje
            Else
                UserList(UserIndex).char.Alas = Obj.Ropaje
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                                UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                             UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

            End If

        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z42")
            Exit Sub

        End If

        '[/MaTeO 9]

    Case eOBJType.otWeapon

        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
            If ObjData(ObjIndex).DosManos = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes usar un arma de dos manos si estás usando un escudo." & FONTTYPE_INFO)
                Exit Sub
            End If
        End If

        'Si esta equipado lo quita
        If UserList(UserIndex).Invent.Object(Slot).Equipped Then
            'Quitamos del inv el item
            Call Desequipar(UserIndex, Slot)

            'Animacion por defecto
            If UserList(UserIndex).flags.Mimetizado = 1 Then
                UserList(UserIndex).CharMimetizado.WeaponAnim = NingunArma
            Else
                UserList(UserIndex).char.WeaponAnim = NingunArma
                '[MaTeO 9]
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                                UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                             UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

                '[/MaTeO 9]
            End If

            Exit Sub

        End If

        'Quitamos el elemento anterior
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)

        End If

        UserList(UserIndex).Invent.Object(Slot).Equipped = 1
        UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        UserList(UserIndex).Invent.WeaponEqpSlot = Slot

        'Sonido
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_SACARARMA)

        If UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).CharMimetizado.WeaponAnim = Obj.WeaponAnim
        Else
            UserList(UserIndex).char.WeaponAnim = Obj.WeaponAnim
            '[MaTeO 9]
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                            UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                         UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

            '[/MaTeO 9]
        End If

        If Obj.ObjetoEspecial > 0 Then
            Call DarObjetoEspecial(UserIndex, val(Obj.ObjetoEspecial))

        End If

        If EspadaSagrada.EspadaSagrada(UserList(UserIndex).Invent.WeaponEqpObjIndex) Then
            Call ChangeSagradaHit(UserIndex)

        End If

    Case eOBJType.otPARAA

        If UserList(UserIndex).flags.Muerto = 1 Then
            Exit Sub
        End If

        If UserList(UserIndex).flags.Paralizado = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡No estás Paralizado!! " & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(UserIndex).flags.Paralizado = 1 Then
            UserList(UserIndex).flags.Paralizado = 0
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "PARADOW")
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).pos.X & "," & UserList(UserIndex).pos.Y)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te has quitado la paralisis." & FONTTYPE_INFO)

            Call QuitarUserInvItem(UserIndex, Slot, 1)
            Call UpdateUserInv(False, UserIndex, Slot)
        End If

    Case eOBJType.otAmuletoDefensa
        If UserList(UserIndex).flags.Muerto = 1 Then
            Exit Sub
        End If

        If UserList(UserIndex).Invent.Object(Slot).Equipped Then
            'Quitamos del inv el item
            Call Desequipar(UserIndex, Slot)
            Exit Sub
        End If

        If UserList(UserIndex).Invent.AmuletoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.AmuletoEqpSlot)
        End If

        UserList(UserIndex).Invent.Object(Slot).Equipped = 1
        UserList(UserIndex).Invent.AmuletoEqpObjIndex = ObjIndex
        UserList(UserIndex).Invent.AmuletoEqpSlot = Slot

    Case eOBJType.otherramientas

        'Si esta equipado lo quita
        If UserList(UserIndex).Invent.Object(Slot).Equipped Then
            'Quitamos del inv el item
            Call Desequipar(UserIndex, Slot)
            Exit Sub

        End If

        'Quitamos el elemento anterior
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
        End If

        UserList(UserIndex).Invent.Object(Slot).Equipped = 1
        UserList(UserIndex).Invent.HerramientaEqpObjIndex = ObjIndex
        UserList(UserIndex).Invent.HerramientaEqpSlot = Slot


    Case eOBJType.otFlechas

        'Si esta equipado lo quita
        If UserList(UserIndex).Invent.Object(Slot).Equipped Then
            'Quitamos del inv el item
            Call Desequipar(UserIndex, Slot)
            Exit Sub

        End If

        'Quitamos el elemento anterior
        If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)

        End If

        UserList(UserIndex).Invent.Object(Slot).Equipped = 1
        UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        UserList(UserIndex).Invent.MunicionEqpSlot = Slot

        If Obj.ObjetoEspecial > 0 Then
            Call DarObjetoEspecial(UserIndex, val(Obj.ObjetoEspecial))

        End If

    Case eOBJType.otArmadura

        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub

        'Si esta equipado lo quita
        If UserList(UserIndex).Invent.Object(Slot).Equipped Then
            Call Desequipar(UserIndex, Slot)
            Call DarCuerpoDesnudo(UserIndex, UserList(UserIndex).flags.Mimetizado = 1)

            If Not UserList(UserIndex).flags.Mimetizado = 1 Then
                '[MaTeO 9]
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                                UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                             UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

                '[/MaTeO 9]
            End If

            Exit Sub

        End If

        'Quita el anterior
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)

        End If

        'Lo equipa

        UserList(UserIndex).Invent.Object(Slot).Equipped = 1
        UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        UserList(UserIndex).Invent.ArmourEqpSlot = Slot

        If UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).CharMimetizado.Body = Obj.Ropaje
        Else
            UserList(UserIndex).char.Body = Obj.Ropaje
            '[MaTeO 9]
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                            UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                         UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

            '[/MaTeO 9]
        End If

        UserList(UserIndex).flags.Desnudo = 0

        If Obj.ObjetoEspecial > 0 Then
            Call DarObjetoEspecial(UserIndex, val(Obj.ObjetoEspecial))

        End If

    Case eOBJType.otCASCO

        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub

        'Si esta equipado lo quita
        If UserList(UserIndex).Invent.Object(Slot).Equipped Then
            Call Desequipar(UserIndex, Slot)

            If UserList(UserIndex).flags.Mimetizado = 1 Then
                UserList(UserIndex).CharMimetizado.CascoAnim = NingunCasco
            Else
                UserList(UserIndex).char.CascoAnim = NingunCasco
                '[MaTeO 9]
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                                UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                             UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

                '[/MaTeO 9]
            End If

            Exit Sub

        End If

        'Quita el anterior
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)

        End If

        'Lo equipa

        UserList(UserIndex).Invent.Object(Slot).Equipped = 1
        UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        UserList(UserIndex).Invent.CascoEqpSlot = Slot

        If UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).CharMimetizado.CascoAnim = Obj.CascoAnim
        Else
            UserList(UserIndex).char.CascoAnim = Obj.CascoAnim
            '[MaTeO 9]
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                            UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                         UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

            '[/MaTeO 9]
        End If

        If Obj.ObjetoEspecial > 0 Then
            Call DarObjetoEspecial(UserIndex, val(Obj.ObjetoEspecial))

        End If

    Case eOBJType.otESCUDO

        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).DosManos = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes usar un arma de dos manos si estás usando un escudo." & FONTTYPE_INFO)
                Exit Sub
            End If
        End If

        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub

        'Si esta equipado lo quita
        If UserList(UserIndex).Invent.Object(Slot).Equipped Then
            Call Desequipar(UserIndex, Slot)

            If UserList(UserIndex).flags.Mimetizado = 1 Then
                UserList(UserIndex).CharMimetizado.ShieldAnim = NingunEscudo
            Else
                UserList(UserIndex).char.ShieldAnim = NingunEscudo
                '[MaTeO 9]
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                                UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                             UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

                '[/MaTeO 9]
            End If

            Exit Sub

        End If

        'Quita el anterior
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)

        End If

        'Lo equipa

        UserList(UserIndex).Invent.Object(Slot).Equipped = 1
        UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        UserList(UserIndex).Invent.EscudoEqpSlot = Slot

        If UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).CharMimetizado.ShieldAnim = Obj.ShieldAnim
        Else
            UserList(UserIndex).char.ShieldAnim = Obj.ShieldAnim

            '[MaTeO 9]
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                            UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList( _
                                                                                                                                                                                                                         UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

            '[/MaTeO 9]
        End If


        If Obj.ObjetoEspecial > 0 Then
            Call DarObjetoEspecial(UserIndex, val(Obj.ObjetoEspecial))

        End If

    End Select

    'Actualiza
    Call UpdateUserInv(False, UserIndex, Slot)
    Call SendUserHitBox(UserIndex)
    Exit Sub
errhandler:
    Call LogError("EquiparInvItem Slot:" & Slot)

End Sub

Private Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean

    On Error GoTo errhandler

    If UserList(UserIndex).Raza = "Humano" Or UserList(UserIndex).Raza = "Elfo" Or UserList(UserIndex).Raza = "Elfo Oscuro" Or _
       UserList(UserIndex).Raza = "Licantropo" Or UserList(UserIndex).Raza _
     = "Ciclope" Then

        If ObjData(ItemIndex).RazaEnana = 0 And ObjData(ItemIndex).RazaHobbit = 0 And _
           ObjData(ItemIndex).RazaOrco = 0 And ObjData(ItemIndex).RazaVampiro = 0 Then
            CheckRazaUsaRopa = True
            Exit Function
        End If

    End If


    Select Case UserList(UserIndex).Raza
    Case "Enano", "Gnomo"
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
        Exit Function

    Case "Hobbit"
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaHobbit = 1)
        Exit Function

    Case "Orco"
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaOrco = 1)
        Exit Function

    Case "Vampiro"
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaVampiro = 1)
        Exit Function
    End Select

    CheckRazaUsaRopa = False

    'Verifica si la raza puede usar la ropa
    'If UserList(UserIndex).Raza = "Humano" Or UserList(UserIndex).Raza = "Elfo" Or UserList(UserIndex).Raza = "Elfo Oscuro" Or UserList( _
     '        UserIndex).Raza = "Orco" Or UserList(UserIndex).Raza = "Licantropo" Or UserList(UserIndex).Raza = "Vampiro" Or UserList(UserIndex).Raza _
     '        = "Ciclope" Then
    '    CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
    'Else
    '    CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
    'End If

    Exit Function
errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

    'Usa un item del inventario
    Dim Obj      As ObjData

    Dim ObjIndex As Integer

    Dim TargObj  As ObjData

    Dim MiObj    As Obj

    If UserList(UserIndex).Invent.Object(Slot).Amount = 0 Then Exit Sub

    Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)

    If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo los newbies pueden usar este objeto." & FONTTYPE_INFO)
        Exit Sub

    End If

    If Obj.ObjType = eOBJType.otWeapon Then

        If Obj.proyectil = 1 Then

            If UserList(UserIndex).Invent.Object(Slot).Equipped = 0 Then
                'Call SendData(SendTarget.toIndex, UserIndex, 0, "||No has seleccionado o equipado su objeto de combate!" & FONTTYPE_INFO)
                Exit Sub

            End If

            'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
            If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
        Else

            'dagas
            If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub

        End If

    Else

        If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub

    End If

    ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    UserList(UserIndex).flags.TargetObjInvIndex = ObjIndex
    UserList(UserIndex).flags.TargetObjInvSlot = Slot

    If Obj.ObjetoEspecial > 0 Then
        If UserList(UserIndex).Invent.Object(Slot).Equipped = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes equiparte el objeto para que surja efecto!" & FONTTYPE_INFO)
            Exit Sub

        End If

        Call UseObjetoEspecial(UserIndex, val(Obj.ObjetoEspecial))

    End If

    Select Case Obj.ObjType

        Case eOBJType.otPack

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            Call AbrirPack(UserIndex, ObjIndex, Slot)

            Call QuitarUserInvItem(UserIndex, Slot, 1)
            Call UpdateUserInv(False, UserIndex, Slot)

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "TW131")

            Exit Sub

        Case eOBJType.otVales

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            If UserList(UserIndex).Stats.ELV = STAT_MAXELV Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Ya eres level maximo!" & FONTTYPE_INFO)
                Exit Sub

            End If

            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + Obj.Expe

            Call CheckUserLevel(UserIndex)
            Call EnviarExp(UserIndex)
            'Quitamos del inv el item

            Call QuitarUserInvItem(UserIndex, Slot, 1)

            Call UpdateUserInv(False, UserIndex, Slot)

        Case eOBJType.otUseOnce

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            If UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MaxHam Then
                Exit Sub

            End If

            'Usa el item
            UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam + Obj.MinHam

            If UserList(UserIndex).Stats.MinHam > UserList(UserIndex).Stats.MaxHam Then UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MaxHam
            UserList(UserIndex).flags.Hambre = 0
            Call EnviarHambreYsed(UserIndex)
            'Sonido

            If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
                Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, e_SoundIndex.MORFAR_MANZANA)
            Else
                Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, e_SoundIndex.SOUND_COMIDA)

            End If

            'Quitamos del inv el item

            'CRAW; 03/04/2020 --> DESACTIVAMOS ESTO POR AHORA
            Call QuitarUserInvItem(UserIndex, Slot, 1)

            Call UpdateUserInv(False, UserIndex, Slot)

        Case eOBJType.otGuita

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(UserIndex).Invent.Object(Slot).Amount
            UserList(UserIndex).Invent.Object(Slot).Amount = 0
            UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
            UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1

            Call UpdateUserInv(False, UserIndex, Slot)
            Call EnviarOro(UserIndex)

        Case eOBJType.otWeapon

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            If ObjData(ObjIndex).proyectil = 1 Then

                If Not UserList(UserIndex).flags.SeguroCombate Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate." & FONTTYPE_Motd4)
                    Exit Sub

                End If

                Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Proyectiles)

                If Obj.ObjetoEspecial > 0 Then
                    UserList(UserIndex).flags.EspecialArco = 1
                    UserList(UserIndex).flags.EspecialObjArco = Obj.ObjetoEspecial
                    'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Es un Arco Especial" & FONTTYPE_FIGHT)

                End If

            Else

                If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub

                '¿El target-objeto es leña?
                If UserList(UserIndex).flags.TargetObj = Leña Then
                    If UserList(UserIndex).Invent.Object(Slot).ObjIndex = DAGA Then
                        Call TratarDeHacerFogata(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY, UserIndex)

                    End If

                End If

            End If

        Case eOBJType.otAmuleto

            If UserList(UserIndex).flags.Muerto = 1 Then
                Exit Sub

            End If

            Call AmuTeleport(UserIndex)
            Call QuitarUserInvItem(UserIndex, Slot, 1)

            Call UpdateUserInv(False, UserIndex, Slot)
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW100")

        Case eOBJType.otRegalos

            If UserList(UserIndex).flags.Muerto = 1 Then
                Exit Sub

            End If

            Dim RandomRegalo As Integer

            RandomRegalo = RandomNumber(1, NumRegalos)

            MiObj.ObjIndex = Regalos(RandomRegalo).ObjIndex
            MiObj.Amount = 1
            Call MeterItemEnInventario(UserIndex, MiObj)
            Call QuitarUserInvItem(UserIndex, Slot, 1)

            Call UpdateUserInv(False, UserIndex, Slot)

        Case eOBJType.otPociones

            'If UserList(userindex).Lac.LPociones.Puedo = False Then Exit Sub
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            If UserList(UserIndex).Counters.TimerAttack > 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes esperar un momento para tomar otra poción." & FONTTYPE_INFO)
                Exit Sub

            End If

            Select Case Obj.TipoPocion

                    Dim MXATRIBUTOS As String

                Case 1    'Modif la agilidad
                    UserList(UserIndex).flags.DuracionEfectoAmarillas = Obj.DuracionEfecto

                    'Usa el item
                    UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)

                    MXATRIBUTOS = val(MAXATRIBUTOS + UserList(UserIndex).flags.EspecialAgilidad)

                    If UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) > MXATRIBUTOS Then UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = MXATRIBUTOS

                    'Quitamos del inv el item
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    UserList(UserIndex).flags.TomoPocionAmarilla = True
                    Call EnviarAmarillas(UserIndex)
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)

                Case 2    'Modif la fuerza
                    UserList(UserIndex).flags.DuracionEfectoVerdes = Obj.DuracionEfecto

                    'Usa el item
                    UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)

                    MXATRIBUTOS = val(MAXATRIBUTOS + UserList(UserIndex).flags.EspecialFuerza)

                    If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) > MXATRIBUTOS Then UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = MXATRIBUTOS

                    'Quitamos del inv el item
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    UserList(UserIndex).flags.TomoPocionVerde = True
                    Call EnviarVerdes(UserIndex)
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)

                Case 3    'Pocion roja, restaura HP

                    If UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP Then
                        Exit Sub

                    End If

                    'Usa el item
                    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + RandomNumber(Obj.MinModificador, Obj.MaxModificador)

                    If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP

                    'Quitamos del inv el item
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)
                    Call EnviarHP(UserIndex)

                Case 4    'Pocion azul, restaura MANA

                    If UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN Then
                        Exit Sub

                    End If

                    'Usa el item
                    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + Porcentaje(UserList(UserIndex).Stats.MaxMAN, 5)

                    If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN

                    'Quitamos del inv el item
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)
                    Call EnviarMn(UserIndex)

                Case 5    ' Pocion violeta

                    If UserList(UserIndex).flags.Envenenado = 1 Then
                        UserList(UserIndex).flags.Envenenado = 0
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te has curado del envenenamiento." & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No estás envenenado!!" & FONTTYPE_INFO)

                    End If

                Case 6    ' Pocion desparalizar

                    If Obj.MinSkill > UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, "0", "||Necesitarás " & Obj.MinSkill & " Skills en Navegación para utilizarlo. " & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    If UserList(UserIndex).flags.Paralizado = 0 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La pócima no puede funcionar por que no estás paralizado." & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    If UserList(UserIndex).flags.Paralizado = 1 Then
                        UserList(UserIndex).flags.Paralizado = 0
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "PARADOW")
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).pos.X & "," & UserList(UserIndex).pos.Y)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La pocima surte su efecto y te has desparalizado!" & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)

                    End If

                    'Quitamos del inv el item
                    Call QuitarUserInvItem(UserIndex, Slot, 1)

                Case 7    'Pocion invisibilidad

                    If Obj.MinSkill > UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, "0", "||Necesitarás " & Obj.MinSkill & " Skills en Navegación para utilizarlo. " & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    If UserList(UserIndex).pos.Map = mapainvo Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes lanzar invisibilidad en sala de invocaciones!" & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    If UserList(UserIndex).pos.Map = MAPADUELO Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes lanzar invisibilidad en duelo!" & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    If UserList(UserIndex).flags.Invisible = 1 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Ya estás invisible!!" & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    If UserList(UserIndex).flags.Invisible = 0 Then
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        UserList(UserIndex).flags.Invisible = 1
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te has vuelto invisible." & FONTTYPE_INFO)
                        Call UpdateUserInv(False, UserIndex, Slot)
                        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)
                        Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, "NOVER" & UserList(UserIndex).char.CharIndex & ",1," & UserList(UserIndex).PartyIndex)

                    End If

                Case 8    'Pocion Telepatia

                    If Obj.MinSkill > UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, "0", "||Necesitarás " & Obj.MinSkill & " Skills en Navegación para utilizarlo. " & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    If UserList(UserIndex).Telepatia = 1 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya tienes aprendida la telepatia!" & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Call QuitarUserInvItem(UserIndex, Slot, 1)

                    UserList(UserIndex).Telepatia = 1

                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has aprendido la telepatia!" & FONTTYPE_INFO)
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)

                Case 9    'Teleport a Nix

                    If Obj.MinSkill > UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, "0", "||Necesitarás " & Obj.MinSkill & " Skills en Navegación para utilizarlo. " & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Call QuitarUserInvItem(UserIndex, Slot, 1)

                    Call WarpUserChar(UserIndex, 34, 40, 50, True)

                Case 10    'pocion energia

                    If UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta Then
                        Exit Sub

                    End If

                    Call QuitarUserInvItem(UserIndex, Slot, 1)

                    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta + RandomNumber(Obj.MinModificador, Obj.MaxModificador)

                    If UserList(UserIndex).Stats.MinSta > UserList(UserIndex).Stats.MaxSta Then UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
                    Call EnviarSta(UserIndex)
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)

                Case 11    'Pocion para remover Ceguera

                    If Obj.MinSkill > UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, "0", "||Necesitarás " & Obj.MinSkill & " Skills en Navegación para utilizarlo. " & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
                    If UserList(UserIndex).flags.Ceguera = 0 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No estás ciego" & FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    Call SendData(ToIndex, UserIndex, 0, "NSEGUE")
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La pocima surte su efecto y recuperas la visión." & FONTTYPE_INFO)
                    Call QuitarUserInvItem(UserIndex, Slot, 1)

                Case 12    'Pocion para remover Estupidez

                    If Obj.MinSkill > UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, "0", "||Necesitarás " & Obj.MinSkill & " Skills en Navegación para utilizarlo. " & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
                    If UserList(UserIndex).flags.Estupidez = 0 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No estás estupido." & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    UserList(UserIndex).flags.Estupidez = 0
                    Call SendData(ToIndex, UserIndex, 0, "NESTUP")
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La pocima surte su efecto y recobras la cordura." & FONTTYPE_INFO)
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)

                Case 13    'Teleport a Ulla

                    If Obj.MinSkill > UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, "0", "||Necesitarás " & Obj.MinSkill & " Skills en Navegación para utilizarlo. " & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call WarpUserChar(UserIndex, 1, 52, 53, True)
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)

                Case 14    'Teleport a Bander

                    If Obj.MinSkill > UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, "0", "||Necesitarás " & Obj.MinSkill & " Skills en Navegación para utilizarlo. " & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call WarpUserChar(UserIndex, 59, 50, 41, True)
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)

                Case 15    'Pocion azul, restaura MANA

                    If UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN Then
                        Exit Sub

                    End If

                    'Usa el item
                    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + Porcentaje(UserList(UserIndex).Stats.MaxMAN, 5)

                    If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN

                    'Quitamos del inv el item
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)
                    Call EnviarMn(UserIndex)

                Case 16    'Pocion sistema criatura

                    If DiaEspecialExp = True Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Los Dioses de AoMania no te permiten usar tu poder para cambiar este día especial." & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    If DiaEspecialOro = True Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Los Dioses de AoMania no te permiten usar tu poder para cambiar este día especial." & FONTTYPE_INFO)
                        Exit Sub

                    End If

                    RandomCase = RandomNumber(1, 15)
                    Call CriaturasNormales(RandomCase)
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)

            End Select

            Call UpdateUserInv(False, UserIndex, Slot)

        Case eOBJType.otLibromagico

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No puedes usar el item si estás muerto!" & FONTTYPE_INFO)
                Exit Sub

            End If
        
            If UserList(UserIndex).flags.UsoLibroHP = 15 Then
                SendData SendTarget.ToIndex, UserIndex, 0, "||¡Ya no puedes ganar más vida!." & FONTTYPE_INFO
                Exit Sub

            End If
        
            If UserList(UserIndex).Stats.MaxHP > STAT_MAXHP Then UserList(UserIndex).Stats.MaxHP = STAT_MAXHP

            UserList(UserIndex).Stats.MaxHP = UserList(UserIndex).Stats.MaxHP + 1
        
            '        If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then _
            '           UserList(UserIndex).Stats.MaxHP = UserList(UserIndex).Stats.MaxHP + 1

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has ganado 1 punto de vida!" & FONTTYPE_INFO)
            Call SendUserStatsBox(UserIndex)
            Call QuitarUserInvItem(UserIndex, Slot, 1)    'te quito el item que ya te hice el efecto
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_CHIRP)            'Plus de sonido escojan el que quieran
            Call UpdateUserInv(True, UserIndex, 0)

        Case eOBJType.otBebidas

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            If UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU Then
                Exit Sub

            End If

            UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU + Obj.MinSed

            If UserList(UserIndex).Stats.MinAGU > UserList(UserIndex).Stats.MaxAGU Then UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
            UserList(UserIndex).flags.Sed = 0
            Call EnviarHambreYsed(UserIndex)

            'Quitamos del inv el item
            Call QuitarUserInvItem(UserIndex, Slot, 1)

            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_BEBER)

            Call UpdateUserInv(False, UserIndex, Slot)

        Case eOBJType.otLlaves

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
            TargObj = ObjData(UserList(UserIndex).flags.TargetObj)

            '¿El objeto clickeado es una puerta?
            If TargObj.ObjType = eOBJType.otPuertas Then

                '¿Esta cerrada?
                If TargObj.Cerrada = 1 Then

                    '¿Cerrada con llave?
                    If TargObj.Llave > 0 Then
                        If TargObj.clave = Obj.clave Then

                            MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerrada
                            UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has abierto la puerta." & FONTTYPE_INFO)
                            Exit Sub
                        Else
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La llave no sirve." & FONTTYPE_INFO)
                            Exit Sub

                        End If

                    Else

                        If TargObj.clave = Obj.clave Then
                            MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerradaLlave
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has cerrado con llave la puerta." & FONTTYPE_INFO)
                            UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
                            Exit Sub
                        Else
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La llave no sirve." & FONTTYPE_INFO)
                            Exit Sub

                        End If

                    End If

                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No esta cerrada." & FONTTYPE_INFO)
                    Exit Sub

                End If

            End If

        Case eOBJType.otCheques

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + 10000

            'Quitamos del inv el item
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            Call UpdateUserInv(False, UserIndex, Slot)
            Call SendUserStatsBox(UserIndex)

        Case eOBJType.otBotellaVacia

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            If Not HayAgua(UserList(UserIndex).pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay agua allí." & FONTTYPE_INFO)
                Exit Sub

            End If

            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).IndexAbierta
            Call QuitarUserInvItem(UserIndex, Slot, 1)

            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)

            End If

            Call UpdateUserInv(False, UserIndex, Slot)

        Case eOBJType.otBotellaLlena

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU + Obj.MinSed

            If UserList(UserIndex).Stats.MinAGU > UserList(UserIndex).Stats.MaxAGU Then UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
            UserList(UserIndex).flags.Sed = 0
            Call EnviarHambreYsed(UserIndex)
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).IndexCerrada
            Call QuitarUserInvItem(UserIndex, Slot, 1)

            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)

            End If

        Case eOBJType.otherramientas

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            If Not UserList(UserIndex).Stats.MinSta > 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas muy cansado" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).Invent.Object(Slot).Equipped = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Antes de usar la herramienta deberias equipartela." & FONTTYPE_INFO)
                Exit Sub

            End If

            UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta

            If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then UserList(UserIndex).Reputacion.PlebeRep = MAXREP

            Select Case ObjIndex

                Case TIJERA
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Sastreria)

                Case AGUJA
                    Call EnviarObjSastre(UserIndex)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ABRS")

                Case HOZ_DE_MANO
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Recolectar)

                Case MORTERO
                    Call EnviarObjHechiceria(UserIndex)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ABRH")

                Case CAÑA_PESCA, RED_PESCA
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Pesca)

                Case HACHA_LEÑADOR
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & talar)

                Case PIQUETE_MINERO
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Mineria)

                Case MARTILLO_HERRERO
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Herreria)

                Case MARTILLO_HERRERO_MAGICO
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Herrero)

                Case SERRUCHO_CARPINTERO
                    Call EnivarObjConstruibles(UserIndex)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "SFC")

            End Select

        Case eOBJType.otPergaminos

            If UserList(UserIndex).flags.Muerto = 1 Then
                ' Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo puedes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).Stats.MinHam <= 30 And UserList(UserIndex).Stats.MinAGU <= 30 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado hambriento y sediento." & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).flags.Privilegios = PlayerType.User Then

                If ObjData(ObjIndex).Gm = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Este hechizo solo puede aprenderlos los gamemaster," & FONTTYPE_INFO)
                    Exit Sub

                End If

                If ObjData(ObjIndex).Real = 1 Or ObjData(ObjIndex).Caos = 1 Or ObjData(ObjIndex).Nemes = 1 Or ObjData(ObjIndex).Templ = 1 Then
                    If Not FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu faccion no puede aprender este hechizo." & FONTTYPE_INFO)
                        Exit Sub

                    End If

                End If

                If UCase$(UserList(UserIndex).Clase) = "GUERRERO" Or UCase(UserList(UserIndex).Clase) = "ARQUERO" Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tú clase no puede aprender este hechizo." & FONTTYPE_INFO)
                    Exit Sub

                End If

                If ClasePuedeUsarItem(UserIndex, ObjIndex) Or FaccionPuedeUsarItem(UserIndex, ObjIndex) Then

                    If UCase$(Len(Hechizos(Obj.HechizoIndex).ExclusivoClase)) > 0 Or UCase$(Hechizos(Obj.HechizoIndex).ExclusivoClase) <> UCase$(UserList(UserIndex).Clase) Then

                        If ClasePuedeLanzarHechizo(UserIndex, Obj.HechizoIndex) = False Then
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tú clase no puede aprender este hechizo." & FONTTYPE_INFO)
                            Exit Sub

                        End If

                        Call AgregarHechizo(UserIndex, Slot)
                        Call UpdateUserInv(False, UserIndex, Slot)
                        Exit Sub

                    End If

                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tú clase no puede aprender este hechizo." & FONTTYPE_INFO)
                    Exit Sub

                End If

            Else
                Call AgregarHechizo(UserIndex, Slot)
                Call UpdateUserInv(False, UserIndex, Slot)

            End If

        Case eOBJType.otMinerales

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & FundirMetal)

        Case eOBJType.otInstrumentos

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub

            End If

            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & Obj.Snd1)

        Case eOBJType.otBarcos

            If UserList(UserIndex).flags.Montado = True Then Exit Sub
            If UserList(UserIndex).flags.Metamorfosis = 1 Then Exit Sub
            If UserList(UserIndex).flags.Licantropo = 1 Then Exit Sub
                                                            
            If Not FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "||No perteneces a la faccion requerida para usar este barco." & FONTTYPE_INFO)
            ElseIf ClasePuedeUsarItem(UserIndex, ObjIndex) And ((LegalPos(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X - 1, UserList(UserIndex).pos.Y, True) Or LegalPos(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1, True) Or LegalPos(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X + 1, UserList(UserIndex).pos.Y, True) Or LegalPos(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y + 1, True)) And UserList(UserIndex).flags.Navegando = 0) Or UserList(UserIndex).flags.Navegando = 1 Then
                    
                If HayAgua(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y) And UserList(UserIndex).flags.Navegando = 1 Then
                    If Not ((LegalPos(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X - 1, UserList(UserIndex).pos.Y, False) Or LegalPos(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1, False) Or LegalPos(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X + 1, UserList(UserIndex).pos.Y, False) Or LegalPos(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y + 1, False))) Then
                        Call SendData(ToIndex, UserIndex, 0, "||¡No puedes sacarte el barco si estás en el agua!" & FONTTYPE_INFO)
                        Exit Sub

                    End If

                End If

                UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.BarcoSlot = Slot
                Call DoNavega(UserIndex, Obj)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Debes aproximarte al agua para usar el barco!" & FONTTYPE_INFO)

            End If

        Case eOBJType.otMontura
            ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            Obj = ObjData(ObjIndex)

            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas muerto!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).flags.Navegando = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas navegando!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).pos.Map = 154 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes usar tu mascota en guerra!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).pos.Map = 162 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes usar tu mascota en guerra!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).pos.Map = 163 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes usar tu mascota en guerra!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).pos.Map = 96 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes usar tu mascota en guerra!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).pos.Map = 98 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes usar tu mascota en el castillo!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).pos.Map = 99 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes usar tu mascota en el castillo!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).pos.Map = 100 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes usar tu mascota en el castillo!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).pos.Map = 101 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes usar tu mascota en el castillo!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).pos.Map = 102 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes usar tu mascota en el castillo!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).pos.Map = 164 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes usar tu mascota en fortaleza fuerte!!" & FONTTYPE_INFO)
                Exit Sub

            End If

            If UserList(UserIndex).flags.Montado = True Then
                UserList(UserIndex).char.Body = UserList(UserIndex).flags.NumeroMont
                '[MaTeO 9]
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)
                '[/MaTeO 9]
                UserList(UserIndex).flags.Montado = False
                Exit Sub

            End If

            If UserList(UserIndex).flags.Montado = False Then
                UserList(UserIndex).flags.NumeroMont = UserList(UserIndex).char.Body
                UserList(UserIndex).char.Body = Obj.Ropaje
                '[MaTeO 9]
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)
                '[/MaTeO 9]
                UserList(UserIndex).flags.Montado = True

            End If

    End Select

    'Actualiza
    Call UpdateUserInv(False, UserIndex, Slot)

End Sub

Sub EnviarArmasMagicasConstruibles(ByVal UserIndex As Integer)

    Dim i As Integer, cad$

    For i = 1 To UBound(ObjArmaHerreroMagico)
        If ObjData(ObjArmaHerreroMagico(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(eSkill.Herrero) \ ModHerreriA(UserList(UserIndex).Clase) Then
            If ObjData(ObjArmaHerreroMagico(i)).ObjType = eOBJType.otWeapon Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "OBHM" & ObjArmaHerreroMagico(i) & "@" & ObjData(ObjArmaHerreroMagico(i)).Name & " (" & ObjData(ObjArmaHerreroMagico(i)).MinHit & "/" & ObjData(ObjArmaHerreroMagico(i)).MaxHit & ")")
            End If
        End If
    Next i

End Sub

Sub EnviarArmadurasMagicasConstruibles(ByVal UserIndex As Integer)
    Dim i As Integer, cad$

    For i = 1 To UBound(ObjArmaduraHerreroMagico)
        If ObjData(ObjArmaduraHerreroMagico(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(eSkill.Herrero) \ ModHerreriA(UserList(UserIndex).Clase) Then
            If ObjData(ObjArmaduraHerreroMagico(i)).ObjType = eOBJType.otArmadura Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "OBHM" & ObjArmaduraHerreroMagico(i) & "@" & ObjData(ObjArmaduraHerreroMagico(i)).Name & " (" & ObjData(ObjArmaduraHerreroMagico(i)).MinDef & "/" & ObjData(ObjArmaduraHerreroMagico(i)).MaxDef & ")")
            End If
        End If
    Next i

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)

    Dim i As Integer, cad$

    For i = 1 To UBound(ArmasHerrero)

        If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) \ ModHerreriA(UserList(UserIndex).Clase) Then

            If ObjData(ArmasHerrero(i)).ObjType = eOBJType.otWeapon Then
                '[DnG!]
                cad$ = cad$ & ObjData(ArmasHerrero(i)).Name & " (" & ObjData(ArmasHerrero(i)).LingH & "-" & ObjData(ArmasHerrero(i)).LingP & "-" & _
                       ObjData(ArmasHerrero(i)).LingO & ")" & "," & ArmasHerrero(i) & ","
                '[/DnG!]
            Else
                cad$ = cad$ & ObjData(ArmasHerrero(i)).Name & "," & ArmasHerrero(i) & ","

            End If

        End If

    Next i

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "LAH" & cad$)

End Sub

Sub EnviarObjSastre(ByVal UserIndex As Integer)

    Dim i As Integer

    For i = 1 To UBound(ObjSastre)
        If ObjData(ObjSastre(i)).SkSastreria <= UserList(UserIndex).Stats.UserSkills(eSkill.Sastreria) / ModSastreria(UserList(UserIndex).Clase) _
           Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "LSTS" & ObjSastre(i) & "@" & ObjData(ObjSastre(i)).Name & " (Lana: " & ObjData(ObjSastre(i)).Lana & ") (Piel Lobo: " & ObjData(ObjSastre(i)).Lobo & ") (Piel Osos: " & ObjData(ObjSastre(i)).Osos & ") (Piel Tigre: " & ObjData(ObjSastre(i)).Tigre & ") (P.Oso Polar: " & ObjData(ObjSastre(i)).OsoPolar & ") (Piel Vaca: " & ObjData(ObjSastre(i)).Vaca & ") (Piel Jabali: " & ObjData(ObjSastre(i)).Jabali & ")")
        End If
    Next i

End Sub

Sub EnviarObjHechiceria(ByVal UserIndex As Integer)

    Dim i As Integer

    For i = 1 To UBound(ObjHechizeria)
        If ObjData(ObjHechizeria(i)).SkHechiceria <= UserList(UserIndex).Stats.UserSkills(eSkill.Hechiceria) / ModHechizeria(UserList(UserIndex).Clase) _
           Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "OBJH" & ObjHechizeria(i) & "@" & ObjData(ObjHechizeria(i)).Name & " (" & ObjData(ObjHechizeria(i)).Hierba & ")")
        End If
    Next i

End Sub

Sub EnivarObjConstruibles(ByVal UserIndex As Integer)

    Dim i As Integer, cad$

    For i = 1 To UBound(ObjCarpintero)

        If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) / ModCarpinteria(UserList( _
                                                                                                                                UserIndex).Clase) Then cad$ = cad$ & ObjData(ObjCarpintero(i)).Name & " (" & ObjData(ObjCarpintero(i)).Madera & ")" & "," & _
                                                                                                                                                                             ObjCarpintero(i) & ","
    Next i

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "OBR" & cad$)

End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)

    Dim i As Integer, cad$

    For i = 1 To UBound(ArmadurasHerrero)

        If ObjData(ArmadurasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList( _
                                                                                                                          UserIndex).Clase) Then
            '[DnG!]
            cad$ = cad$ & ObjData(ArmadurasHerrero(i)).Name & " (" & ObjData(ArmadurasHerrero(i)).LingH & "-" & ObjData(ArmadurasHerrero(i)).LingP _
                 & "-" & ObjData(ArmadurasHerrero(i)).LingO & ")" & "," & ArmadurasHerrero(i) & ","

            '[/DnG!]
        End If

    Next i

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "LAR" & cad$)

End Sub

Sub TirarTodo(ByVal UserIndex As Integer)

    On Error Resume Next

    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Trigger = eTrigger.ZONAPELEA Then Exit Sub

    Call TirarTodosLosItems(UserIndex)

    If UserList(UserIndex).Stats.GLD < 100001 Then Call TirarOro(UserList(UserIndex).Stats.GLD, UserIndex)

End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean

    ItemSeCae = (ObjData(Index).Real <> 1 Or ObjData(Index).Cae = 0) And (ObjData(Index).Caos <> 1 Or ObjData(Index).Cae = 0) And ObjData( _
                Index).ObjType <> eOBJType.otLlaves And ObjData(Index).ObjType <> eOBJType.otBarcos And ObjData(Index).Cae = 0

End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    If MapInfo(UserList(UserIndex).pos.Map).Cae = 1 Then Exit Sub

    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex

        If ItemIndex > 0 Then
            If ItemSeCae(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0

                'Creo el Obj
                MiObj.Amount = UserList(UserIndex).Invent.Object(i).Amount
                MiObj.ObjIndex = ItemIndex

                Call Tilelibre(UserList(UserIndex).pos, NuevaPos)

                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

                End If

            End If

        End If

    Next i

End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

    ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer

    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Trigger = 6 Then Exit Sub

    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex

        If ItemIndex > 0 Then
            If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0

                'Creo MiObj
                MiObj.Amount = UserList(UserIndex).Invent.Object(i).ObjIndex
                MiObj.ObjIndex = ItemIndex

                Tilelibre UserList(UserIndex).pos, NuevaPos

                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).OBJInfo.ObjIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, _
                                                                                                            NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

                End If

            End If

        End If

    Next i

End Sub

Public Sub DarObjetoEspecial(UserIndex As Integer, Objeto As Long)

    With UserList(UserIndex)

        Select Case Objeto

        Case "2"
            .flags.EspecialFuerza = val(.flags.EspecialFuerza + "5")
            .Stats.UserAtributos(eAtributos.Fuerza) = val(.Stats.UserAtributos(eAtributos.Fuerza) + "5")
            .Stats.UserAtributosBackUP(eAtributos.Fuerza) = .Stats.UserAtributosBackUP(eAtributos.Fuerza) + 5
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MF" & .Stats.UserAtributos(eAtributos.Fuerza))
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
            Call EnviarVerdes(UserIndex)
            Exit Sub

        Case "3"
            .flags.EspecialFuerza = val(.flags.EspecialFuerza + "2")
            .Stats.UserAtributos(eAtributos.Fuerza) = val(.Stats.UserAtributos(eAtributos.Fuerza) + "2")
            .Stats.UserAtributosBackUP(eAtributos.Fuerza) = .Stats.UserAtributosBackUP(eAtributos.Fuerza) + 2
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MF" & .Stats.UserAtributos(eAtributos.Fuerza))
            Call EnviarVerdes(UserIndex)
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
            Exit Sub

        Case "4"
            .flags.EspecialFuerza = val(.flags.EspecialFuerza + "3")
            .Stats.UserAtributos(eAtributos.Fuerza) = val(.Stats.UserAtributos(eAtributos.Fuerza) + "3")
            .Stats.UserAtributosBackUP(eAtributos.Fuerza) = .Stats.UserAtributosBackUP(eAtributos.Fuerza) + 3
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MF" & .Stats.UserAtributos(eAtributos.Fuerza))
            Call EnviarVerdes(UserIndex)
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
            Exit Sub

        Case "5"
            .flags.EspecialAgilidad = val(.flags.EspecialAgilidad + "5")
            .Stats.UserAtributos(eAtributos.Agilidad) = val(.Stats.UserAtributos(eAtributos.Agilidad) + "5")
            .Stats.UserAtributosBackUP(eAtributos.Fuerza) = .Stats.UserAtributosBackUP(eAtributos.Agilidad) + 5
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MA" & .Stats.UserAtributos(eAtributos.Agilidad))
            Call EnviarAmarillas(UserIndex)
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
            Exit Sub

        Case "6"
            .flags.EspecialAgilidad = val(.flags.EspecialAgilidad + "2")
            .Stats.UserAtributos(eAtributos.Agilidad) = val(.Stats.UserAtributos(eAtributos.Agilidad) + "2")
            .Stats.UserAtributosBackUP(eAtributos.Fuerza) = .Stats.UserAtributosBackUP(eAtributos.Agilidad) + 2
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MA" & .Stats.UserAtributos(eAtributos.Agilidad))
            Call EnviarAmarillas(UserIndex)
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
            Exit Sub

        Case "7"
            .flags.EspecialAgilidad = val(.flags.EspecialAgilidad + "3")
            .Stats.UserAtributos(eAtributos.Agilidad) = val(.Stats.UserAtributos(eAtributos.Agilidad) + "3")
            .Stats.UserAtributosBackUP(eAtributos.Fuerza) = .Stats.UserAtributosBackUP(eAtributos.Agilidad) + 3
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MA" & .Stats.UserAtributos(eAtributos.Agilidad))
            Call EnviarAmarillas(UserIndex)
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")

        Case "8"
            .Stats.MaxMAN = .Stats.MaxMAN + 100
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MM" & .Stats.MaxMAN)
            Call EnviarMn(UserIndex)
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
            Exit Sub

        Case "9"
            .Stats.MaxMAN = .Stats.MaxMAN + 200
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MM" & .Stats.MaxMAN)
            Call EnviarMn(UserIndex)
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
            Exit Sub

        Case "10"
            .Stats.MaxMAN = .Stats.MaxMAN + 300
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MM" & .Stats.MaxMAN)
            Call EnviarMn(UserIndex)
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
            Exit Sub

        End Select

    End With

End Sub

Private Sub QuitarObjetoEspecial(UserIndex As Integer, Objeto As Long)

    With UserList(UserIndex)

        Select Case Objeto

        Case "2"
            .flags.EspecialFuerza = val(.flags.EspecialFuerza - "5")
            .Stats.UserAtributos(eAtributos.Fuerza) = val(.Stats.UserAtributos(eAtributos.Fuerza) - "5")
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MF" & .Stats.UserAtributos(eAtributos.Fuerza))
            .Stats.UserAtributosBackUP(eAtributos.Fuerza) = .Stats.UserAtributosBackUP(eAtributos.Fuerza) - 5
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
            Call EnviarVerdes(UserIndex)

        Case "3"
            .flags.EspecialFuerza = val(.flags.EspecialFuerza - "2")
            .Stats.UserAtributos(eAtributos.Fuerza) = val(.Stats.UserAtributos(eAtributos.Fuerza) - "2")
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MF" & .Stats.UserAtributos(eAtributos.Fuerza))
            Call EnviarVerdes(UserIndex)
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")

        Case "4"
            .flags.EspecialFuerza = val(.flags.EspecialFuerza - "3")
            .Stats.UserAtributos(eAtributos.Fuerza) = val(.Stats.UserAtributos(eAtributos.Fuerza) - "3")
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MF" & .Stats.UserAtributos(eAtributos.Fuerza))
            Call EnviarVerdes(UserIndex)
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")

        Case "5"
            .flags.EspecialAgilidad = val(.flags.EspecialAgilidad - "5")
            .Stats.UserAtributos(eAtributos.Agilidad) = val(.Stats.UserAtributos(eAtributos.Agilidad) - "5")
            .Stats.UserAtributosBackUP(eAtributos.Fuerza) = .Stats.UserAtributosBackUP(eAtributos.Agilidad) - 5
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MA" & .Stats.UserAtributos(eAtributos.Agilidad))
            Call EnviarAmarillas(UserIndex)
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")

        Case "6"
            .flags.EspecialAgilidad = val(.flags.EspecialAgilidad - "2")
            .Stats.UserAtributos(eAtributos.Agilidad) = val(.Stats.UserAtributos(eAtributos.Agilidad) - "2")
            .Stats.UserAtributosBackUP(eAtributos.Fuerza) = .Stats.UserAtributosBackUP(eAtributos.Agilidad) - 2
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MA" & .Stats.UserAtributos(eAtributos.Agilidad))
            Call EnviarAmarillas(UserIndex)
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")

        Case "7"
            .flags.EspecialAgilidad = val(.flags.EspecialAgilidad - "3")
            .Stats.UserAtributos(eAtributos.Agilidad) = val(.Stats.UserAtributos(eAtributos.Agilidad) - "3")
            .Stats.UserAtributosBackUP(eAtributos.Fuerza) = .Stats.UserAtributosBackUP(eAtributos.Agilidad) - 3
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MA" & .Stats.UserAtributos(eAtributos.Agilidad))
            Call EnviarAmarillas(UserIndex)
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")

        Case "8"
            .Stats.MaxMAN = .Stats.MaxMAN - 100
            If .Stats.MinMAN > .Stats.MaxMAN Then
            .Stats.MinMAN = .Stats.MinMAN - 100
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MM" & .Stats.MaxMAN)
            Call EnviarMn(UserIndex)
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
            Exit Sub

        Case "9"
            .Stats.MaxMAN = .Stats.MaxMAN - 200
            If .Stats.MinMAN > .Stats.MaxMAN Then
            .Stats.MinMAN = .Stats.MinMAN - 200
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MM" & .Stats.MaxMAN)
            Call EnviarMn(UserIndex)
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
            Exit Sub

        Case "10"
            .Stats.MaxMAN = .Stats.MaxMAN - 300
            If .Stats.MinMAN > .Stats.MaxMAN Then
            .Stats.MinMAN = .Stats.MinMAN - 300
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MM" & .Stats.MaxMAN)
            Call EnviarMn(UserIndex)
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
            Exit Sub

        End Select

    End With

End Sub

Private Sub UseObjetoEspecial(UserIndex As Integer, Objeto As Long)
    Dim Recuperador As Integer

    With UserList(UserIndex)

        Select Case Objeto

        Case "50"

            If UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN Then
                Exit Sub

            End If

            Recuperador = RandomNumber(2, 8)

            'Usa el item
            UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + Recuperador

            If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then UserList(UserIndex).Stats.MinMAN = UserList( _
               UserIndex).Stats.MaxMAN

            'Quitamos del inv el item
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)
            Call EnviarMn(UserIndex)

        Case "51"

            If UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP Then
                Exit Sub

            End If

            Recuperador = RandomNumber(8, 15)

            'Usa el item
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + Recuperador

            If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList( _
               UserIndex).Stats.MaxHP

            'Quitamos del inv el item
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)
            Call EnviarHP(UserIndex)

        Case "55"

            Dim MXATRIBUTOS As String
            Recuperador = RandomNumber(5, 7)

            UserList(UserIndex).flags.DuracionEfectoAmarillas = 4800

            'Usa el item
            UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + _
                                                                           Recuperador

            MXATRIBUTOS = val(MAXATRIBUTOS + UserList(UserIndex).flags.EspecialAgilidad)

            If UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) > MXATRIBUTOS Then UserList(UserIndex).Stats.UserAtributos( _
               eAtributos.Agilidad) = MXATRIBUTOS

            'Quitamos del inv el item
            UserList(UserIndex).flags.TomoPocionAmarilla = True
            Call EnviarAmarillas(UserIndex)

            UserList(UserIndex).flags.DuracionEfectoVerdes = 4800

            'Usa el item
            UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + Recuperador

            MXATRIBUTOS = val(MAXATRIBUTOS + UserList(UserIndex).flags.EspecialFuerza)

            If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) > MXATRIBUTOS Then UserList(UserIndex).Stats.UserAtributos( _
               eAtributos.Fuerza) = MXATRIBUTOS

            'Quitamos del inv el item
            UserList(UserIndex).flags.TomoPocionVerde = True
            Call EnviarVerdes(UserIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & SND_POTEAR)

        End Select

    End With

End Sub

Sub RevObjetoEspecial(UserIndex As Integer, Objeto As Long)

    With UserList(UserIndex)

        Select Case Objeto

        Case "2"
            .flags.EspecialFuerza = val(.flags.EspecialFuerza + "5")

        Case "3"
            .flags.EspecialFuerza = val(.flags.EspecialFuerza + "2")

        Case "4"
            .flags.EspecialFuerza = val(.flags.EspecialFuerza + "3")

        Case "5"
            .flags.EspecialAgilidad = val(.flags.EspecialAgilidad + "5")

        Case "6"
            .flags.EspecialAgilidad = val(.flags.EspecialAgilidad + "2")

        Case "7"
            .flags.EspecialAgilidad = val(.flags.EspecialAgilidad + "3")

        End Select

    End With

End Sub

Private Sub AbrirPack(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Slot As Byte)

    Dim i As Integer
    Dim Objeto As Integer
    Dim Cantidad As Integer
    Dim MiObj As Obj

    For i = 1 To ObjData(ObjIndex).Pack.NumObjs

        Objeto = ObjData(ObjIndex).Pack.Obj(i).Objeto
        Cantidad = ObjData(ObjIndex).Pack.Obj(i).Cantidad

        MiObj.ObjIndex = Objeto
        MiObj.Amount = Cantidad

        Call MeterItemEnInventario(UserIndex, MiObj)

    Next i

    Call UpdateUserInv(False, UserIndex, Slot)

End Sub

Sub DarCabezaDragon(ByVal UserIndex As Integer, ByVal Color As String)

    Dim Obj As Obj

    Select Case UCase(Color)
    Case "ROJA"

        If UCase$(UserList(UserIndex).Raza) = "ENANO" Then
            Obj.ObjIndex = 483
        Else
            If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
                Obj.ObjIndex = 481
            Else
                Obj.ObjIndex = 482
            End If
        End If

    Case "NEGRA"

        If UCase$(UserList(UserIndex).Raza) = "ENANO" Then
            Obj.ObjIndex = 902
        Else
            If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
                Obj.ObjIndex = 900
            Else
                Obj.ObjIndex = 901
            End If
        End If

    Case "VERDE"

        If UCase$(UserList(UserIndex).Raza) = "ENANO" Then
            Obj.ObjIndex = 914
        Else
            If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
                Obj.ObjIndex = 912
            Else
                Obj.ObjIndex = 913
            End If
        End If

    Case "LILA"

        If UCase$(UserList(UserIndex).Raza) = "ENANO" Then
            Obj.ObjIndex = 929
        Else
            If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
                Obj.ObjIndex = 1169
            Else
                Obj.ObjIndex = 1170
            End If
        End If

    Case "BLANCA"

        If UCase$(UserList(UserIndex).Raza) = "ENANO" Then
            Obj.ObjIndex = 911
        Else
            If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
                Obj.ObjIndex = 909
            Else
                Obj.ObjIndex = 910
            End If
        End If

    Case "NARANJA"

        If UCase$(UserList(UserIndex).Raza) = "ENANO" Then
            Obj.ObjIndex = 908
        Else
            If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
                Obj.ObjIndex = 906
            Else
                Obj.ObjIndex = 907
            End If
        End If

    Case "AZUL"

        If UCase$(UserList(UserIndex).Raza) = "ENANO" Then
            Obj.ObjIndex = 905
        Else
            If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
                Obj.ObjIndex = 903
            Else
                Obj.ObjIndex = 904
            End If
        End If

    End Select

    Obj.Amount = 1

    Call MeterItemEnInventario(UserIndex, Obj)

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & "&H4580FF" & "°" & "Ahí tienes tu armadura!" & "°" _
                                                  & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))

End Sub

Sub MezclarAlas(ByVal UserIndex As Integer, ByVal iAV As Integer, ByVal iAN As Integer)

    Dim Obj As Obj
    Dim Porc As Byte

    If TieneObjetos(ObjCreacionAlas, 1, UserIndex) Then
        Porc = "100"
    Else
        Porc = RandomNumber(1, 100)
    End If

    If Porc > 75 Then

        If iAV > 0 Then
            Call QuitarObjetos(iAV, 1, UserIndex)
        End If

        Obj.ObjIndex = iAN
        Obj.Amount = 1
        Call MeterItemEnInventario(UserIndex, Obj)

        Call SendData(SendTarget.ToAll, 0, 0, "TW147")
        Call SendData(SendTarget.ToAll, 0, 0, "||El usuario " & UserList(UserIndex).Name & " ha creado " & Chr(147) & ObjData(iAN).Name & Chr(147) & " exitosamente." & FONTTYPE_TALKMSG)

    Else

        If iAV > 0 Then
            Call QuitarObjetos(iAV, 1, UserIndex)
        End If

        Call SendData(SendTarget.ToAll, 0, 0, "TW140")
        Call SendData(SendTarget.ToAll, 0, 0, "||El usuario " & UserList(UserIndex).Name & " ha fallado en crear " & Chr(147) & ObjData(iAN).Name & Chr(147) & " y ha perdido los items de la mezcla." & FONTTYPE_TALKMSG)

    End If

    If TieneObjetos(ObjCreacionAlas, 1, UserIndex) Then
        Call QuitarObjetos(ObjCreacionAlas, 1, UserIndex)
    End If

    Call QuitarObjetos(Plumas.Ampere, 1, UserIndex)
    Call QuitarObjetos(Plumas.Bassinger, 1, UserIndex)
    Call QuitarObjetos(Plumas.Seth, 1, UserIndex)

End Sub
