Attribute VB_Name = "Mod_Monedas"
'~~ Modulo: Mod_Monedas ~~
'~~ Desc: Monedas AoMCreditos y AoMCanjeos. ~~

Option Explicit

Public Type tAoMCreditos
    Name As String
    Index As Long
    Monedas As Long
End Type

Public Type tAoMCanjes
    Name As String
    Index As Long
    Monedas As Long
    Cantidad As Long
End Type

Public AoMCreditos() As tAoMCreditos
Public NumAoMCreditos As Long
Public NpcAoMCreditos As Long
Public AoMCanjes() As tAoMCanjes
Public NumAoMCanjes As Long
Public NpcAoMCanjes As Long

Public Sub Load_Creditos()

    Dim i As Long
    Dim NType As Long

    For i = 1 To NumNPCs

        NType = val(GetVar(App.Path & "\Dat\NPCs.dat", "NPC" & i, "NPCType"))

        If NType = eNPCType.Creditos Then
            NpcAoMCreditos = val(i)
        End If

    Next i

    NumAoMCreditos = val(GetVar(App.Path & "\Dat\AoMCreditos.dat", "INIT", "NumItems"))

    If NumAoMCreditos = 0 Then Exit Sub

    ReDim AoMCreditos(1 To MAX_INVENTORY_SLOTS)

    For i = 1 To NumAoMCreditos
        AoMCreditos(i).Name = GetVar(App.Path & "\Dat\AoMCreditos.dat", "Item" & i, "Name")
        AoMCreditos(i).Index = val(GetVar(App.Path & "\Dat\AoMCreditos.dat", "Item" & i, "Index"))
        AoMCreditos(i).Monedas = val(GetVar(App.Path & "\Dat\AoMCreditos.dat", "Item" & i, "Monedas"))
    Next i

End Sub

Public Sub IniciarComercioCreditos(ByVal UserIndex As Integer)

'Mandamos el Inventario
    Call EnviarCredInv(UserIndex, UserList(UserIndex).flags.TargetNpc)
    'Hacemos un Update del inventario del usuario
    Call UpdateUserInv(True, UserIndex, 0)
    'Atcualizamos las monedas
    Call EnviarCreditos(UserIndex)
    'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
    UserList(UserIndex).flags.Comerciando = True
    SendData SendTarget.ToIndex, UserIndex, 0, "INITCRE"
    Exit Sub

End Sub

Sub EnviarCredInv(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'Enviamos el inventario del npc con el cual el user va a comerciar...
    Dim i As Integer
    Dim val As Long

    For i = 1 To MAX_INVENTORY_SLOTS

        If i <= NumAoMCreditos Then

            SendData SendTarget.ToIndex, UserIndex, 0, "NPCC" & AoMCreditos(i).Name & "," & _
                                                       AoMCreditos(i).Index & "," & AoMCreditos(i).Monedas & "," & ObjData(AoMCreditos(i).Index).GrhIndex & "," & _
                                                       ObjData(AoMCreditos(i).Index).def & "," & ObjData(AoMCreditos(i).Index).MinHit & "," & ObjData(AoMCreditos(i).Index).MaxHit & "," & _
                                                       ObjData(AoMCreditos(i).Index).ObjType

        Else
            SendData SendTarget.ToIndex, UserIndex, 0, "NPCC" & "Nada" & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," _
                                                     & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0

        End If

    Next i

End Sub

Sub NpcVentaCredito(ByVal UserIndex As Integer, ByVal ObjIndex As Long, ByVal Creditos As Long, ByVal NpcIndex As Integer)
    Dim Slot As Integer
    Dim MiObj As Obj

    With UserList(UserIndex)


        Slot = 1
        Do Until .Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes tener mas objetos." & FONTTYPE_INFO)
                Exit Sub
            End If
        Loop

        If .AoMCreditos >= Creditos Then
            Call LogCreditos(.Name, .AoMCreditos, Creditos, ObjData(ObjIndex).Name)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Enhorabuena! Has comprado " & ObjData(ObjIndex).Name & " por " & Creditos & " AoMCreditos." & FONTTYPE_INFO)
            .AoMCreditos = .AoMCreditos - Creditos
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahora tienes " & .AoMCreditos & " AoMCreditos." & FONTTYPE_INFO)
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjIndex
            Call MeterItemEnInventario(UserIndex, MiObj)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes suficiente AoMCreditos." & FONTTYPE_WARNING)
            Exit Sub
        End If

    End With
    Call EnviarCreditos(UserIndex)
    Call UpdateVentanaCred(UserIndex)
    Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
End Sub

Public Sub UpdateVentanaCred(ByVal UserIndex As Integer)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "TRANSAC")
End Sub

Public Sub EnviarCreditos(ByVal UserIndex As Integer)
    If UserList(UserIndex).AoMCreditos < 0 Then UserList(UserIndex).AoMCreditos = 0
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "CRE" & UserList(UserIndex).AoMCreditos)
End Sub

'*****Canjes
Public Sub Load_Canjes()

    Dim i As Long
    Dim NType As Long

    For i = 1 To NumNPCs
        NType = val(GetVar(App.Path & "\Dat\NPCs.dat", "NPC" & i, "NPCType"))

        If NType = eNPCType.Canjes Then
            NpcAoMCanjes = val(i)
        End If
    Next i

    NumAoMCanjes = val(GetVar(App.Path & "\dat\AoMCanjes.dat", "INIT", "NumItems"))

    If NumAoMCanjes = 0 Then Exit Sub

    ReDim AoMCanjes(1 To MAX_INVENTORY_SLOTS)

    If NumAoMCanjes > 0 Then
        For i = 1 To NumAoMCanjes
            AoMCanjes(i).Name = GetVar(App.Path & "\dat\AoMCanjes.dat", "Item" & i, "Name")
            AoMCanjes(i).Index = val(GetVar(App.Path & "\dat\AoMCanjes.dat", "Item" & i, "Index"))
            AoMCanjes(i).Monedas = val(GetVar(App.Path & "\dat\AoMCanjes.dat", "Item" & i, "Monedas"))
            AoMCanjes(i).Cantidad = val(GetVar(App.Path & "\dat\AoMCanjes.dat", "Item" & i, "Cantidad"))
        Next i
    End If

End Sub

Public Sub IniciarComercioCanjes(ByVal UserIndex As Integer)
'Mandamos el Inventario
    Call EnviarCanjInv(UserIndex, UserList(UserIndex).flags.TargetNpc)
    'Hacemos un Update del inventario del usuario
    Call UpdateUserInv(True, UserIndex, 0)
    'Atcualizamos las monedas
    Call EnviarCanjes(UserIndex)
    'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
    UserList(UserIndex).flags.Comerciando = True
    SendData SendTarget.ToIndex, UserIndex, 0, "INITCANJ"
    Exit Sub
End Sub

Sub EnviarCanjInv(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'Enviamos el inventario del npc con el cual el user va a comerciar...
    Dim i As Integer
    Dim val As Long

    For i = 1 To MAX_INVENTORY_SLOTS

        If i <= NumAoMCanjes Then

            SendData SendTarget.ToIndex, UserIndex, 0, "NPCJ" & AoMCanjes(i).Name & "," & _
                                                       AoMCanjes(i).Index & "," & AoMCanjes(i).Monedas & "," & ObjData(AoMCanjes(i).Index).GrhIndex & "," & _
                                                       ObjData(AoMCanjes(i).Index).def & "," & ObjData(AoMCanjes(i).Index).MinHit & "," & ObjData(AoMCanjes(i).Index).MaxHit & "," & _
                                                       ObjData(AoMCanjes(i).Index).ObjType & "," & AoMCanjes(i).Cantidad

        Else
            SendData SendTarget.ToIndex, UserIndex, 0, "NPCJ" & "Nada" & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," _
                                                     & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0

        End If

    Next i

End Sub

Sub NpcVentaCanjes(ByVal UserIndex As Integer, ByVal i As Long, ByVal Cantidad As Integer, ByVal NpcIndex As Integer)
    Dim Slot As Integer
    Dim MiObj As Obj
    Dim Precio As Long

    With UserList(UserIndex)

        If Cantidad < 1 Then Exit Sub

        Call EnviarCanjes(UserIndex)

        If Cantidad > MAX_INVENTORY_OBJS Then
            Call SendData(SendTarget.ToAll, 0, 0, UserList(UserIndex).Name & " ha sido baneado por el sistema anti-cheats." & FONTTYPE_FIGHT)
            Call Ban(UserList(UserIndex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio " & Cantidad)
            UserList(UserIndex).flags.Ban = 1
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRHas sido baneado por el sistema anti cheats")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If

        Slot = 1
        Do Until .Invent.Object(Slot).ObjIndex = 0 Or .Invent.Object(Slot).ObjIndex = AoMCanjes(i).Index
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes tener mas objetos." & FONTTYPE_INFO)
                Exit Sub
            End If
        Loop

        Precio = AoMCanjes(i).Monedas * Cantidad

        If .AoMCanjes >= Precio Then
            Call LogCanjes("1", .Name, .AoMCanjes, Precio, ObjData(AoMCanjes(i).Index).Name)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Enhorabuena! Has comprado " & Cantidad & " " & AoMCanjes(i).Name & " por " & Precio & " AoMCanjes." & FONTTYPE_INFO)
            .AoMCanjes = .AoMCanjes - Precio
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahora tienes " & .AoMCanjes & " AoMCanjes." & FONTTYPE_INFO)
            AoMCanjes(i).Cantidad = AoMCanjes(i).Cantidad - Cantidad
            MiObj.Amount = Cantidad
            MiObj.ObjIndex = AoMCanjes(i).Index
            Call MeterItemEnInventario(UserIndex, MiObj)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes suficiente AoMCreditos." & FONTTYPE_WARNING)
            Exit Sub
        End If

    End With

    Call EnviarCanjes(UserIndex)
    Call UpdateVentanaCanjes(UserIndex)
    Call EnviarCanjInv(UserIndex, UserList(UserIndex).flags.TargetNpc)
End Sub

Public Sub UpdateVentanaCanjes(ByVal UserIndex As Integer)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "TRANSAJ")
End Sub

Public Sub EnviarCanjes(ByVal UserIndex As Integer)
    If UserList(UserIndex).AoMCanjes < 0 Then UserList(UserIndex).AoMCanjes = 0
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "CRJ" & UserList(UserIndex).AoMCanjes)
End Sub

Sub NPCCompraCanjes(ByVal UserIndex As Integer, ByVal i As Integer, ByVal Cantidad As Integer)

    Dim ObjIndex As Boolean
    Dim Id As Integer
    Dim Loopi As Long
    Dim Precio As Long

    With UserList(UserIndex)
        For Loopi = 1 To NumAoMCanjes
            If .Invent.Object(i).ObjIndex = AoMCanjes(Loopi).Index Then
                ObjIndex = True
                Id = Loopi
            End If
        Next Loopi

        If ObjIndex Then
            Call EnviarCanjes(UserIndex)
            If .Invent.Object(i).Amount >= Cantidad Then
                Precio = AoMCanjes(Id).Monedas * Cantidad
                Call LogCanjes("2", .Name, .AoMCanjes, Precio, ObjData(AoMCanjes(Id).Index).Name)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Enhorabuena! Has vendido " & Cantidad & " " & AoMCanjes(Id).Name & " por " & Precio & " AoMCanjes." & FONTTYPE_INFO)
                .AoMCanjes = .AoMCanjes + Precio
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahora tienes " & .AoMCanjes & " AoMCanjes." & FONTTYPE_INFO)
                Call QuitarObjetos(AoMCanjes(Id).Index, Cantidad, UserIndex)
                AoMCanjes(Id).Cantidad = AoMCanjes(Id).Cantidad + Cantidad
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes tanta cantidad de item." & FONTTYPE_INFO)
            End If

        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tengo interés de comprar ese item." & FONTTYPE_INFO)
        End If
    End With

    Call EnviarCanjes(UserIndex)
    Call UpdateVentanaCanjes(UserIndex)
    Call EnviarCanjInv(UserIndex, UserList(UserIndex).flags.TargetNpc)
End Sub
