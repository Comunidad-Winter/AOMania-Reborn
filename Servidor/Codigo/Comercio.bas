Attribute VB_Name = "Comercio"

Option Explicit

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%          MODULO DE COMERCIO NPC-USER              %%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Function UserCompraObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal NpcIndex As Integer, ByVal Cantidad As Integer) As Boolean

    On Error GoTo errorh

    Dim infla As Long
    Dim Descuento As String
    Dim Unidad As Long, monto As Long
    Dim Slot As Integer
    Dim obji As Integer
    Dim Encontre As Boolean

    UserCompraObj = False

    If (Npclist(UserList(UserIndex).flags.TargetNpc).Invent.Object(ObjIndex).Amount <= 0) Then Exit Function

    obji = Npclist(UserList(UserIndex).flags.TargetNpc).Invent.Object(ObjIndex).ObjIndex

    'es una armadura real y el tipo no es faccion?
    'If ObjData(obji).Real = 1 Then
    '    If Npclist(NpcIndex).Name <> "SR" Then
    '        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbWhite & "°" & _
             '                "Lo siento, la ropa faccionaria solo es para muestra, no tengo autorización para venderla. Diríjete al sastre de tu ejército." _
             '                & "°" & str(Npclist(NpcIndex).char.CharIndex))
    '        Exit Function
    '
    '        End If
    '
    '    End If

    'If ObjData(obji).Caos = 1 Then
    ' If Npclist(NpcIndex).Name <> "SC" Then
    '     Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbWhite & "°" & _
          "Lo siento, la ropa faccionaria solo es para muestra, no tengo autorización para venderla. Diríjete al sastre de tu ejército." _
          & "°" & str(Npclist(NpcIndex).char.CharIndex))
    '   Exit Function

    '    End If

    '   End If

    '¿Ya tiene un objeto de este tipo?
    Slot = 1

    Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = obji And UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad <= _
       MAX_INVENTORY_OBJS

        Slot = Slot + 1

        If Slot > MAX_INVENTORY_SLOTS Then
            Exit Do

        End If

    Loop

    'Sino se fija por un slot vacio
    If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1

        Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes tener más objetos." & FONTTYPE_INFO)
                Exit Function
            End If

        Loop
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1

    End If

    'desde aca para abajo se realiza la transaccion
    UserCompraObj = True

    'Mete el obj en el slot
    If UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        'Menor que MAX_INV_OBJS
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = obji
        UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad

        'Le sustraemos el valor en oro del obj comprado

        infla = (Npclist(NpcIndex).Inflacion * ObjData(obji).Valor) / 100
        Descuento = UserList(UserIndex).flags.Descuento

        If Descuento = 0 Then Descuento = 1    'evitamos dividir por 0!
        If Npclist(NpcIndex).Numero = 265 Then
            Unidad = ObjData(Npclist(NpcIndex).Invent.Object(ObjIndex).ObjIndex).Valor
        Else
            If Npclist(NpcIndex).Invent.Object(ObjIndex).ObjIndex = 1944 Then
              Unidad = ObjData(Npclist(NpcIndex).Invent.Object(ObjIndex).ObjIndex).Valor
            Else
            Unidad = ((ObjData(Npclist(NpcIndex).Invent.Object(ObjIndex).ObjIndex).Valor + infla) / Descuento)
            End If
        End If
        monto = Unidad * Cantidad
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - monto


        'tal vez suba el skill comerciar ;-)
        Call SubirSkill(UserIndex, comerciar)
        Call EnviarOro(UserIndex)

        If ObjData(obji).ObjType = eOBJType.otLlaves Then Call logVentaCasa(UserList(UserIndex).Name & " compro " & ObjData(obji).Name)

        Call QuitarNpcInvItem(UserList(UserIndex).flags.TargetNpc, CByte(ObjIndex), Cantidad)
    Else
        Call SendData(SendTarget.toindex, UserIndex, 0, "||No puedes tener más objetos." & FONTTYPE_INFO)

    End If

    Exit Function

errorh:
    Call LogError("Error en USERCOMPRAOBJ. " & err.Description)

End Function

Sub NpcCompraObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

    On Error GoTo errorh

    Dim Slot As Integer
    Dim obji As Integer
    Dim NpcIndex As Integer
    Dim infla As Long
    Dim monto As Long

    If Cantidad < 1 Then Exit Sub

    NpcIndex = UserList(UserIndex).flags.TargetNpc
    obji = UserList(UserIndex).Invent.Object(ObjIndex).ObjIndex

    If ObjData(obji).Newbie = 1 Then
        Call SendData(SendTarget.toindex, UserIndex, 0, "||No comercio objetos para newbies." & FONTTYPE_INFO)
        Exit Sub

    End If

    If Npclist(NpcIndex).TipoItems <> eOBJType.otCualquiera Then

        '¿Son los items con los que comercia el npc?
        If Npclist(NpcIndex).TipoItems <> ObjData(obji).ObjType Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||El npc no esta interesado en comprar ese objeto." & FONTTYPE_WARNING)
            Exit Sub

        End If

    End If

    If obji = iORO Then
        Call SendData(SendTarget.toindex, UserIndex, 0, "||El npc no esta interesado en comprar ese objeto." & FONTTYPE_WARNING)
        Exit Sub

    End If

    '¿Ya tiene un objeto de este tipo?
    Slot = 1

    Do Until (Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = obji And Npclist(NpcIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS)

        Slot = Slot + 1

        If Slot > MAX_INVENTORY_SLOTS Then Exit Do
    Loop

    'Sino se fija por un slot vacio antes del slot devuelto
    If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1

        Do Until Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then Exit Do
        Loop

        If Slot <= MAX_INVENTORY_SLOTS Then Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1

    End If

    If Slot <= MAX_INVENTORY_SLOTS Then    'Slot valido
        'Mete el obj en el slot
        Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = obji

        If Npclist(NpcIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
            Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount + Cantidad
        Else
            Npclist(NpcIndex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS

        End If

    End If

    Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)

    'Le sumamos al user el valor en oro del obj vendido

    If Npclist(NpcIndex).Numero = 265 Then
        If ObjData(obji).ObjType = eOBJType.otPLATA Then
            monto = ((ObjData(obji).Valor) * Cantidad)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + monto
        Else
            monto = ((ObjData(obji).Valor) * Cantidad)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + monto
        End If
    Else
    
        If ObjData(obji).ObjType = eOBJType.otPLATA Then
            monto = ((ObjData(obji).Valor \ 2 + infla) * Cantidad)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + monto
        ElseIf ObjData(obji).ObjType = eOBJType.otCheques Then
              monto = ((ObjData(obji).Valor) * Cantidad)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + monto
        Else
            monto = ((ObjData(obji).Valor \ 3 + infla) * Cantidad)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + monto
        End If
    End If

    If UserList(UserIndex).Stats.GLD > MaxOro Then UserList(UserIndex).Stats.GLD = MaxOro

    'tal vez suba el skill comerciar ;-)
    Call SubirSkill(UserIndex, comerciar)
    Call EnviarOro(UserIndex)
    Exit Sub

errorh:
    Call LogError("Error en NPCCOMPRAOBJ. " & err.Description)

End Sub

Sub IniciarCOmercioNPC(ByVal UserIndex As Integer)

    On Error GoTo errhandler

    'Mandamos el Inventario
    Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNpc)
    'Hacemos un Update del inventario del usuario
    Call UpdateUserInv(True, UserIndex, 0)
    'Atcualizamos el dinero
    Call EnviarOro(UserIndex)
    'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
    UserList(UserIndex).flags.Comerciando = True
    SendData SendTarget.toindex, UserIndex, 0, "INITCOM"
    Exit Sub

errhandler:
    Dim str As String
    str = "Error en IniciarComercioNPC. UI=" & UserIndex

    If UserIndex > 0 Then
        str = str & ".Nombre: " & UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " comerciando con "

        If UserList(UserIndex).flags.TargetNpc > 0 Then
            str = str & Npclist(UserList(UserIndex).flags.TargetNpc).Name
        Else
            str = str & "<NPCINDEX 0>"

        End If

    Else
        str = str & "<USERINDEX 0>"

    End If

End Sub

Sub NPCVentaItem(ByVal UserIndex As Integer, ByVal i As Integer, ByVal Cantidad As Integer, ByVal NpcIndex As Integer)

'listindex+1, cantidad
    On Error GoTo errhandler

    Dim infla As Long
    Dim val As Long
    Dim Desc As String

    If Cantidad < 1 Then Exit Sub

    'NPC VENDE UN OBJ A UN USUARIO
    Call EnviarOro(UserIndex)

    If i > MAX_INVENTORY_SLOTS Then
        Call SendData(SendTarget.ToAdmins, 0, 0, "Posible intento de romper el sistema de comercio. Usuario: " & UserList(UserIndex).Name & _
                                                 FONTTYPE_WARNING)
        Exit Sub
    End If

    If Cantidad > MAX_INVENTORY_OBJS Then
        Call SendData(SendTarget.ToAll, 0, 0, UserList(UserIndex).Name & " ha sido baneado por el sistema anti-cheats." & FONTTYPE_FIGHT)
        Call Ban(UserList(UserIndex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio " & Cantidad)
        UserList(UserIndex).flags.Ban = 1
        Call SendData(SendTarget.toindex, UserIndex, 0, "ERRHas sido baneado por el sistema anti cheats")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If

    'Calculamos el valor unitario


    infla = (Npclist(NpcIndex).Inflacion * ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).Valor) / 100
    Desc = Descuento(UserIndex)

    If Desc = 0 Then Desc = 1    'evitamos dividir por 0!

    If Npclist(NpcIndex).Numero = 265 Then
        val = ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).Valor
    Else
      If Npclist(NpcIndex).Invent.Object(i).ObjIndex = 1944 Then
         val = ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).Valor
         Else
        val = (ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).Valor + infla) / Desc
     End If
    End If

    If UserList(UserIndex).Stats.GLD >= (val * Cantidad) Then
        If Npclist(UserList(UserIndex).flags.TargetNpc).Invent.Object(i).Amount > 0 Then
            If Cantidad > Npclist(UserList(UserIndex).flags.TargetNpc).Invent.Object(i).Amount Then Cantidad = Npclist(UserList( _
                                                                                                                       UserIndex).flags.TargetNpc).Invent.Object(i).Amount

            'Agregamos el obj que compro al inventario
            If Not UserCompraObj(UserIndex, CInt(i), UserList(UserIndex).flags.TargetNpc, Cantidad) Then
                'Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes comprar este ítem." & FONTTYPE_INFO)
            End If

            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            'Actualizamos el oro
            Call EnviarOro(UserIndex)
            'Actualizamos la ventana de comercio
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNpc)
            Call UpdateVentanaComercio(i, 0, UserIndex)

        End If

    Else
        Call SendData(SendTarget.toindex, UserIndex, 0, "||No tienes suficiente dinero." & FONTTYPE_INFO)

    End If

    Exit Sub

errhandler:
    Call LogError("Error en comprar item: " & err.Description)

End Sub

Sub NPCCompraItem(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer, Optional ByVal NpcDrag As Integer = 0)

    On Error GoTo errhandler

    Dim NpcIndex As Integer

    If NpcDrag > 0 Then
        NpcIndex = NpcDrag
    Else
        NpcIndex = UserList(UserIndex).flags.TargetNpc

    End If

    'Si es una armadura faccionaria vemos que la está intentando vender al sastre
    If ObjData(UserList(UserIndex).Invent.Object(Item).ObjIndex).Real = 1 And ObjData(UserList(UserIndex).Invent.Object(Item).ObjIndex).ObjType = otArmadura Then
        If Npclist(NpcIndex).Name <> "SR" Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||Las armaduras faccionarias sólo las puedes vender a sus respectivos Sastres" & _
                                                            FONTTYPE_WARNING)

            'Actualizamos la ventana de comercio
            Call UpdateVentanaComercio(Item, 1, UserIndex)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNpc)
            Exit Sub

        End If

    ElseIf ObjData(UserList(UserIndex).Invent.Object(Item).ObjIndex).Caos = 1 And ObjData(UserList(UserIndex).Invent.Object(Item).ObjIndex).ObjType = otArmadura Then

        If Npclist(NpcIndex).Name <> "SC" Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||Las armaduras faccionarias sólo las puedes vender a sus respectivos Sastres" & _
                                                            FONTTYPE_WARNING)

            'Actualizamos la ventana de comercio
            Call UpdateVentanaComercio(Item, 1, UserIndex)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNpc)
            Exit Sub
        End If

    ElseIf ObjData(UserList(UserIndex).Invent.Object(Item).ObjIndex).Nemes = 1 And ObjData(UserList(UserIndex).Invent.Object(Item).ObjIndex).ObjType = otArmadura Then

        If Npclist(NpcIndex).Name <> "SC" Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||Las armaduras faccionarias sólo las puedes vender a sus respectivos Sastres" & _
                                                            FONTTYPE_WARNING)

            'Actualizamos la ventana de comercio
            Call UpdateVentanaComercio(Item, 1, UserIndex)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNpc)
            Exit Sub
        End If

    ElseIf ObjData(UserList(UserIndex).Invent.Object(Item).ObjIndex).Templ = 1 And ObjData(UserList(UserIndex).Invent.Object(Item).ObjIndex).ObjType = otArmadura Then

        If Npclist(NpcIndex).Name <> "SC" Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "||Las armaduras faccionarias sólo las puedes vender a sus respectivos Sastres" & _
                                                            FONTTYPE_WARNING)

            'Actualizamos la ventana de comercio
            Call UpdateVentanaComercio(Item, 1, UserIndex)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNpc)
            Exit Sub
        End If

    End If

    'NPC COMPRA UN OBJ A UN USUARIO
    Call EnviarOro(UserIndex)

    If UserList(UserIndex).Invent.Object(Item).Amount > 0 And UserList(UserIndex).Invent.Object(Item).Equipped = 0 Then

        If Cantidad > 0 And Cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Item).Amount
        'Agregamos el obj que compro al inventario
        Call NpcCompraObj(UserIndex, CInt(Item), Cantidad)
        'Actualizamos el inventario del usuario
        Call UpdateUserInv(True, UserIndex, 0)
        'Actualizamos el oro
        Call EnviarOro(UserIndex)

        Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNpc)
        'Actualizamos la ventana de comercio
        Call UpdateVentanaComercio(Item, 1, UserIndex)

    End If

    Exit Sub

errhandler:
    Call LogError("Error en vender item: " & err.Description)

End Sub

Sub UpdateVentanaComercio(ByVal Slot As Integer, ByVal NpcInv As Byte, ByVal UserIndex As Integer)
    Call SendData(SendTarget.toindex, UserIndex, 0, "TRANSOK" & Slot & "," & NpcInv)

End Sub

Function Descuento(ByVal UserIndex As Integer) As String

'Establece el descuento en funcion del skill comercio
Dim PtsComercio As Integer
PtsComercio = UserList(UserIndex).Stats.UserSkills(comerciar)

If PtsComercio <= 11 And PtsComercio > 5 Then
    UserList(UserIndex).flags.Descuento = 1.1
    Descuento = 1.1
ElseIf PtsComercio <= 20 And PtsComercio >= 12 Then
    UserList(UserIndex).flags.Descuento = 1.2
    Descuento = 1.2
ElseIf PtsComercio <= 30 And PtsComercio >= 19 Then
    UserList(UserIndex).flags.Descuento = 1.3
    Descuento = 1.3
ElseIf PtsComercio <= 40 And PtsComercio >= 29 Then
    UserList(UserIndex).flags.Descuento = 1.4
    Descuento = 1.4
ElseIf PtsComercio <= 50 And PtsComercio >= 39 Then
    UserList(UserIndex).flags.Descuento = 1.5
    Descuento = 1.5
ElseIf PtsComercio <= 60 And PtsComercio >= 49 Then
    UserList(UserIndex).flags.Descuento = 1.6
    Descuento = 1.6
ElseIf PtsComercio <= 70 And PtsComercio >= 59 Then
    UserList(UserIndex).flags.Descuento = 1.7
    Descuento = 1.7
ElseIf PtsComercio <= 80 And PtsComercio >= 69 Then
    UserList(UserIndex).flags.Descuento = 1.8
    Descuento = 1.8
ElseIf PtsComercio <= 99 And PtsComercio >= 79 Then
    UserList(UserIndex).flags.Descuento = 1.9
    Descuento = 1.9
ElseIf PtsComercio <= 999999 And PtsComercio >= 99 Then
    UserList(UserIndex).flags.Descuento = 2
    Descuento = 2
Else
    UserList(UserIndex).flags.Descuento = 0
    Descuento = 0
End If

End Function

Sub EnviarNpcInv(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    'Enviamos el inventario del npc con el cual el user va a comerciar...
    Dim i     As Integer

    Dim infla As Long

    Dim Desc  As String

    Dim val   As Long

    Desc = Descuento(UserIndex)

    If Desc = 0 Then Desc = 1    'evitamos dividir por 0!

    For i = 1 To MAX_INVENTORY_SLOTS

        If Npclist(NpcIndex).Invent.Object(i).ObjIndex > 0 Then

            'Calculamos el porc de inflacion del npc
            If Npclist(NpcIndex).Numero = 265 Then
                val = ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).Valor
            Else
                If Npclist(NpcIndex).Invent.Object(i).ObjIndex = 1944 Then
                    val = ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).Valor
                Else
                infla = (Npclist(NpcIndex).Inflacion * ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).Valor) / 100
                val = (ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).Valor + infla) / Desc
                End If

            End If

            SendData SendTarget.toindex, UserIndex, 0, "NPCI" & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).Name & "," & Npclist(NpcIndex).Invent.Object(i).Amount & "," & val & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).GrhIndex & "," & Npclist(NpcIndex).Invent.Object(i).ObjIndex & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).ObjType & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).MaxHit & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).MinHit & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).MaxDef
        Else
            SendData SendTarget.toindex, UserIndex, 0, "NPCI" & "Nada" & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0

        End If

    Next i

End Sub

Sub IniciarComercioSagradoNPC(ByVal UserIndex As Integer)

    On Error GoTo errhandler

    Call EnviaSagradoNpcInv(UserIndex)
    Call UpdateUserInv(True, UserIndex, 0)
    UserList(UserIndex).flags.Comerciando = True
    SendData SendTarget.toindex, UserIndex, 0, "INITSAG"

    Exit Sub

errhandler:

    Call LogError("Error en IniciarComercioSagradoNPC")

End Sub

Sub EnviaSagradoNpcInv(ByVal UserIndex As Integer)
          
      Dim ObjIndex As Integer
      Dim i As Integer
      
      For i = 1 To MAX_INVENTORY_SLOTS
         If i <= NumCanjeSagrado Then
          ObjIndex = val(CanjeSagrado(i).Index)
           Else
           ObjIndex = 0
         End If
          
          If ObjIndex > 0 Then
            SendData SendTarget.toindex, UserIndex, 0, "NPCI" & ObjData(ObjIndex).Name & "," & "10000" & "," & CanjeSagrado(i).Valor & "," & ObjData(ObjIndex).GrhIndex & "," & ObjIndex & "," & ObjData(ObjIndex).ObjType & "," & ObjData(ObjIndex).MaxHit & "," & ObjData(ObjIndex).MinHit & "," & ObjData(ObjIndex).MaxDef
          Else
          
             SendData SendTarget.toindex, UserIndex, 0, "NPCI" & "Nada" & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0
             
          End If
          
      Next i
      
          
End Sub

Sub NPCVentaItemSagrado(ByVal UserIndex As Integer, _
                        ByVal i As Integer, _
                        ByVal Cantidad As Integer, _
                        ByVal NpcIndex As Integer)
                        
    Dim Unidad As Integer
    Dim LooPc As Integer
    Dim Obj As Obj
           
    If Cantidad < 1 Then Exit Sub
      


    If Cantidad > MAX_INVENTORY_OBJS Then
        Call SendData(SendTarget.ToAll, 0, 0, UserList(UserIndex).Name & " ha sido baneado por el sistema anti-cheats." & FONTTYPE_FIGHT)
        Call Ban(UserList(UserIndex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio " & Cantidad)
        UserList(UserIndex).flags.Ban = 1
        Call SendData(SendTarget.toindex, UserIndex, 0, "ERRHas sido baneado por el sistema anti cheats")
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
    
    For LooPc = 1 To NumCanjeSagrado
         
        If CanjeSagrado(LooPc).Index = i Then
            Unidad = CanjeSagrado(LooPc).Valor
        End If
         
    Next LooPc
    
    Unidad = Unidad * Cantidad
    
    If Not TieneObjetos(TicketSagrado, Unidad, UserIndex) Then
        Call SendData(toindex, UserIndex, 0, "||No tienes suficientes vales!" & FONTTYPE_WARNING)
        Exit Sub
    End If
    
    Obj.ObjIndex = i
    Obj.Amount = Cantidad
    
    Call MeterItemEnInventario(UserIndex, Obj)
    Call QuitarObjetos(TicketSagrado, Unidad, UserIndex)
      
End Sub
