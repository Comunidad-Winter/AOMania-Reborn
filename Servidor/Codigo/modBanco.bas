Attribute VB_Name = "modBanco"
Option Explicit

'MODULO PROGRAMADO POR NEB
'Kevin Birmingham
'kbneb@hotmail.com

Sub IniciarDeposito(ByVal Userindex As Integer)

    On Error GoTo errhandler

    'Hacemos un Update del inventario del usuario
    Call UpdateBanUserInv(True, Userindex, 0)
    'Atcualizamos el dinero
    Call EnviarOro(Userindex)
    'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
    SendData SendTarget.ToIndex, Userindex, 0, "INITBANCO"
    UserList(Userindex).flags.Comerciando = True

errhandler:

End Sub

Sub SendBanObj(Userindex As Integer, Slot As Byte, Object As UserOBJ)

    UserList(Userindex).BancoInvent.Object(Slot) = Object

    If Object.ObjIndex > 0 Then

        Call SendData(SendTarget.ToIndex, Userindex, 0, "SBO" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & _
                Object.Amount & "," & ObjData(Object.ObjIndex).GrhIndex & "," & ObjData(Object.ObjIndex).ObjType & "," & ObjData( _
                Object.ObjIndex).MaxHit & "," & ObjData(Object.ObjIndex).MinHit & "," & ObjData(Object.ObjIndex).MaxDef & "," & ObjData( _
                Object.ObjIndex).MinDef)

    Else

        Call SendData(SendTarget.ToIndex, Userindex, 0, "SBO" & Slot & ",0," & "(Vacío)" & ",0,0,0")

    End If

End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal Userindex As Integer, ByVal Slot As Byte)

    Dim NullObj As UserOBJ
    Dim LoopC   As Byte

    'Actualiza un solo slot
    If Not UpdateAll Then

        'Actualiza el inventario
        If UserList(Userindex).BancoInvent.Object(Slot).ObjIndex > 0 Then
            Call SendBanObj(Userindex, Slot, UserList(Userindex).BancoInvent.Object(Slot))
        Else
            Call SendBanObj(Userindex, Slot, NullObj)

        End If

    Else

        'Actualiza todos los slots
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS

            'Actualiza el inventario
            If UserList(Userindex).BancoInvent.Object(LoopC).ObjIndex > 0 Then
                Call SendBanObj(Userindex, LoopC, UserList(Userindex).BancoInvent.Object(LoopC))
            Else
            
                Call SendBanObj(Userindex, LoopC, NullObj)
            
            End If

        Next LoopC

    End If

End Sub

Sub UserRetiraItem(ByVal Userindex As Integer, ByVal i As Integer, ByVal Cantidad As Integer)

    On Error GoTo errhandler

    If Cantidad < 1 Then Exit Sub
   
    If UserList(Userindex).BancoInvent.Object(i).Amount > 0 Then
        If Cantidad > UserList(Userindex).BancoInvent.Object(i).Amount Then Cantidad = UserList(Userindex).BancoInvent.Object(i).Amount
        'Agregamos el obj que compro al inventario
        Call UserReciveObj(Userindex, CInt(i), Cantidad)
        'Actualizamos el inventario del usuario
        Call UpdateUserInv(True, Userindex, 0)
        'Actualizamos el banco
        Call UpdateBanUserInv(True, Userindex, 0)
        'Actualizamos la ventana de comercio
        Call UpdateVentanaBanco(i, 0, Userindex)

    End If

errhandler:

End Sub

Sub UserReciveObj(ByVal Userindex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

    Dim Slot As Integer
    Dim obji As Integer

    If UserList(Userindex).BancoInvent.Object(ObjIndex).Amount <= 0 Then Exit Sub

    obji = UserList(Userindex).BancoInvent.Object(ObjIndex).ObjIndex

    '¿Ya tiene un objeto de este tipo?
    Slot = 1

    Do Until UserList(Userindex).Invent.Object(Slot).ObjIndex = obji And UserList(Userindex).Invent.Object(Slot).Amount + Cantidad <= _
            MAX_INVENTORY_OBJS
    
        Slot = Slot + 1

        If Slot > MAX_INVENTORY_SLOTS Then
            Exit Do

        End If

    Loop

    'Sino se fija por un slot vacio
    If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1

        Do Until UserList(Userindex).Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "||No puedes tener mas objetos." & FONTTYPE_INFO)
                Exit Sub

            End If

        Loop
        UserList(Userindex).Invent.NroItems = UserList(Userindex).Invent.NroItems + 1

    End If

    'Mete el obj en el slot
    If UserList(Userindex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
        'Menor que MAX_INV_OBJS
        UserList(Userindex).Invent.Object(Slot).ObjIndex = obji
        UserList(Userindex).Invent.Object(Slot).Amount = UserList(Userindex).Invent.Object(Slot).Amount + Cantidad
    
        Call QuitarBancoInvItem(Userindex, CByte(ObjIndex), Cantidad)
    Else
        Call SendData(SendTarget.ToIndex, Userindex, 0, "||No puedes tener mas objetos." & FONTTYPE_INFO)

    End If

End Sub

Sub QuitarBancoInvItem(ByVal Userindex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)

    Dim ObjIndex As Integer
    ObjIndex = UserList(Userindex).BancoInvent.Object(Slot).ObjIndex

    'Quita un Obj

    UserList(Userindex).BancoInvent.Object(Slot).Amount = UserList(Userindex).BancoInvent.Object(Slot).Amount - Cantidad
        
    If UserList(Userindex).BancoInvent.Object(Slot).Amount <= 0 Then
        UserList(Userindex).BancoInvent.NroItems = UserList(Userindex).BancoInvent.NroItems - 1
        UserList(Userindex).BancoInvent.Object(Slot).ObjIndex = 0
        UserList(Userindex).BancoInvent.Object(Slot).Amount = 0

    End If
    
End Sub

Sub UpdateVentanaBanco(ByVal Slot As Integer, ByVal NpcInv As Byte, ByVal Userindex As Integer)
 
    Call SendData(SendTarget.ToIndex, Userindex, 0, "BANCOOK" & Slot & "," & NpcInv)
 
End Sub

Sub UserDepositaItem(ByVal Userindex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)

    On Error GoTo errhandler

    'El usuario deposita un item
   
    If UserList(Userindex).Invent.Object(Item).Amount > 0 And UserList(Userindex).Invent.Object(Item).Equipped = 0 Then
            
        If Cantidad > 0 And Cantidad > UserList(Userindex).Invent.Object(Item).Amount Then Cantidad = UserList(Userindex).Invent.Object(Item).Amount
        'Agregamos el obj que compro al inventario
        Call UserDejaObj(Userindex, CInt(Item), Cantidad)
        'Actualizamos el inventario del usuario
        Call UpdateUserInv(True, Userindex, 0)
        'Actualizamos el inventario del banco
        Call UpdateBanUserInv(True, Userindex, 0)
        'Actualizamos la ventana del banco
            
        Call UpdateVentanaBanco(Item, 1, Userindex)
            
    End If

errhandler:

End Sub

Sub UserDejaObj(ByVal Userindex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

    Dim Slot As Integer
    Dim obji As Integer

    If Cantidad < 1 Then Exit Sub

    obji = UserList(Userindex).Invent.Object(ObjIndex).ObjIndex

    '¿Ya tiene un objeto de este tipo?
    Slot = 1

    Do Until UserList(Userindex).BancoInvent.Object(Slot).ObjIndex = obji And UserList(Userindex).BancoInvent.Object(Slot).Amount + Cantidad <= _
            MAX_INVENTORY_OBJS
        Slot = Slot + 1
        
        If Slot > MAX_BANCOINVENTORY_SLOTS Then
            Exit Do

        End If

    Loop

    'Sino se fija por un slot vacio antes del slot devuelto
    If Slot > MAX_BANCOINVENTORY_SLOTS Then
        Slot = 1

        Do Until UserList(Userindex).BancoInvent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "||No tienes mas espacio en el banco!!" & FONTTYPE_INFO)
                Exit Sub
                Exit Do

            End If

        Loop

        If Slot <= MAX_BANCOINVENTORY_SLOTS Then UserList(Userindex).BancoInvent.NroItems = UserList(Userindex).BancoInvent.NroItems + 1
        
    End If

    If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido

        'Mete el obj en el slot
        If UserList(Userindex).BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        
            'Menor que MAX_INV_OBJS
            UserList(Userindex).BancoInvent.Object(Slot).ObjIndex = obji
            UserList(Userindex).BancoInvent.Object(Slot).Amount = UserList(Userindex).BancoInvent.Object(Slot).Amount + Cantidad
        
            Call QuitarUserInvItem(Userindex, CByte(ObjIndex), Cantidad)

        Else
            Call SendData(SendTarget.ToIndex, Userindex, 0, "||El banco no puede cargar tantos objetos." & FONTTYPE_INFO)

        End If

    Else
        Call QuitarUserInvItem(Userindex, CByte(ObjIndex), Cantidad)

    End If

End Sub

Sub SendUserBovedaTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)

    On Error Resume Next

    Dim j As Integer
  Call SendData(SendTarget.ToIndex, sendIndex, 0, "||" & UserList(Userindex).Name & " Tiene " & UserList(Userindex).BancoInvent.NroItems & " objetos." & FONTTYPE_INFO)

    For j = 1 To MAX_BANCOINVENTORY_SLOTS

        If UserList(Userindex).BancoInvent.Object(j).ObjIndex > 0 Then
            Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| Objeto " & j & ":  """ & ObjData(UserList(Userindex).BancoInvent.Object( _
                    j).ObjIndex).Name & """ (ObjIndex: " & UserList(Userindex).BancoInvent.Object(j).ObjIndex & ")" & " Cantidad: " & UserList(Userindex).BancoInvent.Object(j).Amount & FONTTYPE_INFO)

        End If

    Next

End Sub

Sub SendUserBovedaTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)

    On Error Resume Next

    Dim j        As Integer
    Dim CharFile As String, Tmp As String
    Dim ObjInd   As Long, ObjCant As Long

    CharFile = CharPath & CharName & ".chr"

    If FileExist(CharFile, vbNormal) Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||" & CharName & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| Tiene " & GetVar(CharFile, "BancoInventory", "CantidadItems") & " objetos." & _
                FONTTYPE_INFO)

        For j = 1 To MAX_BANCOINVENTORY_SLOTS
            Tmp = GetVar(CharFile, "BancoInventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))

            If ObjInd > 0 Then
                Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant & _
                        FONTTYPE_INFO)

            End If

        Next
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Usuario inexistente: " & CharName & FONTTYPE_INFO)

    End If

End Sub

