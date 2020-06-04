Attribute VB_Name = "ModSubasta"
Option Explicit

Private Const Max_Subasta As Integer = 100
Public Cant_Subasta As Integer
Private Const OroCrearSubasta As Integer = 100
Public SavingSubasta As Boolean

Public Type tSubasta
     
     Subastador As String
     Objeto As Integer
     Cantidad As Integer
     Valor As Long
     Timer As Long
     Comprador As String
     
End Type

Public Subasta(1 To Max_Subasta) As tSubasta

Public bancoObj(1 To MAX_BANCOINVENTORY_SLOTS) As SubastaBanco

Public Type SubastaBanco
    ObjIndex As Integer
    Amount As Integer
End Type

Public Sub CargarSubastas()
     
     Dim LoopC As Long
     Dim Leer As New clsIniManager
     
     Call Leer.Initialize(DatPath & "\Subastas.dat")
     
     Cant_Subasta = Leer.GetValue("INIT", "NumSubasta")
     
     For LoopC = 1 To Cant_Subasta
              
         Subasta(LoopC).Subastador = CStr(Leer.GetValue("Subasta" & LoopC, "Subastador"))
         Subasta(LoopC).Objeto = CLng(Leer.GetValue("Subasta" & LoopC, "Objeto"))
         Subasta(LoopC).Cantidad = CLng(Leer.GetValue("Subasta" & LoopC, "Cantidad"))
         Subasta(LoopC).Valor = CLng(Leer.GetValue("Subasta" & LoopC, "Valor"))
         Subasta(LoopC).Timer = CLng(Leer.GetValue("Subasta" & LoopC, "Timer"))
         Subasta(LoopC).Comprador = CStr(Leer.GetValue("Subasta" & LoopC, "Comprador"))
              
     Next LoopC
     
     Set Leer = Nothing
     
End Sub

Public Sub GuardarSubastas()
        
     Dim Cant As Integer
     Dim LoopC As Long
     Dim Leer As clsIniManager
     Dim Round As Integer
     
     Set Leer = New clsIniManager
     
     If FileExist(DatPath & "\Subastas.dat", vbNormal) Then Call Kill(DatPath & "\Subastas.dat")
     
     For LoopC = 1 To Max_Subasta
        
        If Subasta(LoopC).Subastador <> "" And Cant_Subasta > 0 Then
            Round = Cant + 1
            If Cant_Subasta >= Round Then
            Cant = Cant + 1
            Call Leer.ChangeValue("INIT", "NumSubasta", Cant)
            Call Leer.ChangeValue("Subasta" & Cant, "Subastador", Subasta(LoopC).Subastador)
            Call Leer.ChangeValue("Subasta" & Cant, "Objeto", Subasta(LoopC).Objeto)
            Call Leer.ChangeValue("Subasta" & Cant, "Cantidad", Subasta(LoopC).Cantidad)
            Call Leer.ChangeValue("Subasta" & Cant, "Valor", Subasta(LoopC).Valor)
            Call Leer.ChangeValue("Subasta" & Cant, "Timer", Subasta(LoopC).Timer)
            Call Leer.ChangeValue("Subasta" & Cant, "Comprador", Subasta(LoopC).Comprador)
            End If
        End If
        
     Next LoopC
     
     If Cant = 0 Then
         Call Leer.ChangeValue("INIT", "NumSubasta", Cant)
     End If
     
     Call Leer.DumpFile(DatPath & "\Subastas.dat")
     
     Set Leer = Nothing
     
     Call CargarSubastas
        
End Sub

Public Sub IniciarVentanaSubasta(ByVal UserIndex As Integer)

    Call EnviaListSubasta(UserIndex)
    
    Call SendData(ToIndex, UserIndex, 0, "INITSUB")
       
End Sub

Public Sub EnviaListSubasta(ByVal UserIndex As Integer)
     
     Dim LoopC As Long
     
     Call SendData(ToIndex, UserIndex, 0, "RESETSB" & Cant_Subasta)
     
     If Cant_Subasta > 0 Then
         
         For LoopC = 1 To Cant_Subasta
                
                Call SendData(ToIndex, UserIndex, 0, "PAQSUBS" & Subasta(LoopC).Objeto & "," & ObjData(Subasta(LoopC).Objeto).Name & "," & Subasta(LoopC).Cantidad & "," & Subasta(LoopC).Valor & "," & _
                         Subasta(LoopC).Subastador & "," & VerTimerSubasta(Subasta(LoopC).Timer) & "," & Subasta(LoopC).Comprador & "," & ObjData(Subasta(LoopC).Objeto).GrhIndex)
                
         Next LoopC
         
     End If
     
End Sub

Public Sub CrearSubasta(ByVal UserIndex As Integer, ByVal Objeto As Integer, ByVal Cantidad As Integer, ByVal Precio As Long, ByVal Timer As String)
        
        If UserList(UserIndex).Stats.GLD < OroCrearSubasta Then
             Call SendData(ToIndex, UserIndex, 0, "||No tienes suficiente oro para poder crear una subasta." & FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Not TieneObjetos(Objeto, Cantidad, UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes suficiente objeto para poder crear una subasta." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If ObjData(Objeto).Real = 1 Or ObjData(Objeto).Caos = 1 Or ObjData(Objeto).Nemes = 1 Or ObjData(Objeto).Templ = 1 Then
             Call SendData(ToIndex, UserIndex, 0, "||No puedes subastar un objeto de la faccion." & FONTTYPE_INFO)
             Exit Sub
        End If
        
        If ObjData(Objeto).sagrado = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes subastar un objeto sagrado." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).Invent.EscudoEqpObjIndex = Objeto Or UserList(UserIndex).Invent.AlaEqpObjIndex = Objeto Or _
            UserList(UserIndex).Invent.AmuletoEqpObjIndex = Objeto Or UserList(UserIndex).Invent.ArmourEqpObjIndex = Objeto Or _
            UserList(UserIndex).Invent.CascoEqpObjIndex = Objeto Or _
            UserList(UserIndex).Invent.CascoEqpObjIndex = Objeto Or UserList(UserIndex).Invent.HerramientaEqpObjIndex = Objeto Or _
            UserList(UserIndex).Invent.MunicionEqpObjIndex = Objeto Or UserList(UserIndex).Invent.WeaponEqpObjIndex = Objeto Then
            Call SendData(ToIndex, UserIndex, 0, "||Debes desequiparte el objeto para poder subastarlo." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        Cant_Subasta = Cant_Subasta + 1
        
        Subasta(Cant_Subasta).Objeto = Objeto
        Subasta(Cant_Subasta).Cantidad = Cantidad
        Subasta(Cant_Subasta).Valor = Precio
        Subasta(Cant_Subasta).Subastador = UserList(UserIndex).Name
        Subasta(Cant_Subasta).Timer = DarTimerSubasta(Timer)
        Subasta(Cant_Subasta).Comprador = ""
        
         Call GuardarSubastas
        
        Call EnviaListSubasta(UserIndex)
        
        Call SendData(ToIndex, UserIndex, 0, "RELOADS")
        
        Call QuitarObjetos(Objeto, Cantidad, UserIndex)
        
        Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " está subastando: " & ObjData(Objeto).Name & " (Cantidad: " & Cantidad & ") con un precio inicial de " & Precio & " monedas de oro. Ve al NPC para ofertar si deseas participar." & FONTTYPE_GUERRA)
End Sub

Public Sub IntervaloSubasta(ByVal Id As Integer)
     
     If Subasta(Id).Timer >= 0 Then
         Subasta(Id).Timer = Subasta(Id).Timer - 1
     End If
     
     If Subasta(Id).Timer < 0 Then
         SavingSubasta = True
         Call SubastaFinalizada(Id)
     End If
     
End Sub

Public Sub SubastaFinalizada(Id)
       
       Dim TIndex As Integer
       Dim Obj As Obj
       
       If Subasta(Id).Comprador = "" Then
       
           Call SendData(ToAll, 0, 0, "||La subasta de " & Subasta(Id).Subastador & " ha finalizado sin ofertantes." & FONTTYPE_GUERRA)
               
           TIndex = NameIndex(Subasta(Id).Subastador)
           
           If TIndex > 0 Then
           
               Obj.ObjIndex = Subasta(Id).Objeto
               Obj.Amount = Subasta(Id).Cantidad
               Call MeterItemEnInventario(TIndex, Obj)
               
           Else
                
               Call DepositarItemOffline(UCase$(Subasta(Id).Subastador), Subasta(Id).Objeto, Subasta(Id).Cantidad)
                
           End If
           
           Else
           
           Call SendData(ToAll, 0, 0, "||La Subasta de " & Subasta(Id).Subastador & " ha terminado, y ha vendido " & ObjData(Subasta(Id).Objeto).Name & " al personaje " & Subasta(Id).Comprador & "." & FONTTYPE_GUERRA)
           
           TIndex = NameIndex(Subasta(Id).Comprador)
           
           If TIndex > 0 Then
           
               Obj.ObjIndex = Subasta(Id).Objeto
               Obj.Amount = Subasta(Id).Cantidad
               Call MeterItemEnInventario(TIndex, Obj)
               
           Else
                
               Call DepositarItemOffline(UCase$(Subasta(Id).Comprador), Subasta(Id).Objeto, Subasta(Id).Cantidad)
                
           End If
           
           Call EntregarOroSubastador(Subasta(Id).Subastador, Subasta(Id).Valor)
           
       End If
       
       
       Subasta(Id).Cantidad = 0
       Subasta(Id).Comprador = ""
       Subasta(Id).Objeto = 0
       Subasta(Id).Subastador = ""
       Subasta(Id).Timer = -1
       Subasta(Id).Valor = -1
       
       Cant_Subasta = Cant_Subasta - 1
       
End Sub

Public Sub OfertaSubasta(ByVal UserIndex As Integer, ByVal Id As Integer, ByVal Subastador As String, ByVal Objeto As Integer, ByVal Oferta As Integer)
     
     If Subasta(Id).Subastador <> Subastador Or Subasta(Id).Objeto <> Objeto Then
         Call SendData(ToIndex, UserIndex, 0, "||Hubo un problema en subasta! Actualiza la subasta!" & FONTTYPE_INFO)
         Exit Sub
     End If
     
     If UCase$(Subasta(Id).Subastador) = UCase$(UserList(UserIndex).Name) Then
         Call SendData(ToIndex, UserIndex, 0, "||No puedes hacerte una auto oferta a tu propia subasta." & FONTTYPE_INFO)
         Exit Sub
     End If
     
     If UCase$(Subasta(Id).Comprador) = UCase$(UserList(UserIndex).Name) Then
         Call SendData(ToIndex, UserIndex, 0, "||Actualmente, hay una oferta tuya sobre esta subasta." & FONTTYPE_INFO)
         Exit Sub
     End If
     
     If Subasta(Id).Valor >= Oferta Then
          Call SendData(ToIndex, UserIndex, 0, "||La oferta debe ser superior de los " & Subasta(Id).Valor & " oro." & FONTTYPE_INFO)
          Exit Sub
     End If
     
     If Subasta(Id).Comprador <> "" And UCase$(Subasta(Id).Comprador) <> UCase$(UserList(UserIndex).Name) Then
         Call SendData(ToAll, 0, 0, "||La oferta de " & Subasta(Id).Comprador & ", con el objeto: " & ObjData(Subasta(Id).Objeto).Name & ", fue superada por " & UserList(UserIndex).Name & "." & FONTTYPE_TALK)
     End If
     
     Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " ha Ofertado la Cantidad: " & Oferta & " monedas de oro. por el " & ObjData(Subasta(Id).Objeto).Name & " del Usuario: " & Subasta(Id).Subastador & FONTTYPE_GUERRA)
     
     Subasta(Id).Comprador = UserList(UserIndex).Name
     Subasta(Id).Valor = Oferta
     
End Sub

Public Sub EntregarOroSubastador(ByVal Subastador As String, ByVal Oro As Long)
       
       Dim TIndex As Integer
       Dim Valor As Long
       
       TIndex = NameIndex(Subastador)
       
       If TIndex > 0 Then
           
           UserList(TIndex).Stats.GLD = UserList(TIndex).Stats.GLD + Oro
           Call EnviarOro(TIndex)
           
       Else
            
            Subastador = UCase$(Subastador)
            
            Valor = GetVar(App.Path & "\Charfile\" & Subastador & ".chr", "STATS", "GLD")
            
            Valor = Valor + Oro
            
            Call WriteVar(App.Path & "\Charfile\" & Subastador & ".chr", "STATS", "GLD", Valor)
            
       End If
       
End Sub

Function DarTimerSubasta(ByVal Timer As Integer) As Long
      
      DarTimerSubasta = (60 * Timer)
      
End Function

Function VerTimerSubasta(ByVal Timer As Integer) As String
      
      Dim Datos As String
      Dim c As Long
      
      If Timer >= 60 Then
          c = val((Timer / 60))
          Datos = c & " Hrs "
          Timer = Timer - (60 * c)
      End If
      
      If Timer < 60 Then
          Datos = Datos & Timer & " Mins"
      End If
      
      
      VerTimerSubasta = Datos
      
End Function

Sub DepositarItemOffline(ByVal Comprador As String, _
                         ByVal ObjIndex As Integer, _
                         ByVal Cantidad As Integer)

    Dim Slot   As Integer

    Dim obji   As Integer

    Dim Nitems As Integer

    Dim LoopC  As Integer

    Dim ln     As String

    If Cantidad < 1 Then Exit Sub

    obji = ObjIndex

    Nitems = GetVar(App.Path & "\Charfile\" & Comprador & ".chr", "BancoInventory", "CantidadItems")

    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
        ln = GetVar(App.Path & "\Charfile\" & Comprador & ".chr", "BancoInventory", "Obj" & LoopC)
        Debug.Print ln
        bancoObj(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
        bancoObj(LoopC).Amount = CInt(ReadField(2, ln, 45))
    Next LoopC

    '¿Ya tiene un objeto de este tipo?
    Slot = 1

    Do Until bancoObj(Slot).ObjIndex = obji And bancoObj(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
        Slot = Slot + 1
        
        If Slot > MAX_BANCOINVENTORY_SLOTS Then

            Exit Do

        End If

    Loop

    'Sino se fija por un slot vacio antes del slot devuelto
    If Slot > MAX_BANCOINVENTORY_SLOTS Then

        Slot = 1

        Do Until bancoObj(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_BANCOINVENTORY_SLOTS Then

                Call ItemLastPos(Comprador, Cantidad, obji)
                Exit Sub
                Exit Do

            End If

        Loop

        If Slot <= MAX_BANCOINVENTORY_SLOTS Then Call WriteVar(App.Path & "\Charfile\" & Comprador & ".chr", "BancoInventory", "CantidadItems", Nitems + 1)
        
    End If

    If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido

        'Mete el obj en el slot
        If bancoObj(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        
            'Menor que MAX_INV_OBJS
            Call WriteVar(App.Path & "\Charfile\" & Comprador & ".chr", "BancoInventory", "Obj" & Slot, obji & "-" & bancoObj(Slot).Amount + Cantidad)

        Else
            Call ItemLastPos(Comprador, Cantidad, obji)

        End If

    Else

        'Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
    End If

End Sub

Sub ItemLastPos(ByVal PJ As String, sCant As Integer, objndX As Integer)

    Dim DameLastPos As String

    Dim lastPos     As WorldPos

    Dim MiObj       As Obj

    MiObj.Amount = sCant
    MiObj.ObjIndex = objndX

    DameLastPos = GetVar(App.Path & "\Charfile\" & PJ & ".chr", "INIT", "Position")

    lastPos.Map = CInt(ReadField(1, DameLastPos, 45))
    lastPos.X = CInt(ReadField(2, DameLastPos, 45))
    lastPos.Y = CInt(ReadField(3, DameLastPos, 45))

    Call TirarItemAlPiso(lastPos, MiObj)

End Sub

