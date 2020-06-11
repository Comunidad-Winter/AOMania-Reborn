Attribute VB_Name = "Mod_TCP2"
Option Explicit

Sub HandleData2(ByVal rData As String)
   
    Dim Rs        As Integer

    Dim LooPC     As Integer

    Dim charindex As Long

    Dim X         As Integer
    
    Dim i As Integer
   
    Select Case UCase$(Left$(rData, 2))
        
        Case "PL"

            For X = 1 To 10
                  
                frmMain.UserClanPos(X).Visible = False
                  
            Next X

            Exit Sub
        
        Case "PO"
            rData = Right$(rData, Len(rData) - 2)
            X = Val(ReadField(3, rData, 44))
            ClanPos(X).X = Val(ReadField(1, rData, 44))
            ClanPos(X).Y = Val(ReadField(2, rData, 44))
            Call ActualizarShpClanPos
            Exit Sub

        Case "XN"             '>>>>>> Coge información de quest NPC
            rData = Right$(rData, Len(rData) - 2)
            charindex = ReadField(1, rData, 44)
            CharList(charindex).NpcType = ReadField(2, rData, 44)
            CharList(charindex).NombreNpc = ReadField(3, rData, 44)
            CharList(charindex).Hostile = ReadField(4, rData, 44)
            Exit Sub
       
        Case "XU"          '>>>>>>> Coge datos de quest usuario y abre frmquest
            rData = Right$(rData, Len(rData) - 2)
            Quest.NumQuests = ReadField(1, rData, 44)
         
            For LooPC = 1 To NumQuests
                Quest.InfoUser.UserQuest(LooPC) = ReadField(LooPC, rData, 44)
            Next LooPC
            
            frmQuest.Show , frmMain
            Exit Sub
        
        Case "XP"       '>>>>>>> Actualiza el proceso de la quest
            rData = Right$(rData, Len(rData) - 2)
           
            charindex = ReadField(1, rData, 44)
            ProcesoQuest = Val(ReadField(2, rData, 44))
           
            Exit Sub
           
        Case "XI"    '>>>>>>> Actualiza icono npc misiones
            rData = Right$(rData, Len(rData) - 2)
           
            charindex = Val(ReadField(1, rData, 44))
            CharList(charindex).Icono = ReadField(2, rData, 44)
            Exit Sub
        
        Case "XV" '>>>>>>> Ejecuta ventana hablar Npc
            rData = Right$(rData, Len(rData) - 2)
           
            HablarQuest.NumMsj = Val(ReadField(1, rData, 44))
           
            For LooPC = 1 To HablarQuest.NumMsj
                  
                X = LooPC + 1
                  
                HablarQuest.Mensaje(LooPC) = ReadField(X, rData, 44)
                  
            Next LooPC
           
            FrmHablarNpc.Show , frmMain
           
            Exit Sub

    End Select
   
    Select Case UCase$(Left$(rData, 3))
        
        Case "RIG"
            rData = Right$(rData, Len(rData) - 3)
           
            charindex = ReadField(1, rData, 44)
           
            CharList(charindex).Gm = Val(ReadField(2, rData, 44))
           
            Exit Sub
    
        Case "EMC"

            Dim NColor As Integer

            rData = Right$(rData, Len(rData) - 3)
            CountMEC = 380 + Len(CStr(rData))
            MensajeEnvio = String(380, " ") + rData
            frmMain.EnvioMsj.SelStart = 0
            frmMain.EnvioMsj.SelLength = Len(frmMain.EnvioMsj)
        
            NColor = RandomNumber(1, 4)
        
            If NColor = 1 Then
                frmMain.EnvioMsj.SelColor = vbCyan
            ElseIf NColor = 2 Then
                frmMain.EnvioMsj.SelColor = vbWhite
            ElseIf NColor = 3 Then
                frmMain.EnvioMsj.SelColor = vbYellow
            ElseIf NColor = 4 Then
                frmMain.EnvioMsj.SelColor = vbRed

            End If
        
            frmMain.EnvioMsj.SelBold = True
        
            frmMain.TimerMsj.Enabled = True
            Exit Sub
        
        Case "ACT"    'binmode: correccion de posicion
    
            rData = Right$(rData, Len(rData) - 3)
           
            Call ActualizaPosicion(rData)
       
            Exit Sub
        
        Case "SMN"
            MapInfo.Name = Right$(rData, Len(rData) - 3)
            TextoMapa = MapInfo.Name & " (  " & UserMap & "   X: " & CharList(UserCharIndex).pos.X & " Y: " & CharList(UserCharIndex).pos.Y & ")"
            Exit Sub
       
        Case "VPA"
           
            rData = Right$(rData, Len(rData) - 3)
            Rs = Val(ReadField(1, rData, 44))
           
            If Rs = 0 Then
                frmParty.Label1.Visible = True
            ElseIf Rs = 1 Then
              
                For LooPC = 1 To MaxVerParty
              
                    frmParty.Label2(LooPC).Caption = PartyData(LooPC).Name
                    frmParty.Label2(LooPC).Visible = True
                    frmParty.Label3(LooPC).Caption = PartyData(LooPC).MinHP & "/" & PartyData(LooPC).MaxHP
                    frmParty.Label3(LooPC).Visible = True
                    frmParty.Label4(LooPC).Visible = True
                    frmParty.Shape1(LooPC).Visible = True
                  
                    If PartyData(LooPC).MinHP > 0 Then
                        frmParty.Label4(LooPC).Width = (((PartyData(LooPC).MinHP / 100) / (PartyData(LooPC).MaxHP / 100)) * 101)
                    Else
                        frmParty.Label4(LooPC).Width = 0

                    End If
                 
                Next LooPC
              
                frmParty.cmdSalir.Visible = True
              
            End If
           
            frmParty.Show , frmMain
           
            Exit Sub
       
        Case "IVP"
           
            rData = Right$(rData, Len(rData) - 3)
            Rs = Val(ReadField(1, rData, 44))
           
            PartyData(Rs).Name = ReadField(2, rData, 44)
            PartyData(Rs).MinHP = Val(ReadField(3, rData, 44))
            PartyData(Rs).MaxHP = Val(ReadField(4, rData, 44))
           
            MaxVerParty = Rs
           
            Exit Sub
       
        Case "VPT"
            rData = Right$(rData, Len(rData) - 3)
             
            charindex = Val(ReadField(1, rData, 44))
            
            CharList(charindex).Stats.MinHP = Val(ReadField(2, rData, 44))
            CharList(charindex).Stats.MaxHP = Val(ReadField(3, rData, 44))
            CharList(charindex).PartyIndex = Val(ReadField(4, rData, 44))
            Exit Sub
             
    End Select
   
    Select Case UCase$(Left$(rData, 4))
        
       Case "MOTD"
            Call frmMain.ClearConsolas
            Call LeerMotd
            Call Audio.StopWave
            Call Audio.StopMidi
       Exit Sub
        
        Case "HUCT"
            rData = Right$(rData, Len(rData) - 4)
            TimeChange = rData
            Call DayNameChange(rData)
            Exit Sub
   
        Case "VLDB"
            frmValidarBanco.Show , frmMain
            Exit Sub
             
        Case "BANP"
            rData = Right$(rData, Len(rData) - 4)
            
            frmBancoInfo.LblBanco = ReadField(1, rData, 44)
            frmBancoInfo.LblOro = ReadField(2, rData, 44)
            frmBancoInfo.LblObj = ReadField(3, rData, 44)
            
            frmBancoInfo.Show , frmMain
            Exit Sub
       
        Case "BAND"
            rData = Right$(rData, Len(rData) - 4)
           
            frmBancoDepositar.LblOro = rData
           
            frmBancoDepositar.Show , frmMain
            Exit Sub
        
        Case "BANF"
            rData = Right$(rData, Len(rData) - 4)
           
            frmBancoFinal.LblBanco = ReadField(1, rData, 44)
            frmBancoFinal.LblOro = ReadField(2, rData, 44)
           
            frmBancoFinal.Show , frmMain
            Exit Sub
           
        Case "BANR"
            rData = Right$(rData, Len(rData) - 4)
           
            frmBancoRetirar.LblBanco = rData
           
            frmBancoRetirar.Show , frmMain
            Exit Sub
        
        Case "HECA"
            frmOlvidarHechizo.Show , frmMain
            Exit Sub
        
        Case "LSTH"
            rData = Right$(rData, Len(rData) - 4)
           
            frmOlvidarHechizo.List1.AddItem rData
            Exit Sub
        
        Case "ABRC"
          
            rData = Right$(rData, Len(rData) - 4)
            Rs = ReadField(1, rData, 64)
            rData = ReadField(2, rData, 64)
          
            ReDim Heads(1 To Rs)
            frmCabezas.List1.Clear
          
            For LooPC = 1 To Rs
                frmCabezas.List1.AddItem "Cabeza" & LooPC
                Heads(LooPC) = ReadField(LooPC, rData, 44)
            Next LooPC
          
            frmCabezas.Show , frmMain
            Exit Sub
        
        Case "LSTS"
            rData = Right$(rData, Len(rData) - 4)
          
            ObjSastre(NumSastre) = ReadField(1, rData, 64)
            frmSastre.List1.AddItem ReadField(2, rData, 64)
            NumSastre = NumSastre + 1
          
            Exit Sub
          
        Case "ABRS"
            frmSastre.Show , frmMain
            Exit Sub
        
        Case "OBJH"
            rData = Right$(rData, Len(rData) - 4)
          
            ObjHechizeria(NumHechizeria) = ReadField(1, rData, 64)
            frmHechiceria.List1.AddItem ReadField(2, rData, 64)
            NumHechizeria = NumHechizeria + 1
            Exit Sub
        
        Case "ABRH"
            frmHechiceria.Show , frmMain
            Exit Sub
           
        Case "OBHM"
            rData = Right$(rData, Len(rData) - 4)
           
            ObjHerreroMagico(NumHerrero) = ReadField(1, rData, 64)
            frmHerreroMagico.List1.AddItem ReadField(2, rData, 64)
            NumHerrero = NumHerrero + 1
            Exit Sub
      
        Case "ABHM"
            frmHerreroMagico.Show , frmMain
            Exit Sub
          
    End Select
    
    Select Case UCase$(Left$(rData, 6))
         
        Case "TNAVEG"
            rData = Right$(rData, Len(rData) - 6)
           
            CharList(UserCharIndex).VelocidadBarco = Val(rData)
            Debug.Print "TIPO DE VELOCIDAD DE TU BARCO ES: " & Val(rData)
           
            If UserNavegando Then
                If CharList(UserCharIndex).VelocidadBarco = 2 Then
                    ScrollPixelFrame = ScrollPixelFrame + 1
                ElseIf CharList(UserCharIndex).VelocidadBarco = 3 Then
                    ScrollPixelFrame = ScrollPixelFrame + 2

                End If

            ElseIf Not UserNavegando Then

                If CharList(UserCharIndex).VelocidadBarco = 2 Then
                    ScrollPixelFrame = ScrollPixelFrame - 1
                ElseIf CharList(UserCharIndex).VelocidadBarco = 3 Then
                    ScrollPixelFrame = ScrollPixelFrame - 2

                End If

            End If
            
            Debug.Print "ScrollPixelFrame " & ScrollPixelFrame
           
            Exit Sub
         
    End Select
    
    Select Case UCase(Left$(rData, 7))
         
         Case "INITSAG"           ' >>>>> Inicia Comerciar /SAGRADO ::
        i = 1

        Do While i <= MAX_INVENTORY_SLOTS

            If Inventario.ObjIndex(i) <> 0 Then
                frmComerciar.List1(1).AddItem Inventario.ItemName(i)
            Else
                frmComerciar.List1(1).AddItem "Nada"

            End If

            i = i + 1
        Loop
        Comerciando = True
        CanjeSagrado = True
        frmComerciar.Show , frmMain
        Exit Sub
        
        Case "RESETSB"
              
              rData = Val(Right$(rData, Len(rData) - 7))
              
              NumSubasta = rData
              
              ReDim Preserve Subasta(1 To NumSubasta) As tSubasta
              
              NumSubasta = 0
              
        Exit Sub
              
        Case "PAQSUBS"
              
              rData = Right$(rData, Len(rData) - 7)
              
              NumSubasta = NumSubasta + 1
              
              Subasta(NumSubasta).IdObjeto = ReadField(1, rData, 44)
              Subasta(NumSubasta).Objeto = ReadField(2, rData, 44)
              Subasta(NumSubasta).Cantidad = ReadField(3, rData, 44)
              Subasta(NumSubasta).Valor = ReadField(4, rData, 44)
              Subasta(NumSubasta).Subastador = ReadField(5, rData, 44)
              Subasta(NumSubasta).Timer = ReadField(6, rData, 44)
              Subasta(NumSubasta).Comprador = ReadField(7, rData, 44)
              Subasta(NumSubasta).GrhIndex = ReadField(8, rData, 44)
              
        Exit Sub
        
        Case "INITSUB"
              frmSubasta.Show , frmMain
        Exit Sub
        
        Case "RELOADS"
             Unload frmSubastaCrear
             Call frmSubasta.ReloadVentanaSubasta
        Exit Sub
        
        Case "RLDUSUB"
             Call frmSubasta.ReloadVentanaSubasta
        Exit Sub
         
    End Select
    
    Select Case UCase(Left$(rData, 9))
        Case "COMUSUINV"
         rData = Right(rData, Len(rData) - 9)
         OtroInventario(1).ObjIndex = ReadField(2, rData, 44)
         OtroInventario(1).Name = ReadField(3, rData, 44)
         OtroInventario(1).Amount = ReadField(4, rData, 44)
         OtroInventario(1).Equipped = ReadField(5, rData, 44)
         OtroInventario(1).GrhIndex = Val(ReadField(6, rData, 44))
         OtroInventario(1).ObjType = Val(ReadField(7, rData, 44))
         OtroInventario(1).MaxHit = Val(ReadField(8, rData, 44))
         OtroInventario(1).MinHit = Val(ReadField(9, rData, 44))
         OtroInventario(1).MaxDef = Val(ReadField(10, rData, 44))
         OtroInventario(1).Valor = Val(ReadField(11, rData, 44))
         
         frmComerciarUsu.List2.Clear
         
         frmComerciarUsu.List2.AddItem OtroInventario(1).Name
         frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = OtroInventario(1).Amount
         
         frmComerciarUsu.lblEstadoResp.Visible = False
    End Select

End Sub
