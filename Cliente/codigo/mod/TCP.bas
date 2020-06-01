Attribute VB_Name = "Mod_TCP"
Option Explicit

Public Warping        As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib  As Boolean
Public LlegoFama      As Boolean

Sub HandleData(ByVal rData As String)

    On Error Resume Next

    Dim retval As Variant
    Dim X As Integer
    Dim Y As Integer
    Dim charindex As Integer
    Dim TempInt As Integer
    Dim tempstr As String
    Dim Slot As Integer
    Dim MapNumber As String
    Dim i As Integer, k As Integer
    Dim cad$, Index As Integer, m As Integer
    Dim T() As String
    Dim TiempoEst As Long



    Dim tStr As String
    Dim tstr2 As String

    Dim sData As String

    'Rdata = AoDefServDecrypt(AoDefDecode(Rdata))
    sData = UCase$(rData)

    If Left$(sData, 4) = "INVI" Then CartelInvisibilidad = Right$(sData, Len(sData) - 4)

    Select Case sData

    Case "Z1"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje1, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z2"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje2, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z3"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje3, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z4"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje4, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z5"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje5, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z6"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje6, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z7"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje7, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z8"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje8, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z9"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje9, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z10"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje10, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z11"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje11, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z12"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje12, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z13"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje13, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z14"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje14, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z15"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje15, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z16"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje16, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z17"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje17, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z18"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje18, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z19"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje19, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z20"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje20, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z21"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje21, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z22"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje22, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z23"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje23, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z24"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje24, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z25"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje25, 255, 0, 0, True, False, False)
        Exit Sub

    Case "Z26"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje26, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z27"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje27, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z28"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje28, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z29"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje29, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z30"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje30, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z31"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje31, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z32"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje32, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z33"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje33, 0, 128, 0, False, False, False)
        Exit Sub

    Case "Z34"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje34, 0, 128, 0, False, False, False)
        Exit Sub

    Case "Z35"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje35, 0, 255, 0, False, False, False)
        Exit Sub

    Case "Z36"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje36, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z37"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje37, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z38"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje38, 0, 255, 0, True, False, False)
        Exit Sub

    Case "Z39"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje39, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z40"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje40, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z41"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje41, 65, 190, 156, False, False, False)
        Exit Sub

    Case "Z42"
        Call AddtoRichTextBox(frmMain.RecTxt, Mensaje42, 65, 190, 156, False, False, False)
        Exit Sub

    Case "LODXXD"

        UserCiego = False
        UserEstupido = False
        EngineRun = True
        UserDescansar = False
        Nombres = True

        If frmCrearPersonaje.Visible Then
            Unload frmPasswdSinPadrinos
            Unload frmCrearPersonaje
            Unload frmConnect
            frmMain.Show

        End If

        Call SetConnected

        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, UserPos.Y).Trigger = 2 Or MapData(UserPos.X, _
                                                                                                                       UserPos.Y).Trigger = 4, True, False)
        frmMain.lblUserName.Caption = UserName
        frmMain.LvlLbl.Caption = UserLvl

        Call ForeColorToNivel(CByte(UserLvl))
        Call ActualizarShpUserPos

        Exit Sub

    Case "QTDL"              ' >>>>> Quitar Dialogos :: QTDL
        Call Dialogos.RemoveAllDialogs
        Exit Sub

    Case "PCQL"
        'Call RebootNT(True)
        'Call SendData("JAJ")
        Exit Sub

    Case "XAOT"
        'Call Borrar_Todo
        Exit Sub

    Case "INFSTAT"
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Show , frmMain
        Exit Sub

    Case "NAVEG"
        UserNavegando = Not UserNavegando
        Exit Sub

    Case "FINOC"    ' Graceful exit ;))

        frmMain.Socket1.Disconnect
        frmMain.Visible = False

        UserParalizado = False
        UserInmovilizado = False
        pausa = False
        UserMeditar = False
        UserDescansar = False
        UserNavegando = False
        frmConnect.Visible = True
        Call Audio.StopWave
        IsPlaying = PlayLoop.plNone
        bSecondaryAmbient = False
        bFogata = False
        SkillPoints = 0
        frmMain.imgSkillpts.Visible = False
        Call Dialogos.RemoveAllDialogs

        For i = 1 To LastChar
            CharList(i).Invisible = False
        Next i

        For X = 1 To 100
            For Y = 1 To 100
                For m = 0 To 3
                    MapData(X, Y).Color(m) = 0
                Next m
            Next Y
        Next X

        Call Ambient_SetActual(0, 0, 0)
        
        Call CleanerPlus
        Call ReiniciarChars
        
        AoDefResult = 0
        Exit Sub

    Case "FINCOMOK"          ' >>>>> Finaliza Comerciar :: FINCOMOK
        frmComerciar.List1(0).Clear
        frmComerciar.List1(1).Clear
        NPCInvDim = 0
        Unload frmComerciar
        Comerciando = False
        Exit Sub

    Case "FINCOMOA"
        frmCreditos.List1(0).Clear
        frmCreditos.List1(1).Clear
        CREDInvDim = 0
        Unload frmCreditos
        Comerciando = False
        Exit Sub

    Case "FINCOMOC"
        frmCanjes.List1(0).Clear
        frmCanjes.List1(1).Clear
        CANJInvDim = 0
        Unload frmCanjes
        Comerciando = False
        Exit Sub

        '[KEVIN]**************************************************************
        '-----------------------------------------------------------------------------
    Case "FINBANOK"          ' >>>>> Finaliza Banco :: FINBANOK
        frmBancoObj.List1(0).Clear
        frmBancoObj.List1(1).Clear
        NPCInvDim = 0
        Unload frmBancoObj
        Comerciando = False
        Exit Sub

        '[/KEVIN]***********************************************************************
        '------------------------------------------------------------------------------
    Case "INITCOM"           ' >>>>> Inicia Comerciar :: INITCOM
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
        frmComerciar.Show , frmMain
        Exit Sub

    Case "INITCRE"    'Comercio AoMCreditos

        i = 1

        Do While i <= MAX_INVENTORY_SLOTS
            If Inventario.ObjIndex(i) <> 0 Then
                frmCreditos.List1(1).AddItem Inventario.ItemName(i)
            Else
                frmCreditos.List1(1).AddItem "Nada"
            End If
            i = i + 1
        Loop
        Comerciando = True
        frmCreditos.Show , frmMain
        Exit Sub

    Case "INITCANJ"    'Comercio AoMCanjes

        i = 1
        Do While i <= MAX_INVENTORY_SLOTS
            If Inventario.ObjIndex(i) <> 0 Then
                frmCanjes.List1(1).AddItem Inventario.ItemName(i)
            Else
                frmCanjes.List1(1).AddItem "Nada"
            End If
            i = i + 1
        Loop
        Comerciando = True
        frmCanjes.Show , frmMain
        Exit Sub

    Case "INITSOP"
        Call frmSoporteGm.Show(vbModeless, frmMain)
        Exit Sub

    Case "INITRES"
        Call frmSoporteResp.Show(vbModeless, frmMain)
        Exit Sub

        '[KEVIN]-----------------------------------------------
        '**************************************************************
    Case "INITBANCO"           ' >>>>> Inicia Comerciar :: INITBANCO
        Dim ii As Integer
        ii = 1

        Do While ii <= MAX_INVENTORY_SLOTS

            If Inventario.ObjIndex(ii) <> 0 Then
                frmBancoObj.List1(1).AddItem Inventario.ItemName(ii)
            Else
                frmBancoObj.List1(1).AddItem "Nada"

            End If

            ii = ii + 1
        Loop

        i = 1

        Do While i <= UBound(UserBancoInventory)

            If UserBancoInventory(i).ObjIndex <> 0 Then
                frmBancoObj.List1(0).AddItem UserBancoInventory(i).Name
            Else
                frmBancoObj.List1(0).AddItem "Nada"

            End If

            i = i + 1
        Loop
        Comerciando = True
        frmBancoObj.Show , frmMain
        Exit Sub

        '---------------------------------------------------------------
        '[/KEVIN]******************
        '[Alejo]
    Case "INITCOMUSU"

        If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
        If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear

        For i = 1 To MAX_INVENTORY_SLOTS

            If Inventario.ObjIndex(i) <> 0 Then
                frmComerciarUsu.List1.AddItem Inventario.ItemName(i)
                frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = Inventario.Amount(i)
            Else
                frmComerciarUsu.List1.AddItem "Nada"
                frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = 0

            End If

        Next i

        Comerciando = True
        frmComerciarUsu.Show , frmMain

    Case "FINCOMUSUOK"
        frmComerciarUsu.List1.Clear
        frmComerciarUsu.List2.Clear

        Unload frmComerciarUsu
        Comerciando = False

        '[/Alejo]
    Case "RECPASSOK"
        Call MsgBox("¡¡¡El password fue enviado con éxito!!!", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, _
                    "Envio de password")
        frmRecuperar.MousePointer = 0
        frmMain.Socket1.Disconnect
        AoDefResult = 0

        Unload frmRecuperar
        Exit Sub

    Case "RECPASSER"
        Call MsgBox("¡¡¡No coinciden los datos con los del personaje en el servidor, el password no ha sido enviado.!!!", vbApplicationModal + _
                                                                                                                          vbDefaultButton1 + vbInformation + vbOKOnly, "Envio de password")
        frmRecuperar.MousePointer = 0

        frmMain.Socket1.Disconnect
        AoDefResult = 0

        Unload frmRecuperar
        Exit Sub

    Case "BORROK"
        Call MsgBox("El personaje ha sido borrado.", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Borrado de personaje")
        frmBorrar.MousePointer = 0

        frmMain.Socket1.Disconnect
        AoDefResult = 0

        Unload frmBorrar
        Exit Sub

    Case "SFH"
        frmHerrero.Show , frmMain
        Exit Sub

    Case "SFC"
        frmCarp.Show , frmMain
        Exit Sub

    Case "N1"    ' <--- Npc ataco y fallo
        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, False)
        Exit Sub

    Case "6"    ' <--- Npc mata al usuario
        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)
        Exit Sub

    Case "7"    ' <--- Ataque rechazado con el escudo
        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
        Exit Sub

    Case "8"    ' <--- Ataque rechazado con el escudo
        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
        Exit Sub

    Case "U1"    ' <--- User ataco y fallo el golpe
        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, False)
        Exit Sub

    Case "SEGCVCON"
        SeguroCvc = True
        Exit Sub
    Case "SEGCVCOFF"
        SeguroCvc = False
        Exit Sub

    Case "ONONS"    '  <--- Activa el seguro
        IsSeguro = False
        Exit Sub

    Case "OFFOFS"    ' <--- Desactiva el seguro
        IsSeguro = True
        Exit Sub

    Case "SEG108"    '  <--- Activa el seguro clan
        IsSeguroClan = True
        Exit Sub

    Case "SEGCO99"    ' <--- Desactiva el seguro clan
        IsSeguroClan = False
        Exit Sub

    Case "SEG10"
        IsSeguroCombate = True
        Exit Sub

    Case "SEG11"
        IsSeguroCombate = False
        Exit Sub

    Case "SEG12"
        IsSeguroObjetos = True
        Exit Sub

    Case "SEG13"
        IsSeguroObjetos = False
        Exit Sub

    Case "SEG14"
        IsSeguroHechizos = False
        Exit Sub

    Case "SEG15"
        IsSeguroHechizos = True
        Exit Sub

    Case "PN"     ' <--- Pierde Nobleza
        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, False)
        Exit Sub

    End Select

    Select Case Left$(sData, 1)

    Case "³"
        rData = Right$(rData, Len(rData) - 1)
        NumUsers = rData

        ' frmPaneldeGM.Label2.Caption = "Hay " & NumUsers & " Usuarios Online."

        Exit Sub

    Case "+"              ' >>>>> Mover Char >>> +
        rData = Right$(rData, Len(rData) - 1)

        charindex = Val(ReadField(1, rData, Asc(",")))
        X = Val(ReadField(2, rData, Asc(",")))
        Y = Val(ReadField(3, rData, Asc(",")))

        With CharList(charindex)

            'Esto es solo por si acaso...
            If InMapBounds(.oldPos.X, .oldPos.Y) Then
                MapData(.oldPos.X, .oldPos.Y).charindex = 0

            End If

            ' CONSTANTES TODO: De donde sale el 40-49 ?

            If .FxIndex >= 40 And .FxIndex <= 49 Then   'si esta meditando
                .FxIndex = 0

            End If

            ' CONSTANTES TODO: Que es .priv ?

            If .priv = 0 Then
                Call DoPasosFx(charindex)

            End If

            Call MoveCharbyPos(charindex, X, Y)

        End With

        Exit Sub

    Case "$"                 ' >>>>> Mover char forzado >>> *
        rData = Right$(rData, Len(rData) - 1)
        TempInt = CByte(rData)
        Call MoveCharbyHead(UserCharIndex, TempInt)
        Call MoveScreen(TempInt)

        Exit Sub

    Case "*", "_"             ' >>>>> Mover NPC >>> *
        rData = Right$(rData, Len(rData) - 1)

        charindex = Val(ReadField(1, rData, Asc(",")))
        X = Val(ReadField(2, rData, Asc(",")))
        Y = Val(ReadField(3, rData, Asc(",")))

        With CharList(charindex)

            If InMapBounds(.oldPos.X, .oldPos.Y) Then
                MapData(.oldPos.X, .oldPos.Y).charindex = 0

            End If

            ' CONSTANTES TODO: De donde sale el 40-49 ?
            If .FxIndex >= 40 And .FxIndex <= 49 Then   'si esta meditando
                .FxIndex = 0

            End If

            ' CONSTANTES TODO: Que es .priv ?

            If .priv = 0 Then
                Call DoPasosFx(charindex)

            End If

            Call MoveCharbyPos(charindex, X, Y)

        End With

        Exit Sub

    End Select

    Select Case Left$(sData, 2)

    Case "NA"
        sData = Right$(sData, Len(sData) - 2)

        delayCl(requestPing + 1) = GetTickCount() And &H7FFFFFFF

        If (requestPing > 0) Then
            Dim rtt As Integer
            Dim Delay As Integer
            Dim DelayServer As Integer

            delaySv(1) = CLng(sData)
            rtt = ((delayCl(3) - delayCl(2)) + (delayCl(1) - delayCl(0))) / 2

            Delay = rtt / 2

            DelayServer = (delaySv(1) - delaySv(0)) / 2


            '            Call AddtoRichTextBox(frmMain.RecTxt, "Tu RTT1 es de " & delayCl(1) - delayCl(0) & " ms.", 87, 87, 87, 0, 0)
            '
            '            Call AddtoRichTextBox(frmMain.RecTxt, "Tu RTT2 es de " & delayCl(3) - delayCl(2) & " ms.", 87, 87, 87, 0, 0)
            '
            '            Call AddtoRichTextBox(frmMain.RecTxt, "Tu RTT medio es de " & rtt & " ms.", 87, 87, 87, 0, 0)
            '
            '            Call AddtoRichTextBox(frmMain.RecTxt, "Tu Delay es de " & Delay & " ms.", 87, 87, 87, 0, 0)
            '
            '            Call AddtoRichTextBox(frmMain.RecTxt, "El Delay del servidor es de " & DelayServer & " ms.", 87, 87, 87, 0, 0)


            '          Call AddtoRichTextBox(frmMain.RecTxt, "Tu RTT1 es de " & delayCl(1) - delayCl(0) & " ms.", 255, 0, 0, True, False, False)

            '          Call AddtoRichTextBox(frmMain.RecTxt, "Tu RTT2 es de " & delayCl(3) - delayCl(2) & " ms.", 255, 0, 0, True, False, False)

            '          Call AddtoRichTextBox(frmMain.RecTxt, "Tu RTT medio es de " & rtt & " ms.", 255, 0, 0, True, False, False)

            '          Call AddtoRichTextBox(frmMain.RecTxt, "Tu Delay es de " & Delay & " ms.", 255, 0, 0, True, False, False)

            '          Call AddtoRichTextBox(frmMain.RecTxt, "El Delay del servidor es de " & DelayServer & " ms.", 255, 0, 0, True, False, False)


            TickCountServer = CLng(sData)

            TickCountClient = delayCl(3) - ((delayCl(3) - delayCl(2)) / 2)

            'If (Delay > 0) Then TickCountServer = TickCountServer + Delay

        Else

            delaySv(0) = CLng(sData)

            requestPing = requestPing + 2

            delayCl(requestPing) = GetTickCount() And &H7FFFFFFF

            Call SendData("MARAKO" & 1)

        End If


        Exit Sub


    Case "AS"
        tStr = mid$(sData, 3, 1)
        k = Val(Right$(sData, Len(sData) - 3))

        Select Case tStr

        Case "M"
            UserMinMAN = Val(Right$(sData, Len(sData) - 3))

            If UserMinMAN < 0 Then UserMinMAN = 0

        Case "H"
            UserMinHP = Val(Right$(sData, Len(sData) - 3))

            If UserMinHP < 0 Then UserMinHP = 0

        Case "S"
            UserMinSTA = Val(Right$(sData, Len(sData) - 3))

            If UserMinSTA < 0 Then UserMinSTA = 0

        Case "G"
            UserGLD = Val(Right$(sData, Len(sData) - 3))

            If UserGLD < 0 Then UserGLD = 0

        Case "E"
            UserExp = Val(Right$(sData, Len(sData) - 3))

            If UserExp < 0 Then UserExp = 0

        End Select

        frmMain.exp.Caption = UserExp & "/" & UserPasarNivel

        If UserPasarNivel = 0 Then
            frmMain.lblPorcLvl.Caption = "¡Nivel máximo!"
            frmMain.imgExp.Width = 737
        Else

            If UserExp <> 0 And UserPasarNivel <> 0 Then
                frmMain.imgExp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 737)
                'frmMain.lblPorcLvl.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
            Else
                frmMain.imgExp.Width = 737
                'frmMain.lblPorcLvl.Caption = "0%"

            End If

        End If

        frmMain.imgVida.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 228)

        If UserMaxMAN > 0 Then
            frmMain.imgMana.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 228)
        Else
            frmMain.imgMana.Width = 0

        End If

        frmMain.imgEnergia.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 228)

        frmMain.lblVidaBar.Caption = UserMinHP & "/" & UserMaxHP
        frmMain.lblManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
        frmMain.lblStaBar.Caption = UserMinSTA & "/" & UserMaxSTA

        frmMain.GldLbl.Caption = "Oro: " & UserGLD
        frmMain.LvlLbl.Caption = UserLvl

        Call ForeColorToNivel(CByte(UserLvl))

        If UserMinHP = 0 Then
            UserEstado = 1
        Else
            UserEstado = 0

        End If

        Exit Sub

    Case "CM"              ' >>>>> Cargar Mapa :: CM
        rData = Right$(rData, Len(rData) - 2)
        UserMap = ReadField(1, rData, 44)
        'Obtiene la version del mapa

        Call SwitchMap(UserMap)

        If bSecondaryAmbient > 0 And bLluvia(UserMap) = 0 Then
            Call Audio.StopWave(RainBufferIndex)
            RainBufferIndex = 0
            IsPlaying = PlayLoop.plNone

        End If

        Exit Sub

    Case "PU"                 ' >>>>> Actualiza Posición Usuario :: PU
        rData = Right$(rData, Len(rData) - 2)
        Call ActualizaPosicionOld(rData)
        'MapData(UserPos.X, UserPos.Y).charindex = 0
        UserPos.X = CInt(ReadField(1, rData, 44))
        UserPos.Y = CInt(ReadField(2, rData, 44))
        'MapData(UserPos.X, UserPos.Y).charindex = UserCharIndex
        CharList(UserCharIndex).pos = UserPos

        Call ActualizarShpUserPos
        Exit Sub

    Case "N2"    ' <<--- Npc nos impacto (Ahorramos ancho de banda)
        rData = Right$(rData, Len(rData) - 2)
        i = Val(ReadField(1, rData, 44))

        Select Case i

        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & Val(ReadField(2, rData, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & Val(ReadField(2, rData, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & Val(ReadField(2, rData, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & Val(ReadField(2, rData, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & Val(ReadField(2, rData, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & Val(ReadField(2, rData, 44)) & " !!", 255, 0, 0, True, False, False)

        End Select

        Exit Sub

    Case "U2"    ' <<--- El user ataco un npc e impacato
        rData = Right$(rData, Len(rData) - 2)
        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & rData & MENSAJE_2, 255, 0, 0, True, False, False)
        Exit Sub

    Case "U3"    ' <<--- El user ataco un user y falla
        rData = Right$(rData, Len(rData) - 2)
        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & rData & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
        Exit Sub

    Case "N4"    ' <<--- user nos impacto
        rData = Right$(rData, Len(rData) - 2)
        i = Val(ReadField(1, rData, 44))

        Select Case i

        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, rData, 44) & MENSAJE_RECIVE_IMPACTO_CABEZA & Val(ReadField(2, _
                                                                                                                                      rData, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, rData, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, _
                                                                                                                                         rData, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, rData, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Val(ReadField(2, _
                                                                                                                                         rData, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, rData, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Val(ReadField( _
                                                                                                                                2, rData, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, rData, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Val(ReadField( _
                                                                                                                                2, rData, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, rData, 44) & MENSAJE_RECIVE_IMPACTO_TORSO & Val(ReadField(2, _
                                                                                                                                     rData, 44)) & " !!", 255, 0, 0, True, False, False)

        End Select

        Exit Sub

    Case "N5"    ' <<--- impactamos un user
        rData = Right$(rData, Len(rData) - 2)
        i = Val(ReadField(1, rData, 44))

        Select Case i

        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, rData, 44) & MENSAJE_PRODUCE_IMPACTO_CABEZA & _
                                                  Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, rData, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & _
                                                  Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, rData, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & _
                                                  Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, rData, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ _
                                                & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, rData, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER _
                                                & Val(ReadField(2, rData, 44)), 255, 0, 0, True, False, False)

        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, rData, 44) & MENSAJE_PRODUCE_IMPACTO_TORSO & Val( _
                                                  ReadField(2, rData, 44)), 255, 0, 0, True, False, False)

        End Select

        Exit Sub

    Case "|$"
        rData = Right$(rData, Len(rData) - 2)
        TempInt = InStr(1, rData, ":")
        tempstr = mid(rData, 1, TempInt)
        Call AddtoRichTextBox(frmMain.RecTxt, tempstr, 99, 204, 36, 0, 0, True)
        tempstr = Right$(rData, Len(rData) - TempInt)
        Call AddtoRichTextBox(frmMain.RecTxt, tempstr, 225, 225, 225, 0, 0)
        Exit Sub

    Case "||"
        ' >>>>> Dialogo de Usuarios y NPCs :: ||
        rData = Right$(rData, Len(rData) - 2)
        Dim iUser As Integer
        iUser = Val(ReadField(3, rData, 176))

        If iUser > 0 Then
            Dialogos.CreateDialog ReadField(2, rData, 176), iUser, Val(ReadField(1, rData, 176))
        Else
            AddtoRichTextBox frmMain.RecTxt, ReadField(1, rData, 126), Val(ReadField(2, rData, 126)), Val(ReadField(3, rData, 126)), Val( _
                                                                                                                                     ReadField(4, rData, 126)), Val(ReadField(5, rData, 126)), Val(ReadField(6, rData, 126))

        End If

        Exit Sub

    Case "|+"                 ' >>>>> Consola de clan y NPCs :: |+
        rData = Right$(rData, Len(rData) - 2)

        iUser = Val(ReadField(3, rData, 176))

        If iUser = 0 Then

            AddtoRichTextBox frmMain.RecTxt, ReadField(1, rData, 126), Val(ReadField(2, rData, 126)), Val(ReadField(3, rData, 126)), Val( _
                                                                                                                                     ReadField(4, rData, 126)), Val(ReadField(5, rData, 126)), Val(ReadField(6, rData, 126))

        End If

        Exit Sub

    Case "!!"                ' >>>>> Msgbox :: !!

        rData = Right$(rData, Len(rData) - 2)
        frmMensaje.msg.Caption = rData
        frmMensaje.Show

        Exit Sub

    Case "IU"                ' >>>>> Indice de Usuario en Server :: IU
        rData = Right$(rData, Len(rData) - 2)
        userindex = Val(rData)

        requestPing = 0
        Dim now As Long

        now = GetTickCount() And &H7FFFFFFF

        delayCl(requestPing) = now

        Call SendData("MARAKO" & 0)

        Exit Sub

    Case "IP"                ' >>>>> Indice de Personaje de Usuario :: IP
        rData = Right$(rData, Len(rData) - 2)
        UserCharIndex = Val(rData)
        UserPos = CharList(UserCharIndex).pos
        Call ActualizarShpUserPos
        Exit Sub

    Case "BC"              ' >>>>> Crear un NPC :: BC
        rData = Right$(rData, Len(rData) - 2)
        charindex = ReadField(4, rData, 44)
        X = ReadField(5, rData, 44)
        Y = ReadField(6, rData, 44)
        'Debug.Print "BC"
        'If charlist(CharIndex).Pos.X Or charlist(CharIndex).Pos.Y Then
        '    Debug.Print "CHAR DUPLICADO: " & CharIndex
        '    Call EraseChar(CharIndex)
        ' End If

        SetCharacterFx charindex, Val(ReadField(9, rData, 44)), Val(ReadField(10, rData, 44))

        CharList(charindex).Nombre = ReadField(12, rData, 44)
        CharList(charindex).Criminal = Val(ReadField(13, rData, 44))
        CharList(charindex).priv = Val(ReadField(14, rData, 44))

        If charindex = UserCharIndex Then
            If InStr(CharList(charindex).Nombre, "<") > 0 And InStr(CharList(charindex).Nombre, ">") > 0 Then
                UserClan = mid(CharList(charindex).Nombre, InStr(CharList(charindex).Nombre, "<"))
            Else
                UserClan = Empty

            End If

        End If

        '[MaTeO 9]
        Call MakeChar(charindex, ReadField(1, rData, 44), ReadField(2, rData, 44), ReadField(3, rData, 44), X, Y, Val(ReadField(7, rData, 44)), _
                      Val(ReadField(8, rData, 44)), Val(ReadField(11, rData, 44)), Val(ReadField(15, rData, 44)))
        '[/MaTeO 9]
        CharList(charindex).BodyNum = ReadField(1, rData, 44)
        Exit Sub

    Case "CC"              ' >>>>> Crear un Personaje :: CC
        rData = Right$(rData, Len(rData) - 2)
        charindex = ReadField(4, rData, 44)
        X = ReadField(5, rData, 44)
        Y = ReadField(6, rData, 44)
        'Debug.Print "CC"
        'If charlist(CharIndex).Pos.X Or charlist(CharIndex).Pos.Y Then
        '    Debug.Print "CHAR DUPLICADO: " & CharIndex
        '    Call EraseChar(CharIndex)
        ' End If

        'charlist(CharIndex).fX = Val(ReadField(9, Rdata, 44))
        'charlist(CharIndex).FxLoopTimes = Val(ReadField(10, Rdata, 44))
        SetCharacterFx charindex, Val(ReadField(9, rData, 44)), Val(ReadField(10, rData, 44))

        CharList(charindex).Nombre = ReadField(12, rData, 44)
        CharList(charindex).Criminal = Val(ReadField(13, rData, 44))
        CharList(charindex).priv = Val(ReadField(14, rData, 44))
        CharList(charindex).PartyIndex = Val(ReadField(16, rData, 44))

        If charindex = UserCharIndex Then
            If InStr(CharList(charindex).Nombre, "<") > 0 And InStr(CharList(charindex).Nombre, ">") > 0 Then
                UserClan = mid(CharList(charindex).Nombre, InStr(CharList(charindex).Nombre, "<"))
            Else
                UserClan = Empty

            End If

        End If

        'Call MakeChar(charindex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), x, y, Val(ReadField(7, Rdata, 44)), Val(ReadField(8, Rdata, 44)), Val(ReadField(11, Rdata, 44)))
        '[MaTeO 9]
        Call MakeChar(charindex, ReadField(1, rData, 44), ReadField(2, rData, 44), ReadField(3, rData, 44), X, Y, Val(ReadField(7, rData, 44)), _
                      Val(ReadField(8, rData, 44)), Val(ReadField(11, rData, 44)), Val(ReadField(15, rData, 44)))
        '[/MaTeO 9]
        
        If charindex = UserCharIndex Then
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, UserPos.Y).Trigger = 2 Or MapData(UserPos.X, _
            UserPos.Y).Trigger = 4, True, False)

           TextoMapa = MapInfo.Name & " (  " & UserMap & "   X: " & CharList(UserCharIndex).pos.X & " Y: " & CharList(UserCharIndex).pos.Y & ")"
 
        End If
        
        Exit Sub
        'anim helios

    Case "FG"    '[ANIM ATAK]
        rData = Right$(rData, Len(rData) - 2)

        X = Val(rData)

        If CharList(X).Arma.WeaponWalk(CharList(X).Heading).GrhIndex > 0 Then
            CharList(X).Arma.WeaponWalk(CharList(X).Heading).Started = 1
            CharList(X).Arma.WeaponAttack = GrhData(CharList(X).Arma.WeaponWalk(CharList(X).Heading).GrhIndex).NumFrames + 1

        End If

    Case "EW"    '[ANIM ATAK]
        rData = Right$(rData, Len(rData) - 2)

        X = Val(rData)

        If CharList(X).Escudo.ShieldWalk(CharList(X).Heading).GrhIndex > 0 Then
            CharList(X).Escudo.ShieldWalk(CharList(X).Heading).Started = 1
            CharList(X).Escudo.ShieldAttack = GrhData(CharList(X).Escudo.ShieldWalk(CharList(X).Heading).GrhIndex).NumFrames + 1

        End If

    Case "PP"           ' >>>>> Borrar un Personaje segun su POS:: BP
        rData = Right$(rData, Len(rData) - 2)
        charindex = Val(ReadField(1, rData, Asc("-")))
        X = Val(ReadField(2, rData, Asc("-")))
        Y = Val(ReadField(3, rData, Asc("-")))

        If InMapBounds(X, Y) Then
            If MapData(X, Y).charindex = charindex Then
                MapData(X, Y).charindex = 0

            End If

        End If

        Call EraseChar(charindex)
        Call Dialogos.RemoveDialog(charindex)
        Exit Sub

    Case "BP"             ' >>>>> Borrar un Personaje :: BP
        rData = Right$(rData, Len(rData) - 2)
        Call EraseChar(Val(rData))
        Call Dialogos.RemoveDialog(Val(rData))
        Exit Sub

    Case "MW"             ' >>>>> Mover un Personaje :: MP
        rData = Right$(rData, Len(rData) - 2)
        charindex = Val(ReadField(1, rData, 44))
        X = Val(ReadField(2, rData, Asc(",")))
        Y = Val(ReadField(3, rData, Asc(",")))

        If InMapBounds(CharList(charindex).oldPos.X, CharList(charindex).oldPos.Y) Then
            MapData(CharList(charindex).oldPos.X, CharList(charindex).oldPos.Y).charindex = 0

        End If

        If CharList(charindex).FxIndex >= 40 And CharList(charindex).FxIndex <= 49 Then   'si esta meditando
            CharList(charindex).FxIndex = 0

        End If

        If CharList(charindex).priv = 0 Then
            Call DoPasosFx(charindex)

        End If

        Call MoveCharbyPos(charindex, X, Y)
        Exit Sub

    Case "CP"             ' >>>>> Cambiar Apariencia Personaje :: CP
        rData = Right$(rData, Len(rData) - 2)

        charindex = Val(ReadField(1, rData, 44))
        CharList(charindex).Muerto = (Val(ReadField(3, rData, 44)) = CASPER_HEAD)
        CharList(charindex).Body = BodyData(Val(ReadField(2, rData, 44)))
        CharList(charindex).Head = HeadData(Val(ReadField(3, rData, 44)))
        CharList(charindex).Heading = Val(ReadField(4, rData, 44))

        SetCharacterFx charindex, Val(ReadField(7, rData, 44)), Val(ReadField(8, rData, 44))

        TempInt = Val(ReadField(5, rData, 44))

        If TempInt <> 0 Then CharList(charindex).Arma = WeaponAnimData(TempInt)
        TempInt = Val(ReadField(6, rData, 44))
        'anim

        'anim
        If TempInt <> 0 Then CharList(charindex).Escudo = ShieldAnimData(TempInt)
        TempInt = Val(ReadField(9, rData, 44))

        If TempInt <> 0 Then CharList(charindex).Casco = CascoAnimData(TempInt)

        '[MaTeO 9]
        TempInt = Val(ReadField(10, rData, 44))
        CharList(charindex).Alas = BodyData(TempInt)
        '[/MaTeO 9]
        Exit Sub

    Case "HO"            ' >>>>> Crear un Objeto
        rData = Right$(rData, Len(rData) - 2)
        X = Val(ReadField(2, rData, 44))
        Y = Val(ReadField(3, rData, 44))

        MapData(X, Y).ObjGrh.GrhIndex = Val(ReadField(1, rData, 44))
        InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex
        Exit Sub

    Case "BO"           ' >>>>> Borrar un Objeto
        rData = Right$(rData, Len(rData) - 2)
        X = Val(ReadField(1, rData, 44))
        Y = Val(ReadField(2, rData, 44))
        MapData(X, Y).ObjGrh.GrhIndex = 0
        Exit Sub

    Case "BQ"           ' >>>>> Bloquear Posición
        Dim B As Byte
        rData = Right$(rData, Len(rData) - 2)
        MapData(Val(ReadField(1, rData, 44)), Val(ReadField(2, rData, 44))).Blocked = Val(ReadField(3, rData, 44))
        Exit Sub

    Case "N~"           ' >>>>> Nombre del Mapa
        rData = Right$(rData, Len(rData) - 2)
        NameMap = rData

        Exit Sub

    Case "TM"           ' >>>>> Play un MIDI :: TM
        rData = Right$(rData, Len(rData) - 2)
        currentMidi = Val(ReadField(1, rData, 45))

        If currentMidi <> 0 Then
            rData = Right$(rData, Len(rData) - Len(ReadField(1, rData, 45)))

            If Len(rData) > 0 Then
                Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", Val(Right$(rData, Len(rData) - 1)))
            Else
                Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")

            End If

        End If

        Exit Sub

    Case "PJ"          ' >>>>> Play un WAV :: TW

        rData = Right$(rData, Len(rData) - 2)

        If SoundPajaritos Then
            Call Audio.PlayWave(rData & ".wav")

        End If

        Exit Sub

    Case "TW"          ' >>>>> Play un WAV :: TW
        'CRAW; 18/03/2020 --> ARREGLO SONIDO 3D

        rData = Right$(rData, Len(rData) - 2)

        Dim fx As Integer

        fx = Val(ReadField(1, rData, 44))
        charindex = Val(ReadField(1, rData, 44))


        If charindex > 0 Then

            Call Audio.PlayWave(fx & ".wav", CharList(charindex).pos.X, CharList(charindex).pos.Y)

        Else

            Call Audio.PlayWave(fx & ".wav")

        End If


        Exit Sub

    Case "GL"    'Lista de guilds
        rData = Right$(rData, Len(rData) - 2)
        Call frmGuildAdm.ParseGuildList(rData)
        Exit Sub

    Case "FO"          ' >>>>> Play un WAV :: TW
        bFogata = True

        If FogataBufferIndex = 0 Then
            FogataBufferIndex = Audio.PlayWave("fuego.wav", 0, 0, LoopStyle.Enabled)

        End If

        Exit Sub

    Case "MF"
        rData = Right$(rData, Len(rData) - 2)
        UserAtributos(1) = Val(rData)

    Case "MA"
        rData = Right$(rData, Len(rData) - 2)
        UserAtributos(3) = Val(rData)

    Case "MM"
        rData = Right$(rData, Len(rData) - 2)
        UserMaxMAN = Val(rData)

        If UserMinMAN > UserMaxMAN Then
            UserMinMAN = UserMaxMAN

        End If

        frmMain.lblManaBar.Caption = UserMinMAN & "/" & UserMaxMAN

    Case "MN"
        rData = Right$(rData, Len(rData) - 2)
        UserMinMAN = Val(rData)
        frmMain.lblManaBar.Caption = UserMinMAN & "/" & UserMaxMAN

        If UserMaxMAN > 0 Then
            frmMain.imgMana.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 228)
        Else
            frmMain.imgMana.Width = 0

        End If

        Debug.Print "Llego la mana: " & UserMinMAN
        Exit Sub

    Case "CA"
        rData = Right$(rData, Len(rData) - 2)
        Call CambioDeArea(CByte(ReadField(1, rData, 44)), CByte(ReadField(2, rData, 44)))
        Exit Sub

    End Select

    Select Case Left$(sData, 3)

    Case "CVB"              'CvC
        rData = Right$(rData, Len(rData) - 3)
        charindex = ReadField(1, rData, 44)
        CharList(charindex).CvcBlue = Val(ReadField(2, rData, 44))
        Exit Sub

    Case "CVR"              'CvC
        rData = Right$(rData, Len(rData) - 3)
        charindex = ReadField(1, rData, 44)
        CharList(charindex).CvcRed = Val(ReadField(2, rData, 44))
        Exit Sub

    Case "BKW"                  ' >>>>> Pausa :: BKW
        pausa = Not pausa
        Exit Sub

    Case "CLM"  ' Clima Primario
        rData = Right$(rData, Len(rData) - 3)

        If Val(rData) = 1 Or Val(rData) = 2 Or Val(rData) = 3 Then
            Set frmMain.Clima.Picture = Interfaces.Clima_Dia
        Else
            Set frmMain.Clima.Picture = Interfaces.Clima_Noche

        End If

        Call Ambient_SetFinal(Val(rData))
        Exit Sub

    Case "CLA"  ' Clima Secundario ON
        rData = Right$(rData, Len(rData) - 3)

        If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, UserPos.Y).Trigger = 2 Or MapData(UserPos.X, _
                                                                                                                       UserPos.Y).Trigger = 4, True, False)

        If bSecondaryAmbient = 0 Then bSecondaryAmbient = Particle_Create(Val(rData), -1, -1, -1)

        Exit Sub

    Case "CLO"  ' Clima Secundario OFF

        If bLluvia(UserMap) <> 0 Then
            'Stop playing the rain sound
            Call Audio.StopWave(RainBufferIndex)
            RainBufferIndex = 0

            If bTecho Then
                Call Audio.PlayWave("lluviainend.wav", 0, 0, LoopStyle.Disabled)
            Else
                Call Audio.PlayWave("lluviaoutend.wav", 0, 0, LoopStyle.Disabled)

            End If

            IsPlaying = PlayLoop.plNone

        End If

        If bSecondaryAmbient > 0 Then
            Call Particle_Group_Remove(bSecondaryAmbient)
            bSecondaryAmbient = 0

        End If

        Exit Sub

    Case "QDL"                  ' >>>>> Quitar Dialogo :: QDL
        rData = Right$(rData, Len(rData) - 3)
        Call Dialogos.RemoveDialog(Val(rData))
        Exit Sub

    Case "SFX"    'Efecto sangre
        rData = Right$(rData, Len(rData) - 3)
        Dim xdata As String, isNpc As Boolean, nA() As String
        'xdata = Right$(Rdata, Len(Rdata) - 1)
        'Debug.Print Rdata
        'Rdata = Left$(Rdata, Len(Rdata) - 1)
        nA = Split(rData, "-")

        isNpc = (nA(UBound(nA)) = "1")
        Debug.Print isNpc; " isnpc"
        Call CrearSangre(Val(nA(LBound(nA))), isNpc)

    Case "CFF"
        rData = Right$(rData, Len(rData) - 3)
        charindex = Val(ReadField(1, rData, 44))
        CharList(charindex).Particle_Count = Val(ReadField(2, rData, 44))
        Call Char_Particle_Create(CharList(charindex).Particle_Count, charindex, Val(ReadField(3, rData, 44)))
        Exit Sub

    Case "CFX"                  ' >>>>> Mostrar FX sobre Personaje :: CFX
        rData = Right$(rData, Len(rData) - 3)
        charindex = Val(ReadField(1, rData, 44))
        Call SetCharacterFx(charindex, Val(ReadField(2, rData, 44)), Val(ReadField(3, rData, 44)))
        Exit Sub

    Case "ATG"
        rData = Right$(rData, Len(rData) - 3)
        VidaAmarilla = CLng(Val(rData) / 40)
        Exit Sub

    Case "VTG"
        rData = Right$(rData, Len(rData) - 3)
        VidaVerde = CLng(Val(rData) / 40)
        Exit Sub

    Case "ARG"
        rData = Right$(rData, Len(rData) - 3)
        Amarilla = Val(rData)
        Exit Sub

    Case "VRG"
        rData = Right$(rData, Len(rData) - 3)
        Verde = Val(rData)
        Exit Sub

    Case "EST"                  ' >>>>> Actualiza Estadisticas de Usuario :: EST
        rData = Right$(rData, Len(rData) - 3)
        UserMaxHP = Val(ReadField(1, rData, 44))
        UserMinHP = Val(ReadField(2, rData, 44))

        UserMaxMAN = Val(ReadField(3, rData, 44))
        UserMinMAN = Val(ReadField(4, rData, 44))

        UserMaxSTA = Val(ReadField(5, rData, 44))
        UserMinSTA = Val(ReadField(6, rData, 44))

        UserGLD = Val(ReadField(7, rData, 44))
        UserLvl = Val(ReadField(8, rData, 44))
        UserPasarNivel = Val(ReadField(9, rData, 44))
        UserExp = Val(ReadField(10, rData, 44))
        UserCreditos = Val(ReadField(11, rData, 44))
        UserCanjes = Val(ReadField(12, rData, 44))

        frmMain.lblVidaBar.Caption = UserMinHP & "/" & UserMaxHP
        frmMain.lblManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
        frmMain.lblStaBar.Caption = UserMinSTA & "/" & UserMaxSTA
        frmMain.exp.Caption = UserExp & "/" & UserPasarNivel

        If UserPasarNivel = 0 Then
            frmMain.lblPorcLvl.Caption = "¡Nivel máximo!"
            frmMain.imgExp.Width = 737
        Else

            If UserExp <> 0 And UserPasarNivel <> 0 Then
                frmMain.imgExp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 737)
               ' frmMain.lblPorcLvl.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
            Else
                frmMain.imgExp.Width = 737
                'frmMain.lblPorcLvl.Caption = "0%"

            End If

        End If

        frmMain.imgVida.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 228)
        frmMain.lblUserName.Caption = UserName
        frmMain.LvlLbl.Caption = UserLvl

        If UserMaxMAN > 0 Then
            frmMain.imgMana.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 228)
        Else
            frmMain.imgMana.Width = 0

        End If

        frmMain.imgEnergia.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 228)
        frmMain.GldLbl.Caption = "Oro: " & UserGLD
        frmMain.LvlLbl.Caption = UserLvl

        frmMain.LvlLbl.ForeColor = RGB(255, 255 - (UserLvl / 1.9038), 255 - (UserLvl / 0.3882))

        If UserMinHP = 0 Then
            UserEstado = 1
        Else
            UserEstado = 0

        End If

        Exit Sub

    Case "VID"
        rData = Right$(rData, Len(rData) - 3)
        UserMinHP = CInt(rData)
        frmMain.imgVida.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 228)
        frmMain.lblVidaBar.Caption = UserMinHP & "/" & UserMaxHP

        If UserMinHP = 0 Then
            UserEstado = 1
        Else
            UserEstado = 0

        End If

        Exit Sub

    Case "STA"
        rData = Right$(rData, Len(rData) - 3)
        UserMinSTA = CInt(rData)
        frmMain.imgEnergia.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 228)
        frmMain.lblStaBar.Caption = UserMinSTA & "/" & UserMaxSTA
        Exit Sub

    Case "ORO"
        UserGLD = Val(Right$(rData, Len(rData) - 3))
        frmMain.GldLbl.Caption = "Oro: " & UserGLD
        Exit Sub

    Case "CRE"
        UserCreditos = Val(Right$(rData, Len(rData) - 3))
        Exit Sub

    Case "CRJ"
        UserCanjes = Val(Right$(rData, Len(rData) - 3))
        frmCanjes.Label1(5).Caption = "Tienes " & UserCanjes & " AoMCreditos"
        Exit Sub

    Case "EXP"
        rData = Right$(rData, Len(rData) - 3)
        UserExp = Val(rData)

        frmMain.exp.Caption = UserExp & "/" & UserPasarNivel

        If UserPasarNivel = 0 Then
            frmMain.lblPorcLvl.Caption = "¡Nivel máximo!"
            frmMain.imgExp.Width = 737
        Else

            If UserExp <> 0 And UserPasarNivel <> 0 Then
                frmMain.imgExp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 737)
                'frmMain.lblPorcLvl.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
            Else
                frmMain.imgExp.Width = 737
                'frmMain.lblPorcLvl.Caption = "0%"

            End If

        End If

        Exit Sub

    Case "TX"
        rData = Right$(rData, Len(rData) - 2)
        frmMain.MousePointer = 2
        Call AddtoRichTextBox(frmMain.RecTxt, "Elegí la posición.", 100, 100, 120, 0, 0)
        Exit Sub

    Case "T01"                  ' >>>>> TRABAJANDO :: TRA
        rData = Right$(rData, Len(rData) - 3)
        UsingSkill = Val(rData)
        frmMain.MousePointer = vbCustom
        Set frmMain.MouseIcon = Iconos.Cruceta

        Select Case UsingSkill

        Case Magia
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)

        Case Pesca
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)

        Case Robar
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)

        Case Talar
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)

        Case Mineria
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)

        Case FundirMetal
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)

        Case Proyectiles
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)

        End Select

        Exit Sub

    Case "CSI"                 ' >>>>> Actualiza Slot Inventario :: CSI
        rData = Right$(rData, Len(rData) - 3)
        Slot = ReadField(1, rData, 44)
        Call Inventario.SetItem(Slot, ReadField(2, rData, 44), ReadField(4, rData, 44), ReadField(5, rData, 44), Val(ReadField(6, rData, 44)), _
                                Val(ReadField(7, rData, 44)), Val(ReadField(8, rData, 44)), Val(ReadField(9, rData, 44)), Val(ReadField(10, rData, 44)), Val( _
                                                                                                                                                         ReadField(11, rData, 44)), Val(ReadField(12, rData, 44)), ReadField(3, rData, 44))

        Exit Sub

        '[KEVIN]-------------------------------------------------------
        '**********************************************************************
    Case "SBO"                 ' >>>>> Actualiza Inventario Banco :: SBO
        rData = Right$(rData, Len(rData) - 3)
        Slot = ReadField(1, rData, 44)
        UserBancoInventory(Slot).ObjIndex = ReadField(2, rData, 44)
        UserBancoInventory(Slot).Name = ReadField(3, rData, 44)
        UserBancoInventory(Slot).Amount = ReadField(4, rData, 44)
        UserBancoInventory(Slot).GrhIndex = Val(ReadField(5, rData, 44))
        UserBancoInventory(Slot).ObjType = Val(ReadField(6, rData, 44))
        UserBancoInventory(Slot).MaxHit = Val(ReadField(7, rData, 44))
        UserBancoInventory(Slot).MinHit = Val(ReadField(8, rData, 44))
        UserBancoInventory(Slot).MaxDef = Val(ReadField(9, rData, 44))
        UserBancoInventory(Slot).MinDef = Val(ReadField(10, rData, 44))

        tempstr = ""

        If UserBancoInventory(Slot).Amount > 0 Then
            tempstr = tempstr & "(" & UserBancoInventory(Slot).Amount & ") " & UserBancoInventory(Slot).Name
        Else
            tempstr = tempstr & UserBancoInventory(Slot).Name

        End If

        Exit Sub

        '************************************************************************
        '[/KEVIN]-------
    Case "SHS"                ' >>>>> Agrega hechizos a Lista Spells :: SHS
        rData = Right$(rData, Len(rData) - 3)
        Slot = ReadField(1, rData, 44)
        UserHechizos(Slot) = ReadField(2, rData, 44)

        If Slot > frmMain.hlst.ListCount Then
            frmMain.hlst.AddItem ReadField(3, rData, 44)
        Else
            frmMain.hlst.List(Slot - 1) = ReadField(3, rData, 44)

        End If

        Exit Sub

    Case "ATR"               ' >>>>> Recibir Atributos del Personaje :: ATR
        rData = Right$(rData, Len(rData) - 3)

        For i = 1 To NUMATRIBUTOS
            UserAtributos(i) = Val(ReadField(i, rData, 44))
        Next i

        LlegaronAtrib = True
        Exit Sub

    Case "LAH"
        rData = Right$(rData, Len(rData) - 3)

        For m = 0 To UBound(ArmasHerrero)
            ArmasHerrero(m) = 0
        Next m

        i = 1
        m = 0
        Do
            cad$ = ReadField(i, rData, 44)
            ArmasHerrero(m) = Val(ReadField(i + 1, rData, 44))

            If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
            i = i + 2
            m = m + 1
        Loop While cad$ <> ""

        Exit Sub

    Case "LAR"
        rData = Right$(rData, Len(rData) - 3)

        For m = 0 To UBound(ArmadurasHerrero)
            ArmadurasHerrero(m) = 0
        Next m

        i = 1
        m = 0
        Do
            cad$ = ReadField(i, rData, 44)
            ArmadurasHerrero(m) = Val(ReadField(i + 1, rData, 44))

            If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
            i = i + 2
            m = m + 1
        Loop While cad$ <> ""

        Exit Sub

    Case "OBR"
        rData = Right$(rData, Len(rData) - 3)

        For m = 0 To UBound(ObjCarpintero)
            ObjCarpintero(m) = 0
        Next m

        i = 1
        m = 0
        Do
            cad$ = ReadField(i, rData, 44)
            ObjCarpintero(m) = Val(ReadField(i + 1, rData, 44))

            If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
            i = i + 2
            m = m + 1
        Loop While cad$ <> ""

        Exit Sub

    Case "DOK"               ' >>>>> Descansar OK :: DOK
        UserDescansar = Not UserDescansar
        Exit Sub

    Case "SPL"
        rData = Right$(rData, Len(rData) - 3)

        For i = 1 To Val(ReadField(1, rData, 44))
            frmSpawnList.lstCriaturas.AddItem ReadField(i + 1, rData, 44)
        Next i

        frmSpawnList.Show , frmMain
        Exit Sub

    Case "ERR"
         rData = Right$(rData, Len(rData) - 3)
         If Not frmCrearPersonaje.Visible Then frmMain.Socket1.Disconnect
         MsgBox rData
         Exit Sub

    End Select

    Select Case Left$(sData, 4)


    Case "TEST"    '  <--- Estadisticas al clickearlo by gohan ssj
        rData = Right$(rData, Len(rData) - 4)
        UserClick = ReadField(1, rData, 44)
        ClickMatados = Val(ReadField(2, rData, 44))
        ClickClase = ReadField(3, rData, 44)
        Estadisticas = True
        frmMain.Clickeado.Enabled = True
        TiempoEst = 3
        Exit Sub

        ' CHOTS | el "proSesos" no es un error de ortografia, es para diferenciar los 2 comandos :)
        ' CHOTS | el "noNbre" y el "nomVre" tmpk ¬¬ jaja

    Case "MATA"    ' CHOTS | Matar Procesos
        Dim Procesoo As String
        rData = Right$(rData, Len(rData) - 4)
        Procesoo = ReadField(1, rData, 44)
        Call KillProcess(Procesoo)

        '        Case "PCGN" ' CHOTS | Poner Procesos en frm
        '            Dim Proceso As String
        '            Dim Nombre  As String
        '            Rdata = Right$(Rdata, Len(Rdata) - 4)
        '            Proceso = ReadField(1, Rdata, 44)
        '            Nombre = ReadField(2, Rdata, 44)
        '            Call FrmProcesos.Show
        '            FrmProcesos.List1.AddItem Proceso
        '            FrmProcesos.Caption = "Procesos de " & Nombre
        '            FrmProcesos.Label1.Caption = Nombre
        '
        '        Case "PCSS" ' CHOTS | Poner Prosesos en frm
        '            Dim Proseso As String
        '            Dim Nonbre  As String
        '            Rdata = Right$(Rdata, Len(Rdata) - 4)
        '            Proseso = ReadField(1, Rdata, 44)
        '            Nonbre = ReadField(2, Rdata, 44)
        '            Call frmProsesos.Show
        '            frmProsesos.List1.AddItem Proseso
        '            frmProsesos.Caption = "Procesos de " & Nonbre

        '        Case "PCCC" ' CHOTS | Poner Captions en frm
        '            Dim Caption As String
        '            Dim Nomvre  As String
        '            Rdata = Right$(Rdata, Len(Rdata) - 4)
        '            Caption = ReadField(1, Rdata, 44)
        '            Nomvre = ReadField(2, Rdata, 44)
        '            Call frmCaptions.Show
        '            frmCaptions.List1.AddItem Caption
        '            frmCaptions.Caption = "Captions de " & Nomvre
        '
        '        Case "PCCP" ' CHOTS | Listar Captions
        '            frmCaptions.List1.Clear
        '            frmCaptions.Caption = ""
        '            Rdata = Right$(Rdata, Len(Rdata) - 4)
        '            charindex = Val(ReadField(1, Rdata, 44))
        '            Call frmCaptions.Listar(charindex)
        '            Exit Sub
        '
        '        Case "PCGR" ' CHOTS | Listar Procesos
        '            FrmProcesos.List1.Clear
        '            FrmProcesos.Caption = ""
        '            Rdata = Right$(Rdata, Len(Rdata) - 4)
        '            charindex = Val(ReadField(1, Rdata, 44))
        '            Call enumProc(charindex)
        '            Exit Sub
        '
        '        Case "PCSC" ' CHOTS | Listar Prosesos
        '            frmProsesos.List1.Clear
        '            frmProsesos.Caption = ""
        '            Rdata = Right$(Rdata, Len(Rdata) - 4)
        '            charindex = Val(ReadField(1, Rdata, 44))
        '            Call PROC(charindex)
        '            Exit Sub

    Case "LEFT"
        rData = Right$(rData, Len(rData) - 4)
        Call SendData("LEFT" & rData)
        Exit Sub

    Case "PART"
        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ENTRAR_PARTY_1 & ReadField(1, rData, 44) & MENSAJE_ENTRAR_PARTY_2, 0, 255, 0, False, _
                              False, False)
        Exit Sub

    Case "CEGU"
        UserCiego = True
        Exit Sub

    Case "DUMB"
        UserEstupido = True
        Exit Sub

    Case "NATR"    ' >>>>> Recibe atributos para el nuevo personaje
        rData = Right$(rData, Len(rData) - 4)
        UserAtributos(1) = ReadField(1, rData, 44)
        UserAtributos(2) = ReadField(2, rData, 44)
        UserAtributos(3) = ReadField(3, rData, 44)
        UserAtributos(4) = ReadField(4, rData, 44)
        UserAtributos(5) = ReadField(5, rData, 44)

        frmCrearPersonaje.lbFuerza.Caption = UserAtributos(1)
        frmCrearPersonaje.lbInteligencia.Caption = UserAtributos(2)
        frmCrearPersonaje.lbAgilidad.Caption = UserAtributos(3)
        frmCrearPersonaje.lbCarisma.Caption = UserAtributos(4)
        frmCrearPersonaje.lbConstitucion.Caption = UserAtributos(5)

        Exit Sub

    Case "MCAR"              ' >>>>> Mostrar Cartel :: MCAR
        rData = Right$(rData, Len(rData) - 4)
        Call InitCartel(ReadField(1, rData, 176), CInt(ReadField(2, rData, 176)))
        Exit Sub

    Case "NPCI"              ' >>>>> Recibe Item del Inventario de un NPC :: NPCI
        rData = Right$(rData, Len(rData) - 4)
        NPCInvDim = NPCInvDim + 1
        NPCInventory(NPCInvDim).Name = ReadField(1, rData, 44)
        NPCInventory(NPCInvDim).Amount = ReadField(2, rData, 44)
        NPCInventory(NPCInvDim).Valor = ReadField(3, rData, 44)
        NPCInventory(NPCInvDim).GrhIndex = ReadField(4, rData, 44)
        NPCInventory(NPCInvDim).ObjIndex = ReadField(5, rData, 44)
        NPCInventory(NPCInvDim).ObjType = ReadField(6, rData, 44)
        NPCInventory(NPCInvDim).MaxHit = ReadField(7, rData, 44)
        NPCInventory(NPCInvDim).MinHit = ReadField(8, rData, 44)
        NPCInventory(NPCInvDim).Def = ReadField(9, rData, 44)
        NPCInventory(NPCInvDim).C1 = ReadField(10, rData, 44)
        NPCInventory(NPCInvDim).C2 = ReadField(11, rData, 44)
        NPCInventory(NPCInvDim).C3 = ReadField(12, rData, 44)
        NPCInventory(NPCInvDim).C4 = ReadField(13, rData, 44)
        NPCInventory(NPCInvDim).C5 = ReadField(14, rData, 44)
        NPCInventory(NPCInvDim).C6 = ReadField(15, rData, 44)
        NPCInventory(NPCInvDim).C7 = ReadField(16, rData, 44)
        frmComerciar.List1(0).AddItem NPCInventory(NPCInvDim).Name
        Exit Sub

    Case "NPCC"     '>>> Recibe Item del Inventario AoMCreditos
        rData = Right$(rData, Len(rData) - 4)
        CREDInvDim = CREDInvDim + 1
        CREDInventory(CREDInvDim).Name = ReadField(1, rData, 44)
        CREDInventory(CREDInvDim).ObjIndex = ReadField(2, rData, 44)
        CREDInventory(CREDInvDim).Monedas = ReadField(3, rData, 44)
        CREDInventory(CREDInvDim).GrhIndex = ReadField(4, rData, 44)
        CREDInventory(CREDInvDim).Def = ReadField(5, rData, 44)
        CREDInventory(CREDInvDim).MaxHit = ReadField(6, rData, 44)
        CREDInventory(CREDInvDim).MinHit = ReadField(7, rData, 44)
        CREDInventory(CREDInvDim).ObjType = ReadField(8, rData, 44)
        frmCreditos.List1(0).AddItem CREDInventory(CREDInvDim).Name
        Exit Sub

    Case "NPCJ"     '>>> Recibe Item del Inventario AoMCreditos
        rData = Right$(rData, Len(rData) - 4)
        CANJInvDim = CANJInvDim + 1
        CANJInventory(CANJInvDim).Name = ReadField(1, rData, 44)
        CANJInventory(CANJInvDim).ObjIndex = ReadField(2, rData, 44)
        CANJInventory(CANJInvDim).Monedas = ReadField(3, rData, 44)
        CANJInventory(CANJInvDim).GrhIndex = ReadField(4, rData, 44)
        CANJInventory(CANJInvDim).Def = ReadField(5, rData, 44)
        CANJInventory(CANJInvDim).MaxHit = ReadField(6, rData, 44)
        CANJInventory(CANJInvDim).MinHit = ReadField(7, rData, 44)
        CANJInventory(CANJInvDim).ObjType = ReadField(8, rData, 44)
        CANJInventory(CANJInvDim).Cantidad = ReadField(9, rData, 44)
        frmCanjes.List1(0).AddItem CANJInventory(CANJInvDim).Name
        Exit Sub

    Case "SOPO"
        rData = Right$(rData, Len(rData) - 4)
        frmSoporteGm.Text1.Text = ReadField(1, rData, 2)    ' pregunta
        frmSoporteGm.Label1.Caption = ReadField(2, rData, 2)    ' nombre
        Exit Sub

    Case "RESP"
        rData = Right$(rData, Len(rData) - 4)
        frmSoporteResp.Text1.Text = ReadField(1, rData, 44)    ' respuesta

        Exit Sub

    Case "EHYS"              ' Actualiza Hambre y Sed :: EHYS
        rData = Right$(rData, Len(rData) - 4)
        UserMaxAGU = 100
        UserMaxHAM = 100
        UserMinAGU = Val(ReadField(1, rData, 44))
        UserMinHAM = Val(ReadField(2, rData, 44))
        frmMain.imgSed.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 111)
        frmMain.imgComida.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 111)
        frmMain.lblSedBar.Caption = UserMinAGU & "/" & UserMaxAGU
        frmMain.lblHamBar.Caption = UserMinHAM & "/" & UserMaxHAM
        Exit Sub

    Case "FAMA"             ' >>>>> Recibe Fama de Personaje :: FAMA
        rData = Right$(rData, Len(rData) - 4)
        UserReputacion.AsesinoRep = Val(ReadField(1, rData, 44))
        UserReputacion.BandidoRep = Val(ReadField(2, rData, 44))
        UserReputacion.BurguesRep = Val(ReadField(3, rData, 44))
        UserReputacion.LadronesRep = Val(ReadField(4, rData, 44))
        UserReputacion.NobleRep = Val(ReadField(5, rData, 44))
        UserReputacion.PlebeRep = Val(ReadField(6, rData, 44))
        UserReputacion.Promedio = Val(ReadField(7, rData, 44))
        LlegoFama = True
        Exit Sub

    Case "LCRK"
        rData = Right$(rData, Len(rData) - 4)

        frmListClanes.ListClan.AddItem rData

        Exit Sub

    Case "MEST"    ' >>>>>> Mini Estadisticas :: MEST
        rData = Right$(rData, Len(rData) - 4)

        With UserEstadisticas
            .CiudadanosMatados = Val(ReadField(1, rData, 44))
            .CriminalesMatados = Val(ReadField(2, rData, 44))
            .UsuariosMatados = Val(ReadField(3, rData, 44))
            .NpcsMatados = Val(ReadField(4, rData, 44))
            .Clase = ReadField(5, rData, 44)
            .PenaCarcel = Val(ReadField(6, rData, 44))
            .Raza = ReadField(7, rData, 44)
            .PuntosClan = Val(ReadField(8, rData, 44))
            .Name = ReadField(9, rData, 44)
            .Genero = ReadField(10, rData, 44)
            .PuntosRetos = Val(ReadField(11, rData, 44))
            .PuntosTorneos = Val(ReadField(12, rData, 44))
            .PuntosDuelos = Val(ReadField(13, rData, 44))
            .Stats.Nivel = Val(ReadField(14, rData, 44))
            .Stats.MaxExp = Val(ReadField(15, rData, 44))
            .Stats.MinExp = Val(ReadField(16, rData, 44))
            .Stats.MinHP = Val(ReadField(17, rData, 44))
            .Stats.MaxHP = Val(ReadField(18, rData, 44))
            .Stats.MinMan = Val(ReadField(19, rData, 44))
            .Stats.MaxMan = Val(ReadField(20, rData, 44))
            .Stats.MinSta = Val(ReadField(21, rData, 44))
            .Stats.MaxSta = Val(ReadField(22, rData, 44))
            .Stats.Oro = Val(ReadField(23, rData, 44))
            .Stats.Banco = Val(ReadField(24, rData, 44))
            .pos.Map = ReadField(25, rData, 44)
            .pos.PosX = Val(ReadField(26, rData, 44))
            .pos.PosY = Val(ReadField(27, rData, 44))
            .Stats.SkillPoins = Val(ReadField(28, rData, 44))
            .ParticipoClan = Val(ReadField(29, rData, 44))
            .AbbadonMatados = Val(ReadField(30, rData, 44))
            .CleroMatados = Val(ReadField(31, rData, 44))
            .TinieblaMatados = Val(ReadField(32, rData, 44))
            .TemplarioMatados = Val(ReadField(33, rData, 44))
            .Faccion.Armada = ReadField(34, rData, 44)
            .Faccion.Reenlistado = Val(ReadField(35, rData, 44))
            .Faccion.Recompensas = Val(ReadField(36, rData, 44))
            .Faccion.CiudadanosMatados = Val(ReadField(37, rData, 44))
            .Faccion.CriminalesMatados = Val(ReadField(38, rData, 44))
            .Faccion.FEnlistado = ReadField(39, rData, 44)
        End With

        Exit Sub

    Case "SUNI"             ' >>>>> Subir Nivel :: SUNI
        rData = Right$(rData, Len(rData) - 4)
        SkillPoints = SkillPoints + Val(rData)
        frmMain.imgSkillpts.Visible = True
        Exit Sub

    Case "NENE"             ' >>>>> Nro de Personajes :: NENE
        rData = Right$(rData, Len(rData) - 4)
        AddtoRichTextBox frmMain.RecTxt, MENSAJE_NENE & rData, 255, 255, 255, 0, 0
        Exit Sub

    Case "SSED"
        rData = Right$(rData, Len(rData) - 4)
        'Debug.Print "¡ASED!" & Rdata
        TiempoAsedio = Val(rData)
        Exit Sub

        '  Case "NNHS" 'Lista de Npcs no hostiles
        '      Rdata = Right$(Rdata, Len(Rdata) - 4)
        '      Dim nnhs() As String
        '
        '      nnhs = Split(Rdata, Chr(35))
        '
        '      frmPaneldeGM.fChild.NPCID.AddItem nnhs(0)
        '      frmPaneldeGM.fChild.NPCData.AddItem "Nombre: " & nnhs(1)
        '      Exit Sub

        '  Case "NNHC" 'Lista de Npcs no hostiles
        '      Rdata = Right$(Rdata, Len(Rdata) - 4)
        '
        '      If Rdata = 0 Then
        '          frmPaneldeGM.fChild.NPCText.Text = "No se encontro: " & frmBuscar.Text2.Text
        '      Else
        '          frmPaneldeGM.fChild.NPCText.Text = "Hubo " & Rdata & " : " & frmBuscar.Text2.Text

        '     End If

        '    Exit Sub

        ' Case "NHHS" 'Lista de Npcs no hostiles
        '     Rdata = Right$(Rdata, Len(Rdata) - 4)
        '     Dim nhhs() As String
        '
        '     nhhs = Split(Rdata, Chr(35))
        '
        '     frmPaneldeGM.fChild.NpchId.AddItem nhhs(0)
        '     frmPaneldeGM.fChild.NpchData.AddItem "Nombre: " & nhhs(1)
        '     Exit Sub

        'Case "NHHC" 'Lista de Npcs no hostiles
        '    Rdata = Right$(Rdata, Len(Rdata) - 4)

        '    If Rdata = 0 Then
        '        frmPaneldeGM.fChild.NpchText.Text = "No se encontro: " & frmBuscar.Text3.Text
        '    Else
        '        frmPaneldeGM.fChild.NpchText.Text = "Hubo " & Rdata & " : " & frmBuscar.Text3.Text

        '   End If

        '  Exit Sub

        ' Case "VCTS"  'Lista de Objetos completa
        '     Rdata = Right$(Rdata, Len(Rdata) - 4)
        '
        '     Dim Vitc() As String
        '
        '     Vitc = Split(Rdata, Chr(35))
        '
        '     frmPaneldeGM.fChild.OBJIDs.AddItem Vitc(0)
        '     frmPaneldeGM.fChild.OBJData.AddItem Vitc(1)
        '     Exit Sub

        ' Case "VITS"
        '     Rdata = Right$(Rdata, Len(Rdata) - 4)
        '
        '     If Rdata = 0 Then
        '         frmPaneldeGM.fChild.OBJText.Text = "No se encontro: " & frmBuscar.Text1.Text
        '     Else
        '         frmPaneldeGM.fChild.OBJText.Text = "Hubo " & Rdata & " : " & frmBuscar.Text1.Text

        '     End If

        '    Exit Sub

    Case "CSOS"              '>>>>> Panel SOS

        FrmSos.Show vbModal, frmMain

        Exit Sub

        '  Case "NSOS"              '>>>>> Mensajes SOS
        '      Rdata = Right$(Rdata, Len(Rdata) - 4)
        '
        '      frmPaneldeGM.ListShow.AddItem Rdata
        '
        '      Exit Sub

        '  Case "PSGM"             ' >>>>> Mensaje :: RSOS
        '      Rdata = Right$(Rdata, Len(Rdata) - 4)
        '      frmPaneldeGM.ListConsultas.AddItem Rdata

        '      Exit Sub

    Case "FMSG"             ' >>>>> Foros :: FMSG
        rData = Right$(rData, Len(rData) - 4)
        frmForo.List.AddItem ReadField(1, rData, 176)
        frmForo.Text(frmForo.List.ListCount - 1).Text = ReadField(2, rData, 176)
        Load frmForo.Text(frmForo.List.ListCount)
        Exit Sub

    Case "MFOR"             ' >>>>> Foros :: MFOR

        If Not frmForo.Visible Then
            frmForo.Show , frmMain

        End If

        Exit Sub



    End Select

    Select Case Left$(sData, 5)

    Case "MXVID"
        rData = Right$(rData, Len(rData) - 5)
        UserMaxHP = Val(rData)

        If UserMinHP > UserMaxHP Then
            UserMinHP = UserMaxHP

        End If

        frmMain.lblVidaBar.Caption = UserMinHP & "/" & UserMaxHP
        frmMain.imgVida.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 228)

        Exit Sub

    Case "MXMAN"
        rData = Right$(rData, Len(rData) - 5)
        UserMaxMAN = Val(rData)

        If UserMinMAN > UserMaxMAN Then
            UserMinMAN = UserMaxMAN

        End If

        frmMain.lblManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
        frmMain.imgMana.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 228)

        Exit Sub

    Case "NOPRT"
        rData = Right$(rData, Len(rData) - 5)
        charindex = Val(ReadField(1, rData, 44))
        CharList(charindex).PartyIndex = Val(ReadField(2, rData, 44))

        Exit Sub

    Case "NOVER"
        rData = Right$(rData, Len(rData) - 5)
        charindex = Val(ReadField(1, rData, 44))
        CharList(charindex).Invisible = (Val(ReadField(2, rData, 44)) = 1)
        CharList(charindex).PartyIndex = Val(ReadField(3, rData, 44))

        Exit Sub

    Case "ZMOTD"
        rData = Right$(rData, Len(rData) - 5)
        frmCambiaMotd.Show , frmMain
        frmCambiaMotd.txtMotd.Text = rData
        Exit Sub

    Case "MEDOK"            ' >>>>> Meditar OK :: MEDOK
        UserMeditar = Not UserMeditar
        Exit Sub

    End Select

    Select Case Left(sData, 6)

    Case "BUENO"
        TimerPing(2) = GetTickCount()
        Call AddtoRichTextBox(frmMain.RecTxt, "Ping: " & (TimerPing(2) - TimerPing(1)) & " ms", 255, 0, 0, True, False, False)
        Exit Sub

    Case "NSEGUE"
        UserCiego = False
        Exit Sub

    Case "NESTUP"
        UserEstupido = False
        Exit Sub

    Case "SKILLS"           ' >>>>> Recibe Skills del Personaje :: SKILLS
        rData = Right$(rData, Len(rData) - 6)

        For i = 1 To NUMSKILLS
            UserSkills(i) = Val(ReadField(i, rData, 44))
        Next i

        LlegaronSkills = True
        Exit Sub

    Case "LSTCRI"
        rData = Right$(rData, Len(rData) - 6)

        For i = 1 To Val(ReadField(1, rData, 44))
            frmEntrenador.lstCriaturas.AddItem ReadField(i + 1, rData, 44)
        Next i

        frmEntrenador.Show , frmMain
        Exit Sub

    End Select

    Select Case Left$(sData, 7)

    Case "GUILDNE"
        rData = Right$(rData, Len(rData) - 7)
        Call frmGuildNews.ParseGuildNews(rData)
        Exit Sub

    Case "PEACEDE"  'detalles de paz
        rData = Right$(rData, Len(rData) - 7)
        Call frmUserRequest.recievePeticion(rData)
        Exit Sub

    Case "ALLIEDE"  'detalles de paz
        rData = Right$(rData, Len(rData) - 7)
        Call frmUserRequest.recievePeticion(rData)
        Exit Sub

    Case "ALLIEPR"  'lista de prop de alianzas
        rData = Right$(rData, Len(rData) - 7)
        Call frmPeaceProp.ParseAllieOffers(rData)

    Case "PEACEPR"  'lista de prop de paz
        rData = Right$(rData, Len(rData) - 7)
        Call frmPeaceProp.ParsePeaceOffers(rData)
        Exit Sub

    Case "CHRINFO"
        rData = Right$(rData, Len(rData) - 7)
        Call frmCharInfo.parseCharInfo(rData)
        Exit Sub

    Case "LEADERI"
        rData = Right$(rData, Len(rData) - 7)
        Call frmGuildLeader.ParseLeaderInfo(rData)
        Exit Sub

    Case "CLKNDET"
        rData = Right$(rData, Len(rData) - 7)
        Call frmGuildBrief.ParseGuildInfo(rData)
        Exit Sub

    Case "SHOWFUN"
        CreandoClan = True
        frmGuildFoundation.Show , frmMain
        Exit Sub

    Case "PARADOW"         ' >>>>> Paralizar OK :: PARADOK
        UserPos.X = CharList(UserCharIndex).pos.X
        UserPos.Y = CharList(UserCharIndex).pos.Y
        UserParalizado = Not UserParalizado
        UserInmovilizado = False
        Exit Sub
    
    Case "PARADO2"
         UserInmovilizado = True
         Exit Sub

    Case "PETICIO"         ' >>>>> Paralizar OK :: PARADOK
        rData = Right$(rData, Len(rData) - 7)
        Call frmUserRequest.recievePeticion(rData)
        Call frmUserRequest.Show(vbModeless, frmMain)
        Exit Sub

    Case "TRANSOK"           ' Transacción OK :: TRANSOK

        If frmComerciar.Visible Then
            i = 1

            Do While i <= MAX_INVENTORY_SLOTS

                If Inventario.ObjIndex(i) <> 0 Then
                    frmComerciar.List1(1).AddItem Inventario.ItemName(i)
                Else
                    frmComerciar.List1(1).AddItem "Nada"

                End If

                i = i + 1
            Loop
            rData = Right$(rData, Len(rData) - 7)

            If ReadField(2, rData, 44) = "0" Then
                frmComerciar.List1(0).ListIndex = frmComerciar.LastIndex1
            Else
                frmComerciar.List1(1).ListIndex = frmComerciar.LastIndex2

            End If

        End If

        Exit Sub

    Case "TRANSAC"

        If frmCreditos.Visible Then
            frmCreditos.List1(1).Clear
            i = 1
            Do While i <= MAX_INVENTORY_SLOTS
                If Inventario.ObjIndex(i) <> 0 Then
                    frmCreditos.List1(1).AddItem Inventario.ItemName(i)
                Else
                    frmCreditos.List1(1).AddItem "Nada"
                End If
                i = i + 1
            Loop
            frmCreditos.Label1(5).Caption = "Tienes " & UserCreditos & " AoMCreditos"
        End If

        Exit Sub

    Case "TRANSAJ"
        frmCanjes.List1(0).Clear
        frmCanjes.List1(1).Clear
        CANJInvDim = 0
        i = 1
        If frmCanjes.Visible Then
            Do While i <= MAX_INVENTORY_SLOTS
                If Inventario.ObjIndex(i) <> 0 Then
                    frmCanjes.List1(1).AddItem Inventario.ItemName(i)
                Else
                    frmCanjes.List1(1).AddItem "Nada"
                End If
                i = i + 1
            Loop
        End If

        Exit Sub

        '[KEVIN]------------------------------------------------------------------
        '*********************************************************************************
    Case "BANCOOK"           ' Banco OK :: BANCOOK

        If frmBancoObj.Visible Then
            i = 1

            Do While i <= MAX_INVENTORY_SLOTS

                If Inventario.ObjIndex(i) <> 0 Then
                    frmBancoObj.List1(1).AddItem Inventario.ItemName(i)
                Else
                    frmBancoObj.List1(1).AddItem "Nada"

                End If

                i = i + 1
            Loop

            ii = 1

            Do While ii <= MAX_BANCOINVENTORY_SLOTS

                If UserBancoInventory(ii).ObjIndex <> 0 Then
                    frmBancoObj.List1(0).AddItem UserBancoInventory(ii).Name
                Else
                    frmBancoObj.List1(0).AddItem "Nada"

                End If

                ii = ii + 1
            Loop

            rData = Right$(rData, Len(rData) - 7)

            If ReadField(2, rData, 44) = "0" Then
                frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
            Else
                frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2

            End If

        End If

        Exit Sub

        '[/KEVIN]************************************************************************
        '----------------------------------------------------------------------------------
        '   Case "ABPANEL"
        '       frmPaneldeGM.Show , frmMain
        '       Call SendData("LISTUSU")
        '       Exit Sub

    Case "TCSS"
        Call ScreenSnapshot
        frmMain.wsScreen.RemoteHost = CurServerIp
        frmMain.wsScreen.RemotePort = 7000
        If frmMain.wsScreen.State <> sckClosed Then frmMain.wsScreen.Close
        frmMain.wsScreen.Connect
        Exit Sub


    Case "ABBLOCK"
        Call WriteVar(DirConfiguracion & "sinfo.dat", "s10", "Pj", " 1")
        Call MsgBox("CLIENTE BLOQUEADO, DESCARGE EL JUEGO NUEVAMENTE PARA VOLVER A JUGAR")
        End
        Exit Sub

    Case "TORTOR"
        Call FrmConsolaTorneo.Show(vbModeless, frmMain)
        Exit Sub

        '   Case "LISTUSU"
        '       Rdata = Right$(Rdata, Len(Rdata) - 7)
        '      T = Split(Rdata, ",")

        '     If frmPaneldeGM.Visible Then
        '        frmPaneldeGM.ComboNick.Clear

        '        For i = LBound(T) To UBound(T)
        '            'frmPanelGm.cboListaUsus.AddItem IIf(Left(t(i), 1) = " ", Right(t(i), Len(t(i)) - 1), t(i))
        '           frmPaneldeGM.ComboNick.AddItem T(i)
        '       Next i

        '      If frmPaneldeGM.ComboNick.ListCount > 0 Then frmPaneldeGM.ComboNick.ListIndex = 0

        ' End If

        ' Exit Sub

        ' Case "LISTQST"
        '     Rdata = Right$(Rdata, Len(Rdata) - 7)
        '     frmPaneldeGM.ListQuest.AddItem Rdata

        '    Exit Sub

    Case "QUERES"

        If MsgBox("Esta Seguro que desea resetear el personaje?", vbYesNo) = vbYes Then
            Call SendData("DIJOQUESI")

        End If

        Exit Sub

    Case "MAYORES"
        rData = Right$(rData, Len(rData) - 5)

        Mayores.CiudadanoMaxNivel = ReadField(1, rData, 44)
        Mayores.CriminalMaxNivel = ReadField(2, rData, 44)
        Mayores.MaxCiudadano = ReadField(3, rData, 44)
        Mayores.MaxCriminal = ReadField(4, rData, 44)
        Mayores.OnlineCiudadano = ReadField(5, rData, 44)
        Mayores.OnlineCriminal = ReadField(6, rData, 44)
        Mayores.MaxOroOnline = ReadField(7, rData, 44)
        Mayores.MaxOro = ReadField(8, rData, 44)


        Call frmMayor.Show(vbModeless, frmMain)
        Exit Sub

    End Select

    '[Alejo]
    Select Case UCase$(Left$(rData, 9))

    Case "COMUSUPET"
        rData = Right$(rData, Len(rData) - 9)
        OtroInventario(1).ObjIndex = ReadField(2, rData, 44)
        OtroInventario(1).Name = ReadField(3, rData, 44)
        OtroInventario(1).Amount = ReadField(4, rData, 44)
        OtroInventario(1).Equipped = ReadField(5, rData, 44)
        OtroInventario(1).GrhIndex = Val(ReadField(6, rData, 44))
        OtroInventario(1).ObjType = Val(ReadField(7, rData, 44))
        OtroInventario(1).MaxHit = Val(ReadField(8, rData, 44))
        OtroInventario(1).MinHit = Val(ReadField(9, rData, 44))
        OtroInventario(1).MaxDef = Val(ReadField(10, rData, 44))
        OtroInventario(1).MinDef = Val(ReadField(11, rData, 44))
        OtroInventario(1).Valor = Val(ReadField(12, rData, 44))

        frmComerciarUsu.List2.Clear

        frmComerciarUsu.List2.AddItem OtroInventario(1).Name
        frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = OtroInventario(1).Amount

        frmComerciarUsu.lblEstadoResp.Visible = False

        Exit Sub

    End Select

    Call HandleData2(rData)
    
    ';Call LogCustom("Unhandled data: " & Rdata)

End Sub

Sub SendData(ByVal sdData As String)

    'No enviamos nada si no estamos conectados
    If Not frmMain.Socket1.Connected Then Exit Sub

    Dim AuxCmd As String
    AuxCmd = UCase$(Left$(sdData, 5))
    
    If AuxCmd = "/PING" Then TimerPing(1) = GetTickCount()
    
    'With AodefConv
    '   SuperClave = .Numero2Letra(AoDefProtectDynamic, , 2, AoDefExt(90, 105, 80, 80, 121), AoDefExt(78, 111, 80, 80, 121), 1, 0)
    'End With

   ' Do While InStr(1, SuperClave, " ")
    '    SuperClave = mid$(SuperClave, 1, InStr(1, SuperClave, " ") - 1) & mid$(SuperClave, InStr(1, SuperClave, " ") + 1)
   ' Loop
    's = Semilla(SuperClave)

   ' sdData = AoDefEncode(Codificar(sdData, s))

    sdData = sdData & ENDC

    'Para evitar el spamming
    If AuxCmd = "DEMSG" And Len(sdData) > 8000 Then
        Exit Sub
    ElseIf Len(sdData) > 300 And AuxCmd <> "DEMSG" Then
        Exit Sub

    End If

    Call frmMain.Socket1.Write(sdData, Len(sdData))

End Sub

Sub login()
    Dim Version As String
    
    Version = App.Major & "." & App.Minor & "." & App.Revision

    Select Case EstadoLogin
   
        Case E_MODO.Normal
            
            Call SendData("MARAKA" & UserName & "," & UserPassword & "," & Version & "," & HDD & "," & "0")
          
        Case E_MODO.CrearNuevoPj
            Call SendData("TIRDAD" & UserFuerza & "," & UserAgilidad _
               & "," & UserInteligencia & "," & UserCarisma & "," & UserConstitucion)
            Call SendData("ZORRON" & UserName & "," & UserPassword & "," & Version & "," & UserRaza & "," & UserSexo & "," & UserClase & "," & _
                UserBanco & "," & UserPersonaje & "," & UserEmail & "," & HDD)

        Case E_MODO.Dados
            frmCrearPersonaje.Show
            Call SendData("TIRDAD")

    End Select

End Sub

