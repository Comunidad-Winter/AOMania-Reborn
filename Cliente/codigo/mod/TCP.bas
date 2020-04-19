Attribute VB_Name = "Mod_TCP"
Option Explicit

Public Warping        As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib  As Boolean
Public LlegoFama      As Boolean

Sub HandleData(ByVal Rdata As String)

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
    '[SEGURIDAD SATUROS]
    ' Rdata = AoDefServDecrypt(AoDefDecode(Rdata))
    sData = UCase$(Rdata)

    '[SEGURIDAD SATUROS]
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

        Unload frmRecuperar
        Exit Sub

    Case "RECPASSER"
        Call MsgBox("¡¡¡No coinciden los datos con los del personaje en el servidor, el password no ha sido enviado.!!!", vbApplicationModal + _
                                                                                                                          vbDefaultButton1 + vbInformation + vbOKOnly, "Envio de password")
        frmRecuperar.MousePointer = 0

        frmMain.Socket1.Disconnect

        Unload frmRecuperar
        Exit Sub

    Case "BORROK"
        Call MsgBox("El personaje ha sido borrado.", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Borrado de personaje")
        frmBorrar.MousePointer = 0

        frmMain.Socket1.Disconnect

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
        Rdata = Right$(Rdata, Len(Rdata) - 1)
        NumUsers = Rdata

        ' frmPaneldeGM.Label2.Caption = "Hay " & NumUsers & " Usuarios Online."

        Exit Sub

    Case "+"              ' >>>>> Mover Char >>> +
        Rdata = Right$(Rdata, Len(Rdata) - 1)

        charindex = Val(ReadField(1, Rdata, Asc(",")))
        X = Val(ReadField(2, Rdata, Asc(",")))
        Y = Val(ReadField(3, Rdata, Asc(",")))

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
        Rdata = Right$(Rdata, Len(Rdata) - 1)
        TempInt = CByte(Rdata)
        Call MoveCharbyHead(UserCharIndex, TempInt)
        Call MoveScreen(TempInt)

        Exit Sub

    Case "*", "_"             ' >>>>> Mover NPC >>> *
        Rdata = Right$(Rdata, Len(Rdata) - 1)

        charindex = Val(ReadField(1, Rdata, Asc(",")))
        X = Val(ReadField(2, Rdata, Asc(",")))
        Y = Val(ReadField(3, Rdata, Asc(",")))

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
            frmMain.imgExp.Width = 153
        Else

            If UserExp <> 0 And UserPasarNivel <> 0 Then
                frmMain.imgExp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 153)
                frmMain.lblPorcLvl.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
            Else
                frmMain.imgExp.Width = 153
                frmMain.lblPorcLvl.Caption = "0%"

            End If

        End If

        frmMain.imgVida.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 101)

        If UserMaxMAN > 0 Then
            frmMain.imgMana.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 101)
        Else
            frmMain.imgMana.Width = 0

        End If

        frmMain.imgEnergia.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 101)

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
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        UserMap = ReadField(1, Rdata, 44)
        'Obtiene la version del mapa

        Call SwitchMap(UserMap)

        If bSecondaryAmbient > 0 And bLluvia(UserMap) = 0 Then
            Call Audio.StopWave(RainBufferIndex)
            RainBufferIndex = 0
            IsPlaying = PlayLoop.plNone

        End If

        Exit Sub

    Case "PU"                 ' >>>>> Actualiza Posición Usuario :: PU
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        MapData(UserPos.X, UserPos.Y).charindex = 0
        UserPos.X = CInt(ReadField(1, Rdata, 44))
        UserPos.Y = CInt(ReadField(2, Rdata, 44))
        MapData(UserPos.X, UserPos.Y).charindex = UserCharIndex
        CharList(UserCharIndex).pos = UserPos

        Call ActualizarShpUserPos
        Exit Sub

    Case "N2"    ' <<--- Npc nos impacto (Ahorramos ancho de banda)
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        i = Val(ReadField(1, Rdata, 44))

        Select Case i

        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & Val(ReadField(2, Rdata, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & Val(ReadField(2, Rdata, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & Val(ReadField(2, Rdata, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & Val(ReadField(2, Rdata, 44)) & " !!", 255, 0, 0, True, False, False)

        End Select

        Exit Sub

    Case "U2"    ' <<--- El user ataco un npc e impacato
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & Rdata & MENSAJE_2, 255, 0, 0, True, False, False)
        Exit Sub

    Case "U3"    ' <<--- El user ataco un user y falla
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & Rdata & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
        Exit Sub

    Case "N4"    ' <<--- user nos impacto
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        i = Val(ReadField(1, Rdata, 44))

        Select Case i

        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_CABEZA & Val(ReadField(2, _
                                                                                                                                      Rdata, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, _
                                                                                                                                         Rdata, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Val(ReadField(2, _
                                                                                                                                         Rdata, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Val(ReadField( _
                                                                                                                                2, Rdata, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Val(ReadField( _
                                                                                                                                2, Rdata, 44)) & " !!", 255, 0, 0, True, False, False)

        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_TORSO & Val(ReadField(2, _
                                                                                                                                     Rdata, 44)) & " !!", 255, 0, 0, True, False, False)

        End Select

        Exit Sub

    Case "N5"    ' <<--- impactamos un user
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        i = Val(ReadField(1, Rdata, 44))

        Select Case i

        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_CABEZA & _
                                                  Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & _
                                                  Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & _
                                                  Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ _
                                                & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER _
                                                & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)

        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_TORSO & Val( _
                                                  ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)

        End Select

        Exit Sub

    Case "|$"
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        TempInt = InStr(1, Rdata, ":")
        tempstr = mid(Rdata, 1, TempInt)
        Call AddtoRichTextBox(frmMain.RecTxt, tempstr, 99, 204, 36, 0, 0, True)
        tempstr = Right$(Rdata, Len(Rdata) - TempInt)
        Call AddtoRichTextBox(frmMain.RecTxt, tempstr, 225, 225, 225, 0, 0)
        Exit Sub

    Case "||"
        ' >>>>> Dialogo de Usuarios y NPCs :: ||
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        Dim iUser As Integer
        iUser = Val(ReadField(3, Rdata, 176))

        If iUser > 0 Then
            Dialogos.CreateDialog ReadField(2, Rdata, 176), iUser, Val(ReadField(1, Rdata, 176))
        Else
            AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val( _
                                                                                                                                     ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))

        End If

        Exit Sub

    Case "|+"                 ' >>>>> Consola de clan y NPCs :: |+
        Rdata = Right$(Rdata, Len(Rdata) - 2)

        iUser = Val(ReadField(3, Rdata, 176))

        If iUser = 0 Then

            AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val( _
                                                                                                                                     ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))

        End If

        Exit Sub

    Case "!!"                ' >>>>> Msgbox :: !!

        Rdata = Right$(Rdata, Len(Rdata) - 2)
        frmMensaje.msg.Caption = Rdata
        frmMensaje.Show

        Exit Sub

    Case "IU"                ' >>>>> Indice de Usuario en Server :: IU
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        userindex = Val(Rdata)

        requestPing = 0
        Dim now As Long

        now = GetTickCount() And &H7FFFFFFF

        delayCl(requestPing) = now

        Call SendData("MARAKO" & 0)

        Exit Sub

    Case "IP"                ' >>>>> Indice de Personaje de Usuario :: IP
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        UserCharIndex = Val(Rdata)
        UserPos = CharList(UserCharIndex).pos
        Call ActualizarShpUserPos
        Exit Sub

    Case "BC"              ' >>>>> Crear un NPC :: BC
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        charindex = ReadField(4, Rdata, 44)
        X = ReadField(5, Rdata, 44)
        Y = ReadField(6, Rdata, 44)
        'Debug.Print "BC"
        'If charlist(CharIndex).Pos.X Or charlist(CharIndex).Pos.Y Then
        '    Debug.Print "CHAR DUPLICADO: " & CharIndex
        '    Call EraseChar(CharIndex)
        ' End If

        SetCharacterFx charindex, Val(ReadField(9, Rdata, 44)), Val(ReadField(10, Rdata, 44))

        CharList(charindex).nombre = ReadField(12, Rdata, 44)
        CharList(charindex).Criminal = Val(ReadField(13, Rdata, 44))
        CharList(charindex).priv = Val(ReadField(14, Rdata, 44))

        If charindex = UserCharIndex Then
            If InStr(CharList(charindex).nombre, "<") > 0 And InStr(CharList(charindex).nombre, ">") > 0 Then
                UserClan = mid(CharList(charindex).nombre, InStr(CharList(charindex).nombre, "<"))
            Else
                UserClan = Empty

            End If

        End If

        '[MaTeO 9]
        Call MakeChar(charindex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), X, Y, Val(ReadField(7, Rdata, 44)), _
                      Val(ReadField(8, Rdata, 44)), Val(ReadField(11, Rdata, 44)), Val(ReadField(15, Rdata, 44)))
        '[/MaTeO 9]
        CharList(charindex).BodyNum = ReadField(1, Rdata, 44)
        Exit Sub

    Case "CC"              ' >>>>> Crear un Personaje :: CC
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        charindex = ReadField(4, Rdata, 44)
        X = ReadField(5, Rdata, 44)
        Y = ReadField(6, Rdata, 44)
        'Debug.Print "CC"
        'If charlist(CharIndex).Pos.X Or charlist(CharIndex).Pos.Y Then
        '    Debug.Print "CHAR DUPLICADO: " & CharIndex
        '    Call EraseChar(CharIndex)
        ' End If

        'charlist(CharIndex).fX = Val(ReadField(9, Rdata, 44))
        'charlist(CharIndex).FxLoopTimes = Val(ReadField(10, Rdata, 44))
        SetCharacterFx charindex, Val(ReadField(9, Rdata, 44)), Val(ReadField(10, Rdata, 44))

        CharList(charindex).nombre = ReadField(12, Rdata, 44)
        CharList(charindex).Criminal = Val(ReadField(13, Rdata, 44))
        CharList(charindex).priv = Val(ReadField(14, Rdata, 44))
        CharList(charindex).PartyIndex = Val(ReadField(16, Rdata, 44))

        If charindex = UserCharIndex Then
            If InStr(CharList(charindex).nombre, "<") > 0 And InStr(CharList(charindex).nombre, ">") > 0 Then
                UserClan = mid(CharList(charindex).nombre, InStr(CharList(charindex).nombre, "<"))
            Else
                UserClan = Empty

            End If

        End If

        'Call MakeChar(charindex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), x, y, Val(ReadField(7, Rdata, 44)), Val(ReadField(8, Rdata, 44)), Val(ReadField(11, Rdata, 44)))
        '[MaTeO 9]
        Call MakeChar(charindex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), X, Y, Val(ReadField(7, Rdata, 44)), _
                      Val(ReadField(8, Rdata, 44)), Val(ReadField(11, Rdata, 44)), Val(ReadField(15, Rdata, 44)))
        '[/MaTeO 9]
        Exit Sub
        'anim helios

    Case "FG"    '[ANIM ATAK]
        Rdata = Right$(Rdata, Len(Rdata) - 2)

        X = Val(Rdata)

        If CharList(X).Arma.WeaponWalk(CharList(X).Heading).GrhIndex > 0 Then
            CharList(X).Arma.WeaponWalk(CharList(X).Heading).Started = 1
            CharList(X).Arma.WeaponAttack = GrhData(CharList(X).Arma.WeaponWalk(CharList(X).Heading).GrhIndex).NumFrames + 1

        End If

    Case "EW"    '[ANIM ATAK]
        Rdata = Right$(Rdata, Len(Rdata) - 2)

        X = Val(Rdata)

        If CharList(X).Escudo.ShieldWalk(CharList(X).Heading).GrhIndex > 0 Then
            CharList(X).Escudo.ShieldWalk(CharList(X).Heading).Started = 1
            CharList(X).Escudo.ShieldAttack = GrhData(CharList(X).Escudo.ShieldWalk(CharList(X).Heading).GrhIndex).NumFrames + 1

        End If

    Case "PP"           ' >>>>> Borrar un Personaje segun su POS:: BP
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        charindex = Val(ReadField(1, Rdata, Asc("-")))
        X = Val(ReadField(2, Rdata, Asc("-")))
        Y = Val(ReadField(3, Rdata, Asc("-")))

        If InMapBounds(X, Y) Then
            If MapData(X, Y).charindex = charindex Then
                MapData(X, Y).charindex = 0

            End If

        End If

        Call EraseChar(charindex)
        Call Dialogos.RemoveDialog(charindex)
        Exit Sub

    Case "BP"             ' >>>>> Borrar un Personaje :: BP
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        Call EraseChar(Val(Rdata))
        Call Dialogos.RemoveDialog(Val(Rdata))
        Exit Sub

    Case "MW"             ' >>>>> Mover un Personaje :: MP
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        charindex = Val(ReadField(1, Rdata, 44))
        X = Val(ReadField(2, Rdata, Asc(",")))
        Y = Val(ReadField(3, Rdata, Asc(",")))

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
        Rdata = Right$(Rdata, Len(Rdata) - 2)

        charindex = Val(ReadField(1, Rdata, 44))
        CharList(charindex).Muerto = (Val(ReadField(3, Rdata, 44)) = CASPER_HEAD)
        CharList(charindex).Body = BodyData(Val(ReadField(2, Rdata, 44)))
        CharList(charindex).Head = HeadData(Val(ReadField(3, Rdata, 44)))
        CharList(charindex).Heading = Val(ReadField(4, Rdata, 44))

        SetCharacterFx charindex, Val(ReadField(7, Rdata, 44)), Val(ReadField(8, Rdata, 44))

        TempInt = Val(ReadField(5, Rdata, 44))

        If TempInt <> 0 Then CharList(charindex).Arma = WeaponAnimData(TempInt)
        TempInt = Val(ReadField(6, Rdata, 44))
        'anim

        'anim
        If TempInt <> 0 Then CharList(charindex).Escudo = ShieldAnimData(TempInt)
        TempInt = Val(ReadField(9, Rdata, 44))

        If TempInt <> 0 Then CharList(charindex).Casco = CascoAnimData(TempInt)

        '[MaTeO 9]
        TempInt = Val(ReadField(10, Rdata, 44))
        CharList(charindex).Alas = BodyData(TempInt)
        '[/MaTeO 9]
        Exit Sub

    Case "HO"            ' >>>>> Crear un Objeto
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        X = Val(ReadField(2, Rdata, 44))
        Y = Val(ReadField(3, Rdata, 44))

        MapData(X, Y).ObjGrh.GrhIndex = Val(ReadField(1, Rdata, 44))
        InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex
        Exit Sub

    Case "BO"           ' >>>>> Borrar un Objeto
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        X = Val(ReadField(1, Rdata, 44))
        Y = Val(ReadField(2, Rdata, 44))
        MapData(X, Y).ObjGrh.GrhIndex = 0
        Exit Sub

    Case "BQ"           ' >>>>> Bloquear Posición
        Dim b As Byte
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        MapData(Val(ReadField(1, Rdata, 44)), Val(ReadField(2, Rdata, 44))).Blocked = Val(ReadField(3, Rdata, 44))
        Exit Sub

    Case "N~"           ' >>>>> Nombre del Mapa
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        NameMap = Rdata

        Exit Sub

    Case "TM"           ' >>>>> Play un MIDI :: TM
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        currentMidi = Val(ReadField(1, Rdata, 45))

        If currentMidi <> 0 Then
            Rdata = Right$(Rdata, Len(Rdata) - Len(ReadField(1, Rdata, 45)))

            If Len(Rdata) > 0 Then
                Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", Val(Right$(Rdata, Len(Rdata) - 1)))
            Else
                Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")

            End If

        End If

        Exit Sub

    Case "PJ"          ' >>>>> Play un WAV :: TW

        Rdata = Right$(Rdata, Len(Rdata) - 2)

        If SoundPajaritos Then
            Call Audio.PlayWave(Rdata & ".wav")

        End If

        Exit Sub

    Case "TW"          ' >>>>> Play un WAV :: TW
        'CRAW; 18/03/2020 --> ARREGLO SONIDO 3D

        Rdata = Right$(Rdata, Len(Rdata) - 2)

        Dim fx As Integer

        fx = Val(ReadField(1, Rdata, 44))
        charindex = Val(ReadField(1, Rdata, 44))
        

        If charindex > 0 Then

            Call Audio.PlayWave(fx & ".wav", CharList(charindex).pos.X, CharList(charindex).pos.Y)

        Else

            Call Audio.PlayWave(fx & ".wav")

        End If


        Exit Sub

    Case "GL"    'Lista de guilds
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        Call frmGuildAdm.ParseGuildList(Rdata)
        Exit Sub

    Case "FO"          ' >>>>> Play un WAV :: TW
        bFogata = True

        If FogataBufferIndex = 0 Then
            FogataBufferIndex = Audio.PlayWave("fuego.wav", 0, 0, LoopStyle.Enabled)

        End If

        Exit Sub

    Case "MF"
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        UserAtributos(1) = Val(Rdata)

    Case "MA"
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        UserAtributos(3) = Val(Rdata)

    Case "MM"
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        UserMaxMAN = Val(Rdata)

        If UserMinMAN > UserMaxMAN Then
            UserMinMAN = UserMaxMAN

        End If

        frmMain.lblManaBar.Caption = UserMinMAN & "/" & UserMaxMAN

    Case "MN"
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        UserMinMAN = Val(Rdata)
        frmMain.lblManaBar.Caption = UserMinMAN & "/" & UserMaxMAN

        If UserMaxMAN > 0 Then
            frmMain.imgMana.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 101)
        Else
            frmMain.imgMana.Width = 0

        End If

        Debug.Print "Llego la mana: " & UserMinMAN
        Exit Sub

    Case "CA"
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        Call CambioDeArea(CByte(ReadField(1, Rdata, 44)), CByte(ReadField(2, Rdata, 44)))
        Exit Sub

    End Select

    Select Case Left$(sData, 3)

    Case "BKW"                  ' >>>>> Pausa :: BKW
        pausa = Not pausa
        Exit Sub

    Case "CLM"  ' Clima Primario
        Rdata = Right$(Rdata, Len(Rdata) - 3)

        If Val(Rdata) = 1 Or Val(Rdata) = 2 Or Val(Rdata) = 3 Then
            Set frmMain.Clima.Picture = Interfaces.Clima_Dia
        Else
            Set frmMain.Clima.Picture = Interfaces.Clima_Noche

        End If

        Call Ambient_SetFinal(Val(Rdata))
        Exit Sub

    Case "CLA"  ' Clima Secundario ON
        Rdata = Right$(Rdata, Len(Rdata) - 3)

        If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, UserPos.Y).Trigger = 2 Or MapData(UserPos.X, _
                                                                                                                       UserPos.Y).Trigger = 4, True, False)

        If bSecondaryAmbient = 0 Then bSecondaryAmbient = Particle_Create(Val(Rdata), -1, -1, -1)

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
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        Call Dialogos.RemoveDialog(Val(Rdata))
        Exit Sub

    Case "CFF"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        charindex = Val(ReadField(1, Rdata, 44))
        CharList(charindex).Particle_Count = Val(ReadField(2, Rdata, 44))
        Call Char_Particle_Create(CharList(charindex).Particle_Count, charindex, Val(ReadField(3, Rdata, 44)))
        Exit Sub

    Case "CFX"                  ' >>>>> Mostrar FX sobre Personaje :: CFX
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        charindex = Val(ReadField(1, Rdata, 44))
        Call SetCharacterFx(charindex, Val(ReadField(2, Rdata, 44)), Val(ReadField(3, Rdata, 44)))
        Exit Sub

    Case "ATG"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        VidaAmarilla = CLng(Val(Rdata) / 40)
        Exit Sub

    Case "VTG"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        VidaVerde = CLng(Val(Rdata) / 40)
        Exit Sub

    Case "ARG"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        Amarilla = Val(Rdata)
        frmMain.lblAgi.Caption = Amarilla
        Exit Sub

    Case "VRG"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        Verde = Val(Rdata)
        frmMain.lblFuerza.Caption = Verde
        Exit Sub

    Case "ARM"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        ArmaMin = Val(ReadField(1, Rdata, 44))
        ArmaMax = Val(ReadField(2, Rdata, 44))

        ArmorMin = Val(ReadField(3, Rdata, 44))
        ArmorMax = Val(ReadField(4, Rdata, 44))

        EscuMin = Val(ReadField(5, Rdata, 44))
        EscuMax = Val(ReadField(6, Rdata, 44))

        CascMin = Val(ReadField(7, Rdata, 44))
        CascMax = Val(ReadField(8, Rdata, 44))
        'MagMin = Val(ReadField(9, Rdata, 44))
        'MagMax = Val(ReadField(10, Rdata, 44))

        frmMain.lblArmor.Caption = ArmorMin & "/" & ArmorMax
        frmMain.lblArma.Caption = ArmaMin & "/" & ArmaMax
        frmMain.lblEscudo.Caption = EscuMin & "/" & EscuMax
        frmMain.LblCasc.Caption = CascMin & "/" & CascMax

        Dim SR As RECT, DR As RECT

        SR.Left = 0
        SR.Top = 0
        SR.Right = 32
        SR.bottom = 32

        DR.Left = 0
        DR.Top = 0
        DR.Right = 32
        DR.bottom = 32

        Dim j As Integer

        For j = 1 To 20

            If Inventario.Equipped(j) = True Then

                ' espada
                If Inventario.ObjType(j) = 2 Then
                    Call DrawGrhtoHdc(frmMain.Picture2.hdc, Inventario.GrhIndex(j), DR)
                    frmMain.Picture2.Refresh

                End If

                ' armadura
                If Inventario.ObjType(j) = 3 Then
                    Call DrawGrhtoHdc(frmMain.Picture1.hdc, Inventario.GrhIndex(j), DR)
                    frmMain.Picture1.Refresh

                End If

                ' casco
                If Inventario.ObjType(j) = 17 Then
                    Call DrawGrhtoHdc(frmMain.Picture3.hdc, Inventario.GrhIndex(j), DR)
                    frmMain.Picture3.Refresh

                End If

                ' escudo
                If Inventario.ObjType(j) = 16 Then
                    Call DrawGrhtoHdc(frmMain.Picture4.hdc, Inventario.GrhIndex(j), DR)
                    frmMain.Picture4.Refresh

                End If

            End If

        Next j

        If frmMain.lblArmor.Caption = "0/0" Then
            frmMain.Picture1.Picture = LoadPicture(vbNullString)
            frmMain.Picture1.Refresh

        End If

        If frmMain.lblArma.Caption = "0/0" Then
            frmMain.Picture2.Picture = LoadPicture(vbNullString)
            frmMain.Picture2.Refresh

        End If

        If frmMain.lblEscudo.Caption = "0/0" Then
            frmMain.Picture4.Picture = LoadPicture(vbNullString)
            frmMain.Picture4.Refresh

        End If

        If frmMain.LblCasc.Caption = "0/0" Then
            frmMain.Picture3.Picture = LoadPicture(vbNullString)
            frmMain.Picture3.Refresh

        End If

        Exit Sub

        'frmMain.lblMagica.Caption = MagMin & "/" & MagMax
    Case "EST"                  ' >>>>> Actualiza Estadisticas de Usuario :: EST
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        UserMaxHP = Val(ReadField(1, Rdata, 44))
        UserMinHP = Val(ReadField(2, Rdata, 44))

        UserMaxMAN = Val(ReadField(3, Rdata, 44))
        UserMinMAN = Val(ReadField(4, Rdata, 44))

        UserMaxSTA = Val(ReadField(5, Rdata, 44))
        UserMinSTA = Val(ReadField(6, Rdata, 44))

        UserGLD = Val(ReadField(7, Rdata, 44))
        UserLvl = Val(ReadField(8, Rdata, 44))
        UserPasarNivel = Val(ReadField(9, Rdata, 44))
        UserExp = Val(ReadField(10, Rdata, 44))
        UserCreditos = Val(ReadField(11, Rdata, 44))
        UserCanjes = Val(ReadField(12, Rdata, 44))

        frmMain.lblVidaBar.Caption = UserMinHP & "/" & UserMaxHP
        frmMain.lblManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
        frmMain.lblStaBar.Caption = UserMinSTA & "/" & UserMaxSTA
        frmMain.exp.Caption = UserExp & "/" & UserPasarNivel

        If UserPasarNivel = 0 Then
            frmMain.lblPorcLvl.Caption = "¡Nivel máximo!"
            frmMain.imgExp.Width = 208
        Else

            If UserExp <> 0 And UserPasarNivel <> 0 Then
                frmMain.imgExp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 153)
                frmMain.lblPorcLvl.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
            Else
                frmMain.imgExp.Width = 153
                frmMain.lblPorcLvl.Caption = "0%"

            End If

        End If

        frmMain.imgVida.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 101)
        frmMain.lblUserName.Caption = UserName
        frmMain.LvlLbl.Caption = UserLvl

        If UserMaxMAN > 0 Then
            frmMain.imgMana.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 101)
        Else
            frmMain.imgMana.Width = 0

        End If

        frmMain.imgEnergia.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 101)
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
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        UserMinHP = CInt(Rdata)
        frmMain.imgVida.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 101)
        frmMain.lblVidaBar.Caption = UserMinHP & "/" & UserMaxHP

        If UserMinHP = 0 Then
            UserEstado = 1
        Else
            UserEstado = 0

        End If

        Exit Sub

    Case "STA"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        UserMinSTA = CInt(Rdata)
        frmMain.imgEnergia.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 104)
        frmMain.lblStaBar.Caption = UserMinSTA & "/" & UserMaxSTA
        Exit Sub

    Case "ORO"
        UserGLD = Val(Right$(Rdata, Len(Rdata) - 3))
        frmMain.GldLbl.Caption = "Oro: " & UserGLD
        Exit Sub

    Case "CRE"
        UserCreditos = Val(Right$(Rdata, Len(Rdata) - 3))
        Exit Sub

    Case "CRJ"
        UserCanjes = Val(Right$(Rdata, Len(Rdata) - 3))
        frmCanjes.Label1(5).Caption = "Tienes " & UserCanjes & " AoMCreditos"
        Exit Sub

    Case "EXP"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        UserExp = Val(Rdata)

        frmMain.exp.Caption = UserExp & "/" & UserPasarNivel

        If UserPasarNivel = 0 Then
            frmMain.lblPorcLvl.Caption = "¡Nivel máximo!"
            frmMain.imgExp.Width = 153
        Else

            If UserExp <> 0 And UserPasarNivel <> 0 Then
                frmMain.imgExp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 153)
                frmMain.lblPorcLvl.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
            Else
                frmMain.imgExp.Width = 153
                frmMain.lblPorcLvl.Caption = "0%"

            End If

        End If

        Exit Sub

    Case "TX"
        Rdata = Right$(Rdata, Len(Rdata) - 2)
        frmMain.MousePointer = 2
        Call AddtoRichTextBox(frmMain.RecTxt, "Elegí la posición.", 100, 100, 120, 0, 0)
        Exit Sub

    Case "T01"                  ' >>>>> TRABAJANDO :: TRA
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        UsingSkill = Val(Rdata)
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
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        Slot = ReadField(1, Rdata, 44)
        Call Inventario.SetItem(Slot, ReadField(2, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), Val(ReadField(6, Rdata, 44)), _
                                Val(ReadField(7, Rdata, 44)), Val(ReadField(8, Rdata, 44)), Val(ReadField(9, Rdata, 44)), Val(ReadField(10, Rdata, 44)), Val( _
                                                                                                                                                         ReadField(11, Rdata, 44)), Val(ReadField(12, Rdata, 44)), ReadField(3, Rdata, 44))

        Exit Sub

        '[KEVIN]-------------------------------------------------------
        '**********************************************************************
    Case "SBO"                 ' >>>>> Actualiza Inventario Banco :: SBO
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        Slot = ReadField(1, Rdata, 44)
        UserBancoInventory(Slot).ObjIndex = ReadField(2, Rdata, 44)
        UserBancoInventory(Slot).Name = ReadField(3, Rdata, 44)
        UserBancoInventory(Slot).Amount = ReadField(4, Rdata, 44)
        UserBancoInventory(Slot).GrhIndex = Val(ReadField(5, Rdata, 44))
        UserBancoInventory(Slot).ObjType = Val(ReadField(6, Rdata, 44))
        UserBancoInventory(Slot).MaxHit = Val(ReadField(7, Rdata, 44))
        UserBancoInventory(Slot).MinHit = Val(ReadField(8, Rdata, 44))
        UserBancoInventory(Slot).MaxDef = Val(ReadField(9, Rdata, 44))
        UserBancoInventory(Slot).MinDef = Val(ReadField(10, Rdata, 44))

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
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        Slot = ReadField(1, Rdata, 44)
        UserHechizos(Slot) = ReadField(2, Rdata, 44)

        If Slot > frmMain.hlst.ListCount Then
            frmMain.hlst.AddItem ReadField(3, Rdata, 44)
        Else
            frmMain.hlst.List(Slot - 1) = ReadField(3, Rdata, 44)

        End If

        Exit Sub

    Case "ATR"               ' >>>>> Recibir Atributos del Personaje :: ATR
        Rdata = Right$(Rdata, Len(Rdata) - 3)

        For i = 1 To NUMATRIBUTOS
            UserAtributos(i) = Val(ReadField(i, Rdata, 44))
        Next i

        LlegaronAtrib = True
        Exit Sub

    Case "LAH"
        Rdata = Right$(Rdata, Len(Rdata) - 3)

        For m = 0 To UBound(ArmasHerrero)
            ArmasHerrero(m) = 0
        Next m

        i = 1
        m = 0
        Do
            cad$ = ReadField(i, Rdata, 44)
            ArmasHerrero(m) = Val(ReadField(i + 1, Rdata, 44))

            If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
            i = i + 2
            m = m + 1
        Loop While cad$ <> ""

        Exit Sub

    Case "LAR"
        Rdata = Right$(Rdata, Len(Rdata) - 3)

        For m = 0 To UBound(ArmadurasHerrero)
            ArmadurasHerrero(m) = 0
        Next m

        i = 1
        m = 0
        Do
            cad$ = ReadField(i, Rdata, 44)
            ArmadurasHerrero(m) = Val(ReadField(i + 1, Rdata, 44))

            If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
            i = i + 2
            m = m + 1
        Loop While cad$ <> ""

        Exit Sub

    Case "OBR"
        Rdata = Right$(Rdata, Len(Rdata) - 3)

        For m = 0 To UBound(ObjCarpintero)
            ObjCarpintero(m) = 0
        Next m

        i = 1
        m = 0
        Do
            cad$ = ReadField(i, Rdata, 44)
            ObjCarpintero(m) = Val(ReadField(i + 1, Rdata, 44))

            If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
            i = i + 2
            m = m + 1
        Loop While cad$ <> ""

        Exit Sub

    Case "DOK"               ' >>>>> Descansar OK :: DOK
        UserDescansar = Not UserDescansar
        Exit Sub

    Case "NIX"               ' >>>>> castillo nix

        If frmMain.Tnix.Enabled = False Then
            frmMain.Nix.Visible = True
            frmMain.Tnix.Enabled = True

        End If

        Exit Sub

    Case "SPL"
        Rdata = Right$(Rdata, Len(Rdata) - 3)

        For i = 1 To Val(ReadField(1, Rdata, 44))
            frmSpawnList.lstCriaturas.AddItem ReadField(i + 1, Rdata, 44)
        Next i

        frmSpawnList.Show , frmMain
        Exit Sub

    Case "LLE"
        frmMain.Label10.Visible = True
        Exit Sub

    Case "ERR"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        frmPasswdSinPadrinos.MousePointer = 1

        If Not frmCrearPersonaje.Visible Then

            frmMain.Socket1.Disconnect

        End If

        If frmConnect.Visible = True Then
            MsgBox Rdata

            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup

            frmMain.Socket1.Disconnect

        Else
            MsgBox (Rdata)
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup

        End If

        frmConnect.MousePointer = 1
        Exit Sub

    End Select

    Select Case Left$(sData, 4)


    Case "TEST"    '  <--- Estadisticas al clickearlo by gohan ssj
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        UserClick = ReadField(1, Rdata, 44)
        ClickMatados = Val(ReadField(2, Rdata, 44))
        ClickClase = ReadField(3, Rdata, 44)
        Estadisticas = True
        frmMain.Clickeado.Enabled = True
        TiempoEst = 3
        Exit Sub

        ' CHOTS | el "proSesos" no es un error de ortografia, es para diferenciar los 2 comandos :)
        ' CHOTS | el "noNbre" y el "nomVre" tmpk ¬¬ jaja

    Case "MATA"    ' CHOTS | Matar Procesos
        Dim Procesoo As String
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        Procesoo = ReadField(1, Rdata, 44)
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

    Case "ULLA"    ' castillo ulla

        If frmMain.Tulla.Enabled = False Then
            frmMain.Ulla.Visible = True
            frmMain.Tulla.Enabled = True

        End If

    Case "LEMU"    ' castillo lemuria

        If frmMain.Tlemu.Enabled = False Then
            frmMain.Lemu.Visible = True
            frmMain.Tlemu.Enabled = True

        End If

    Case "TALE"  ' castillo tale

        If frmMain.Ttale.Enabled = False Then
            frmMain.Tale.Visible = True
            frmMain.Ttale.Enabled = True

        End If

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
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        Call SendData("LEFT" & Rdata)
        Exit Sub

    Case "PART"
        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ENTRAR_PARTY_1 & ReadField(1, Rdata, 44) & MENSAJE_ENTRAR_PARTY_2, 0, 255, 0, False, _
                              False, False)
        Exit Sub

    Case "CEGU"
        UserCiego = True
        Exit Sub

    Case "DUMB"
        UserEstupido = True
        Exit Sub

    Case "NATR"    ' >>>>> Recibe atributos para el nuevo personaje
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        UserAtributos(1) = ReadField(1, Rdata, 44)
        UserAtributos(2) = ReadField(2, Rdata, 44)
        UserAtributos(3) = ReadField(3, Rdata, 44)
        UserAtributos(4) = ReadField(4, Rdata, 44)
        UserAtributos(5) = ReadField(5, Rdata, 44)

        frmCrearPersonaje.lbFuerza.Caption = UserAtributos(1)
        frmCrearPersonaje.lbInteligencia.Caption = UserAtributos(2)
        frmCrearPersonaje.lbAgilidad.Caption = UserAtributos(3)
        frmCrearPersonaje.lbCarisma.Caption = UserAtributos(4)
        frmCrearPersonaje.lbConstitucion.Caption = UserAtributos(5)

        Exit Sub

    Case "MCAR"              ' >>>>> Mostrar Cartel :: MCAR
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        Call InitCartel(ReadField(1, Rdata, 176), CInt(ReadField(2, Rdata, 176)))
        Exit Sub

    Case "NPCI"              ' >>>>> Recibe Item del Inventario de un NPC :: NPCI
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        NPCInvDim = NPCInvDim + 1
        NPCInventory(NPCInvDim).Name = ReadField(1, Rdata, 44)
        NPCInventory(NPCInvDim).Amount = ReadField(2, Rdata, 44)
        NPCInventory(NPCInvDim).Valor = ReadField(3, Rdata, 44)
        NPCInventory(NPCInvDim).GrhIndex = ReadField(4, Rdata, 44)
        NPCInventory(NPCInvDim).ObjIndex = ReadField(5, Rdata, 44)
        NPCInventory(NPCInvDim).ObjType = ReadField(6, Rdata, 44)
        NPCInventory(NPCInvDim).MaxHit = ReadField(7, Rdata, 44)
        NPCInventory(NPCInvDim).MinHit = ReadField(8, Rdata, 44)
        NPCInventory(NPCInvDim).Def = ReadField(9, Rdata, 44)
        NPCInventory(NPCInvDim).C1 = ReadField(10, Rdata, 44)
        NPCInventory(NPCInvDim).C2 = ReadField(11, Rdata, 44)
        NPCInventory(NPCInvDim).C3 = ReadField(12, Rdata, 44)
        NPCInventory(NPCInvDim).C4 = ReadField(13, Rdata, 44)
        NPCInventory(NPCInvDim).C5 = ReadField(14, Rdata, 44)
        NPCInventory(NPCInvDim).C6 = ReadField(15, Rdata, 44)
        NPCInventory(NPCInvDim).C7 = ReadField(16, Rdata, 44)
        frmComerciar.List1(0).AddItem NPCInventory(NPCInvDim).Name
        Exit Sub

    Case "NPCC"     '>>> Recibe Item del Inventario AoMCreditos
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        CREDInvDim = CREDInvDim + 1
        CREDInventory(CREDInvDim).Name = ReadField(1, Rdata, 44)
        CREDInventory(CREDInvDim).ObjIndex = ReadField(2, Rdata, 44)
        CREDInventory(CREDInvDim).Monedas = ReadField(3, Rdata, 44)
        CREDInventory(CREDInvDim).GrhIndex = ReadField(4, Rdata, 44)
        CREDInventory(CREDInvDim).Def = ReadField(5, Rdata, 44)
        CREDInventory(CREDInvDim).MaxHit = ReadField(6, Rdata, 44)
        CREDInventory(CREDInvDim).MinHit = ReadField(7, Rdata, 44)
        CREDInventory(CREDInvDim).ObjType = ReadField(8, Rdata, 44)
        frmCreditos.List1(0).AddItem CREDInventory(CREDInvDim).Name
        Exit Sub

    Case "NPCJ"     '>>> Recibe Item del Inventario AoMCreditos
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        CANJInvDim = CANJInvDim + 1
        CANJInventory(CANJInvDim).Name = ReadField(1, Rdata, 44)
        CANJInventory(CANJInvDim).ObjIndex = ReadField(2, Rdata, 44)
        CANJInventory(CANJInvDim).Monedas = ReadField(3, Rdata, 44)
        CANJInventory(CANJInvDim).GrhIndex = ReadField(4, Rdata, 44)
        CANJInventory(CANJInvDim).Def = ReadField(5, Rdata, 44)
        CANJInventory(CANJInvDim).MaxHit = ReadField(6, Rdata, 44)
        CANJInventory(CANJInvDim).MinHit = ReadField(7, Rdata, 44)
        CANJInventory(CANJInvDim).ObjType = ReadField(8, Rdata, 44)
        CANJInventory(CANJInvDim).Cantidad = ReadField(9, Rdata, 44)
        frmCanjes.List1(0).AddItem CANJInventory(CANJInvDim).Name
        Exit Sub

    Case "SOPO"
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        frmSoporteGm.Text1.Text = ReadField(1, Rdata, 2)    ' pregunta
        frmSoporteGm.Label1.Caption = ReadField(2, Rdata, 2)    ' nombre
        Exit Sub

    Case "RESP"
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        frmSoporteResp.Text1.Text = ReadField(1, Rdata, 44)    ' respuesta

        Exit Sub

    Case "EHYS"              ' Actualiza Hambre y Sed :: EHYS
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        UserMaxAGU = 100
        UserMaxHAM = 100
        UserMinAGU = Val(ReadField(1, Rdata, 44))
        UserMinHAM = Val(ReadField(2, Rdata, 44))
        frmMain.imgSed.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 93)
        frmMain.imgComida.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 93)
        frmMain.lblSedBar.Caption = UserMinAGU & "/" & UserMaxAGU
        frmMain.lblHamBar.Caption = UserMinHAM & "/" & UserMaxHAM
        Exit Sub

    Case "FAMA"             ' >>>>> Recibe Fama de Personaje :: FAMA
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        UserReputacion.AsesinoRep = Val(ReadField(1, Rdata, 44))
        UserReputacion.BandidoRep = Val(ReadField(2, Rdata, 44))
        UserReputacion.BurguesRep = Val(ReadField(3, Rdata, 44))
        UserReputacion.LadronesRep = Val(ReadField(4, Rdata, 44))
        UserReputacion.NobleRep = Val(ReadField(5, Rdata, 44))
        UserReputacion.PlebeRep = Val(ReadField(6, Rdata, 44))
        UserReputacion.Promedio = Val(ReadField(7, Rdata, 44))
        LlegoFama = True
        Exit Sub

    Case "LCRK"
        Rdata = Right$(Rdata, Len(Rdata) - 4)

        frmListClanes.ListClan.AddItem Rdata

        Exit Sub

    Case "MEST"    ' >>>>>> Mini Estadisticas :: MEST
        Rdata = Right$(Rdata, Len(Rdata) - 4)

        With UserEstadisticas
            .CiudadanosMatados = Val(ReadField(1, Rdata, 44))
            .CriminalesMatados = Val(ReadField(2, Rdata, 44))
            .UsuariosMatados = Val(ReadField(3, Rdata, 44))
            .NpcsMatados = Val(ReadField(4, Rdata, 44))
            .Clase = ReadField(5, Rdata, 44)
            .PenaCarcel = Val(ReadField(6, Rdata, 44))
            .Raza = ReadField(7, Rdata, 44)
            .PuntosClan = Val(ReadField(8, Rdata, 44))
            .Name = ReadField(9, Rdata, 44)
            .Genero = ReadField(10, Rdata, 44)
            .PuntosRetos = Val(ReadField(11, Rdata, 44))
            .PuntosTorneos = Val(ReadField(12, Rdata, 44))
            .PuntosDuelos = Val(ReadField(13, Rdata, 44))
            .Stats.Nivel = Val(ReadField(14, Rdata, 44))
            .Stats.MaxExp = Val(ReadField(15, Rdata, 44))
            .Stats.MinExp = Val(ReadField(16, Rdata, 44))
            .Stats.MinHP = Val(ReadField(17, Rdata, 44))
            .Stats.MaxHP = Val(ReadField(18, Rdata, 44))
            .Stats.MinMan = Val(ReadField(19, Rdata, 44))
            .Stats.MaxMan = Val(ReadField(20, Rdata, 44))
            .Stats.MinSta = Val(ReadField(21, Rdata, 44))
            .Stats.MaxSta = Val(ReadField(22, Rdata, 44))
            .Stats.Oro = Val(ReadField(23, Rdata, 44))
            .Stats.Banco = Val(ReadField(24, Rdata, 44))
            .pos.Map = ReadField(25, Rdata, 44)
            .pos.PosX = Val(ReadField(26, Rdata, 44))
            .pos.PosY = Val(ReadField(27, Rdata, 44))
            .Stats.SkillPoins = Val(ReadField(28, Rdata, 44))
            .ParticipoClan = Val(ReadField(29, Rdata, 44))
            .AbbadonMatados = Val(ReadField(30, Rdata, 44))
            .CleroMatados = Val(ReadField(31, Rdata, 44))
            .TinieblaMatados = Val(ReadField(32, Rdata, 44))
            .TemplarioMatados = Val(ReadField(33, Rdata, 44))
            .Faccion.Armada = ReadField(34, Rdata, 44)
            .Faccion.Reenlistado = Val(ReadField(35, Rdata, 44))
            .Faccion.Recompensas = Val(ReadField(36, Rdata, 44))
            .Faccion.CiudadanosMatados = Val(ReadField(37, Rdata, 44))
            .Faccion.CriminalesMatados = Val(ReadField(38, Rdata, 44))
            .Faccion.FEnlistado = ReadField(39, Rdata, 44)
        End With

        Exit Sub

    Case "SUNI"             ' >>>>> Subir Nivel :: SUNI
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        SkillPoints = SkillPoints + Val(Rdata)
        frmMain.imgSkillpts.Visible = True
        Exit Sub

    Case "NENE"             ' >>>>> Nro de Personajes :: NENE
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        AddtoRichTextBox frmMain.RecTxt, MENSAJE_NENE & Rdata, 255, 255, 255, 0, 0
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
        Rdata = Right$(Rdata, Len(Rdata) - 4)
        frmForo.List.AddItem ReadField(1, Rdata, 176)
        frmForo.Text(frmForo.List.ListCount - 1).Text = ReadField(2, Rdata, 176)
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
        Rdata = Right$(Rdata, Len(Rdata) - 5)
        UserMaxHP = Val(Rdata)

        If UserMinHP > UserMaxHP Then
            UserMinHP = UserMaxHP

        End If

        frmMain.lblVidaBar.Caption = UserMinHP & "/" & UserMaxHP
        frmMain.imgVida.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 101)

        Exit Sub

    Case "MXMAN"
        Rdata = Right$(Rdata, Len(Rdata) - 5)
        UserMaxMAN = Val(Rdata)

        If UserMinMAN > UserMaxMAN Then
            UserMinMAN = UserMaxMAN

        End If

        frmMain.lblManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
        frmMain.imgMana.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 101)

        Exit Sub

    Case "NOPRT"
        Rdata = Right$(Rdata, Len(Rdata) - 5)
        charindex = Val(ReadField(1, Rdata, 44))
        CharList(charindex).PartyIndex = Val(ReadField(2, Rdata, 44))

        Exit Sub

    Case "NOVER"
        Rdata = Right$(Rdata, Len(Rdata) - 5)
        charindex = Val(ReadField(1, Rdata, 44))
        CharList(charindex).Invisible = (Val(ReadField(2, Rdata, 44)) = 1)
        CharList(charindex).PartyIndex = Val(ReadField(3, Rdata, 44))

        Exit Sub

    Case "ZMOTD"
        Rdata = Right$(Rdata, Len(Rdata) - 5)
        frmCambiaMotd.Show , frmMain
        frmCambiaMotd.txtMotd.Text = Rdata
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
        Rdata = Right$(Rdata, Len(Rdata) - 6)

        For i = 1 To NUMSKILLS
            UserSkills(i) = Val(ReadField(i, Rdata, 44))
        Next i

        LlegaronSkills = True
        Exit Sub

    Case "LSTCRI"
        Rdata = Right$(Rdata, Len(Rdata) - 6)

        For i = 1 To Val(ReadField(1, Rdata, 44))
            frmEntrenador.lstCriaturas.AddItem ReadField(i + 1, Rdata, 44)
        Next i

        frmEntrenador.Show , frmMain
        Exit Sub

    End Select

    Select Case Left$(sData, 7)

    Case "GUILDNE"
        Rdata = Right$(Rdata, Len(Rdata) - 7)
        Call frmGuildNews.ParseGuildNews(Rdata)
        Exit Sub

    Case "PEACEDE"  'detalles de paz
        Rdata = Right$(Rdata, Len(Rdata) - 7)
        Call frmUserRequest.recievePeticion(Rdata)
        Exit Sub

    Case "ALLIEDE"  'detalles de paz
        Rdata = Right$(Rdata, Len(Rdata) - 7)
        Call frmUserRequest.recievePeticion(Rdata)
        Exit Sub

    Case "ALLIEPR"  'lista de prop de alianzas
        Rdata = Right$(Rdata, Len(Rdata) - 7)
        Call frmPeaceProp.ParseAllieOffers(Rdata)

    Case "PEACEPR"  'lista de prop de paz
        Rdata = Right$(Rdata, Len(Rdata) - 7)
        Call frmPeaceProp.ParsePeaceOffers(Rdata)
        Exit Sub

    Case "CHRINFO"
        Rdata = Right$(Rdata, Len(Rdata) - 7)
        Call frmCharInfo.parseCharInfo(Rdata)
        Exit Sub

    Case "LEADERI"
        Rdata = Right$(Rdata, Len(Rdata) - 7)
        Call frmGuildLeader.ParseLeaderInfo(Rdata)
        Exit Sub

    Case "CLKNDET"
        Rdata = Right$(Rdata, Len(Rdata) - 7)
        Call frmGuildBrief.ParseGuildInfo(Rdata)
        Exit Sub

    Case "SHOWFUN"
        CreandoClan = True
        frmGuildFoundation.Show , frmMain
        Exit Sub

    Case "PARADOW"         ' >>>>> Paralizar OK :: PARADOK
        UserParalizado = Not UserParalizado
        Exit Sub

    Case "PETICIO"         ' >>>>> Paralizar OK :: PARADOK
        Rdata = Right$(Rdata, Len(Rdata) - 7)
        Call frmUserRequest.recievePeticion(Rdata)
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
            Rdata = Right$(Rdata, Len(Rdata) - 7)

            If ReadField(2, Rdata, 44) = "0" Then
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

            Rdata = Right$(Rdata, Len(Rdata) - 7)

            If ReadField(2, Rdata, 44) = "0" Then
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
        Rdata = Right$(Rdata, Len(Rdata) - 5)

        Mayores.CiudadanoMaxNivel = ReadField(1, Rdata, 44)
        Mayores.CriminalMaxNivel = ReadField(2, Rdata, 44)
        Mayores.MaxCiudadano = ReadField(3, Rdata, 44)
        Mayores.MaxCriminal = ReadField(4, Rdata, 44)
        Mayores.OnlineCiudadano = ReadField(5, Rdata, 44)
        Mayores.OnlineCriminal = ReadField(6, Rdata, 44)
        Mayores.MaxOroOnline = ReadField(7, Rdata, 44)
        Mayores.MaxOro = ReadField(8, Rdata, 44)


        Call frmMayor.Show(vbModeless, frmMain)
        Exit Sub

    End Select

    '[Alejo]
    Select Case UCase$(Left$(Rdata, 9))

    Case "COMUSUPET"
        Rdata = Right$(Rdata, Len(Rdata) - 9)
        OtroInventario(1).ObjIndex = ReadField(2, Rdata, 44)
        OtroInventario(1).Name = ReadField(3, Rdata, 44)
        OtroInventario(1).Amount = ReadField(4, Rdata, 44)
        OtroInventario(1).Equipped = ReadField(5, Rdata, 44)
        OtroInventario(1).GrhIndex = Val(ReadField(6, Rdata, 44))
        OtroInventario(1).ObjType = Val(ReadField(7, Rdata, 44))
        OtroInventario(1).MaxHit = Val(ReadField(8, Rdata, 44))
        OtroInventario(1).MinHit = Val(ReadField(9, Rdata, 44))
        OtroInventario(1).MaxDef = Val(ReadField(10, Rdata, 44))
        OtroInventario(1).MinDef = Val(ReadField(11, Rdata, 44))
        OtroInventario(1).Valor = Val(ReadField(12, Rdata, 44))

        frmComerciarUsu.List2.Clear

        frmComerciarUsu.List2.AddItem OtroInventario(1).Name
        frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = OtroInventario(1).Amount

        frmComerciarUsu.lblEstadoResp.Visible = False

        Exit Sub

    End Select

    Call HandleData2(Rdata)

    ';Call LogCustom("Unhandled data: " & Rdata)

End Sub

Sub SendData(ByVal sdData As String)

    'No enviamos nada si no estamos conectados
    If Not frmMain.Socket1.Connected Then Exit Sub

    Dim AuxCmd As String
    AuxCmd = UCase$(Left$(sdData, 5))
    
    If AuxCmd = "/PING" Then TimerPing(1) = GetTickCount()

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
    Dim version As String
    
    version = App.Major & "." & App.Minor & "." & App.Revision

    Select Case EstadoLogin
   
        Case E_MODO.Normal
            
            Call SendData("MARAKA" & UserName & "," & UserPassword & "," & version & "," & HDD & "," & "0")
          
        Case E_MODO.CrearNuevoPj
            Call SendData("TIRDAD" & UserFuerza & "," & UserAgilidad _
               & "," & UserInteligencia & "," & UserCarisma & "," & UserConstitucion)
            Call SendData("ZORRON" & UserName & "," & UserPassword & "," & version & "," & UserRaza & "," & UserSexo & "," & UserClase & "," & _
                UserBanco & "," & UserPersonaje & "," & UserEmail & "," & UserHogar & "," & HDD)

        Case E_MODO.Dados
            frmCrearPersonaje.Show
            Call SendData("TIRDAD")

    End Select

End Sub

