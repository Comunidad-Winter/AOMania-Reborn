Attribute VB_Name = "ES"
Option Explicit
''Function EsAdmin(ByVal Name As String) As Boolean
''Function EsDios(ByVal Name As String) As Boolean
''Function EsSemiDios(ByVal Name As String) As Boolean
''Function EsConsejero(ByVal Name As String) As Boolean
''Function EsRolesMaster(ByVal Name As String) As Boolean

Public Administradores As clsIniManager

Function GmTrue(ByRef Name As String, ByRef HDD As String)
    Dim NumGms As Integer
   
    NumGms = val(GetVar(App.Path & "\dat\gmsmac.dat", "INIT", "Num"))
   
    If NumGms = 0 Then
        GmTrue = False
        Exit Function

    End If
   
    Dim i      As Integer
    Dim GMName As String
    Dim GmMac  As String
   
    For i = 1 To NumGms
        
        GMName = GetVar(App.Path & "\dat\gmsmac.dat", "GM" & i, "Nombre")
        GmMac = GetVar(App.Path & "\dat\gmsmac.dat", "GM" & i, "MAC")
        
        If UCase(Name) = GMName Then
            If MD5String(HDD) = GmMac Then
                GmTrue = True
                Exit Function

            End If

        End If
        
    Next i
   
    GmTrue = False

End Function

Public Function EsGmChar(ByRef Name As String) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 27/03/2011
    'Returns true if char is administrative user.
    '***************************************************
    
    Dim EsGm As Boolean
    
    ' Admin?
    EsGm = EsAdmin(Name)

    ' Dios?
    If Not EsGm Then EsGm = EsDios(Name)

    ' Semidios?
    If Not EsGm Then EsGm = EsSemiDios(Name)

    ' Consejero?
    If Not EsGm Then EsGm = EsConsejero(Name)

    EsGmChar = EsGm

End Function

Function EsHDD(ByRef Name As String, ByRef HDD As String) As Boolean
    Dim tHDD As String
    
    tHDD = Administradores.GetValue("HDD", Name)
    
    If Len(tHDD) <= 0 Then
        EsHDD = True
        Exit Function

    End If
    
    EsHDD = (StrComp(tHDD, HDD) = 0)

End Function

Function EsAdmin(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsAdmin = (val(Administradores.GetValue("Admin", Name)) = 1)

End Function

Function EsDios(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsDios = (val(Administradores.GetValue("Dios", Name)) = 1)

End Function

Function EsSemiDios(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsSemiDios = (val(Administradores.GetValue("SemiDios", Name)) = 1)

End Function

Function EsConsejero(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsConsejero = (val(Administradores.GetValue("Consejero", Name)) = 1)

End Function

Function EsRolesMaster(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsRolesMaster = (val(Administradores.GetValue("RM", Name)) = 1)

End Function

Public Sub LoadAdministrativeUsers()
    'Admines     => Admin
    'Dioses      => Dios
    'SemiDioses  => SemiDios
    'Especiales  => Especial
    'Consejeros  => Consejero
    'RoleMasters => RM

    'Si esta mierda tuviese array asociativos el código sería tan lindo.
    Dim buf  As Integer
    Dim i    As Long
    Dim tStr As String
    Dim Name As String
    Dim HDD  As String
    
    ' Public container
    Set Administradores = New clsIniManager
    
    ' Server ini info file
    Dim ServerIni As clsIniManager
    Set ServerIni = New clsIniManager
    
    Call ServerIni.Initialize(IniPath & "Server.ini")
       
    ' Admines
    buf = val(ServerIni.GetValue("INIT", "Administradores"))
    
    For i = 1 To buf
        tStr = UCase$(ServerIni.GetValue("Administradores", "Admin" & i))
        
        Name = ReadField(1, tStr, Asc("@"))
        HDD = ReadField(2, tStr, Asc("@"))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Admin", Name, "1")
        Call Administradores.ChangeValue("HDD", Name, HDD)

    Next i
    
    ' Dioses
    buf = val(ServerIni.GetValue("INIT", "Dioses"))
    
    For i = 1 To buf
        tStr = UCase$(ServerIni.GetValue("Dioses", "Dios" & i))
        
        Name = ReadField(1, tStr, Asc("@"))
        HDD = ReadField(2, tStr, Asc("@"))

        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Dios", Name, "1")
        Call Administradores.ChangeValue("HDD", Name, HDD)

    Next i
        
    ' SemiDioses
    buf = val(ServerIni.GetValue("INIT", "SemiDioses"))
    
    For i = 1 To buf
        tStr = UCase$(ServerIni.GetValue("SemiDioses", "SemiDios" & i))
        
        Name = ReadField(1, tStr, Asc("@"))
        HDD = ReadField(2, tStr, Asc("@"))

        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("SemiDios", Name, "1")
        Call Administradores.ChangeValue("HDD", Name, HDD)
    Next i
    
    ' Consejeros
    buf = val(ServerIni.GetValue("INIT", "Consejeros"))
        
    For i = 1 To buf
        tStr = UCase$(ServerIni.GetValue("Consejeros", "Consejero" & i))
        
        Name = ReadField(1, tStr, Asc("@"))
        HDD = ReadField(2, tStr, Asc("@"))

        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Consejero", Name, "1")
        Call Administradores.ChangeValue("HDD", Name, HDD)
    Next i
    
    ' RolesMasters
    buf = val(ServerIni.GetValue("INIT", "RolesMasters"))
        
    For i = 1 To buf
        tStr = UCase$(ServerIni.GetValue("RolesMasters", "RM" & i))
        
        Name = ReadField(1, tStr, Asc("@"))
        HDD = ReadField(2, tStr, Asc("@"))

        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("RM", Name, "1")
        Call Administradores.ChangeValue("HDD", Name, HDD)
    Next i
    
    Set ServerIni = Nothing
    
End Sub

Public Sub CargarSpawnList()

    Dim n As Integer, loopc As Integer
    n = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(n) As tCriaturasEntrenador

    For loopc = 1 To n
        SpawnList(loopc).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & loopc))
        SpawnList(loopc).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & loopc)
    Next loopc
    
End Sub

Public Function TxtDimension(ByVal Name As String) As Long

    Dim n As Integer, cad As String, Tam As Long
    n = FreeFile(1)
    Open Name For Input As #n
    Tam = 0

    Do While Not EOF(n)
        Tam = Tam + 1
        Line Input #n, cad
    Loop
    Close n
    TxtDimension = Tam

End Function

Public Sub CargarForbidenWords()

    ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
    Dim n As Integer, i As Integer
    n = FreeFile(1)
    Open DatPath & "NombresInvalidos.txt" For Input As #n

    For i = 1 To UBound(ForbidenNames)
        Line Input #n, ForbidenNames(i)
    Next i

    Close n

End Sub

Public Sub CargarHechizos()


    On Error GoTo errhandler

    If frmMain.Visible Then frmMain.txStatus.caption = "Cargando Hechizos."

    Dim Hechizo As Integer
    Dim Leer    As New clsIniManager
    Dim i As Integer

    Call Leer.Initialize(DatPath & "Hechizos.dat")

    'obtiene el numero de hechizos
    NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))
    ReDim Hechizos(1 To NumeroHechizos) As tHechizo

    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumeroHechizos
    frmCargando.cargar.value = 0

    'Llena la lista
    For Hechizo = 1 To NumeroHechizos

        Hechizos(Hechizo).nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
        Hechizos(Hechizo).Desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
        Hechizos(Hechizo).PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
    
        Hechizos(Hechizo).HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
        Hechizos(Hechizo).TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
        Hechizos(Hechizo).PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
    
        Hechizos(Hechizo).Tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
        Hechizos(Hechizo).WAV = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
        Hechizos(Hechizo).FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
    
        Hechizos(Hechizo).loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
    
        Hechizos(Hechizo).Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
    
        Hechizos(Hechizo).SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
        Hechizos(Hechizo).MinHP = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
        Hechizos(Hechizo).MaxHP = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
    
        Hechizos(Hechizo).SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
        Hechizos(Hechizo).MinMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
        Hechizos(Hechizo).ManMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
    
        Hechizos(Hechizo).SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
        Hechizos(Hechizo).MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
        Hechizos(Hechizo).MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
    
        Hechizos(Hechizo).SubeHam = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
        Hechizos(Hechizo).MinHam = val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
        Hechizos(Hechizo).MaxHam = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
    
        Hechizos(Hechizo).SubeSed = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
        Hechizos(Hechizo).MinSed = val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
        Hechizos(Hechizo).MaxSed = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
        
    
        Hechizos(Hechizo).SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
        Hechizos(Hechizo).MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
        Hechizos(Hechizo).MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
    
        Hechizos(Hechizo).SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
        Hechizos(Hechizo).MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
        Hechizos(Hechizo).MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
    
        Hechizos(Hechizo).SubeCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
        Hechizos(Hechizo).MinCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
        Hechizos(Hechizo).MaxCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
    
        Hechizos(Hechizo).Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
        Hechizos(Hechizo).Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
        Hechizos(Hechizo).Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
        Hechizos(Hechizo).RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
        Hechizos(Hechizo).RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
        Hechizos(Hechizo).RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
        Hechizos(Hechizo).RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
    
        Hechizos(Hechizo).CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
        Hechizos(Hechizo).Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
        Hechizos(Hechizo).Maldicion = val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
        Hechizos(Hechizo).RemoverMaldicion = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
        Hechizos(Hechizo).Bendicion = val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
        Hechizos(Hechizo).Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
        Hechizos(Hechizo).ExclusivoClase = Leer.GetValue("Hechizo" & Hechizo, "ExclusivoClase")
    
        Hechizos(Hechizo).Ceguera = val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
        Hechizos(Hechizo).Estupidez = val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
    
        Hechizos(Hechizo).invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
        Hechizos(Hechizo).NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
        Hechizos(Hechizo).Cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
        Hechizos(Hechizo).Mimetiza = val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
    
        Hechizos(Hechizo).Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
        Hechizos(Hechizo).ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
    
        Hechizos(Hechizo).MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
        Hechizos(Hechizo).ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
    
        'Barrin 30/9/03
        Hechizos(Hechizo).StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
    
        Hechizos(Hechizo).Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
        frmCargando.cargar.value = frmCargando.cargar.value + 1
    
        Hechizos(Hechizo).NeedStaff = val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
        Hechizos(Hechizo).StaffAffected = CBool(val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
        
        For i = 1 To 20
            Hechizos(Hechizo).ClaseProhibida(i) = StringToClase(Leer.GetValue("Hechizo" & Hechizo, "CP" & i))
        Next i
        
    Next Hechizo

    Set Leer = Nothing
    Exit Sub

errhandler:
    MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.Description
 
End Sub

Sub LoadMotd()
    Dim i As Integer

    MaxLines = val(GetVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines"))
    ReDim MOTD(1 To MaxLines)

    For i = 1 To MaxLines
        MOTD(i).texto = GetVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & i)
        MOTD(i).Formato = ""
    Next i

End Sub

Public Sub DoBackUp()

    'Call LogTarea("Sub DoBackUp")
    haciendoBK = True
    Dim i As Integer

    ' Lo saco porque elimina elementales y mascotas - Maraxus
    ''''''''''''''lo pongo aca x sugernecia del yind
    'For i = 1 To LastNPC
    '    If Npclist(i).flags.NPCActive Then
    '        If Npclist(i).Contadores.TiempoExistencia > 0 Then
    '            Call MuereNpc(i, 0)
    '        End If
    '    End If
    'Next i
    '''''''''''/'lo pongo aca x sugernecia del yind

    Call SendData(SendTarget.toall, 0, 0, "BKW")

    Call LimpiarObjs
    Call modGuilds.v_RutinaElecciones
    Call WorldSave
 
    'Call ResetCentinelaInfo     'Reseteamos al centinela

    Call SendData(SendTarget.toall, 0, 0, "BKW")

    'Call EstadisticasWeb.Informar(EVENTO_NUEVO_CLAN, 0)

    haciendoBK = False

    'Log
    On Error Resume Next

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time
    Close #nfile

End Sub

Public Sub GrabarMapa(ByVal Map As Long, ByRef MAPFILE As String)

    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2011
    '10/08/2010 - Pato: Implemento el clsByteBuffer para el grabado de mapas
    '28/10/2010:ZaMa - Ahora no se hace backup de los pretorianos.
    '12/01/2011 - Amraphen: Ahora no se hace backup de NPCs prohibidos (Pretorianos, Mascotas, Invocados y Centinela)
    '***************************************************

    'On Error Resume Next

    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y           As Long
    Dim X           As Long
    Dim ByFlags     As Byte
    Dim loopc       As Long
    
    Dim MapWriter   As clsByteBuffer
    Dim InfWriter   As clsByteBuffer
    Dim IniManager  As clsIniManager

    Set MapWriter = New clsByteBuffer
    Set InfWriter = New clsByteBuffer
    Set IniManager = New clsIniManager
    
    If FileExist(MAPFILE & ".map", vbNormal) Then
        Kill MAPFILE & ".map"

    End If
    
    If FileExist(MAPFILE & ".inf", vbNormal) Then
        Kill MAPFILE & ".inf"

    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open MAPFILE & ".Map" For Binary As FreeFileMap
    
    Call MapWriter.initializeWriter(FreeFileMap)
    
    'Open .inf file
    FreeFileInf = FreeFile
    Open MAPFILE & ".Inf" For Binary As FreeFileInf
    
    Call InfWriter.initializeWriter(FreeFileInf)
    
    'map Header
    Call MapWriter.putInteger(MapInfo(Map).MapVersion)
        
    Call MapWriter.putString(MiCabecera.Desc, False)
    Call MapWriter.putLong(MiCabecera.crc)
    Call MapWriter.putLong(MiCabecera.MagicWord)
    
    Call MapWriter.putDouble(0)
    
    'inf Header
    Call InfWriter.putDouble(0)
    Call InfWriter.putInteger(0)
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            With MapData(Map, X, Y)
                ByFlags = 0
                
                If .Blocked Then ByFlags = ByFlags Or 1
                If .Graphic(2) Then ByFlags = ByFlags Or 2
                If .Graphic(3) Then ByFlags = ByFlags Or 4
                If .Graphic(4) Then ByFlags = ByFlags Or 8
                If .trigger Then ByFlags = ByFlags Or 16
                
                Call MapWriter.putByte(ByFlags)
                
                Call MapWriter.putLong(.Graphic(1))
                
                For loopc = 2 To 4

                    If .Graphic(loopc) Then Call MapWriter.putLong(.Graphic(loopc))
                Next loopc
                
                If .trigger Then Call MapWriter.putInteger(CInt(.trigger))
                
                '.inf file
                ByFlags = 0
                
                If .OBJInfo.ObjIndex > 0 Then
                    If ObjData(.OBJInfo.ObjIndex).OBJType = eOBJType.otFogata Then
                        .OBJInfo.ObjIndex = 0
                        .OBJInfo.Amount = 0

                    End If

                End If
    
                If .TileExit.Map Then ByFlags = ByFlags Or 1
                If .NpcIndex Then ByFlags = ByFlags Or 2
                If .OBJInfo.ObjIndex Then ByFlags = ByFlags Or 4
                
                Call InfWriter.putByte(ByFlags)
                
                If .TileExit.Map Then
                    Call InfWriter.putInteger(.TileExit.Map)
                    Call InfWriter.putInteger(.TileExit.X)
                    Call InfWriter.putInteger(.TileExit.Y)

                End If
                
                If .NpcIndex Then Call InfWriter.putInteger(Npclist(.NpcIndex).Numero)
                
                If .OBJInfo.ObjIndex Then
                    Call InfWriter.putInteger(.OBJInfo.ObjIndex)
                    Call InfWriter.putInteger(.OBJInfo.Amount)

                End If
                
            End With

        Next X
    Next Y
    
    Call MapWriter.saveBuffer
    Call InfWriter.saveBuffer
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf
    
    Set MapWriter = Nothing
    Set InfWriter = Nothing

    With MapInfo(Map)
        'write .dat file
        Call IniManager.ChangeValue("Mapa" & Map, "Name", .Name)
        Call IniManager.ChangeValue("Mapa" & Map, "MusicNum", .Music)
        Call IniManager.ChangeValue("Mapa" & Map, "MagiaSinefecto", .MagiaSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "OcultarSinEfecto", .OcultarSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "StartPos", .StartPos.Map & "-" & .StartPos.X & "-" & .StartPos.Y)
    
        Call IniManager.ChangeValue("Mapa" & Map, "Terreno", .Terreno)
        Call IniManager.ChangeValue("Mapa" & Map, "Zona", .Zona)
        Call IniManager.ChangeValue("Mapa" & Map, "Restringir", .Restringir)
        Call IniManager.ChangeValue("Mapa" & Map, "BackUp", CStr(.BackUp))
    
        If .Pk Then
            Call IniManager.ChangeValue("Mapa" & Map, "Pk", "0")
        Else
            Call IniManager.ChangeValue("Mapa" & Map, "Pk", "1")

        End If
    
        Call IniManager.DumpFile(MAPFILE & ".dat")

    End With
    
    Set IniManager = Nothing

End Sub

Sub LoadArmasHerreria()

    Dim n As Integer, lc As Integer

    n = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

    ReDim Preserve ArmasHerrero(1 To n) As Integer

    For lc = 1 To n
        ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
    Next lc

End Sub

Sub LoadArmadurasHerreria()

    Dim n As Integer, lc As Integer

    n = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

    ReDim Preserve ArmadurasHerrero(1 To n) As Integer

    For lc = 1 To n
        ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
    Next lc

End Sub

Sub LoadObjCarpintero()

    Dim n As Integer, lc As Integer

    n = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))

    ReDim Preserve ObjCarpintero(1 To n) As Integer

    For lc = 1 To n
        ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
    Next lc

End Sub

Sub LoadOBJData()

    '###################################################
    '#               ATENCION PELIGRO                  #
    '###################################################
    '
    '¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
    '
    'El que ose desafiar esta LEY, se las tendrá que ver
    'con migo. Para leer desde el OBJ.DAT se deberá usar
    'la nueva clase clsLeerInis.
    '
    'Alejo
    '
    '###################################################

    'Call LogTarea("Sub LoadOBJData")

    On Error GoTo errhandler

    If frmMain.Visible Then frmMain.txStatus.caption = "Cargando base de datos de los objetos."

    '*****************************************************************
    'Carga la lista de objetos
    '*****************************************************************
    Dim Object As Integer
    Dim Leer   As New clsIniManager

    Call Leer.Initialize(DatPath & "Obj.dat")

    'obtiene el numero de obj
    NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))

    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumObjDatas
    frmCargando.cargar.value = 0

    ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
    'Llena la lista
    For Object = 1 To NumObjDatas

        If Object = 246 Then
            Object = Object

        End If
    
        ObjData(Object).Name = Leer.GetValue("OBJ" & Object, "Name")
    
        ObjData(Object).GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
    
        If ObjData(Object).GrhIndex <= 0 Then
            ObjData(Object).GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex "))

        End If
    
        ObjData(Object).OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
        ObjData(Object).Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
    
        Select Case ObjData(Object).OBJType

            Case eOBJType.otArmadura
                ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                ObjData(Object).Nemes = val(Leer.GetValue("OBJ" & Object, "Nemes"))
                ObjData(Object).Templ = val(Leer.GetValue("OBJ" & Object, "Templ"))
                ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                ObjData(Object).ObjetoEspecial = val(Leer.GetValue("OBJ" & Object, "ObjetoEspecial"))
        
            Case eOBJType.otESCUDO
                ObjData(Object).ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                ObjData(Object).Nemes = val(Leer.GetValue("OBJ" & Object, "Nemes"))
                ObjData(Object).Templ = val(Leer.GetValue("OBJ" & Object, "Templ"))
                ObjData(Object).ObjetoEspecial = val(Leer.GetValue("OBJ" & Object, "ObjetoEspecial"))
        
            Case eOBJType.otCASCO
                ObjData(Object).CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                ObjData(Object).Nemes = val(Leer.GetValue("OBJ" & Object, "Nemes"))
                ObjData(Object).Templ = val(Leer.GetValue("OBJ" & Object, "Templ"))
                ObjData(Object).ObjetoEspecial = val(Leer.GetValue("OBJ" & Object, "ObjetoEspecial"))
        
            Case eOBJType.otWeapon
                ObjData(Object).WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                ObjData(Object).Apuñala = val(Leer.GetValue("OBJ" & Object, "Apuñala"))
                ObjData(Object).Pegadoble = val(Leer.GetValue("OBJ" & Object, "PegaDoble"))
                ObjData(Object).DosManos = val(Leer.GetValue("OBJ" & Object, "DosManos"))
                ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                ObjData(Object).MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                ObjData(Object).MinHit = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                ObjData(Object).Proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
                ObjData(Object).Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
                ObjData(Object).StaffPower = val(Leer.GetValue("OBJ" & Object, "StaffPower"))
                ObjData(Object).StaffDamageBonus = val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
                ObjData(Object).VaraDragon = val(Leer.GetValue("OBJ" & Object, "VaraDragon"))
          
                ObjData(Object).Refuerzo = val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
            
                ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                ObjData(Object).Nemes = val(Leer.GetValue("OBJ" & Object, "Nemes"))
                ObjData(Object).Templ = val(Leer.GetValue("OBJ" & Object, "Templ"))
                ObjData(Object).ObjetoEspecial = val(Leer.GetValue("OBJ" & Object, "ObjetoEspecial"))
        
            Case eOBJType.otHerramientas
                ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
        
            Case eOBJType.otInstrumentos
                ObjData(Object).Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
                ObjData(Object).Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
                ObjData(Object).Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
        
            Case eOBJType.otMinerales
                ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
        
            Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
                ObjData(Object).IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
                ObjData(Object).IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
                ObjData(Object).IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
            
            Case eOBJType.otAmuleto
                ObjData(Object).Cae = val(Leer.GetValue("OBJ" & Object, "Cae"))
            
            Case eOBJType.otPociones
                ObjData(Object).TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
                ObjData(Object).MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
                ObjData(Object).MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
                ObjData(Object).DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
                ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                ObjData(Object).Cae = val(Leer.GetValue("OBJ" & Object, "Cae"))
            
            Case eOBJType.otVales
                ObjData(Object).Expe = val(Leer.GetValue("OBJ" & Object, "Expe"))
                ObjData(Object).Cae = val(Leer.GetValue("OBJ" & Object, "Cae"))

            Case eOBJType.otBarcos
                ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                ObjData(Object).MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                ObjData(Object).MinHit = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                ObjData(Object).Nemes = val(Leer.GetValue("OBJ" & Object, "Nemes"))
                ObjData(Object).Templ = val(Leer.GetValue("OBJ" & Object, "Templ"))
        
            Case eOBJType.otFlechas
                ObjData(Object).MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                ObjData(Object).MinHit = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                ObjData(Object).Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
                ObjData(Object).ObjetoEspecial = val(Leer.GetValue("OBJ" & Object, "ObjetoEspecial"))
            
            Case eOBJType.otPasaje
                ObjData(Object).Zona = val(Leer.GetValue("OBJ" & Object, "Zona"))

        End Select
    
        ObjData(Object).Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
        ObjData(Object).HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
    
        ObjData(Object).LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
    
        ObjData(Object).MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
    
        ObjData(Object).MaxHP = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
        ObjData(Object).MinHP = val(Leer.GetValue("OBJ" & Object, "MinHP"))
    
        ObjData(Object).Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
        ObjData(Object).Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
    
        ObjData(Object).MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
        ObjData(Object).MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
        ObjData(Object).Nivel = val(Leer.GetValue("OBJ" & Object, "Nivel"))
        
    
        ObjData(Object).MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
        ObjData(Object).MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
    
        ObjData(Object).RazaEnana = val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
        ObjData(Object).RazaHobbit = val(Leer.GetValue("OBJ" & Object, "RazaHobbit"))
        ObjData(Object).RazaVampiro = val(Leer.GetValue("OBJ" & Object, "RazaVampiro"))
        ObjData(Object).RazaOrco = val(Leer.GetValue("OBJ" & Object, "RazaOrco"))
    
        ObjData(Object).Valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
    
        ObjData(Object).Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
    
        ObjData(Object).Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta"))

        If ObjData(Object).Cerrada = 1 Then
            ObjData(Object).Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
            ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))

        End If
    
        'Puertas y llaves
        ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
    
        ObjData(Object).texto = Leer.GetValue("OBJ" & Object, "Texto")
        ObjData(Object).GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
    
        ObjData(Object).Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))
        ObjData(Object).ForoID = Leer.GetValue("OBJ" & Object, "ID")
    
        Dim i As Integer

        For i = 1 To 20
            ObjData(Object).ClaseProhibida(i) = StringToClase(Leer.GetValue("OBJ" & Object, "CP" & i))
        Next i
    
        ObjData(Object).DefensaMagicaMax = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
        ObjData(Object).DefensaMagicaMin = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
    
        ObjData(Object).SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
    
        If ObjData(Object).SkCarpinteria > 0 Then ObjData(Object).Madera = val(Leer.GetValue("OBJ" & Object, "Madera"))
    
        'Bebidas
        ObjData(Object).MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
    
        ObjData(Object).Cae = val(Leer.GetValue("OBJ" & Object, "Cae"))
    
        frmCargando.cargar.value = frmCargando.cargar.value + 1
    Next Object

    Set Leer = Nothing

    Exit Sub

errhandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.Description

End Sub

Sub LoadUserStats(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)

    On Error Resume Next

    Dim loopc As Integer

    For loopc = 1 To NUMATRIBUTOS
        UserList(UserIndex).Stats.UserAtributos(loopc) = CInt(UserFile.GetValue("ATRIBUTOS", "AT" & loopc))
        UserList(UserIndex).Stats.UserAtributosBackUP(loopc) = UserList(UserIndex).Stats.UserAtributos(loopc)
    Next loopc

    For loopc = 1 To NUMSKILLS
        UserList(UserIndex).Stats.UserSkills(loopc) = CInt(UserFile.GetValue("SKILLS", "SK" & loopc))
    Next loopc

    For loopc = 1 To MAXUSERHECHIZOS
        UserList(UserIndex).Stats.UserHechizos(loopc) = CInt(UserFile.GetValue("Hechizos", "H" & loopc))
    Next loopc

    UserList(UserIndex).Stats.GLD = CLng(UserFile.GetValue("STATS", "GLD"))
    UserList(UserIndex).Stats.Banco = CLng(UserFile.GetValue("STATS", "BANCO"))

    UserList(UserIndex).Stats.MET = CInt(UserFile.GetValue("STATS", "MET"))
    UserList(UserIndex).Stats.MaxHP = CInt(UserFile.GetValue("STATS", "MaxHP"))
    UserList(UserIndex).Stats.MinHP = CInt(UserFile.GetValue("STATS", "MinHP"))

    UserList(UserIndex).Stats.MinSta = CInt(UserFile.GetValue("STATS", "MinSTA"))
    UserList(UserIndex).Stats.MaxSta = CInt(UserFile.GetValue("STATS", "MaxSTA"))
    UserList(UserIndex).Stats.TrofOro = CInt(UserFile.GetValue("STATS", "TrofOro"))
    UserList(UserIndex).Stats.TrofBronce = CInt(UserFile.GetValue("STATS", "TrofBronce"))
    UserList(UserIndex).Stats.TrofPlata = CInt(UserFile.GetValue("STATS", "TrofPlata"))
    UserList(UserIndex).Stats.TrofMadera = CInt(UserFile.GetValue("STATS", "TrofMadera"))
    UserList(UserIndex).Stats.MaxMAN = CInt(UserFile.GetValue("STATS", "MaxMAN"))
    UserList(UserIndex).Stats.MinMAN = CInt(UserFile.GetValue("STATS", "MinMAN"))

    ' puntos
    UserList(UserIndex).Stats.PuntosDeath = CInt(UserFile.GetValue("STATS", "PuntosDeath"))
    UserList(UserIndex).Stats.PuntosDuelos = CInt(UserFile.GetValue("STATS", "PuntosDuelos"))
    UserList(UserIndex).Stats.PuntosTorneo = CInt(UserFile.GetValue("STATS", "PuntosTorneo"))
    UserList(UserIndex).Stats.PuntosRetos = CInt(UserFile.GetValue("STATS", "PuntosRetos"))
    UserList(UserIndex).Stats.PuntosPlante = CInt(UserFile.GetValue("STATS", "PuntosPlante"))

    UserList(UserIndex).Stats.MaxHit = CInt(UserFile.GetValue("STATS", "MaxHIT"))
    UserList(UserIndex).Stats.MinHit = CInt(UserFile.GetValue("STATS", "MinHIT"))

    UserList(UserIndex).Stats.MaxAGU = CInt(UserFile.GetValue("STATS", "MaxAGU"))
    UserList(UserIndex).Stats.MinAGU = CInt(UserFile.GetValue("STATS", "MinAGU"))

    UserList(UserIndex).Stats.MaxHam = CInt(UserFile.GetValue("STATS", "MaxHAM"))
    UserList(UserIndex).Stats.MinHam = CInt(UserFile.GetValue("STATS", "MinHAM"))

    UserList(UserIndex).Stats.SkillPts = CInt(UserFile.GetValue("STATS", "SkillPtsLibres"))

    UserList(UserIndex).Stats.Exp = CDbl(UserFile.GetValue("STATS", "EXP"))
    UserList(UserIndex).Stats.ELU = CLng(UserFile.GetValue("STATS", "ELU"))
    UserList(UserIndex).Stats.ELV = CLng(UserFile.GetValue("STATS", "ELV"))

    UserList(UserIndex).Stats.UsuariosMatados = CInt(UserFile.GetValue("MUERTES", "UserMuertes"))
    UserList(UserIndex).Stats.CriminalesMatados = CInt(UserFile.GetValue("MUERTES", "CrimMuertes"))
    UserList(UserIndex).Stats.NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))

    UserList(UserIndex).flags.PertAlCons = CByte(UserFile.GetValue("CONSEJO", "PERTENECE"))
    UserList(UserIndex).flags.PertAlConsCaos = CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS"))
    UserList(UserIndex).flags.Silenciado = CByte(UserFile.GetValue("FLAGS", "Silenciado"))
    
    UserList(UserIndex).Stats.CleroMatados = CInt(UserFile.GetValue("STATS", "CleroMatados"))
    UserList(UserIndex).Stats.AbbadonMatados = CInt(UserFile.GetValue("STATS", "AbbadonMatados"))
    UserList(UserIndex).Stats.TemplarioMatados = CInt(UserFile.GetValue("STATS", "TemplarioMatados"))
    UserList(UserIndex).Stats.TinieblaMatados = CInt(UserFile.GetValue("STATS", "TinieblaMatados"))

    Call CargarELU

End Sub

Sub LoadUserReputacion(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)

    UserList(UserIndex).Reputacion.AsesinoRep = CDbl(UserFile.GetValue("REP", "Asesino"))
    UserList(UserIndex).Reputacion.BandidoRep = CDbl(UserFile.GetValue("REP", "Bandido"))
    UserList(UserIndex).Reputacion.BurguesRep = CDbl(UserFile.GetValue("REP", "Burguesia"))
    UserList(UserIndex).Reputacion.LadronesRep = CDbl(UserFile.GetValue("REP", "Ladrones"))
    UserList(UserIndex).Reputacion.NobleRep = CDbl(UserFile.GetValue("REP", "Nobles"))
    UserList(UserIndex).Reputacion.PlebeRep = CDbl(UserFile.GetValue("REP", "Plebe"))
    UserList(UserIndex).Reputacion.Promedio = CDbl(UserFile.GetValue("REP", "Promedio"))

End Sub

Sub LoadUserInit(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)

    Dim loopc As Long
    Dim ln    As String
    Dim Obj   As Long
    
    UserList(UserIndex).AoMCreditos = val(UserFile.GetValue("PUNTOS", "AoMCreditos"))
    UserList(UserIndex).AoMCanjes = val(UserFile.GetValue("PUNTOS", "AoMCanjes"))

    UserList(UserIndex).Pareja = UserFile.GetValue("INIT", "Pareja")
    UserList(UserIndex).flags.Casado = CByte(UserFile.GetValue("FLAGS", "Casado"))

    UserList(UserIndex).Faccion.ArmadaReal = val(UserFile.GetValue("FACCIONES", "EJERCITOREAL"))
    UserList(UserIndex).Faccion.FuerzasCaos = CByte(UserFile.GetValue("FACCIONES", "EjercitoCaos"))

    UserList(UserIndex).Faccion.Templario = CByte(UserFile.GetValue("FACCIONES", "Templario"))
    UserList(UserIndex).Faccion.Nemesis = CByte(UserFile.GetValue("FACCIONES", "Nemesis"))

    UserList(UserIndex).Faccion.CiudadanosMatados = CDbl(UserFile.GetValue("FACCIONES", "CiudMatados"))
    UserList(UserIndex).Faccion.CriminalesMatados = CDbl(UserFile.GetValue("FACCIONES", "CrimMatados"))

    UserList(UserIndex).Faccion.RecibioArmaduraCaos = CByte(UserFile.GetValue("FACCIONES", "rArCaos"))
    UserList(UserIndex).Faccion.RecibioArmaduraReal = CByte(UserFile.GetValue("FACCIONES", "rArReal"))

    UserList(UserIndex).Faccion.RecibioArmaduraNemesis = CByte(UserFile.GetValue("FACCIONES", "rArNemesis"))
    UserList(UserIndex).Faccion.RecibioArmaduraTemplaria = CByte(UserFile.GetValue("FACCIONES", "rArTemplaria"))

    UserList(UserIndex).Faccion.RecibioExpInicialCaos = CByte(UserFile.GetValue("FACCIONES", "rExCaos"))
    UserList(UserIndex).Faccion.RecibioExpInicialReal = CByte(UserFile.GetValue("FACCIONES", "rExReal"))

    UserList(UserIndex).Faccion.RecibioExpInicialNemesis = CByte(UserFile.GetValue("FACCIONES", "rExNemesis"))
    UserList(UserIndex).Faccion.RecibioExpInicialTemplaria = CByte(UserFile.GetValue("FACCIONES", "rExTemplaria"))

    UserList(UserIndex).Faccion.RecompensasCaos = CByte(UserFile.GetValue("FACCIONES", "recCaos"))
    UserList(UserIndex).Faccion.RecompensasReal = CByte(UserFile.GetValue("FACCIONES", "recReal"))

    UserList(UserIndex).Faccion.RecompensasNemesis = CByte(UserFile.GetValue("FACCIONES", "recNemesis"))
    UserList(UserIndex).Faccion.RecompensasTemplaria = CByte(UserFile.GetValue("FACCIONES", "recTemplaria"))

    UserList(UserIndex).Faccion.Reenlistadas = CByte(UserFile.GetValue("FACCIONES", "Reenlistadas"))
    UserList(UserIndex).Faccion.ArmaduraFaccionaria = CInt(UserFile.GetValue("FACCIONES", "ArmaduraFaccionaria"))
    UserList(UserIndex).Faccion.NextRecompensas = CInt(UserFile.GetValue("FACCIONES", "NextRecompensas"))
    
    UserList(UserIndex).Faccion.FEnlistado = UserFile.GetValue("FACCIONES", "FechaIngreso")
    
    UserList(UserIndex).flags.Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
    UserList(UserIndex).flags.Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))
    
    UserList(UserIndex).Gladiador = CByte(UserFile.GetValue("INIT", "Gladiador")) 'gladiador magico

    UserList(UserIndex).flags.Hambre = CByte(UserFile.GetValue("FLAGS", "Hambre"))
    UserList(UserIndex).flags.Sed = CByte(UserFile.GetValue("FLAGS", "Sed"))
    UserList(UserIndex).flags.Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))

    UserList(UserIndex).flags.Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
    UserList(UserIndex).flags.Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))

    If UserList(UserIndex).flags.Paralizado = 1 Then
        UserList(UserIndex).Counters.Paralisis = IntervaloParalizado

    End If

    UserList(UserIndex).flags.Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
    UserList(UserIndex).flags.Embarcado = CByte(UserFile.GetValue("FLAGS", "Embarcado"))

    UserList(UserIndex).Counters.Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))

    UserList(UserIndex).Email = UserFile.GetValue("CONTACTO", "Email")
    
    UserList(UserIndex).Telepatia = UserFile.GetValue("INIT", "Telepatia")

    UserList(UserIndex).Zona = UserFile.GetValue("INIT", "Zona")
    UserList(UserIndex).Genero = UserFile.GetValue("INIT", "Genero")
    UserList(UserIndex).Clase = UserFile.GetValue("INIT", "Clase")
    UserList(UserIndex).Raza = UserFile.GetValue("INIT", "Raza")
    UserList(UserIndex).Hogar = UserFile.GetValue("INIT", "Hogar")
    UserList(UserIndex).char.Heading = CInt(UserFile.GetValue("INIT", "Heading"))

    UserList(UserIndex).OrigChar.Head = CInt(UserFile.GetValue("INIT", "Head"))
    UserList(UserIndex).OrigChar.Body = CInt(UserFile.GetValue("INIT", "Body"))
    UserList(UserIndex).OrigChar.WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
    UserList(UserIndex).OrigChar.ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
    UserList(UserIndex).OrigChar.CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))
    UserList(UserIndex).OrigChar.Alas = val(UserFile.GetValue("INIT", "Alas"))
    UserList(UserIndex).OrigChar.Heading = eHeading.SOUTH

    If UserList(UserIndex).flags.Muerto = 0 Then
        UserList(UserIndex).char = UserList(UserIndex).OrigChar
    Else
        UserList(UserIndex).char.Body = iCuerpoMuerto
        UserList(UserIndex).char.Head = iCabezaMuerto
        UserList(UserIndex).char.WeaponAnim = NingunArma
        UserList(UserIndex).char.ShieldAnim = NingunEscudo
        UserList(UserIndex).char.CascoAnim = NingunCasco

    End If

    UserList(UserIndex).Desc = UserFile.GetValue("INIT", "Desc")
    ' soporte
    UserList(UserIndex).Pregunta = UserFile.GetValue("INIT", "Pregunta")
    UserList(UserIndex).Respuesta = UserFile.GetValue("INIT", "Respuesta")

    UserList(UserIndex).pos.Map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
    UserList(UserIndex).pos.X = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
    UserList(UserIndex).pos.Y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))

    UserList(UserIndex).Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))

    '[KEVIN]--------------------------------------------------------------------
    '***********************************************************************************
    UserList(UserIndex).BancoInvent.NroItems = CInt(UserFile.GetValue("BancoInventory", "CantidadItems"))

    'Lista de objetos del banco
    For loopc = 1 To MAX_BANCOINVENTORY_SLOTS
        ln = UserFile.GetValue("BancoInventory", "Obj" & loopc)
        UserList(UserIndex).BancoInvent.Object(loopc).ObjIndex = CInt(ReadField(1, ln, 45))
        UserList(UserIndex).BancoInvent.Object(loopc).Amount = CInt(ReadField(2, ln, 45))
    Next loopc

    '------------------------------------------------------------------------------------
    '[/KEVIN]*****************************************************************************

    'Lista de objetos
    For loopc = 1 To MAX_INVENTORY_SLOTS
        ln = UserFile.GetValue("Inventory", "Obj" & loopc)
        UserList(UserIndex).Invent.Object(loopc).ObjIndex = CInt(ReadField(1, ln, 45))
        UserList(UserIndex).Invent.Object(loopc).Amount = CInt(ReadField(2, ln, 45))
        UserList(UserIndex).Invent.Object(loopc).Equipped = CByte(ReadField(3, ln, 45))
    Next loopc
    
    UserList(UserIndex).Stats.ELV = CLng(UserFile.GetValue("STATS", "ELV"))

    'Obtiene el indice-objeto del arma
    UserList(UserIndex).Invent.WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))

    If UserList(UserIndex).Invent.WeaponEqpSlot > 0 Then
        UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.WeaponEqpSlot).ObjIndex
        Obj = UserList(UserIndex).Invent.WeaponEqpObjIndex

        If val(ObjData(Obj).ObjetoEspecial) > 0 Then

            Call RevObjetoEspecial(UserIndex, ObjData(Obj).ObjetoEspecial)

        End If
           
        If EspadaSagrada.EspadaSagrada(UserList(UserIndex).Invent.WeaponEqpObjIndex) Then
            Call ChangeSagradaHit(UserIndex)
        End If

    End If

    'Obtiene el indice-objeto de ala
    UserList(UserIndex).Invent.AlaEqpSlot = val(UserFile.GetValue("Inventory", "AlaEqpSlot"))

    If UserList(UserIndex).Invent.AlaEqpSlot > 0 Then
        UserList(UserIndex).Invent.AlaEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.AlaEqpSlot).ObjIndex

    End If

    'Obtiene el indice-objeto del armadura
    UserList(UserIndex).Invent.ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))

    If UserList(UserIndex).Invent.ArmourEqpSlot > 0 Then
        UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.ArmourEqpSlot).ObjIndex
        UserList(UserIndex).flags.Desnudo = 0
        Obj = UserList(UserIndex).Invent.ArmourEqpObjIndex

        If val(ObjData(Obj).ObjetoEspecial) > 0 Then

            Call RevObjetoEspecial(UserIndex, ObjData(Obj).ObjetoEspecial)

        End If

    Else
        UserList(UserIndex).flags.Desnudo = 1

    End If

    'Obtiene el indice-objeto del escudo
    UserList(UserIndex).Invent.EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))

    If UserList(UserIndex).Invent.EscudoEqpSlot > 0 Then
        UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.EscudoEqpSlot).ObjIndex
        Obj = UserList(UserIndex).Invent.EscudoEqpObjIndex

        If val(ObjData(Obj).ObjetoEspecial) > 0 Then

            Call RevObjetoEspecial(UserIndex, ObjData(Obj).ObjetoEspecial)

        End If

    End If

    'Obtiene el indice-objeto del casco
    UserList(UserIndex).Invent.CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))

    If UserList(UserIndex).Invent.CascoEqpSlot > 0 Then
        UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.CascoEqpSlot).ObjIndex
        Obj = UserList(UserIndex).Invent.CascoEqpObjIndex

        If val(ObjData(Obj).ObjetoEspecial) > 0 Then

            Call RevObjetoEspecial(UserIndex, ObjData(Obj).ObjetoEspecial)

        End If

    End If

    'Obtiene el indice-objeto barco
    UserList(UserIndex).Invent.BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))

    If UserList(UserIndex).Invent.BarcoSlot > 0 Then
        UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.BarcoSlot).ObjIndex

    End If

    'Obtiene el indice-objeto municion
    UserList(UserIndex).Invent.MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))

    If UserList(UserIndex).Invent.MunicionEqpSlot > 0 Then
        UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MunicionEqpSlot).ObjIndex

    End If

    '[Alejo]
    'Obtiene el indice-objeto herramienta
    UserList(UserIndex).Invent.HerramientaEqpSlot = CInt(UserFile.GetValue("Inventory", "HerramientaSlot"))

    If UserList(UserIndex).Invent.HerramientaEqpSlot > 0 Then
        UserList(UserIndex).Invent.HerramientaEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.HerramientaEqpSlot).ObjIndex

    End If

    UserList(UserIndex).NroMacotas = 0

    ln = UserFile.GetValue("Guild", "GUILDINDEX")

    If IsNumeric(ln) Then
        UserList(UserIndex).GuildIndex = CInt(ln)
    Else
        UserList(UserIndex).GuildIndex = 0

    End If
    
    UserList(UserIndex).Clan.FundoClan = val(UserFile.GetValue("GUILD", "FundoClan"))
    UserList(UserIndex).Clan.PuntosClan = val(UserFile.GetValue("GUILD", "PuntosClan"))
    UserList(UserIndex).Clan.UMuerte = UserFile.GetValue("GUILD", "UltimaMuerte")
    UserList(UserIndex).Clan.ParticipoClan = val(UserFile.GetValue("GUILD", "PARTICIPOCLAN"))
    

End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String

    Dim sSpaces  As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found
  
    szReturn = ""
  
    sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
  
    GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
  
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()

    If frmMain.Visible Then frmMain.txStatus.caption = "Cargando backup."

    Dim Map       As Integer
    Dim TempInt   As Integer
    Dim tFileName As String
    Dim npcfile   As String

    On Error GoTo man
    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.value = 0
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
      
    For Map = 1 To NumMaps
        
        If val(GetVar(App.Path & MapPath & "Mapa" & Map & ".Dat", "Mapa" & Map, "BackUp")) <> 0 Then
            tFileName = App.Path & "\WorldBackUp\Mapa" & Map
            
            If Not FileExist(tFileName & ".*") Then 'Miramos que exista al menos uno de los 3 archivos, sino lo cargamos de la carpeta de los mapas
                tFileName = App.Path & MapPath & "Mapa" & Map

            End If

        Else
            tFileName = App.Path & MapPath & "Mapa" & Map

        End If
        
        Call CargarMapa(Map, tFileName)
        
        frmCargando.cargar.value = frmCargando.cargar.value + 1
        DoEvents
    Next Map

    Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)
 
End Sub

Sub LoadMapData()
    'Call LogTarea("Sub LoadMapData")

    If frmMain.Visible Then frmMain.txStatus.caption = "Cargando mapas."

    Dim Map       As Integer
    Dim tFileName As String

    On Error GoTo man

    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas

    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.value = 0

    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")

    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
  
    For Map = 1 To NumMaps
    
        tFileName = App.Path & MapPath & "Mapa" & Map
        Call CargarMapa(Map, tFileName)
        
        frmCargando.cargar.value = frmCargando.cargar.value + 1
        DoEvents
    Next Map

    Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)

End Sub

Public Sub CargarMapa(ByVal Map As Long, ByRef MAPFl As String)

    '***************************************************
    'Author: Unknown
    'Last Modification: 10/08/2010
    '10/08/2010 - Pato: Implemento el clsByteBuffer y el clsIniManager para la carga de mapa
    '***************************************************

    On Error GoTo errh

    Dim hFile     As Integer
    Dim X         As Long
    Dim Y         As Long
    Dim ByFlags   As Byte
    Dim npcfile   As String
    Dim Leer      As clsIniManager
    Dim MapReader As clsByteBuffer
    Dim InfReader As clsByteBuffer
    Dim Buff()    As Byte
    
    Set MapReader = New clsByteBuffer
    Set InfReader = New clsByteBuffer
    Set Leer = New clsIniManager
    
    npcfile = DatPath & "NPCs.dat"
    
    hFile = FreeFile

    Open MAPFl & ".map" For Binary As #hFile
    Seek hFile, 1

    ReDim Buff(LOF(hFile) - 1) As Byte
    
    Get #hFile, , Buff
    Close hFile
    
    Call MapReader.initializeReader(Buff)

    'inf
    Open MAPFl & ".inf" For Binary As #hFile
    Seek hFile, 1

    ReDim Buff(LOF(hFile) - 1) As Byte
    
    Get #hFile, , Buff
    Close hFile
    
    Call InfReader.initializeReader(Buff)
    
    'map Header
    MapInfo(Map).MapVersion = MapReader.getInteger
    
    MiCabecera.Desc = MapReader.getString(Len(MiCabecera.Desc))
    MiCabecera.crc = MapReader.getLong
    MiCabecera.MagicWord = MapReader.getLong
    
    Call MapReader.getDouble

    'inf Header
    Call InfReader.getDouble
    Call InfReader.getInteger

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            With MapData(Map, X, Y)
                '.map file
                ByFlags = MapReader.getByte

                If ByFlags And 1 Then .Blocked = 1

                .Graphic(1) = MapReader.getLong

                'Layer 2 used?
                If ByFlags And 2 Then .Graphic(2) = MapReader.getLong

                'Layer 3 used?
                If ByFlags And 4 Then .Graphic(3) = MapReader.getLong

                'Layer 4 used?
                If ByFlags And 8 Then .Graphic(4) = MapReader.getLong

                'Trigger used?
                If ByFlags And 16 Then .trigger = MapReader.getInteger

                '.inf file
                ByFlags = InfReader.getByte

                If ByFlags And 1 Then
                    .TileExit.Map = InfReader.getInteger
                    .TileExit.X = InfReader.getInteger
                    .TileExit.Y = InfReader.getInteger

                End If

                If ByFlags And 2 Then
                    'Get and make NPC
                    .NpcIndex = InfReader.getInteger

                    If .NpcIndex > 0 Then

                        'Si el npc debe hacer respawn en la pos
                        'original la guardamos
                        If val(Leer.GetValue("NPC" & .NpcIndex, "PosOrig")) = 1 Then
                            .NpcIndex = OpenNPC(.NpcIndex)
                            Npclist(.NpcIndex).Orig.Map = Map
                            Npclist(.NpcIndex).Orig.X = X
                            Npclist(.NpcIndex).Orig.Y = Y
                        Else
                            .NpcIndex = OpenNPC(.NpcIndex)

                        End If

                        Npclist(.NpcIndex).pos.Map = Map
                        Npclist(.NpcIndex).pos.X = X
                        Npclist(.NpcIndex).pos.Y = Y

                        'Call MakeNPCChar(True, 0, .NpcIndex, map, X, Y)
                        'Call MakeNPCChar(ToNone, 0, 0, MapData(map, X, Y).NpcIndex, map, X, Y)
                        Call MakeNPCChar(SendTarget.ToMap, 0, 0, .NpcIndex, 1, 1, 1)

                    End If

                End If

                If ByFlags And 4 Then
                    'Get and make Object
                    .OBJInfo.ObjIndex = InfReader.getInteger
                    .OBJInfo.Amount = InfReader.getInteger

                End If

            End With

        Next X
    Next Y
    
    Call Leer.Initialize(MAPFl & ".dat")
    
    With MapInfo(Map)
        .Name = Leer.GetValue("Mapa" & Map, "Name")
        .Music = Leer.GetValue("Mapa" & Map, "MusicNum")
        
        .StartPos.Map = val(ReadField(1, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
        .StartPos.X = val(ReadField(2, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
        .StartPos.Y = val(ReadField(3, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
        
        .MagiaSinEfecto = val(Leer.GetValue("Mapa" & Map, "MagiaSinEfecto"))
        .OcultarSinEfecto = val(Leer.GetValue("Mapa" & Map, "OcultarSinEfecto"))

        If val(Leer.GetValue("Mapa" & Map, "Pk")) = 0 Then
            .Pk = True
        Else
            .Pk = False

        End If
        
        .Terreno = Leer.GetValue("Mapa" & Map, "Terreno")
        .Zona = Leer.GetValue("Mapa" & Map, "Zona")
        .Restringir = Leer.GetValue("Mapa" & Map, "Restringir")
        .BackUp = val(Leer.GetValue("Mapa" & Map, "BACKUP"))

    End With
    
    Set MapReader = Nothing
    Set InfReader = Nothing
    Set Leer = Nothing
    
    Erase Buff
    Exit Sub

errh:
    Call LogError("Error cargando mapa: " & Map & " - Pos: " & X & "," & Y & "." & Err.Description)

    Set MapReader = Nothing
    Set InfReader = Nothing
    Set Leer = Nothing

End Sub

Sub SaveConfig()
    
    Call WriteVar(App.Path & "\Dat\Ini\Config.ini", "NOTOKAR", "NOPSD", MaxLevel)
    Call WriteVar(App.Path & "\Dat\Ini\Config.ini", "NOTOKAR", "USUARIO", UserMaxLevel)

End Sub

Sub LoadSini()

    Dim Temporal As Long

    If frmMain.Visible Then frmMain.txStatus.caption = "Cargando info de inicio del server."
    
    MaxLevel = val(GetVar(App.Path & "\Dat\Ini\Config.ini", "NOTOKAR", "NOPSD"))
    UserMaxLevel = GetVar(App.Path & "\Dat\Ini\Config.ini", "NOTOKAR", "USUARIO")
    
    Multexp = val(GetVar(App.Path & "\Dat\ini\multipli.ini", "Multiplicadores", "Exp"))
    MultOro = val(GetVar(App.Path & "\Dat\ini\multipli.ini", "Multiplicadores", "Oro"))
    MultMsg = GetVar(App.Path & "\Dat\ini\multipli.ini", "Multiplicadores", "Msg")
    
    BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))

    Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
    LastSockListen = val(GetVar(IniPath & "Server.ini", "INIT", "LastSockListen"))
    HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
    AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
    IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
    'Lee la version correcta del cliente
    ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")

    PuedeCrearPersonajes = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))
    CamaraLenta = val(GetVar(IniPath & "Server.ini", "INIT", "CamaraLenta"))
    ServerSoloGMs = val(GetVar(IniPath & "Server.ini", "init", "ServerSoloGMs"))

    MAPA_PRETORIANO = val(GetVar(IniPath & "Server.ini", "INIT", "MapaPretoriano"))

    ClientsCommandsQueue = val(GetVar(IniPath & "Server.ini", "INIT", "ClientsCommandsQueue"))
    EnTesting = val(GetVar(IniPath & "Server.ini", "INIT", "Testing"))
    EncriptarProtocolosCriticos = val(GetVar(IniPath & "Server.ini", "INIT", "Encriptar"))

    'Start pos
    StartPos.Map = val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
    StartPos.X = val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
    StartPos.Y = val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))

    'Intervalos
    SanaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
    FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar

    StaminaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
    FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar

    SanaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
    FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar

    StaminaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
    FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar

    IntervaloSed = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))
    FrmInterv.txtIntervaloSed.Text = IntervaloSed

    IntervaloHambre = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre"))
    FrmInterv.txtIntervaloHambre.Text = IntervaloHambre

    IntervaloVeneno = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
    FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno

    IntervaloParalizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
    FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado

    IntervaloInvisible = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
    FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible

    IntervaloFrio = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
    FrmInterv.txtIntervaloFrio.Text = IntervaloFrio

    IntervaloWavFx = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX"))
    FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx

    IntervaloInvocacion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
    FrmInterv.txtInvocacion.Text = IntervaloInvocacion

    IntervaloParaConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
    FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion

    '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&

    IntervaloUserPuedeCastear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo"))
    FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear

    frmMain.TIMER_AI.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI"))
    FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval

    frmMain.npcataca.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
    FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.Interval

    IntervaloUserPuedeTrabajar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
    FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar

    IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
    FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar

    MinutosWs = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))

    If MinutosWs < 60 Then MinutosWs = 180
    
    MinutosGuardarUsuarios = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGuardarUsuarios"))
    
    MinutosLimpia = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "MinutosLimpia"))
    
    IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))
    IntervaloUserPuedeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar"))
    IntervaloFlechasCazadores = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechasCazadores"))

    entrarReto = val(GetVar(IniPath & "Server.ini", "EVENTOS", "PRECIORETO"))
    entrarPlante = val(GetVar(IniPath & "Server.ini", "EVENTOS", "PRECIOPLANTE"))
    entrarReto2v2 = val(GetVar(IniPath & "Server.ini", "EVENTOS", "PRECIO2V2"))

    lvlGuerra = val(GetVar(IniPath & "Server.ini", "EVENTOS", "LVLGUERRA"))
    lvlMedusa = val(GetVar(IniPath & "Server.ini", "EVENTOS", "LVLMEDUSA"))
    lvlTorneo = val(GetVar(IniPath & "Server.ini", "EVENTOS", "LVLTORNEO"))
    lvlDeath = val(GetVar(IniPath & "Server.ini", "EVENTOS", "LVLDEATH"))

    'Ressurect pos
    ResPos.Map = val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
    ResPos.X = val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
    ResPos.Y = val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
  
    recordusuarios = val(GetVar(IniPath & "Server.ini", "INIT", "Record"))
  
    'Max users
    Temporal = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))

    If MaxUsers = 0 Then
        MaxUsers = Temporal
        ReDim UserList(1 To MaxUsers) As User

    End If

    Nix.Map = GetVar(DatPath & "Ciudades.dat", "NIX", "Mapa")
    Nix.X = GetVar(DatPath & "Ciudades.dat", "NIX", "X")
    Nix.Y = GetVar(DatPath & "Ciudades.dat", "NIX", "Y")

    Ullathorpe.Map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
    Ullathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
    Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")

    Banderbill.Map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
    Banderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
    Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")

    Lindos.Map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
    Lindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
    Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")
    
    'OPCIONES
    NumClan = val(GetVar(IniPath & "Server.ini", "OPCIONES", "Clan"))
    
    ' Cargar Administradores
    Call LoadAdministrativeUsers
    
    Call LoadNosfe
    Call LoadSagradas

End Sub

Sub LoadZonas()

    NumZonas = val(GetVar(DatPath & "Zonas.dat", "MAIN", "NumZonas")) - 1
    Dim i As Long

    For i = 0 To NumZonas
        Zonas(i).nombre = (GetVar(DatPath & "Zonas.dat", "Zona" & i + 1, "Nombre"))
        Zonas(i).Map = val(GetVar(DatPath & "Zonas.dat", "Zona" & i + 1, "Map"))
        Zonas(i).X = val(GetVar(DatPath & "Zonas.dat", "Zona" & i + 1, "X"))
        Zonas(i).Y = val(GetVar(DatPath & "Zonas.dat", "Zona" & i + 1, "Y"))
    Next i

End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
    '*****************************************************************
    'Escribe VAR en un archivo
    '*****************************************************************

    writeprivateprofilestring Main, Var, value, File
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, ByVal UserFile As String)

    On Error GoTo errhandler

    Dim Manager As clsIniManager
    Dim Existe  As Boolean
    Dim loopc   As Long

    With UserList(UserIndex)

        'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
        If Len(.Clase) = 0 Or .Stats.ELV = 0 Then
            Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(UserIndex).Name)
            Exit Sub

        End If
        
        Set Manager = New clsIniManager
    
        If FileExist(UserFile) Then
            Call Manager.Initialize(UserFile)
        
            If FileExist(UserFile & ".bk") Then Call Kill(UserFile & ".bk")
            Name UserFile As UserFile & ".bk"
        
            Existe = True

        End If

        If .flags.Mimetizado = 1 Then
            .char.Body = .CharMimetizado.Body
            .char.Head = .CharMimetizado.Head
            .char.CascoAnim = .CharMimetizado.CascoAnim
            .char.ShieldAnim = .CharMimetizado.ShieldAnim
            .char.WeaponAnim = .CharMimetizado.WeaponAnim

            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0

        End If
        
        
        Call Manager.ChangeValue("PUNTOS", "AoMCreditos", val(.AoMCreditos))
        Call Manager.ChangeValue("PUNTOS", "AoMCanjes", val(.AoMCanjes))
        
        Call Manager.ChangeValue("FLAGS", "Muerto", CStr(.flags.Muerto))
        Call Manager.ChangeValue("FLAGS", "Escondido", CStr(.flags.Escondido))
        Call Manager.ChangeValue("FLAGS", "Hambre", CStr(.flags.Hambre))
        Call Manager.ChangeValue("FLAGS", "Sed", CStr(.flags.Sed))
        Call Manager.ChangeValue("FLAGS", "Desnudo", CStr(.flags.Desnudo))
        Call Manager.ChangeValue("FLAGS", "Ban", CStr(.flags.Ban))
        Call Manager.ChangeValue("FLAGS", "Silenciado", CStr(.flags.Silenciado))
        Call Manager.ChangeValue("FLAGS", "Navegando", CStr(.flags.Navegando))
        Call Manager.ChangeValue("FLAGS", "Embarcado", CStr(.flags.Embarcado))
        
        Call Manager.ChangeValue("INIT", "Gladiador", .Gladiador)
        
        Call Manager.ChangeValue("FLAGS", "Envenenado", CStr(.flags.Envenenado))
        Call Manager.ChangeValue("FLAGS", "Paralizado", CStr(.flags.Paralizado))

        Call Manager.ChangeValue("INIT", "Pareja", .Pareja)
        Call Manager.ChangeValue("FLAGS", "Casado", CStr(.flags.Casado))

        Call Manager.ChangeValue("CONSEJO", "PERTENECE", CStr(.flags.PertAlCons))
        Call Manager.ChangeValue("CONSEJO", "PERTENECECAOS", CStr(.flags.PertAlConsCaos))

        Call Manager.ChangeValue("COUNTERS", "Pena", CStr(.Counters.Pena))

        Call Manager.ChangeValue("FACCIONES", "EjercitoReal", CStr(.Faccion.ArmadaReal))
        Call Manager.ChangeValue("FACCIONES", "EjercitoCaos", CStr(.Faccion.FuerzasCaos))
    
        Call Manager.ChangeValue("FACCIONES", "Nemesis", CStr(.Faccion.Nemesis))
        Call Manager.ChangeValue("FACCIONES", "Templario", CStr(.Faccion.Templario))
    
        Call Manager.ChangeValue("FACCIONES", "CiudMatados", CStr(.Faccion.CiudadanosMatados))
        Call Manager.ChangeValue("FACCIONES", "CrimMatados", CStr(.Faccion.CriminalesMatados))
    
        Call Manager.ChangeValue("FACCIONES", "rArCaos", CStr(.Faccion.RecibioArmaduraCaos))
        Call Manager.ChangeValue("FACCIONES", "rArReal", CStr(.Faccion.RecibioArmaduraReal))
    
        Call Manager.ChangeValue("FACCIONES", "rArNemesis", CStr(.Faccion.RecibioArmaduraNemesis))
        Call Manager.ChangeValue("FACCIONES", "rArTemplaria", CStr(.Faccion.RecibioArmaduraTemplaria))
        
        Call Manager.ChangeValue("FACCIONES", "rExCaos", CStr(.Faccion.RecibioExpInicialCaos))
        Call Manager.ChangeValue("FACCIONES", "rExReal", CStr(.Faccion.RecibioExpInicialReal))
    
        Call Manager.ChangeValue("FACCIONES", "rExNemesis", CStr(.Faccion.RecibioExpInicialNemesis))
        Call Manager.ChangeValue("FACCIONES", "rExTemplaria", CStr(.Faccion.RecibioExpInicialTemplaria))
    
        Call Manager.ChangeValue("FACCIONES", "recCaos", CStr(.Faccion.RecompensasCaos))
        Call Manager.ChangeValue("FACCIONES", "recReal", CStr(.Faccion.RecompensasReal))
    
        Call Manager.ChangeValue("FACCIONES", "recNemesis", CStr(.Faccion.RecompensasNemesis))
        Call Manager.ChangeValue("FACCIONES", "recTemplaria", CStr(.Faccion.RecompensasTemplaria))
    
        Call Manager.ChangeValue("FACCIONES", "Reenlistadas", CStr(.Faccion.Reenlistadas))
        Call Manager.ChangeValue("FACCIONES", "ArmaduraFaccionaria", CStr(.Faccion.ArmaduraFaccionaria))
        Call Manager.ChangeValue("FACCIONES", "NextRecompensas", CStr(.Faccion.NextRecompensas))
        
        Call Manager.ChangeValue("FACCIONES", "FechaIngreso", .Faccion.FEnlistado)
        
        '¿Fueron modificados los atributos del usuario?
        If Not .flags.TomoPocionAmarilla And Not .flags.TomoPocionVerde Then

            For loopc = 1 To UBound(.Stats.UserAtributos)
                Call Manager.ChangeValue("ATRIBUTOS", "AT" & loopc, CStr(.Stats.UserAtributos(loopc)))
            Next loopc

        Else

            For loopc = 1 To UBound(.Stats.UserAtributos)
                '.Stats.UserAtributos(LoopC) = .Stats.UserAtributosBackUP(LoopC)
                Call Manager.ChangeValue("ATRIBUTOS", "AT" & loopc, CStr(.Stats.UserAtributosBackUP(loopc)))
            Next

        End If

        For loopc = 1 To UBound(.Stats.UserSkills)
            Call Manager.ChangeValue("SKILLS", "SK" & loopc, CStr(.Stats.UserSkills(loopc)))
        Next loopc

        Call Manager.ChangeValue("CONTACTO", "Email", .Email)

        Call Manager.ChangeValue("INIT", "Genero", .Genero)
        Call Manager.ChangeValue("INIT", "Raza", .Raza)
        Call Manager.ChangeValue("INIT", "Hogar", .Hogar)
        Call Manager.ChangeValue("INIT", "Clase", .Clase)
        Call Manager.ChangeValue("INIT", "Password", .Password)
        Call Manager.ChangeValue("INIT", "Desc", .Desc)
        Call Manager.ChangeValue("INIT", "Zona", CStr(.Zona))
        Call Manager.ChangeValue("INIT", "Telepatia", .Telepatia)

        Call Manager.ChangeValue("INIT", "Heading", CStr(.char.Heading))
        Call Manager.ChangeValue("INIT", "Head", CStr(.OrigChar.Head))

        If .flags.Muerto = 0 Then
            Call Manager.ChangeValue("INIT", "Body", CStr(.char.Body))

        End If
    
        Call Manager.ChangeValue("INIT", "Alas", CStr(.char.Alas))
        Call Manager.ChangeValue("INIT", "Arma", CStr(.char.WeaponAnim))
        Call Manager.ChangeValue("INIT", "Escudo", CStr(.char.ShieldAnim))
        Call Manager.ChangeValue("INIT", "Casco", CStr(.char.CascoAnim))

        Call Manager.ChangeValue("INIT", "LastIP", .ip)
        Call Manager.ChangeValue("INIT", "Position", .pos.Map & "-" & .pos.X & "-" & .pos.Y)
        'soporte
        Call Manager.ChangeValue("INIT", "Pregunta", .Pregunta)
        Call Manager.ChangeValue("INIT", "Respuesta", .Respuesta)

        Call Manager.ChangeValue("STATS", "GLD", CStr(.Stats.GLD))
        Call Manager.ChangeValue("STATS", "BANCO", CStr(.Stats.Banco))

        Call Manager.ChangeValue("STATS", "MET", CStr(.Stats.MET))
        Call Manager.ChangeValue("STATS", "MaxHP", CStr(.Stats.MaxHP))
        Call Manager.ChangeValue("STATS", "MinHP", CStr(.Stats.MinHP))

        Call Manager.ChangeValue("STATS", "MaxSTA", CStr(.Stats.MaxSta))
        Call Manager.ChangeValue("STATS", "MinSTA", CStr(.Stats.MinSta))

        Call Manager.ChangeValue("STATS", "MaxMAN", CStr(.Stats.MaxMAN))
        Call Manager.ChangeValue("STATS", "MinMAN", CStr(.Stats.MinMAN))

        Call Manager.ChangeValue("STATS", "TrofOro", CStr(.Stats.TrofOro))
        Call Manager.ChangeValue("STATS", "TrofPlata", CStr(.Stats.TrofPlata))
        Call Manager.ChangeValue("STATS", "TrofBronce", CStr(.Stats.TrofBronce))
        Call Manager.ChangeValue("STATS", "TrofMadera", CStr(.Stats.TrofMadera))

        'puntos
        Call Manager.ChangeValue("STATS", "PuntosDeath", CStr(.Stats.PuntosDeath))
        Call Manager.ChangeValue("STATS", "PuntosDuelos", CStr(.Stats.PuntosDuelos))
        Call Manager.ChangeValue("STATS", "PuntosTorneo", CStr(.Stats.PuntosTorneo))
        Call Manager.ChangeValue("STATS", "PuntosRetos", CStr(.Stats.PuntosRetos))
        Call Manager.ChangeValue("STATS", "PuntosPlante", CStr(.Stats.PuntosPlante))

        Call Manager.ChangeValue("STATS", "MaxHIT", CStr(.Stats.MaxHit))
        Call Manager.ChangeValue("STATS", "MinHIT", CStr(.Stats.MinHit))
        Call Manager.ChangeValue("STATS", "MaxAGU", CStr(.Stats.MaxAGU))
        Call Manager.ChangeValue("STATS", "MinAGU", CStr(.Stats.MinAGU))

        Call Manager.ChangeValue("STATS", "MaxHAM", CStr(.Stats.MaxHam))
        Call Manager.ChangeValue("STATS", "MinHAM", CStr(.Stats.MinHam))

        Call Manager.ChangeValue("STATS", "SkillPtsLibres", CStr(.Stats.SkillPts))
  
        Call Manager.ChangeValue("STATS", "EXP", CStr(.Stats.Exp))
        Call Manager.ChangeValue("STATS", "ELV", CStr(.Stats.ELV))

        Call Manager.ChangeValue("STATS", "ELU", CStr(.Stats.ELU))
        
        Call Manager.ChangeValue("STATS", "CLEROMATADOS", CStr(.Stats.CleroMatados))
        Call Manager.ChangeValue("STATS", "ABBADONMATADOS", CStr(.Stats.AbbadonMatados))
        Call Manager.ChangeValue("STATS", "TEMPLARIOMATADOS", CStr(.Stats.TemplarioMatados))
        Call Manager.ChangeValue("STATS", "TINIEBLAMATADOS", CStr(.Stats.TinieblaMatados))
        
        Call Manager.ChangeValue("MUERTES", "UserMuertes", CStr(.Stats.UsuariosMatados))
        Call Manager.ChangeValue("MUERTES", "CrimMuertes", CStr(.Stats.CriminalesMatados))
        Call Manager.ChangeValue("MUERTES", "NpcsMuertes", CStr(.Stats.NPCsMuertos))
  
        '[KEVIN]----------------------------------------------------------------------------
        '*******************************************************************************************
        Call Manager.ChangeValue("BancoInventory", "CantidadItems", val(.BancoInvent.NroItems))

        For loopc = 1 To MAX_BANCOINVENTORY_SLOTS
            Call Manager.ChangeValue("BancoInventory", "Obj" & loopc, .BancoInvent.Object(loopc).ObjIndex & "-" & .BancoInvent.Object(loopc).Amount)
        Next loopc

        '*******************************************************************************************
        '[/KEVIN]-----------
  
        'Save Inv
        Call Manager.ChangeValue("Inventory", "CantidadItems", val(.Invent.NroItems))

        For loopc = 1 To MAX_INVENTORY_SLOTS
            Call Manager.ChangeValue("Inventory", "Obj" & loopc, .Invent.Object(loopc).ObjIndex & "-" & .Invent.Object(loopc).Amount & "-" & _
                .Invent.Object(loopc).Equipped)
        Next

        Call Manager.ChangeValue("Inventory", "WeaponEqpSlot", CStr(.Invent.WeaponEqpSlot))
        Call Manager.ChangeValue("Inventory", "AlaEqpSlot", CStr(.Invent.AlaEqpSlot))
        Call Manager.ChangeValue("Inventory", "ArmourEqpSlot", CStr(.Invent.ArmourEqpSlot))
        Call Manager.ChangeValue("Inventory", "CascoEqpSlot", CStr(.Invent.CascoEqpSlot))
        Call Manager.ChangeValue("Inventory", "EscudoEqpSlot", CStr(.Invent.EscudoEqpSlot))
        Call Manager.ChangeValue("Inventory", "BarcoSlot", CStr(.Invent.BarcoSlot))
        Call Manager.ChangeValue("Inventory", "MunicionSlot", CStr(.Invent.MunicionEqpSlot))
        Call Manager.ChangeValue("Inventory", "HerramientaSlot", CStr(.Invent.HerramientaEqpSlot))

        'Reputacion
        Call Manager.ChangeValue("REP", "Asesino", CStr(.Reputacion.AsesinoRep))
        Call Manager.ChangeValue("REP", "Bandido", CStr(.Reputacion.BandidoRep))
        Call Manager.ChangeValue("REP", "Burguesia", CStr(.Reputacion.BurguesRep))
        Call Manager.ChangeValue("REP", "Ladrones", CStr(.Reputacion.LadronesRep))
        Call Manager.ChangeValue("REP", "Nobles", CStr(.Reputacion.NobleRep))
        Call Manager.ChangeValue("REP", "Plebe", CStr(.Reputacion.PlebeRep))

        Dim L As Long
        L = (-.Reputacion.AsesinoRep) + (-.Reputacion.BandidoRep) + UserList(UserIndex).Reputacion.BurguesRep + (-.Reputacion.LadronesRep) + _
            UserList(UserIndex).Reputacion.NobleRep + .Reputacion.PlebeRep
        L = L / 6
        Call Manager.ChangeValue("REP", "Promedio", CStr(L))

        Dim cad As String

        For loopc = 1 To MAXUSERHECHIZOS
            cad = .Stats.UserHechizos(loopc)
            Call Manager.ChangeValue("HECHIZOS", "H" & loopc, cad)
        Next

        Dim NroMascotas As Long
        NroMascotas = .NroMacotas

        For loopc = 1 To MAXMASCOTAS

            ' Mascota valida?
            If .MascotasIndex(loopc) > 0 Then

                ' Nos aseguramos que la criatura no fue invocada
                If Npclist(.MascotasIndex(loopc)).Contadores.TiempoExistencia = 0 Then
                    cad = .MascotasType(loopc)
                Else 'Si fue invocada no la guardamos
                    cad = "0"
                    NroMascotas = NroMascotas - 1

                End If

                Call Manager.ChangeValue("MASCOTAS", "MAS" & loopc, cad)

            End If

        Next

        Call Manager.ChangeValue("MASCOTAS", "NroMascotas", CStr(NroMascotas))

        'Devuelve el head de muerto
        If .flags.Muerto = 1 Then
            .char.Head = iCabezaMuerto
        End If
        
        Call Manager.ChangeValue("GUILD", "FundoClan", .Clan.FundoClan)
        Call Manager.ChangeValue("GUILD", "PuntosClan", .Clan.PuntosClan)
        Call Manager.ChangeValue("GUILD", "UltimaMuerte", .Clan.UMuerte)
        Call Manager.ChangeValue("GUILD", "PARTICIPOCLAN", .Clan.ParticipoClan)

    End With

    Call Manager.DumpFile(UserFile)

    Set Manager = Nothing

    If Existe Then Call Kill(UserFile & ".bk")
    
    Call SaveConfig
    
    Call Save_Rank(UserIndex)
    
#If MYSQL = 1 Then
    If UserList(UserIndex).flags.Privilegios = PlayerType.User Then
        Call Add_DataBase(UserIndex, "SaveUser")
        DoEvents
        Call Add_DataBase(UserIndex, "Account")
    End If
#End If
    
    Exit Sub

errhandler:
    Call LogError("Error en SaveUser")
    Set Manager = Nothing

End Sub

Function Criminal(ByVal UserIndex As Integer) As Boolean
    
    Dim L As Long
    
    With UserList(UserIndex).Reputacion
        L = (-.AsesinoRep) + (-.BandidoRep) + .BurguesRep + (-.LadronesRep) + .NobleRep + .PlebeRep
        L = L / 6
        Criminal = (L < 0)

    End With

End Function

Sub BackUPnPc(ByVal NpcIndex As Integer, ByVal hFile As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 10/09/2010
    '10/09/2010 - Pato: Optimice el BackUp de NPCs
    '***************************************************

    Dim loopc As Long
    
    Print #hFile, "[NPC" & Npclist(NpcIndex).Numero & "]"
    
    With Npclist(NpcIndex)
    
        'General
        Print #hFile, "Name=" & .Name
        Print #hFile, "Desc=" & .Desc
        Print #hFile, "Head=" & val(.char.Head)
        Print #hFile, "Body=" & val(.char.Body)
        Print #hFile, "Heading=" & val(.char.Heading)
        Print #hFile, "Movement=" & val(.Movement)
        Print #hFile, "Attackable=" & val(.Attackable)
        Print #hFile, "Comercia=" & val(.Comercia)
        Print #hFile, "TipoItems=" & val(.TipoItems)
        Print #hFile, "Hostil=" & val(.Hostile)
        Print #hFile, "GiveEXP=" & val(.GiveEXP)
        Print #hFile, "GiveGLD=" & val(.GiveGLD)
        Print #hFile, "InvReSpawn=" & val(.InvReSpawn)
        Print #hFile, "NpcType=" & val(.NPCtype)
        
        'Stats
        Print #hFile, "Alineacion=" & val(.Stats.Alineacion)
        Print #hFile, "DEF=" & val(.Stats.def)
        Print #hFile, "MaxHit=" & val(.Stats.MaxHit)
        Print #hFile, "MaxHp=" & val(.Stats.MaxHP)
        Print #hFile, "MinHit=" & val(.Stats.MinHit)
        Print #hFile, "MinHp=" & val(.Stats.MinHP)
        
        'Flags
        Print #hFile, "ReSpawn=" & val(.flags.Respawn)
        Print #hFile, "BackUp=" & val(.flags.BackUp)
        Print #hFile, "Domable=" & val(.flags.Domable)
        
        'Inventario
        Print #hFile, "NroItems=" & val(.Invent.NroItems)

        If .Invent.NroItems > 0 Then

            For loopc = 1 To .Invent.NroItems
                Print #hFile, "Obj" & loopc & "=" & .Invent.Object(loopc).ObjIndex & "-" & .Invent.Object(loopc).Amount
            Next loopc

        End If
        
        Print #hFile, vbNullString

    End With

End Sub

Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NPCNumber As Integer)

    'Status
    If frmMain.Visible Then frmMain.txStatus.caption = "Cargando backup Npc"

    Dim npcfile As String
    Dim Leer    As clsIniManager
    
    Set Leer = New clsIniManager

    Leer.Initialize DatPath & "bkNPCs.dat"

    With Npclist(NpcIndex)
        .Numero = NPCNumber
        .Name = Leer.GetValue("NPC" & NPCNumber, "Name")
        .Desc = Leer.GetValue("NPC" & NPCNumber, "Desc")
        .Movement = val(Leer.GetValue("NPC" & NPCNumber, "Movement"))
        .NPCtype = val(Leer.GetValue("NPC" & NPCNumber, "NpcType"))
        .DefensaMagica = val(Leer.GetValue("NPC" & NPCNumber, "DefensaMagica"))

        .char.Body = val(Leer.GetValue("NPC" & NPCNumber, "Body"))
        .char.Head = val(Leer.GetValue("NPC" & NPCNumber, "Head"))
        .char.Heading = val(Leer.GetValue("NPC" & NPCNumber, "Heading"))

        .Attackable = val(Leer.GetValue("NPC" & NPCNumber, "Attackable"))
        .Comercia = val(Leer.GetValue("NPC" & NPCNumber, "Comercia"))
        .Hostile = val(Leer.GetValue("NPC" & NPCNumber, "Hostile"))
        
        If DiaEspecialExp = True Then
            .GiveEXP = Round((val(Leer.GetValue("NPC" & NPCNumber, "GiveEXP")) * Multexp) * LoteriaCriatura)
        Else
            
            If SistemaCriatura.ExpCriatura = True Then
                If Npclist(NpcIndex).Numero = NpcCriatura Then
                    .GiveEXP = Round((val(Leer.GetValue("NPC" & NPCNumber, "GiveEXP")) * Multexp) * LoteriaCriatura)
                Else
                    .GiveEXP = val(Leer.GetValue("NPC" & NPCNumber, "GiveEXP")) * Multexp

                End If

            Else
                .GiveEXP = val(Leer.GetValue("NPC" & NPCNumber, "GiveEXP")) * Multexp

            End If

        End If
    
        .flags.ExpCount = .GiveEXP

        If DiaEspecialOro = True Then
            Npclist(NpcIndex).GiveGLD = Round(val(Leer.GetValue("NPC" & NPCNumber, "GiveGLD")) * MultOro) * LoteriaCriatura
        Else

            If SistemaCriatura.OroCriatura = True Then
                If Npclist(NpcIndex).Numero = NpcCriatura Then
                    .GiveGLD = Round(val(Leer.GetValue("NPC" & NPCNumber, "GiveGLD")) * MultOro) * LoteriaCriatura
                Else
                    .GiveGLD = val(Leer.GetValue("NPC" & NPCNumber, "GiveGLD")) * MultOro

                End If

            Else
                .GiveGLD = val(Leer.GetValue("NPC" & NPCNumber, "GiveGLD")) * MultOro

            End If

        End If

        .InvReSpawn = val(Leer.GetValue("NPC" & NPCNumber, "InvReSpawn"))

        .Stats.MaxHP = val(Leer.GetValue("NPC" & NPCNumber, "MaxHP"))
        .Stats.MinHP = val(Leer.GetValue("NPC" & NPCNumber, "MinHP"))
        .Stats.MaxHit = val(Leer.GetValue("NPC" & NPCNumber, "MaxHIT"))
        .Stats.MinHit = val(Leer.GetValue("NPC" & NPCNumber, "MinHIT"))
        .Stats.def = val(Leer.GetValue("NPC" & NPCNumber, "DEF"))
        .Stats.Alineacion = val(Leer.GetValue("NPC" & NPCNumber, "Alineacion"))

        Dim loopc As Long
        Dim ln    As String
        .Invent.NroItems = val(Leer.GetValue("NPC" & NPCNumber, "NROITEMS"))

        If .Invent.NroItems > 0 Then

            For loopc = 1 To MAX_INVENTORY_SLOTS
                ln = Leer.GetValue("NPC" & NPCNumber, "Obj" & loopc)
                .Invent.Object(loopc).ObjIndex = val(ReadField(1, ln, 45))
                .Invent.Object(loopc).Amount = val(ReadField(2, ln, 45))
       
            Next loopc

        Else

            For loopc = 1 To MAX_INVENTORY_SLOTS
                .Invent.Object(loopc).ObjIndex = 0
                .Invent.Object(loopc).Amount = 0
            Next loopc

        End If

        .Inflacion = val(Leer.GetValue("NPC" & NPCNumber, "Inflacion"))

        .flags.NPCActive = True
        .flags.Respawn = val(Leer.GetValue("NPC" & NPCNumber, "ReSpawn"))
        .flags.BackUp = val(Leer.GetValue("NPC" & NPCNumber, "BackUp"))
        .flags.Domable = val(Leer.GetValue("NPC" & NPCNumber, "Domable"))
        .flags.RespawnOrigPos = val(Leer.GetValue("NPC" & NPCNumber, "OrigPos"))

        'Tipo de items con los que comercia
        .TipoItems = val(Leer.GetValue("NPC" & NPCNumber, "TipoItems"))

    End With
    
    Set Leer = Nothing

End Sub

Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal motivo As String)

    Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", motivo)

    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, UserList(BannedIndex).Name
    Close #mifile

End Sub

Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal motivo As String)

    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)

    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub

Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal motivo As String)

    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)

    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub

Public Sub CargaApuestas()

    Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

End Sub
