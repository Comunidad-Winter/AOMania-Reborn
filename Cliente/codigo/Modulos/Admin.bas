Attribute VB_Name = "Admin"
Option Explicit

Public Type tMotd

    texto As String
    Formato As String

End Type

Public MaxLines As Integer
Public MOTD()   As tMotd

Public Type tAPuestas

    Ganancias As Long
    Perdidas As Long
    Jugadas As Long

End Type

Public Apuestas                     As tAPuestas

Public NPCs                         As Long
Public DebugSocket                  As Boolean

Public Horas                        As Long
Public Dias                         As Long
Public MinsRunning                  As Long

Public ReiniciarServer              As Long

Public tInicioServer                As Long

Public SanaIntervaloSinDescansar    As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar       As Integer
Public StaminaIntervaloDescansar    As Integer
Public IntervaloSed                 As Integer
Public IntervaloHambre              As Integer
Public IntervaloVeneno              As Integer
Public IntervaloParalizado          As Integer
Public IntervaloInvisible           As Integer
Public IntervaloFrio                As Integer
Public IntervaloWavFx               As Integer
Public IntervaloLanzaHechizo        As Integer
Public IntervaloNPCPuedeAtacar      As Integer
Public IntervaloNPCAI               As Integer
Public IntervaloInvocacion          As Integer
Public IntervaloUserPuedeAtacar     As Long
Public IntervaloUserPuedeCastear    As Long
Public IntervaloUserPuedeTrabajar   As Long
Public IntervaloParaConexion        As Long
Public IntervaloCerrarConexion      As Long '[Gonzalo]
Public IntervaloUserPuedeUsar       As Long
Public IntervaloFlechasCazadores    As Long

Public MinutosWs                    As Long
Public MinutosGuardarUsuarios       As Long
Public MinutosLimpia                As Long
Public Puerto                       As Integer

Public MAXPASOS                     As Long

Public BootDelBackUp                As Byte
Public DeNoche                      As Boolean

Public IpList                       As New Collection
Public ClientsCommandsQueue         As Byte

Public Type TCPESStats

    BytesEnviados As Double
    BytesRecibidos As Double
    BytesEnviadosXSEG As Long
    BytesRecibidosXSEG As Long
    BytesEnviadosXSEGMax As Long
    BytesRecibidosXSEGMax As Long
    BytesEnviadosXSEGCuando As Date
    BytesRecibidosXSEGCuando As Date

End Type

Public TCPESStats As TCPESStats

'Public ResetThread As New clsThreading

Function VersionOK(ByVal Ver As String) As Boolean
    VersionOK = (Ver = ULTIMAVERSION)

End Function

Public Function ValidarLoginMSG(ByVal n As Integer) As Integer

    On Error Resume Next

    Dim AuxInteger  As Integer
    Dim AuxInteger2 As Integer
    AuxInteger = SD(n)
    AuxInteger2 = SDM(n)
    ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)

End Function

Sub ReSpawnOrigPosNpcs()

    On Error Resume Next

    Dim i     As Integer
    Dim MiNPC As npc
        
    For i = 1 To LastNPC

        'OJO
        If Npclist(i).flags.NPCActive Then
              
            If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)

            End If
              
            'tildada por sugerencia de yind
            'If Npclist(i).Contadores.TiempoExistencia > 0 Then
            '        Call MuereNpc(i, 0)
            'End If
        End If
         
    Next i

End Sub

Sub RespGuerrasAngeles()

    On Error Resume Next

    Dim i       As Integer
    Dim MiNPC   As npc
    Dim Npc4    As Integer
    Dim Npc4Pos As WorldPos
    Npc4 = 941
    Npc4Pos.Map = 66
    Npc4Pos.X = 77
    Npc4Pos.Y = 77
            
    For i = 1 To LastNPC

        'OJO
        If Npclist(i).flags.NPCActive Then
              
            If Npclist(i).Numero = Npc4 Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                     
            End If
                
            'tildada por sugerencia de yind
            'If Npclist(i).Contadores.TiempoExistencia > 0 Then
            '        Call MuereNpc(i, 0)
            'End If
        End If
         
    Next i

End Sub

Sub RespGuerrasDemonio()

    On Error Resume Next

    Dim i       As Integer
    Dim MiNPC   As npc
    Dim Npc3    As Integer
    Dim Npc3Pos As WorldPos
    Npc3 = 940
    Npc3Pos.Map = 66
    Npc3Pos.X = 77
    Npc3Pos.Y = 23
            
    For i = 1 To LastNPC

        'OJO
        If Npclist(i).flags.NPCActive Then
              
            If Npclist(i).Numero = Npc3 Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                    
            End If
                
            'tildada por sugerencia de yind
            'If Npclist(i).Contadores.TiempoExistencia > 0 Then
            '        Call MuereNpc(i, 0)
            'End If
        End If
         
    Next i

End Sub

Sub WorldSave()

    Dim loopX As Long
    Dim hFile As Integer

    Call SendData(SendTarget.toall, 0, 0, "||°¨¨°(_.·´¯`·«¤°GUARDANDO CONFIGURACIÓN AOMANÍA°¤»·´¯`·._)°¨¨°" & FONTTYPE_WorldCarga)

    Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
    
    Call SaveConfig

    Dim j As Integer, k As Integer

    For j = 1 To NumMaps

        If MapInfo(j).BackUp = 1 Then k = k + 1
    Next j

    FrmStat.ProgressBar1.min = 0
    FrmStat.ProgressBar1.max = k
    FrmStat.ProgressBar1.value = 0

    For loopX = 1 To NumMaps
          
        If MapInfo(loopX).BackUp = 1 Then
          
            Call GrabarMapa(loopX, App.Path & "\WorldBackUp\Mapa" & loopX)
            FrmStat.ProgressBar1.value = FrmStat.ProgressBar1.value + 1

        End If

        'DoEvents
    Next loopX

    FrmStat.Visible = False

    If FileExist(DatPath & "\bkNpcs.dat", vbNormal) Then Kill (DatPath & "bkNpcs.dat")
    hFile = FreeFile()
    
    Open DatPath & "\bkNpcs.dat" For Output As hFile
    
    For loopX = 1 To LastNPC

        If Npclist(loopX).flags.BackUp = 1 Then
            Call BackUPnPc(loopX, hFile)

        End If

    Next loopX

    Close hFile
    
    Call SendData(SendTarget.toall, 0, 0, "||°¨¨°(_.·´¯`·«¤°CONFIGURACIÓN AOMANÍA GUARDADA°¤»·´¯`·._)°¨¨°" & FONTTYPE_WorldSave)

End Sub

Public Sub PurgarPenas()

    Dim i As Integer

    For i = 1 To LastUser

        If UserList(i).flags.UserLogged Then
          
            If UserList(i).Counters.Pena > 0 Then
                      
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                      
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                    Call SendData(SendTarget.toIndex, i, 0, "||Has sido liberado!" & FONTTYPE_INFO)

                End If
                      
            End If
              
        End If

    Next i

End Sub

Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GMName As String = "")
              
    UserList(UserIndex).Counters.Pena = Minutos
              
    Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
              
    ' If GmName = "" Then
    'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has sido encarcelado, deberas permanecer en la carcel " & Minutos & " minutos. Para poder salir ilegalmente de la carcel, deberás escribir /FIANZA y según la cantidad de minutos que te hayan encarcelado, pagás 200k por cada minuto y te libera." & FONTTYPE_INFO)
    ' Else
    'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & GmName & " te ha encarcelado, deberas permanecer en la carcel " & Minutos & " minutos o bien, pagar una fianza por tu pena. Sale 200.000 monedas de oro por minuto que te quede, tu oro se quitará automáticamente. Si decides pagar la fianza, escribe /FIANZA. Recuerda, no es gratis." & FONTTYPE_INFO)
    ' End If
              
End Sub

Public Sub BorrarUsuario(ByVal UserName As String)

    On Error Resume Next

    If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
        Kill CharPath & UCase$(UserName) & ".chr"

    End If

End Sub

Public Function BANCheck(ByVal Name As String) As Boolean

    BANCheck = (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban")) = 1)

End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean

    PersonajeExiste = FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)

End Function

Public Function UnBan(ByVal Name As String) As Boolean

    'Unban the character
    Call WriteVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban", "0")

    'Remove it from the banned people database
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NO REASON")

End Function

Public Function MD5ok(ByVal MD5Formateado As String) As Boolean

    Dim i As Integer

    If MD5ClientesActivado = 1 Then

        For i = 0 To UBound(MD5s)

            If (MD5Formateado = MD5s(i)) Then
                MD5ok = True
                Exit Function

            End If

        Next i

        MD5ok = False
    Else
        MD5ok = True

    End If

End Function

Public Sub BanIpAgrega(ByVal ip As String)

    BanIps.Add ip

    Call BanIpGuardar

End Sub

Public Function BanIpBuscar(ByVal ip As String) As Long

    Dim Dale  As Boolean
    Dim loopc As Long

    Dale = True
    loopc = 1

    Do While loopc <= BanIps.Count And Dale
        Dale = (BanIps.Item(loopc) <> ip)
        loopc = loopc + 1
    Loop

    If Dale Then
        BanIpBuscar = 0
    Else
        BanIpBuscar = loopc - 1

    End If

End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean

    On Error Resume Next

    Dim n As Long

    n = BanIpBuscar(ip)

    If n > 0 Then
        BanIps.Remove n
        BanIpGuardar
        BanIpQuita = True
    Else
        BanIpQuita = False

    End If

End Function

Public Sub BanIpGuardar()

    Dim ArchivoBanIp As String
    Dim ArchN        As Long
    Dim loopc        As Long

    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"

    ArchN = FreeFile()
    Open ArchivoBanIp For Output As #ArchN

    For loopc = 1 To BanIps.Count
        Print #ArchN, BanIps.Item(loopc)
    Next loopc

    Close #ArchN

End Sub

Public Sub BanIpCargar()

    Dim ArchN        As Long
    Dim Tmp          As String
    Dim ArchivoBanIp As String

    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"

    Do While BanIps.Count > 0
        BanIps.Remove 1
    Loop

    ArchN = FreeFile()
    Open ArchivoBanIp For Input As #ArchN

    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        BanIps.Add Tmp
    Loop

    Close #ArchN

End Sub

Public Sub ActualizaStatsES()

    Static TUlt      As Single
    Dim Transcurrido As Single

    Transcurrido = Timer - TUlt

    If Transcurrido >= 5 Then
        TUlt = Timer

        With TCPESStats
            .BytesEnviadosXSEG = CLng(.BytesEnviados / Transcurrido)
            .BytesRecibidosXSEG = CLng(.BytesRecibidos / Transcurrido)
            .BytesEnviados = 0
            .BytesRecibidos = 0
              
            If .BytesEnviadosXSEG > .BytesEnviadosXSEGMax Then
                .BytesEnviadosXSEGMax = .BytesEnviadosXSEG
                .BytesEnviadosXSEGCuando = CDate(Now)

            End If
              
            If .BytesRecibidosXSEG > .BytesRecibidosXSEGMax Then
                .BytesRecibidosXSEGMax = .BytesRecibidosXSEG
                .BytesRecibidosXSEGCuando = CDate(Now)

            End If
              
            If frmEstadisticas.Visible Then
                Call frmEstadisticas.ActualizaStats

            End If

        End With

    End If

End Sub

Public Function UserDarPrivilegioLevel(ByVal Name As String) As Long

    If EsDios(Name) Then
        UserDarPrivilegioLevel = 3
    ElseIf EsSemiDios(Name) Then
        UserDarPrivilegioLevel = 2
    ElseIf EsConsejero(Name) Then
        UserDarPrivilegioLevel = 1
    Else
        UserDarPrivilegioLevel = 0

    End If

End Function

