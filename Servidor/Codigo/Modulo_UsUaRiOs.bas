Attribute VB_Name = "UsUaRiOs"
Option Explicit

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)

    Dim DaExp As Integer

    DaExp = CInt(UserList(VictimIndex).Stats.ELV * 2)

    UserList(AttackerIndex).Stats.Exp = UserList(AttackerIndex).Stats.Exp + DaExp

    If UserList(AttackerIndex).Stats.Exp > MAXEXP Then UserList(AttackerIndex).Stats.Exp = MAXEXP

    'Lo mata
    Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||Has matado a " & UserList(VictimIndex).Name & "!" & FONTTYPE_Motd4)
    Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||Has ganado " & DaExp & " puntos de experiencia." & FONTTYPE_Motd4)

    Call SendData(SendTarget.ToIndex, VictimIndex, 0, "||" & UserList(AttackerIndex).Name & " te ha matado!" & FONTTYPE_Motd4)

    If TriggerZonaPelea(VictimIndex, AttackerIndex) <> TRIGGER6_PERMITE Then
        If (Not Criminal(VictimIndex)) Then
            UserList(AttackerIndex).Reputacion.AsesinoRep = UserList(AttackerIndex).Reputacion.AsesinoRep + vlASESINO * 2

            If UserList(AttackerIndex).Reputacion.AsesinoRep > MAXREP Then UserList(AttackerIndex).Reputacion.AsesinoRep = MAXREP
            UserList(AttackerIndex).Reputacion.BurguesRep = 0
            UserList(AttackerIndex).Reputacion.NobleRep = 0
            UserList(AttackerIndex).Reputacion.PlebeRep = 0
        Else
            UserList(AttackerIndex).Reputacion.NobleRep = UserList(AttackerIndex).Reputacion.NobleRep + vlNoble

            If UserList(AttackerIndex).Reputacion.NobleRep > MAXREP Then UserList(AttackerIndex).Reputacion.NobleRep = MAXREP

        End If

    End If

    If UserList(VictimIndex).GranPoder = 1 Then
        Call mod_GranPoder.UserMataPoder(VictimIndex, AttackerIndex)

    End If
    
    If UserList(AttackerIndex).Quest.ValidMatarUser > 0 Then
    
        If UserList(AttackerIndex).Quest.ValidMatarUser > 0 Then
            Call UserMataQuest(AttackerIndex, VictimIndex, UserList(AttackerIndex).Quest.Quest)
        End If
    
    End If

    Call UserDie(VictimIndex)

    Call modGuilds.PuntosClan(AttackerIndex, VictimIndex)

    If UserList(VictimIndex).Faccion.ArmadaReal = 1 Then
        UserList(AttackerIndex).Stats.CleroMatados = UserList(AttackerIndex).Stats.CleroMatados + 1
    ElseIf UserList(VictimIndex).Faccion.FuerzasCaos = 1 Then
        UserList(AttackerIndex).Stats.AbbadonMatados = UserList(AttackerIndex).Stats.AbbadonMatados + 1
    ElseIf UserList(VictimIndex).Faccion.Nemesis = 1 Then
        UserList(AttackerIndex).Stats.TinieblaMatados = UserList(AttackerIndex).Stats.TinieblaMatados + 1
    ElseIf UserList(VictimIndex).Faccion.Templario = 1 Then
        UserList(AttackerIndex).Stats.TemplarioMatados = UserList(AttackerIndex).Stats.TemplarioMatados + 1

    End If

    If UserList(AttackerIndex).Stats.UsuariosMatados < 32000 Then UserList(AttackerIndex).Stats.UsuariosMatados = UserList(AttackerIndex).Stats.UsuariosMatados + 1
    'Log
    Call LogAsesinato(UserList(AttackerIndex).Name & " asesino a " & UserList(VictimIndex).Name)

End Sub

Sub RevivirUsuario(ByVal UserIndex As Integer)
     
    If NameDay = "Noche" And UCase$(UserList(UserIndex).Raza) = "LICANTROPO" And UserList(UserIndex).flags.Navegando = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes resucitar navegando, la oscuridad te lo impide" & FONTTYPE_INFO)
        Exit Sub

    End If

    UserList(UserIndex).flags.Muerto = 0
    UserList(UserIndex).Stats.MinHP = 35

    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta

    'No puede estar empollando
    UserList(UserIndex).flags.EstaEmpo = 0
    UserList(UserIndex).EmpoCont = 0

    If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP

    End If

    Call DarCuerpoDesnudo(UserIndex)

    If NameDay = "Noche" And UCase$(UserList(UserIndex).Raza) = "LICANTROPO" And UserList(UserIndex).flags.Licantropo = 1 Then
        Call DoLicantropo(UserIndex)

    End If

   Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
            UserIndex).OrigChar.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, _
            UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)
            
    '[/MaTeO 9]
    Call EnviarHP(UserIndex)
    Call EnviarSta(UserIndex)

    If UserList(UserIndex).flags.bandas = True Then
        Call Transforma(UserIndex)

    End If

End Sub

'[MaTeO 9]
Sub ChangeUserChar(ByVal sndRoute As Byte, _
                   ByVal sndIndex As Integer, _
                   ByVal sndMap As Integer, _
                   ByVal UserIndex As Integer, _
                   ByVal Body As Integer, _
                   ByVal Head As Integer, _
                   ByVal heading As Byte, _
                   ByVal Arma As Integer, _
                   ByVal Escudo As Integer, _
                   ByVal Casco As Integer, _
                   ByVal Alas As Integer)
'[/MaTeO 9]

    If UserList(UserIndex).Asedio.Participando Then
        If UserList(UserIndex).Raza = "Humano" Or _
           UserList(UserIndex).Raza = "Elfo" Or _
           UserList(UserIndex).Raza = "Elfo Oscuro" Then
            If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
                Select Case UserList(UserIndex).Asedio.Team
                Case Equipos.Azul
                    Body = 516
                Case Equipos.Negro
                    Body = 508
                Case Equipos.Rojo
                    Body = 520
                Case Equipos.Verde
                    Body = 512
                End Select
            Else
                Select Case UserList(UserIndex).Asedio.Team
                Case Equipos.Azul
                    Body = 514
                Case Equipos.Negro
                    Body = 506
                Case Equipos.Rojo
                    Body = 518
                Case Equipos.Verde
                    Body = 510
                End Select
            End If
        Else
            If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
                Select Case UserList(UserIndex).Asedio.Team
                Case Equipos.Azul
                    Body = 515
                Case Equipos.Negro
                    Body = 507
                Case Equipos.Rojo
                    Body = 519
                Case Equipos.Verde
                    Body = 511
                End Select
            Else
                Select Case UserList(UserIndex).Asedio.Team
                Case Equipos.Azul
                    Body = 517
                Case Equipos.Negro
                    Body = 509
                Case Equipos.Rojo
                    Body = 521
                Case Equipos.Verde
                    Body = 513
                End Select
            End If
        End If
    End If
   
 
  If UserList(UserIndex).flags.MetamorfosisLicantropo = 1 Then
        Body = 173
        Head = 0
       heading = heading
       Arma = NingunArma
       Escudo = NingunEscudo
       Casco = NingunCasco
       Alas = NingunAlas
    Else
    UserList(UserIndex).char.Body = Body
    UserList(UserIndex).char.Head = Head
    UserList(UserIndex).char.heading = heading
    UserList(UserIndex).char.WeaponAnim = Arma
    UserList(UserIndex).char.ShieldAnim = Escudo
    UserList(UserIndex).char.CascoAnim = Casco
    UserList(UserIndex).char.Alas = Alas
  End If
  
 
    
    

    If sndRoute = SendTarget.ToMap Then
        '[MaTeO 9]
        Call SendToUserArea(UserIndex, "CP" & UserList(UserIndex).char.CharIndex & "," & Body & "," & Head & "," & heading & "," & Arma & "," & _
                                       Escudo & "," & UserList(UserIndex).char.FX & "," & UserList(UserIndex).char.loops & "," & Casco & "," & Alas)
        '[/MaTeO 9]
    Else
        '[MaTeO 9]
        Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(UserIndex).char.CharIndex & "," & Body & "," & Head & "," & heading & "," & Arma _
                                                & "," & Escudo & "," & UserList(UserIndex).char.FX & "," & UserList(UserIndex).char.loops & "," & Casco & "," & Alas)

        '[/MaTeO 9]
    End If

End Sub

Sub EnviarSubirNivel(ByVal UserIndex As Integer, ByVal Puntos As Integer)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "SUNI" & Puntos)

End Sub

Sub EnviarSkills(ByVal UserIndex As Integer)
    Dim i As Integer
    Dim cad As String

    For i = 1 To NUMSKILLS
        cad = cad & UserList(UserIndex).Stats.UserSkills(i) & ","
    Next i

    SendData SendTarget.ToIndex, UserIndex, 0, "SKILLS" & cad$

End Sub

Sub EnviarFama(ByVal UserIndex As Integer)
    Dim cad As String

    cad = cad & UserList(UserIndex).Reputacion.AsesinoRep & ","
    cad = cad & UserList(UserIndex).Reputacion.BandidoRep & ","
    cad = cad & UserList(UserIndex).Reputacion.BurguesRep & ","
    cad = cad & UserList(UserIndex).Reputacion.LadronesRep & ","
    cad = cad & UserList(UserIndex).Reputacion.NobleRep & ","
    cad = cad & UserList(UserIndex).Reputacion.PlebeRep & ","

    Dim L As Long

    L = (-UserList(UserIndex).Reputacion.AsesinoRep) + (-UserList(UserIndex).Reputacion.BandidoRep) + UserList(UserIndex).Reputacion.BurguesRep + ( _
        -UserList(UserIndex).Reputacion.LadronesRep) + UserList(UserIndex).Reputacion.NobleRep + UserList(UserIndex).Reputacion.PlebeRep
    L = L / 6

    UserList(UserIndex).Reputacion.Promedio = L

    cad = cad & UserList(UserIndex).Reputacion.Promedio

    SendData SendTarget.ToIndex, UserIndex, 0, "FAMA" & cad

End Sub

Sub EnviarFamaGM(ByVal UserIndex As Integer, ByVal rData As Integer)
    Dim cad As String

    cad = cad & UserList(rData).Reputacion.AsesinoRep & ","
    cad = cad & UserList(rData).Reputacion.BandidoRep & ","
    cad = cad & UserList(rData).Reputacion.BurguesRep & ","
    cad = cad & UserList(rData).Reputacion.LadronesRep & ","
    cad = cad & UserList(rData).Reputacion.NobleRep & ","
    cad = cad & UserList(rData).Reputacion.PlebeRep & ","

    Dim L As Long

    L = (-UserList(rData).Reputacion.AsesinoRep) + (-UserList(rData).Reputacion.BandidoRep) + UserList(rData).Reputacion.BurguesRep + ( _
        -UserList(rData).Reputacion.LadronesRep) + UserList(rData).Reputacion.NobleRep + UserList(rData).Reputacion.PlebeRep
    L = L / 6

    UserList(rData).Reputacion.Promedio = L

    cad = cad & UserList(rData).Reputacion.Promedio

    SendData SendTarget.ToIndex, UserIndex, 0, "FAMA" & cad

End Sub

Sub EnviarAtrib(ByVal UserIndex As Integer)
    Dim i As Integer
    Dim cad As String

    For i = 1 To NUMATRIBUTOS
        cad = cad & UserList(UserIndex).Stats.UserAtributos(i) & ","
    Next
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ATR" & cad)

End Sub

Sub EnviarAtribGM(ByVal UserIndex As Integer, ByVal rData As Integer)
    Dim i As Integer
    Dim cad As String

    For i = 1 To NUMATRIBUTOS
        cad = cad & UserList(rData).Stats.UserAtributos(i) & ","
    Next
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ATR" & cad)

End Sub

Public Sub EnviarMiniEstadisticasGM(ByVal UserIndex As Integer, ByVal rData As Integer)
    If UserList(rData).Faccion.ArmadaReal = 1 Then
        UserArmada = "CLERO"
        UserRecompensas = UserList(rData).Faccion.RecompensasReal
    ElseIf UserList(rData).Faccion.FuerzasCaos = 1 Then
        UserArmada = "ABBADON"
        UserRecompensas = UserList(rData).Faccion.RecompensasCaos
    ElseIf UserList(rData).Faccion.Nemesis = 1 Then
        UserArmada = "TINIEBLA"
        UserRecompensas = UserList(rData).Faccion.RecompensasNemesis
    ElseIf UserList(rData).Faccion.Templario = 1 Then
        UserArmada = "TEMPLARIO"
        UserRecompensas = UserList(rData).Faccion.RecompensasTemplaria
    End If


    With UserList(rData)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MEST" & .Faccion.CiudadanosMatados & "," & .Faccion.CriminalesMatados & "," & _
                                                        .Stats.UsuariosMatados & "," & .Stats.NPCsMuertos & "," & .Clase & "," & .Counters.Pena & "," & .Raza & "," & .Clan.PuntosClan & "," & .Name & "," & _
                                                        .Genero & "," & .Stats.PuntosRetos & "," & .Stats.PuntosTorneo & "," & .Stats.PuntosDuelos & "," & .Stats.ELV & "," & .Stats.ELU & "," & .Stats.Exp & "," & _
                                                        .Stats.MinHP & "," & .Stats.MaxHP & "," & .Stats.MinMAN & "," & .Stats.MaxMAN & "," & .Stats.MinSta & "," & .Stats.MaxSta & "," & .Stats.GLD & "," & _
                                                        .Stats.Banco & "," & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & "," & .Stats.SkillPts & "," & .Clan.ParticipoClan & "," & .Stats.AbbadonMatados & "," & .Stats.CleroMatados & "," & _
                                                        .Stats.TinieblaMatados & "," & .Stats.TemplarioMatados & "," & UserArmada & "," & .Faccion.Reenlistadas & "," & UserRecompensas & "," & _
                                                        .Faccion.CiudadanosMatados & "," & .Faccion.CriminalesMatados & "," & .Faccion.FEnlistado)
    Debug.Print .Genero
    End With
    
    
End Sub

Public Sub EnviarMiniEstadisticas(ByVal UserIndex As Integer)

    If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
        UserArmada = "CLERO"
        UserRecompensas = UserList(UserIndex).Faccion.RecompensasReal
    ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
        UserArmada = "ABBADON"
        UserRecompensas = UserList(UserIndex).Faccion.RecompensasCaos
    ElseIf UserList(UserIndex).Faccion.Nemesis = 1 Then
        UserArmada = "TINIEBLA"
        UserRecompensas = UserList(UserIndex).Faccion.RecompensasNemesis
    ElseIf UserList(UserIndex).Faccion.Templario = 1 Then
        UserArmada = "TEMPLARIO"
        UserRecompensas = UserList(UserIndex).Faccion.RecompensasTemplaria
    End If


    With UserList(UserIndex)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MEST" & .Faccion.CiudadanosMatados & "," & .Faccion.CriminalesMatados & "," & _
                                                        .Stats.UsuariosMatados & "," & .Stats.NPCsMuertos & "," & .Clase & "," & .Counters.Pena & "," & .Raza & "," & .Clan.PuntosClan & "," & .Name & "," & _
                                                        .Genero & "," & .Stats.PuntosRetos & "," & .Stats.PuntosTorneo & "," & .Stats.PuntosDuelos & "," & .Stats.ELV & "," & .Stats.ELU & "," & .Stats.Exp & "," & _
                                                        .Stats.MinHP & "," & .Stats.MaxHP & "," & .Stats.MinMAN & "," & .Stats.MaxMAN & "," & .Stats.MinSta & "," & .Stats.MaxSta & "," & .Stats.GLD & "," & _
                                                        .Stats.Banco & "," & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & "," & .Stats.SkillPts & "," & .Clan.ParticipoClan & "," & .Stats.AbbadonMatados & "," & .Stats.CleroMatados & "," & _
                                                        .Stats.TinieblaMatados & "," & .Stats.TemplarioMatados & "," & UserArmada & "," & .Faccion.Reenlistadas & "," & UserRecompensas & "," & _
                                                        .Faccion.CiudadanosMatados & "," & .Faccion.CriminalesMatados & "," & .Faccion.FEnlistado)
    End With

End Sub

Sub EraseUserChar(sndRoute As Byte, _
                  sndIndex As Integer, _
                  sndMap As Integer, _
                  UserIndex As Integer)

    On Error GoTo ErrorHandler
   
    CharList(UserList(UserIndex).char.CharIndex) = 0
    
    If UserList(UserIndex).char.CharIndex = LastChar Then

        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1

            If LastChar = 0 Then Exit Do
        Loop

    End If
    
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
    
    'binmode: grrrrr.. para algo esta sndroute, sndindex, sndmap...
    Call SendData(sndRoute, sndIndex, sndMap, "BP" & UserList(UserIndex).char.CharIndex)
    
    UserList(UserIndex).char.CharIndex = 0
    
    NumChars = NumChars - 1
    
    Exit Sub
    
ErrorHandler:
    Call LogError("Error en EraseUserchar")

End Sub

Sub MakeUserChar(ByVal sndRoute As SendTarget, _
                 ByVal sndIndex As Integer, _
                 ByVal sndMap As Integer, _
                 ByVal UserIndex As Integer, _
                 ByVal Map As Integer, _
                 ByVal X As Integer, _
                 ByVal Y As Integer)

    On Local Error GoTo hayerror
    
    If UserList(UserIndex).flags.MetamorfosisLicantropo = 1 Then
        UserList(UserIndex).char.Body = 173
       UserList(UserIndex).char.Head = 0
       UserList(UserIndex).char.heading = UserList(UserIndex).char.heading
       UserList(UserIndex).char.WeaponAnim = NingunArma
       UserList(UserIndex).char.ShieldAnim = NingunEscudo
       UserList(UserIndex).char.CascoAnim = NingunCasco
       UserList(UserIndex).char.Alas = NingunAlas
  End If
  
    Dim CharIndex As Integer

    If InMapBounds(Map, X, Y) Then

        With UserList(UserIndex)

            'If needed make a new character in list
            If .char.CharIndex = 0 Then
                CharIndex = NextOpenCharIndex
                .char.CharIndex = CharIndex
                CharList(CharIndex) = UserIndex

            End If

            'Place character on map
            MapData(Map, X, Y).UserIndex = UserIndex

            'Send make character command to clients
            Dim klan As String

            If .GuildIndex > 0 Then
                klan = Guilds(.GuildIndex).GuildName

            End If

            Dim bCr As Byte
            Dim SendPrivilegios As Byte

            If Criminal(UserIndex) Then
                bCr = 1
            Else
                bCr = 0

            End If

            If .Faccion.FuerzasCaos = 1 Then
                bCr = 2
            End If

            If .Faccion.Templario = 1 Then
                bCr = 3
            End If

            If .Faccion.ArmadaReal = 1 Then
                bCr = 4
            End If

            If .Faccion.Nemesis = 1 Then
                bCr = 5
            End If
            Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "CVB" & UserList(UserIndex).char.CharIndex & "," & UserList(UserIndex).flags.CvcBlue)
            Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "CVR" & UserList(UserIndex).char.CharIndex & "," & UserList(UserIndex).flags.CvcRed)

            If Len(klan) <> 0 Then
                If sndRoute = SendTarget.ToIndex Then

                    Dim code As String

                    If .flags.Privilegios > PlayerType.User Then
                        If .showName Then

                            code = .char.Body & "," & .char.Head & "," & .char.heading & "," & .char.CharIndex & "," & X & "," & Y & "," & _
                                   .char.WeaponAnim & "," & .char.ShieldAnim & "," & .char.FX & "," & 999 & "," & .char.CascoAnim & "," & .Name & _
                                 " <" & klan & ">" & "" & "," & bCr & "," & IIf(.flags.EsRolesMaster, 5, .flags.Privilegios) & "," & .char.Alas _
                                 & "," & .PartyIndex

                            Call SendData(sndRoute, sndIndex, sndMap, "CC" & code)
                        Else

                            'Hide the name and clan
                            code = .char.Body & "," & .char.Head & "," & .char.heading & "," & .char.CharIndex & "," & X & "," & Y & "," & _
                                   .char.WeaponAnim & "," & .char.ShieldAnim & "," & .char.FX & "," & 999 & "," & .char.CascoAnim & ",," & bCr & _
                                   "," & IIf(.flags.EsRolesMaster, 5, .flags.Privilegios) & "," & .char.Alas & "," & .PartyIndex

                            Call SendData(sndRoute, sndIndex, sndMap, "CC" & code)

                        End If

                    Else

                        code = .char.Body & "," & .char.Head & "," & .char.heading & "," & .char.CharIndex & "," & X & "," & Y & "," & _
                               .char.WeaponAnim & "," & .char.ShieldAnim & "," & .char.FX & "," & 999 & "," & .char.CascoAnim & "," & .Name & " <" _
                             & klan & ">" & "" & "," & bCr & "," & IIf(.flags.PertAlCons = 1, 4, IIf(.flags.PertAlConsCaos = 1, 6, 0)) & "," & _
                               .char.Alas & "," & .PartyIndex

                        Call SendData(sndRoute, sndIndex, sndMap, "CC" & code)

                    End If
                 'Paquete envio de quest
                 Call SendData(sndRoute, sndIndex, sndMap, "XC" & 1)
                 
                ElseIf sndRoute = SendTarget.ToMap Then
                    Call AgregarUser(UserIndex, .Pos.Map)

                End If

            Else    'if tiene clan

                If sndRoute = SendTarget.ToIndex Then

                    If .flags.Privilegios > PlayerType.User Then
                        If .showName Then
                            Call SendData(SendTarget.ToIndex, sndIndex, sndMap, "BC" & .char.Body & "," & .char.Head & "," & .char.heading & "," & _
                                                                                .char.CharIndex & "," & X & "," & Y & "," & .char.WeaponAnim & "," & .char.ShieldAnim & "," & .char.FX & "," & _
                                                                                999 & "," & .char.CascoAnim & "," & .Name & "" & "," & bCr & "," & IIf(.flags.EsRolesMaster, 5, _
                                                                                                                                                       .flags.Privilegios) & "," & .char.Alas)
                        Else
                            Call SendData(SendTarget.ToIndex, sndIndex, sndMap, "BC" & .char.Body & "," & .char.Head & "," & .char.heading & "," & _
                                                                                .char.CharIndex & "," & X & "," & Y & "," & .char.WeaponAnim & "," & .char.ShieldAnim & "," & .char.FX & "," & _
                                                                                999 & "," & .char.CascoAnim & ",," & bCr & "," & IIf(.flags.EsRolesMaster, 5, .flags.Privilegios) & "," & _
                                                                                .char.Alas)

                        End If

                    Else
                        Call SendData(SendTarget.ToIndex, sndIndex, sndMap, "BC" & .char.Body & "," & .char.Head & "," & .char.heading & "," & _
                                                                            .char.CharIndex & "," & X & "," & Y & "," & .char.WeaponAnim & "," & .char.ShieldAnim & "," & .char.FX & "," & 999 _
                                                                          & "," & .char.CascoAnim & "," & .Name & "" & "," & bCr & "," & IIf(.flags.PertAlCons = 1, 4, IIf( _
                                                                                                                                                                         .flags.PertAlConsCaos = 1, 6, 0)) & "," & .char.Alas)

                    End If

                ElseIf sndRoute = SendTarget.ToMap Then
                    Call AgregarUser(UserIndex, .Pos.Map)

                End If

            End If   'if clan

        End With

    End If

    Exit Sub

hayerror:
    LogError ("MakeUserChar: num: " & err.Number & " desc: " & err.Description)
    'Resume Next
    Call CloseSocket(UserIndex)

End Sub

Sub CheckUserLevel(ByVal UserIndex As Integer)

    On Error GoTo errhandler

    Dim Pts         As Integer

    Dim AumentoLVL  As Byte

    Dim AumentoHIT  As Integer

    Dim AumentoMANA As Integer

    Dim AumentoSTA  As Integer

    Dim AumentoHP   As Integer

    Dim LastLvl     As Byte

    Dim LastHit     As Integer

    Dim LastMana    As Integer

    Dim LastSTA     As Integer

    Dim LastHp      As Integer

    Dim Promedio    As Double

    Dim ExPromedio  As Double

    Dim WasNewbie   As Boolean

    WasNewbie = EsNewbie(UserIndex)

    '¿Alcanzo el maximo nivel?
    With UserList(UserIndex)

        LastLvl = .Stats.ELV
        LastHit = .Stats.MaxHit
        LastMana = .Stats.MaxMAN
        LastSTA = .Stats.MaxSta
        LastHp = .Stats.MaxHP

        'Si exp >= then Exp para subir de nivel entonce subimos el nivel
        'If .Stats.Exp >= .Stats.ELU Then
        Do While .Stats.Exp >= .Stats.ELU

            'Checkea si alcanzó el máximo nivel
            If .Stats.ELV = STAT_MAXELV Then

                .Stats.Exp = 0
                .Stats.ELU = 0
                Exit Do

            End If

            If .Stats.ELV = 1 Then

                Pts = 10
            Else

                If .Clase = "TRABAJADOR" Then

                    Pts = Pts + 10
                Else
                    Pts = Pts + 5

                End If

            End If

            ' rodra , no avisa total no hay =)
            .Stats.ELV = .Stats.ELV + 1
            .Stats.Exp = .Stats.Exp - .Stats.ELU

            .Stats.ELU = levelELU(.Stats.ELV)
            
            If UserList(UserIndex).Stats.ELV < 11 Then

                UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.5
            ElseIf UserList(UserIndex).Stats.ELV < 25 Then
                UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.4
            ElseIf UserList(UserIndex).Stats.ELV < 30 Then
                UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
            ElseIf UserList(UserIndex).Stats.ELV < 40 Then
                UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.2
            Else
                UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.1

            End If

            Call AumentoStatsClase(UserIndex, UCase$(.Clase), AumentoHP, AumentoMANA, AumentoSTA, AumentoHIT)

            'Actualizamos HitPoints
            .Stats.MaxHP = .Stats.MaxHP + AumentoHP

            If .Stats.MaxHP > STAT_MAXHP Then .Stats.MaxHP = STAT_MAXHP

            'Actualizamos Stamina
            .Stats.MaxSta = .Stats.MaxSta + AumentoSTA

            If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA

            'Actualizamos Mana
            .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA

            If .Stats.ELV < 36 Then
                If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN

            Else

                If .Stats.MaxMAN > 9999 Then .Stats.MaxMAN = 9999

            End If

            'Actualizamos Golpe Máximo
            .Stats.MaxHit = .Stats.MaxHit + AumentoHIT

            If .Stats.ELV < 36 Then
                If .Stats.MaxHit > STAT_MAXHIT_UNDER36 Then .Stats.MaxHit = STAT_MAXHIT_UNDER36

            Else

                If .Stats.MaxHit > STAT_MAXHIT_OVER36 Then .Stats.MaxHit = STAT_MAXHIT_OVER36

            End If

            'Actualizamos Golpe Mínimo
            .Stats.MinHit = .Stats.MinHit + AumentoHIT

            If .Stats.ELV < 36 Then
                If .Stats.MinHit > STAT_MAXHIT_UNDER36 Then .Stats.MinHit = STAT_MAXHIT_UNDER36

            Else

                If .Stats.MinHit > STAT_MAXHIT_OVER36 Then .Stats.MinHit = STAT_MAXHIT_OVER36

            End If

            'Promedio CHOTS

            Call LogDesarrollo(Date$ & " " & .Name & " paso a nivel " & .Stats.ELV & " gano HP: " & AumentoHP)

        Loop

        AumentoLVL = .Stats.ELV - LastLvl

        If AumentoLVL > 0 Then

            If .Stats.ELV = STAT_MAXELV Then
                If Criminal(UserIndex) Then

                    Call AgregarHechizoEspecial(UserIndex, H_Demonio)
                    Call AgregarHechizoEspecial(UserIndex, H_DemonioII)
                Else
                    Call AgregarHechizoEspecial(UserIndex, H_Angel)
                    Call AgregarHechizoEspecial(UserIndex, H_AngelII)

                End If

            End If

            AumentoHIT = .Stats.MaxHit - LastHit
            AumentoMANA = .Stats.MaxMAN - LastMana
            AumentoSTA = .Stats.MaxSta - LastSTA
            AumentoHP = .Stats.MaxHP - LastHp

            .Stats.MinHP = .Stats.MaxHP
            .Stats.MinMAN = .Stats.MaxMAN

            Call SendData(SendTarget.ToPCArea, UserIndex, .Pos.Map, "TW" & SND_NIVEL)
            Call SendData(SendTarget.ToPCArea, UserIndex, .Pos.Map, "||" & vbCyan & "°" & "Has pasado al nivel " & .Stats.ELV & "°" & CStr(.char.CharIndex))

            If AumentoLVL = 1 Then

                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has subido de nivel!" & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has subido " & AumentoLVL & " Niveles!." & FONTTYPE_INFO)

            End If
            
            If .Quest.Start = 1 Then
                If QuestList(.Quest.Quest).MaxNivel > 0 Then

                    Call SuperaNivelMaximoQuest(UserIndex, .Quest.Quest)

                End If

            End If

            'Notificamos al user
            If AumentoHP > 0 Then

                SendData SendTarget.ToIndex, UserIndex, 0, "||Has ganado " & AumentoHP & " puntos de vida." & FONTTYPE_INFO

            End If

            If AumentoSTA > 0 Then

                SendData SendTarget.ToIndex, UserIndex, 0, "||Has ganado " & AumentoSTA & " puntos de stamina." & FONTTYPE_INFO

            End If

            If AumentoMANA > 0 Then

                SendData SendTarget.ToIndex, UserIndex, 0, "||Has ganado " & AumentoMANA & " puntos de mana." & FONTTYPE_INFO

            End If

            If AumentoHIT > 0 Then

                SendData SendTarget.ToIndex, UserIndex, 0, "||Tu golpe maximo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
                SendData SendTarget.ToIndex, UserIndex, 0, "||Tu golpe minimo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO

            End If

            'Borrar final del testeo
            If EsNewbie(UserIndex) Then

                .Stats.GLD = .Stats.GLD + "8000"

            End If
            
            'Quest Faccion
            If .Faccion.ArmadaReal = 1 Or .Faccion.FuerzasCaos = 1 Or .Faccion.Nemesis = 1 Or .Faccion.Templario = 1 Then
                If .Stats.ELV = 45 Then

                    .Faccion.ActiveQuest = 1
                    .Faccion.Quest = .Faccion.Quest + 1
                    Call SendData(ToIndex, UserIndex, 0, "||¡Nueva misión faccion! Traeme 3 plumas y entregamelas con el comando /QUESTENTREGA para recibir la recompensa." & FONTTYPE_Motd5)
                    Call DarIconoFaccion(UserIndex)
                    
                ElseIf .Stats.ELV = 55 Then
                
                    .Faccion.ActiveQuest = 1
                    .Faccion.Quest = .Faccion.Quest + 1
                    Call SendData(ToIndex, UserIndex, 0, "||¡Nueva misión faccion! Traeme 3 plumas y entregamelas con el comando /QUESTENTREGA para recibir la recompensa." & FONTTYPE_Motd5)
                    Call DarIconoFaccion(UserIndex)

                End If
            End If
            
            If .flags.Privilegios = PlayerType.User Then
                If .Stats.ELV > MaxLevel Then

                    MaxLevel = .Stats.ELV
                    UserMaxLevel = .Name

                End If

                Call CriCiuMaxLvl(UserIndex)

            End If

            If .Stats.ELV > 13 Then

                ExPromedio = Round((.Stats.MaxHP - AumentoHP) / (.Stats.ELV - 1), 2)
                Promedio = Round(.Stats.MaxHP / .Stats.ELV, 2)

                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El Promedio de vida de tu Personaje era de " & CStr(ExPromedio) & FONTTYPE_ORO)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahora el Promedio es de " & CStr(Promedio) & FONTTYPE_ORO)

            End If

            Call SendUserStatsBox(UserIndex)

        End If

        If Not EsNewbie(UserIndex) And WasNewbie Then

            Call QuitarNewbieObj(UserIndex)

            If UCase$(MapInfo(.Pos.Map).Restringir) = "SI" Then

                Call WarpUserChar(UserIndex, 34, 45, 50, True)
                '  Call SendData(SendTarget.toindex, UserIndex, 0, "||Debes abandonar el Dungeon Newbie." & _
                   FONTTYPE_WARNING)

            End If

        End If

        If Pts > 0 Then

            Call EnviarSkills(UserIndex)
            Call EnviarSubirNivel(UserIndex, Pts)

            .Stats.SkillPts = .Stats.SkillPts + Pts

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has ganado un total de " & CStr(Pts) & " skillpoints." & FONTTYPE_INFO)

        End If

        If .Sagrada.Enabled = 1 Then

            Call ChangeSagradaHit(UserIndex)

        End If

    End With

    Exit Sub

errhandler:
    LogError ("Error en la subrutina CheckUserLevel")

End Sub

Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean

    PuedeAtravesarAgua = UserList(UserIndex).flags.Navegando = 1

End Function

'Private Sub NuevaPosCasper(ByVal Muerto As Integer)
'
'    Dim WorldPos As WorldPos
'    Dim WorldPos2 As WorldPos
'    Dim WorldPos3 As WorldPos
'    Dim WorldPos4 As WorldPos
'
'    WorldPos.Y = UserList(Muerto).pos.Y + 1
''    WorldPos.X = UserList(Muerto).pos.X
 '   WorldPos.Map = UserList(Muerto).pos.Map
'
'    WorldPos2.Y = UserList(Muerto).pos.Y
'    WorldPos2.X = UserList(Muerto).pos.X + 1
'    WorldPos2.Map = UserList(Muerto).pos.Map
'
'    WorldPos3.Y = UserList(Muerto).pos.Y - 1
'    WorldPos3.X = UserList(Muerto).pos.X
'    WorldPos3.Map = UserList(Muerto).pos.Map
'
'    WorldPos4.Y = UserList(Muerto).pos.Y
'    WorldPos4.X = UserList(Muerto).pos.X - 1
'    WorldPos4.Map = UserList(Muerto).pos.Map
'
'    If LegalPos(WorldPos.Map, WorldPos.X, WorldPos.Y, False) Then
'        Call MoveUserChar(Muerto, eHeading.NORTH)
'        Exit Sub
'    ElseIf LegalPos(WorldPos2.Map, WorldPos2.X, WorldPos2.Y, False) Then
'        Call MoveUserChar(Muerto, eHeading.EAST)
'        Exit Sub
'    ElseIf LegalPos(WorldPos3.Map, WorldPos3.X, WorldPos3.Y, False) Then
'        Call MoveUserChar(Muerto, eHeading.SOUTH)
'        Exit Sub
'    ElseIf LegalPos(WorldPos4.Map, WorldPos4.X, WorldPos4.Y, False) Then
'        Call MoveUserChar(Muerto, eHeading.WEST)
'        Exit Sub
'
'    End If
'
'End Sub

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As Byte, codigo As Byte, SuX As Byte, SuY As Byte)


    Dim nPos          As WorldPos

    Dim sailing       As Boolean

    Dim CasperIndex   As Integer

    Dim CasperHeading As eHeading

    Dim CasPerPos     As WorldPos

    Dim isAdminInvi   As Boolean

    Dim i             As Integer

    With UserList(UserIndex)
            
        sailing = PuedeAtravesarAgua(UserIndex)

        nPos = .Pos
        Call HeadtoPos(nHeading, nPos)

        isAdminInvi = (.flags.AdminInvisible = 1)

        'If CasperIndex > 0 Then
        'If UserList(CasperIndex).flags.Muerto = 1 And UserList(UserIndex).flags.Muerto = 0 Then
        'Call NuevaPosCasper(CasperIndex)
        'End If
        'End If

        If MoveToLegalPos(.Pos.Map, nPos.X, nPos.Y, sailing, Not sailing) Then
        
          Call Corr_MandaPosicion(UserIndex, SuX, SuY, codigo, 1)

            'si no estoy solo en el mapa...
            If MapInfo(.Pos.Map).NumUsers > 1 Then

                CasperIndex = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex

                'Si hay un usuario, y paso la validacion, entonces es un casper
                If CasperIndex > 0 Then

                    ' Los admins invisibles no pueden patear caspers
                    If Not isAdminInvi Then

                        CasperHeading = InvertHeading(nHeading)
                        CasPerPos = UserList(CasperIndex).Pos
                        Call HeadtoPos(CasperHeading, CasPerPos)

                        With UserList(CasperIndex)

                            ' Si es un admin invisible, no se avisa a los demas clientes
                            If Not .flags.AdminInvisible = 1 Then
                                Call SendToUserAreaButindex(CasperIndex, "+" & .char.CharIndex & "," & CasPerPos.X & "," & CasPerPos.Y)

                            End If

                            Call SendData(SendTarget.ToIndex, CasperIndex, 0, "$" & CasperHeading)

                            'Update map and user pos
                            .Pos = CasPerPos
                            .char.heading = CasperHeading
                            MapData(.Pos.Map, CasPerPos.X, CasPerPos.Y).UserIndex = CasperIndex

                        End With
                        
                        
                        'Actualizamos las áreas de ser necesario
                        Call ModAreas.CheckUpdateNeededUser(CasperIndex, CasperHeading)

                    End If
                    
                    Else
                    
                    Call Corr_MandaPosicion(UserIndex, SuX, SuY, codigo, 1)

                End If

                If Not isAdminInvi Then Call SendToUserAreaButindex(UserIndex, "+" & .char.CharIndex & "," & nPos.X & "," & nPos.Y)
               
            End If

            ' Los admins invisibles no pueden patear caspers
            If Not (isAdminInvi And CasperIndex <> 0) Then

                Dim oldUserIndex As Integer

                oldUserIndex = MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex

                ' Si no hay intercambio de pos con nadie
                If oldUserIndex = UserIndex Then
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0

                End If

                'Update map and user pos
                .Pos = nPos
                .char.heading = nHeading
                MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex

                If ZonaCura(UserIndex) Then Call AutoCuraUser(UserIndex)

                'Actualizamos las áreas de ser necesario
                
                Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading)

            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "PU" & .Pos.X & "," & .Pos.Y)
        

                .Counters.validInputs = (GetTickCount() And &H7FFFFFFF) + .char.delay + 20

            End If

        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "PU" & .Pos.X & "," & .Pos.Y)

            .Counters.validInputs = (GetTickCount() And &H7FFFFFFF) + .char.delay + 20

        End If

        If .Counters.Trabajando Then .Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1
        
        If UserList(UserIndex).GuildIndex > 0 Then
                 
               Call EnviaPosClan(UserIndex)
            
        End If
        
    End With

End Sub

Public Sub EnviaPosClan(ByVal UserIndex As Integer)

    Dim IDUser As Integer

    Dim i      As Integer
          
     If Guilds(UserList(UserIndex).GuildIndex).GuildName <> "" Then
          
    For i = 1 To LastUser

        If UCase$(UserList(i).Name) <> UCase$(UserList(UserIndex).Name) Then
           If UserList(i).GuildIndex > 0 Then
            If UCase$(Guilds(UserList(i).GuildIndex).GuildName) = UCase$(Guilds(UserList(UserIndex).GuildIndex).GuildName) Then
                If UserList(i).Pos.Map = UserList(UserIndex).Pos.Map Then
                    IDUser = IDUser + 1
                    Call SendData(ToIndex, UserIndex, 0, "PO" & UserList(i).Pos.X & "," & UserList(i).Pos.Y & "," & IDUser)
                    'Call SendData(ToIndex, i, 0, "PO" & UserList(UserIndex).pos.x & "," & UserList(UserIndex).pos.Y & "," & IDUser)

                End If

            End If

        End If
        End If

    Next i
    
    End If

End Sub

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading

'*************************************************
'Author: ZaMa
'Last modified: 30/03/2009
'Returns the heading opposite to the one passed by val.
'*************************************************
    Select Case nHeading

    Case eHeading.EAST
        InvertHeading = WEST

    Case eHeading.WEST
        InvertHeading = EAST

    Case eHeading.SOUTH
        InvertHeading = NORTH

    Case eHeading.NORTH
        InvertHeading = SOUTH

    End Select

End Function

Sub AutoCuraUser(ByVal UserIndex As Integer)

    If UserList(UserIndex).flags.Muerto = 1 Then
        Call RevivirUsuario(UserIndex)
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El sacerdote te ha resucitado y curado." & FONTTYPE_INFO)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & 9 & "," & 1)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW106")
        Call SendUserStatsBox(UserIndex)

    End If

    If UserList(UserIndex).Stats.MinHP < UserList(UserIndex).Stats.MaxHP Then
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El sacerdote te ha curado." & FONTTYPE_INFO)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & 9 & "," & 1)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW106")
        Call SendUserStatsBox(UserIndex)

    End If

    If UserList(UserIndex).flags.Envenenado = 1 Then UserList(UserIndex).flags.Envenenado = 0

End Sub

Sub ChangeUserInv(UserIndex As Integer, Slot As Byte, Object As UserOBJ)

    If Object.ObjIndex > 0 Then

        ' cambiamos precio divido en 2 si es cheke de oro
        If ObjData(Object.ObjIndex).Name = "Cheque por valor de 10k" Then
            PrecioQl = 1
        Else
            If ObjData(Object.ObjIndex).ObjType = eOBJType.otPLATA Then
                PrecioQl = 2
            Else
                PrecioQl = 3
            End If
        End If
        
          If ObjData(Object.ObjIndex).ObjType = eOBJType.otWeapon Then
              If ObjData(Object.ObjIndex).sagrado = 1 Then
                 Dim Ver As String
                 
                 Ver = ViewSagradaHit(UserIndex, Object.ObjIndex)
                 
                 Call SendData(SendTarget.ToIndex, UserIndex, 0, "CSI" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & _
                                                        Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," & ObjData(Object.ObjIndex).ObjType & "," & _
                                                        Ver & "," & ObjData(Object.ObjIndex).MaxDef & "," & ObjData( _
                                                        Object.ObjIndex).MinDef & "," & ObjData(Object.ObjIndex).Valor \ PrecioQl)
              
              ElseIf ObjData(Object.ObjIndex).sagrado = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "CSI" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & _
                                                        Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," & ObjData(Object.ObjIndex).ObjType & "," & _
                                                        ObjData(Object.ObjIndex).MaxHit & "," & ObjData(Object.ObjIndex).MinHit & "," & ObjData(Object.ObjIndex).MaxDef & "," & ObjData( _
                                                        Object.ObjIndex).MinDef & "," & ObjData(Object.ObjIndex).Valor \ PrecioQl)
              End If
          
          'ElseIf ObjData(Object.ObjIndex).ObjType = eOBJType.otUseOnce Then
          
          '  If UserList(UserIndex).Sagrada.Enabled = 1 Then
          '       Call SendData(SendTarget.ToIndex, UserIndex, 0, "CSI" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & _
                                                        Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," & ObjData(Object.ObjIndex).ObjType & "," & _
                                                        UserList(UserIndex).Sagrada.MaxHit & "," & UserList(UserIndex).Sagrada.MinHit & "," & ObjData(Object.ObjIndex).MaxDef & "," & ObjData( _
                                                        Object.ObjIndex).MinDef & "," & ObjData(Object.ObjIndex).Valor \ PrecioQl)
              
          '    ElseIf UserList(UserIndex).Sagrada.Enabled = 0 Then
          '        Call SendData(SendTarget.ToIndex, UserIndex, 0, "CSI" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & _
                                                        Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," & ObjData(Object.ObjIndex).ObjType & "," & _
                                                        ObjData(Object.ObjIndex).MaxHit & "," & ObjData(Object.ObjIndex).MinHit & "," & ObjData(Object.ObjIndex).MaxDef & "," & ObjData( _
                                                        Object.ObjIndex).MinDef & "," & ObjData(Object.ObjIndex).Valor \ PrecioQl)
              'End If
          
          Else
          
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "CSI" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & _
                                                        Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," & ObjData(Object.ObjIndex).ObjType & "," & _
                                                        ObjData(Object.ObjIndex).MaxHit & "," & ObjData(Object.ObjIndex).MinHit & "," & ObjData(Object.ObjIndex).MaxDef & "," & ObjData( _
                                                        Object.ObjIndex).MinDef & "," & ObjData(Object.ObjIndex).Valor \ PrecioQl)
          End If

    Else
        'Call SendData(SendTarget.ToIndex, UserIndex, 0, "CSI" & Slot & ",0," & "(Vacío)" & ",0,0,0")
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "CSI" & Slot & "," & "0" & "," & "(None)" & "," & "0" & "," & "0")

    End If

End Sub

Function NextOpenCharIndex() As Integer

Dim LoopC As Integer

For LoopC = 1 To LastChar + 1
    If CharList(LoopC) = 0 Then
        NextOpenCharIndex = LoopC
        NumChars = NumChars + 1
        If LoopC > LastChar Then LastChar = LoopC
        Exit Function
    End If
Next LoopC

End Function

Function NextOpenUser() As Integer
    Dim LoopC As Long
    
    For LoopC = 1 To MaxUsers + 1
        If LoopC > MaxUsers Then Exit For
        If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
    Next LoopC
    
    NextOpenUser = LoopC
End Function

Sub SendUserHitBox(ByVal UserIndex As Integer)
    Dim lagaminarma As Integer
    Dim lagamaxarma As Integer

    Dim lagaminarmor As Integer
    Dim lagamaxarmor As Integer

    Dim lagaminescu As Integer
    Dim lagamaxescu As Integer

    Dim lagamincasc As Integer
    Dim lagamaxcasc As Integer

    Dim Index As Integer

    lagaminarma = 0
    lagamaxarma = 0

    If UserList(UserIndex).Invent.WeaponEqpSlot > 0 Then
        Index = UserList(UserIndex).Invent.WeaponEqpObjIndex

        If Index > 0 Then
            lagaminarma = ObjData(Index).MinHit
            lagamaxarma = ObjData(Index).MaxHit

        End If

    End If

    lagaminarmor = 0
    lagamaxarmor = 0

    If UserList(UserIndex).Invent.ArmourEqpSlot > 0 Then
        Index = UserList(UserIndex).Invent.ArmourEqpObjIndex

        If Index > 0 Then
            lagaminarmor = ObjData(Index).MinDef
            lagamaxarmor = ObjData(Index).MaxDef

        End If

    End If

    lagaminescu = 0
    lagamaxescu = 0

    If UserList(UserIndex).Invent.EscudoEqpSlot > 0 Then
        Index = UserList(UserIndex).Invent.EscudoEqpObjIndex

        If Index > 0 Then
            lagaminescu = ObjData(Index).MinDef
            lagamaxescu = ObjData(Index).MaxDef

        End If

    End If

    lagamincasc = 0
    lagamaxcasc = 0

    If UserList(UserIndex).Invent.CascoEqpSlot > 0 Then
        Index = UserList(UserIndex).Invent.CascoEqpObjIndex

        If Index > 0 Then
            lagamincasc = ObjData(Index).MinDef
            lagamaxcasc = ObjData(Index).MaxDef

        End If

    End If

    'CRAW; 03/04/2020 --> QUITAMOS PORQUE NO HACE FALTA
    'Call SendData(SendTarget.ToIndex, UserIndex, 0, "ARM" & lagaminarma & "," & lagamaxarma & "," & lagaminarmor & "," & lagamaxarmor & "," & _
     lagaminescu & "," & lagamaxescu & "," & lagamincasc & "," & lagamaxcasc)

End Sub

Sub EnviarVerdes(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "VTG" & .flags.DuracionEfectoVerdes)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "VRG" & .Stats.UserAtributos(eAtributos.Fuerza))

    End With

End Sub

Sub EnviarAmarillas(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ATG" & .flags.DuracionEfectoAmarillas)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ARG" & .Stats.UserAtributos(eAtributos.Agilidad))

    End With

End Sub

Sub SendUserStatsBox(ByVal UserIndex As Integer)

    Call CompruebaOroRank(UserIndex)


    Call SendData(SendTarget.ToIndex, UserIndex, 0, "EST" & UserList(UserIndex).Stats.MaxHP & "," & UserList(UserIndex).Stats.MinHP & "," & _
                                                    UserList(UserIndex).Stats.MaxMAN & "," & UserList(UserIndex).Stats.MinMAN & "," & UserList(UserIndex).Stats.MaxSta & "," & UserList( _
                                                    UserIndex).Stats.MinSta & "," & UserList(UserIndex).Stats.GLD & "," & UserList(UserIndex).Stats.ELV & "," & UserList( _
                                                    UserIndex).Stats.ELU & "," & UserList(UserIndex).Stats.Exp & "," & UserList(UserIndex).AoMCreditos & "," & UserList(UserIndex).AoMCanjes)

End Sub

Sub EnviarHP(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        If .Stats.MinHP < 0 Then .Stats.MinHP = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "VID" & .Stats.MinHP)


        If .PartyIndex <> 0 Then

            If .Stats.MinHP <> 0 And .Stats.MaxHP <> 0 Then

                Call SendData(SendTarget.ToPartyArea, UserIndex, .Pos.Map, "VPT" & .char.CharIndex & "," & .Stats.MinHP & "," & .Stats.MaxHP & "," & .PartyIndex)

            End If

        End If

    End With

End Sub

Sub EnviarMn(ByVal UserIndex As Integer)

    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "MN" & UserList(UserIndex).Stats.MinMAN)

End Sub

Sub EnviarSta(ByVal UserIndex As Integer)

    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "STA" & UserList(UserIndex).Stats.MinSta)

End Sub

Sub EnviarOro(ByVal UserIndex As Integer)

    If UserList(UserIndex).Stats.GLD < 0 Then UserList(UserIndex).Stats.GLD = 0
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ORO" & UserList(UserIndex).Stats.GLD)

    Call CompruebaOroRank(UserIndex)
End Sub

Sub EnviarExp(ByVal UserIndex As Integer)

    If UserList(UserIndex).Stats.Exp < 0 Then UserList(UserIndex).Stats.Exp = 0
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "EXP" & UserList(UserIndex).Stats.Exp)

End Sub

Sub EnviarHambreYsed(ByVal UserIndex As Integer)

    If UserList(UserIndex).Stats.MinAGU < 0 Then UserList(UserIndex).Stats.MinAGU = 0
    If UserList(UserIndex).Stats.MinHam < 0 Then UserList(UserIndex).Stats.MinHam = 0

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "EHYS" & UserList(UserIndex).Stats.MinAGU & "," & UserList(UserIndex).Stats.MinHam)

End Sub

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
    Dim GuildI As Integer

    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Estadisticas de: " & UserList(UserIndex).Name & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.Exp & "/" & _
                                                    UserList(UserIndex).Stats.ELU & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Salud: " & UserList(UserIndex).Stats.MinHP & "/" & UserList(UserIndex).Stats.MaxHP & _
                                                  "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList( _
                                                    UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta & FONTTYPE_INFO)

    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHit & "/" & UserList( _
                                                        UserIndex).Stats.MaxHit & " (" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHit & "/" & ObjData(UserList( _
                                                                                                                                                                      UserIndex).Invent.WeaponEqpObjIndex).MaxHit & ")" & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHit & "/" & UserList( _
                                                        UserIndex).Stats.MaxHit & FONTTYPE_INFO)

    End If

    If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: " & ObjData(UserList( _
                                                                                                 UserIndex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: 0" & FONTTYPE_INFO)

    End If

    If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: " & ObjData(UserList( _
                                                                                                 UserIndex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: 0" & FONTTYPE_INFO)

    End If

    GuildI = UserList(UserIndex).GuildIndex

    If GuildI > 0 Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Clan: " & Guilds(GuildI).GuildName & FONTTYPE_INFO)

        If UCase$(Guilds(GuildI).GetLeader) = UCase$(UserList(sendIndex).Name) Then
            Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Status: Lider" & FONTTYPE_INFO)

        End If

        'guildpts no tienen objeto
        'Call SendData(SendTarget.ToIndex, sendIndex, 0, "||User GuildPoints: " & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
    End If

    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Oro: " & UserList(UserIndex).Stats.GLD & "  Posicion: " & UserList(UserIndex).Pos.X & "," & _
                                                    UserList(UserIndex).Pos.Y & " en mapa " & UserList(UserIndex).Pos.Map & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Dados: " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) & ", " & UserList( _
                                                    UserIndex).Stats.UserAtributos(eAtributos.Agilidad) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) & ", " & _
                                                    UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) & _
                                                    FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Trofeos de Oro: " & UserList(UserIndex).Stats.TrofOro & "~255~255~6~0~0~")
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Trofeos de Plata: " & UserList(UserIndex).Stats.TrofPlata & "~255~255~251~0~0~")
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Trofeos de Bronce: " & UserList(UserIndex).Stats.TrofBronce & "~187~0~0~0~0~")
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Amuletos de Madera: " & UserList(UserIndex).Stats.TrofMadera & "~237~207~139~0~0~")

End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

    With UserList(UserIndex)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Pj: " & .Name & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||CiudadanosMatados: " & .Faccion.CiudadanosMatados & " CriminalesMatados: " & _
                                                        .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||NPCsMuertos: " & .Stats.NPCsMuertos & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Clase: " & .Clase & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Pena: " & .Counters.Pena & FONTTYPE_INFO)

    End With

End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
    Dim CharFile As String
    Dim Ban As String
    Dim BanDetailPath As String

    BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & CharName & ".chr"

    If FileExist(CharFile) Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Pj: " & CharName & FONTTYPE_INFO)
        ' 3 en uno :p
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||CiudadanosMatados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & _
                                                      " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " UsuariosMatados: " & GetVar(CharFile, "MUERTES", _
                                                                                                                                                            "UserMuertes") & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||NPCsMuertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes") & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Clase: " & GetVar(CharFile, "INIT", "Clase") & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Pena: " & GetVar(CharFile, "COUNTERS", "PENA") & FONTTYPE_INFO)
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Ban: " & Ban & FONTTYPE_INFO)

        If Ban = "1" Then
            Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Baneado por: " & GetVar(CharFile, CharName, "BannedBy") & " El Motivo Fue: " & _
                                                            GetVar(BanDetailPath, CharName, "Reason") & FONTTYPE_INFO)

        End If

    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||El pj no existe: " & CharName & FONTTYPE_INFO)

    End If

End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim j As Long

    Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| " & UserList(UserIndex).Name & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| Tiene " & UserList(UserIndex).Invent.NroItems & " objetos." & FONTTYPE_INFO)

    For j = 1 To MAX_INVENTORY_SLOTS

        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
            Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| Objeto " & j & "  (Indice " & UserList(UserIndex).Invent.Object(j).ObjIndex & ") " & ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Name & _
                                                          " Cantidad: " & UserList(UserIndex).Invent.Object(j).Amount & FONTTYPE_INFO)

        End If

    Next j

End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)

    On Error Resume Next

    Dim j As Long
    Dim CharFile As String, Tmp As String
    Dim ObjInd As Long, ObjCant As Long

    CharFile = CharPath & CharName & ".chr"

    If FileExist(CharFile, vbNormal) Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||" & CharName & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos." & FONTTYPE_INFO)

        For j = 1 To MAX_INVENTORY_SLOTS
            Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))

            If ObjInd > 0 Then
                Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant & _
                                                                FONTTYPE_INFO)

            End If

        Next j

    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Usuario inexistente: " & CharName & FONTTYPE_INFO)

    End If

End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim j As Integer
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||" & UserList(UserIndex).Name & FONTTYPE_INFO)

    For j = 1 To NUMSKILLS
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| " & SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j) & FONTTYPE_INFO)
    Next
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| SkillLibres:" & UserList(UserIndex).Stats.SkillPts & FONTTYPE_INFO)

End Sub

Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

    If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not Criminal(Npclist(NpcIndex).MaestroUser)

        If EsMascotaCiudadano Then Call SendData(SendTarget.ToIndex, Npclist(NpcIndex).MaestroUser, 0, "||¡¡" & UserList(UserIndex).Name & _
                                                                                                     " esta atacando tu mascota!!" & FONTTYPE_FIGHT)

    End If

End Function

Sub NpcAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

'Guardamos el usuario que ataco el npc
    Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).Name

    If Npclist(NpcIndex).MaestroUser > 0 Then Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)

    If EsMascotaCiudadano(NpcIndex, UserIndex) Then
            Call VolverCriminal(UserIndex)
            ' Npclist(NpcIndex).flags.AttackedBy
            Npclist(NpcIndex).Movement = NPCDEFENSA
            Npclist(NpcIndex).Hostile = 1
Else
    'Reputacion
    If Npclist(NpcIndex).Stats.Alineacion = 0 Then
       If Npclist(NpcIndex).NPCtype = 1 Then
                Call VolverCriminal(UserIndex)
       Else
               Call AddtoVar(UserList(UserIndex).Reputacion.BandidoRep, vlASALTO, MAXREP)
       End If
    ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
             Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlCAZADOR / 2, MAXREP)
    End If

        'hacemos que el npc se defienda
        Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
        Npclist(NpcIndex).Hostile = 1

    End If

End Sub

Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean

    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        PuedeApuñalar = ((UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR) And (ObjData(UserList( _
                                                                                                              UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1)) Or ((UCase$(UserList(UserIndex).Clase) = "ASESINO") And (ObjData(UserList( _
                                                                                                                                                                                                                                  UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1))
    Else
        PuedeApuñalar = False

    End If

End Function

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)

    Dim Aumenta As Integer

    Dim Prob    As Integer
    
    If UserList(UserIndex).Stats.ELV <= 3 Then
        Prob = 15
    ElseIf UserList(UserIndex).Stats.ELV > 3 And UserList(UserIndex).Stats.ELV < 6 Then
        Prob = 15
    ElseIf UserList(UserIndex).Stats.ELV >= 6 And UserList(UserIndex).Stats.ELV < 10 Then
        Prob = 20
    ElseIf UserList(UserIndex).Stats.ELV >= 10 And UserList(UserIndex).Stats.ELV < 20 Then
        Prob = 20
    Else
        Prob = 25

    End If
    
    Aumenta = Int(RandomNumber(1, Prob))
    
    Dim lvl As Integer

    lvl = UserList(UserIndex).Stats.ELV
    
    If lvl >= UBound(LevelSkill) Then Exit Sub
    If UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
    
    If Aumenta = 7 And UserList(UserIndex).Stats.UserSkills(Skill) < LevelSkill(lvl).LevelValue Then
        Call AddtoVar(UserList(UserIndex).Stats.UserSkills(Skill), 1, MAXSKILLPOINTS)
        Call SendData(ToIndex, UserIndex, 0, "||¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(UserIndex).Stats.UserSkills(Skill) & " pts." & FONTTYPE_INFO)

        If lvl < 15 Then
            Call AddtoVar(UserList(UserIndex).Stats.Exp, lvl * 10, MAXEXP)
            Call SendData(ToIndex, UserIndex, 0, "||¡Has ganado " & lvl * 10 & " puntos de experiencia!" & FONTTYPE_FIGHT)
        Else
            Call AddtoVar(UserList(UserIndex).Stats.Exp, 50, MAXEXP)
            Call SendData(ToIndex, UserIndex, 0, "||¡Has ganado 50 puntos de experiencia!" & FONTTYPE_FIGHT)

        End If

        Call EnviarSkills(UserIndex)
        Call CheckUserLevel(UserIndex)

    End If

End Sub

Sub UserDie(ByVal UserIndex As Integer)

    On Error GoTo ErrorHandler
    
    Dim i As Byte
    
    If UserList(UserIndex).flags.MetamorfosisLicantropo = 1 Then Call DoLicantropo(UserIndex)

    If UserList(UserIndex).GranPoder = 1 Then
        Call mod_GranPoder.MuerePoder(UserIndex)
    End If

    'Sonido
    If UCase$(UserList(UserIndex).Genero) = "MUJER" Then
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, e_SoundIndex.MUERTE_MUJER)
    Else
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, e_SoundIndex.MUERTE_HOMBRE)
    End If

    'Quitar el dialogo del user muerto
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "QDL" & UserList(UserIndex).char.CharIndex)
    
    'enviar efecto de sangre
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "SFX" & UserList(UserIndex).char.CharIndex & "-0")
    
    For i = 1 To MAXMASCOTAS
    
         If UserList(UserIndex).MascotasIndex(i) > 0 Then
              Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0
              Call MuereNpc(UserList(UserIndex).MascotasIndex(i), UserIndex)
              
         End If
    
    Next i
    
    UserList(UserIndex).Stats.MinHP = 0
    UserList(UserIndex).Stats.MinSta = 0
    UserList(UserIndex).flags.AtacadoPorNpc = 0
    UserList(UserIndex).flags.AtacadoPorUser = 0
    UserList(UserIndex).flags.Envenenado = 0
    UserList(UserIndex).flags.Muerto = 1

    '  2vs2
    If HayPareja = True Then
        If UserList(Pareja.Jugador1).flags.EnPareja = True And UserList(Pareja.Jugador2).flags.EnPareja = True And UserList(Pareja.Jugador1).flags.Muerto = 1 And UserList(Pareja.Jugador2).flags.Muerto = 1 Then
            Call WarpUserChar(Pareja.Jugador1, 34, 30, 50)
            Call WarpUserChar(Pareja.Jugador2, 34, 30, 51)
            Call WarpUserChar(Pareja.Jugador3, 34, 30, 52)
            Call WarpUserChar(Pareja.Jugador4, 34, 30, 53)
            UserList(Pareja.Jugador1).flags.EnPareja = False
            UserList(Pareja.Jugador1).flags.EsperaPareja = False
            UserList(Pareja.Jugador1).flags.SuPareja = 0
            UserList(Pareja.Jugador2).flags.EnPareja = False
            UserList(Pareja.Jugador2).flags.EsperaPareja = False
            UserList(Pareja.Jugador2).flags.SuPareja = 0
            UserList(Pareja.Jugador3).flags.EnPareja = False
            UserList(Pareja.Jugador3).flags.EsperaPareja = False
            UserList(Pareja.Jugador3).flags.SuPareja = 0
            UserList(Pareja.Jugador4).flags.EnPareja = False
            UserList(Pareja.Jugador4).flags.EsperaPareja = False
            UserList(Pareja.Jugador4).flags.SuPareja = 0
            HayPareja = False
            Call SendData(SendTarget.ToAll, 0, 0, "||2 vs 2 > " & UserList(Pareja.Jugador1).Name & " y " & UserList(Pareja.Jugador2).Name & " han sido derrotados" & FONTTYPE_GUILD)
        End If

        If UserList(Pareja.Jugador3).flags.EnPareja = True And UserList(Pareja.Jugador4).flags.EnPareja = True And UserList(Pareja.Jugador3).flags.Muerto = 1 And UserList(Pareja.Jugador4).flags.Muerto = 1 Then
            Call WarpUserChar(Pareja.Jugador1, 34, 30, 50)
            Call WarpUserChar(Pareja.Jugador2, 34, 30, 51)
            Call WarpUserChar(Pareja.Jugador3, 34, 30, 52)
            Call WarpUserChar(Pareja.Jugador4, 34, 30, 53)
            UserList(Pareja.Jugador1).flags.EnPareja = False
            UserList(Pareja.Jugador1).flags.EsperaPareja = False
            UserList(Pareja.Jugador1).flags.SuPareja = 0
            UserList(Pareja.Jugador2).flags.EnPareja = False
            UserList(Pareja.Jugador2).flags.EsperaPareja = False
            UserList(Pareja.Jugador2).flags.SuPareja = 0
            UserList(Pareja.Jugador3).flags.EnPareja = False
            UserList(Pareja.Jugador3).flags.EsperaPareja = False
            UserList(Pareja.Jugador3).flags.SuPareja = 0
            UserList(Pareja.Jugador4).flags.EnPareja = False
            UserList(Pareja.Jugador4).flags.EsperaPareja = False
            UserList(Pareja.Jugador4).flags.SuPareja = 0
            HayPareja = False
            Call SendData(SendTarget.ToAll, 0, 0, "||2 vs 2 > " & UserList(Pareja.Jugador3).Name & " y " & UserList(Pareja.Jugador4).Name & " han sido derrotados" & FONTTYPE_GUILD)
        End If
    End If

    Dim aN As Integer

    aN = UserList(UserIndex).flags.AtacadoPorNpc

    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = ""

    End If

    '<<<< Verdes >>>>
    If UserList(UserIndex).flags.DuracionEfectoVerdes > 0 Then
        UserList(UserIndex).flags.DuracionEfectoVerdes = 0
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Fuerza)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "VTG" & UserList(UserIndex).flags.DuracionEfectoVerdes)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "VRG" & UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
    End If

    '<<<< Amarillas >>>>
    If UserList(UserIndex).flags.DuracionEfectoAmarillas > 0 Then
        UserList(UserIndex).flags.DuracionEfectoAmarillas = 0
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Agilidad)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ATG" & UserList(UserIndex).flags.DuracionEfectoAmarillas)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ARG" & UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
    End If

    '<<<< Paralisis >>>>
    If UserList(UserIndex).flags.Paralizado = 1 Then
        UserList(UserIndex).flags.Paralizado = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "PARADOW")

    End If

    '<<< Estupidez >>>
    If UserList(UserIndex).flags.Estupidez = 1 Then
        UserList(UserIndex).flags.Estupidez = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "NESTUP")

    End If

    '<<<< Descansando >>>>
    If UserList(UserIndex).flags.Descansar Then
        UserList(UserIndex).flags.Descansar = False
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")

    End If

    '<<<< Meditando >>>>
    If UserList(UserIndex).flags.Meditando Then
        UserList(UserIndex).flags.Meditando = False
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MEDOK")

    End If

    '<<<< Invisible >>>>
    If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).Counters.Ocultando = 0
        UserList(UserIndex).flags.Invisible = 0

        'no hace falta encriptar este NOVER
        Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).char.CharIndex & ",0," & UserList( _
                                                                        UserIndex).PartyIndex)

        If UserList(UserIndex).PartyIndex <> 0 Then

            If UserList(UserIndex).Stats.MinHP <> 0 And UserList(UserIndex).Stats.MaxHP <> 0 Then

                Call SendData(SendTarget.ToPartyArea, UserIndex, UserList(UserIndex).Pos.Map, "VPT" & UserList(UserIndex).char.CharIndex & "," & UserList(UserIndex).Stats.MinHP & "," & UserList(UserIndex).Stats.MaxHP & "," & UserList(UserIndex).PartyIndex)

            End If

        End If

    End If

    If TriggerZonaPelea(UserIndex, UserIndex) <> TRIGGER6_PERMITE Then

        ' << Si es newbie no pierde el inventario >>
        If Not EsNewbie(UserIndex) Or Criminal(UserIndex) Then
            Call TirarTodo(UserIndex)
        Else

            'If EsNewbie(UserIndex) Then Call TirarTodosLosItemsNoNewbies(UserIndex)

        End If

    End If

    ' DESEQUIPA TODOS LOS OBJETOS
    'desequipar armadura
    If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)

    End If

    If UserList(UserIndex).Invent.AlaEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.AlaEqpSlot)

    End If

    'desequipar arma
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)

    End If

    'desequipar casco
    If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)

    End If

    'desequipar herramienta
    If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)

    End If

    'desequipar municiones
    If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)

    End If

    'desequipar escudo
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)

    End If
    If UserList(UserIndex).EnCvc Then
        'Dim ijaji As Integer
        'For ijaji = 1 To LastUser
        With UserList(UserIndex)
            If Guilds(.GuildIndex).GuildName = Nombre1 Then
                If .EnCvc = True Then
                    If .flags.Muerto Then
                        modGuilds.UsuariosEnCvcClan1 = modGuilds.UsuariosEnCvcClan1 - 1
                        If modGuilds.UsuariosEnCvcClan1 = 0 Then
                            Call SendData(SendTarget.ToAll, UserIndex, 0, "||" & "El clan " & Nombre2 & " derrotó al clan " & Nombre1 & "." & FONTTYPE_GUILD)
                            CvcFunciona = False
                            Call LlevarUsuarios
                        End If
                    End If
                End If
            End If


            If Guilds(.GuildIndex).GuildName = Nombre2 Then
                If .EnCvc = True Then
                    If .flags.Muerto Then
                        modGuilds.UsuariosEnCvcClan2 = modGuilds.UsuariosEnCvcClan2 - 1
                        If modGuilds.UsuariosEnCvcClan2 = 0 Then
                            Call SendData(SendTarget.ToAll, UserIndex, 0, "||" & "El clan " & Nombre1 & " derrotó al clan " & Nombre2 & "." & FONTTYPE_GUILD)
                            CvcFunciona = False
                            Call LlevarUsuarios
                        End If
                    End If
                End If
            End If
        End With
        'Next ijaji
    End If

    ' << Reseteamos los posibles FX sobre el personaje >>
    If UserList(UserIndex).char.loops = LoopAdEternum Then
        UserList(UserIndex).char.FX = 0
        UserList(UserIndex).char.loops = 0

    End If

    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    If UserList(UserIndex).flags.automatico = True Then
        Call Rondas_UsuarioMuere(UserIndex)

    End If

    If UserList(UserIndex).flags.bandas = True Then
        Call Ban_Muere(UserIndex)

    End If

    If UserList(UserIndex).flags.Montado = True Then
        UserList(UserIndex).flags.NumeroMont = 0
        UserList(UserIndex).flags.Montado = False

    End If

    ' <<Si pierde el duelo se va>>
    If UserList(UserIndex).Pos.Map = MAPADUELO And UserIndex = duelosespera Then
        Call WarpUserChar(UserIndex, 34, 30, 50, True)
        Call SendData(SendTarget.ToAll, 0, 0, "||Duelos: el perdedor " & UserList(duelosespera).Name & " a salido de duelos." & FONTTYPE_TALK)
        duelosespera = duelosreta
        numduelos = 0

    End If

    If UserList(UserIndex).Pos.Map = MAPADUELO And UserIndex = duelosreta Then
        Call WarpUserChar(UserIndex, 34, 30, 50, True)
        numduelos = numduelos + 1
        UserList(duelosespera).Stats.PuntosDuelos = UserList(duelosespera).Stats.PuntosDuelos + 1
        Call SendData(SendTarget.ToAll, 0, 0, "||Duelos: el perdedor " & UserList(duelosreta).Name & " a salido de duelos." & FONTTYPE_TALK)

        If numduelos Mod 5 = 0 Then
            Call SendData(SendTarget.ToAll, 0, 0, "TW123")
            Call SendData(SendTarget.ToAll, 0, 0, "||Duelos: " & UserList(duelosespera).Name & " ha ganado " & numduelos & " consecutivos!" & _
                                                  FONTTYPE_TALK)

        End If

        Call SendData(SendTarget.ToAll, 0, 0, "||Duelos: " & UserList(duelosespera).Name & " ha ganado el duelo y espera otro rival." & FONTTYPE_TALK)

    End If

    ' << Restauramos el mimetismo
    If UserList(UserIndex).flags.Mimetizado = 1 Then
        UserList(UserIndex).char.Body = UserList(UserIndex).CharMimetizado.Body
        UserList(UserIndex).char.Head = UserList(UserIndex).CharMimetizado.Head
        UserList(UserIndex).char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim

        UserList(UserIndex).char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
        UserList(UserIndex).char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
        UserList(UserIndex).Counters.Mimetismo = 0
        UserList(UserIndex).flags.Mimetizado = 0

    End If

    '<< Cambiamos la apariencia del char >>
    If UserList(UserIndex).flags.Navegando = 0 Then
        UserList(UserIndex).char.Body = iCuerpoMuerto
        UserList(UserIndex).char.Head = iCabezaMuerto
        UserList(UserIndex).char.ShieldAnim = NingunEscudo
        UserList(UserIndex).char.WeaponAnim = NingunArma
        UserList(UserIndex).char.CascoAnim = NingunCasco

    Else
        UserList(UserIndex).char.Body = iFragataFantasmal    ';)

    End If

    UserList(UserIndex).NroMacotas = 0

    'If Criminal(UserIndex) Then
    '   Call SendData(SendTarget.toIndex, UserIndex, 0, "Z33")
    ' Else
    '     Call SendData(SendTarget.toIndex, UserIndex, 0, "Z34")

    'End If

    'If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
    '        Dim MiObj As Obj
    '        Dim nPos As WorldPos
    '        MiObj.ObjIndex = RandomNumber(554, 555)
    '        MiObj.Amount = 1
    '        nPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    '        Dim ManchaSangre As New cGarbage
    '        ManchaSangre.Map = nPos.Map
    '        ManchaSangre.X = nPos.X
    '        ManchaSangre.Y = nPos.Y
    '        Call TrashCollector.Add(ManchaSangre)
    'End If

    '<< Actualizamos clientes >>
    '[MaTeO 9]
    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, val(UserIndex), UserList(UserIndex).char.Body, UserList( _
                                                                                                                         UserIndex).char.Head, UserList(UserIndex).char.heading, NingunArma, NingunEscudo, NingunCasco, NingunAlas)
    '[/MaTeO 9]

    Call SendUserStatsBox(UserIndex)
    Call SendUserHitBox(UserIndex)
    Call EnviarAmarillas(UserIndex)
    Call EnviarVerdes(UserIndex)

    '<<Castigos por party>>
    'If UserList(UserIndex).PartyIndex > 0 Then
    '    Call mdParty.ObtenerExito(UserIndex, UserList(UserIndex).Stats.ELV * -10 * mdParty.CantMiembros(UserIndex), UserList(UserIndex).pos.Map, _
         '            UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y)
    '
    '    End If

    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & MuereSpell & "," _
                                                                             & LoopSpell)

    'Reset Spell

    MuereSpell = 0
    LoopSpell = 0


    If UserList(UserIndex).flags.EnDosVDos = True Then
        Call VerificarDosVDos(UserIndex)

    End If

    If UserList(UserIndex).flags.EstaDueleando = True Then
        Call TerminarDuelo(UserList(UserIndex).flags.Oponente, UserIndex)

    End If

    If UserList(UserIndex).Asedio.Participando Then
        Call modAsedio.MuereUser(UserIndex)
    End If

    Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & err.Number & " Descripción: " & err.Description)

End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)

    On Error GoTo ErrorHandler

    If EsNewbie(Muerto) Then Exit Sub

    If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub

    If Criminal(Muerto) Then
        If UserList(Atacante).flags.LastCrimMatado <> UserList(Muerto).Name Then
            UserList(Atacante).flags.LastCrimMatado = UserList(Muerto).Name

            If UserList(Atacante).Faccion.CriminalesMatados < 65000 Then UserList(Atacante).Faccion.CriminalesMatados = UserList( _
               Atacante).Faccion.CriminalesMatados + 1

        End If

        If UserList(Atacante).Faccion.CriminalesMatados > MAXUSERMATADOS Then
            UserList(Atacante).Faccion.CriminalesMatados = 0
            UserList(Atacante).Faccion.RecompensasReal = 0

        End If

        If UserList(Atacante).Faccion.RecibioExpInicialCaos = 1 And UserList(Muerto).Faccion.FuerzasCaos = 1 Then
            UserList(Atacante).Faccion.Reenlistadas = 200  'jaja que trucho

            'con esto evitamos que se vuelva a reenlistar
        End If

    Else

        If UserList(Atacante).flags.LastCiudMatado <> UserList(Muerto).Name Then
            UserList(Atacante).flags.LastCiudMatado = UserList(Muerto).Name

            If UserList(Atacante).Faccion.CiudadanosMatados < 65000 Then UserList(Atacante).Faccion.CiudadanosMatados = UserList( _
               Atacante).Faccion.CiudadanosMatados + 1

        End If

        If UserList(Atacante).Faccion.CiudadanosMatados > MAXUSERMATADOS Then
            UserList(Atacante).Faccion.CiudadanosMatados = 0
            UserList(Atacante).Faccion.RecompensasCaos = 0

        End If

    End If

ErrorHandler:
    '  Call LogError("Error en SUB CONTARMUERTE. Error: " & Err.Number & " Descripción: " & Err.Description)

End Sub

Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos)
    'Call LogTarea("Sub Tilelibre")

    Dim Notfound As Boolean

    Dim LoopC    As Integer

    Dim Tx       As Integer

    Dim Ty       As Integer

    Dim hayobj   As Boolean

    hayobj = False
    nPos.Map = Pos.Map

    Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y) Or hayobj
    
        If LoopC > 15 Then
            Notfound = True
            Exit Do

        End If
    
        For Ty = Pos.Y - LoopC To Pos.Y + LoopC
            For Tx = Pos.X - LoopC To Pos.X + LoopC
        
                If LegalPos(nPos.Map, Tx, Ty) = True Then
                    hayobj = (MapData(nPos.Map, Tx, Ty).OBJInfo.ObjIndex > 0)

                    If Not hayobj And MapData(nPos.Map, Tx, Ty).TileExit.Map = 0 Then
                        nPos.X = Tx
                        nPos.Y = Ty
                        Tx = Pos.X + LoopC
                        Ty = Pos.Y + LoopC

                    End If

                End If
        
            Next Tx
        Next Ty
    
        LoopC = LoopC + 1
    
    Loop

    If Notfound = True Then
        nPos.X = 0
        nPos.Y = 0

    End If

End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)

    Dim OldMap As Integer
    Dim OldX As Integer
    Dim OldY As Integer
    Dim NP As Byte

    With UserList(UserIndex)

        If .Pos.Map = MAPADUELO Then
            If MapInfo(MAPADUELO).NumUsers > 0 Then
                If .flags.Privilegios = PlayerType.Dios Or .flags.Privilegios = PlayerType.SemiDios Or .flags.Privilegios = PlayerType.Consejero Then

                Else

                    If .flags.Muerto = 1 Then

                    Else
                        Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos: " & .Name & " ha salido de la sala de torneos." & FONTTYPE_TALK)

                    End If

                End If

            End If

        End If

        'Quitar el dialogo
        Call SendToUserArea(UserIndex, "QDL" & .char.CharIndex)
        Call SendData(SendTarget.ToIndex, UserIndex, .Pos.Map, "QTDL")

        OldMap = .Pos.Map
        OldX = .Pos.X
        OldY = .Pos.Y

        Call EraseUserChar(SendTarget.ToMap, 0, OldMap, UserIndex)

        If OldMap <> Map Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "CM" & Map & "," & MapInfo(.Pos.Map).MapVersion)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "TM" & MapInfo(Map).Music)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "N~" & MapInfo(Map).Name)

            'Update new Map Users
            MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1

            'Update old Map Users
            MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
            
            Call Corr_ActualizarPosicion(UserIndex, X, Y)

            If MapInfo(OldMap).NumUsers < 0 Then
                MapInfo(OldMap).NumUsers = 0

            End If

        End If

        .Pos.X = X
        .Pos.Y = Y
        .Pos.Map = Map

        'Anti Pisadas
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex <> 0 Then
            Dim nPos As WorldPos
            Call ClosestStablePos(.Pos, nPos)

            If nPos.X <> 0 And nPos.Y <> 0 Then
                .Pos.Map = nPos.Map
                .Pos.X = nPos.X
                .Pos.Y = nPos.Y

            End If

        End If

        'Anti Pisadas

        Call MakeUserChar(SendTarget.ToMap, 0, Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "IP" & .char.CharIndex)

        'Seguis invisible al pasar de mapa
        If (.flags.Invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then
            Call SendToUserArea(UserIndex, "NOVER" & .char.CharIndex & ",1," & .PartyIndex)
        End If
        
        If .PartyIndex <> 0 Then

                If .Stats.MinHP <> 0 And .Stats.MaxHP <> 0 Then
                
                    NP = UserList(UserIndex).PartyIndex
                          
                      Call Parties(NP).EnviarBarra

                End If

        End If

        If FX And .flags.AdminInvisible = 0 Then    'FX
            Call SendData(SendTarget.ToPCArea, UserIndex, .Pos.Map, "TW" & SND_WARP & "," & .char.CharIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, .Pos.Map, "CFX" & .char.CharIndex & "," & FXIDs.FXWARP & ",0")

        End If

        Call WarpMascotas(UserIndex)
        
        'Encontrar mapas.
        
        If .Quest.Start = 1 Then
            If .Quest.NumMap > 0 Then
                Call EncuentraMapaQuest(UserIndex, .Quest.Quest)
            End If
        End If
        
        If .GuildIndex > 0 Then
            Call SendData(ToGuildMembers, .GuildIndex, 0, "PL")
            Call EnviaPosClan(UserIndex)
        End If
        
        
    End With

End Sub

Sub UpdateUserMap(ByVal UserIndex As Integer)

    Dim Map As Integer
    Dim X As Integer
    Dim Y As Integer

    'EnviarNoche UserIndex

    On Error GoTo 0

    Map = UserList(UserIndex).Pos.Map
    Call SendData(ToIndex, UserIndex, 0, "SMN" & UCase$(MapInfo(UserList(UserIndex).Pos.Map).Name))

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(Map, X, Y).UserIndex > 0 And UserIndex <> MapData(Map, X, Y).UserIndex Then
                Call MakeUserChar(SendTarget.ToIndex, UserIndex, 0, MapData(Map, X, Y).UserIndex, Map, X, Y)

                If UserList(MapData(Map, X, Y).UserIndex).flags.Invisible = 1 Or UserList(MapData(Map, X, Y).UserIndex).flags.Oculto = 1 Then Call _
                   SendData(SendTarget.ToIndex, UserIndex, 0, "NOVER" & UserList(MapData(Map, X, Y).UserIndex).char.CharIndex & ",1," & _
                                                              UserList(MapData(Map, X, Y).UserIndex).PartyIndex)

            End If

            If MapData(Map, X, Y).NpcIndex > 0 Then
                Call MakeNPCChar(SendTarget.ToIndex, UserIndex, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)

            End If

            If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
                If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ObjType <> eOBJType.otArboles Then
                    Call MakeObj(SendTarget.ToIndex, UserIndex, 0, MapData(Map, X, Y).OBJInfo, Map, X, Y)

                    If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ObjType = eOBJType.otPuertas Then
                        Call Bloquear(SendTarget.ToIndex, UserIndex, 0, Map, X, Y, MapData(Map, X, Y).Blocked)
                        Call Bloquear(SendTarget.ToIndex, UserIndex, 0, Map, X - 1, Y, MapData(Map, X - 1, Y).Blocked)

                    End If

                End If

            End If

        Next X
    Next Y

End Sub

Sub WarpMascotas(ByVal UserIndex As Integer)
    Dim i As Integer

    Dim UMascRespawn As Boolean
    Dim miflag As Byte, MascotasReales As Integer
    Dim prevMacotaType As Integer

    Dim PetTypes(1 To MAXMASCOTAS) As Integer
    Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
    Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

    Dim NroPets As Integer, InvocadosMatados As Integer

    NroPets = UserList(UserIndex).NroMacotas
    InvocadosMatados = 0


    For i = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasIndex(i) > 0 Then
            PetRespawn(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.Respawn = 0
            PetTypes(i) = UserList(UserIndex).MascotasType(i)
            PetTiempoDeVida(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia
            Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))

        End If

    Next i

    For i = 1 To MAXMASCOTAS

        If PetTypes(i) > 0 Then
            UserList(UserIndex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(UserIndex).Pos, False, PetRespawn(i))
            UserList(UserIndex).MascotasType(i) = PetTypes(i)

            'Controlamos que se sumoneo OK
            If UserList(UserIndex).MascotasIndex(i) = 0 Then
                UserList(UserIndex).MascotasIndex(i) = 0
                UserList(UserIndex).MascotasType(i) = 0

                If UserList(UserIndex).NroMacotas > 0 Then UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
                Exit Sub

            End If

            Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
            Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = TipoAI.SigueAmo
            Npclist(UserList(UserIndex).MascotasIndex(i)).Target = 0
            Npclist(UserList(UserIndex).MascotasIndex(i)).TargetNpc = 0
            Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
            Call FollowAmo(UserList(UserIndex).MascotasIndex(i))

        End If

    Next i

    UserList(UserIndex).NroMacotas = NroPets

End Sub

Sub RepararMascotas(ByVal UserIndex As Integer)
    Dim i As Integer
    Dim MascotasReales As Integer

    For i = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasType(i) > 0 Then MascotasReales = MascotasReales + 1
    Next i

    If MascotasReales <> UserList(UserIndex).NroMacotas Then UserList(UserIndex).NroMacotas = 0

End Sub

Sub Cerrar_Usuario(ByVal UserIndex As Integer, Optional ByVal Tiempo As Integer = -1)

    If Tiempo = -1 Then Tiempo = IntervaloCerrarConexion

    If UserList(UserIndex).flags.Privilegios > User Then
        If UserList(UserIndex).flags.UserLogged And Not UserList(UserIndex).Counters.Saliendo Then
            UserList(UserIndex).Counters.Saliendo = True
            UserList(UserIndex).Counters.Salir = IIf(UserList(UserIndex).flags.Privilegios > PlayerType.User Or Not MapInfo(UserList( _
                                                                                                                            UserIndex).Pos.Map).Pk, 0, Tiempo)

        End If

    Else

        If UserList(UserIndex).flags.UserLogged And Not UserList(UserIndex).Counters.Saliendo Then
            UserList(UserIndex).Counters.Saliendo = True
            UserList(UserIndex).Counters.Salir = IIf(UserList(UserIndex).flags.Privilegios > PlayerType.User Or Not MapInfo(UserList( _
                                                                                                                            UserIndex).Pos.Map).Pk, IntervaloCerrarConexion, Tiempo)

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Cerrando...Se cerrará el juego en " & UserList(UserIndex).Counters.Salir & _
                                                          " segundos..." & FONTTYPE_INFO)

        End If

    End If
    If UserList(UserIndex).flags.EnCvc = True Then
        UserList(UserIndex).flags.EnCvc = False
        WarpUserChar UserIndex, 34, 30, 50, True
    End If

End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal UserIndex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)

    Dim ViejoNick As String
    Dim ViejoCharBackup As String

    If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
    ViejoNick = UserList(UserIndexDestino).Name

    If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
        'hace un backup del char
        ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
        Name CharPath & ViejoNick & ".chr" As ViejoCharBackup

    End If

End Sub

Public Sub Empollando(ByVal UserIndex As Integer)

    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.EstaEmpo = 1
    Else
        UserList(UserIndex).flags.EstaEmpo = 0
        UserList(UserIndex).EmpoCont = 0

    End If

End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal nombre As String)

    If FileExist(CharPath & nombre & ".chr", vbArchive) = False Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Pj Inexistente" & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Estadisticas de: " & nombre & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Nivel: " & GetVar(CharPath & nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar( _
                                                        CharPath & nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & nombre & ".chr", "stats", "elu") & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Vitalidad: " & GetVar(CharPath & nombre & ".chr", "stats", "minsta") & "/" & GetVar( _
                                                        CharPath & nombre & ".chr", "stats", "maxSta") & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Salud: " & GetVar(CharPath & nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath _
                                                                                                                                        & nombre & ".chr", "Stats", "MaxHP") & "  Mana: " & GetVar(CharPath & nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & _
                                                                                                                                                                                                                                                                   nombre & ".chr", "Stats", "MaxMAN") & FONTTYPE_INFO)

        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & GetVar(CharPath & nombre & ".chr", "stats", "MaxHIT") & _
                                                        FONTTYPE_INFO)

        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Oro: " & GetVar(CharPath & nombre & ".chr", "stats", "GLD") & FONTTYPE_INFO)

        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Trofeos de Oro: " & GetVar(CharPath & nombre & ".chr", "stats", "TrofOro") & _
                                                        "~255~255~6~0~0~")
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Trofeos de Plata: " & GetVar(CharPath & nombre & ".chr", "stats", "TrofPlata") & _
                                                        "~255~255~251~0~0~")
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Trofeos de Bronce: " & GetVar(CharPath & nombre & ".chr", "stats", "TrofBronce") & _
                                                        "~187~0~0~0~0~")
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Amuletos de Madera: " & GetVar(CharPath & nombre & ".chr", "stats", "TrofMadera") & _
                                                        "~237~207~139~0~0~")

    End If

    Exit Sub

End Sub

Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)

    On Error Resume Next

    Dim j As Integer
    Dim CharFile As String, Tmp As String
    Dim ObjInd As Long, ObjCant As Long

    CharFile = CharPath & CharName & ".chr"

    If FileExist(CharFile, vbNormal) Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||" & CharName & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco." & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Usuario inexistente: " & CharName & FONTTYPE_INFO)

    End If

End Sub

Public Sub FxDoMeditar(ByVal UserIndex As Integer)

    Dim esNivel As Byte
    Dim esFaccion As Boolean

    Dim Fxs As Integer

    With UserList(UserIndex)
        esNivel = .Stats.ELV
        esFaccion = (.Faccion.Nemesis = 1 Or .Faccion.Templario = 1 Or .Faccion.ArmadaReal = 1 Or .Faccion.FuerzasCaos = 1)

        If EsGmChar(.Name) Then
            Fxs = 17
            Call SendData(SendTarget.ToPCArea, UserIndex, .Pos.Map, "CFX" & .char.CharIndex & "," & Fxs & "," & LoopAdEternum)

            .char.FX = Fxs
            Exit Sub

        End If

        If esFaccion Then

            Dim esNemesis As Byte
            Dim esTemplario As Byte
            Dim esArmada As Byte
            Dim esCaos As Byte

            esNemesis = .Faccion.Nemesis
            esTemplario = .Faccion.Templario
            esArmada = .Faccion.ArmadaReal
            esCaos = .Faccion.FuerzasCaos

            If esNemesis Then

                If esNivel < 35 Then
                    Fxs = 43
                ElseIf esNivel < 45 Then
                    Fxs = 40
                ElseIf esNivel < 55 Then
                    Fxs = 41
                ElseIf esNivel = STAT_MAXELV Then
                    Fxs = 46

                End If

            End If

            If esTemplario Then

                If esNivel < 35 Then
                    Fxs = 33
                ElseIf esNivel < 45 Then
                    Fxs = 35
                ElseIf esNivel < 55 Then
                    Fxs = 36
                ElseIf esNivel = STAT_MAXELV Then
                    Fxs = 47

                End If

            End If

            If esArmada Then

                If esNivel < 25 Then
                    Fxs = 31
                ElseIf esNivel < 31 Then
                    Fxs = 20
                ElseIf esNivel < 45 Then
                    Fxs = 32
                ElseIf esNivel < 55 Then
                    Fxs = 28
                ElseIf esNivel = STAT_MAXELV Then
                    Fxs = 45

                End If

            End If

            If esCaos Then

                If esNivel < 25 Then
                    Fxs = 29
                ElseIf esNivel < 31 Then
                    Fxs = 21
                ElseIf esNivel < 45 Then
                    Fxs = 30
                ElseIf esNivel < 55 Then
                    Fxs = 27
                ElseIf esNivel = STAT_MAXELV Then
                    Fxs = 44

                End If

            End If

        Else

            If esNivel < 10 Then
                Fxs = 26
            ElseIf esNivel < 15 Then

                If Criminal(UserIndex) Then
                    Fxs = 4
                Else
                    Fxs = 48

                End If

            ElseIf esNivel < 20 Then

                If Criminal(UserIndex) Then
                    Fxs = 6
                Else
                    Fxs = 5

                End If

            ElseIf esNivel < 35 Then

                If Criminal(UserIndex) Then
                    Fxs = 16
                Else
                    Fxs = 49

                End If

            ElseIf esNivel < 45 Then

                If Criminal(UserIndex) Then
                    Fxs = 50
                Else
                    Fxs = 23

                End If

            ElseIf esNivel <= 55 Then

                If Criminal(UserIndex) Then
                    Fxs = 51
                Else
                    Fxs = 18

                End If

            End If

        End If

        Call SendData(SendTarget.ToPCArea, UserIndex, .Pos.Map, "CFX" & .char.CharIndex & "," & Fxs & "," & LoopAdEternum)

        .char.FX = Fxs

    End With

End Sub

Private Sub AumentoStatsClase(ByVal UserIndex As Integer, _
                              ByVal Clase As String, _
                              ByRef AumentoHP As Integer, _
                              ByRef AumentoMANA As Integer, _
                              ByRef AumentoSTA As Integer, _
                              ByRef AumentoHIT As Integer)

    Select Case Clase

    Case "GUERRERO"

        Select Case UserList(UserIndex).Stats.UserAtributos(constitucion)

        Case 21
            AumentoHP = RandomNumber(GCONST21MINVIDA, GCONST21MAXVIDA)

        Case 20
            AumentoHP = RandomNumber(GCONST20MINVIDA, GCONST20MAXVIDA)

        Case 19
            AumentoHP = RandomNumber(GCONST19MINVIDA, GCONST19MAXVIDA)

        Case 18
            AumentoHP = RandomNumber(GCONST18MINVIDA, GCONST18MAXVIDA)

        Case 17
            AumentoHP = RandomNumber(GCONST17MINVIDA, GCONST17MAXVIDA)

        Case Else
            AumentoHP = RandomNumber(GCONSTOTROMINVIDA, GCONSTOTROMAXVIDA)

        End Select

        'AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 2, 3)
        If UserList(UserIndex).Stats.ELV = 1 Then
            AumentoHIT = 2

        ElseIf UserList(UserIndex).Stats.ELV < 36 Then
            AumentoHIT = 3

        ElseIf UserList(UserIndex).Stats.ELV >= 36 And UserList(UserIndex).Stats.ELV < 46 Then
            AumentoHIT = 5

        Else
            AumentoHIT = 7

        End If

        AumentoSTA = AumentoSTDef

    Case "CAZADOR"

        Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion)

        Case 21
            AumentoHP = RandomNumber(CCONST21MINVIDA, CCONST21MAXVIDA)

        Case 20
            AumentoHP = RandomNumber(CCONST20MINVIDA, CCONST20MAXVIDA)

        Case 19
            AumentoHP = RandomNumber(CCONST19MINVIDA, CCONST19MAXVIDA)

        Case 18
            AumentoHP = RandomNumber(CCONST18MINVIDA, CCONST18MAXVIDA)

        Case 17
            AumentoHP = RandomNumber(CCONST17MINVIDA, CCONST17MAXVIDA)

        Case Else
            AumentoHP = RandomNumber(CCONSTOTROMINVIDA, CCONSTOTROMAXVIDA)

        End Select

        AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 2, 3)
        AumentoSTA = AumentoSTDef

    Case "PALADIN"

        Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion)

        Case 21
            AumentoHP = RandomNumber(PCONST21MINVIDA, PCONST21MAXVIDA)

        Case 20
            AumentoHP = RandomNumber(PCONST20MINVIDA, PCONST20MAXVIDA)

        Case 19
            AumentoHP = RandomNumber(PCONST19MINVIDA, PCONST19MAXVIDA)

        Case 18
            AumentoHP = RandomNumber(PCONST18MINVIDA, PCONST18MAXVIDA)

        Case 17
            AumentoHP = RandomNumber(PCONST17MINVIDA, PCONST17MAXVIDA)

        Case Else
            AumentoHP = RandomNumber(PCONSTOTROMINVIDA, PCONSTOTROMAXVIDA)

        End Select

        'AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 1, 3)
        If UserList(UserIndex).Stats.ELV = 1 Then
            AumentoHIT = 2

        ElseIf UserList(UserIndex).Stats.ELV > 2 And UserList(UserIndex).Stats.ELV < 36 Then
            AumentoHIT = 3

        ElseIf UserList(UserIndex).Stats.ELV >= 36 And UserList(UserIndex).Stats.ELV <= 45 Then
            AumentoHIT = 4

        Else
            AumentoHIT = 5

        End If

        AumentoMANA = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
        AumentoSTA = AumentoSTDef

    Case "MAGO"

        Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion)

        Case 21
            AumentoHP = RandomNumber(MCONST21MINVIDA, MCONST21MAXVIDA)

        Case 20
            AumentoHP = RandomNumber(MCONST20MINVIDA, MCONST20MAXVIDA)

        Case 19
            AumentoHP = RandomNumber(MCONST19MINVIDA, MCONST19MAXVIDA)

        Case 18
            AumentoHP = RandomNumber(MCONST18MINVIDA, MCONST18MAXVIDA)

        Case 17
            AumentoHP = RandomNumber(MCONST17MINVIDA, MCONST17MAXVIDA)

        Case Else
            AumentoHP = RandomNumber(MCONSTOTROMINVIDA, MCONSTOTROMAXVIDA)

        End Select

        If AumentoHP < 1 Then AumentoHP = 4

        AumentoHIT = 2
        AumentoMANA = 3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
        AumentoSTA = AumentoSTMago

    Case "CLERIGO"

        Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion)

        Case 21
            AumentoHP = RandomNumber(CLCONST21MINVIDA, CLCONST21MAXVIDA)

        Case 20
            AumentoHP = RandomNumber(CLCONST20MINVIDA, CLCONST20MAXVIDA)

        Case 19
            AumentoHP = RandomNumber(CLCONST19MINVIDA, CLCONST19MAXVIDA)

        Case 18
            AumentoHP = RandomNumber(CLCONST18MINVIDA, CLCONST18MAXVIDA)

        Case 17
            AumentoHP = RandomNumber(CLCONST17MINVIDA, CLCONST17MAXVIDA)

        Case Else
            AumentoHP = RandomNumber(CLCONSTOTROMINVIDA, CLCONSTOTROMAXVIDA)

        End Select

        AumentoHIT = 2
        AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
        AumentoSTA = AumentoSTDef

    Case "ASESINO"

        Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion)

        Case 21
            AumentoHP = RandomNumber(ACONST21MINVIDA, ACONST21MAXVIDA)

        Case 20
            AumentoHP = RandomNumber(ACONST20MINVIDA, ACONST20MAXVIDA)

        Case 19
            AumentoHP = RandomNumber(ACONST19MINVIDA, ACONST19MAXVIDA)

        Case 18
            AumentoHP = RandomNumber(ACONST18MINVIDA, ACONST18MAXVIDA)

        Case 17
            AumentoHP = RandomNumber(ACONST17MINVIDA, ACONST17MAXVIDA)

        Case Else
            AumentoHP = RandomNumber(ACONSTOTROMINVIDA, ACONSTOTROMAXVIDA)

        End Select

        'AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 1, 3)
        If UserList(UserIndex).Stats.ELV = 1 Then
            AumentoHIT = 2

        ElseIf UserList(UserIndex).Stats.ELV < 26 Then
            AumentoHIT = 3

        Else
            AumentoHIT = 2

        End If

        AumentoMANA = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) * 1.2
        AumentoSTA = AumentoSTDef

    Case "BARDO"

        Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion)

        Case 21
            AumentoHP = RandomNumber(BACONST21MINVIDA, BACONST21MAXVIDA)

        Case 20
            AumentoHP = RandomNumber(BACONST20MINVIDA, BACONST20MAXVIDA)

        Case 19
            AumentoHP = RandomNumber(BACONST19MINVIDA, BACONST19MAXVIDA)

        Case 18
            AumentoHP = RandomNumber(BACONST18MINVIDA, BACONST18MAXVIDA)

        Case 17
            AumentoHP = RandomNumber(BACONST17MINVIDA, BACONST17MAXVIDA)

        Case Else
            AumentoHP = RandomNumber(BACONSTOTROMINVIDA, BACONSTOTROMAXVIDA)

        End Select

        'AumentoHIT = 2

        If UserList(UserIndex).Stats.ELV = 1 Then
            AumentoHIT = 2

        Else
            AumentoHIT = 3

        End If

        AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
        AumentoSTA = AumentoSTDef

    Case "LADRON"

        Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion)

        Case 21
            AumentoHP = RandomNumber(LCONST21MINVIDA, LCONST21MAXVIDA)

        Case 20
            AumentoHP = RandomNumber(LCONST20MINVIDA, LCONST20MAXVIDA)

        Case 19
            AumentoHP = RandomNumber(LCONST19MINVIDA, LCONST19MAXVIDA)

        Case 18
            AumentoHP = RandomNumber(LCONST18MINVIDA, LCONST18MAXVIDA)

        Case 17
            AumentoHP = RandomNumber(LCONST17MINVIDA, LCONST17MAXVIDA)

        Case Else
            AumentoHP = RandomNumber(LCONSTOTROMINVIDA, LCONSTOTROMAXVIDA)

        End Select

        AumentoHIT = 2
        AumentoSTA = AumentoSTDef

    Case "DRUIDA"

        Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion)

        Case 21
            AumentoHP = RandomNumber(DCONST21MINVIDA, DCONST21MAXVIDA)

        Case 20
            AumentoHP = RandomNumber(DCONST20MINVIDA, DCONST20MAXVIDA)

        Case 19
            AumentoHP = RandomNumber(DCONST19MINVIDA, DCONST19MAXVIDA)

        Case 18
            AumentoHP = RandomNumber(DCONST18MINVIDA, DCONST18MAXVIDA)

        Case 17
            AumentoHP = RandomNumber(DCONST17MINVIDA, DCONST17MAXVIDA)

        Case Else
            AumentoHP = RandomNumber(DCONSTOTROMINVIDA, DCONSTOTROMAXVIDA)

        End Select

        AumentoHIT = 2
        AumentoMANA = 2.15 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
        AumentoSTA = AumentoSTDef

    Case "TRABAJADOR"

        Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion)

        Case 21
            AumentoHP = RandomNumber(TCONST21MINVIDA, TCONST21MAXVIDA)

        Case 20
            AumentoHP = RandomNumber(TCONST20MINVIDA, TCONST20MAXVIDA)

        Case 19
            AumentoHP = RandomNumber(TCONST19MINVIDA, TCONST19MAXVIDA)

        Case 18
            AumentoHP = RandomNumber(TCONST18MINVIDA, TCONST18MAXVIDA)

        Case 17
            AumentoHP = RandomNumber(TCONST17MINVIDA, TCONST17MAXVIDA)

        Case Else
            AumentoHP = RandomNumber(TCONSTOTROMINVIDA, TCONSTOTROMAXVIDA)

        End Select

        AumentoHIT = 2
        AumentoSTA = AumentoSTDef

    Case "PIRATA"

        Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion)

        Case 21
            AumentoHP = RandomNumber(TCONST21MINVIDA, TCONST21MAXVIDA)

        Case 20
            AumentoHP = RandomNumber(TCONST20MINVIDA, TCONST20MAXVIDA)

        Case 19
            AumentoHP = RandomNumber(TCONST19MINVIDA, TCONST19MAXVIDA)

        Case 18
            AumentoHP = RandomNumber(TCONST18MINVIDA, TCONST18MAXVIDA)

        Case 17
            AumentoHP = RandomNumber(TCONST17MINVIDA, TCONST17MAXVIDA)

        Case Else
            AumentoHP = RandomNumber(TCONSTOTROMINVIDA, TCONSTOTROMAXVIDA)

        End Select

        AumentoHIT = 2
        AumentoSTA = AumentoSTDef

    Case "BRUJO"

        Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion)

        Case 21
            AumentoHP = RandomNumber(BCONST21MINVIDA, BCONST21MAXVIDA)

        Case 20
            AumentoHP = RandomNumber(BCONST20MINVIDA, BCONST20MAXVIDA)

        Case 19
            AumentoHP = RandomNumber(BCONST19MINVIDA, BCONST19MAXVIDA)

        Case 18
            AumentoHP = RandomNumber(BCONST18MINVIDA, BCONST18MAXVIDA)

        Case 17
            AumentoHP = RandomNumber(BCONST17MINVIDA, BCONST17MAXVIDA)

        Case Else
            AumentoHP = RandomNumber(BCONSTOTROMINVIDA, BCONSTOTROMAXVIDA)

        End Select

        AumentoHIT = 2
        AumentoMANA = 2.5 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
        AumentoSTA = AumentoSTDef

    Case "ARQUERO"

        Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion)

        Case 21
            AumentoHP = RandomNumber(ARCONST21MINVIDA, ARCONST21MAXVIDA)

        Case 20
            AumentoHP = RandomNumber(ARCONST20MINVIDA, ARCONST20MAXVIDA)

        Case 19
            AumentoHP = RandomNumber(ARCONST19MINVIDA, ARCONST19MAXVIDA)

        Case 18
            AumentoHP = RandomNumber(ARCONST18MINVIDA, ARCONST18MAXVIDA)

        Case 17
            AumentoHP = RandomNumber(ARCONST17MINVIDA, ARCONST17MAXVIDA)

        Case Else
            AumentoHP = RandomNumber(ARCONSTOTROMINVIDA, ARCONSTOTROMAXVIDA)

        End Select

        AumentoHIT = 2
        AumentoSTA = AumentoSTDef


    Case "DIOS"

        Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion)

        Case 21
            AumentoHP = RandomNumber(ARCONST21MINVIDA, ARCONST21MAXVIDA)

        Case 20
            AumentoHP = RandomNumber(ARCONST20MINVIDA, ARCONST20MAXVIDA)

        Case 19
            AumentoHP = RandomNumber(ARCONST19MINVIDA, ARCONST19MAXVIDA)

        Case 18
            AumentoHP = RandomNumber(ARCONST18MINVIDA, ARCONST18MAXVIDA)

        Case 17
            AumentoHP = RandomNumber(ARCONST17MINVIDA, ARCONST17MAXVIDA)

        Case Else
            AumentoHP = RandomNumber(ARCONSTOTROMINVIDA, ARCONSTOTROMAXVIDA)

        End Select

        If AumentoHP < 1 Then AumentoHP = 4

        AumentoHIT = 2
        AumentoMANA = 3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
        AumentoSTA = AumentoSTMago


    Case Else

        Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion)

        Case 21
            AumentoHP = RandomNumber(6, 9)

        Case 20
            AumentoHP = RandomNumber(5, 9)

        Case 19, 18
            AumentoHP = RandomNumber(4, 8)

        Case Else
            AumentoHP = RandomNumber(5, UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) \ 2) - AdicionalHPCazador

        End Select

        AumentoHIT = 2
        AumentoSTA = AumentoSTDef

    End Select


End Sub

Public Sub SetearInv(ByVal UserIndex As Integer, ByVal Clase As String, ByVal Raza As String)

    Dim Slot As Byte

    Dim ArmaObjIndex As Integer
    Dim ArmaSlot As Byte

    Dim ArmorObjIndex As Integer
    Dim ArmorSlot As Byte

    Dim WeapongObjIndex As Integer
    Dim WeaponSlot As Integer

    Dim CascoObjIndex As Integer
    Dim CascoSlot As Byte

    Dim EscuObjIndex As Integer
    Dim EscuSlot As Byte

    With UserList(UserIndex).Invent

        '???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿

        Slot = Slot + 1
        .Object(Slot).ObjIndex = 467
        .Object(Slot).Amount = 100

        Slot = Slot + 1
        .Object(Slot).ObjIndex = 468
        .Object(Slot).Amount = 100

        Slot = Slot + 1
        .Object(Slot).ObjIndex = 460
        .Object(Slot).Amount = 1

        Select Case Clase

        Case "GUERRERO"

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 461
            .Object(Slot).Amount = 50

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 462
            .Object(Slot).Amount = 50

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 948
            .Object(Slot).Amount = 50

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1177
            .Object(Slot).Amount = 50

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1178
            .Object(Slot).Amount = 1
            .Object(Slot).Equipped = 1

            ArmaObjIndex = 1178
            ArmaSlot = Slot

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1179
            .Object(Slot).Amount = 1
            .Object(Slot).Equipped = 1

            EscuObjIndex = 1179
            EscuSlot = Slot

        Case "PALADIN", "ASESINO", "CLERIGO", "BARDO"

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1177
            .Object(Slot).Amount = 50

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1178
            .Object(Slot).Amount = 1
            .Object(Slot).Equipped = 1

            ArmaObjIndex = 1178
            ArmaSlot = Slot

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1179
            .Object(Slot).Amount = 1
            .Object(Slot).Equipped = 1

            EscuObjIndex = 1179
            EscuSlot = Slot

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 395
            .Object(Slot).Amount = 50

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 461
            .Object(Slot).Amount = 50

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 462
            .Object(Slot).Amount = 50

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 948
            .Object(Slot).Amount = 50

        Case "CAZADOR", "ARQUERO"
            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1173
            .Object(Slot).Amount = 1

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1174
            .Object(Slot).Amount = 500

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1177
            .Object(Slot).Amount = 50

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 462
            .Object(Slot).Amount = 50

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 948
            .Object(Slot).Amount = 50

        Case "BRUJO", "MAGO"
            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1175
            .Object(Slot).Amount = 1
            .Object(Slot).Equipped = 1

            CascoObjIndex = 1175
            CascoSlot = Slot

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1176
            .Object(Slot).Amount = 1
            .Object(Slot).Equipped = 1

            ArmaObjIndex = 1176
            ArmaSlot = Slot

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1177
            .Object(Slot).Amount = 50

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 395
            .Object(Slot).Amount = 50

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 461
            .Object(Slot).Amount = 50

        Case "TRABAJADOR"
            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1173
            .Object(Slot).Amount = 1

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1174
            .Object(Slot).Amount = 500

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1177
            .Object(Slot).Amount = 50

        Case "DRUIDA"

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 462
            .Object(Slot).Amount = 25

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 948
            .Object(Slot).Amount = 25

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 461
            .Object(Slot).Amount = 50

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 395
            .Object(Slot).Amount = 50

        Case "PIRATA"

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1178
            .Object(Slot).Amount = 1

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1179
            .Object(Slot).Amount = 1

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1177
            .Object(Slot).Amount = 50

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 395
            .Object(Slot).Amount = 25

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 461
            .Object(Slot).Amount = 25

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 462
            .Object(Slot).Amount = 25

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 948
            .Object(Slot).Amount = 25

        End Select

        Select Case Raza

        Case "HOBBIT"
            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1130
            .Object(Slot).Amount = 1
            .Object(Slot).Equipped = 1

            ArmorObjIndex = 1130
            ArmorSlot = Slot

        Case "ORCO"
            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1131
            .Object(Slot).Amount = 1
            .Object(Slot).Equipped = 1

            ArmorObjIndex = 1131
            ArmorSlot = Slot

        Case "VAMPIRO"
            Slot = Slot + 1
            .Object(Slot).ObjIndex = 1171
            .Object(Slot).Amount = 1
            .Object(Slot).Equipped = 1

            ArmorObjIndex = 1171
            ArmorSlot = Slot

        Case "ENANO", "GNOMO"
            Slot = Slot + 1
            .Object(Slot).ObjIndex = 466
            .Object(Slot).Amount = 1
            .Object(Slot).Equipped = 1

            ArmorObjIndex = 466
            ArmorSlot = Slot

        Case Else

            Slot = Slot + 1
            .Object(Slot).ObjIndex = 463
            .Object(Slot).Amount = 1
            .Object(Slot).Equipped = 1

            ArmorObjIndex = 463
            ArmorSlot = Slot

        End Select

        .ArmourEqpObjIndex = ArmorObjIndex
        .ArmourEqpSlot = ArmorSlot

        .CascoEqpObjIndex = CascoObjIndex
        .CascoEqpSlot = CascoSlot

        .EscudoEqpObjIndex = EscuObjIndex
        .EscudoEqpSlot = EscuSlot

        .WeaponEqpObjIndex = ArmaObjIndex
        .WeaponEqpSlot = ArmaSlot

        If .ArmourEqpObjIndex <> 0 Then
            UserList(UserIndex).char.Body = ObjData(.ArmourEqpObjIndex).Ropaje

        End If

        If .CascoEqpObjIndex <> 0 Then
            UserList(UserIndex).char.CascoAnim = ObjData(.CascoEqpObjIndex).CascoAnim

        End If

        If .EscudoEqpObjIndex <> 0 Then
            UserList(UserIndex).char.ShieldAnim = ObjData(.EscudoEqpObjIndex).ShieldAnim

        End If

        If .WeaponEqpObjIndex <> 0 Then
            UserList(UserIndex).char.WeaponAnim = ObjData(.WeaponEqpObjIndex).WeaponAnim

        End If

        '[MaTeO 9]
        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                        UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, _
                            UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

        '[/MaTeO 9]

        .NroItems = Slot

    End With

End Sub

Public Sub DragToUser(ByVal UserIndex As Integer, ByVal TIndex As Integer, ByVal Slot As Byte, ByVal Amount As Integer)

' @@ Drag un slot a un usuario.

    Dim tObj As Obj
    Dim tString As String
    Dim errorFound As String
    Dim Espacio As Boolean

    ' Puede dragear ?
    If Not CanDragObj(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.Navegando, UserList(UserIndex).flags.Muerto, errorFound) Then
        WriteConsoleMsg UserIndex, errorFound, FONTTYPE_INFOBOLD

        Exit Sub

    End If

    ' Puede dragear ?
    If Not CanDragObj(UserList(TIndex).Pos.Map, UserList(TIndex).flags.Navegando, UserList(TIndex).flags.Muerto, errorFound) Then
        WriteConsoleMsg UserIndex, errorFound, FONTTYPE_INFOBOLD

        Exit Sub

    End If

    'Preparo el objeto.
    tObj.Amount = Amount
    tObj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex

    If (Amount <= 0 Or Amount > UserList(UserIndex).Invent.Object(Slot).Amount) Then
        Call WriteConsoleMsg(UserIndex, "Cantidad invalida", FONTTYPE_INFO)
        Exit Sub

    End If

    'TmpIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    Espacio = MeterItemEnInventario(TIndex, tObj)

    'No tiene espacio.

    If Not Espacio Then
        WriteConsoleMsg UserIndex, "El usuario no tiene espacio en su inventario.", FONTTYPE_INFOBOLD
        Exit Sub

    End If

    'Quito el objeto.
    QuitarUserInvItem UserIndex, Slot, Amount

    'Hago un update de su inventario.
    UpdateUserInv False, UserIndex, Slot

    'Preparo el mensaje para userINdex (quien dragea)

    tString = "Le has arrojado"

    If tObj.Amount <> 1 Then
        tString = tString & " " & tObj.Amount & " - " & ObjData(tObj.ObjIndex).Name
    Else
        tString = tString & " Tu " & ObjData(tObj.ObjIndex).Name

    End If

    tString = tString & " ah " & UserList(TIndex).Name

    'Envio el mensaje
    WriteConsoleMsg UserIndex, tString, FONTTYPE_INFOBOLD

    'Preparo el mensaje para el otro usuario (quien recibe)
    tString = UserList(UserIndex).Name & " Te ha arrojado"

    If tObj.Amount <> 1 Then
        tString = tString & " " & tObj.Amount & " - " & ObjData(tObj.ObjIndex).Name
    Else
        tString = tString & " su " & ObjData(tObj.ObjIndex).Name

    End If

    'Envio el mensaje al otro usuario
    WriteConsoleMsg TIndex, tString & ".", FONTTYPE_INFOBOLD

End Sub

Public Sub DragToNPC(ByVal UserIndex As Integer, ByVal tNpc As Integer, ByVal Slot As Byte, ByVal Amount As Integer)

' @@ Drag un slot a un npc.

    On Error GoTo errhandler

    Dim TeniaOro As Long
    Dim TeniaObj As Integer
    Dim TmpIndex As Integer
    Dim errorFound As String

    TmpIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    TeniaOro = UserList(UserIndex).Stats.GLD
    TeniaObj = UserList(UserIndex).Invent.Object(Slot).Amount

    ' Puede dragear ?
    If Not CanDragObj(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.Navegando, UserList(UserIndex).flags.Muerto, errorFound) Then
        WriteConsoleMsg UserIndex, errorFound, FONTTYPE_INFOBOLD

        Exit Sub

    End If

    If (Amount <= 0 Or Amount > UserList(UserIndex).Invent.Object(Slot).Amount) Then
        Call WriteConsoleMsg(UserIndex, "Cantidad invalida", FONTTYPE_INFO)
        Exit Sub

    End If

    'Es un banquero?

    If Npclist(tNpc).NPCtype = eNPCType.Banquero Then
        Call UserDejaObj(UserIndex, Slot, Amount)

        'No tiene más el mismo amount que antes? entonces depositó.

        If TeniaObj <> UserList(UserIndex).Invent.Object(Slot).Amount Then
            WriteConsoleMsg UserIndex, "Has depositado " & Amount & " - " & ObjData(TmpIndex).Name & ".", FONTTYPE_INFOBOLD
            UpdateUserInv False, UserIndex, Slot

        End If

        'Es un npc comerciante?
    ElseIf Npclist(tNpc).Comercia = 1 Then
        'El npc compra cualquier tipo de items?

        If Not Npclist(tNpc).TipoItems <> eOBJType.otCualquiera Or Npclist(tNpc).TipoItems = ObjData(UserList(UserIndex).Invent.Object( _
                                                                                                     Slot).ObjIndex).ObjType Then

            Call NPCCompraItem(UserIndex, Slot, Amount, tNpc)

            'Ganó oro? si es así es porque lo vendió.

            If TeniaOro <> UserList(UserIndex).Stats.GLD Then
                WriteConsoleMsg UserIndex, "Le has vendido al " & Npclist(tNpc).Name & " " & Amount & " - " & ObjData(TmpIndex).Name & ".", _
                                FONTTYPE_INFOBOLD

            End If

        Else
            WriteConsoleMsg UserIndex, "El npc no está interesado en comprar este tipo de objetos.", FONTTYPE_INFOBOLD

        End If

    End If

    Exit Sub

errhandler:

End Sub

Public Sub DragToPos(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Slot As Byte, ByVal Amount As Integer)

'            Drag un slot a una posición.

    Dim errorFound As String
    Dim tObj As Obj
    Dim tString As String
    Dim TmpIndex As Integer

    'No puede dragear en esa pos?

    If Not UserList(UserIndex).flags.SeguroObjetos Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes tirar. Tienes el seguro de objetos activados!!!" & _
                                                        FONTTYPE_Motd4)
        Exit Sub
    End If

    If UserList(UserIndex).flags.Privilegios = PlayerType.User Then
        If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Newbie = 1 And EsNewbie(UserIndex) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes tirar el objeto." & FONTTYPE_INFO)
            Exit Sub
        End If

        If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Caos = 1 Or ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Real = 1 Or ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Templ = 1 Or ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Nemes = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes tirar el objeto." & FONTTYPE_INFO)
            Exit Sub
        End If
    End If


    ' Puede dragear ?
    If Not CanDragObj(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.Navegando, UserList(UserIndex).flags.Muerto, errorFound) Then
        WriteConsoleMsg UserIndex, errorFound, FONTTYPE_INFOBOLD

        Exit Sub

    End If

    'Creo el objeto.
    tObj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    tObj.Amount = Amount

    If (Amount <= 0 Or Amount > UserList(UserIndex).Invent.Object(Slot).Amount) Then
        Call WriteConsoleMsg(UserIndex, "Cantidad invalida", FONTTYPE_INFO)
        Exit Sub

    End If

    'Agrego el objeto a la posición.
    MakeObj SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, tObj, UserList(UserIndex).Pos.Map, X, Y

    'Quito el objeto.
    QuitarUserInvItem UserIndex, Slot, Amount

    'Actualizo el inventario
    UpdateUserInv False, UserIndex, Slot

    'Preparo el mensaje.
    tString = "Has arrojado "

    If tObj.Amount <> 1 Then
        tString = tString & tObj.Amount & " - " & ObjData(tObj.ObjIndex).Name
    Else
        tString = tString & "tu " & ObjData(tObj.ObjIndex).Name    'faltaba el tstring &

    End If

    'ENvio.
    WriteConsoleMsg UserIndex, tString & ".", FONTTYPE_INFOBOLD

End Sub

Private Function CanDragToPos(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByRef error As String) As Boolean

'            Devuelve si se puede dragear un item a x posición.

    CanDragToPos = False

    'Zona segura?

    If Not (MapInfo(Map).Pk) Then
        error = "No está permitido arrojar objetos al suelo en zonas seguras."
        Exit Function

    End If

    'Ya hay objeto?

    If Not (MapData(Map, X, Y).OBJInfo.ObjIndex = 0) Then
        error = "Hay un objeto en esa posición!"
        Exit Function

    End If

    'Tile bloqueado?

    If Not (MapData(Map, X, Y).TileExit.Map = 0) Then
        error = "No puedes arrojar objetos en esa posición"
        Exit Function

    End If

    'Tile bloqueado?

    If Not (MapData(Map, X, Y).Blocked = 0) Then
        error = "No puedes arrojar objetos en esa posición"
        Exit Function

    End If

    If (HayAgua(Map, X, Y)) Then
        error = "No puedes arrojar objetos al agua"
        Exit Function

    End If

    CanDragToPos = True

End Function

Public Sub WriteConsoleMsg(ByVal sendIndex As Integer, ByVal PacketId As String, ByVal Font As String)

    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||" & PacketId & Font)

End Sub

Private Function CanDragObj(ByVal ObjIndex As Integer, ByVal Navegando As Boolean, ByVal Muerto As Byte, ByRef error As String) As Boolean

'            Devuelve si un objeto es drageable.

    CanDragObj = False

    If ObjIndex < 1 Or ObjIndex > UBound(ObjData()) Then Exit Function

    'Objeto newbie?

    If ObjData(ObjIndex).Newbie <> 0 Then
        error = "No puedes arrojar objetos newbies!"
        Exit Function

    End If

    'Está navgeando?
    If Navegando Then
        error = "No puedes arrojar un objeto a un usuario en barco!"
        Exit Function

    End If

    If Muerto = 1 Then
        error = "No puedes arrojar objetos a un muerto!"
        Exit Function

    End If

    CanDragObj = True

End Function

Public Sub moveItem(ByVal UserIndex As Integer, ByVal originalSlot As Integer, ByVal NewSlot As Integer)

    Dim tmpObj As UserOBJ
    Dim newObjIndex As Byte, originalObjIndex As Byte

    If (originalSlot <= 0) Or (NewSlot <= 0) Then Exit Sub

    With UserList(UserIndex)

        If (originalSlot > MAX_INVENTORY_SLOTS) Or (NewSlot > MAX_INVENTORY_SLOTS) Then Exit Sub

        tmpObj = .Invent.Object(originalSlot)
        .Invent.Object(originalSlot) = .Invent.Object(NewSlot)
        .Invent.Object(NewSlot) = tmpObj

        'Viva VB6 y sus putas deficiencias.

        If .Invent.ArmourEqpSlot = originalSlot Then
            .Invent.ArmourEqpSlot = NewSlot
        ElseIf .Invent.ArmourEqpSlot = NewSlot Then
            .Invent.ArmourEqpSlot = originalSlot

        End If

        If .Invent.BarcoSlot = originalSlot Then
            .Invent.BarcoSlot = NewSlot
        ElseIf .Invent.BarcoSlot = NewSlot Then
            .Invent.BarcoSlot = originalSlot

        End If

        If .Invent.CascoEqpSlot = originalSlot Then
            .Invent.CascoEqpSlot = NewSlot
        ElseIf .Invent.CascoEqpSlot = NewSlot Then
            .Invent.CascoEqpSlot = originalSlot

        End If

        If .Invent.EscudoEqpSlot = originalSlot Then
            .Invent.EscudoEqpSlot = NewSlot
        ElseIf .Invent.EscudoEqpSlot = NewSlot Then
            .Invent.EscudoEqpSlot = originalSlot

        End If

        If .Invent.MunicionEqpSlot = originalSlot Then
            .Invent.MunicionEqpSlot = NewSlot
        ElseIf .Invent.MunicionEqpSlot = NewSlot Then
            .Invent.MunicionEqpSlot = originalSlot

        End If

        If .Invent.WeaponEqpSlot = originalSlot Then
            .Invent.WeaponEqpSlot = NewSlot
        ElseIf .Invent.WeaponEqpSlot = NewSlot Then
            .Invent.WeaponEqpSlot = originalSlot

        End If

        If .Invent.AlaEqpSlot = originalSlot Then
            .Invent.AlaEqpSlot = NewSlot
        ElseIf .Invent.AlaEqpSlot = NewSlot Then
            .Invent.AlaEqpSlot = originalSlot

        End If

        Call UpdateUserInv(False, UserIndex, originalSlot)
        Call UpdateUserInv(False, UserIndex, NewSlot)

    End With

End Sub

Sub IniciarChangeHead(ByVal UserIndex As Integer)

    Dim CantHead As Byte
    Dim Heads As String

    Select Case UCase$(UserList(UserIndex).Raza)

    Case "HUMANO"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"
            CantHead = 25
            Heads = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,16,17,18,19,20,21,22,23,24,25,26"

        Case "MUJER"
            CantHead = "7"
            Heads = "68,69,70,71,72,74,75"

        End Select

    Case "ELFO"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"
            CantHead = 10
            Heads = "102,103,104,105,106,107,108,109,110,210"

        Case "MUJER"
            CantHead = "5"
            Heads = "107,108,170,171,172"

        End Select

    Case "ELFO OSCURO"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"
            CantHead = 2
            Heads = "202,203"

        Case "MUJER"
            CantHead = "3"
            Heads = "270,271,272"

        End Select

    Case "ENANO"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"
            CantHead = 10
            Heads = "301,302,303,304,305,306,307,308,309,310"

        Case "MUJER"
            CantHead = "1"
            Heads = "370"

        End Select

    Case "GNOMO"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"
            CantHead = 1
            Heads = "401"

        Case "MUJER"
            CantHead = "4"
            Heads = "470,471,472,473"

        End Select

    Case "HOBBIT"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"
            CantHead = 3
            Heads = "609,610,611"

        Case "MUJER"
            CantHead = "4"
            Heads = "612,613,614,615"

        End Select

    Case "ORCO"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"
            CantHead = 4
            Heads = "602,603,604,605"

        Case "MUJER"
            CantHead = "8"
            Heads = "606,607"

        End Select

    Case "LICANTROPO"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"
            CantHead = 12
            Heads = "1,2,3,4,5,6,7,8,9,10,11,19"

        Case "MUJER"
            CantHead = "5"
            Heads = "68,69,70,71,72"

        End Select

    Case "VAMPIRO"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"
            CantHead = 3
            Heads = "710,711,712"

        Case "MUJER"
            CantHead = "3"
            Heads = "710,711,712"

        End Select

    Case "CICLOPE"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"
            CantHead = 3
            Heads = "530,531,532"

        Case "MUJER"
            CantHead = "3"
            Heads = "533,534,535"

        End Select

    End Select

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ABRC" & CantHead & "@" & Heads)

End Sub

Public Function MismaParty(ByVal UserIndex As Integer, ByVal PIndex As Integer) As Boolean

    If UserList(UserIndex).PartyIndex > 0 Then

        If UserList(PIndex).PartyIndex > 0 Then

            If UserList(UserIndex).PartyIndex = UserList(PIndex).PartyIndex Then

                MismaParty = True
                Exit Function

            End If

        End If

    End If

    MismaParty = False

End Function

Public Function MismoClan(ByVal UserIndex As Integer, ByVal CIndex As Integer) As Boolean

    If UserList(UserIndex).GuildIndex > 0 Then
        If UserList(CIndex).GuildIndex > 0 Then
            If UserList(UserIndex).GuildIndex = UserList(CIndex).GuildIndex Then
                If UserList(UserIndex).flags.SeguroClan = True Then
                    MismoClan = False
                    Exit Function
                Else
                    MismoClan = True
                    Exit Function
                End If
            End If
        End If
    End If

    MismoClan = False

End Function

Sub LlevarUsuarios()

    Dim ijaji As Integer

    For ijaji = 1 To LastUser

        If UserList(ijaji).Pos.Map = 8 And UserList(ijaji).EnCvc = True Then
            Call WarpUserChar(ijaji, UserList(ijaji).ViejaPos.Map, UserList(ijaji).ViejaPos.X, UserList(ijaji).ViejaPos.Y, True)
            UserList(ijaji).EnCvc = False
            UserList(ijaji).flags.CvcBlue = 0
            UserList(ijaji).flags.CvcRed = 0
            Call SendData(SendTarget.ToMap, 0, UserList(ijaji).Pos.Map, "CVB" & UserList(ijaji).char.CharIndex & "," & UserList(ijaji).flags.CvcBlue)
            Call SendData(SendTarget.ToMap, 0, UserList(ijaji).Pos.Map, "CVR" & UserList(ijaji).char.CharIndex & "," & UserList(ijaji).flags.CvcRed)

        End If

    Next ijaji

End Sub

Public Function EsUserMute(ByVal IndexHabla As Integer, ByVal IndexRecibe As Integer) As Boolean
        
        Dim LoopC As Integer
        
        If IndexHabla = IndexRecibe Then
            EsUserMute = False
            Exit Function
        End If
        
        If UserList(IndexRecibe).Ignore.NumIgnores > 0 Then
             
             For LoopC = 1 To UserList(IndexRecibe).Ignore.NumIgnores
                  
                  If UCase$(UserList(IndexRecibe).Ignore.Usuario(LoopC)) = UCase$(UserList(IndexHabla).Name) Then
                      EsUserMute = True
                      Exit Function
                  End If
                  
             Next LoopC
             
        End If
        
        EsUserMute = False
        
End Function
