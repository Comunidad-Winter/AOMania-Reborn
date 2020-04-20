Attribute VB_Name = "mod_Castillos"
' Encapsulo el sistema de castillos que tenian aca
Option Explicit

Public Const NpcRey As Integer = 657
Public Const NpcFortaleza As Integer = 663

Public Const CastilloNorte As Byte = 98
Public Const CastilloSur As Byte = 99
Public Const CastilloEste As Byte = 100
Public Const CastilloOeste As Byte = 101
Public Const MapaFortaleza As Byte = 102
Public Const MapaFuerte As Byte = 164

Private TiempoCura As Integer
Private CuraMinimaRey As Integer
Private CuraMaximaRey As Integer
Public ExpConquista As Integer
Public OroConquista As Integer

Private RecompensaCastillo As Integer
Private IntervaloRecompensa As Integer

Public Norte As String
Public Sur As String
Public Oeste As String
Public Este As String
Public Fortaleza As String

Private HoraSur As String
Private HoraNorte As String
Private HoraEste As String
Private HoraOeste As String
Private HoraForta As String

Public Sub RecompensaCastillos()
    Dim GuildIndex As Integer

    RecompensaCastillo = RecompensaCastillo + 1

    If RecompensaCastillo >= IntervaloRecompensa Then

        If Len(Norte) > 0 Then

            GuildIndex = modGuilds.GuildIndex(Norte)

            Call modGuilds.RecompensasCastillos(GuildIndex, ExpConquista, OroConquista, "Norte")

        End If

        If Len(Sur) > 0 Then

            GuildIndex = modGuilds.GuildIndex(Sur)

            Call modGuilds.RecompensasCastillos(GuildIndex, ExpConquista, OroConquista, "Sur")

        End If

        If Len(Oeste) > 0 Then

            GuildIndex = modGuilds.GuildIndex(Oeste)

            Call modGuilds.RecompensasCastillos(GuildIndex, ExpConquista, OroConquista, "Oeste")

        End If

        If Len(Este) > 0 Then

            GuildIndex = modGuilds.GuildIndex(Este)

            Call modGuilds.RecompensasCastillos(GuildIndex, ExpConquista, OroConquista, "Este")

        End If

        If Len(Fortaleza) > 0 Then

            GuildIndex = modGuilds.GuildIndex(Fortaleza)

            Call modGuilds.RecompensasCastillos(GuildIndex, ExpConquista, OroConquista, "Fortaleza")

        End If

        RecompensaCastillo = 0

    End If

End Sub

Public Sub CuraRey(ByVal NpcIndex As Integer)
    Static Tiempo As Integer
    Dim HpCura As Integer

    If Npclist(NpcIndex).Stats.MinHP < 15000 And Not Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP Then
        Tiempo = Tiempo + 1
    Else
        Exit Sub
    End If

    If Tiempo >= TiempoCura Then

        With Npclist(NpcIndex)

            If .Numero = NpcRey Or .Numero = NpcFortaleza Then

                If .Stats.MinHP < 15000 And Not .Stats.MinHP = .Stats.MaxHP Then
                    HpCura = RandomNumber(CuraMinimaRey, CuraMaximaRey)
                    .Stats.MinHP = .Stats.MinHP + HpCura
                    Call SendData(SendTarget.ToMap, 0, .pos.Map, "||¡El rey del castillo se ha curado " & HpCura & " puntos de vida!" & FONTTYPE_TALKMSG)
                    If .Stats.MinHP > .Stats.MaxHP Then .Stats.MinHP = .Stats.MaxHP

                End If

            End If

        End With

        Tiempo = 0

    End If

End Sub

Public Sub ConnectFuerte(ByVal UserIndex As Integer)

    Dim GuildName As String

    With UserList(UserIndex)

        If .GuildIndex = 0 Then
            Call WarpUserChar(UserIndex, 34, 30, 50, False)
        ElseIf .GuildIndex > 0 Then
            GuildName = Guilds(.GuildIndex).GuildName

            If UCase$(GuildName) <> UCase$(Fortaleza) Then
                Call WarpUserChar(UserIndex, 34, 30, 50, False)
            End If

        End If
    End With

End Sub

Public Sub WarpCastillo(ByVal UserIndex As Integer, ByVal Castillo As String)

    With UserList(UserIndex)

        If .flags.EstaDueleando1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes defender el castillo estando en DUELOS." & FONTTYPE_WARNING)
            Exit Sub

        End If

        If .flags.Paralizado = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes teletransportarte porque estás afectado por un hechizo que te lo impide." _
                                                          & FONTTYPE_WARNING)
            Exit Sub

        End If

        If .Counters.Pena > 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes salir de la cárcel." & FONTTYPE_WARNING)
            Exit Sub

        End If

        If .GuildIndex = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes clan!" & FONTTYPE_INFO)
            Exit Sub
        End If

        If .flags.Montado = True Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes ir a defender castillo montado en tu mascota!!." & FONTTYPE_INFO)
            Exit Sub
        End If

        If .Stats.MinHP < .Stats.MaxHP Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu salud debe estar completa." & FONTTYPE_INFO)
            Exit Sub
        End If

        Dim X As Integer
        Dim Y As Integer
        Dim Map As Integer
        Dim GuildName As String
        Dim WarpUser As Boolean

        GuildName = Guilds(.GuildIndex).GuildName
        X = RandomNumber(57, 64)
        Y = RandomNumber(37, 40)

        Select Case UCase$(Castillo)

        Case "NORTE"
            Map = CastilloNorte
            WarpUser = (StrComp(GuildName, Norte, vbTextCompare) = 0)

        Case "SUR"
            Map = CastilloSur
            WarpUser = (StrComp(GuildName, Sur, vbTextCompare) = 0)

        Case "ESTE"
            Map = CastilloEste
            WarpUser = (StrComp(GuildName, Este, vbTextCompare) = 0)

        Case "OESTE"
            Map = CastilloOeste
            WarpUser = (StrComp(GuildName, Oeste, vbTextCompare) = 0)

        Case "FORTALEZA"
            Map = MapaFortaleza
            WarpUser = (StrComp(GuildName, Fortaleza, vbTextCompare) = 0)

        End Select

        If WarpUser Then

            Select Case UCase(Castillo)
            Case "NORTE"
                If UserList(UserIndex).Castillos.Norte = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo puedes ir a defender el castillo " & LCase$(Castillo) & " una vez cada 2 minutos." & FONTTYPE_INFO)
                    Exit Sub
                End If

            Case "OESTE"
                If UserList(UserIndex).Castillos.Oeste = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo puedes ir a defender el castillo " & LCase(Castillo) & " una vez cada 2 minutos." & FONTTYPE_INFO)
                    Exit Sub
                End If

            Case "ESTE"
                If UserList(UserIndex).Castillos.Este = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo puedes ir a defender el castillo " & LCase(Castillo) & " una vez cada 2 minutos." & FONTTYPE_INFO)
                    Exit Sub
                End If

            Case "SUR"
                If UserList(UserIndex).Castillos.Sur = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo puedes ir a defender el castillo " & LCase(Castillo) & " una vez cada 2 minutos." & FONTTYPE_INFO)
                    Exit Sub
                End If

            Case "FORTALEZA"
                If UserList(UserIndex).Castillos.Fortaleza = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo puedes ir a defender el castillo " & LCase(Castillo) & " una vez cada 2 minutos." & FONTTYPE_INFO)
                    Exit Sub
                End If
            End Select

            Call WarpUserChar(UserIndex, Map, X, Y, True)

            Select Case UCase$(Castillo)
            Case "NORTE"
                UserList(UserIndex).Castillos.Norte = 1
                UserList(UserIndex).Castillos.tNorte = 2

            Case "OESTE"
                UserList(UserIndex).Castillos.Oeste = 1
                UserList(UserIndex).Castillos.tOeste = 2

            Case "ESTE"
                UserList(UserIndex).Castillos.Este = 1
                UserList(UserIndex).Castillos.tEste = 2

            Case "SUR"
                UserList(UserIndex).Castillos.Sur = 1
                UserList(UserIndex).Castillos.tSur = 2

            Case "FORTALEZA"
                UserList(UserIndex).Castillos.Fortaleza = 1
                UserList(UserIndex).Castillos.tFortaleza = 2
            End Select
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El castillo " & LCase(Castillo) & " no le pertenece a tu clan." & FONTTYPE_INFO)

        End If

    End With

End Sub

Public Function GolpeNpcCastillo(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean

    Dim NumberNpc As Integer
    Dim UserMap As Integer
    Dim GuildName As String

    GolpeNpcCastillo = False

    With UserList(UserIndex)

        '  Si no tiene elegido ningun npc, par aque hacer el resto..
        If NpcIndex < 0 Then Exit Function
        NumberNpc = Npclist(NpcIndex).Numero

        ' si no es el Rey de castillo o fortaleza, al pedo.
        If NumberNpc = NpcRey Or NumberNpc = NpcFortaleza Then

            UserMap = .pos.Map

            If .GuildIndex = 0 And (UserMap >= CastilloNorte And UserMap <= MapaFortaleza) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes clan!" & FONTTYPE_INFO)
                Exit Function

            End If

            GuildName = Guilds(.GuildIndex).GuildName

            Select Case UserMap

            Case CastilloNorte

                If GuildName = Norte And NumberNpc = NpcRey Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                    Exit Function

                End If

            Case CastilloSur

                If GuildName = Sur And NumberNpc = NpcRey Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                    Exit Function

                End If

            Case CastilloEste

                If GuildName = Este And NumberNpc = NpcRey Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                    Exit Function

                End If

            Case CastilloOeste

                If GuildName = Oeste And NumberNpc = NpcRey Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                    Exit Function

                End If

            Case MapaFortaleza

                If GuildName = Fortaleza And NumberNpc = NpcFortaleza Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                    Exit Function

                End If

                If Not GuildName = Norte And NumberNpc = NpcFortaleza Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clan le falta el castillo Norte por conquistar!!" & FONTTYPE_INFO)
                    Exit Function

                End If

                If Not GuildName = Oeste And NumberNpc = NpcFortaleza Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clan le falta el castillo Oeste por conquistar!!" & FONTTYPE_INFO)
                    Exit Function

                End If

                If Not GuildName = Este And NumberNpc = NpcFortaleza Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clan le falta el castillo Este por conquistar!!" & FONTTYPE_INFO)
                    Exit Function

                End If

                If Not GuildName = Sur And NumberNpc = NpcFortaleza Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clan le falta el castillo Sur por conquistar!!" & FONTTYPE_INFO)
                    Exit Function

                End If

            End Select

        End If

    End With

    GolpeNpcCastillo = True

End Function

Public Function HechizoNpcCastillo(ByVal UserIndex As Integer, ByVal h As Integer) As Boolean

    Dim NumberNpc As Integer
    Dim UserMap As Integer
    Dim GuildName As String
    HechizoNpcCastillo = False

    With UserList(UserIndex)

        '  Si no tiene elegido ningun npc, par aque hacer el resto..
        If .flags.TargetNpc < 0 Then Exit Function
        NumberNpc = Npclist(.flags.TargetNpc).Numero

        ' si no es el Rey de castillo o fortaleza, al pedo.
        If NumberNpc = NpcRey Or NumberNpc = NpcFortaleza Then

            UserMap = .pos.Map

            If .GuildIndex = 0 And (UserMap >= CastilloNorte And UserMap <= MapaFortaleza) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes clan!" & FONTTYPE_INFO)
                Exit Function

            End If

            GuildName = Guilds(.GuildIndex).GuildName

            Select Case UserMap

            Case CastilloNorte

                If GuildName = Norte And NumberNpc = NpcRey And Not Hechizos(h).SubeHP = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                    Exit Function

                End If

            Case CastilloSur

                If GuildName = Sur And NumberNpc = NpcRey And Not Hechizos(h).SubeHP = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                    Exit Function

                End If

            Case CastilloEste

                If GuildName = Este And NumberNpc = NpcRey And Not Hechizos(h).SubeHP = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                    Exit Function

                End If

            Case CastilloOeste

                If GuildName = Oeste And NumberNpc = NpcRey And Not Hechizos(h).SubeHP = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                    Exit Function

                End If

            Case MapaFortaleza

                If GuildName = Fortaleza And NumberNpc = NpcFortaleza And Not Hechizos(h).SubeHP = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                    Exit Function

                End If

                If Not GuildName = Norte And NumberNpc = NpcFortaleza Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clan le falta el castillo Norte por conquistar!!" & FONTTYPE_INFO)
                    Exit Function

                End If

                If Not GuildName = Oeste And NumberNpc = NpcFortaleza Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clan le falta el castillo Oeste por conquistar!!" & FONTTYPE_INFO)
                    Exit Function

                End If

                If Not GuildName = Este And NumberNpc = NpcFortaleza Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clan le falta el castillo Este por conquistar!!" & FONTTYPE_INFO)
                    Exit Function

                End If

                If Not GuildName = Sur And NumberNpc = NpcFortaleza Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clan le falta el castillo Sur por conquistar!!" & FONTTYPE_INFO)
                    Exit Function

                End If

            End Select

        End If

    End With

    HechizoNpcCastillo = True

End Function

Public Sub SendInfoCastillos(ByVal UserIndex As Integer)

    If Norte = vbNullString Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Castillo Norte: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Castillo Norte: " & Norte & " " & HoraNorte & FONTTYPE_CONSEJOCAOSVesA)

    End If

    If Sur = vbNullString Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Castillo Sur: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Castillo Sur: " & Sur & " " & HoraSur & FONTTYPE_CONSEJOCAOSVesA)

    End If

    If Oeste = vbNullString Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Castillo Oeste: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Castillo Oeste: " & Oeste & " " & HoraOeste & FONTTYPE_CONSEJOCAOSVesA)

    End If

    If Este = vbNullString Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Castillo Este: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Castillo Este: " & Este & " " & HoraEste & FONTTYPE_CONSEJOCAOSVesA)

    End If

    If Fortaleza = vbNullString Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Fortaleza: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Fortaleza: " & Fortaleza & " " & HoraForta & FONTTYPE_CONSEJOCAOSVesA)

    End If

End Sub

Public Sub CargarCastillos()

    TiempoCura = val(GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "TiempoCura"))
    ExpConquista = val(GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "ExpConquista"))
    OroConquista = val(GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "OroConquista"))
    CuraMinimaRey = val(GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "CuraMinimaRey"))
    CuraMaximaRey = val(GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "CuraMaximaRey"))
    IntervaloRecompensa = val(GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "IntervaloRecompensa"))

    Norte = GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Norte")
    Sur = GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Sur")
    Fortaleza = GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Fortaleza")
    Este = GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Este")
    Oeste = GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Oeste")

    RecompensaCastillo = 0

    HoraSur = GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraSur")
    HoraNorte = GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraNorte")
    HoraOeste = GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraOeste")
    HoraEste = GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraEste")
    HoraForta = GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraForta")

    #If MYSQL = 1 Then
        Call Add_DataBase("0", "Castillos")
    #End If

End Sub

Public Sub AccionNpcCastillos(ByVal NPCNumber As Integer, ByVal UserIndex As Integer)
    Dim NpcPos As WorldPos

    NpcPos.X = 57
    NpcPos.Y = 75

    With UserList(UserIndex)
        Dim UserMap As Integer
        UserMap = .pos.Map

        Select Case NPCNumber

        Case NpcRey

            Select Case UserMap

            Case 98
                Norte = Guilds(.GuildIndex).GuildName
                HoraNorte = now

                Call SendData(SendTarget.ToAll, 0, 0, "||EL CLAN " & UCase$(Norte) & " HA CONQUISTADO EL CASTILLO NORTE." & FONTTYPE_GUILD)

                Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Norte", Norte)
                Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraNorte", HoraNorte)

                NpcPos.Map = 98

                Call SpawnNpc(NpcRey, NpcPos, True, False)

                Call Guilds(.GuildIndex).SetCastleAddPunto(UserIndex)

                #If MYSQL = 1 Then
                    Call Add_DataBase("0", "Castillos")
                #End If

            Case 99

                Sur = Guilds(.GuildIndex).GuildName
                HoraSur = now

                Call SendData(SendTarget.ToAll, 0, 0, "||EL CLAN " & UCase$(Sur) & " HA CONQUISTADO EL CASTILLO SUR." & FONTTYPE_GUILD)

                Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Sur", Sur)
                Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraSur", HoraSur)

                NpcPos.Map = 99

                Call SpawnNpc(NpcRey, NpcPos, True, False)

                Call Guilds(.GuildIndex).SetCastleAddPunto(UserIndex)

                #If MYSQL = 1 Then
                    Call Add_DataBase("0", "Castillos")
                #End If

            Case 100
                Este = Guilds(.GuildIndex).GuildName
                HoraEste = now

                Call SendData(SendTarget.ToAll, 0, 0, "||EL CLAN " & UCase$(Este) & " HA CONQUISTADO EL CASTILLO ESTE." & FONTTYPE_GUILD)

                Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Este", Este)
                Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraEste", HoraEste)

                NpcPos.Map = 100

                Call SpawnNpc(NpcRey, NpcPos, True, False)

                Call Guilds(.GuildIndex).SetCastleAddPunto(UserIndex)

                #If MYSQL = 1 Then
                    Call Add_DataBase("0", "Castillos")
                #End If

            Case 101
                Oeste = Guilds(.GuildIndex).GuildName
                HoraOeste = now

                Call SendData(SendTarget.ToAll, 0, 0, "||EL CLAN " & UCase$(Oeste) & " HA CONQUISTADO EL CASTILLO OESTE." & FONTTYPE_GUILD)

                Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Oeste", Oeste)
                Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraOeste", HoraOeste)

                NpcPos.Map = 101

                Call SpawnNpc(NpcRey, NpcPos, True, False)

                Call Guilds(.GuildIndex).SetCastleAddPunto(UserIndex)

                #If MYSQL = 1 Then
                    Call Add_DataBase("0", "Castillos")
                #End If

            End Select

            Call SendData(SendTarget.ToAll, UserIndex, .pos.Map, "TW44")

        Case NpcFortaleza

            If UserMap = 102 Then

                Fortaleza = Guilds(.GuildIndex).GuildName
                HoraForta = now

                Call SendData(SendTarget.ToAll, 0, 0, "||EL CLAN " & UCase$(Fortaleza) & " HA CONQUISTADO EL CASTILLO FORTALEZA." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToAll, UserIndex, .pos.Map, "TW44")

                Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Fortaleza", Fortaleza)
                Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraForta", HoraForta)

                NpcPos.Map = 102

                Call SpawnNpc(NpcFortaleza, NpcPos, True, False)

                Call Guilds(.GuildIndex).SetCastleAddPunto(UserIndex)

                #If MYSQL = 1 Then
                    Call Add_DataBase("0", "Castillos")
                #End If

            End If

        End Select

    End With

End Sub

