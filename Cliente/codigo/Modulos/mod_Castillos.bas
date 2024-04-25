Attribute VB_Name = "mod_Castillos"
' Encapsulo el sistema de castillos que tenian aca
Option Explicit

Private Const NpcRey        As Integer = 657
Private Const NpcFortaleza  As Integer = 663

Private Const CastilloNorte As Integer = 98
Private Const CastilloSur   As Integer = 99
Private Const CastilloEste  As Integer = 100
Private Const CastilloOeste As Integer = 101
Private Const MapaFortaleza As Integer = 102

Private TiempoCura          As Integer
Private CuraMinimaRey       As Integer
Private CuraMaximaRey       As Integer
Private ExpConquista        As Integer
Private OroConquista        As Integer

Private RecompensaNorte     As Integer
Private RecompensaSur       As Integer
Private RecompensaOeste     As Integer
Private RecompensaEste      As Integer
Private RecompensaFortaleza As Integer

Public Norte               As String
Public Sur                 As String
Public Oeste               As String
Public Este                As String
Public Fortaleza           As String

Private HoraSur             As String
Private HoraNorte           As String
Private HoraEste            As String
Private HoraOeste           As String
Private HoraForta           As String

Public Sub RecompensaCastillos()
    Dim GuildIndex As Integer

    RecompensaNorte = RecompensaNorte + 1
    RecompensaSur = RecompensaSur + 1
    RecompensaOeste = RecompensaOeste + 1
    RecompensaEste = RecompensaEste + 1
    RecompensaFortaleza = RecompensaFortaleza + 1

    If RecompensaNorte >= 120 Then
    
        If Len(Norte) > 0 Then
    
            Call SendData(SendTarget.toall, 0, 0, "||Recompensa de Castillo Norte para el clan " & Norte & FONTTYPE_CONSEJOCAOSVesA)
                
            GuildIndex = modGuilds.GuildIndex(Norte)
        
            Call modGuilds.RecompensasCastillos(GuildIndex, ExpConquista, OroConquista)
       
        Else
  
            Call SendData(SendTarget.toall, 0, 0, "||Recompensa Castillo Norte: Nadie." & FONTTYPE_CONSEJOCAOSVesA)

        End If

        RecompensaNorte = 0

    End If

    If RecompensaSur >= 120 Then
    
        If Len(Sur) > 0 Then
    
            Call SendData(SendTarget.toall, 0, 0, "||Recompensa de Castillo Sur para el clan " & Sur & FONTTYPE_CONSEJOCAOSVesA)
                
            GuildIndex = modGuilds.GuildIndex(Sur)
        
            Call modGuilds.RecompensasCastillos(GuildIndex, ExpConquista, OroConquista)
       
        Else
  
            Call SendData(SendTarget.toall, 0, 0, "||Recompensa Castillo Sur: Nadie." & FONTTYPE_CONSEJOCAOSVesA)

        End If

        RecompensaSur = 0

    End If

    If RecompensaOeste >= 120 Then
    
        If Len(Oeste) > 0 Then
    
            Call SendData(SendTarget.toall, 0, 0, "||Recompensa de Castillo Oeste para el clan " & Oeste & FONTTYPE_CONSEJOCAOSVesA)
                
            GuildIndex = modGuilds.GuildIndex(Oeste)
        
            Call modGuilds.RecompensasCastillos(GuildIndex, ExpConquista, OroConquista)
       
        Else
  
            Call SendData(SendTarget.toall, 0, 0, "||Recompensa Castillo Oeste: Nadie." & FONTTYPE_CONSEJOCAOSVesA)

        End If

        RecompensaOeste = 0

    End If
    
    If RecompensaEste >= 120 Then
    
        If Len(Este) > 0 Then
    
            Call SendData(SendTarget.toall, 0, 0, "||Recompensa de Castillo Este para el clan " & Este & FONTTYPE_CONSEJOCAOSVesA)
                
            GuildIndex = modGuilds.GuildIndex(Este)
        
            Call modGuilds.RecompensasCastillos(GuildIndex, ExpConquista, OroConquista)
       
        Else
  
            Call SendData(SendTarget.toall, 0, 0, "||Recompensa Castillo Este: Nadie." & FONTTYPE_CONSEJOCAOSVesA)

        End If

        RecompensaEste = 0

    End If

    If RecompensaFortaleza >= 120 Then
    
        If Len(Fortaleza) > 0 Then
    
            Call SendData(SendTarget.toall, 0, 0, "||Recompensa de Fortaleza para el clan " & Fortaleza & FONTTYPE_CONSEJOCAOSVesA)
                
            GuildIndex = modGuilds.GuildIndex(Fortaleza)
        
            Call modGuilds.RecompensasCastillos(GuildIndex, ExpConquista, OroConquista)
       
        Else
  
            Call SendData(SendTarget.toall, 0, 0, "||Recompensa de Fortaleza: Nadie." & FONTTYPE_CONSEJOCAOSVesA)

        End If

        RecompensaFortaleza = 0

    End If

End Sub

Public Sub CuraRey(ByVal NpcIndex As Integer)
    Static Tiempo As Integer
    
    Tiempo = Tiempo + 1
    
    If Tiempo >= TiempoCura Then
    
        With Npclist(NpcIndex)

            If .Numero = NpcRey Or .Numero = NpcFortaleza Then
        
                If .Stats.MinHP > 5000 And Not .Stats.MinHP = .Stats.MaxHP Then
                    .Stats.MinHP = .Stats.MinHP + RandomNumber(CuraMinimaRey, CuraMaximaRey)

                    If .Stats.MinHP > .Stats.MaxHP Then .Stats.MinHP = .Stats.MaxHP

                End If

            End If

        End With
        
        Tiempo = 0

    End If

End Sub

Public Sub WarpCastillo(ByVal UserIndex As Integer, ByVal Castillo As String)

    With UserList(UserIndex)
    
        If .flags.EstaDueleando1 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes defender el castillo estando en DUELOS." & FONTTYPE_WARNING)
            Exit Sub

        End If

        If .flags.Paralizado = 1 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes teletransportarte porque estás afectado por un hechizo que te lo impide." _
                & FONTTYPE_WARNING)
            Exit Sub

        End If
        
        If .Counters.Pena > 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes salir de la cárcel." & FONTTYPE_WARNING)
            Exit Sub

        End If
            
        If .GuildIndex = 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No tienes clan!" & FONTTYPE_INFO)
            Exit Sub

        End If
        
        Dim X         As Integer
        Dim Y         As Integer
        Dim Map       As Integer
        Dim GuildName As String
        Dim WarpUser  As Boolean
        
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
                WarpUser = (StrComp(GuildName, Norte, vbTextCompare) = 0) And (StrComp(GuildName, Sur, vbTextCompare) = 0) And (StrComp(GuildName, _
                    Este, vbTextCompare) = 0) And (StrComp(GuildName, Oeste, vbTextCompare) = 0)

        End Select
            
        If WarpUser Then
            Call WarpUserChar(UserIndex, Map, X, Y, True)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No cumples con los requisitos para ir a este castillo." & FONTTYPE_INFO)

        End If

    End With

End Sub

Public Function GolpeNpcCastillo(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean

    Dim NumberNpc As Integer
    Dim UserMap   As Integer
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
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No tienes clan!" & FONTTYPE_INFO)
                Exit Function

            End If
        
            GuildName = Guilds(.GuildIndex).GuildName

            Select Case UserMap
        
                Case CastilloNorte

                    If GuildName = Norte And NumberNpc = NpcRey Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                        Exit Function

                    End If

                Case CastilloSur

                    If GuildName = Sur And NumberNpc = NpcRey Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                        Exit Function

                    End If

                Case CastilloEste

                    If GuildName = Oeste And NumberNpc = NpcRey Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                        Exit Function

                    End If

                Case CastilloOeste

                    If GuildName = Norte And NumberNpc = NpcRey Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                        Exit Function

                    End If

                Case MapaFortaleza
            
                    If GuildName = Fortaleza And NumberNpc = NpcFortaleza Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                        Exit Function

                    End If
       
                    If Not GuildName = Norte And NumberNpc = NpcFortaleza Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu clan le falta el castillo Norte por conquistar!!" & FONTTYPE_INFO)
                        Exit Function

                    End If

                    If Not GuildName = Oeste And NumberNpc = NpcFortaleza Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu clan le falta el castillo Oeste por conquistar!!" & FONTTYPE_INFO)
                        Exit Function

                    End If

                    If Not GuildName = Este And NumberNpc = NpcFortaleza Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu clan le falta el castillo Este por conquistar!!" & FONTTYPE_INFO)
                        Exit Function

                    End If

                    If Not GuildName = Sur And NumberNpc = NpcFortaleza Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu clan le falta el castillo Sur por conquistar!!" & FONTTYPE_INFO)
                        Exit Function

                    End If
                
            End Select

        End If

    End With

    GolpeNpcCastillo = True

End Function

Public Function HechizoNpcCastillo(ByVal UserIndex As Integer, ByVal H As Integer) As Boolean

    Dim NumberNpc As Integer
    Dim UserMap   As Integer
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
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No tienes clan!" & FONTTYPE_INFO)
                Exit Function

            End If
        
            GuildName = Guilds(.GuildIndex).GuildName

            Select Case UserMap
        
                Case CastilloNorte

                    If GuildName = Norte And NumberNpc = NpcRey And Not Hechizos(H).SubeHP = 1 Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                        Exit Function

                    End If

                Case CastilloSur

                    If GuildName = Sur And NumberNpc = NpcRey And Not Hechizos(H).SubeHP = 1 Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                        Exit Function

                    End If

                Case CastilloEste

                    If GuildName = Oeste And NumberNpc = NpcRey And Not Hechizos(H).SubeHP = 1 Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                        Exit Function

                    End If

                Case CastilloOeste

                    If GuildName = Norte And NumberNpc = NpcRey And Not Hechizos(H).SubeHP = 1 Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                        Exit Function

                    End If

                Case MapaFortaleza
            
                    If GuildName = Fortaleza And NumberNpc = NpcFortaleza And Not Hechizos(H).SubeHP = 1 Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes atacar a tu rey!!" & FONTTYPE_INFO)
                        Exit Function

                    End If

                    If Not GuildName = Norte And NumberNpc = NpcFortaleza Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu clan le falta el castillo Norte por conquistar!!" & FONTTYPE_INFO)
                        Exit Function

                    End If

                    If Not GuildName = Oeste And NumberNpc = NpcFortaleza Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu clan le falta el castillo Oeste por conquistar!!" & FONTTYPE_INFO)
                        Exit Function

                    End If

                    If Not GuildName = Este And NumberNpc = NpcFortaleza Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu clan le falta el castillo Este por conquistar!!" & FONTTYPE_INFO)
                        Exit Function

                    End If

                    If Not GuildName = Sur And NumberNpc = NpcFortaleza Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu clan le falta el castillo Sur por conquistar!!" & FONTTYPE_INFO)
                        Exit Function

                    End If
                
            End Select

        End If

    End With

    HechizoNpcCastillo = True

End Function

Public Sub SendInfoCastillos(ByVal UserIndex As Integer)

    If Norte = vbNullString Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Castillo Norte: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
    Else
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Castillo Norte: " & Norte & " " & HoraNorte & FONTTYPE_CONSEJOCAOSVesA)

    End If

    If Sur = vbNullString Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Castillo Sur: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
    Else
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Castillo Sur: " & Sur & " " & HoraSur & FONTTYPE_CONSEJOCAOSVesA)

    End If

    If Oeste = vbNullString Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Castillo Oeste: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
    Else
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Castillo Oeste: " & Oeste & " " & HoraOeste & FONTTYPE_CONSEJOCAOSVesA)

    End If

    If Este = vbNullString Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Castillo Este: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
    Else
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Castillo Este: " & Este & " " & HoraEste & FONTTYPE_CONSEJOCAOSVesA)

    End If

    If Fortaleza = vbNullString Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Fortaleza: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
    Else
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Fortaleza: " & Fortaleza & " " & HoraForta & FONTTYPE_CONSEJOCAOSVesA)

    End If

End Sub

Public Sub CargarCastillos()

    TiempoCura = val(GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "TiempoCura"))
    ExpConquista = val(GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "ExpConquista"))
    OroConquista = val(GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "OroConquista"))
    CuraMinimaRey = val(GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "CuraMinimaRey"))
    CuraMaximaRey = val(GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "CuraMaximaRey"))
    
    Norte = GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Norte")
    Sur = GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Sur")
    Fortaleza = GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Fortaleza")
    Este = GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Este")
    Oeste = GetVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Oeste")
    
    RecompensaNorte = 0
    RecompensaSur = 0
    RecompensaOeste = 0
    RecompensaEste = 0
    RecompensaFortaleza = 0

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
    NpcPos.Y = 72

    With UserList(UserIndex)
        Dim UserMap As Integer
        UserMap = .pos.Map

        Select Case NPCNumber
            
            Case NpcRey

                Select Case UserMap
                        
                    Case 98
                        Norte = Guilds(.GuildIndex).GuildName
                        HoraNorte = Now
            
                        Call SendData(SendTarget.toall, 0, 0, "||EL CLAN " & Norte & " HA CONQUISTADO EL CASTILLO NORTE." & FONTTYPE_GUILD)
          
                        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Norte", Norte)
                        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraNorte", HoraNorte)
                        
                        NpcPos.Map = 98
                        RecompensaNorte = 0
                        Call SpawnNpc(NpcRey, NpcPos, True, False)
                        
                        Call Guilds(.GuildIndex).SetCastleAddPunto(UserIndex)
                        
#If MYSQL = 1 Then
                        Call Add_DataBase("0", "Castillos")
#End If

                    Case 99
        
                        Sur = Guilds(.GuildIndex).GuildName
                        HoraSur = Now
            
                        Call SendData(SendTarget.toall, 0, 0, "||EL CLAN " & Sur & " HA CONQUISTADO EL CASTILLO SUR." & FONTTYPE_GUILD)
                        
                        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Sur", Sur)
                        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraSur", HoraSur)
                        
                        NpcPos.Map = 99
                        RecompensaSur = 0
                        Call SpawnNpc(NpcRey, NpcPos, True, False)
                        
                        Call Guilds(.GuildIndex).SetCastleAddPunto(UserIndex)
                        
#If MYSQL = 1 Then
                        Call Add_DataBase("0", "Castillos")
#End If
        
                    Case 100
                        Este = Guilds(.GuildIndex).GuildName
                        HoraEste = Now
                        
                        Call SendData(SendTarget.toall, 0, 0, "||EL CLAN " & UCase$(Este) & " HA CONQUISTADO EL CASTILLO ESTE." & FONTTYPE_GUILD)
                     
                        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Este", Este)
                        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraEste", HoraEste)
                        
                        NpcPos.Map = 100
                        RecompensaEste = 0
                        Call SpawnNpc(NpcRey, NpcPos, True, False)
                        
                        Call Guilds(.GuildIndex).SetCastleAddPunto(UserIndex)
                        
#If MYSQL = 1 Then
                        Call Add_DataBase("0", "Castillos")
#End If
           
                    Case 101
                        Oeste = Guilds(.GuildIndex).GuildName
                        HoraOeste = Now
                        
                        Call SendData(SendTarget.toall, 0, 0, "||EL CLAN " & UCase$(Oeste) & " HA CONQUISTADO EL CASTILLO OESTE." & FONTTYPE_GUILD)
                      
                        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Oeste", Oeste)
                        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraOeste", HoraOeste)
                        
                        NpcPos.Map = 101
                        RecompensaOeste = 0
                        Call SpawnNpc(NpcRey, NpcPos, True, False)
                        
                        Call Guilds(.GuildIndex).SetCastleAddPunto(UserIndex)
                        
#If MYSQL = 1 Then
                        Call Add_DataBase("0", "Castillos")
#End If
  
                End Select
                
                Call SendData(SendTarget.toall, UserIndex, .pos.Map, "TW44")
        
            Case NpcFortaleza

                If UserMap = 102 Then
        
                    Fortaleza = Guilds(.GuildIndex).GuildName
                    HoraForta = Now
            
                    Call SendData(SendTarget.toall, 0, 0, "||EL CLAN " & Fortaleza & " HA CONQUISTADO EL CASTILLO FORTALEZA." & FONTTYPE_GUILD)
                    Call SendData(SendTarget.toall, UserIndex, .pos.Map, "TW44")
  
                    Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Fortaleza", Fortaleza)
                    Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraForta", HoraForta)
                   
                    NpcPos.Map = 102
                    RecompensaFortaleza = 0
                    Call SpawnNpc(NpcFortaleza, NpcPos, True, False)
                    
                    Call Guilds(.GuildIndex).SetCastleAddPunto(UserIndex)
                    
#If MYSQL = 1 Then
                    Call Add_DataBase("0", "Castillos")
#End If
                  
                End If
        
        End Select

    End With

End Sub

