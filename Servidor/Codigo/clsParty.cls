VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_members(1 To PARTY_MAXMEMBERS) As tPartyMember
'miembros

Private p_expTotal As Long
'Estadistica :D

Private p_Fundador As Integer
'el creador

Private p_CantMiembros As Integer
'cantidad de miembros

'datos en los pjs: | indexParty(indice en p_members), partyLeader(userindex del lider) |

'Constructor de clase
Public Sub Class_Initialize()

    p_expTotal = 0
    p_CantMiembros = 0

End Sub

'Destructor de clase
Public Sub Class_Terminate()

End Sub

Public Function MiExperiencia(ByVal UserIndex As Integer) As Long

'Me dice cuanta experiencia tengo colectada ya en la party
    Dim i As Integer
    i = 0

    While i <= PARTY_MAXMEMBERS And p_members(i).UserIndex <> UserIndex

        i = i + 1
    Wend

    If i <= PARTY_MAXMEMBERS Then
        MiExperiencia = p_members(i).Experiencia
    Else        'esto no deberia pasar :p
        MiExperiencia = -1

    End If

End Function

Public Sub ObtenerExito(ByVal ExpGanada As Long, ByVal Mapa As Integer, X As Integer, Y As Integer)
'Se produjo un evento que da experiencia en la wp referenciada
    Dim i As Integer
    Dim UI As Integer
    Dim expPerCapita As Double
    Dim CantidadDistancia As Byte

    For i = 1 To PARTY_MAXMEMBERS
        UI = p_members(i).UserIndex
        If UI > 0 Then
            If Mapa = UserList(UI).pos.Map And UserList(UI).flags.Muerto = 0 Then
                If Distance(UserList(UI).pos.X, UserList(UI).pos.Y, X, Y) <= PARTY_MAXDISTANCIA Then
                    CantidadDistancia = CantidadDistancia + 1
                End If
            End If
        End If
    Next i

    expPerCapita = CLng((ExpGanada + Porcentaje(ExpGanada, (3 * CantidadDistancia))) / CantidadDistancia)

    For i = 1 To PARTY_MAXMEMBERS
        UI = p_members(i).UserIndex
        If UI > 0 Then
            If Mapa = UserList(UI).pos.Map And UserList(UI).flags.Muerto = 0 Then
                If Distance(UserList(UI).pos.X, UserList(UI).pos.Y, X, Y) <= PARTY_MAXDISTANCIA Then
                    UserList(p_members(i).UserIndex).Stats.Exp = UserList(p_members(i).UserIndex).Stats.Exp + expPerCapita
                    Call SendData(SendTarget.ToIndex, (UI), 0, "||¡Has ganado " & expPerCapita & " puntos de experiencia!" & FONTTYPE_Motd4)
                    ' p_members(i).Experiencia = p_members(i).Experiencia + expPerCapita
                    If expPerCapita < 0 Then
                        expPerCapita = 0
                    End If
                    If UserList(UI).Stats.Exp > MAXEXP Then _
                       UserList(UI).Stats.Exp = MAXEXP
                    Call CheckUserLevel(UI)
                    Call SendUserStatsBox(UI)    'Experiencia obviamente xd
                    If PARTY_EXPERIENCIAPORGOLPE Then
                        ' UserList(UI).Stats.Exp = UserList(UI).Stats.Exp + expPerCapita
                        UserList(p_members(i).UserIndex).Stats.Exp = UserList(p_members(i).UserIndex).Stats.Exp + expPerCapita
                        Call SendData(SendTarget.ToIndex, (UI), 0, "||Has ganado " & expPerCapita & " puntos de experiencia." & FONTTYPE_AMARILLON)
                        If UserList(UI).Stats.Exp > MAXEXP Then _
                           UserList(UI).Stats.Exp = MAXEXP
                        Call CheckUserLevel(UI)
                        Call SendUserStatsBox(UI)    ' lo mismo, exp
                    End If
                End If
            End If
        End If
    Next i

End Sub
Public Sub MandarMensajeAConsola(ByVal texto As String, ByVal Sender As String)

'feo feo, muy feo acceder a senddata desde aca, pero BUEEEEEEEEEEE...
    Dim i As Integer

    For i = 1 To PARTY_MAXMEMBERS

        If p_members(i).UserIndex > 0 Then
            Call SendData(SendTarget.ToIndex, p_members(i).UserIndex, 0, "||" & texto & FONTTYPE_TALKMSG)

        End If

    Next i

End Sub

Public Sub EnviarBarra()
     
     Dim i As Integer
     
     For i = 1 To PARTY_MAXMEMBERS
          
          If p_members(i).UserIndex > 0 Then
             Call SendData(SendTarget.ToPartyArea, p_members(i).UserIndex, UserList(p_members(i).UserIndex).pos.Map, "VPT" & UserList(p_members(i).UserIndex).char.CharIndex & "," & UserList(p_members(i).UserIndex).Stats.MinHP & "," & UserList(p_members(i).UserIndex).Stats.MaxHP & "," & UserList(p_members(i).UserIndex).PartyIndex)
                
          End If
          
     Next i
     
End Sub

Public Function EsPartyLeader(ByVal UserIndex As Integer) As Boolean

    EsPartyLeader = (UserIndex = p_Fundador)

End Function

Public Function IndexLeader(ByVal UserIndex As Integer)
    IndexLeader = p_Fundador
End Function

Public Function NuevoMiembro(ByVal UserIndex As Integer) As Boolean

    Dim i As Integer
    i = 1

    While i <= PARTY_MAXMEMBERS And p_members(i).UserIndex > 0

        i = i + 1
    Wend

    If i <= PARTY_MAXMEMBERS Then
        p_members(i).Experiencia = 0
        p_members(i).UserIndex = UserIndex
        NuevoMiembro = True
        p_CantMiembros = p_CantMiembros + 1
    Else
        NuevoMiembro = False

    End If

End Function

Public Function SaleMiembro(ByVal UserIndex As Integer) As Boolean

    Dim i As Integer
    Dim j As Integer
    Dim Leader As Integer

    i = 1
    SaleMiembro = False
    Leader = Parties(UserList(i).PartyIndex).IndexLeader(p_members(i).UserIndex)

    While i <= PARTY_MAXMEMBERS And p_members(i).UserIndex <> UserIndex
        i = i + 1
    Wend

    If i = 1 Then
        'sale el founder, la party se disuelve
        SaleMiembro = True

        For j = PARTY_MAXMEMBERS To 1 Step -1
            If p_members(j).UserIndex > 0 Then

                If j > 1 Then
                    Call MandarMensajeAConsola(UserList(p_members(j).UserIndex).Name & " ha salido de la party.", "Servidor")
                Else
                    Call MandarMensajeAConsola("La party ha sido cerrada.", "Servidor")
                    Call EnviarHP(Leader)
                    Call SendData(SendTarget.ToIndex, Leader, 0, "NOPRT" & UserList(Leader).char.CharIndex & "," & UserList(Leader).PartyIndex)

                End If
                UserList(p_members(j).UserIndex).PartyIndex = 0
                Call EnviarHP(p_members(j).UserIndex)
                Call SendData(SendTarget.ToIndex, p_members(j).UserIndex, 0, "NOPRT" & UserList(p_members(j).UserIndex).char.CharIndex & "," & UserList(p_members(j).UserIndex).PartyIndex)

                p_CantMiembros = p_CantMiembros - 1
                p_members(j).UserIndex = 0
            End If
        Next j

    Else

        If i <= PARTY_MAXMEMBERS Then

            Call MandarMensajeAConsola(UserList(p_members(i).UserIndex).Name & " ha salido de la party.", "Servidor")

            UserList(i).PartyIndex = 0
            p_CantMiembros = p_CantMiembros - 1
            Call EnviarHP(p_members(i).UserIndex)
            Call SendData(SendTarget.ToIndex, p_members(i).UserIndex, 0, "NOPRT" & UserList(p_members(i).UserIndex).char.CharIndex & "," & UserList(p_members(i).UserIndex).PartyIndex)

            p_members(i).UserIndex = 0
            CompactMemberList

            If p_CantMiembros < 2 Then
                Call MandarMensajeAConsola("La party ha sido cerrada.", "Servidor")
                UserList(Leader).PartyIndex = 0
                Call EnviarHP(Leader)
                Call SendData(SendTarget.ToIndex, Leader, 0, "NOPRT" & UserList(Leader).char.CharIndex & "," & UserList(Leader).PartyIndex)

                p_CantMiembros = p_CantMiembros - 1
                p_members("1").UserIndex = 0
            End If
        End If

    End If

End Function
Public Function HacerLeader(ByVal UserIndex As Integer) As Boolean

    Dim i As Integer
    Dim OldLeader As Integer
    Dim oldExp As Long
    Dim UserIndexIndex As Integer

    UserIndexIndex = 0
    HacerLeader = True

    For i = 1 To PARTY_MAXMEMBERS

        If p_members(i).UserIndex > 0 Then
            If p_members(i).UserIndex = UserIndex Then
                UserIndexIndex = i
                Exit For

            End If

            'chequeo la dif entre levels
            If Abs(UserList(UserIndex).Stats.ELV - UserList(p_members(i).UserIndex).Stats.ELV) > MAXPARTYDELTALEVEL Then
                Call SendData(SendTarget.ToIndex, p_members(i).UserIndex, 0, "|| No se puede hacer líder a " & UserList(UserIndex).Name & _
                                                                           " pues la diferencia entre niveles de los miembros de la party sería mayor a " & MAXPARTYDELTALEVEL & FONTTYPE_PARTY)
                HacerLeader = False

            End If

        End If

    Next i

    If Not HacerLeader Then Exit Function

    If UserIndexIndex = 0 Then
        'catastrofe! esto no deberia pasar nunca! pero como es AO.... :p
        Call LogError("INCONSISTENCIA DE PARTIES")
        Call SendData(SendTarget.ToAdmins, 0, 0, _
                      "|| Inconsistencia de parties en HACERLEADER (UII = 0), AVISE A UN PROGRAMADOR ESTO ES UNA CATASTROFE!!!!")
        HacerLeader = False
        Exit Function

    End If

    'aca esta todo bien y doy vuelta las collections
    OldLeader = p_members(1).UserIndex
    oldExp = p_members(1).Experiencia

    p_members(1).UserIndex = p_members(UserIndexIndex).UserIndex     'que en realdiad es el userindex, pero no quiero inconsistencias moviendo experiencias
    p_members(1).Experiencia = p_members(UserIndexIndex).Experiencia

    p_members(i).UserIndex = OldLeader
    p_members(i).Experiencia = oldExp

    p_Fundador = p_members(1).UserIndex

    'no need to compact
End Function

Public Sub UpdateUserParty()

    Dim i As Integer

    For i = 1 To PARTY_MAXMEMBERS

        If p_members(i).UserIndex > 0 Then
            Call UpdatePartyMap(p_members(i).UserIndex)
            Call EnviarHP(p_members(i).UserIndex)
        End If


    Next i

End Sub

Public Sub ObtenerMiembrosOnline(ByRef MemberList As String)

'esta funcion viola un poco la correcta forma de hacer la clase
'pero bue...
    Dim i As Integer
    MemberList = "OnlineParty: "

    For i = 1 To PARTY_MAXMEMBERS

        If p_members(i).UserIndex > 0 Then

            If i = 1 Then
                MemberList = MemberList & UserList(p_members(i).UserIndex).Name
            Else
                MemberList = MemberList & ", " & UserList(p_members(i).UserIndex).Name
            End If

        End If

    Next i

    MemberList = MemberList & "." & FONTTYPE_INFO

End Sub

Public Sub ObtenerVerParty(ByVal IndexU As Integer)

    Dim i As Integer
    Dim LoopC As Integer
    Dim NameUser As String
    Dim MinHpUser As Integer
    Dim MaxHpUser As Integer

    For LoopC = 1 To PARTY_MAXMEMBERS

        If p_members(LoopC).UserIndex > 0 Then

            i = i + 1
            NameUser = UserList(p_members(LoopC).UserIndex).Name
            MinHpUser = UserList(p_members(LoopC).UserIndex).Stats.MinHP
            MaxHpUser = UserList(p_members(LoopC).UserIndex).Stats.MaxHP

            Call SendData(SendTarget.ToIndex, IndexU, 0, "IVP" & i & "," & NameUser & "," & MinHpUser & "," & MaxHpUser)


        End If

    Next LoopC

    Call SendData(SendTarget.ToIndex, IndexU, 0, "VPA" & "1")

End Sub


Public Function PuedeEntrar(ByVal UserIndex As Integer, ByRef razon As String) As Boolean

'DEFINE LAS REGLAS DEL JUEGO PARA DEJAR ENTRAR A MIEMBROS
    Dim esArmada As Boolean
    Dim esCaos As Boolean
    Dim MyLevel As Integer
    Dim i As Integer
    Dim rv As Boolean
    Dim UI As Integer

    rv = True
    esArmada = (UserList(UserIndex).Faccion.ArmadaReal = 1) Or (UserList(UserIndex).Faccion.Templario = 1)
    esCaos = (UserList(UserIndex).Faccion.FuerzasCaos = 1) Or (UserList(UserIndex).Faccion.Nemesis = 1)
    MyLevel = UserList(UserIndex).Stats.ELV

    rv = Distancia(UserList(p_members(1).UserIndex).pos, UserList(UserIndex).pos) <= MAXDISTANCIAINGRESOPARTY

    If rv Then
        rv = (p_members(PARTY_MAXMEMBERS).UserIndex = 0)

        If rv Then

            For i = 1 To PARTY_MAXMEMBERS
                UI = p_members(i).UserIndex

                'pongo los casos que evitarian que pueda entrar
                'aspirante armada en party crimi
                If UI > 0 Then


                    'diferencia entre niveles
                    If Abs(UserList(UI).Stats.ELV - UserList(UserIndex).Stats.ELV) > MAXPARTYDELTALEVEL Then
                        razon = "La diferencia con " & UserList(UI).Name & " es mayor a " & MAXPARTYDELTALEVEL & " niveles."
                        rv = False

                    End If

                    If Not rv Then Exit For    'violate una programacion estructurada

                End If

            Next i

        Else
            razon = "La mayor cantidad de miembros es " & PARTY_MAXMEMBERS

        End If

    Else
        razon = "Te encuentras muy lejos del fundador."

    End If

    PuedeEntrar = rv

End Function


Private Sub CompactMemberList()

    Dim i As Integer
    Dim freeIndex As Integer
    Dim tPlayer As tPartyMember
    i = 1

    While i <= PARTY_MAXMEMBERS

        If p_members(i).UserIndex = 0 And freeIndex = 0 Then
            freeIndex = i
        ElseIf p_members(i).UserIndex > 0 And freeIndex > 0 Then
            p_members(freeIndex).Experiencia = p_members(i).Experiencia
            p_members(freeIndex).UserIndex = p_members(i).UserIndex
            p_members(i).UserIndex = 0
            p_members(i).Experiencia = 0
            'muevo el de la pos i a freeindex
            i = freeIndex
            freeIndex = 0

        End If

        i = i + 1
    Wend

End Sub

Public Function CantMiembros() As Integer

    CantMiembros = p_CantMiembros

End Function
