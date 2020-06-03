Attribute VB_Name = "mdParty"
Option Explicit

'cantidad maxima de parties en el servidor
Public Const MAX_PARTIES As Integer = 300

''
'nivel minimo para crear party
Public Const MINPARTYLEVEL As Byte = 13

'nivel minimo para aceptar miembro party
Public Const MINACLEVEL As Byte = 7

''
'Cantidad maxima de gente en la party
Public Const PARTY_MAXMEMBERS As Byte = 10

''
'Si esto esta en True, la exp sale por cada golpe que le da
'Si no, la exp la recibe al salirse de la party (pq las partys, floodean)
Public Const PARTY_EXPERIENCIAPORGOLPE As Boolean = False

''
'maxima diferencia de niveles permitida en una party
Public Const MAXPARTYDELTALEVEL As Byte = 10

''
'distancia al leader para que este acepte el ingreso
Public Const MAXDISTANCIAINGRESOPARTY As Byte = 7

''
'maxima distancia a un exito para obtener su experiencia
Public Const PARTY_MAXDISTANCIA As Byte = 18

''
'restan las muertes de los miembros?
Public Const CASTIGOS As Boolean = False

Public Type tPartyMember

    UserIndex As Integer
    Experiencia As Long

End Type

Public Function NextParty() As Integer

    Dim i As Integer
    NextParty = -1

    For i = 1 To MAX_PARTIES

        If Parties(i) Is Nothing Then
            NextParty = i
            Exit Function

        End If

    Next i

End Function

Public Function PuedeCrearParty(ByVal UserIndex As Integer) As Boolean

    PuedeCrearParty = True

    '    If UserList(UserIndex).Stats.ELV < MINPARTYLEVEL Then
    If UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) <= 14 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Tu carisma no es suficiente para liderar una party." & FONTTYPE_PARTY)
        PuedeCrearParty = False
    ElseIf UserList(UserIndex).flags.Muerto = 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Estás muerto!" & FONTTYPE_PARTY)
        PuedeCrearParty = False
    ElseIf UserList(UserIndex).PartyIndex > 0 Then
        PuedeCrearParty = False
    End If

End Function

Public Sub CrearParty(ByVal UserIndex As Integer)

    Dim tInt As Integer

    If UserList(UserIndex).PartyIndex = 0 Then
        If UserList(UserIndex).flags.Muerto = 0 Then
            If UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) >= 0 Then
                tInt = mdParty.NextParty

                If tInt = -1 Then
                    'Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Por el momento no se pueden crear mas partys" & FONTTYPE_PARTY)
                    Exit Sub
                Else
                    Set Parties(tInt) = New clsParty

                    If Not Parties(tInt).NuevoMiembro(UserIndex) Then
                        ' Call SendData(SendTarget.toIndex, UserIndex, 0, "|| La party está llena, no puedes entrar" & FONTTYPE_PARTY)
                        Set Parties(tInt) = Nothing
                        Exit Sub
                    Else
                        'Call SendData(SendTarget.toIndex, UserIndex, 0, "|| ¡ Has formado una party !" & FONTTYPE_PARTY)
                        UserList(UserIndex).PartyIndex = tInt
                        UserList(UserIndex).PartySolicitud = 0
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "NOPRT" & UserList(UserIndex).char.CharIndex & "," & UserList(UserIndex).PartyIndex)
                        Call UpdatePartyMap(UserIndex)

                        If Not Parties(tInt).HacerLeader(UserIndex) Then
                            'Call SendData(SendTarget.toIndex, UserIndex, 0, "|| No puedes hacerte líder." & FONTTYPE_PARTY)
                        Else
                            'Call SendData(SendTarget.toIndex, UserIndex, 0, "|| ¡ Te has convertido en líder de la party !" & FONTTYPE_PARTY)

                            Call EnviarHP(UserIndex)

                        End If

                    End If

                End If

            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| La party están deshabilitada." & _
                                                                FONTTYPE_PARTY)

            End If

        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Estás muerto!" & FONTTYPE_PARTY)

        End If

    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Ya perteneces a una party." & FONTTYPE_PARTY)

    End If

End Sub

Public Sub ResetPartyInfo(ByVal UserIndex As Integer)
        
        With UserList(UserIndex)
             
             .PartySolicitud = 0
             .PartyIndex = 0
             
        End With
        
End Sub

Public Sub CerrarParty(ByVal PI As Integer)
     
     Dim i As Integer
     
     Call Parties(PI).MandarMensajeAConsola("La party ha sido cerrada.", "Servidor")
     
     For i = 1 To PARTY_MAXMEMBERS
        If Parties(PI).IDMember(i) > 0 Then
             UserList(Parties(PI).IDMember(i)).PartyIndex = 0
             Call SendData(SendTarget.ToIndex, Parties(PI).IDMember(i), 0, "NOPRT" & UserList(Parties(PI).IDMember(i)).char.CharIndex & "," & UserList(Parties(PI).IDMember(i)).PartyIndex)
             Call Parties(PI).DeleteMember(i)
         End If
     Next i
     
     Parties(i).p_CantMiembros = 0
     
End Sub

Public Sub SalirDeParty(ByVal UserIndex As Integer)

    Dim PI As Integer
    Dim LoopC As Integer
    
    If UserList(UserIndex).PartyIndex = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No perteneces a ninguna party!!" & FONTTYPE_INFO)
         Exit Sub
    End If
    
    PI = UserList(UserIndex).PartyIndex
    
    If Parties(PI).EsPartyLeader(UserIndex) = True Then
        Call CerrarParty(PI)
    End If
    
    Call Parties(PI).MandarMensajeAConsola(UserList(UserIndex).Name & " ha salido de la party.", "Servidor")

    Parties(PI).p_CantMiembros = Parties(PI).p_CantMiembros - 1
    
    If Parties(PI).CantMiembros < 2 Then
        Call CerrarParty(PI)
        
        Else
        
        Call ResetPartyInfo(UserIndex)
        
        For LoopC = 1 To PARTY_MAXMEMBERS
        
           If UserList(Parties(PI).IDMember(LoopC)).Name = UserList(UserIndex).Name Then
                Call Parties(PI).DeleteMember(LoopC)
                Call SendData(ToAll, 0, 0, "NOPRT" & UserList(UserIndex).char.CharIndex & "," & UserList(UserIndex).PartyIndex)
           End If
        
        Next LoopC
        
    End If
    
End Sub

Public Sub ExpulsarDeParty(ByVal Leader As Integer, ByVal OldMember As Integer)

    Dim PI As Integer
    Dim razon As String
    PI = UserList(Leader).PartyIndex

    If PI > 0 Then
        If PI = UserList(OldMember).PartyIndex Then
            If Parties(PI).EsPartyLeader(Leader) Then
                If Parties(PI).SaleMiembro(OldMember) Then
                    Set Parties(PI) = Nothing
                Else
                    If Parties(PI).CantMiembros < 2 Then
                        Set Parties(PI) = Nothing
                    End If

                    UserList(OldMember).PartyIndex = 0

                End If

            Else
                Call SendData(SendTarget.ToIndex, Leader, 0, "|| Solo el fundador puede expulsar miembros de una party." & FONTTYPE_INFO)

            End If

        Else
            Call SendData(SendTarget.ToIndex, Leader, 0, "|| " & UserList(OldMember).Name & " no pertenece a tu party." & FONTTYPE_INFO)

        End If

    Else
        Call SendData(SendTarget.ToIndex, Leader, 0, "|| No eres miembro de ninguna party." & FONTTYPE_INFO)

    End If

End Sub

Public Sub BroadCastParty(ByVal UserIndex As Integer, ByRef texto As String)

    Dim PI As Integer

    PI = UserList(UserIndex).PartyIndex

    If PI > 0 Then
        Call Parties(PI).MandarMensajeAConsola(texto, UserList(UserIndex).Name)

    End If

End Sub

Public Sub OnlineParty(ByVal UserIndex As Integer)

    Dim PI As Integer
    Dim texto As String

    PI = UserList(UserIndex).PartyIndex

    If PI > 0 Then
        Call Parties(PI).ObtenerMiembrosOnline(texto)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & texto & FONTTYPE_PARTY)

    End If

End Sub

Public Sub TransformarEnLider(ByVal OldLeader As Integer, ByVal NewLeader As Integer)

    Dim PI As Integer

    If OldLeader = NewLeader Then Exit Sub

    PI = UserList(OldLeader).PartyIndex

    If PI > 0 Then
        If PI = UserList(NewLeader).PartyIndex Then
            If UserList(NewLeader).flags.Muerto = 0 Then
                If Parties(PI).EsPartyLeader(OldLeader) Then
                    If Parties(PI).HacerLeader(NewLeader) Then
                        Call Parties(PI).MandarMensajeAConsola("El nuevo líder de la party es " & UserList(NewLeader).Name, UserList(OldLeader).Name)
                    Else
                        Call SendData(SendTarget.ToIndex, OldLeader, 0, "||¡No se ha hecho el cambio de mando!" & FONTTYPE_PARTY)

                    End If

                Else
                    Call SendData(SendTarget.ToIndex, OldLeader, 0, "||¡No eres el líder!" & FONTTYPE_PARTY)

                End If

            Else
                Call SendData(SendTarget.ToIndex, OldLeader, 0, "||¡Está muerto!" & FONTTYPE_INFO)

            End If

        Else
            Call SendData(SendTarget.ToIndex, OldLeader, 0, "||" & UserList(NewLeader).Name & " no pertenece a tu party." & FONTTYPE_INFO)

        End If

    End If

End Sub

Public Sub ObtenerExito(ByVal UserIndex As Integer, ByVal Exp As Long, Mapa As Integer, X As Integer, Y As Integer)

    If Exp <= 0 Then
        If Not CASTIGOS Then Exit Sub

    End If

    Call Parties(UserList(UserIndex).PartyIndex).ObtenerExito(Exp, Mapa, X, Y)

End Sub

Public Function CantMiembros(ByVal UserIndex As Integer) As Integer

    CantMiembros = 0

    If UserList(UserIndex).PartyIndex > 0 Then
        CantMiembros = Parties(UserList(UserIndex).PartyIndex).CantMiembros

    End If

End Function


'--------> A PARTIR DE AQUÍ EMPIEZA PARTY BY BASSINGER

'New sub por Bassinger
Public Sub EnviarParty(ByVal UserIndex As Integer, ByVal UserRecibe As Integer)

    Dim tInt As Integer
    Dim tParty As Integer

    tParty = UserList(UserIndex).PartyIndex

    If UserList(UserIndex).flags.Muerto = 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡ Estás Muerto !!" & FONTTYPE_INFO)
        Exit Sub
    End If

    If UserList(UserRecibe).PartyIndex > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El otro usuario ya esta en party!!." & FONTTYPE_INFO)
        Exit Sub
    End If

    If Not Parties(tParty).EsPartyLeader(UserIndex) Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya estás en un party. Para salir debes ir al boton party y clickear en salir de la party." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UserList(UserRecibe).PartySolicitud = tParty Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes esperar a que el usuario al que ya le has propuesto acepte o cancele la propuesta." & FONTTYPE_INFO)
        Exit Sub
    End If



    UserList(UserRecibe).PartySolicitud = tParty
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Le has ofrecido party a " & UserList(UserRecibe).Name & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, UserRecibe, 0, "||" & UserList(UserIndex).Name & " te a ofrecido party, si deseas continuar escribe /ACEPTAR de otro modo /CANCELAR." & FONTTYPE_TALKMSG)

End Sub

Public Sub AceptarParty(ByVal UserIndex As Integer)

    Dim tParty As Integer
    Dim razon As String
    Dim LoopC As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Map As Integer


    If UserList(UserIndex).PartySolicitud = 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No te han ofrecido ninguna party." & FONTTYPE_INFO)
        Exit Sub
    End If

    tParty = UserList(UserIndex).PartySolicitud

    If Parties(tParty).PuedeEntrar(UserIndex, razon) Then
        If Parties(tParty).NuevoMiembro(UserIndex) Then
            UserList(UserIndex).PartyIndex = tParty
            UserList(UserIndex).PartySolicitud = 0

            Call BroadCastParty(UserIndex, UserList(UserIndex).Name & " se a unido a la party.")

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "NOPRT" & UserList(UserIndex).char.CharIndex & "," & UserList(UserIndex).PartyIndex)

            Parties(tParty).UpdateUserParty

            Call EnviarHP(UserIndex)
        End If
    End If

End Sub

Public Sub CancelarParty(ByVal UserIndex As Integer)

    Dim Leader As Integer
    Dim PI As Integer

    PI = UserList(UserIndex).PartySolicitud

    If UserList(UserIndex).PartyIndex > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Para abandonar una party debes ir a el botón party." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UserList(UserIndex).PartySolicitud > 0 Then

        Leader = Parties(PI).IndexLeader(UserIndex)

        Call SendData(SendTarget.ToIndex, Leader, 0, "||" & UserList(UserIndex).Name & " ha rechazado tu proposición." & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has rechazado la proposición de una party." & FONTTYPE_INFO)
        UserList(UserIndex).PartySolicitud = 0

        If Parties(PI).CantMiembros < "2" Then
            Set Parties(PI) = Nothing
            UserList(Leader).PartyIndex = 0
        End If

    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No te han ofrecido ninguna party." & FONTTYPE_INFO)
        Exit Sub
    End If

End Sub

Sub UpdatePartyMap(ByVal UserIndex As Integer)

    Dim Map As Integer
    Dim X As Integer
    Dim Y As Integer

    Map = UserList(UserIndex).Pos.Map

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(Map, X, Y).UserIndex > 0 And UserIndex <> MapData(Map, X, Y).UserIndex Then
                Call MakeUserChar(SendTarget.ToIndex, UserIndex, 0, MapData(Map, X, Y).UserIndex, Map, X, Y)

                If UserList(MapData(Map, X, Y).UserIndex).flags.Invisible = 1 Or UserList(MapData(Map, X, Y).UserIndex).flags.Oculto = 1 Then Call _
                   SendData(SendTarget.ToIndex, UserIndex, 0, "NOVER" & UserList(MapData(Map, X, Y).UserIndex).char.CharIndex & ",1," & _
                                                              UserList(MapData(Map, X, Y).UserIndex).PartyIndex)

            End If
        Next X
    Next Y
End Sub

Sub VerParty(ByVal UserIndex As Integer)
    Dim Rs As Integer

    Rs = UserList(UserIndex).PartyIndex

    If Rs = 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "VPA" & Rs)
    ElseIf Rs > 0 Then
        Parties(Rs).ObtenerVerParty (UserIndex)
    End If

End Sub
