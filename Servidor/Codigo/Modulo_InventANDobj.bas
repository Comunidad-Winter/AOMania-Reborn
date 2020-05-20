Attribute VB_Name = "InvNpc"

Option Explicit

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Inv & Obj
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Modulo para controlar los objetos y los inventarios.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Public Function TirarItemAlPiso(Pos As WorldPos, Obj As Obj) As WorldPos

    On Error GoTo errhandler

    Dim NuevaPos As WorldPos
    NuevaPos.X = 0
    NuevaPos.Y = 0

    Call Tilelibre(Pos, NuevaPos)

    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
        Call MakeObj(SendTarget.ToMap, 0, Pos.Map, Obj, Pos.Map, NuevaPos.X, NuevaPos.Y)
        TirarItemAlPiso = NuevaPos

    End If

    Exit Function
errhandler:

End Function

Public Sub NPC_TIRAR_ITEMS(ByRef npc As npc, ByVal UserIndex As Integer)

'TIRA TODOS LOS ITEMS DEL NPC
'On Error Resume Next

    With npc

        If .Drops.NumDrop > 0 Then
            Dim LagaRLDrop
            Dim LagaI
            Dim MiObj As Obj
            Dim ObjIndex As Integer
            Dim iSkill As Integer
            Dim Suerte As Integer

            iSkill = UserList(UserIndex).Stats.UserSkills(eSkill.Suerte)

            Select Case iSkill
            Case 0
                Suerte = 0
            Case 1 To 10
                Suerte = 5
            Case 11 To 20
                Suerte = 10
            Case 21 To 30
                Suerte = 15
            Case 31 To 40
                Suerte = 20
            Case 41 To 50
                Suerte = 25
            Case 51 To 60
                Suerte = 30
            Case 61 To 70
                Suerte = 35
            Case 71 To 80
                Suerte = 40
            Case 81 To 99
                Suerte = 45
            Case 100
                Suerte = 50
            End Select


            For LagaI = 1 To npc.Drops.NumDrop

                LagaRLDrop = RandomNumber(1, 100)

                If npc.Drops.Porcentaje(LagaI) = 0 Then

                    MiObj.Amount = npc.Drops.Amount(LagaI)
                    MiObj.ObjIndex = npc.Drops.DropIndex(LagaI)
                    Call TirarItemAlPiso(.Pos, MiObj)

                ElseIf LagaRLDrop <= val(npc.Drops.Porcentaje(LagaI) + Porcentaje(npc.Drops.Porcentaje(LagaI), Suerte)) Then

                    MiObj.Amount = npc.Drops.Amount(LagaI)
                    MiObj.ObjIndex = npc.Drops.DropIndex(LagaI)
                    Call TirarItemAlPiso(.Pos, MiObj)

                End If

            Next LagaI

        End If

    End With

    If RandomNumber(1, 300) = 1 Then
        Call SubirSkill(UserIndex, Suerte)
    End If

End Sub

Function QuedanItems(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Boolean

'On Error Resume Next

'Call LogTarea("Function QuedanItems npcindex:" & NpcIndex & " objindex:" & ObjIndex)

    Dim i As Integer

    If Npclist(NpcIndex).Invent.NroItems > 0 Then

        For i = 1 To MAX_INVENTORY_SLOTS

            If Npclist(NpcIndex).Invent.Object(i).ObjIndex = ObjIndex Then
                QuedanItems = True
                Exit Function

            End If

        Next

    End If

    QuedanItems = False

End Function

Function EncontrarCant(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Integer

'On Error Resume Next

'Devuelve la cantidad original del obj de un npc

    Dim ln As String, npcfile As String
    Dim i As Integer

    npcfile = DatPath & "NPCs.dat"

    For i = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & i)

        If ObjIndex = val(ReadField(1, ln, 45)) Then
            EncontrarCant = val(ReadField(2, ln, 45))
            Exit Function

        End If

    Next

    EncontrarCant = 50

End Function

Sub ResetNpcInv(ByVal NpcIndex As Integer)

'On Error Resume Next

    Dim i As Integer

    Npclist(NpcIndex).Invent.NroItems = 0

    For i = 1 To MAX_INVENTORY_SLOTS
        Npclist(NpcIndex).Invent.Object(i).ObjIndex = 0
        Npclist(NpcIndex).Invent.Object(i).Amount = 0
    Next i

    Npclist(NpcIndex).InvReSpawn = 0

End Sub

Sub QuitarNpcInvItem(ByVal NpcIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)

    Dim ObjIndex As Integer
    ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex

    'Quita un Obj
    If ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Crucial = 0 Then
        Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount - Cantidad

        If Npclist(NpcIndex).Invent.Object(Slot).Amount <= 0 Then
            Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
            Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
            Npclist(NpcIndex).Invent.Object(Slot).Amount = 0

            If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
                Call CargarInvent(NpcIndex)    'Reponemos el inventario

            End If

        End If

    Else
        Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount - Cantidad

        If Npclist(NpcIndex).Invent.Object(Slot).Amount <= 0 Then
            Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
            Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
            Npclist(NpcIndex).Invent.Object(Slot).Amount = 0

            If Not QuedanItems(NpcIndex, ObjIndex) Then

                Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = ObjIndex
                Npclist(NpcIndex).Invent.Object(Slot).Amount = EncontrarCant(NpcIndex, ObjIndex)
                Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1

            End If

            If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
                Call CargarInvent(NpcIndex)    'Reponemos el inventario

            End If

        End If

    End If

End Sub

Sub CargarInvent(ByVal NpcIndex As Integer)

'Vuelve a cargar el inventario del npc NpcIndex
    Dim LoopC As Integer
    Dim ln As String
    Dim npcfile As String

    npcfile = DatPath & "NPCs.dat"

    Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "NROITEMS"))

    For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
        ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & LoopC)
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))

    Next LoopC

End Sub

