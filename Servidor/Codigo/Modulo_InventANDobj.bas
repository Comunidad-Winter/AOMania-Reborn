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
          Call MakeObj(ToMap, 0, Pos.Map, _
          Obj, Pos.Map, NuevaPos.X, NuevaPos.Y)
          TirarItemAlPiso = NuevaPos
    
    




Exit Function
End If
errhandler:

End Function

Public Sub NPC_TIRAR_ITEMS(ByRef npc As npc, ByVal UserIndex As Integer)

    'TIRA TODOS LOS ITEMS DEL NPC
    'On Error Resume Next
    
    Dim LagaRLDrop

    Dim LagaI

    Dim MiObj        As Obj

    Dim ObjIndex     As Integer

    Dim iSkill       As Integer

    Dim Suerte       As Integer

    Dim Probabilidad As Long

    Dim p            As Double

    If npc.Invent.NroItems > 0 Then

        For LagaI = 1 To npc.Drops.NumDrop
            
            If npc.Invent.Object(LagaI).ObjIndex > 0 Then
                
                MiObj.Amount = npc.Drops.Amount(LagaI)
                MiObj.ObjIndex = npc.Drops.DropIndex(LagaI)
                
                If npc.Drops.Porcentaje(LagaI) <> 0 Then
                
                    Suerte = ((UserList(UserIndex).Stats.UserSkills(eSkill.Suerte) + 5) / 10)
                    p = (100 / npc.Drops.Porcentaje(LagaI))
                    Probabilidad = Int(p)
                 
                    If Suerte = 0 Then Suerte = 1
                    If UserList(UserIndex).Stats.UserSkills(eSkill.Suerte) = 100 Then Suerte = 11
                 
                    Probabilidad = Probabilidad - Porcentaje(Probabilidad, Suerte * 2)
                 
                    Dim mirandom As Double
                
                    mirandom = RandomNumber2(1, Probabilidad)
                
                    Debug.Print "MISUERTE--> " & Suerte & " prob " & Probabilidad & "MIRANDOM ->" & mirandom
                
                    If mirandom = 1 Then
                        Call TirarItemAlPiso(npc.Pos, MiObj)
                    
                        If npc.Drops.Porcentaje(LagaI) < 50 Then Call SubirSkill(UserIndex, Suerte)
                    
                    End If

                Else
                    Call TirarItemAlPiso(npc.Pos, MiObj)

                End If

            End If
        
        Next LagaI
      
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

    On Error Resume Next

    'Devuelve la cantidad original del obj de un npc

    Dim ln As String, npcfile As String

    Dim i  As Integer

    If Npclist(NpcIndex).Numero > 499 Then
        npcfile = DatPath & "NPCs-HOSTILES.dat"
    Else
        npcfile = DatPath & "NPCs.dat"

    End If
 
    For i = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & i)

        If ObjIndex = val(readfield2(1, ln, 45)) Then
            EncontrarCant = val(readfield2(2, ln, 45))
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

On Error Resume Next

Dim ObjIndex As Integer
ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex

    'Quita un Obj                                                          TTT///////////7
    If ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Crucial = 0 Then
        Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount - Cantidad
        
        If Npclist(NpcIndex).Invent.Object(Slot).Amount <= 0 Then
            Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
            Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
            Npclist(NpcIndex).Invent.Object(Slot).Amount = 0
            If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
               Call CargarInvent(NpcIndex) 'Reponemos el inventario
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
               Call CargarInvent(NpcIndex) 'Reponemos el inventario
            End If
        End If
    
    
    
    End If
End Sub


Sub CargarInvent(ByVal NpcIndex As Integer)
On Error Resume Next
'Vuelve a cargar el inventario del npc NpcIndex
Dim LoopC As Integer
Dim ln As String

Dim npcfile As String

If Npclist(NpcIndex).Numero > 499 Then
    npcfile = DatPath & "NPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "NPCs.dat"
End If

Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "NROITEMS"))

For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
    ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & LoopC)
    Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(readfield2(1, ln, 45))
    Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(readfield2(2, ln, 45))
     'TTT////////////////77
       ' Npclist(NpcIndex).Prob(LoopC) = val(Leer.DarValor("NPC" & NpcIndex, "prob" & LoopC))
    'TTT////////////////77
Next LoopC

End Sub
