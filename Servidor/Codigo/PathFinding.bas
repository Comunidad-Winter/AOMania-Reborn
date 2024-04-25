Attribute VB_Name = "PathFinding"
Option Explicit

Private Const ROWS = 100
Private Const COLUMS = 100
Private Const MAXINT = 1000
Private Const Walkable = 0

Private Type tIntermidiateWork
    Known As Boolean
    DistV As Integer
    PrevV As tVertice
End Type

Dim TmpArray(1 To ROWS, 1 To COLUMS) As tIntermidiateWork

Dim TilePosX As Integer, TilePosY As Integer

Dim MyVert As tVertice
Dim MyFin As tVertice

Dim Iter As Integer

Private Function Limites(ByVal vfila As Integer, ByVal vcolu As Integer)
Limites = vcolu >= 1 And vcolu <= COLUMS And vfila >= 1 And vfila <= ROWS
End Function

Private Function IsWalkable(ByVal Map As Integer, ByVal row As Integer, ByVal Col As Integer, ByVal NpcIndex As Integer) As Boolean
IsWalkable = MapData(Map, row, Col).Blocked = 0 And MapData(Map, row, Col).NpcIndex = 0

If MapData(Map, row, Col).UserIndex <> 0 Then
     If MapData(Map, row, Col).UserIndex <> Npclist(NpcIndex).PFINFO.TargetUser Then IsWalkable = False
End If

End Function

Private Sub ProcessAdjacents(ByVal MapIndex As Integer, ByRef T() As tIntermidiateWork, ByRef vfila As Integer, ByRef vcolu As Integer, ByVal NpcIndex As Integer)
    Dim V As tVertice
    Dim j As Integer
    'Look to North
    j = vfila - 1
    If Limites(j, vcolu) Then
            If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then
                    'Nos aseguramos que no hay un camino más corto
                    If T(j, vcolu).DistV = MAXINT Then
                        'Actualizamos la tabla de calculos intermedios
                        T(j, vcolu).DistV = T(vfila, vcolu).DistV + 1
                        T(j, vcolu).PrevV.X = vcolu
                        T(j, vcolu).PrevV.Y = vfila
                        'Mete el vertice en la cola
                        V.X = vcolu
                        V.Y = j
                        Call Push(V)
                    End If
            End If
    End If
    j = vfila + 1
    'look to south
    If Limites(j, vcolu) Then
            If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then
                'Nos aseguramos que no hay un camino más corto
                If T(j, vcolu).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    T(j, vcolu).DistV = T(vfila, vcolu).DistV + 1
                    T(j, vcolu).PrevV.X = vcolu
                    T(j, vcolu).PrevV.Y = vfila
                    'Mete el vertice en la cola
                    V.X = vcolu
                    V.Y = j
                    Call Push(V)
                End If
            End If
    End If
    'look to west
    If Limites(vfila, vcolu - 1) Then
            If IsWalkable(MapIndex, vfila, vcolu - 1, NpcIndex) Then
                'Nos aseguramos que no hay un camino más corto
                If T(vfila, vcolu - 1).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    T(vfila, vcolu - 1).DistV = T(vfila, vcolu).DistV + 1
                    T(vfila, vcolu - 1).PrevV.X = vcolu
                    T(vfila, vcolu - 1).PrevV.Y = vfila
                    'Mete el vertice en la cola
                    V.X = vcolu - 1
                    V.Y = vfila
                    Call Push(V)
                End If
            End If
    End If
    'look to east
    If Limites(vfila, vcolu + 1) Then
            If IsWalkable(MapIndex, vfila, vcolu + 1, NpcIndex) Then
                'Nos aseguramos que no hay un camino más corto
                If T(vfila, vcolu + 1).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    T(vfila, vcolu + 1).DistV = T(vfila, vcolu).DistV + 1
                    T(vfila, vcolu + 1).PrevV.X = vcolu
                    T(vfila, vcolu + 1).PrevV.Y = vfila
                    'Mete el vertice en la cola
                    V.X = vcolu + 1
                    V.Y = vfila
                    Call Push(V)
                End If
            End If
    End If
   
   
End Sub


Public Sub SeekPath(ByVal NpcIndex As Integer, Optional ByVal MaxSteps As Integer = 30)
'############################################################
'This Sub seeks a path from the npclist(npcindex).pos
'to the location NPCList(NpcIndex).PFINFO.Target.
'The optional parameter MaxSteps is the maximum of steps
'allowed for the path.
'############################################################
On Error GoTo err
Dim cur_npc_pos As tVertice
Dim tar_npc_pos As tVertice
Dim V As tVertice
Dim NpcMap As Integer
Dim steps As Integer
Dim error As Integer
error = 1
NpcMap = Npclist(NpcIndex).pos.Map

steps = 0

cur_npc_pos.X = Npclist(NpcIndex).pos.Y
cur_npc_pos.Y = Npclist(NpcIndex).pos.X

tar_npc_pos.X = Npclist(NpcIndex).PFINFO.Target.X '  UserList(NPCList(NpcIndex).PFINFO.TargetUser).Pos.X
tar_npc_pos.Y = Npclist(NpcIndex).PFINFO.Target.Y '  UserList(NPCList(NpcIndex).PFINFO.TargetUser).Pos.Y
error = 2
Call InitializeTable(TmpArray, cur_npc_pos)
error = 3
Call InitQueue
error = 4
'We add the first vertex to the Queue
Call Push(cur_npc_pos)
error = 5
Do While (Not IsEmpty)
    If steps > MaxSteps Then Exit Do
    V = Pop
    If V.X = tar_npc_pos.X And V.Y = tar_npc_pos.Y Then Exit Do
    Call ProcessAdjacents(NpcMap, TmpArray, V.Y, V.X, NpcIndex)
Loop
error = 6
Call MakePath(NpcIndex)
error = 7
Exit Sub

err:
Call LogError("Error en seekpath: " & err.Number & "  " & err.Description & " SubProceso:" & err.source & "--->" & error)

End Sub

Private Sub MakePath(ByVal NpcIndex As Integer)
'#######################################################
'Builds the path previously calculated
'#######################################################
On Error GoTo err
Dim Pasos As Integer
Dim miV As tVertice
Dim i As Integer
Dim error As Integer

error = 1
Pasos = TmpArray(Npclist(NpcIndex).PFINFO.Target.Y, Npclist(NpcIndex).PFINFO.Target.X).DistV
error = 2
Npclist(NpcIndex).PFINFO.PathLenght = Pasos
error = 3

If Pasos = MAXINT Then
error = 4
    'MsgBox "There is no path."
    Npclist(NpcIndex).PFINFO.NoPath = True
    error = 5
    Npclist(NpcIndex).PFINFO.PathLenght = 0
    Exit Sub
End If

error = 6
ReDim Npclist(NpcIndex).PFINFO.Path(0 To Pasos) As tVertice
error = 7
miV.X = Npclist(NpcIndex).PFINFO.Target.X
miV.Y = Npclist(NpcIndex).PFINFO.Target.Y
error = 8
For i = Pasos To 1 Step -1
error = 9
    Npclist(NpcIndex).PFINFO.Path(i) = miV
    error = 10
    miV = TmpArray(miV.Y, miV.X).PrevV
Next i

error = 11
Npclist(NpcIndex).PFINFO.CurPos = 1
error = 12
Npclist(NpcIndex).PFINFO.NoPath = False
Exit Sub

err:
'LogError ("ERROR en : MakePath " & err.Description & "--error-> " & error & " X_Y-> " & miV.X & " " & miV.Y)
Npclist(NpcIndex).PFINFO.PathLenght = 0
End Sub

Private Sub InitializeTable(ByRef T() As tIntermidiateWork, ByRef S As tVertice, Optional ByVal MaxSteps As Integer = 30)
'#########################################################
'Initialize the array where we calculate the path
'#########################################################

Dim j As Integer, K As Integer
Const anymap = 1
For j = S.Y - MaxSteps To S.Y + MaxSteps
    For K = S.X - MaxSteps To S.X + MaxSteps
        If InMapBounds(anymap, j, K) Then
            T(j, K).Known = False
            T(j, K).DistV = MAXINT
            T(j, K).PrevV.X = 0
            T(j, K).PrevV.Y = 0
        End If
    Next
Next

T(S.Y, S.X).Known = False
T(S.Y, S.X).DistV = 0

End Sub

