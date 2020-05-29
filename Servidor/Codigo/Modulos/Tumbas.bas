Attribute VB_Name = "Tumbas"
'Mod_Tumbas.bas by JM
Option Explicit

Public Const TUMBA_TOTAN = 1773
Public Const TUMBA_SATUROS = 1774
Public objtotan(1 To 10) As Integer


Sub AomaniaTumba()

Call AomaniaTumbaTotan
Call AomaniaTumbaSatu

End Sub


Sub AomaniaTumbaTotan()


On Error GoTo ErrorTumbas
Dim Pos As WorldPos
Dim newpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim Iteraciones As Long


Dim Map As Integer
Dim X As Integer
Dim Y As Integer

Dim error As Integer

error = 1

error = 2
    
  objtotan(1) = 1
  objtotan(2) = 26
  objtotan(3) = 3
  objtotan(4) = 19
  objtotan(5) = 30
  objtotan(6) = 40
  objtotan(7) = 80
  objtotan(8) = 123
  objtotan(9) = 592
  objtotan(10) = 617
  
    
  objetoTotan = RandomNumber(1, 10)
               
   Do While Not PosicionValida
    Pos.Map = RandomNumber(1, 190)
                Do While True
                
                If (Pos.Map <> 1 And Pos.Map <> 20 And Pos.Map <> 33 And _
                 Pos.Map <> 34 And Pos.Map <> 37 And Pos.Map <> 40 And _
                 Pos.Map <> 41 And Pos.Map <> 42 And Pos.Map <> 43 And _
                 Pos.Map <> 44 And Pos.Map <> 45 And Pos.Map <> 47 And Pos.Map <> 48 And _
                 Pos.Map <> 50 And Pos.Map <> 51 And Pos.Map <> 52 And _
                 Pos.Map <> 58 And Pos.Map <> 59 And Pos.Map <> 60 And _
                 Pos.Map <> 61 And Pos.Map <> 62 And Pos.Map <> 63 And _
                 Pos.Map <> 64 And Pos.Map <> 66 And Pos.Map <> 81 And _
                 Pos.Map <> 82 And Pos.Map <> 83 And Pos.Map <> 84 And _
                 Pos.Map <> 85 And Pos.Map <> 86 And Pos.Map <> 111 And _
                 Pos.Map <> 112 And Pos.Map <> 114 And Pos.Map <> 115 And _
                 Pos.Map <> 116 And Pos.Map <> 117 And Pos.Map <> 118 And _
                 Pos.Map <> 119 And Pos.Map <> 121 And Pos.Map <> 132 And _
                 Pos.Map <> 140 And Pos.Map <> 141 And Pos.Map <> 142 And _
                 Pos.Map <> 143 And Pos.Map <> 144 And Pos.Map <> 145 And _
                 Pos.Map <> 146 And Pos.Map <> 150 And Pos.Map <> 151 And _
                 Pos.Map <> 156 And Pos.Map <> 157 And Pos.Map <> 158 And _
                 Pos.Map <> 159 And Pos.Map <> 160 And Pos.Map <> 161 And _
                 Pos.Map <> 163 And Pos.Map <> 164 And Pos.Map <> 165 And _
                 Pos.Map <> 166 And Pos.Map <> 167 And Pos.Map <> 168 And _
                 Pos.Map <> 169 And Pos.Map <> 170 And Pos.Map <> 171 And _
                 Pos.Map <> 183 And Pos.Map <> 185 And Pos.Map <> 186 And _
                 Pos.Map <> 187 And Pos.Map <> 188 And Pos.Map <> 189 And _
                 Pos.Map <> 190 And Pos.Map < 190) Then Exit Do
                 
                 Pos.Map = RandomNumber(1, 190)
                 'Debug.Print Pos.Map
                ' DoEvents
               Loop
    
    
    
error = 3
        X = RandomNumber(5, 95)
        Y = RandomNumber(5, 95)
        
            
        Pos.X = X
        Pos.Y = Y
        
        Call ClosestLegalPosTumba(Pos, newpos)  'Nos devuelve la posicion valida mas cercana
error = 4
        'Si X e Y son iguales a 0 significa que no se encontro posicion valida
        If LegalPosTumba(newpos.Map, newpos.X, newpos.Y) Then
            'Asignamos las nuevas coordenas solo si son validas
error = 5
           ' Npclist(nIndex).pos.Map = newpos.Map
           ' Npclist(nIndex).pos.X = newpos.X
           ' Npclist(nIndex).pos.Y = newpos.Y
            PosicionValida = True
        
        End If
            
       ' DoEvents
    Loop
    
    
error = 6



error = 7

'Call borrarTumba(posTumbaTotan)

posTumbaTotan.Map = newpos.Map
posTumbaTotan.X = newpos.X
posTumbaTotan.Y = newpos.Y

Call crearTumba(posTumbaTotan, TUMBA_TOTAN)

error = 9
Debug.Print "TumabTotan ha aparecido en el mapa " & newpos.Map & " " & newpos.X & " " & newpos.Y
Exit Sub
ErrorTumbas:
        LogError ("Error en Sub AomaniaTumbaTotan() " & error)
End Sub


Sub AomaniaTumbaSatu()


On Error GoTo ErrorTumbas
Dim Pos As WorldPos
Dim newpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim Iteraciones As Long


Dim Map As Integer
Dim X As Integer
Dim Y As Integer

Dim error As Integer

error = 1

error = 2
    
               
   Do While Not PosicionValida
    Pos.Map = RandomNumber(1, 190)
                Do While True
                
                If (Pos.Map <> 1 And Pos.Map <> 20 And Pos.Map <> 33 And _
                 Pos.Map <> 34 And Pos.Map <> 37 And Pos.Map <> 40 And _
                 Pos.Map <> 41 And Pos.Map <> 42 And Pos.Map <> 43 And _
                 Pos.Map <> 44 And Pos.Map <> 45 And Pos.Map <> 47 And Pos.Map <> 48 And _
                 Pos.Map <> 50 And Pos.Map <> 51 And Pos.Map <> 52 And _
                 Pos.Map <> 58 And Pos.Map <> 59 And Pos.Map <> 60 And _
                 Pos.Map <> 61 And Pos.Map <> 62 And Pos.Map <> 63 And _
                 Pos.Map <> 64 And Pos.Map <> 66 And Pos.Map <> 81 And _
                 Pos.Map <> 82 And Pos.Map <> 83 And Pos.Map <> 84 And _
                 Pos.Map <> 85 And Pos.Map <> 86 And Pos.Map <> 111 And _
                 Pos.Map <> 112 And Pos.Map <> 114 And Pos.Map <> 115 And _
                 Pos.Map <> 116 And Pos.Map <> 117 And Pos.Map <> 118 And _
                 Pos.Map <> 119 And Pos.Map <> 121 And Pos.Map <> 132 And _
                 Pos.Map <> 140 And Pos.Map <> 141 And Pos.Map <> 142 And _
                 Pos.Map <> 143 And Pos.Map <> 144 And Pos.Map <> 145 And _
                 Pos.Map <> 146 And Pos.Map <> 150 And Pos.Map <> 151 And _
                 Pos.Map <> 156 And Pos.Map <> 157 And Pos.Map <> 158 And _
                 Pos.Map <> 159 And Pos.Map <> 160 And Pos.Map <> 161 And _
                 Pos.Map <> 163 And Pos.Map <> 164 And Pos.Map <> 165 And _
                 Pos.Map <> 166 And Pos.Map <> 167 And Pos.Map <> 168 And _
                 Pos.Map <> 169 And Pos.Map <> 170 And Pos.Map <> 171 And _
                 Pos.Map <> 183 And Pos.Map <> 185 And Pos.Map <> 186 And _
                 Pos.Map <> 187 And Pos.Map <> 188 And Pos.Map <> 189 And _
                 Pos.Map <> 190 And Pos.Map < 190) Then Exit Do
                 
                 Pos.Map = RandomNumber(1, 190)
                 'Debug.Print Pos.Map
                ' DoEvents
               Loop
    
    
    
error = 3
        X = RandomNumber(5, 95)
        Y = RandomNumber(5, 95)
        
            
        Pos.X = X
        Pos.Y = Y
        
        Call ClosestLegalPosTumba(Pos, newpos)  'Nos devuelve la posicion valida mas cercana
error = 4
        'Si X e Y son iguales a 0 significa que no se encontro posicion valida
        If LegalPosTumba(newpos.Map, newpos.X, newpos.Y) Then
            'Asignamos las nuevas coordenas solo si son validas
error = 5
           ' Npclist(nIndex).pos.Map = newpos.Map
           ' Npclist(nIndex).pos.X = newpos.X
           ' Npclist(nIndex).pos.Y = newpos.Y
            PosicionValida = True
        
        End If
            
       ' DoEvents
    Loop
    
    
error = 6



error = 7

'Call borrarTumba(posTumbaTotan)

posTumbaSatu.Map = newpos.Map
posTumbaSatu.X = newpos.X
posTumbaSatu.Y = newpos.Y

Call crearTumba(posTumbaSatu, TUMBA_SATUROS)

error = 9
Debug.Print "TumabSatu ha aparecido en el mapa " & newpos.Map & " " & newpos.X & " " & newpos.Y
Exit Sub
ErrorTumbas:
        LogError ("Error en Sub AomaniaTumbaSatu() " & error)
End Sub


Sub borrarTumba(Pos As WorldPos)
    
    If Pos.Map > 0 Then Call EraseObj(ToMap, 0, Pos.Map, 1, Pos.Map, Pos.X, Pos.Y)

End Sub

Sub crearTumba(Pos As WorldPos, ObjIndex As Integer)

 Dim MiObj As Obj
    
 MiObj.Amount = 1
              
 MiObj.ObjIndex = ObjIndex
 
  Call MakeObj(ToMap, 0, Pos.Map, MiObj, Pos.Map, Pos.X, Pos.Y)
 'Call TirarItemAlPiso(pos, MiObj)



End Sub

Sub ClosestLegalPosTumba(Pos As WorldPos, ByRef nPos As WorldPos)
    
Dim interacciones As Integer
interacciones = 0

    nPos = Pos
    
    Do While Not LegalPosTumba(nPos.Map, nPos.X, nPos.Y) And interacciones < 50
    
        nPos.X = nPos.X + (RandomNumber(1, 10) - 5)
        nPos.Y = nPos.Y + (RandomNumber(1, 10) - 5)
        interacciones = interacciones + 1
        
       ' Debug.Print interacciones
        'DoEvents
        
    Loop
    

    
    
End Sub



Function LegalPosTumba(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean


'¿Es un mapa valido?
If (Map <= 0 Or Map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
            LegalPosTumba = False
Else
  
  
        LegalPosTumba = (MapData(Map, X, Y).Blocked <> 1) And _
                   (MapData(Map, X, Y).UserIndex = 0) And _
                   (MapData(Map, X, Y).NpcIndex = 0) And _
                   (Not HayAgua(Map, X, Y))
   
End If

End Function





