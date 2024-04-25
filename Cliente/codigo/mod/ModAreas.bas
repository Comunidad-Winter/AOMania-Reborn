Attribute VB_Name = "ModAreas"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Marquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matias Fernando Pequeno
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez

' WyroX: Pequenia modificacion para que el tamanio de las areas se calcule automaticamente
' en base al tamanio del render y de un valor arbitrario para el buffer (tiles extra)

Option Explicit

' Cantidad de tiles buffer
' (para que graficos grandes se vean desde fuera de la pantalla)
' (debe coincidir con el mismo valor en el server - areas)
Public Const TilesBuffer As Byte = 5

' Tamanio de las areas
Private AreasX As Byte
Private AreasY As Byte

' Area actual
Private CurAreaX As Integer
Private CurAreaY As Integer

Public Sub CalcularAreas(HalfWindowTileWidth As Integer, HalfWindowTileHeight As Integer)
    AreasX = HalfWindowTileWidth + TileBufferSize
    AreasY = HalfWindowTileHeight + TileBufferSize
End Sub

' Elimina todo fuera del area del usuario
Public Sub CambioDeArea(ByVal x As Byte, ByVal y As Byte)

    CurAreaX = x \ AreasX
    CurAreaY = y \ AreasY

    Dim loopX As Integer, loopY As Integer, charindex As Integer

    ' Recorremos el mapa entero (TODO: Se puede optimizar si el server nos enviara la direccion del area que nos movimos)
    For loopX = 1 To 100
        For loopY = 1 To 100
             
            ' Si el tile esta fuera del area
            If Not EstaDentroDelArea(loopX, loopY) Then

                ' Borrar char
                charindex = InMapBounds(loopX, loopY)
                If (charindex > 0) Then
                    If (charindex <> UserCharIndex) Then
                        Call EraseChar(charindex)
                    End If
                End If

                ' Borrar objeto
                If (Map_PosExitsObject(loopX, loopY) > 0) Then
                    Call Map_DestroyObject(loopX, loopY)
                End If

            End If

        Next loopY
    Next loopX

End Sub

Public Sub CleanerPlus()
          Dim x As Long
   Dim y As Long

'edit cambio el rango de valores en x y para solucionar otro bug con respecto al cambio de mapas

  For x = XMinMapSize To XMaxMapSize
       For y = YMinMapSize To YMaxMapSize

       If (MapData(x, y).charindex) Then
         Call EraseChar(MapData(x, y).charindex)
       End If

      If (MapData(x, y).ObjGrh.GrhIndex) Then
        Call Map_DestroyObject(x, y)
      End If

   Next y
 Next x
End Sub

' Calcula si la posicion se encuentra dentro del area del usuario
Public Function EstaDentroDelArea(ByVal x As Integer, ByVal y As Integer) As Boolean
    EstaDentroDelArea = (Abs(CurAreaX - x \ AreasX) <= 1) And (Abs(CurAreaY - y \ AreasY) <= 1)
End Function

Private Function Char_MapPosExits(ByVal x As Byte, ByVal y As Byte) As Integer
 
        '*****************************************************************
        'Checks to see if a tile position has a char_index and return it
        '*****************************************************************
   
        If (InMapBounds(x, y)) Then
                Char_MapPosExits = MapData(x, y).charindex
        Else
                Char_MapPosExits = 0
        End If
  
End Function

Private Sub Map_DestroyObject(ByVal x As Byte, ByVal y As Byte)

      If (InMapBounds(x, y)) Then

            With MapData(x, y)
                  .OBJInfo.ObjIndex = 0
                  .OBJInfo.Amount = 0
                  
                  Dim Grh As Grh
                  .ObjGrh = Grh
        
            End With

      End If

End Sub

Private Function Map_PosExitsObject(ByVal x As Byte, ByVal y As Byte) As Integer
 
      '*****************************************************************
      'Checks to see if a tile position has a char_index and return it
      '*****************************************************************

      If (InMapBounds(x, y)) Then
            Map_PosExitsObject = MapData(x, y).ObjGrh.GrhIndex
      Else
            Map_PosExitsObject = 0
      End If
 
End Function
