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
Public Sub CambioDeArea(ByVal X As Byte, ByVal Y As Byte)

    CurAreaX = X \ AreasX
    CurAreaY = Y \ AreasY

    Dim loopX As Integer, loopY As Integer, CharIndex As Integer

    ' Recorremos el mapa entero (TODO: Se puede optimizar si el server nos enviara la direccion del area que nos movimos)
    For loopX = 1 To 100
        For loopY = 1 To 100

            ' Si el tile esta fuera del area
            If Not EstaDentroDelArea(loopX, loopY) Then

                ' Borrar char
                CharIndex = InMapBounds(loopX, loopY)
                If (CharIndex > 0) Then
                    If (CharIndex <> UserCharIndex) Then
                        Call EraseChar(CharIndex)
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

' Calcula si la posicion se encuentra dentro del area del usuario
Public Function EstaDentroDelArea(ByVal X As Integer, ByVal Y As Integer) As Boolean
    EstaDentroDelArea = (Abs(CurAreaX - X \ AreasX) <= 1) And (Abs(CurAreaY - Y \ AreasY) <= 1)
End Function

Private Function Char_MapPosExits(ByVal X As Byte, ByVal Y As Byte) As Integer
 
        '*****************************************************************
        'Checks to see if a tile position has a char_index and return it
        '*****************************************************************
   
        If (InMapBounds(X, Y)) Then
                Char_MapPosExits = MapData(X, Y).CharIndex
        Else
                Char_MapPosExits = 0
        End If
  
End Function

Private Sub Map_DestroyObject(ByVal X As Byte, ByVal Y As Byte)

      If (InMapBounds(X, Y)) Then

            With MapData(X, Y)
                  .OBJInfo.ObjIndex = 0
                  .OBJInfo.Amount = 0
                  
                  Dim Grh As Grh
                  .ObjGrh = Grh
        
            End With

      End If

End Sub

Private Function Map_PosExitsObject(ByVal X As Byte, ByVal Y As Byte) As Integer
 
      '*****************************************************************
      'Checks to see if a tile position has a char_index and return it
      '*****************************************************************

      If (InMapBounds(X, Y)) Then
            Map_PosExitsObject = MapData(X, Y).ObjGrh.GrhIndex
      Else
            Map_PosExitsObject = 0
      End If
 
End Function
