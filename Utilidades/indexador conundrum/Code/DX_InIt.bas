Attribute VB_Name = "Mod_DX"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public DirectX           As New DirectX7
Public DirectDraw        As DirectDraw7

Public PrimarySurface    As DirectDrawSurface7
Public PrimaryClipper    As DirectDrawClipper
Public SecundaryClipper  As DirectDrawClipper
Public BackBufferSurface As DirectDrawSurface7

Public oldResHeight      As Long, oldResWidth As Long
Public bNoResChange      As Boolean

Private Sub IniciarDXobject(DX As DirectX7)

    Err.Clear

    On Error Resume Next

    Set DX = New DirectX7

    If Err Then
        MsgBox _
                "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        End

    End If

End Sub

Private Sub IniciarDDobject(DD As DirectDraw7)
    Err.Clear

    On Error Resume Next

    Set DD = DirectX.DirectDrawCreate("")

    If Err Then
        MsgBox _
                "No se puede iniciar DirectDraw. Por favor asegurese de tener la ultima version correctamente instalada."
        End

    End If

End Sub

Public Sub IniciarObjetosDirectX()

    On Error Resume Next

    Call IniciarDXobject(DirectX)
    Call IniciarDDobject(DirectDraw)

    Exit Sub

End Sub


