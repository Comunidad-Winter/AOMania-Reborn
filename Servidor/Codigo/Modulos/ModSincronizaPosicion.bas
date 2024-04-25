Attribute VB_Name = "ModSincronizaPosicion"
Option Explicit

Const NUMCORRECCIONES = 256

Public Type Tcorreccion

    OffsetX(0 To NUMCORRECCIONES - 1) As Integer
    OffsetY(0 To NUMCORRECCIONES - 1) As Integer
    
    codigo As Integer

End Type

Public Correcciones() As Tcorreccion

Public Sub InitCorrecciones(NumUsers As Integer)
    ReDim Correcciones(1 To NumUsers) As Tcorreccion

End Sub

Public Sub Corr_ActualizarPosicion(UI As Integer, X As Integer, Y As Integer)
    Call Corr_IgnorarHastaAhora(UI)
    Call SendData(ToIndex, UI, 0, "PU" & X & "," & Y)

End Sub

Public Sub Corr_IgnorarHastaAhora(UI As Integer)
    'bandera
    Correcciones(UI).OffsetX(Correcciones(UI).codigo) = 120 '100 amon
    
    Correcciones(UI).OffsetY(Correcciones(UI).codigo) = 120 '100 amon
    
    Correcciones(UI).codigo = (Correcciones(UI).codigo + 1) Mod NUMCORRECCIONES
    
    Call SendData(ToIndex, UI, 0, "ACT" & Correcciones(UI).codigo & "*0*0")

End Sub

Public Sub Corr_EntraUser(UI As Integer)

    Dim i As Integer
    
    Correcciones(UI).codigo = 0
    
    For i = 0 To NUMCORRECCIONES - 1
        Correcciones(UI).OffsetX(i) = 0
        Correcciones(UI).OffsetY(i) = 0
    Next i
    
    'le enviamos su codigo
    Call SendData(ToIndex, UI, 0, "ACT" & Correcciones(UI).codigo & "*0*0")
    
End Sub

Public Function Corr_Ignorar(UI As Integer, codigo As Byte) As Boolean

    Dim i As Integer, final As Integer

    With Correcciones(UI)
        i = (codigo) Mod NUMCORRECCIONES
        final = (.codigo) Mod NUMCORRECCIONES
            
        Do Until i = final 'si los codigos coinciden no hay correccion alguna
                
            'bandera
            If ((.OffsetX(i) >= 120) And (.OffsetY(i) >= 120)) Then 'amon
                Corr_Ignorar = True
                Exit Function

            End If

            i = i + 1
            i = i Mod NUMCORRECCIONES
        Loop

    End With
    
    Exit Function

End Function

Public Function Corr_MandaPosicion(UI As Integer, _
                                   X As Byte, _
                                   Y As Byte, _
                                   codigo As Byte, _
                                   who As Integer) As Boolean

    On Error GoTo err

    Dim i     As Integer, tmpi As Integer

    Dim final As Integer

    Dim xpos  As Integer, ypos As Integer, mappos As Integer

    xpos = UserList(UI).pos.X
    ypos = UserList(UI).pos.Y
    mappos = UserList(UI).pos.Map

    'Debug.Print "----------- " & who & "-->" & codigo & " " & X & " " & Y & " " & xpos & " " & ypos

    With Correcciones(UI)
    
        'Debug.Print "----------- " & who & "--2> " & .codigo & " " & .OffsetX(codigo) & " " & .OffsetY(codigo)
    
        'aplicamos a su posicion las correcciones enviadas
        'y no recibidas por el cliente
            
        i = (codigo) Mod NUMCORRECCIONES
        final = (.codigo) Mod NUMCORRECCIONES
            
        Do Until i = final 'si los codigos coinciden no hay correccion alguna
                
            'bandera
            If ((.OffsetX(i) >= 120) And (.OffsetY(i) >= 120)) Then 'amon
                X = 0
                Y = 0
                'Debug.Print ")()()()(----->=100"
            Else
                X = X + .OffsetX(i)
                Y = Y + .OffsetY(i)

                '                    .OffsetX(i) = 0
                '                    .OffsetY(i) = 0
                '
            End If
                
            i = i + 1
            i = i Mod NUMCORRECCIONES
        Loop
            
        Corr_MandaPosicion = True
    
        'si la posicion esta bien nos alegramos y no hacemos nada

        'si la posicion esta bien nos alegramos y no hacemos nada
        If ((UserList(UI).pos.X = X) And (UserList(UI).pos.Y = Y)) Then
            'alegrarse
            Exit Function

        End If
        
        '        If ((xpos = X) And (ypos = Y)) Then
        '            'alegrarse
        '            'If MapData(X, Y).TileExit.Map <> 0 Then Corr_IgnorarHastaAhora (UI)
        '            Exit Function
        '        End If
               
        '        If .OffsetX(.codigo) + xpos < 4 Or .OffsetX(.codigo) + xpos > 96 Or _
        '            .OffsetY(.codigo) + ypos < 4 Or .OffsetY(.codigo) + ypos > 96 Or mappos <> UserList(UI).Pos.Map Then
        '            'Call Corr_ActualizarPosicion(UI, xpos, ypos)
        '            'Corr_IgnorarHastaAhora
        '            Call WarpNearestLegalPos(UI, UserList(UI).Pos.Map, UserList(UI).Pos.X, UserList(UI).Pos.Y, False)
        '            'Call Corr_EntraUser(UI)
        '            Exit Function
        '        End If
        '
        '
               
        'Debug.Print "//////////////" & who & "//> " & .codigo & " " & .OffsetX(codigo) & " " & .OffsetY(codigo)
        'calculamos la diferencia
        '        If mappos <> UserList(UI).Pos.Map Then
        '            Corr_IgnorarHastaAhora (UI)
        '            Exit Function
        '        End If
        '

        'calculamos la diferencia
        .OffsetX(.codigo) = UserList(UI).pos.X - X
        .OffsetY(.codigo) = UserList(UI).pos.Y - Y
        tmpi = .codigo
        
        '        .codigo = (.codigo + 1) Mod 256
        '
        '
        '        .OffsetX(.codigo) = xpos - X
        '        .OffsetY(.codigo) = ypos - Y
        '
        '
        '        tmpi = .codigo
        
        .codigo = (.codigo + 1) Mod NUMCORRECCIONES
        
        'le mandamos la correccion
        
        'If .OffsetX(tmpi) >= 0 And .OffsetY(tmpi) >= 0 Then

        Call SendData(ToIndex, UI, 0, "ACT" & .codigo & "*" & .OffsetX(tmpi) & "*" & .OffsetY(tmpi))
                
        Exit Function

    End With
    
    Exit Function
err:
    Debug.Print "ERROR mandapos " & err.Description & " " & i

End Function

