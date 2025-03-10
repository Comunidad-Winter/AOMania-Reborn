Attribute VB_Name = "mod_MouseGamer"
Private Type tClicks
    LCLista As Long
    intClick(4) As Long
    aInt As Byte
    Pend As Boolean
End Type

Public Enum eTipo
    BotonLanzar = 1
    BotonHechizos = 2
    BotonInventario = 3
    ListaHechizos = 4
End Enum

Private Type tClickPositions
    X(4) As Single
    Y(4) As Single
    aInt As Byte
End Type

Public ClickPositions(1 To 4) As tClickPositions


Public GameClick(1 To 3) As tClicks

Private Const LIMITE_INTERVALO As Byte = 50 'Es el limite de diferencia entre intervalos para guardar en logs en milisegundos

' Se almazenan los intervalos entre el click en la lista de hechizos y el boton lanzar
' y luego de 5 intervalos, se comparan

Public Sub ClickCambioInv() 'cambio al menu inventario
    GameClick(2).LCLista = timeGetTime
    GameClick(2).Pend = True
End Sub

Public Sub ClickEnInv() 'Click en algun objeto del inventario
    With GameClick(2)
        If .Pend = True Then
            
            .intClick(.aInt) = timeGetTime - .LCLista
            .Pend = False
            If .aInt = 4 Then
                .aInt = 0
                Call CompararInt(2)
            End If
            .aInt = .aInt + 1
        End If
    End With
    GameClick(3).LCLista = timeGetTime
    GameClick(3).Pend = True
End Sub

Public Sub ClickCambioHech() ' cambio al menu de hechizos
    With GameClick(3)
        If .Pend = True Then
            
            .intClick(.aInt) = timeGetTime - .LCLista
            .Pend = False
            If .aInt = 4 Then
                .aInt = 0
                Call CompararInt(3)
            End If
            .aInt = .aInt + 1
        End If
    End With
End Sub

Public Sub ClickLista()
    GameClick(1).LCLista = timeGetTime
    GameClick(1).Pend = True
End Sub


Public Sub ClickLanzar()
    With GameClick(1)
        If .Pend = True Then
            
            .intClick(.aInt) = timeGetTime - .LCLista
            .Pend = False
            If .aInt = 4 Then
                .aInt = 0
                Call CompararInt(1)
            End If
            .aInt = .aInt + 1
        End If
    End With
End Sub

Private Sub CompararInt(ByVal Index As Byte)
    Dim advs As Byte
    If Index <= 3 Then
        With GameClick(Index)
        
            advs = CompararIntervalos(.intClick)
    
            If advs >= 6 Then _
                Call SendData("SACSAC1" & Index)
                
        End With
    Else
        With ClickPositions(Index - 3)
            
            advs = CompararPosiciones(.X)
            
            If advs >= 8 Then _
                Call SendData("SACSAC1" & Index)
                
            advs = CompararPosiciones(.Y)
    
            If advs >= 8 Then _
                Call SendData("SACSAC1" & Index)
                
        End With
    End If
End Sub
Private Function CompararPosiciones(ByRef ints() As Single) As Byte 'Devuelve la cantidad de posiciones iguales de click encontrados
    Dim xx As Long, yy As Long
    Dim dif() As Single
    Dim advs As Byte
    Dim nDif As Byte
    Dim Advertencias As Byte
    ReDim Preserve dif(0 To 1) As Single
    For xx = 0 To 4
        For yy = 0 To 4
            If yy <> xx Then
                dif(nDif) = ints(xx) - ints(yy)
                If dif(nDif) < 1 And dif(nDif) > -1 Then
                    Advertencias = Advertencias + 1
                End If
                nDif = nDif + 1
                ReDim Preserve dif(0 To nDif) As Single
            End If
        Next yy
    Next xx
    CompararPosiciones = Advertencias
End Function
Private Function CompararIntervalos(ByRef ints() As Long) As Byte 'Devuelve la cantidad de intervalos parecidos encontrados
    Dim xx As Long, yy As Long
    Dim dif() As Long
    Dim advs As Byte
    Dim nDif As Byte
    Dim Advertencias As Byte
    ReDim Preserve dif(0 To 1) As Long
    For xx = 0 To 4
        For yy = 0 To 4
            If yy <> xx Then
                dif(nDif) = ints(xx) - ints(yy)
                If dif(nDif) < 25 And dif(nDif) > -25 Then
                    Advertencias = Advertencias + 1
                End If
                nDif = nDif + 1
                ReDim Preserve dif(0 To nDif) As Long
            End If
        Next yy
    Next xx
    CompararIntervalos = Advertencias
End Function


Public Sub ClickEnObjetoPos(ByVal TIPO As eTipo, ByVal X As Single, ByVal Y As Single)
    With ClickPositions(CByte(TIPO))
        .aInt = .aInt + 1
        .X(.aInt) = X
        .Y(.aInt) = Y
        
        If .aInt = 4 Then
            .aInt = 0
            Call CompararInt(3 + TIPO)
        End If
    End With
End Sub


