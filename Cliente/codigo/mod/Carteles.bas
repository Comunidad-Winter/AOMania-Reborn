Attribute VB_Name = "Carteles"

Option Explicit

Const XPosCartel = 117

Const YPosCartel = 177

Const MAXLONG = 40

'Carteles
Public Cartel              As Boolean
Public Leyenda             As String
Public LeyendaFormateada() As String
Public textura             As Integer

Sub InitCartel(Ley As String, Grh As Integer)

    If Not Cartel Then
        Leyenda = Ley
        textura = Grh
        Cartel = True

        ReDim LeyendaFormateada(0 To (Len(Ley) \ (MAXLONG \ 2)))
                
        Dim i As Integer, k As Integer, anti As Integer

        anti = 1
        k = 0
        i = 0
        Call DarFormato(Leyenda, i, k, anti)
        i = 0

        Do While LenB(LeyendaFormateada(i)) <> 0 And i < UBound(LeyendaFormateada)
        
            i = i + 1
        Loop

        ReDim Preserve LeyendaFormateada(0 To i)
    Else
        Exit Sub

    End If

End Sub

Private Sub DarFormato(s As String, i As Integer, k As Integer, anti As Integer)

    If anti + i <= Len(s) + 1 Then
        If ((i >= MAXLONG) And mid$(s, anti + i, 1) = " ") Or (anti + i = Len(s)) Then
            LeyendaFormateada(k) = mid$(s, anti, i + 1)
            k = k + 1
            anti = anti + i + 1
            i = 0
        Else
            i = i + 1

        End If

        Call DarFormato(s, i, k, anti)

    End If

End Sub

Sub DibujarCartel()

    If Not Cartel Then Exit Sub

    Dim X As Integer, Y As Integer

    X = XPosCartel + 20
    Y = YPosCartel + 60
    Call DrawGrhIndextoSurface(textura, XPosCartel, YPosCartel, 0, White)

    Dim j As Integer, desp As Integer

    For j = 0 To UBound(LeyendaFormateada)
        Text_Draw X, Y + desp, LeyendaFormateada(j), White
        desp = desp + (frmMain.Font.Size) + 5
    Next

End Sub

