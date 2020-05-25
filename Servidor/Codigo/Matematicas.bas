Attribute VB_Name = "Matematicas"

Option Explicit

Public Function Porcentaje(ByVal Total As Long, ByVal Porc As Long) As Long

    Porcentaje = (Total * Porc) / 100

End Function

Public Function SD(ByVal n As Integer) As Integer

'Call LogTarea("Function SD n:" & n)
'Suma digitos

    Do
        SD = SD + (n Mod 10)
        n = n \ 10
    Loop While (n > 0)

End Function

Public Function SDM(ByVal n As Integer) As Integer

'Call LogTarea("Function SDM n:" & n)
'Suma digitos cada digito menos dos

    Do
        SDM = SDM + (n Mod 10) - 1
        n = n \ 10
    Loop While (n > 0)

End Function

Public Function Complex(ByVal n As Integer) As Integer

'Call LogTarea("Complex")

    If n Mod 2 <> 0 Then
        Complex = n * SD(n)
    Else
        Complex = n * SDM(n)

    End If

End Function

Function Distancia(ByRef wp1 As WorldPos, ByRef wp2 As WorldPos) As Long
'Encuentra la distancia entre dos WorldPos
    Distancia = Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.Map - wp2.Map) * 100)

End Function

Function Distance(x1 As Variant, Y1 As Variant, x2 As Variant, Y2 As Variant) As Double

'Encuentra la distancia entre dos puntos

    Distance = Sqr(((Y1 - Y2) ^ 2 + (x1 - x2) ^ 2))

End Function

Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

'Lo puse en sub Main()
'Randomize Timer

'RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
'If RandomNumber > UpperBound Then RandomNumber = UpperBound

RandomNumber = Int(Rnd * (UpperBound - LowerBound + 1)) + LowerBound

End Function

Function RandomNumber2(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Double

    RandomNumber2 = Int(Rnd * (UpperBound - LowerBound + 1)) + LowerBound

End Function

Sub AddtoVar(ByRef Var As Variant, ByVal Addon As Variant, ByVal max As Variant)
    'Le suma un valor a una variable respetando el maximo valor

    If Var >= max Then
        Var = max
    Else
        Var = Var + Addon

        If Var > max Then
            Var = max

        End If

    End If

End Sub
