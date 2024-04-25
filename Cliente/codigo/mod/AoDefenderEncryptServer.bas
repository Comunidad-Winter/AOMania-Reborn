Attribute VB_Name = "AoDefenderEncryptServer"
Option Explicit

Private Function ConvToHex(X As Integer) As String

    If X > 9 Then
        ConvToHex = Chr(X + 55)
    Else
        ConvToHex = CStr(X)

    End If

End Function

' función que codifica el dato
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function AoDefServEncrypt(DataValue As Variant) As Variant
    
    Dim X         As Long

    Dim Temp      As String

    Dim TempNum   As Integer

    Dim TempChar  As String

    Dim TempChar2 As String
    
    For X = 1 To Len(DataValue)
        TempChar2 = mid(DataValue, X, 1)
        TempNum = Int(Asc(TempChar2) / 16)
        
        If ((TempNum * 16) < Asc(TempChar2)) Then
               
            TempChar = ConvToHex(Asc(TempChar2) - (TempNum * 16))
            Temp = Temp & ConvToHex(TempNum) & TempChar
        Else
            Temp = Temp & ConvToHex(TempNum) & "0"
        
        End If

    Next X
    
    AoDefServEncrypt = Temp

End Function

Private Function ConvToInt(X As String) As Integer
    
    Dim X1   As String

    Dim X2   As String

    Dim Temp As Integer
    
    X1 = mid(X, 1, 1)
    X2 = mid(X, 2, 1)
    
    If IsNumeric(X1) Then
        Temp = 16 * Int(X1)
    Else
        Temp = (Asc(X1) - 55) * 16

    End If
    
    If IsNumeric(X2) Then
        Temp = Temp + Int(X2)
    Else
        Temp = Temp + (Asc(X2) - 55)

    End If
    
    ' retorno
    ConvToInt = Temp
    
End Function

' función que decodifica el dato
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function AoDefServDecrypt(DataValue As Variant) As Variant
    
    Dim X       As Long

    Dim Temp    As String

    Dim HexByte As String
    
    For X = 1 To Len(DataValue) Step 2
        
        HexByte = mid(DataValue, X, 2)
        Temp = Temp & Chr(ConvToInt(HexByte))
        
    Next X

    ' retorno
    AoDefServDecrypt = Temp
    
End Function

