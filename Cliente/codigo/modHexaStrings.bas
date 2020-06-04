Attribute VB_Name = "modHexaStrings"

Option Explicit

Public Function hexMd52Asc(ByVal MD5 As String) As String

    Dim i As Integer, L As String
    
    MD5 = UCase$(MD5)

    If Len(MD5) Mod 2 = 1 Then MD5 = "0" & MD5
    
    For i = 1 To Len(MD5) \ 2
        L = mid(MD5, (2 * i) - 1, 2)
        hexMd52Asc = hexMd52Asc & Chr(hexHex2Dec(L))
    Next i

End Function

Public Function hexHex2Dec(ByVal hex As String) As Long

    Dim i As Integer, L As String

    For i = 1 To Len(hex)
        L = mid(hex, i, 1)

        Select Case L

            Case "A": L = 10

            Case "B": L = 11

            Case "C": L = 12

            Case "D": L = 13

            Case "E": L = 14

            Case "F": L = 15

        End Select
        
        hexHex2Dec = (L * 16 ^ ((Len(hex) - i))) + hexHex2Dec
    Next i

End Function

Public Function txtOffset(ByVal Text As String, ByVal off As Integer) As String

    Dim i As Integer, L As String

    For i = 1 To Len(Text)
        L = mid(Text, i, 1)
        txtOffset = txtOffset & Chr((Asc(L) + off) Mod 256)
    Next i

End Function
