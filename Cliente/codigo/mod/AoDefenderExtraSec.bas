Attribute VB_Name = "AoDefenderExtraSec"
Option Explicit

Public Function AoDefExt(ByVal valor1 As Integer, _
                         ByVal valor2 As Integer, _
                         ByVal valor3 As Integer, _
                         ByVal valor4 As Integer, _
                         ByVal valor5 As Integer) As String
    AoDefExt = Chr(valor1) + Chr(valor2) + Chr(valor3) + Chr(valor4) + Chr(valor5)

End Function

