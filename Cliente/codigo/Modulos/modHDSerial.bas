Attribute VB_Name = "modHDSerial"
Option Explicit

Private HD_Banned_List() As String

Public Sub load_HDList()

    '
    ' @ maTih.-
    
    Dim numBanneds As Integer
    Dim classRead  As clsIniManager
    Dim ban_Loop   As Long
    
    Set classRead = New clsIniManager
    
    Call classRead.Initialize(App.Path & "\Dat\HDBans.ini")
    
    numBanneds = val(classRead.GetValue("INIT", "NumBAN"))
    
    If (numBanneds = 0) Then
    
        ReDim HD_Banned_List(1 To 1) As String

        For ban_Loop = 1 To 1
            HD_Banned_List(ban_Loop) = val(classRead.GetValue("LIST", "HD" & CStr(ban_Loop)))
        Next ban_Loop
    
        Set classRead = Nothing
        Exit Sub

    End If
    
    ReDim HD_Banned_List(1 To numBanneds) As String
    
    For ban_Loop = 1 To numBanneds
        HD_Banned_List(ban_Loop) = val(classRead.GetValue("LIST", "HD" & CStr(ban_Loop)))
    Next ban_Loop
    
    Set classRead = Nothing
    
End Sub

Private Sub write_HDList(ByVal hdIndex As Integer)

    '
    ' @ maTih.-
    
    Dim fFile As String
    
    fFile = App.Path & "\Dat\HDBans.ini"
    
    Call WriteVar(fFile, "INIT", "NumBAN", UBound(HD_Banned_List()))
    Call WriteVar(fFile, "LIST", "HD" & CStr(hdIndex), HD_Banned_List(hdIndex))

End Sub

Public Function add_HD(ByRef hdString As String) As Boolean
    
    Dim j As Long
    
    For j = 1 To UBound(HD_Banned_List())

        If (HD_Banned_List(j) = vbNullString) Then Exit For
    Next j
    
    If (j > UBound(HD_Banned_List())) Then
        ReDim Preserve HD_Banned_List(1 To (UBound(HD_Banned_List())) + 1) As String

    End If
    
    HD_Banned_List(j) = hdString
    
    Call write_HDList(j)
    
    add_HD = True

End Function

Public Function remove_HD(ByRef hdString As String) As Boolean
    
    Dim hdIndex As Integer
    
    hdIndex = check_HD(hdString)
    
    If (hdIndex = -1) Then
        remove_HD = False
        Exit Function

    End If
    
    HD_Banned_List(hdIndex) = vbNullString
    
    Dim NewValue As Integer
    NewValue = UBound(HD_Banned_List()) - 1

    If NewValue < 1 Then NewValue = 1
    ReDim Preserve HD_Banned_List(1 To NewValue) As String
    
    Call write_HDList(hdIndex)

End Function

Public Function check_HD(ByRef hdString As String) As Integer
    
    Dim j As Long
    
    For j = 1 To UBound(HD_Banned_List())

        If UCase$(HD_Banned_List(j) = UCase$(hdString)) Then
            check_HD = CInt(j)
            Exit Function

        End If

    Next j
    
    check_HD = -1

End Function

