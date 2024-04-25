Attribute VB_Name = "ModRankAoM"
Option Explicit

Public CiudadanoOnlineMaxLvl As String
Public CriminalOnlineMaxLvl As String
Public LvlOnlineCiudadano As Long
Public LvlOnlineCriminal As Long

Public NvMaxCiudadanoAoM As Long
Public NvMaxCriminalAoM As Long
Public MaxCiudadanoAoM As String
Public MaxCriminalAoM As String

Public OnlineCiudadano As Long
Public OnlineCriminal As Long

Public UserMaxOroOnline As String
Public OroMaxOnline As Long

Public UserMaxOroOff As String
Public OroMaxOff As Long

Function DirINI()
    DirINI = App.Path & "\Dat\ini\"
End Function

Sub CommandMayor(UserIndex)

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "MAYORES" & _
                                                    CiudadanoOnlineMaxLvl & "," & _
                                                    CriminalOnlineMaxLvl & "," & _
                                                    MaxCiudadanoAoM & "," & _
                                                    MaxCriminalAoM & "," & _
                                                    OnlineCiudadano & "," & _
                                                    OnlineCriminal & "," & _
                                                    UserMaxOroOnline & "," & _
                                                    UserMaxOroOff)

End Sub

Sub Load_Rank()

    CiudadanoOnlineMaxLvl = " "
    CriminalOnlineMaxLvl = " "
    LvlOnlineCiudadano = 0
    LvlOnlineCriminal = 0

    NvMaxCiudadanoAoM = val(GetVar(DirINI & "Ranking.ini", "Ranking", "NvMaxCiudadano"))
    MaxCiudadanoAoM = GetVar(DirINI & "Ranking.ini", "Ranking", "MaxCiudadano")
    NvMaxCiudadanoAoM = val(GetVar(DirINI & "Ranking.ini", "Ranking", "NvMaxCriminal"))
    MaxCiudadanoAoM = GetVar(DirINI & "Ranking.ini", "Ranking", "MaxCriminal")
    UserMaxOroOff = GetVar(DirINI & "Ranking.ini", "Ranking", "MaxOro")
    OroMaxOff = GetVar(DirINI & "Ranking.ini", "Ranking", "MonedaOroMax")

End Sub

Sub Save_Rank(UserIndex As Integer)

    With UserList(UserIndex)

        If .flags.Privilegios = PlayerType.User Then

            If Criminal(UserIndex) Then

                If UCase(CriminalOnlineMaxLvl) = UCase(.Name) Then
                    Call UpdateCriCiuMaxLvl(UserIndex)
                End If

            End If

            If Not Criminal(UserIndex) Then

                If UCase(CiudadanoOnlineMaxLvl) = UCase(.Name) Then
                    Call UpdateCriCiuMaxLvl(UserIndex)
                End If

            End If
        End If

        If UCase(UserMaxOroOnline) = UCase(.Name) Then

            Call UpdateMaxOroRank(UserIndex)
        End If

    End With

End Sub

Sub CountCriCi(UserIndex As Integer)

    With UserList(UserIndex)
        If .flags.Privilegios = PlayerType.User Then

            If Criminal(UserIndex) Then
                OnlineCriminal = OnlineCriminal + 1
            End If

            If Not Criminal(UserIndex) Then
                OnlineCiudadano = OnlineCiudadano + 1
            End If

        End If
    End With

    #If MYSQL = 1 Then
        Call Add_DataBase(UserIndex, "Ranking")
    #End If

End Sub

Sub RestCriCi(UserIndex As Integer)

    With UserList(UserIndex)
        If .flags.Privilegios = PlayerType.User Then

            If Criminal(UserIndex) Then
                OnlineCriminal = OnlineCriminal - 1
            End If

            If Not Criminal(UserIndex) Then
                OnlineCiudadano = OnlineCiudadano - 1
            End If

        End If
    End With

    #If MYSQL = 1 Then
        Call Add_DataBase(UserIndex, "Ranking")
    #End If

End Sub


Sub CriCiuMaxLvl(UserIndex As Integer)

    With UserList(UserIndex)

        If .flags.Privilegios = PlayerType.User Then

            If Criminal(UserIndex) Then
                If .Stats.ELV > NvMaxCriminalAoM Then
                    NvMaxCriminalAoM = .Stats.ELV
                    MaxCriminalAoM = UCase(.Name)
                    Call WriteVar(DirINI & "Ranking.ini", "Ranking", "NvMaxCriminal", NvMaxCriminalAoM)
                    Call WriteVar(DirINI & "Ranking.ini", "Ranking", "MaxCriminal", MaxCriminalAoM)
                End If

                If .Stats.ELV > LvlOnlineCriminal Then
                    CriminalOnlineMaxLvl = UCase(.Name)
                    LvlOnlineCriminal = .Stats.ELV
                End If
            End If

            If Not Criminal(UserIndex) Then

                If .Stats.ELV > NvMaxCiudadanoAoM Then
                    NvMaxCiudadanoAoM = .Stats.ELV
                    MaxCiudadanoAoM = UCase(.Name)
                    Call WriteVar(DirINI & "Ranking.ini", "Ranking", "NvMaxCiudadano", NvMaxCiudadanoAoM)
                    Call WriteVar(DirINI & "Ranking.ini", "Ranking", "MaxCiudadano", MaxCiudadanoAoM)
                End If

                If .Stats.ELV > LvlOnlineCiudadano Then
                    CiudadanoOnlineMaxLvl = UCase(.Name)
                    LvlOnlineCiudadano = .Stats.ELV
                End If

            End If

        End If
    End With

    #If MYSQL = 1 Then
        Call Add_DataBase(UserIndex, "Ranking")
    #End If

End Sub

Sub UpdateCriCiuMaxLvl(UserIndex As Integer)

    With UserList(UserIndex)

        If Criminal(UserIndex) Then

            If .Stats.ELV > NvMaxCriminalAoM Then
                NvMaxCriminalAoM = .Stats.ELV
                MaxCriminalAoM = UCase(.Name)
                Call WriteVar(DirINI & "Ranking.ini", "Ranking", "NvMaxCriminal", NvMaxCriminalAoM)
                Call WriteVar(DirINI & "Ranking.ini", "Ranking", "MaxCriminal", MaxCriminalAoM)
            End If

            If .Name = CriminalOnlineMaxLvl Then
                Call EligeNewCriMaxLevel(UserIndex)
            End If

        End If


        If Not Criminal(UserIndex) Then


            If .Stats.ELV > NvMaxCiudadanoAoM Then
                NvMaxCiudadanoAoM = .Stats.ELV
                MaxCiudadanoAoM = UCase(.Name)
                Call WriteVar(DirINI & "Ranking.ini", "Ranking", "NvMaxCiudadano", NvMaxCiudadanoAoM)
                Call WriteVar(DirINI & "Ranking.ini", "Ranking", "MaxCiudadano", MaxCiudadanoAoM)
            End If

            Call EligeNewCiMaxLevel(UserIndex)
        End If

    End With

End Sub

Sub EligeNewCriMaxLevel(UserIndex As Integer)

    Dim i As Integer
    Dim CUser As Integer

    CUser = NumUsers + 1
    LvlOnlineCriminal = 0

    For i = 1 To CUser

        If CUser = 1 Then
            CiudadanoOnlineMaxLvl = " "
            #If MYSQL = 1 Then
                DoEvents
                Call Add_DataBase(UserIndex, "Ranking")
            #End If
            Exit Sub
        End If

        If UserList(i).flags.Privilegios = User Then

            If Criminal(i) Then

                If UCase(UserList(i).Name) <> UCase(CriminalOnlineMaxLvl) Then
                    If UserList(i).Stats.ELV > LvlOnlineCriminal Then
                        CriminalOnlineMaxLvl = UCase(UserList(i).Name)
                        LvlOnlineCriminal = UserList(i).Stats.ELV
                    End If
                End If

            End If

        End If

    Next i

    ' OnlineCriminal = OnlineCriminal - 1

    If OnlineCriminal = 0 Then
        CriminalOnlineMaxLvl = " "
        LvlOnlineCriminal = 0
    End If

    #If MYSQL = 1 Then
        DoEvents
        Call Add_DataBase(UserIndex, "Ranking")
    #End If
End Sub

Sub EligeNewCiMaxLevel(UserIndex As Integer)

    Dim i As Integer
    Dim CUser As Integer

    CUser = NumUsers + 1

    LvlOnlineCiudadano = 0

    For i = 1 To CUser

        If UserList(i).flags.Privilegios = User Then

            If Not Criminal(i) Then

                If UCase(UserList(i).Name) <> UCase(CiudadanoOnlineMaxLvl) Then

                    If UserList(i).Stats.ELV > LvlOnlineCiudadano Then
                        CiudadanoOnlineMaxLvl = UCase(UserList(i).Name)
                        LvlOnlineCiudadano = UserList(i).Stats.ELV
                    End If

                End If

            End If

        End If

    Next i

    ' OnlineCiudadano = OnlineCiudadano - 1

    If OnlineCiudadano = 0 Then
        CiudadanoOnlineMaxLvl = " "
        LvlOnlineCiudadano = 0
    End If

    #If MYSQL = 1 Then
        DoEvents
        Call Add_DataBase(UserIndex, "Ranking")
    #End If

End Sub

Sub MaxOroRank(UserIndex As Integer)

    Dim VerOro As Long

    With UserList(UserIndex)

        If .flags.Privilegios = PlayerType.User Then

            VerOro = .Stats.GLD + .Stats.Banco

            If VerOro > OroMaxOnline Then

                UserMaxOroOnline = UCase(.Name)
                OroMaxOnline = VerOro

            End If

        End If

    End With

    #If MYSQL = 1 Then
        Call Add_DataBase(UserIndex, "Ranking")
    #End If
End Sub

Sub UpdateMaxOroRank(UserIndex As Integer)
    Dim i As Integer
    Dim CUser As Integer
    Dim VerOro As Long

    If OroMaxOnline > val(GetVar(DirINI & "Ranking.ini", "Ranking", "MonedaOroMax")) Then
        Call WriteVar(DirINI & "Ranking.ini", "Ranking", "MaxOro", UserMaxOroOnline)
        Call WriteVar(DirINI & "Ranking.ini", "Ranking", "MonedaOroMax", OroMaxOnline)
        UserMaxOroOff = UCase(UserMaxOroOnline)
        OroMaxOff = OroMaxOnline
    End If

    CUser = NumUsers + 1
    OroMaxOnline = 0

    For i = 1 To CUser

        If UserList(i).flags.Privilegios = PlayerType.User Then

            If UserList(i).Name <> UserMaxOroOnline Then

                VerOro = UserList(i).Stats.GLD + UserList(i).Stats.Banco

                If VerOro > OroMaxOnline Then

                    UserMaxOroOnline = UCase(UserList(i).Name)
                    OroMaxOnline = VerOro

                End If

            End If

        End If

    Next i

    If OnlineCiudadano = 0 And OnlineCriminal = 0 Then
        UserMaxOroOnline = " "
        OroMaxOnline = 0
    End If


    #If MYSQL = 1 Then
        DoEvents
        Call Add_DataBase(UserIndex, "Ranking")
    #End If

End Sub

Sub CompruebaOroRank(UserIndex As Integer)
    Dim i As Integer
    Dim VerOro As Long

    If UserList(UserIndex).Name = UserMaxOroOnline Then
        OroMaxOnline = UserList(UserIndex).Stats.GLD + UserList(UserIndex).Stats.Banco
    End If

    For i = 1 To NumUsers

        If UserList(i).flags.Privilegios = PlayerType.User Then

            VerOro = UserList(i).Stats.GLD + UserList(i).Stats.Banco

            If VerOro > OroMaxOnline Then
                UserMaxOroOnline = UCase(UserList(i).Name)
                OroMaxOnline = VerOro
            End If

        End If

    Next i

    #If MYSQL = 1 Then
        DoEvents
        Call Add_DataBase(UserIndex, "Ranking")
    #End If
End Sub

Sub OroConnectRank(UserIndex As Integer)
    Dim i As Integer
    Dim Name As String

    Name = UCase(UserList(UserIndex).Name)

    If Name = UserMaxOroOff Then
        Call WriteVar(DirINI & "Ranking.ini", "Ranking", "MonedaOroMax", "0")
    Else
        Exit Sub
    End If

    Dim Arch As String
    Dim VerOro As Long

    Arch = Dir(App.Path & "\Charfile\*.chr")

    Do While Arch > ""
        Arch = Dir
        Name = UCase(readfield2(1, Arch, 46))

        If RevNickOro(Name) = False Then
            If val(GetVar(App.Path & "\Charfile\" & Arch, "INIT", "LOGGED")) = 0 Then
                VerOro = val(GetVar(App.Path & "\Charfile\" & Arch, "STATS", "GLD")) + val(GetVar(App.Path & "\Charfile" & Arch, "STATS", "BANCO"))

                If VerOro > val(GetVar(DirINI & "Ranking.ini", "Ranking", "MonedaOroMax")) Then

                    If UCase(UserMaxOroOff) <> UCase(Name) Then

                        UserMaxOroOff = Name
                        OroMaxOff = VerOro
                        Call WriteVar(DirINI & "Ranking.ini", "Ranking", "MaxOro", Name)
                        Call WriteVar(DirINI & "Ranking.ini", "Ranking", "MonedaOroMax", VerOro)

                    End If

                End If

            End If
        End If

    Loop

    #If MYSQL = 1 Then
        DoEvents
        Call Add_DataBase(UserIndex, "Ranking")
    #End If

End Sub

Function RevNickOro(Name As String)
    Dim i As Integer
    Dim ViewGM As String
    Dim CountGM As Integer

    CountGM = val(GetVar(App.Path & "\Dat\gmsmac.dat", "INIT", "Num"))

    For i = 1 To CountGM
        If UCase(Name) = UCase(GetVar(App.Path & "\Dat\gmsmac.dat", "GM" & i, "Nombre")) Then
            RevNickOro = True
            Exit Function
        End If
    Next i

    RevNickOro = False
End Function

Function RevDirOro()
    Dim Arch As String
    Dim Count As Long

    Arch = Dir(App.Path & "\Charfile\*.chr")

    Do While Arch > ""
        Arch = Dir
        Count = Count + 1
    Loop

    RevDirOro = Count

End Function
