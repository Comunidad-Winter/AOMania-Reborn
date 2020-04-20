Attribute VB_Name = "MYSQL"
#If MYSQL = 1 Then

    Option Explicit

    Private Con As New ADODB.Connection

    Private Host_Mysql As String
    Private Database_Mysql As String
    Private User_Mysql As String
    Private Password_Mysql As String

    Public StatusServer As Long

Sub Load_Mysql()

    Host_Mysql = GetVar(App.Path & "\Server.ini", "MYSQL", "Host")
    Database_Mysql = GetVar(App.Path & "\Server.ini", "MYSQL", "DataBase")
    User_Mysql = GetVar(App.Path & "\Server.ini", "MYSQL", "User")
    Password_Mysql = GetVar(App.Path & "\Server.ini", "MYSQL", "Pass")

    DoEvents

    Call Connect_Mysql
End Sub

Sub Connect_Mysql()

    Dim strOpen

    strOpen = "DRIVER={MySQL ODBC 5.1 Driver};" & "SERVER=" & Host_Mysql & ";" & " DATABASE=" & Database_Mysql & ";" & "UID=" & User_Mysql & ";PWD=" & Password_Mysql & "; OPTION=3"

    Con.Open strOpen

    If Con.State = 0 Then
        Call MsgBox("Hubo un 'Error' con la conexión al Mysql, necesita reconfigurarlo.", vbInformation)
    End If

End Sub

Sub Add_DataBase(UserIndex As Integer, Tabla As String)
    Dim Rs As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String

    If Tabla = "Status" Then
        If StatusServer = 0 Then
            StatusServer = 1
            Con.Execute "DELETE FROM Status"
            Con.Execute "INSERT INTO Status VALUES(" & StatusServer & ")"
        Else
            StatusServer = 0
            Con.Execute "DELETE FROM Status"
            Con.Execute "INSERT INTO Status VALUES(" & StatusServer & ")"
        End If
    End If

    If Tabla = "Castillos" Then
        Con.Execute "DELETE FROM Castillos"
        Con.Execute ("INSERT INTO Castillos (Norte, Este, Oeste, Sur, Fortaleza)VALUES('" & Norte & _
                     "', '" & Este & _
                     "', '" & Oeste & _
                     "', '" & Sur & _
                     "', '" & Fortaleza & _
                     "')")
    End If


    If UserIndex = 0 Then Exit Sub

    mUser = UserList(UserIndex)

    If Len(mUser) = 0 Then Exit Sub

    If Tabla = "Online" Then
        Con.Execute "DELETE FROM Online"
        Con.Execute "INSERT INTO Online VALUES(" & NumUsers & ")"
    End If

    If Tabla = "Ranking" Then
        Con.Execute "DELETE FROM Ranking"
        Con.Execute ("INSERT INTO Ranking (MayorNivelCriminalOnline, MayorNivelCiudadanoOnline, MaxCiudadano, MaxCriminal, OnlineCiudadano, OnlineCriminal, UserMaxOroOnline, UserMaxOroOffline)VALUES('" & CriminalOnlineMaxLvl & _
                     "', '" & CiudadanoOnlineMaxLvl & _
                     "', '" & MaxCiudadanoAoM & _
                     "', '" & MaxCriminalAoM & _
                     "', '" & OnlineCiudadano & _
                     "', '" & OnlineCriminal & _
                     "', '" & UserMaxOroOnline & _
                     "', '" & UserMaxOroOff & _
                     "')")
    End If

    If Tabla = "SaveUser" Then
        Set Rs = Nothing
        Set Rs = Con.Execute("SELECT * FROM flags WHERE Usuario='" & mUser.Name & "'")

        If Rs.BOF = True Or Rs.EOF = True Then
            Con.Execute ("INSERT INTO flags (Usuario, Nivel, Raza, Clase, Oro, CiudadanosMatados, CriminalesMatados, AoMCreditos)VALUES('" & mUser.Name & _
                         "', '" & mUser.Stats.ELV & _
                         "', '" & mUser.Raza & _
                         "', '" & mUser.Clase & _
                         "', '" & mUser.Stats.GLD + mUser.Stats.Banco & _
                         "', '" & mUser.Faccion.CiudadanosMatados & _
                         "', '" & mUser.Faccion.CriminalesMatados & _
                         "', '" & mUser.AoMCreditos & _
                         "')")

            Set Rs = Nothing

        Else

            If Rs.BOF Or Rs.EOF Then Call Con.Execute("INSERT INTO flags (Usuario) VALUES (" & mUser.Name & ")")
            Set Rs = Nothing

        End If

        str = "UPDATE `flags` SET"

        str = str & " Usuario='" & mUser.Name & "'"
        str = str & ",Nivel='" & mUser.Stats.ELV & "'"
        str = str & ",Raza='" & mUser.Raza & "'"
        str = str & ",Clase='" & mUser.Clase & "'"
        str = str & ",Oro='" & mUser.Stats.GLD + mUser.Stats.Banco & "'"
        str = str & ",CiudadanosMatados='" & mUser.Faccion.CiudadanosMatados & "'"
        str = str & ",CriminalesMatados='" & mUser.Faccion.CriminalesMatados & "'"
        str = str & ",AoMCreditos='" & mUser.Faccion.CriminalesMatados & "'"

        str = str & " WHERE Usuario='" & mUser.Name & "'"
        Call Con.Execute(str)

    End If

    If Tabla = "Account" Then
        Set Rs = Nothing
        Set Rs = Con.Execute("SELECT * FROM Accounts WHERE Usuario='" & mUser.Name & "'")

        If Rs.BOF = True Or Rs.EOF = True Then
            Con.Execute ("INSERT INTO Accounts (Usuario, Email, Password)VALUES('" & mUser.Name & _
                         "', '" & mUser.Email & _
                         "', '" & mUser.Password & _
                         "')")

            Set Rs = Nothing

        Else

            If Rs.BOF Or Rs.EOF Then Call Con.Execute("INSERT INTO Accounts (Usuario) VALUES (" & mUser.Name & ")")
            Set Rs = Nothing

        End If

        str = "UPDATE `Accounts` SET"

        str = str & " Usuario='" & mUser.Name & "'"
        str = str & ",Email='" & mUser.Email & "'"
        str = str & ",Password='" & mUser.Password & "'"

        str = str & " WHERE Usuario='" & mUser.Name & "'"
        Call Con.Execute(str)

    End If

End Sub

#End If

Sub Reset_Mysql()
    Con.Execute "DELETE FROM Online"
    Con.Execute "DELETE FROM Status"
    Con.Execute "DELETE FROM Flags"
    Con.Execute "DELETE FROM Castillos"
    Con.Execute "DELETE FROM Accounts"
    Con.Execute "DELETE FROM Ranking"
End Sub
