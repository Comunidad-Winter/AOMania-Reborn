Attribute VB_Name = "TCP"

'Pablo Ignacio Márquez

Option Explicit

'RUTAS DE ENVIO DE DATOS
Public Enum SendTarget

    ToIndex = 0         ' Envia a un solo User
    ToAll = 1           ' A todos los Users
    ToMap = 2           ' Todos los Usuarios en el mapa
    ToPCArea = 3        ' Todos los Users en el area de un user determinado
    ToNone = 4          ' Ninguno
    ToAllButIndex = 5   ' Todos menos el index
    ToMapButIndex = 6   ' Todos en el mapa menos el indice
    togm = 7
    ToNPCArea = 8       ' Todos los Users en el area de un user determinado
    ToGuildMembers = 9
    ToAdmins = 10
    ToPCAreaButIndex = 11
    ToAdminsAreaButConsejeros = 12
    ToDiosesYclan = 13
    ToConsejo = 14
    ToClanArea = 15
    ToConsejoCaos = 16
    ToRolesMasters = 17
    ToDeadArea = 18
    ToCiudadanos = 19
    ToCriminales = 20
    ToPartyArea = 21
    ToReal = 22
    ToCaos = 23
    ToCiudadanosYRMs = 24
    ToCriminalesYRMs = 25
    ToRealYRMs = 26
    ToCaosYRMs = 27

End Enum

Sub DarCuerpoYCabeza(ByRef UserBody As Integer, ByRef UserHead As Integer, ByVal Raza As String, ByVal Gen As String)

'TODO: Poner las heads en arrays, así se acceden por índices
'y no hay problemas de discontinuidad de los índices.
'También se debe usar enums para raza y sexo
    Select Case UCase$(Gen)

    Case "HOMBRE"

        Select Case UCase$(Raza)

        Case "HUMANO"
            UserHead = RandomNumber(1, 14)
            UserBody = 1

        Case "ELFO"
            UserHead = RandomNumber(102, 110)
            UserBody = 2

        Case "ELFO OSCURO"
            UserHead = RandomNumber(202, 203)
            UserBody = 32

        Case "ENANO"
            UserHead = RandomNumber(301, 310)
            UserBody = 52

        Case "GNOMO"
            UserHead = 401
            UserBody = 52

        Case "HOBBIT"
            UserHead = RandomNumber(609, 611)
            UserBody = 297

        Case "ORCO"
            UserHead = RandomNumber(602, 605)
            UserBody = 300

        Case "LICANTROPO"
            UserHead = RandomNumber(3, 11)
            UserBody = 1

        Case "VAMPIRO"
            UserHead = RandomNumber(710, 712)
            UserBody = 32

        Case "CICLOPE"
            UserHead = RandomNumber(530, 532)
            UserBody = 1

        Case Else
            UserHead = 1
            UserBody = 1

        End Select

    Case "MUJER"

        Select Case UCase$(Raza)

        Case "HUMANO"
            UserHead = RandomNumber(68, 72)
            UserBody = 1

        Case "ELFO"
            UserHead = RandomNumber(170, 172)
            UserBody = 2

        Case "ELFO OSCURO"
            UserHead = RandomNumber(270, 272)
            UserBody = 40

        Case "GNOMO"
            UserHead = RandomNumber(470, 473)
            UserBody = 52

        Case "ENANO"
            UserHead = 370
            UserBody = 52

        Case "HOBBIT"
            UserHead = RandomNumber(612, 615)
            UserBody = 298

        Case "ORCO"
            UserHead = RandomNumber(606, 607)
            UserBody = 302

        Case "LICANTROPO"
            UserHead = RandomNumber(68, 72)
            UserBody = 39

        Case "VAMPIRO"
            UserHead = RandomNumber(710, 712)
            UserBody = 40

        Case "CICLOPE"
            UserHead = RandomNumber(533, 535)
            UserBody = 39

        Case Else
            UserHead = 70
            UserBody = 1

        End Select

    End Select

End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Integer

    cad = LCase$(cad)

    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))

        If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
            AsciiValidos = False
            Exit Function

        End If

    Next i

    AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Integer

    cad = LCase$(cad)

    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))

        If (car < 48 Or car > 57) Then
            Numeric = False
            Exit Function

        End If

    Next i

    Numeric = True

End Function

Function NombrePermitido(ByVal nombre As String) As Boolean
    Dim i As Integer

    For i = 1 To UBound(ForbidenNames)

        If InStr(nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function

        End If

    Next i

    NombrePermitido = True

End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean

    Dim LoopC As Integer

    For LoopC = 1 To NUMSKILLS

        If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then
            Exit Function

            If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100

        End If

    Next LoopC

    ValidateSkills = True

End Function

'Barrin 3/3/03
'Agregué PadrinoName y Padrino password como opcionales, que se les da un valor siempre y cuando el servidor esté usando el sistema
Sub ConnectNewUser(UserIndex As Integer, _
                   Name As String, _
                   Password As String, _
                   UserRaza As String, _
                   UserSexo As String, _
                   UserClase As String, _
                   UserBanco As String, _
                   UserPersonaje As String, UserEmail As String, ByVal HdSerial As String)

    If Not AsciiValidos(Name) Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRNombre invalido.")
        Exit Sub

    End If

    Dim LoopC As Integer
    Dim totalskpts As Long

    '¿Existe el personaje?
    If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = True Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRYa existe el personaje.")
        Exit Sub

    End If

    'Tiró los dados antes de llegar acá??
    If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRDebe tirar los dados antes de poder crear un personaje.")
        Exit Sub

    End If

    UserList(UserIndex).flags.Muerto = 0
    UserList(UserIndex).flags.Escondido = 0

    UserList(UserIndex).Reputacion.AsesinoRep = 0
    UserList(UserIndex).Reputacion.BandidoRep = 0
    UserList(UserIndex).Reputacion.BurguesRep = 0
    UserList(UserIndex).Reputacion.LadronesRep = 0
    UserList(UserIndex).Reputacion.NobleRep = 1000
    UserList(UserIndex).Reputacion.PlebeRep = 30

    UserList(UserIndex).Reputacion.Promedio = 30 / 6

    UserList(UserIndex).Name = Name
    UserList(UserIndex).Clase = UserClase
    UserList(UserIndex).Raza = UserRaza
    UserList(UserIndex).Genero = UserSexo
    UserList(UserIndex).Email = UserEmail
    UserList(UserIndex).flags.Casado = 0
    UserList(UserIndex).Pareja = ""

    UserList(UserIndex).Telepatia = 0

    Select Case UCase$(UserRaza)

    Case "HUMANO"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + 1
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + 1
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) + 1
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) + 0
        UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) = UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) + 2

    Case "ELFO"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + 1
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + 2
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) + 2
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) + 2
        UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) = UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) + 1

    Case "ELFO OSCURO"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + 2
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + 4
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) + 3
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) - 2
        UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) = UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) + 0

    Case "ENANO"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + 3
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + 1
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) - 3
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) - 2
        UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) = UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) + 3

    Case "GNOMO"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) - 1
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + 3
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) + 2
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) + 1
        UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) = UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) + 1

    Case "HOBBIT"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) - 5
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + 6
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) + 4
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) + 3
        UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) = UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) - 1

    Case "ORCO"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + 5
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) - 6
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) - 5
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) - 3
        UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) = UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) + 3

    Case "VAMPIRO"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + 2
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + 1
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) + 2
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) + 1
        UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) = UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) + 0

    Case "CICLOPE"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + 3
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + 1
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) + 0
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) - 2
        UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) = UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) + 2

    Case "LICANTROPO"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + 0
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + 0
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) - 0
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) + 0
        UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) = UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) + 0
    End Select

    UserList(UserIndex).Stats.SkillPts = SkillPointInicial
    UserList(UserIndex).PalabraSecreta = UserBanco
    UserList(UserIndex).flags.RPasswd = UserPersonaje

    totalskpts = 0

    'Abs PREVINENE EL HACKEO DE LOS SKILLS %%%%%%%%%%%%%
    For LoopC = 1 To NUMSKILLS
        totalskpts = totalskpts + Abs(UserList(UserIndex).Stats.UserSkills(LoopC))
    Next LoopC

    If totalskpts > 10 Then
        Call LogHackAttemp(UserList(UserIndex).Name & " intento hackear los skills.")
        Call BorrarUsuario(UserList(UserIndex).Name)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If

    '%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%

    UserList(UserIndex).Password = MD5String(Password)
    UserList(UserIndex).char.heading = eHeading.SOUTH

    Call DarCuerpoYCabeza(UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).Raza, UserList(UserIndex).Genero)

    UserList(UserIndex).OrigChar = UserList(UserIndex).char

    UserList(UserIndex).char.WeaponAnim = NingunArma
    UserList(UserIndex).char.ShieldAnim = NingunEscudo
    UserList(UserIndex).char.CascoAnim = NingunCasco
    UserList(UserIndex).char.Alas = NingunAlas

    UserList(UserIndex).Stats.MET = 1

    Dim MiInt As Long
    MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) \ 3)

    UserList(UserIndex).Stats.MaxHP = 15 + MiInt
    UserList(UserIndex).Stats.MinHP = 15 + MiInt

    MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) \ 6)

    If MiInt = 1 Then MiInt = 2

    UserList(UserIndex).Stats.MaxSta = 20 * MiInt
    UserList(UserIndex).Stats.MinSta = 20 * MiInt

    UserList(UserIndex).Stats.MaxAGU = 100
    UserList(UserIndex).Stats.MinAGU = 100
    UserList(UserIndex).Stats.TrofOro = 0
    UserList(UserIndex).Stats.TrofPlata = 0
    UserList(UserIndex).Stats.TrofBronce = 0
    UserList(UserIndex).Stats.MaxHam = 100
    UserList(UserIndex).Stats.MinHam = 100

    ' puntos
    UserList(UserIndex).Stats.PuntosDuelos = 0
    UserList(UserIndex).Stats.PuntosTorneo = 0
    UserList(UserIndex).Stats.PuntosRetos = 0

    ' soporte
    UserList(UserIndex).Pregunta = "Ninguna"
    UserList(UserIndex).Respuesta = "Ninguna"

    '<-----------------MANA----------------------->
    Select Case UCase$(UserClase)

    Case "MAGO"

        UserList(UserIndex).Stats.MaxMAN = 100
        UserList(UserIndex).Stats.MinMAN = 100

    Case "BRUJO"

        UserList(UserIndex).Stats.MaxMAN = 100
        UserList(UserIndex).Stats.MinMAN = 100

    Case "ASESINO"

        UserList(UserIndex).Stats.MaxMAN = 50
        UserList(UserIndex).Stats.MinMAN = 50

    Case "CLERIGO"

        UserList(UserIndex).Stats.MaxMAN = 50
        UserList(UserIndex).Stats.MinMAN = 50

    Case "DRUIDA"

        UserList(UserIndex).Stats.MaxMAN = 90
        UserList(UserIndex).Stats.MinMAN = 90

    Case "BARDO"

        UserList(UserIndex).Stats.MaxMAN = 70
        UserList(UserIndex).Stats.MinMAN = 70

    Case Else

        UserList(UserIndex).Stats.MaxMAN = 0
        UserList(UserIndex).Stats.MinMAN = 0

    End Select

    'Hasta aqui

    If UCase$(UserClase) = "MAGO" Or UCase$(UserClase) = "CLERIGO" Or UCase$(UserClase) = "DRUIDA" Or UCase$(UserClase) = "BARDO" Or UCase$( _
       UserClase) = "ASESINO" Or UCase$(UserClase) = "BRUJO" Then
        UserList(UserIndex).Stats.UserHechizos(1) = 2

    End If

    UserList(UserIndex).Stats.MaxHit = 2
    UserList(UserIndex).Stats.MinHit = 1
    UserList(UserIndex).Stats.GLD = 0
    UserList(UserIndex).Stats.Exp = 0
    UserList(UserIndex).Stats.ELV = 1
    UserList(UserIndex).Stats.ELU = levelELU(UserList(UserIndex).Stats.ELV)

    Call SetearInv(UserIndex, UCase$(UserList(UserIndex).Clase), UCase$(UserList(UserIndex).Raza))

    'Open User

    'Call EnviaRegistro(Name, Password, UserEmail, UserClase, UserRaza)

    Call SaveUser(UserIndex, CharPath & UCase$(Name) & ".chr")

    Call ConnectUser(UserIndex, Name, UserList(UserIndex).Password, HdSerial)

End Sub

Sub CloseSocket(ByVal UserIndex As Integer)
    Dim LoopC As Integer
    Dim i As Integer
    Dim Total As Integer

    On Error GoTo errhandler

    If UserIndex = LastUser Then

        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1

            If LastUser < 1 Then Exit Do
        Loop

    End If

    If NocheLicantropo = True Then
        If UserList(UserIndex).flags.Privilegios = PlayerType.User And _
           UCase$(UserList(UserIndex).Raza) = "LICANTROPO" And _
           UserList(UserIndex).flags.Licantropo = "1" Then
            Call QuitarPoderLicantropo(UserIndex)
        End If
    End If
    
    If UserList(UserIndex).GuildIndex > 0 Then
            Call SendData(ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "PL")
    End If

    DoEvents

    Call RestCriCi(UserIndex)

    DoEvents

    If UserList(UserIndex).GranPoder = 1 Then
        Call mod_GranPoder.DesconectaPoder(UserIndex)
    End If

    If UserList(UserIndex).flags.automatico = True Then
        Call Rondas_UsuarioDesconecta(UserIndex)
    End If

    If (UserList(UserIndex).Name <> "") And UserList(UserIndex).flags.Privilegios > PlayerType.User And (UserList(UserIndex).flags.Privilegios < _
                                                                                                         PlayerType.Dios Or UserList(UserIndex).flags.Privilegios >= PlayerType.Dios) Then

        If Not UserList(UserIndex).pos.Map = 47 Then
            Call WarpUserChar(UserIndex, 47, RandomNumber(58, 73), RandomNumber(18, 24), True)

        End If

    End If

    If UserList(UserIndex).flags.Privilegios = PlayerType.User Then
        If UserList(UserIndex).pos.Map = 47 Then
            Call WarpUserChar(UserIndex, 34, 30, 50, True)

        End If

    End If

    If UserList(UserIndex).pos.Map = 79 And UserList(UserIndex).flags.automatico = False Then
        Call WarpUserChar(UserIndex, 1, 45, 49, True)

    End If

    If UserList(UserIndex).flags.bandas = True Then
        Call Ban_Desconecta(UserIndex)
    End If

    If UserList(UserIndex).flags.medusas = True Then
        Call Med_Desconecta(UserIndex)
    End If

    If UserList(UserIndex).flags.EnDosVDos = True Then
        Call CerroEnDuelo(UserIndex)
    End If

    If UserList(UserIndex).flags.Montado = True Then
        UserList(UserIndex).char.Body = UserList(UserIndex).flags.NumeroMont
        '[MaTeO 9]
        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList( _
                                                                                                                        UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, _
                            UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)
        '[/MaTeO 9]
        UserList(UserIndex).flags.NumeroMont = 0
        UserList(UserIndex).flags.Montado = False

    End If

    If UserList(UserIndex).pos.Map = MAPADUELO And UserIndex = duelosespera Then
        Call WarpUserChar(UserIndex, 34, 30, 50, True)
        Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos: " & UserList(UserIndex).Name & " ha salido de la sala de torneos." & FONTTYPE_TALK)
        duelosespera = duelosreta
        numduelos = 0

    End If

    If UserList(UserIndex).pos.Map = MAPADUELO And UserIndex = duelosreta Then
        Call WarpUserChar(UserIndex, 34, 30, 50, True)
        Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos: " & UserList(UserIndex).Name & " ha salido de la sala de torneos." & FONTTYPE_TALK)

    End If

    If UserList(UserIndex).pos.Map = 76 Then
        Call WarpUserChar(UserIndex, 34, 30, 50, True)

    End If

    If UserIndex = Team.Pj1 Or UserIndex = Team.Pj2 Then
        Team.SonDos = False
        Team.Pj1 = 0
        Team.Pj2 = 0

    End If

    If UserList(UserIndex).flags.EstaDueleando = True Then
        Call DesconectarDuelo(UserList(UserIndex).flags.Oponente, UserIndex)

    End If


    '////////////////////////////////////////////////////////////////////////////////////////
    'Call SecurityIp.IpRestarConexion(GetLongIp(UserList(UserIndex).ip))

    If UserList(UserIndex).ConnID <> -1 Then
        Call CloseSocketSL(UserIndex)

    End If

    'Es el mismo user al que está revisando el centinela??
    'IMPORTANTE!!! hacerlo antes de resetear así todavía sabemos el nombre del user
    ' y lo podemos loguear
    If Centinela.RevisandoUserIndex = UserIndex Then Call modCentinela.CentinelaUserLogout

    'mato los comercios seguros
    If UserList(UserIndex).ComUsu.DestUsu > 0 Then
        If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
            If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call SendData(SendTarget.ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)

            End If

        End If

    End If

    If UserList(UserIndex).flags.UserLogged Then
        If NumUsers > 0 Then NumUsers = NumUsers - 1
        Call CloseUser(UserIndex)
    Else
        Call ResetUserSlot(UserIndex)

    End If

    For i = 1 To NumUsers
        If UserList(i).flags.Privilegios = PlayerType.User Then
            Total = Total + 1
        End If
    Next i

    Call SendData(ToAll, 0, 0, "³" & Total)
    UserList(UserIndex).flags.EnDosVDos = False
    UserList(UserIndex).flags.envioSol = False
    UserList(UserIndex).flags.RecibioSol = False
    UserList(UserIndex).flags.ParejaMuerta = False
    UserList(UserIndex).flags.EsperandoDuelo1 = False
    UserList(UserIndex).flags.Oponente1 = 0
    UserList(UserIndex).flags.EstaDueleando1 = False
    UserList(UserIndex).flags.EsperandoDuelo = False
    UserList(UserIndex).flags.Oponente = 0
    UserList(UserIndex).flags.EstaDueleando = False
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False
    
    Call SaveConfig

    #If MYSQL = 1 Then
        Call Add_DataBase(UserIndex, "Online")
    #End If


    Exit Sub

errhandler:
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False

    Call ResetUserSlot(UserIndex)

    If UserList(UserIndex).ConnID <> -1 Then
        Call CloseSocketSL(UserIndex)

    End If

    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.Description & " - UserIndex = " & UserIndex)

End Sub

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal UserIndex As Integer)

    If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
        Call BorraSlotSock(UserList(UserIndex).ConnID)
        Call WSApiCloseSocket(UserList(UserIndex).ConnID)
        UserList(UserIndex).ConnIDValida = False

    End If

End Sub

Public Function EnviarDatosASlot(ByVal UserIndex As Integer, Datos As String) As Long

    On Error GoTo Err

    Dim Ret As Long

    Ret = WsApiEnviar(UserIndex, Datos)

    If Ret <> 0 And Ret <> WSAEWOULDBLOCK Then
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)

    End If

    EnviarDatosASlot = Ret
    Exit Function

Err:

    Call LogError("TCP::EnviarDatosASlot. UI/ConnId/Datos: " & UserIndex & "/" & UserList(UserIndex).ConnID & "/" & Datos)

End Function

Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal sndData As String)

    On Error Resume Next

    Dim LoopC As Integer
    Dim x As Integer
    Dim Y As Integer
    
    
    sndData = AoDefEncode(AoDefServEncrypt(sndData))
    sndData = sndData & ENDC

    Select Case sndRoute

    Case SendTarget.ToPCArea

        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1

                If InMapBounds(sndMap, x, Y) Then
                    If MapData(sndMap, x, Y).UserIndex > 0 Then
                        If UserList(MapData(sndMap, x, Y).UserIndex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, x, Y).UserIndex, sndData)

                        End If

                    End If

                End If

            Next x
        Next Y

        Exit Sub

    Case SendTarget.ToIndex

        If UserList(sndIndex).ConnID <> -1 Then
            Call EnviarDatosASlot(sndIndex, sndData)
            Exit Sub

        End If

    Case SendTarget.ToNone
        Exit Sub

    Case SendTarget.ToAdmins

        For LoopC = 1 To LastUser

            If UserList(LoopC).ConnID <> -1 Then
                If UserList(LoopC).flags.Privilegios > 0 Then
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

            End If

        Next LoopC

        Exit Sub

    Case SendTarget.ToAll

        For LoopC = 1 To LastUser

            If UserList(LoopC).ConnID <> -1 Then
                If UserList(LoopC).flags.UserLogged Then    'Esta logeado como usuario?
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

            End If

        Next LoopC

        Exit Sub

    Case SendTarget.ToAllButIndex

        For LoopC = 1 To LastUser

            If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
                If UserList(LoopC).flags.UserLogged Then    'Esta logeado como usuario?
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

            End If

        Next LoopC

        Exit Sub

    Case SendTarget.ToMap

        For LoopC = 1 To LastUser

            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).pos.Map = sndMap Then
                        Call EnviarDatosASlot(LoopC, sndData)

                    End If

                End If

            End If

        Next LoopC

        Exit Sub

    Case SendTarget.ToMapButIndex

        For LoopC = 1 To LastUser

            If (UserList(LoopC).ConnID <> -1) And LoopC <> sndIndex Then
                If UserList(LoopC).pos.Map = sndMap Then
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

            End If

        Next LoopC

        Exit Sub

    Case SendTarget.ToGuildMembers

        LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

        While LoopC > 0

            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, sndData)

            End If

            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        Wend

        Exit Sub

    Case SendTarget.ToDeadArea

        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1

                If InMapBounds(sndMap, x, Y) Then
                    If MapData(sndMap, x, Y).UserIndex > 0 Then
                        If UserList(MapData(sndMap, x, Y).UserIndex).flags.Muerto = 1 Or UserList(MapData(sndMap, x, _
                                                                                                          Y).UserIndex).flags.Privilegios >= 1 Then

                            If UserList(MapData(sndMap, x, Y).UserIndex).ConnID <> -1 Then
                                Call EnviarDatosASlot(MapData(sndMap, x, Y).UserIndex, sndData)

                            End If

                        End If

                    End If

                End If

            Next x
        Next Y

        Exit Sub

        '[Alejo-18-5]
    Case SendTarget.ToPCAreaButIndex

        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1

                If InMapBounds(sndMap, x, Y) Then
                    If (MapData(sndMap, x, Y).UserIndex > 0) And (MapData(sndMap, x, Y).UserIndex <> sndIndex) Then

                        If UserList(MapData(sndMap, x, Y).UserIndex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, x, Y).UserIndex, sndData)

                        End If

                    End If

                End If

            Next x
        Next Y

        Exit Sub

    Case SendTarget.ToClanArea

        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1

                If InMapBounds(sndMap, x, Y) Then
                    If (MapData(sndMap, x, Y).UserIndex > 0) Then
                        If UserList(MapData(sndMap, x, Y).UserIndex).ConnID <> -1 Then
                            If UserList(sndIndex).GuildIndex > 0 And UserList(MapData(sndMap, x, Y).UserIndex).GuildIndex = UserList( _
                               sndIndex).GuildIndex Then
                                Call EnviarDatosASlot(MapData(sndMap, x, Y).UserIndex, sndData)

                            End If

                        End If

                    End If

                End If

            Next x
        Next Y

        Exit Sub

    Case SendTarget.ToPartyArea

        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1

                If InMapBounds(sndMap, x, Y) Then
                    If (MapData(sndMap, x, Y).UserIndex > 0) Then
                        If UserList(MapData(sndMap, x, Y).UserIndex).ConnID <> -1 Then
                            If UserList(sndIndex).PartyIndex > 0 And UserList(MapData(sndMap, x, Y).UserIndex).PartyIndex = UserList( _
                               sndIndex).PartyIndex Then
                                Call EnviarDatosASlot(MapData(sndMap, x, Y).UserIndex, sndData)

                            End If

                        End If

                    End If

                End If

            Next x
        Next Y

        Exit Sub

        '[CDT 17-02-2004]
    Case SendTarget.ToAdminsAreaButConsejeros

        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1

                If InMapBounds(sndMap, x, Y) Then
                    If (MapData(sndMap, x, Y).UserIndex > 0) And (MapData(sndMap, x, Y).UserIndex <> sndIndex) Then

                        If UserList(MapData(sndMap, x, Y).UserIndex).ConnID <> -1 Then
                            If UserList(MapData(sndMap, x, Y).UserIndex).flags.Privilegios > 1 Then
                                Call EnviarDatosASlot(MapData(sndMap, x, Y).UserIndex, sndData)

                            End If

                        End If

                    End If

                End If

            Next x
        Next Y

        Exit Sub
        '[/CDT]

    Case SendTarget.ToNPCArea

        For Y = Npclist(sndIndex).pos.Y - MinYBorder + 1 To Npclist(sndIndex).pos.Y + MinYBorder - 1
            For x = Npclist(sndIndex).pos.x - MinXBorder + 1 To Npclist(sndIndex).pos.x + MinXBorder - 1

                If InMapBounds(sndMap, x, Y) Then
                    If MapData(sndMap, x, Y).UserIndex > 0 Then
                        If UserList(MapData(sndMap, x, Y).UserIndex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, x, Y).UserIndex, sndData)

                        End If

                    End If

                End If

            Next x
        Next Y

        Exit Sub
        'Call SendToNpcArea(sndIndex, sndData)

        Exit Sub

    Case SendTarget.ToDiosesYclan
        LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

        While LoopC > 0

            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, sndData)

            End If

            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        Wend

        LoopC = modGuilds.Iterador_ProximoGM(sndIndex)

        While LoopC > 0

            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, sndData)

            End If

            LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
        Wend

        Exit Sub

    Case SendTarget.ToConsejo

        For LoopC = 1 To LastUser

            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.PertAlCons > 0 Then
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

            End If

        Next LoopC

        Exit Sub

    Case SendTarget.ToConsejoCaos

        For LoopC = 1 To LastUser

            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.PertAlConsCaos > 0 Then
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

            End If

        Next LoopC

        Exit Sub

    Case SendTarget.ToRolesMasters

        For LoopC = 1 To LastUser

            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

            End If

        Next LoopC

        Exit Sub

    Case SendTarget.ToCiudadanos

        For LoopC = 1 To LastUser

            If (UserList(LoopC).ConnID <> -1) Then
                If Not Criminal(LoopC) Then
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

            End If

        Next LoopC

        Exit Sub

    Case SendTarget.ToCriminales

        For LoopC = 1 To LastUser

            If (UserList(LoopC).ConnID <> -1) Then
                If Criminal(LoopC) Then
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

            End If

        Next LoopC

        Exit Sub

    Case SendTarget.ToReal

        For LoopC = 1 To LastUser

            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).Faccion.ArmadaReal = 1 Then
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

            End If

        Next LoopC

        Exit Sub

    Case SendTarget.ToCaos

        For LoopC = 1 To LastUser

            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).Faccion.FuerzasCaos = 1 Then
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

            End If

        Next LoopC

        Exit Sub

    Case ToCiudadanosYRMs

        For LoopC = 1 To LastUser

            If (UserList(LoopC).ConnID <> -1) Then
                If Not Criminal(LoopC) Or UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

            End If

        Next LoopC

        Exit Sub

    Case ToCriminalesYRMs

        For LoopC = 1 To LastUser

            If (UserList(LoopC).ConnID <> -1) Then
                If Criminal(LoopC) Or UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

            End If

        Next LoopC

        Exit Sub

    Case ToRealYRMs

        For LoopC = 1 To LastUser

            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).Faccion.ArmadaReal = 1 Or UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

            End If

        Next LoopC

        Exit Sub

    Case ToCaosYRMs

        For LoopC = 1 To LastUser

            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).Faccion.FuerzasCaos = 1 Or UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

            End If

        Next LoopC

        Exit Sub

    End Select

End Sub

Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean

    Dim x As Integer, Y As Integer

    For Y = UserList(Index).pos.Y - MinYBorder + 1 To UserList(Index).pos.Y + MinYBorder - 1
        For x = UserList(Index).pos.x - MinXBorder + 1 To UserList(Index).pos.x + MinXBorder - 1

            If MapData(UserList(Index).pos.Map, x, Y).UserIndex = Index2 Then
                EstaPCarea = True
                Exit Function

            End If

        Next x
    Next Y

    EstaPCarea = False

End Function

Function HayPCarea(pos As WorldPos) As Boolean

    Dim x As Integer, Y As Integer

    For Y = pos.Y - MinYBorder + 1 To pos.Y + MinYBorder - 1
        For x = pos.x - MinXBorder + 1 To pos.x + MinXBorder - 1

            If x > 0 And Y > 0 And x < 101 And Y < 101 Then
                If MapData(pos.Map, x, Y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function

                End If

            End If

        Next x
    Next Y

    HayPCarea = False

End Function

Function HayOBJarea(pos As WorldPos, ObjIndex As Integer) As Boolean

    Dim x As Integer, Y As Integer

    For Y = pos.Y - MinYBorder + 1 To pos.Y + MinYBorder - 1
        For x = pos.x - MinXBorder + 1 To pos.x + MinXBorder - 1

            If MapData(pos.Map, x, Y).OBJInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function

            End If

        Next x
    Next Y

    HayOBJarea = False

End Function

Function ValidateChr(ByVal UserIndex As Integer) As Boolean

    ValidateChr = UserList(UserIndex).char.Head <> 0 And UserList(UserIndex).char.Body <> 0 And ValidateSkills(UserIndex)

End Function

Sub ConnectUser(ByVal UserIndex As Integer, Name As String, Password As String, ByVal hdString As String)
    Dim n As Integer
    Dim tStr As String
    Dim x As Integer
    Dim Total As Integer


    'Reseteamos los FLAGS
    With UserList(UserIndex)

        .flags.Escondido = 0
        .flags.TargetNpc = 0
        .flags.TargetNpcTipo = eNPCType.Comun
        .flags.TargetObj = 0
        .flags.TargetUser = 0
        .Counters.Invisibilidad = IntervaloInvisible
        .char.FX = 0

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "INVI0")

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||       .oo .oPYo. o     o               o              .oPYo.    .oPYo.       .oo" & _
                                                        FONTTYPE_Motd1)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||       .P 8 8    8 8b   d8                              8  .o8    8  .o8      .P 8" & _
                                                        FONTTYPE_Motd1)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||      .P  8 8    8 8`b d'8 .oPYo. odYo. o8 .oPYo.       8 .P'8    8 .P'8     .P  8" & _
                                                        FONTTYPE_Motd1)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||     oPooo8 8    8 8 `o' 8 .oooo8 8' `8  8 .oooo8       8.d' 8    8.d' 8         8" & _
                                                        FONTTYPE_Motd1)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||    .P    8 8    8 8     8 8    8 8   8  8 8    8       8o'  8    8o'  8         8" & _
                                                        FONTTYPE_Motd1)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||   .P     8 `YooP' 8     8 `YooP8 8   8  8 `YooP8       `YooP' 88 `YooP'88       8" & _
                                                        FONTTYPE_Motd1)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||                                                                  www.AoMania.Net" & _
                                                        FONTTYPE_Motd1)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||      X = Seguro Objetos   Q = Mapa   P = Mapa  S = Seguro   W = Seguro de clan" & _
                                                        FONTTYPE_Motd2)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||                         F1 = /Meditar   F12 = Macro Interno Para trabajadores." & _
                                                        FONTTYPE_Motd2)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||                     Si tienes alguna duda o necesitas ayuda, escribe /GM TEXTO" & _
                                                        FONTTYPE_Motd2)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||                                                        Version 0.0.1 Año: 2019" & _
                                                        FONTTYPE_Motd2)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||:- Argentumania -:" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||> ¡¡¡Bienvenidos al Servidor Oficial AoManiA 2019!!!" & FONTTYPE_GUILD)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||> Versión Actual v1 de AOMania, Argentumania 2018. Mod Argentum Online" & FONTTYPE_Motd3)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||> Para cualquier duda, /gm consulta" & FONTTYPE_Motd3)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||> Web Oficial aomania.net argentumania.es" & FONTTYPE_Motd4)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||> Foro Oficial foro.argentumania.es" & FONTTYPE_Motd4)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||---------------" & FONTTYPE_SERVER)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Para ver el Mapa de AOMania dejar pulsada la tecla Q ó P." & FONTTYPE_Motd5)

        Call SendInfoCastillos(UserIndex)

        If MaxLevel > 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| " & FONTTYPE_INFO)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El máximo nivel es " & MaxLevel & ", adquirido por " & UserMaxLevel & "." & _
                                                            FONTTYPE_SERVER)

        End If

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| " & FONTTYPE_INFO)

        If MultMsg = "" Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La Experiencia del Servidor esta subido por " & Multexp & "." & FONTTYPE_Motd5)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El Oro del Servidor esta subido por " & MultOro & "." & FONTTYPE_Motd5)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La Experiencia del Servidor esta subido por " & Multexp & ". " & MultMsg & "." & _
                                                            FONTTYPE_Motd5)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El Oro del Servidor esta subido por " & MultOro & ". " & MultMsg & "." & _
                                                            FONTTYPE_Motd5)

        End If

        If StatusNosfe = True Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Nosferatu esta haciendo estragos en el mapa " & MapaNosfe & FONTTYPE_GUILD)
        End If

        If ExpCriatura = True Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Hoy es día de " & NombreCriatura & ", su experencia esta aumentada x" & _
                                                            LoteriaCriatura & "." & FONTTYPE_TALK)

        End If

        If OroCriatura = True Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Hoy es día de " & NombreCriatura & ", su oro esta aumentada x" & LoteriaCriatura & _
                                                            "." & FONTTYPE_TALK)

        End If

        If DiaEspecialExp = True Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Estáis de suerte! Día especial, la experencia esta aumentada por x2" & FONTTYPE_TALK)

        End If

        If DiaEspecialOro = True Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Estáis de suerte! Día especial, el oro esta aumentada por x2" & FONTTYPE_TALK)

        End If

        'Controlamos no pasar el maximo de usuarios
        If NumUsers >= MaxUsers Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, _
                          "ERREl servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        '¿Este IP ya esta conectado?
        If AllowMultiLogins = 0 Then
            If CheckForSameIP(UserIndex, .ip) = True Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRNo es posible usar mas de un personaje al mismo tiempo.")
                Call CloseSocket(UserIndex)
                Exit Sub

            End If

        End If

        '¿Existe el personaje?
        If Not FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERREl personaje no existe.")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        '¿Es el passwd valido?
        If UCase$(Password) <> UCase$(GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password")) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRPassword incorrecto.")

            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        '¿Ya esta conectado el personaje?
        If CheckForSameName(UserIndex, Name) Then
            If UserList(NameIndex(Name)).Counters.Saliendo Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERREl usuario está saliendo.")
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRPerdon, un usuario con el mismo nombre se há logoeado.")

            End If

            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        'Cargamos el personaje
        Dim Leer As New clsIniManager

        Call Leer.Initialize(CharPath & UCase$(Name) & ".chr")

        'Cargamos los datos del personaje
        Call LoadUserInit(UserIndex, Leer)

        Call LoadUserStats(UserIndex, Leer)

        If Not ValidateChr(UserIndex) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRError en el personaje.")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        Call LoadUserReputacion(UserIndex, Leer)

        Set Leer = Nothing

        If .Invent.EscudoEqpSlot = 0 Then .char.ShieldAnim = NingunEscudo
        If .Invent.CascoEqpSlot = 0 Then .char.CascoAnim = NingunCasco
        If .Invent.WeaponEqpSlot = 0 Then .char.WeaponAnim = NingunArma

        Call UpdateUserInv(True, UserIndex, 0)
        Call UpdateUserHechizos(True, UserIndex, 0)

        If .flags.Navegando = 1 Then
            .char.Body = ObjData(.Invent.BarcoObjIndex).Ropaje
            .char.Head = 0
            .char.WeaponAnim = NingunArma
            .char.ShieldAnim = NingunEscudo
            .char.CascoAnim = NingunCasco
            '[MaTeO 9]
            .char.Alas = NingunAlas
            '[/MaTeO 9]

        End If

        If .flags.Paralizado Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "PARADOW")

        End If

        'Feo, esto tiene que ser parche cliente
        If .flags.Estupidez = 0 Then Call SendData(SendTarget.ToIndex, UserIndex, 0, "NESTUP")

        'Posicion de comienzo
        If .pos.Map = 0 Then
            If .Stats.ELV < 13 Then
                .pos.Map = "37"
                .pos.x = "35"
                .pos.Y = "69"
            Else
                .pos.Map = "34"
                .pos.x = "30"
                .pos.Y = "50"

            End If

        Else

            'Anti Pisadas
            If MapData(.pos.Map, .pos.x, .pos.Y).UserIndex <> 0 Then
                Dim nPos As WorldPos
                Call ClosestStablePos(.pos, nPos)

                If nPos.x <> 0 And nPos.Y <> 0 Then
                    .pos.Map = nPos.Map
                    .pos.x = nPos.x
                    .pos.Y = nPos.Y

                End If

            End If

            'Anti Pisadas

            '            ''TELEFRAG
            '            If MapData(.pos.Map, .pos.x, .pos.y).UserIndex <> 0 Then
            '
            '                ''si estaba en comercio seguro...
            '                If UserList(MapData(.pos.Map, .pos.x, .pos.y).UserIndex).ComUsu.DestUsu > 0 Then
            '
            '                    If UserList(UserList(MapData(.pos.Map, .pos.x, _
                                 '                            .pos.y).UserIndex).ComUsu.DestUsu).flags.UserLogged Then
            '
            '                        Call FinComerciarUsu(UserList(MapData(.pos.Map, .pos.x, _
                                     '                                .pos.y).UserIndex).ComUsu.DestUsu)
            '                        Call SendData(SendTarget.ToIndex, UserList(MapData(.pos.Map, .pos.x, _
                                     '                                .pos.y).UserIndex).ComUsu.DestUsu, 0, _
                                     '                                "||Comercio cancelado. El otro usuario se ha desconectado." & FONTTYPE_TALK)
            '
            '                    End If
            '
            '                End If
            '
            '                If UserList(MapData(.pos.Map, .pos.x, .pos.y).UserIndex).flags.UserLogged Then
            '                    Call FinComerciarUsu(MapData(.pos.Map, .pos.x, .pos.y).UserIndex)
            '
            '                End If
            '
            '                Call CloseSocket(MapData(.pos.Map, .pos.x, .pos.y).UserIndex)
            '
            '            End If

            If .flags.Muerto = 1 Then
                Call Empollando(UserIndex)

            End If

            If .flags.Embarcado = 1 Then
                If Barcos.TiempoRest > 60 Then
                    If .Zona <= NumZonas Then
                        .pos.Map = Zonas(.Zona).Map
                        .pos.Y = Zonas(.Zona).Y
                        .pos.x = Zonas(.Zona).x
                        .flags.Embarcado = 0

                    End If

                End If

            End If

        End If

        If Not MapaValido(.pos.Map) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERREL PJ se encuenta en un mapa invalido.")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        'Nombre de sistema
        .Name = Name
        .Password = Password
        .hd_String = hdString

        Call WriteVar(CharPath & .Name & ".chr", "INIT", "LastHD", .hd_String)

        .showName = True    'Por default los nombres son visibles

        'Info
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "IU" & UserIndex)    'Enviamos el User index
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "CM" & .pos.Map & "," & MapInfo(.pos.Map).MapVersion)    'Carga el mapa
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "TM" & MapInfo(.pos.Map).Music)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "N~" & MapInfo(.pos.Map).Name)

        'Vemos que clase de user es (se lo usa para setear los privilegios alcrear el PJ)
        .flags.EsRolesMaster = EsRolesMaster(Name)

        If EsAdmin(Name) Then
            .flags.Privilegios = PlayerType.Admin
            Call LogGM(.Name, "Se conecto con ip:" & .ip & "Hora: " & now)
        ElseIf EsDios(Name) Then
            .flags.Privilegios = PlayerType.Dios
            Call LogGM(.Name, "Se conecto con ip:" & .ip & "Hora: " & now)
        ElseIf EsSemiDios(Name) Then
            .flags.Privilegios = PlayerType.SemiDios
            Call LogGM(.Name, "Se conecto con ip:" & .ip & "Hora: " & now)
        ElseIf EsConsejero(Name) Then
            .flags.Privilegios = PlayerType.Consejero
            Call LogGM(.Name, "Se conecto con ip:" & .ip & "Hora: " & now)
        Else
            .flags.Privilegios = PlayerType.User

        End If

        ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
        .Counters.IdleCount = 0

        'Crea  el personaje del usuario
        Call MakeUserChar(SendTarget.ToMap, 0, .pos.Map, UserIndex, .pos.Map, .pos.x, .pos.Y)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "IP" & .char.CharIndex)
        ''[/el oso]

        Call SendUserStatsBox(UserIndex)
        Call SendUserHitBox(UserIndex)
        Call EnviarHambreYsed(UserIndex)
        Call EnviarAmarillas(UserIndex)
        Call EnviarVerdes(UserIndex)

        'If haciendoBK Then
        '    Call SendData(SendTarget.ToIndex, UserIndex, 0, "BKW")
        '    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||AOMania> Por favor espera algunos segundos, WorldSave esta ejecutandose." & _
             FONTTYPE_SERVER)

        'End If

        If EnPausa Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "BKW")
            Call SendData(SendTarget.ToIndex, UserIndex, 0, _
                          "||AOMania> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde." & _
                          FONTTYPE_SERVER)

        End If

        If EnTesting And .Stats.ELV >= 18 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, _
                          "ERRServidor en Testing por unos minutos, conectese con PJs de nivel menor a 18. No se conecte con Pjs que puedan resultar importantes por ahora pues pueden arruinarse.")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        'Actualiza el Num de usuarios
        'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
        NumUsers = NumUsers + 1
        .flags.UserLogged = True

        For x = 1 To NumUsers
            If UserList(x).flags.Privilegios = PlayerType.User Then
                Total = Total + 1
            End If
        Next x

        Call SendData(ToAll, 0, 0, "³" & Total)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "HUCT" & CountTC)

        If NocheLicantropo = True Then
            If UserList(UserIndex).flags.Privilegios = PlayerType.User And _
               UCase$(UserList(UserIndex).Raza) = "LICANTROPO" And _
               UserList(UserIndex).flags.Licantropo = "0" Then
                Call DarPoderLicantropo(UserIndex)
            End If
        Else
            If UserList(UserIndex).flags.Privilegios = PlayerType.User And _
               UCase$(UserList(UserIndex).Raza) = "LICANTROPO" And _
               UserList(UserIndex).flags.Licantropo = "1" Then
                Call QuitarPoderLicantropo(UserIndex)
            End If
        End If

        'usado para borrar Pjs
        Call WriteVar(CharPath & .Name & ".chr", "INIT", "Logged", "1")

        'Call SendData(ToAll, 0, 0, "³" & NumUsers)
        MapInfo(.pos.Map).NumUsers = MapInfo(.pos.Map).NumUsers + 1

        If .Stats.SkillPts > 0 Then
            Call EnviarSkills(UserIndex)
            Call EnviarSubirNivel(UserIndex, .Stats.SkillPts)

        End If

        If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers

        If Total > recordusuarios Then
            Call SendData(SendTarget.ToAll, 0, 0, "||Record de usuarios conectados simultaneamente." & "Hay " & Total & " usuarios." & _
                                                  FONTTYPE_TURQ)
            recordusuarios = Total
            Call WriteVar(IniPath & "Server.ini", "INIT", "Record", CStr(recordusuarios))
        End If

        If .NroMacotas > 0 Then
            Dim i As Integer

            For i = 1 To MAXMASCOTAS

                If .MascotasType(i) > 0 Then
                    .MascotasIndex(i) = SpawnNpc(.MascotasType(i), .pos, True, True)

                    If .MascotasIndex(i) > 0 Then
                        Npclist(.MascotasIndex(i)).MaestroUser = UserIndex
                        Call FollowAmo(.MascotasIndex(i))
                    Else
                        .MascotasIndex(i) = 0

                    End If

                End If

            Next i

        End If

        If .flags.Navegando = 1 Then Call SendData(SendTarget.ToIndex, UserIndex, 0, "NAVEG")

        If Criminal(UserIndex) Then
            'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Miembro de las fuerzas del caos > Seguro desactivado <" & FONTTYPE_FIGHT)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "OFFOFS")
            .flags.Seguro = False
        Else
            .flags.Seguro = True
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ONONS")

        End If

        .flags.SeguroClan = False
        .flags.SeguroCombate = False
        .flags.SeguroHechizos = True
        .flags.SeguroObjetos = False

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGCO99")
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEG11")
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEG13")
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEG15")
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGCVCON")
        UserList(UserIndex).flags.SeguroCVC = True

        If ServerSoloGMs > 0 Then
            If .flags.Privilegios < ServerSoloGMs Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRServidor restringido a administradores de jerarquia mayor o igual a: " & _
                                                                ServerSoloGMs & ". Por favor intente en unos momentos.")
                Call CloseSocket(UserIndex)
                Exit Sub

            End If

        End If

        If .GuildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, 0, "||" & .Name & " se ha conectado." & FONTTYPE_GUILD)
                 
               Call EnviaPosClan(UserIndex)
               
            If Not modGuilds.m_ConectarMiembroAClan(UserIndex, .GuildIndex) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu estado no te permite entrar al clan." & FONTTYPE_GUILD)

            End If

        End If

        Call SendData(SendTarget.ToPCArea, UserIndex, .pos.Map, "CFX" & .char.CharIndex & "," & FXIDs.FXWARP & "," & 0)

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "LODXXD")

        Call modGuilds.SendGuildNews(UserIndex)

        Call SendMainAmbient(UserIndex)
        Call SendSecondaryAmbient(UserIndex)

        tStr = modGuilds.a_ObtenerRechazoDeChar(.Name)

        If tStr <> vbNullString Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "!!Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr _
                                                          & ENDC)

        End If

        n = FreeFile
        Open App.Path & "\logs\numusers.log" For Output As n
        Print #n, NumUsers
        Close #n

        n = FreeFile
        'Log
        Open App.Path & "\logs\Connect.log" For Append Shared As #n
        Print #n, Time; Date & " " & .Name & " ha entrado al juego. UserIndex: " & UserIndex & " IP: " & .ip & " HDD: " & .hd_String
        Close #n

        If .pos.Map = MAPADUELO Then
            Call WarpUserChar(UserIndex, 34, 30, 50, True)
        End If

        If .pos.Map = MapaFuerte Then
            Call ConnectFuerte(UserIndex)
        End If

        If .flags.Privilegios = PlayerType.User Then
            If MaxLevel = 0 Then
                MaxLevel = 1
                UserMaxLevel = .Name
            End If
        End If

        Call CriCiuMaxLvl(UserIndex)
        Call CountCriCi(UserIndex)
        Call MaxOroRank(UserIndex)
        Call OroConnectRank(UserIndex)
        Call ConnectQuest(UserIndex)

        #If MYSQL = 1 Then
            Call Add_DataBase(UserIndex, "Online")
        #End If

        DoEvents

    End With

End Sub

Sub SendMOTD(ByVal UserIndex As Integer)
    Dim j As Long

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||       .oo .oPYo. o     o               o              .oPYo.    .oPYo.       .oo" & _
                                                    FONTTYPE_Motd1)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||       .P 8 8    8 8b   d8                              8  .o8    8  .o8      .P 8" & _
                                                    FONTTYPE_Motd1)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||      .P  8 8    8 8`b d'8 .oPYo. odYo. o8 .oPYo.       8 .P'8    8 .P'8     .P  8" & _
                                                    FONTTYPE_Motd1)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||     oPooo8 8    8 8 `o' 8 .oooo8 8' `8  8 .oooo8       8.d' 8    8.d' 8         8" & _
                                                    FONTTYPE_Motd1)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||    .P    8 8    8 8     8 8    8 8   8  8 8    8       8o'  8    8o'  8         8" & _
                                                    FONTTYPE_Motd1)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||   .P     8 `YooP' 8     8 `YooP8 8   8  8 `YooP8       `YooP' 88 `YooP'88       8" & _
                                                    FONTTYPE_Motd1)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||                                                                  www.AoMania.Net" & _
                                                    FONTTYPE_Motd1)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||      X = Seguro Objetos   Q = Mapa   P = Mapa  S = Seguro   W = Seguro de clan" & _
                                                    FONTTYPE_Motd2)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||                         F1 = /Meditar   F12 = Macro Interno Para trabajadores." & _
                                                    FONTTYPE_Motd2)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||                     Si tienes alguna duda o necesitas ayuda, escribe /GM TEXTO" & _
                                                    FONTTYPE_Motd2)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||                                                        Version 0.0.1 Año: 2019" & _
                                                    FONTTYPE_Motd2)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||:- Argentumania -:" & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||> ¡¡¡Bienvenidos al Servidor Oficial AoManiA 2019!!!" & FONTTYPE_GUILD)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||> Versión Actual v1 de AOMania, Argentumania 2018. Mod Argentum Online" & FONTTYPE_Motd3)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||> Para cualquier duda, /gm consulta" & FONTTYPE_Motd3)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||> Web Oficial aomania.net argentumania.es" & FONTTYPE_Motd4)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||> Foro Oficial foro.argentumania.es" & FONTTYPE_Motd4)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||---------------" & FONTTYPE_SERVER)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Para ver el Mapa de AOMania dejar pulsada la tecla Q ó P." & FONTTYPE_Motd5)

    Call SendInfoCastillos(UserIndex)

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| " & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El máximo nivel es " & MaxLevel & ", adquirido por " & UserMaxLevel & "." & FONTTYPE_SERVER)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| " & FONTTYPE_INFO)

    If MultMsg = "" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La Experiencia del Servidor esta subido por " & Multexp & "." & FONTTYPE_Motd5)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El Oro del Servidor esta subido por " & MultOro & "." & FONTTYPE_Motd5)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La Experiencia del Servidor esta subido por " & Multexp & ". " & MultMsg & "." & _
                                                        FONTTYPE_Motd5)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El Oro del Servidor esta subido por " & MultOro & ". " & MultMsg & "." & FONTTYPE_Motd5)

    End If

    'For j = 1 To MaxLines
    '    Call SendData(SendTarget.toindex, UserIndex, 0, "||" & Chr$(3) & MOTD(j).texto)
    'Next j

End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)

'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).Faccion
        .ArmadaReal = 0
        .FuerzasCaos = 0
        .Nemesis = 0
        .Templario = 0
        .CiudadanosMatados = 0
        .CriminalesMatados = 0
        .RecibioArmaduraCaos = 0
        .RecibioArmaduraReal = 0
        .RecibioExpInicialCaos = 0
        .RecibioExpInicialReal = 0
        .RecompensasCaos = 0
        .RecompensasReal = 0
        .RecibioArmaduraNemesis = 0
        .RecibioArmaduraTemplaria = 0
        .RecibioExpInicialNemesis = 0
        .RecibioExpInicialTemplaria = 0
        .RecompensasNemesis = 0
        .RecompensasTemplaria = 0
        .Reenlistadas = 0
        .ArmaduraFaccionaria = 0
        .NextRecompensas = 0

    End With

End Sub

Sub ResetContadores(ByVal UserIndex As Integer)

'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).Counters
        .AGUACounter = 0
        .AttackCounter = 0
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Paralisis = 0
        .Pasos = 0
        .Pena = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Trabajando = 0
        .Ocultando = 0

        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0

    End With

End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)

'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).char
        .Body = 0
        .CascoAnim = 0

        .CharIndex = 0
        .FX = 0
        .Head = 0
        .loops = 0
        .heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0

    End With

End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)

'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex)
        .Name = vbNullString
        .hd_String = vbNullString
        .Password = vbNullString
        .Desc = vbNullString
        .Zona = 0
        .Pareja = vbNullString
        .DescRM = vbNullString
        .pos.Map = 0
        .pos.x = 0
        .pos.Y = 0
        .ip = vbNullString
        .RDBuffer = vbNullString
        .Clase = vbNullString
        .Email = vbNullString
        .Genero = vbNullString
        .Hogar = vbNullString
        .Raza = vbNullString

        .EmpoCont = 0
        .PartyIndex = 0
        .PartySolicitud = 0

        With .Stats
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            .CriminalesMatados = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .SkillPts = 0

        End With

    End With

End Sub

Sub ResetReputacion(ByVal UserIndex As Integer)

'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).Reputacion
        .AsesinoRep = 0
        .BandidoRep = 0
        .BurguesRep = 0
        .LadronesRep = 0
        .NobleRep = 0
        .PlebeRep = 0
        .NobleRep = 0
        .Promedio = 0

    End With

End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)

    If UserList(UserIndex).EscucheClan > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(UserIndex, UserList(UserIndex).EscucheClan)
        UserList(UserIndex).EscucheClan = 0

    End If

    If UserList(UserIndex).GuildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(UserIndex, UserList(UserIndex).GuildIndex)

    End If

    UserList(UserIndex).GuildIndex = 0

End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)

'*************************************************
'Author: Unknown
'Last modified: 03/29/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'03/29/2006 Maraxus - Reseteo el CentinelaOK también.
'*************************************************
    With UserList(UserIndex).flags
        .Comerciando = False
        .Ban = 0
        .Escondido = 0
        .DuracionEfectoAmarillas = 0
        .DuracionEfectoVerdes = 0
        .NpcInv = 0
        .StatsChanged = 0
        .TargetNpc = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TomoPocionAmarilla = False
        .TomoPocionVerde = False
        .Descuento = ""
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .ClienteOK = False
        .Navegando = 0
        .Oculto = 0
        .Envenenado = 0
        .Invisible = 0
        .Paralizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .YaDenuncio = 0
        .Privilegios = PlayerType.User
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .Hechizo = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .EstaEmpo = 0
        .PertAlCons = 0
        .PertAlConsCaos = 0
        .EnDosVDos = False
        .ParejaMuerta = False
        .envioSol = False
        .RecibioSol = False
        .CentinelaOK = False
        .Soporteo = False
        .EstaDueleando1 = False
        .Oponente1 = 0
        .EsperandoDuelo1 = False
        .EstaDueleando = False
        .Oponente = 0
        .EsperandoDuelo = False
        .Embarcado = 0
        .Casado = 0
        .Casandose = False
        .Quien = 0

    End With

    UserList(UserIndex).Counters.AntiSH = 0

End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
    Dim LoopC As Long

    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
    Next LoopC

End Sub

Sub ResetUserPets(ByVal UserIndex As Integer)
    Dim LoopC As Long

    UserList(UserIndex).NroMacotas = 0

    For LoopC = 1 To MAXMASCOTAS
        UserList(UserIndex).MascotasIndex(LoopC) = 0
        UserList(UserIndex).MascotasType(LoopC) = 0
    Next LoopC

End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
    Dim LoopC As Long

    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
        UserList(UserIndex).BancoInvent.Object(LoopC).Amount = 0
        UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
        UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = 0
    Next LoopC

    UserList(UserIndex).BancoInvent.NroItems = 0

End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)

    With UserList(UserIndex).ComUsu

        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(UserIndex)

        End If

    End With

End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)

    Dim UsrTMP As User

    UserList(UserIndex).ConnIDValida = False
    UserList(UserIndex).ConnID = -1

    Call LimpiarComercioSeguro(UserIndex)
    Call ResetFacciones(UserIndex)
    Call ResetContadores(UserIndex)
    Call ResetCharInfo(UserIndex)
    Call ResetBasicUserInfo(UserIndex)
    Call ResetReputacion(UserIndex)
    Call ResetGuildInfo(UserIndex)
    Call ResetUserFlags(UserIndex)
    Call LimpiarInventario(UserIndex)
    Call ResetUserSpells(UserIndex)
    Call ResetUserPets(UserIndex)
    Call ResetUserBanco(UserIndex)

    With UserList(UserIndex).ComUsu
        .Acepto = False
        .Cant = 0
        .DestNick = ""
        .DestUsu = 0
        .Objeto = 0

    End With

    UserList(UserIndex) = UsrTMP

End Sub

Sub CloseUser(ByVal UserIndex As Integer)
'Call LogTarea("CloseUser " & UserIndex)

    On Error GoTo errhandler

    Dim n As Integer
    Dim x As Integer
    Dim Y As Integer
    Dim LoopC As Integer
    Dim Map As Integer
    Dim Name As String
    Dim Raza As String
    Dim Clase As String
    Dim i As Integer
     If UserList(UserIndex).flags.EnCvc = True Then
                UserList(UserIndex).flags.EnCvc = False
                WarpUserChar UserIndex, 34, 30, 50, True
            End If

    Dim aN As Integer

    aN = UserList(UserIndex).flags.AtacadoPorNpc

    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = ""

    End If

    UserList(UserIndex).flags.AtacadoPorNpc = 0

    Map = UserList(UserIndex).pos.Map
    x = UserList(UserIndex).pos.x
    Y = UserList(UserIndex).pos.Y
    Name = UCase$(UserList(UserIndex).Name)
    Raza = UserList(UserIndex).Raza
    Clase = UserList(UserIndex).Clase

    UserList(UserIndex).char.FX = 0
    UserList(UserIndex).char.loops = 0
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & 0 & "," & 0)

    UserList(UserIndex).flags.UserLogged = False
    UserList(UserIndex).Counters.Saliendo = False

    'Le devolvemos el body y head originales
    If UserList(UserIndex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(UserIndex)

    'si esta en party le devolvemos la experiencia
    If UserList(UserIndex).PartyIndex > 0 Then Call SalirDeParty(UserIndex)

    If UserList(UserIndex).flags.Metamorfosis = 1 Then
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).CharMimetizado.Fuerza
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).CharMimetizado.Agilidad
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).CharMimetizado.Inteligencia
    End If

    Call ResetFlagsAsedio(UserIndex)

    ' Grabamos el personaje del usuario
    Call SaveUser(UserIndex, CharPath & Name & ".chr")

    'usado para borrar Pjs
    Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "INIT", "Logged", "0")

    'Quitar el dialogo
    'If MapInfo(Map).NumUsers > 0 Then
    '    Call SendToUserArea(UserIndex, "QDL" & UserList(UserIndex).Char.charindex)
    'End If

    If MapInfo(Map).NumUsers > 0 Then
        Call SendData(SendTarget.ToMapButIndex, UserIndex, Map, "QDL" & UserList(UserIndex).char.CharIndex)

    End If

    'Borrar el personaje
    If UserList(UserIndex).char.CharIndex > 0 Then
        Call EraseUserChar(SendTarget.ToMap, UserIndex, Map, UserIndex)

    End If

    'Borrar mascotas
    For i = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasIndex(i) > 0 Then
            If Npclist(UserList(UserIndex).MascotasIndex(i)).flags.NPCActive Then Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))

        End If

    Next i

    'Update Map Users
    MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1

    If MapInfo(Map).NumUsers < 0 Then
        MapInfo(Map).NumUsers = 0

    End If

    ' Si el usuario habia dejado un msg en la gm's queue lo borramos
    If Ayuda.Existe(UserList(UserIndex).Name) Then Call Ayuda.Quitar(UserList(UserIndex).Name)
    If Quest.Existe(UserList(UserIndex).Name) Then Call Quest.Quitar(UserList(UserIndex).Name)
    If Torneo.Existe(UserList(UserIndex).Name) Then Call Torneo.Quitar(UserList(UserIndex).Name)



    n = FreeFile(1)
    With UserList(UserIndex)
        Open App.Path & "\logs\Connect.log" For Append Shared As #n
        Print #n, Time; Date & " " & .Name & " ha dejado el juego. UserIndex: " & UserIndex & " IP: " & .ip & " HDD: " & .hd_String
        Close #n
    End With

    Call ResetUserSlot(UserIndex)
    Exit Sub
    If UserList(UserIndex).EnCvc Then
            'Dim ijaji As Integer
            'For ijaji = 1 To LastUser
                With UserList(UserIndex)
                    If Guilds(.GuildIndex).GuildName = Nombre1 Then
                        If .EnCvc = True Then
                                modGuilds.UsuariosEnCvcClan1 = modGuilds.UsuariosEnCvcClan1 - 1
                                UserList(UserIndex).EnCvc = False
                                If modGuilds.UsuariosEnCvcClan1 = 0 And CvcFunciona = True Then
                                    Call SendData(SendTarget.ToAll, UserIndex, 0, "||" & "El clan " & Nombre2 & " derrotó al clan " & Nombre1 & "." & FONTTYPE_GUILD)
                                    CvcFunciona = False
                                    Call LlevarUsuarios
                                End If
                         End If
                     End If
                     
                    If Guilds(.GuildIndex).GuildName = Nombre2 Then
                        If .EnCvc = True Then
                                modGuilds.UsuariosEnCvcClan2 = modGuilds.UsuariosEnCvcClan2 - 1
                                UserList(UserIndex).EnCvc = False
                                If modGuilds.UsuariosEnCvcClan2 = 0 And CvcFunciona = True Then
                                    Call SendData(SendTarget.ToAll, UserIndex, 0, "||" & "El clan " & Nombre1 & " derrotó al clan " & Nombre2 & "." & FONTTYPE_GUILD)
                                    CvcFunciona = False
                                    Call LlevarUsuarios
                                End If
                        End If
                    End If
                End With
            'Next ijaji
    End If
    

errhandler:
    Call LogError("Error en CloseUser. Número " & Err.Number & " Descripción: " & Err.Description)

End Sub

Sub HandleData(ByVal UserIndex As Integer, ByVal rData As String)

    On Error GoTo ErrorHandler:

    Dim CadenaOriginal As String

    Dim LoopC          As Integer

    Dim nPos           As WorldPos

    Dim tStr           As String

    Dim tInt           As Integer

    Dim tLong          As Long

    Dim TIndex         As Integer

    Dim tName          As String

    Dim tMessage       As String

    Dim AuxInd         As Integer

    Dim Arg1           As String

    Dim Arg2           As String

    Dim Arg3           As String

    Dim Arg4           As String

    Dim Arg5           As String

    Dim Ver            As String

    Dim encpass        As String

    Dim Pass           As String

    Dim Mapa           As Integer

    Dim Name           As String

    Dim ind

    Dim n                  As Integer

    Dim wpaux              As WorldPos

    Dim mifile             As Integer

    Dim x                  As Integer

    Dim Y                  As Integer

    Dim DummyInt           As Integer

    Dim T()                As String

    Dim i                  As Integer

    Dim sndData            As String

    Dim cliMD5             As String

    Dim ClientChecksum     As String

    Dim ServerSideChecksum As Long

    Dim IdleCountBackup    As Long

    Dim hdStr              As String

    Dim tPath              As String
    
    UserList(UserIndex).clave2 = UserList(UserIndex).clave2 + 1

    With AodefConv
        SuperClave = .Numero2Letra(UserList(UserIndex).clave2, , 2, "ZiPPy", "NoPPy", 1, 0)

    End With

    Do While InStr(1, SuperClave, " ")
        SuperClave = mid$(SuperClave, 1, InStr(1, SuperClave, " ") - 1) & mid$(SuperClave, InStr(1, SuperClave, " ") + 1)
    Loop
    SuperClave = Semilla(SuperClave)
    UserList(UserIndex).clave = SuperClave
          
    If UserList(UserIndex).clave2 = 999998 Then
        UserList(UserIndex).clave2 = 0

    End If

    rData = DeCodificar(AoDefDecode(rData), UserList(UserIndex).clave)

    CadenaOriginal = rData

    '¿Tiene un indece valido?
    If UserIndex <= 0 Then
        Call CloseSocket(UserIndex)
        Exit Sub

    End If

    Select Case Left$(rData, 1)

        Case ";"

            With UserList(UserIndex)

                Dim MsgData As String

                MsgData = Right$(rData, Len(rData) - 1)

                If .flags.Privilegios = PlayerType.Dios Then
                    Call LogGM(.Name, "Dijo: " & MsgData)
                ElseIf .flags.Privilegios = PlayerType.SemiDios Then
                    Call LogGM(.Name, "Dijo: " & MsgData)
                ElseIf .flags.Privilegios = PlayerType.Consejero Then
                    Call LogGM(.Name, "Dijo: " & MsgData)
                ElseIf .flags.Privilegios = PlayerType.User Then
                    Call LogUser(.Name, "Dijo: " & MsgData)

                End If
            
                If .Quest.Start = 1 Then
                    If .Quest.ValidNpcDescubre = 1 Then
                        Call RespuestaNpcQuest(UserIndex, .Quest.Quest, MsgData)

                    End If

                End If

            End With

    End Select

    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    IdleCountBackup = UserList(UserIndex).Counters.IdleCount
    UserList(UserIndex).Counters.IdleCount = 0

    If Not UserList(UserIndex).flags.UserLogged Then

        Select Case Left$(rData, 6)

            Case "MARAKA"

                ''Function EsAdmin(ByVal Name As String) As Boolean
                ''Function EsDios(ByVal Name As String) As Boolean
                ''Function EsSemiDios(ByVal Name As String) As Boolean
                ''Function EsConsejero(ByVal Name As String) As Boolean
                ''Function EsRolesMaster(ByVal Name As String) As Boolean

                Dim SeguridadCliente As Long

                If ServerSoloGMs <> 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRServidor restringido a administradores. Consulte la página oficial o el foro oficial para mas información.")
                    Call CloseSocket(UserIndex)
                    Exit Sub

                End If

                Dim HDD As String

                tName = mid(ReadField(1, rData, 44), 7)
                SeguridadCliente = val(ReadField(5, rData, 44))
                HDD = ReadField(4, rData, 44)

                If SeguridadCliente = 0 Then
                    If EsDios(tName) Or EsSemiDios(tName) Or EsConsejero(tName) Or EsRolesMaster(tName) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRPara usar un personaje GM se necesita el Cliente de ADMINS.")
                        Call CloseSocket(UserIndex)
                        Exit Sub

                    End If

                ElseIf SeguridadCliente = 1 Then

                    If EsDios(tName) Or EsSemiDios(tName) Or EsConsejero(tName) Or EsRolesMaster(tName) Then

                        If GmTrue(tName, HDD) Then
                        Else
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRNo eres el dueño de este GM, no puedes loguear.")
                            Call CloseSocket(UserIndex)
                            Exit Sub

                        End If

                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERREntra con el cliente de usuarios no seas tan listo " & UCase(tName) & ".")
                        Call CloseSocket(UserIndex)
                        Exit Sub

                    End If

                End If

                rData = Right$(rData, Len(rData) - 6)
                Ver = ReadField(3, rData, 44)

                If VersionOK(Ver) Then

                    tName = ReadField(1, rData, 44)

                    If Not AsciiValidos(tName) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRNombre invalido.")
                        Call CloseSocket(UserIndex)
                        Exit Sub

                    End If

                    If Not PersonajeExiste(tName) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERREl personaje no existe.")
                        Call CloseSocket(UserIndex)
                        Exit Sub

                    End If

                    If Not BANCheck(tName) Then

                        Dim HDD_Serial As String

                        HDD_Serial = ReadField(4, rData, 44)

                        If modHDSerial.check_HD(HDD_Serial) = -1 Then

                            If EsGmChar(tName) Then

                                If Not EsHDD(tName, HDD_Serial) Then
                                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRSe te ha prohibido la entrada a AOMania")
                                    Exit Sub

                                End If

                            End If

                            Call ConnectUser(UserIndex, tName, ReadField(2, rData, 44), HDD_Serial)

                        Else
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRSe te ha prohibido la entrada a AOMania")

                        End If

                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRSe te ha prohibido la entrada a AOMania")

                    End If

                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRJuego desactualizado, cierra el juego y ejecuta AOMania.exe para Actualizarlo.")

                End If

                Exit Sub

            Case "ZORRON"

                If PuedeCrearPersonajes = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRLa creacion de personajes en este servidor se ha deshabilitado.")
                    Call CloseSocket(UserIndex)
                    Exit Sub

                End If

                If ServerSoloGMs <> 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRServidor restringido a administradores. Consulte la página oficial o el foro oficial para mas información.")
                    Call CloseSocket(UserIndex)
                    Exit Sub

                End If

                If aClon.MaxPersonajes(UserList(UserIndex).ip) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRHas creado demasiados personajes.")
                    Call CloseSocket(UserIndex)
                    Exit Sub

                End If

                rData = Right$(rData, Len(rData) - 6)
                Ver = ReadField(3, rData, 44)

                If VersionOK(Ver) Then

                    HDD_Serial = ReadField(10, rData, 44)

                    If modHDSerial.check_HD(HDD_Serial) = -1 Then

                        Call ConnectNewUser(UserIndex, ReadField(1, rData, 44), ReadField(2, rData, 44), ReadField(4, rData, 44), ReadField(5, rData, 44), ReadField(6, rData, 44), ReadField(7, rData, 44), ReadField(8, rData, 44), ReadField(9, rData, 44), HDD_Serial)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERR1")

                    End If

                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRJuego desactualizado, cierra el juego y ejecuta AOMania.exe para Actualizarlo.")

                End If

                Exit Sub

            Case "TIRDAD"

                rData = Right$(rData, Len(rData) - 6)

                UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = ReadField(1, rData, 44)
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = ReadField(2, rData, 44)
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = ReadField(3, rData, 44)
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = ReadField(4, rData, 44)
                UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) = ReadField(5, rData, 44)

                Exit Sub

        End Select

        Select Case Left$(rData, 4)

            Case "BORR"    ' <<< borra personajes

                On Error GoTo ExitErr1

                rData = Right$(rData, Len(rData) - 4)
                Arg1 = ReadField(1, rData, 44)

                If Not AsciiValidos(Arg1) Then Exit Sub

                '¿Existe el personaje?
                If Not FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) Then
                    Call CloseSocket(UserIndex)
                    Exit Sub

                End If

                '¿Es el passwd valido?
                If UCase$(ReadField(2, rData, 44)) <> UCase$(GetVar(CharPath & UCase$(Arg1) & ".chr", "INIT", "Password")) Then
                    Call CloseSocket(UserIndex)
                    Exit Sub

                End If

                'If FileExist(CharPath & ucase$(Arg1) & ".chr", vbNormal) Then
                Dim rt As String

                rt = App.Path & "\ChrBackUp\" & UCase$(Arg1) & ".bak"

                If FileExist(rt, vbNormal) Then Kill rt
                Name CharPath & UCase$(Arg1) & ".chr" As rt
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "BORROK")
                Exit Sub
ExitErr1:
                Call LogError(Err.Description & " " & rData)
                Exit Sub

                'End If
        End Select

        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        'Si no esta logeado y envia un comando diferente a los
        'de arriba cerramos la conexion.
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        Call LogHackAttemp("Mesaje enviado sin logearse:" & rData)
        Call CloseSocket(UserIndex)
        Exit Sub

    End If    ' if not user logged

    Dim Procesado As Boolean

    ' bien ahora solo procesamos los comandos que NO empiezan
    ' con "/".
    If Left$(rData, 1) <> "/" Then

        Call HandleData_1(UserIndex, rData, Procesado)

        If Procesado Then Exit Sub

        ' bien hasta aca fueron los comandos que NO empezaban con
        ' "/". Ahora adiviná que sigue :)
    Else

        Call HandleData_2(UserIndex, rData, Procesado)

        If Procesado Then Exit Sub

    End If    ' "/"

    If UserList(UserIndex).flags.Privilegios = PlayerType.User Then
        UserList(UserIndex).Counters.IdleCount = IdleCountBackup

    End If

    '>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<
    If UserList(UserIndex).flags.Privilegios = PlayerType.User Then Exit Sub
    '>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<

    If UCase$(rData) = "/PANELGM" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ABPANEL")
        Exit Sub

    End If

    If UCase$(Left$(rData, 13)) = "/ADVERTENCIA " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        '/carcel nick@motivo
        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        rData = Right$(rData, Len(rData) - 13)

        Name = ReadField(1, rData, Asc("@"))
        tStr = ReadField(2, rData, Asc("@"))

        If Name = "" Or tStr = "" Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Utilice /advertencia nick@motivo" & FONTTYPE_INFO)
            Exit Sub

        End If

        TIndex = NameIndex(Name)

        If TIndex <= 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(TIndex).flags.Privilegios > PlayerType.User Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes advertir a administradores." & FONTTYPE_INFO)
            Exit Sub

        End If

        Name = Replace(Name, "\", "")
        Name = Replace(Name, "/", "")

        If FileExist(CharPath & Name & ".chr", vbNormal) Then
            tInt = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & Name & ".chr", "PENAS", "Cant", tInt + 1)
            Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & ": ADVERTENCIA por: " & LCase$(tStr) & " " & Date & " " & Time)

        End If

        Call LogGM(UserList(UserIndex).Name, " advirtio a " & Name)
        Exit Sub

    End If

    If UCase$(Left$(rData, 9)) = "/PANELSOS" Then
        Call CargarArchivosSos(UserIndex)
        Exit Sub

    End If

    If UCase$(Left$(rData, 11)) = "/DROPQUEST " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 11)

        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            If Quest.Existe(rData) Then Call Quest.Quitar(rData)
            Exit Sub

        End If

        If UserList(TIndex).flags.Quest = 1 Then
            If Quest.Existe(UserList(TIndex).Name) Then Call Quest.Quitar(UserList(TIndex).Name)
            UserList(TIndex).flags.Quest = 0
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has borrado el QUEST de: " & rData & FONTTYPE_INFO)
            Exit Sub
        Else

            If Quest.Existe(UserList(TIndex).Name) Then Call Quest.Quitar(UserList(TIndex).Name)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has borrado el QUEST de: " & rData & FONTTYPE_INFO)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 9)) = "/DROPSOS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 9)

        Call DropSOS(rData, UserIndex)

        Exit Sub

    End If

    If UCase$(Left$(rData, 8)) = "/DROPGM " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 8)

        Call BorrarGM(rData, UserIndex)

        Exit Sub

    End If

    If UCase$(Left$(rData, 14)) = "/PANELCONSULTA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Call CargarArchivosGM(UserIndex)
        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "SOSDONE" Then
        rData = Right$(rData, Len(rData) - 7)
        Call Ayuda.Quitar(rData)
        Exit Sub

    End If

    If UCase$(rData) = "LISTUSU" Then
        ' If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        tStr = "LISTUSU"

        For LoopC = 1 To LastUser

            If (UserList(LoopC).Name <> "") Then
                tStr = tStr & UserList(LoopC).Name & ","

            End If

        Next LoopC

        If Len(tStr) > 7 Then
            tStr = Left$(tStr, Len(tStr) - 1)

        End If

        Call SendData(SendTarget.ToIndex, UserIndex, 0, tStr)
        Exit Sub

    End If

    If UCase$(rData) = "LISTQST" Then

        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        Dim mm As String

        For n = 1 To Quest.Longitud
            mm = Quest.VerElemento(n)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "LISTQST" & mm)
        Next n

        Exit Sub

    End If

    If UCase(Left(rData, 13)) = "/SEARCHNPCSH " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 13)

        Dim nCH As Long

        CountNpcH = 0

        For nCH = 500 To 724

            Call LeerNpcH(nCH, rData, UserIndex)

        Next nCH

    End If

    If UCase(Left(rData, 12)) = "/SEARCHNPCS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 12)

        Dim nC As Long

        CountNpc = 0

        For nC = 1 To 301

            Call LeerNpc(nC, rData, UserIndex)

        Next nC

    End If

    If UCase(Left(rData, 13)) = "/SEARCHITEMS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 13)

        Dim xci       As Long

        Dim CID       As Long

        Dim CNameItem As String

        Dim Asi       As Long

        Asi = 0

        For xci = 1 To NumObjDatas
            CID = xci
            CNameItem = ObjData(xci).Name

            If rData = "" Then
                Asi = Asi + 1
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "VCTS" & CID & "#" & CNameItem)
            Else

                If InStr(LCase(CNameItem), LCase(rData)) Then

                    Asi = Asi + 1
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "VCTS" & CID & "#" & CNameItem)

                End If

            End If

        Next xci

        If Asi = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "VITS" & Asi)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "VITS" & Asi)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 11)) = "/BORRAR SOS" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        Call Ayuda.Reset
        Exit Sub

    End If

    If UCase$(Left$(rData, 9)) = "/RESPUES " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 9)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario offline!!." & FONTTYPE_INFO)
            Exit Sub

        End If

        Call MostrarSop(UserIndex, TIndex, rData)
        SendData SendTarget.ToIndex, UserIndex, 0, "INITSOP"
        Exit Sub

    End If

    If UCase$(Left$(rData, 3)) = "SPA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 3)

        If val(rData) > 0 And val(rData) < UBound(SpawnList) + 1 Then Call SpawnNpc(SpawnList(val(rData)).NpcIndex, UserList(UserIndex).pos, True, False)
        Call LogGM(UserList(UserIndex).Name, "Sumoneo " & SpawnList(val(rData)).NpcName)

        Exit Sub

    End If

    If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
        Call HandleGM(UserIndex, rData)

    End If

ErrorHandler:
    Call LogError("HandleData. CadOri:" & CadenaOriginal & " Nom:" & UserList(UserIndex).Name & "UI:" & UserIndex & " N: " & Err.Number & " D: " & Err.Description)
    'Resume
    'Call CloseSocket(UserIndex)
    'Call Cerrar_Usuario(UserIndex)

End Sub

Sub ReloadSokcet()

    On Error GoTo errhandler

    Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)

    If NumUsers <= 0 Then
        Call WSApiReiniciarSockets
    Else

        '       Call apiclosesocket(SockListen)
        '       SockListen = ListenForConnect(Puerto, hWndMsg, "")
    End If

    Exit Sub
errhandler:
    Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.Description)

End Sub

Public Sub EnviarNoche(ByVal UserIndex As Integer)

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "NOC" & IIf(DeNoche And (MapInfo(UserList(UserIndex).pos.Map).Zona = Campo Or MapInfo(UserList( _
                                                                                                                                          UserIndex).pos.Map).Zona = Ciudad), "1", "0"))
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "NOC" & IIf(DeNoche, "1", "0"))

End Sub

Public Sub EcharPjsNoPrivilegiados()

    Dim LoopC As Long

    For LoopC = 1 To LastUser

        If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then

            If UserList(LoopC).flags.Privilegios < PlayerType.Consejero Then
                Call CloseSocket(LoopC)

            End If

        End If

    Next LoopC

End Sub

'CRAW; 18/09/2019 --> Segmento el handledata por estar muy lleno.
Public Sub getServerDelay(UserIndex As Integer)
    Dim tnow As Long
    tnow = GetTickCount() And &H7FFFFFFF
    Call SendData(ToIndex, UserIndex, 0, "NA" & tnow)
End Sub
