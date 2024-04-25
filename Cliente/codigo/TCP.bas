Attribute VB_Name = "TCP"

'Pablo Ignacio Márquez

Option Explicit

'RUTAS DE ENVIO DE DATOS
Public Enum SendTarget

    toIndex = 0         ' Envia a un solo User
    toall = 1           ' A todos los Users
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
    Dim i   As Integer

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
    Dim i   As Integer

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
                   US1 As String, _
                   US2 As String, _
                   US3 As String, _
                   US4 As String, _
                   US5 As String, _
                   US6 As String, _
                   US7 As String, _
                   US8 As String, _
                   US9 As String, _
                   US10 As String, _
                   US11 As String, _
                   US12 As String, _
                   US13 As String, _
                   US14 As String, _
                   US15 As String, _
                   US16 As String, _
                   US17 As String, _
                   US18 As String, _
                   US19 As String, US20 As String, US21 As String, US22 As String, US23, US24, US25, US26, UserEmail As String, Hogar As String, ByVal HdSerial As String)

    If Not AsciiValidos(Name) Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "ERRNombre invalido.")
        Exit Sub

    End If

    Dim LoopC      As Integer
    Dim totalskpts As Long

    '¿Existe el personaje?
    If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = True Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "ERRYa existe el personaje.")
        Exit Sub

    End If

    'Tiró los dados antes de llegar acá??
    If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "ERRDebe tirar los dados antes de poder crear un personaje.")
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
    UserList(UserIndex).Hogar = Hogar
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
    
    UserList(UserIndex).Stats.UserSkills(1) = val(US1)
    UserList(UserIndex).Stats.UserSkills(2) = val(US2)
    UserList(UserIndex).Stats.UserSkills(3) = val(US3)
    UserList(UserIndex).Stats.UserSkills(4) = val(US4)
    UserList(UserIndex).Stats.UserSkills(5) = val(US5)
    UserList(UserIndex).Stats.UserSkills(6) = val(US6)
    UserList(UserIndex).Stats.UserSkills(7) = val(US7)
    UserList(UserIndex).Stats.UserSkills(8) = val(US8)
    UserList(UserIndex).Stats.UserSkills(9) = val(US9)
    UserList(UserIndex).Stats.UserSkills(10) = val(US10)
    UserList(UserIndex).Stats.UserSkills(11) = val(US11)
    UserList(UserIndex).Stats.UserSkills(12) = val(US12)
    UserList(UserIndex).Stats.UserSkills(13) = val(US13)
    UserList(UserIndex).Stats.UserSkills(14) = val(US14)
    UserList(UserIndex).Stats.UserSkills(15) = val(US15)
    UserList(UserIndex).Stats.UserSkills(16) = val(US16)
    UserList(UserIndex).Stats.UserSkills(17) = val(US17)
    UserList(UserIndex).Stats.UserSkills(18) = val(US18)
    UserList(UserIndex).Stats.UserSkills(19) = val(US19)
    UserList(UserIndex).Stats.UserSkills(20) = val(US20)
    UserList(UserIndex).Stats.UserSkills(21) = val(US21)
    UserList(UserIndex).Stats.UserSkills(22) = val(US22)
    UserList(UserIndex).Stats.UserSkills(23) = val(US23)
    UserList(UserIndex).Stats.UserSkills(24) = val(US24)
    UserList(UserIndex).Stats.UserSkills(25) = val(US25)
    UserList(UserIndex).Stats.UserSkills(26) = val(US26)

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
    UserList(UserIndex).char.Heading = eHeading.SOUTH

    Call DarCuerpoYCabeza(UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).Raza, UserList(UserIndex).Genero)
            
    UserList(UserIndex).OrigChar = UserList(UserIndex).char
 
    UserList(UserIndex).char.WeaponAnim = NingunArma
    UserList(UserIndex).char.ShieldAnim = NingunEscudo
    UserList(UserIndex).char.CascoAnim = NingunCasco
    '[MaTeO 9]
    UserList(UserIndex).char.Alas = NingunAlas
    '[/MaTeO 9]

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
    UserList(UserIndex).Stats.PuntosDeath = 0
    UserList(UserIndex).Stats.PuntosDuelos = 0
    UserList(UserIndex).Stats.PuntosTorneo = 0
    UserList(UserIndex).Stats.PuntosRetos = 0
    UserList(UserIndex).Stats.PuntosPlante = 0

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
    UserList(UserIndex).Stats.ELU = 300
    UserList(UserIndex).Stats.ELV = 1

    Call SetearInv(UserIndex, UCase$(UserList(UserIndex).Clase), UCase$(UserList(UserIndex).Raza))

    'Open User
    
    Call EnviaRegistro(Name, Password, UserEmail, UserClase, UserRaza)
    
    Call SaveUser(UserIndex, CharPath & UCase$(Name) & ".chr")
    
    Call ConnectUser(UserIndex, Name, UserList(UserIndex).Password, HdSerial)
  
End Sub

Sub CloseSocket(ByVal UserIndex As Integer)
    Dim LoopC As Integer

    On Error GoTo errhandler

    If UserIndex = LastUser Then

        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1

            If LastUser < 1 Then Exit Do
        Loop

    End If
    
    Call RestCriCi(UserIndex)
    
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

    If UserList(UserIndex).flags.death = True Then
        Call death_muere(UserIndex)

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
                UserIndex).char.Head, UserList(UserIndex).char.Heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, _
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

    If UserList(UserIndex).flags.EstaDueleando1 = True Then
        Call DesconectarDueloPlantes(UserList(UserIndex).flags.Oponente1, UserIndex)

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
                Call SendData(SendTarget.toIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
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

    Call SendData(toall, 0, 0, "³" & NumUsers)
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
    
    Dim ret As Long
        
    ret = WsApiEnviar(UserIndex, Datos)
    
    If ret <> 0 And ret <> WSAEWOULDBLOCK Then
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)

    End If

    EnviarDatosASlot = ret
    Exit Function
    
Err:
  
    Call LogError("TCP::EnviarDatosASlot. UI/ConnId/Datos: " & UserIndex & "/" & UserList(UserIndex).ConnID & "/" & Datos)

End Function

Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal sndData As String)

    On Error Resume Next

    Dim LoopC As Integer
    Dim X     As Integer
    Dim Y     As Integer

    sndData = sndData & ENDC

    Select Case sndRoute

        Case SendTarget.ToPCArea

            For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
                For X = UserList(sndIndex).pos.X - MinXBorder + 1 To UserList(sndIndex).pos.X + MinXBorder - 1

                    If InMapBounds(sndMap, X, Y) Then
                        If MapData(sndMap, X, Y).UserIndex > 0 Then
                            If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                                Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)

                            End If

                        End If

                    End If

                Next X
            Next Y

            Exit Sub
    
        Case SendTarget.toIndex

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
        
        Case SendTarget.toall

            For LoopC = 1 To LastUser

                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call EnviarDatosASlot(LoopC, sndData)

                    End If

                End If

            Next LoopC

            Exit Sub
    
        Case SendTarget.ToAllButIndex

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
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
                For X = UserList(sndIndex).pos.X - MinXBorder + 1 To UserList(sndIndex).pos.X + MinXBorder - 1

                    If InMapBounds(sndMap, X, Y) Then
                        If MapData(sndMap, X, Y).UserIndex > 0 Then
                            If UserList(MapData(sndMap, X, Y).UserIndex).flags.Muerto = 1 Or UserList(MapData(sndMap, X, _
                                    Y).UserIndex).flags.Privilegios >= 1 Then

                                If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                                    Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)

                                End If

                            End If

                        End If

                    End If

                Next X
            Next Y

            Exit Sub

            '[Alejo-18-5]
        Case SendTarget.ToPCAreaButIndex

            For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
                For X = UserList(sndIndex).pos.X - MinXBorder + 1 To UserList(sndIndex).pos.X + MinXBorder - 1

                    If InMapBounds(sndMap, X, Y) Then
                        If (MapData(sndMap, X, Y).UserIndex > 0) And (MapData(sndMap, X, Y).UserIndex <> sndIndex) Then

                            If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                                Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)

                            End If

                        End If

                    End If

                Next X
            Next Y

            Exit Sub
       
        Case SendTarget.ToClanArea

            For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
                For X = UserList(sndIndex).pos.X - MinXBorder + 1 To UserList(sndIndex).pos.X + MinXBorder - 1

                    If InMapBounds(sndMap, X, Y) Then
                        If (MapData(sndMap, X, Y).UserIndex > 0) Then
                            If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                                If UserList(sndIndex).GuildIndex > 0 And UserList(MapData(sndMap, X, Y).UserIndex).GuildIndex = UserList( _
                                        sndIndex).GuildIndex Then
                                    Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)

                                End If

                            End If

                        End If

                    End If

                Next X
            Next Y

            Exit Sub

        Case SendTarget.ToPartyArea

            For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
                For X = UserList(sndIndex).pos.X - MinXBorder + 1 To UserList(sndIndex).pos.X + MinXBorder - 1

                    If InMapBounds(sndMap, X, Y) Then
                        If (MapData(sndMap, X, Y).UserIndex > 0) Then
                            If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                                If UserList(sndIndex).PartyIndex > 0 And UserList(MapData(sndMap, X, Y).UserIndex).PartyIndex = UserList( _
                                        sndIndex).PartyIndex Then
                                    Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)

                                End If

                            End If

                        End If

                    End If

                Next X
            Next Y

            Exit Sub
        
            '[CDT 17-02-2004]
        Case SendTarget.ToAdminsAreaButConsejeros

            For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
                For X = UserList(sndIndex).pos.X - MinXBorder + 1 To UserList(sndIndex).pos.X + MinXBorder - 1

                    If InMapBounds(sndMap, X, Y) Then
                        If (MapData(sndMap, X, Y).UserIndex > 0) And (MapData(sndMap, X, Y).UserIndex <> sndIndex) Then

                            If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                                If UserList(MapData(sndMap, X, Y).UserIndex).flags.Privilegios > 1 Then
                                    Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)

                                End If

                            End If

                        End If

                    End If

                Next X
            Next Y

            Exit Sub
            '[/CDT]

        Case SendTarget.ToNPCArea

            For Y = Npclist(sndIndex).pos.Y - MinYBorder + 1 To Npclist(sndIndex).pos.Y + MinYBorder - 1
                For X = Npclist(sndIndex).pos.X - MinXBorder + 1 To Npclist(sndIndex).pos.X + MinXBorder - 1

                    If InMapBounds(sndMap, X, Y) Then
                        If MapData(sndMap, X, Y).UserIndex > 0 Then
                            If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                                Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)

                            End If

                        End If

                    End If

                Next X
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

    Dim X As Integer, Y As Integer

    For Y = UserList(Index).pos.Y - MinYBorder + 1 To UserList(Index).pos.Y + MinYBorder - 1
        For X = UserList(Index).pos.X - MinXBorder + 1 To UserList(Index).pos.X + MinXBorder - 1

            If MapData(UserList(Index).pos.Map, X, Y).UserIndex = Index2 Then
                EstaPCarea = True
                Exit Function

            End If
        
        Next X
    Next Y

    EstaPCarea = False

End Function

Function HayPCarea(pos As WorldPos) As Boolean

    Dim X As Integer, Y As Integer

    For Y = pos.Y - MinYBorder + 1 To pos.Y + MinYBorder - 1
        For X = pos.X - MinXBorder + 1 To pos.X + MinXBorder - 1

            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(pos.Map, X, Y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function

                End If

            End If

        Next X
    Next Y

    HayPCarea = False

End Function

Function HayOBJarea(pos As WorldPos, ObjIndex As Integer) As Boolean

    Dim X As Integer, Y As Integer

    For Y = pos.Y - MinYBorder + 1 To pos.Y + MinYBorder - 1
        For X = pos.X - MinXBorder + 1 To pos.X + MinXBorder - 1

            If MapData(pos.Map, X, Y).OBJInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function

            End If
        
        Next X
    Next Y

    HayOBJarea = False

End Function

Function ValidateChr(ByVal UserIndex As Integer) As Boolean

    ValidateChr = UserList(UserIndex).char.Head <> 0 And UserList(UserIndex).char.Body <> 0 And ValidateSkills(UserIndex)

End Function

Sub ConnectUser(ByVal UserIndex As Integer, Name As String, Password As String, ByVal hdString As String)
    Dim n    As Integer
    Dim tStr As String

    'Reseteamos los FLAGS
    With UserList(UserIndex)
            
        .flags.Escondido = 0
        .flags.TargetNpc = 0
        .flags.TargetNpcTipo = eNPCType.Comun
        .flags.TargetObj = 0
        .flags.TargetUser = 0
        .Counters.Invisibilidad = IntervaloInvisible
        .char.FX = 0
        
        Call SendData(SendTarget.toIndex, UserIndex, 0, "INVI0")
        
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||       .oo .oPYo. o     o               o              .oPYo.    .oPYo.       .oo" & _
                FONTTYPE_Motd1)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||       .P 8 8    8 8b   d8                              8  .o8    8  .o8      .P 8" & _
                FONTTYPE_Motd1)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||      .P  8 8    8 8`b d'8 .oPYo. odYo. o8 .oPYo.       8 .P'8    8 .P'8     .P  8" & _
                FONTTYPE_Motd1)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||     oPooo8 8    8 8 `o' 8 .oooo8 8' `8  8 .oooo8       8.d' 8    8.d' 8         8" & _
                FONTTYPE_Motd1)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||    .P    8 8    8 8     8 8    8 8   8  8 8    8       8o'  8    8o'  8         8" & _
                FONTTYPE_Motd1)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||   .P     8 `YooP' 8     8 `YooP8 8   8  8 `YooP8       `YooP' 88 `YooP'88       8" & _
                FONTTYPE_Motd1)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||                                                                  www.AoMania.Net" & _
                FONTTYPE_Motd1)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||      X = Seguro Objetos   Q = Mapa   P = Mapa  S = Seguro   W = Seguro de clan" & _
                FONTTYPE_Motd2)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||                         F1 = /Meditar   F12 = Macro Interno Para trabajadores." & _
                FONTTYPE_Motd2)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||                     Si tienes alguna duda o necesitas ayuda, escribe /GM TEXTO" & _
                FONTTYPE_Motd2)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||                                                        Version 0.0.1 Año: 2019" & _
                FONTTYPE_Motd2)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||:- Argentumania -:" & FONTTYPE_INFO)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||> ¡¡¡Bienvenidos al Servidor Oficial AoManiA 2019!!!" & FONTTYPE_GUILD)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||> Versión Actual v1 de AOMania, Argentumania 2018. Mod Argentum Online" & FONTTYPE_Motd3)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||> Para cualquier duda, /gm consulta" & FONTTYPE_Motd3)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||> Web Oficial aomania.net argentumania.es" & FONTTYPE_Motd4)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||> Foro Oficial foro.argentumania.es" & FONTTYPE_Motd4)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||---------------" & FONTTYPE_SERVER)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Para ver el Mapa de AOMania dejar pulsada la tecla Q ó P." & FONTTYPE_Motd5)
                
        Call SendInfoCastillos(UserIndex)
        
        If MaxLevel > 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El máximo nivel es " & MaxLevel & ", adquirido por " & UserMaxLevel & "." & _
                    FONTTYPE_SERVER)

        End If
        
        Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & FONTTYPE_INFO)
        
        If MultMsg = "" Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La Experiencia del Servidor esta subido por " & Multexp & "." & FONTTYPE_Motd5)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El Oro del Servidor esta subido por " & MultOro & "." & FONTTYPE_Motd5)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La Experiencia del Servidor esta subido por " & Multexp & ". " & MultMsg & "." & _
                    FONTTYPE_Motd5)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El Oro del Servidor esta subido por " & MultOro & ". " & MultMsg & "." & _
                    FONTTYPE_Motd5)

        End If
        
        If StatusNosfe = True Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Nosferatu esta haciendo estragos en el mapa " & MapaNosfe & FONTTYPE_GUILD)
        End If
        
        If ExpCriatura = True Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Hoy es día de " & NombreCriatura & ", su experencia esta aumentada x" & _
                    LoteriaCriatura & "." & FONTTYPE_TALK)

        End If
        
        If OroCriatura = True Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Hoy es día de " & NombreCriatura & ", su oro esta aumentada x" & LoteriaCriatura & _
                    "." & FONTTYPE_TALK)

        End If
        
        If DiaEspecialExp = True Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Estáis de suerte! Día especial, la experencia esta aumentada por x2" & FONTTYPE_TALK)

        End If
        
        If DiaEspecialOro = True Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Estáis de suerte! Día especial, el oro esta aumentada por x2" & FONTTYPE_TALK)
            
        End If
        
        'Controlamos no pasar el maximo de usuarios
        If NumUsers >= MaxUsers Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, _
                    "ERREl servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        '¿Este IP ya esta conectado?
        If AllowMultiLogins = 0 Then
            If CheckForSameIP(UserIndex, .ip) = True Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "ERRNo es posible usar mas de un personaje al mismo tiempo.")
                Call CloseSocket(UserIndex)
                Exit Sub

            End If

        End If

        '¿Existe el personaje?
        If Not FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "ERREl personaje no existe.")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        '¿Es el passwd valido?
        If UCase$(Password) <> UCase$(GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password")) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "ERRPassword incorrecto.")
    
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        '¿Ya esta conectado el personaje?
        If CheckForSameName(UserIndex, Name) Then
            If UserList(NameIndex(Name)).Counters.Saliendo Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "ERREl usuario está saliendo.")
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "ERRPerdon, un usuario con el mismo nombre se há logoeado.")

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
            Call SendData(SendTarget.toIndex, UserIndex, 0, "ERRError en el personaje.")
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
            Call SendData(SendTarget.toIndex, UserIndex, 0, "PARADOW")

        End If

        'Feo, esto tiene que ser parche cliente
        If .flags.Estupidez = 0 Then Call SendData(SendTarget.toIndex, UserIndex, 0, "NESTUP")

        'Posicion de comienzo
        If .pos.Map = 0 Then
            If .Stats.ELV < 13 Then
                .pos.Map = "37"
                .pos.X = "35"
                .pos.Y = "69"
            Else
                .pos.Map = "34"
                .pos.X = "30"
                .pos.Y = "50"

            End If

        Else

            'Anti Pisadas
            If MapData(.pos.Map, .pos.X, .pos.Y).UserIndex <> 0 Then
                Dim nPos As WorldPos
                Call ClosestStablePos(.pos, nPos)
                
                If nPos.X <> 0 And nPos.Y <> 0 Then
                    .pos.Map = nPos.Map
                    .pos.X = nPos.X
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
                        .pos.X = Zonas(.Zona).X
                        .flags.Embarcado = 0

                    End If

                End If

            End If
        
        End If

        If Not MapaValido(.pos.Map) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "ERREL PJ se encuenta en un mapa invalido.")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        'Nombre de sistema
        .Name = Name
        .Password = Password
        .hd_String = hdString
        
        Call WriteVar(CharPath & .Name & ".chr", "INIT", "LastHD", .hd_String)
 
        .showName = True 'Por default los nombres son visibles

        'Info
        Call SendData(SendTarget.toIndex, UserIndex, 0, "IU" & UserIndex) 'Enviamos el User index
        Call SendData(SendTarget.toIndex, UserIndex, 0, "CM" & .pos.Map & "," & MapInfo(.pos.Map).MapVersion) 'Carga el mapa
        Call SendData(SendTarget.toIndex, UserIndex, 0, "TM" & MapInfo(.pos.Map).Music)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "N~" & MapInfo(.pos.Map).Name)

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
        Call MakeUserChar(SendTarget.ToMap, 0, .pos.Map, UserIndex, .pos.Map, .pos.X, .pos.Y)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "IP" & .char.CharIndex)
        ''[/el oso]

        Call SendUserStatsBox(UserIndex)
        Call SendUserHitBox(UserIndex)
        Call EnviarHambreYsed(UserIndex)
        Call EnviarAmarillas(UserIndex)
        Call EnviarVerdes(UserIndex)

        If haciendoBK Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "BKW")
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||AOMania> Por favor espera algunos segundos, WorldSave esta ejecutandose." & _
                    FONTTYPE_SERVER)

        End If

        If EnPausa Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "BKW")
            Call SendData(SendTarget.toIndex, UserIndex, 0, _
                    "||AOMania> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde." & _
                    FONTTYPE_SERVER)

        End If

        If EnTesting And .Stats.ELV >= 18 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, _
                    "ERRServidor en Testing por unos minutos, conectese con PJs de nivel menor a 18. No se conecte con Pjs que puedan resultar importantes por ahora pues pueden arruinarse.")
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        'Actualiza el Num de usuarios
        'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
        NumUsers = NumUsers + 1
        .flags.UserLogged = True
        Call SendData(toall, 0, 0, "³" & NumUsers)
        
        'usado para borrar Pjs
        Call WriteVar(CharPath & .Name & ".chr", "INIT", "Logged", "1")

        Call SendData(toall, 0, 0, "³" & NumUsers)
        MapInfo(.pos.Map).NumUsers = MapInfo(.pos.Map).NumUsers + 1

        If .Stats.SkillPts > 0 Then
            Call EnviarSkills(UserIndex)
            Call EnviarSubirNivel(UserIndex, .Stats.SkillPts)

        End If

        If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers

        If NumUsers > recordusuarios Then
            Call SendData(SendTarget.toall, 0, 0, "||Record de usuarios conectados simultaneamente." & "Hay " & NumUsers & " usuarios." & _
                    FONTTYPE_TURQ)
            recordusuarios = NumUsers
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

        If .flags.Navegando = 1 Then Call SendData(SendTarget.toIndex, UserIndex, 0, "NAVEG")

        If Criminal(UserIndex) Then
            'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Miembro de las fuerzas del caos > Seguro desactivado <" & FONTTYPE_FIGHT)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "OFFOFS")
            .flags.Seguro = False
        Else
            .flags.Seguro = True
            Call SendData(SendTarget.toIndex, UserIndex, 0, "ONONS")

        End If
        
        .flags.SeguroClan = False
        .flags.SeguroCombate = False
        .flags.SeguroHechizos = True
        .flags.SeguroObjetos = False
        
        Call SendData(SendTarget.toIndex, UserIndex, 0, "SEGCO99")
        Call SendData(SendTarget.toIndex, UserIndex, 0, "SEG11")
        Call SendData(SendTarget.toIndex, UserIndex, 0, "SEG13")
        Call SendData(SendTarget.toIndex, UserIndex, 0, "SEG15")

        If ServerSoloGMs > 0 Then
            If .flags.Privilegios < ServerSoloGMs Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "ERRServidor restringido a administradores de jerarquia mayor o igual a: " & _
                        ServerSoloGMs & ". Por favor intente en unos momentos.")
                Call CloseSocket(UserIndex)
                Exit Sub

            End If

        End If

        If .GuildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, 0, "||" & .Name & " se ha conectado." & FONTTYPE_GUILD)

            If Not modGuilds.m_ConectarMiembroAClan(UserIndex, .GuildIndex) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu estado no te permite entrar al clan." & FONTTYPE_GUILD)

            End If

        End If

        Call SendData(SendTarget.ToPCArea, UserIndex, .pos.Map, "CFX" & .char.CharIndex & "," & FXIDs.FXWARP & "," & 0)

        Call SendData(SendTarget.toIndex, UserIndex, 0, "LODXXD")

        Call modGuilds.SendGuildNews(UserIndex)
         
        Call SendMainAmbient(UserIndex)
        Call SendSecondaryAmbient(UserIndex)

        tStr = modGuilds.a_ObtenerRechazoDeChar(.Name)

        If tStr <> vbNullString Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "!!Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr _
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
        
    #If MYSQL = 1 Then
       Call Add_DataBase(UserIndex, "Online")
    #End If
    
    DoEvents
    
    If .PalabraSecreta = "" Or .flags.RPasswd = "" Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "SGDP")
    End If
     
    End With

End Sub

Sub SendMOTD(ByVal UserIndex As Integer)
    Dim j As Long
    
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||       .oo .oPYo. o     o               o              .oPYo.    .oPYo.       .oo" & _
            FONTTYPE_Motd1)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||       .P 8 8    8 8b   d8                              8  .o8    8  .o8      .P 8" & _
            FONTTYPE_Motd1)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||      .P  8 8    8 8`b d'8 .oPYo. odYo. o8 .oPYo.       8 .P'8    8 .P'8     .P  8" & _
            FONTTYPE_Motd1)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||     oPooo8 8    8 8 `o' 8 .oooo8 8' `8  8 .oooo8       8.d' 8    8.d' 8         8" & _
            FONTTYPE_Motd1)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||    .P    8 8    8 8     8 8    8 8   8  8 8    8       8o'  8    8o'  8         8" & _
            FONTTYPE_Motd1)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||   .P     8 `YooP' 8     8 `YooP8 8   8  8 `YooP8       `YooP' 88 `YooP'88       8" & _
            FONTTYPE_Motd1)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||                                                                  www.AoMania.Net" & _
            FONTTYPE_Motd1)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||      X = Seguro Objetos   Q = Mapa   P = Mapa  S = Seguro   W = Seguro de clan" & _
            FONTTYPE_Motd2)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||                         F1 = /Meditar   F12 = Macro Interno Para trabajadores." & _
            FONTTYPE_Motd2)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||                     Si tienes alguna duda o necesitas ayuda, escribe /GM TEXTO" & _
            FONTTYPE_Motd2)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||                                                        Version 0.0.1 Año: 2019" & _
            FONTTYPE_Motd2)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||:- Argentumania -:" & FONTTYPE_INFO)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||> ¡¡¡Bienvenidos al Servidor Oficial AoManiA 2019!!!" & FONTTYPE_GUILD)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||> Versión Actual v1 de AOMania, Argentumania 2018. Mod Argentum Online" & FONTTYPE_Motd3)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||> Para cualquier duda, /gm consulta" & FONTTYPE_Motd3)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||> Web Oficial aomania.net argentumania.es" & FONTTYPE_Motd4)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||> Foro Oficial foro.argentumania.es" & FONTTYPE_Motd4)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||---------------" & FONTTYPE_SERVER)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Para ver el Mapa de AOMania dejar pulsada la tecla Q ó P." & FONTTYPE_Motd5)
            
    Call SendInfoCastillos(UserIndex)

    Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & FONTTYPE_INFO)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El máximo nivel es " & MaxLevel & ", adquirido por " & UserMaxLevel & "." & FONTTYPE_SERVER)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & FONTTYPE_INFO)
        
    If MultMsg = "" Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||La Experiencia del Servidor esta subido por " & Multexp & "." & FONTTYPE_Motd5)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||El Oro del Servidor esta subido por " & MultOro & "." & FONTTYPE_Motd5)
    Else
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||La Experiencia del Servidor esta subido por " & Multexp & ". " & MultMsg & "." & _
                FONTTYPE_Motd5)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||El Oro del Servidor esta subido por " & MultOro & ". " & MultMsg & "." & FONTTYPE_Motd5)

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
        .Heading = 0
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
        .pos.X = 0
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

    Dim n     As Integer
    Dim X     As Integer
    Dim Y     As Integer
    Dim LoopC As Integer
    Dim Map   As Integer
    Dim Name  As String
    Dim Raza  As String
    Dim Clase As String
    Dim i     As Integer

    Dim aN    As Integer

    aN = UserList(UserIndex).flags.AtacadoPorNpc

    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = ""

    End If

    UserList(UserIndex).flags.AtacadoPorNpc = 0

    Map = UserList(UserIndex).pos.Map
    X = UserList(UserIndex).pos.X
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
    Dim X                  As Integer
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
                    Call SendData(SendTarget.toIndex, UserIndex, 0, _
                            "ERRServidor restringido a administradores. Consulte la página oficial o el foro oficial para mas información.")
                    Call CloseSocket(UserIndex)
                    Exit Sub

                End If

                Dim HDD As String
                tName = mid(ReadField(1, rData, 44), 7)
                SeguridadCliente = val(ReadField(5, rData, 44))
                HDD = ReadField(4, rData, 44)
                
                If SeguridadCliente = 0 Then
                    If EsDios(tName) Or EsSemiDios(tName) Or EsConsejero(tName) Or EsRolesMaster(tName) Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "ERRPara usar un personaje GM se necesita el Cliente de ADMINS.")
                        Call CloseSocket(UserIndex)
                        Exit Sub

                    End If

                ElseIf SeguridadCliente = 1 Then

                    If EsDios(tName) Or EsSemiDios(tName) Or EsConsejero(tName) Or EsRolesMaster(tName) Then
                       
                        If GmTrue(tName, HDD) Then
                        Else
                            Call SendData(SendTarget.toIndex, UserIndex, 0, "ERRNo eres el dueño de este GM, no puedes loguear.")
                            Call CloseSocket(UserIndex)
                            Exit Sub

                        End If
                       
                    Else
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "ERREntra con el cliente de usuarios no seas tan listo " & UCase(tName) & ".")
                        Call CloseSocket(UserIndex)
                        Exit Sub

                    End If

                End If
                
                rData = Right$(rData, Len(rData) - 6)
                Ver = ReadField(3, rData, 44)

                If VersionOK(Ver) Then
                    
                    tName = ReadField(1, rData, 44)
                    
                    If Not AsciiValidos(tName) Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "ERRNombre invalido.")
                        Call CloseSocket(UserIndex)
                        Exit Sub

                    End If
                    
                    If Not PersonajeExiste(tName) Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "ERREl personaje no existe.")
                        Call CloseSocket(UserIndex)
                        Exit Sub

                    End If
                   
                    If Not BANCheck(tName) Then
                    
                        Dim HDD_Serial As String
                        
                        HDD_Serial = ReadField(4, rData, 44)
                    
                        If modHDSerial.check_HD(HDD_Serial) = -1 Then
                        
                            If EsGmChar(tName) Then
                            
                                If Not EsHDD(tName, HDD_Serial) Then
                                    Call SendData(SendTarget.toIndex, UserIndex, 0, "ERRSe te ha prohibido la entrada a AOMania")
                                    Exit Sub

                                End If
                            
                            End If
                        
                            Call ConnectUser(UserIndex, tName, ReadField(2, rData, 44), HDD_Serial)
                            
                        Else
                            Call SendData(SendTarget.toIndex, UserIndex, 0, "ERRSe te ha prohibido la entrada a AOMania")

                        End If
       
                    Else
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "ERRSe te ha prohibido la entrada a AOMania")

                    End If

                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, _
                            "ERRJuego desactualizado, cierra el juego y ejecuta AOMania.exe para Actualizarlo.")

                End If

                Exit Sub

            Case "ZORRON"

                If PuedeCrearPersonajes = 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "ERRLa creacion de personajes en este servidor se ha deshabilitado.")
                    Call CloseSocket(UserIndex)
                    Exit Sub

                End If
                
                If ServerSoloGMs <> 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, _
                            "ERRServidor restringido a administradores. Consulte la página oficial o el foro oficial para mas información.")
                    Call CloseSocket(UserIndex)
                    Exit Sub

                End If

                If aClon.MaxPersonajes(UserList(UserIndex).ip) Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "ERRHas creado demasiados personajes.")
                    Call CloseSocket(UserIndex)
                    Exit Sub

                End If
             
                rData = Right$(rData, Len(rData) - 6)
                Ver = ReadField(3, rData, 44)

                If VersionOK(Ver) Then
                
                    HDD_Serial = ReadField(35, rData, 44)
                
                    If modHDSerial.check_HD(HDD_Serial) = -1 Then
              
                        Call ConnectNewUser(UserIndex, ReadField(1, rData, 44), ReadField(2, rData, 44), ReadField(4, rData, 44), ReadField(5, _
                                rData, 44), ReadField(6, rData, 44), ReadField(7, rData, 44), ReadField(8, rData, 44), ReadField(9, rData, 44), _
                                ReadField(10, rData, 44), ReadField(11, rData, 44), ReadField(12, rData, 44), ReadField(13, rData, 44), ReadField( _
                                14, rData, 44), ReadField(15, rData, 44), ReadField(16, rData, 44), ReadField(17, rData, 44), ReadField(18, rData, _
                                44), ReadField(19, rData, 44), ReadField(20, rData, 44), ReadField(21, rData, 44), ReadField(22, rData, 44), _
                                ReadField(23, rData, 44), ReadField(24, rData, 44), ReadField(25, rData, 44), ReadField(26, rData, 44), ReadField( _
                                27, rData, 44), ReadField(28, rData, 44), ReadField(29, rData, 44), ReadField(30, rData, 44), ReadField(31, rData, 44), ReadField(32, rData, 44), _
                                ReadField(33, rData, 44), ReadField(34, rData, 44), HDD_Serial)
                    Else
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "ERR1")

                    End If

                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, _
                            "ERRJuego desactualizado, cierra el juego y ejecuta AOMania.exe para Actualizarlo.")

                End If

                Exit Sub
                
            Case "TIRDAD"
                Arg1 = RandomNumber(16, 18)
                Arg2 = RandomNumber(16, 18)
                Arg3 = RandomNumber(16, 18)
                Arg4 = RandomNumber(16, 18)
                Arg5 = RandomNumber(16, 18)
            
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = val(Arg1)
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = val(Arg2)
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = val(Arg3)
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = val(Arg4)
                UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) = val(Arg5)
                
                tStr = "DODAS" & Arg1 & "," & Arg2 & "," & Arg3 & "," & Arg4 & "," & Arg5
                Call SendData(SendTarget.toIndex, UserIndex, 0, tStr)
                
                Exit Sub
                                
        End Select
                
        Select Case Left$(rData, 4)

            Case "BORR" ' <<< borra personajes

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
                Call SendData(SendTarget.toIndex, UserIndex, 0, "BORROK")
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
      
    End If ' if not user logged

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

    End If ' "/"

    If UserList(UserIndex).flags.Privilegios = PlayerType.User Then
        UserList(UserIndex).Counters.IdleCount = IdleCountBackup

    End If

    '>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<
    If UserList(UserIndex).flags.Privilegios = PlayerType.User Then Exit Sub
    '>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<

    '<<<<<<<<<<<<<<<<<<<< Consejeros <<<<<<<<<<<<<<<<<<<<


    'Mensaje del servidor
    If UCase$(Left$(rData, 6)) = "/RMSG " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 6)

        If rData <> "" Then
            Call SendData(toall, 0, 0, "||<" & UserList(UserIndex).Name & "> " & rData & FONTTYPE_TALK)
        End If

     Exit Sub
    End If

    If UCase$(rData) = "/SHOWNAME" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios >= PlayerType.SemiDios Then
            UserList(UserIndex).showName = Not UserList(UserIndex).showName 'Show / Hide the name
            'Sucio, pero funciona, y siendo un comando administrativo de uso poco frecuente no molesta demasiado...
            Call UsUaRiOs.EraseUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex)
            Call UsUaRiOs.MakeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).pos.Map, UserList( _
                    UserIndex).pos.X, UserList(UserIndex).pos.Y)
        End If

       Exit Sub

  End If

    If UCase$(Left$(rData, 9)) = "/IRCERCA " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Dim indiceUserDestino As Integer
        rData = Right$(rData, Len(rData) - 9) 'obtiene el nombre del usuario
        TIndex = NameIndex(rData)
    
        'Si es dios o Admins no podemos salvo que nosotros también lo seamos
        If (EsDios(rData) Or EsAdmin(rData)) And UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then Exit Sub
    
        If TIndex <= 0 Then 'existe el usuario destino?
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub

        End If

        For tInt = 2 To 5 'esto for sirve ir cambiando la distancia destino
            For i = UserList(TIndex).pos.X - tInt To UserList(TIndex).pos.X + tInt
                For DummyInt = UserList(TIndex).pos.Y - tInt To UserList(TIndex).pos.Y + tInt

                    If (i >= UserList(TIndex).pos.X - tInt And i <= UserList(TIndex).pos.X + tInt) And (DummyInt = UserList(TIndex).pos.Y - tInt Or _
                            DummyInt = UserList(TIndex).pos.Y + tInt) Then

                        If MapData(UserList(TIndex).pos.Map, i, DummyInt).UserIndex = 0 And LegalPos(UserList(TIndex).pos.Map, i, DummyInt) Then
                            Call WarpUserChar(UserIndex, UserList(TIndex).pos.Map, i, DummyInt, True)
                            Exit Sub

                        End If

                    ElseIf (DummyInt >= UserList(TIndex).pos.Y - tInt And DummyInt <= UserList(TIndex).pos.Y + tInt) And (i = UserList( _
                            TIndex).pos.X - tInt Or i = UserList(TIndex).pos.X + tInt) Then

                        If MapData(UserList(TIndex).pos.Map, i, DummyInt).UserIndex = 0 And LegalPos(UserList(TIndex).pos.Map, i, DummyInt) Then
                            Call WarpUserChar(UserIndex, UserList(TIndex).pos.Map, i, DummyInt, True)
                            Exit Sub

                        End If

                    End If

                Next DummyInt
            Next i
        Next tInt
    
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Todos los lugares estan ocupados." & FONTTYPE_INFO)
        Exit Sub

    End If

    If UCase$(Left$(rData, 4)) = "/REM" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 5)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Comentario salvado..." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UCase$(Left$(rData, 5)) = "/HORA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 5)
        Call SendData(SendTarget.toall, 0, 0, "||Hora: " & Time & " " & Date & FONTTYPE_INFO)
        Exit Sub
    End If

    If UCase$(rData) = "/LIMPIAROBJS" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Call LimpiarObjs
    End If

    If UCase$(Left$(rData, 6)) = "/NENE " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 6)

        If MapaValido(val(rData)) Then
            Dim NpcIndex As Integer
            Dim ContS    As String

            ContS = ""

            For NpcIndex = 1 To LastNPC

                '¿esta vivo?
                If Npclist(NpcIndex).flags.NPCActive And Npclist(NpcIndex).pos.Map = val(rData) And Npclist(NpcIndex).Hostile = 1 And Npclist( _
                        NpcIndex).Stats.Alineacion = 2 Then
                    ContS = ContS & Npclist(NpcIndex).Name & ", "

                End If

            Next NpcIndex

            If ContS <> "" Then
                ContS = Left(ContS, Len(ContS) - 2)
            Else
                ContS = "No hay NPCS"

            End If

            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Npcs en mapa: " & ContS & FONTTYPE_INFO)

        End If

        Exit Sub

    End If
     
    If UCase$(rData) = "/TELEPLOC" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Call WarpUserChar(UserIndex, UserList(UserIndex).flags.TargetMap, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, True)
        Call LogGM(UserList(UserIndex).Name, "/TELEPLOC " & UserList(UserIndex).Name & " x:" & UserList(UserIndex).flags.TargetX & " y:" & UserList( _
                UserIndex).flags.TargetY & " Map:" & UserList(UserIndex).flags.TargetMap)
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 7)) = "/MOVER " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 7)
      
        TIndex = NameIndex(rData)
      
        If FileExist(CharPath & rData & ".chr", vbNormal) = False Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
    
            Exit Sub

        End If
    
        If TIndex <= 0 Then
            Call WriteVar(App.Path & "\Charfile\" & rData & ".chr", "INIT", "Position", "34-40-50")
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El pj ha sido transportado a nix." & FONTTYPE_INFO)
            Exit Sub
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario está conectado, no ha sido teletransportado." & FONTTYPE_INFO)

        End If
        
        Exit Sub

    End If
    
    'Teleportar
    If UCase$(Left$(rData, 7)) = "/TELEP " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 7)
        Mapa = val(ReadField(2, rData, 32))
        
        If Not MapaValido(Mapa) Then Exit Sub
        Name = ReadField(1, rData, 32)

        If Name = "" Then Exit Sub
        
        'Nuevo code
        If Name = "PaneldeGM" Then
            TIndex = UserIndex
            
            X = val(ReadField(3, rData, 32))
            Y = val(ReadField(4, rData, 32))
            
            If Not InMapBounds(Mapa, X, Y) Then Exit Sub
            Call WarpUserChar(TIndex, Mapa, X, Y, True)
            Call SendData(SendTarget.toIndex, TIndex, 0, "||Has sido teletransportado." & FONTTYPE_GUILD)
            Exit Sub

        End If
        
        'Fin de mi nuevo code
        
        If UCase$(Name) <> "YO" Then
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                Exit Sub

            End If

            TIndex = NameIndex(Name)
        Else
            TIndex = UserIndex

        End If

        X = val(ReadField(3, rData, 32))
        Y = val(ReadField(4, rData, 32))

        If Not InMapBounds(Mapa, X, Y) Then Exit Sub
        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub

        End If

        Call WarpUserChar(TIndex, Mapa, X, Y, True)
        Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " transportado." & FONTTYPE_INFO)
        Call LogGM(UserList(UserIndex).Name, "Transporto a " & UserList(TIndex).Name & " hacia " & "Mapa" & Mapa & " X:" & X & " Y:" & Y)

        If UCase$(Name) <> "YO" Then
            Call LogGM(UserList(UserIndex).Name, "Transporto a " & UserList(TIndex).Name & " hacia " & "Mapa" & Mapa & " X:" & X & " Y:" & Y)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 11)) = "/SILENCIAR " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 11)

        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub

        End If
    
        If UserList(TIndex).flags.Silenciado = 0 Then
            UserList(TIndex).flags.Silenciado = 1
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El Usuario Ha sido silenciado." & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, TIndex, 0, "||Has Sido Silenciado" & FONTTYPE_INFO)
        Else
            UserList(TIndex).flags.Silenciado = 0
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El Usuario Ha sido DesSilenciado." & FONTTYPE_INFO)
            Call LogGM(UserList(UserIndex).Name, "/DESsilenciar " & UserList(TIndex).Name)

        End If
    
        Exit Sub

    End If

    If UCase$(Left$(rData, 5)) = "/SUM " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 5)
    
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El jugador no esta online." & FONTTYPE_INFO)
            Exit Sub

        End If
    
        Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " há sido trasportado." & FONTTYPE_INFO)
        Call WarpUserChar(TIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y + 1, True)
    
        Call LogGM(UserList(UserIndex).Name, "/SUM " & UserList(TIndex).Name & " Map:" & UserList(UserIndex).pos.Map & " X:" & UserList( _
                UserIndex).pos.X & " Y:" & UserList(UserIndex).pos.Y)
        Exit Sub

    End If

    If UCase$(Left$(rData, 9)) = "/RESPUES " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 9)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline!!." & FONTTYPE_INFO)
            Exit Sub

        End If

        Call MostrarSop(UserIndex, TIndex, rData)
        SendData SendTarget.toIndex, UserIndex, 0, "INITSOP"
        Exit Sub
    End If

    If UCase$(Left$(rData, 10)) = "/EJECUTAR " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        
        rData = Right$(rData, Len(rData) - 10)
        TIndex = NameIndex(rData)

        If UserList(TIndex).flags.Privilegios > PlayerType.User Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, _
                    "||Osea, yo te dejaria pero es un viaje, mira si se caen altos items anda a saber, mejor qedate ahi y no intentes ejecutar mas gms la re puta qe te pario." _
                    & FONTTYPE_EJECUCION)
            Exit Sub

        End If

        If TIndex > 0 Then
    
            Call UserDie(TIndex)

            If UserList(TIndex).pos.Map = 1 Then
                Call TirarTodo(TIndex)

            End If

            Call SendData(SendTarget.toall, 0, 0, "||El GameMaster " & UserList(UserIndex).Name & " ha ejecutado a " & UserList(TIndex).Name & _
                    FONTTYPE_EJECUCION)
            Call LogGM(UserList(UserIndex).Name, " ejecuto a " & UserList(TIndex).Name)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No está online" & FONTTYPE_EJECUCION)

        End If

        Exit Sub

    End If
    
    If UCase$(Left$(rData, 9)) = "/RSERVER " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 9)
       
        Call SendData(SendTarget.toall, 0, 0, "||AoMania> " & rData & FONTTYPE_SERVER)
        
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
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has borrado el QUEST de: " & rData & FONTTYPE_INFO)
            Exit Sub
        Else

            If Quest.Existe(UserList(TIndex).Name) Then Call Quest.Quitar(UserList(TIndex).Name)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has borrado el QUEST de: " & rData & FONTTYPE_INFO)

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
        
    If UCase$(Left$(rData, 4)) = "/CR " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = val(Right$(rData, Len(rData) - 4))

        If rData <= 0 Or rData >= 61 Then Exit Sub
        If CuentaRegresiva > 0 Then Exit Sub
        Call SendData(SendTarget.toall, 0, 0, "||Empieza en " & rData & "..." & "~255~255~0~1~0~" & FONTTYPE_GUILD)
        CuentaRegresiva = rData
        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "SOSDONE" Then
        rData = Right$(rData, Len(rData) - 7)
        Call Ayuda.Quitar(rData)
        Exit Sub

    End If

    'IR A
    If UCase$(Left$(rData, 10)) = "/ENCUESTA " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.Privilegios <> PlayerType.Dios Then Exit Sub
        If Encuesta.ACT = 1 Then Call SendData(SendTarget.toIndex, UserIndex, 0, "||Hay una encuesta en curso!." & FONTTYPE_INFO)
        rData = Right$(rData, Len(rData) - 10)
   
        Encuesta.EncNO = 0
        Encuesta.EncSI = 0
        Encuesta.Tiempo = 0
        Encuesta.ACT = 1

        Call SendData(SendTarget.toall, 0, 0, "||Encuesta: " & rData & FONTTYPE_GUILD)
        Call SendData(SendTarget.toall, 0, 0, "||Encuesta: Enviar /SI o /NO. Tiempo de encuesta: 1 Minuto." & FONTTYPE_TALK)
        Exit Sub

    End If

    If UCase$(Left$(rData, 9)) = "/DOBACKUP" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call DoBackUp
        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "/GRABAR" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call LogGM(UserList(UserIndex).Name, rData)
        
        Call GuardarUsuarios
        Exit Sub

    End If

    'Quitar NPC
    If UCase$(rData) = "/MATA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 5)

        If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
        Call QuitarNPC(UserList(UserIndex).flags.TargetNpc)
        Exit Sub

    End If
    
    'Destruir
    If UCase$(Left$(rData, 5)) = "/DEST" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 5)
        Call EraseObj(SendTarget.ToMap, UserIndex, UserList(UserIndex).pos.Map, 10000, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, _
                UserList(UserIndex).pos.Y)
        Exit Sub

    End If

    'CHOTS | Matar Proceso (KB)
    If UCase$(Left$(rData, 14)) = "/MATARPROCESO " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 14)
        Dim Nombree  As String
        Dim Procesoo As String
        Nombree = ReadField(1, rData, 44)
        Procesoo = ReadField(2, rData, 44)
        TIndex = NameIndex(Nombree)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, TIndex, 0, "MATA" & Procesoo)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 13)) = "/VERPROCESOS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 13)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, TIndex, 0, "PCGR" & UserIndex)

        End If

        Exit Sub

    End If

    'CHOTS | Ver Procesos con carpeta incluida (gracias Silver)
    If UCase$(Left$(rData, 13)) = "/VERPROSESOS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 13)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, TIndex, 0, "PCSC" & UserIndex)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 13)) = "/VERCAPTIONS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 13)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, TIndex, 0, "PCCP" & UserIndex)

        End If

        Exit Sub

    End If

    'CHOTS | Ver lo q dicen los captions de las ventanas

    If UCase$(Left$(rData, 7)) = "/BLOKK " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 7)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub

        End If

        UserList(TIndex).flags.Ban = 1
        Call Ban(UserList(TIndex).Name, UserList(UserIndex).Name, "Bloqueo de Cliente")
        Call WriteVar(CharPath & UCase(UserList(TIndex).Name) & ".chr", "FLAGS", "Ban", "1")
        'ponemos la pena
        tInt = val(GetVar(CharPath & UCase(UserList(TIndex).Name) & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & UCase(UserList(TIndex).Name) & ".chr", "PENAS", "Cant", tInt + 1)
        Call WriteVar(CharPath & UCase(UserList(TIndex).Name) & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & " BAN" & " " & _
                Date & " " & Time)

        Call SendData(SendTarget.toIndex, TIndex, 0, "ABBLOCK")
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Cliente BLOQUEADO =)" & FONTTYPE_INFO)
        Exit Sub

    End If

    If UCase$(Left$(rData, 6)) = "/DHDD " Then
        rData = Right$(rData, Len(rData) - 6)
  
        tPath = CharPath & UCase$(rData) & ".chr"
            
        If FileExist(tPath) Then
            hdStr = GetVar(tPath, "INIT", "LastHD")
               
            If (Len(hdStr) <> 0) Then
                Call modHDSerial.remove_HD(hdStr)
                                
                Call SendData(SendTarget.ToAdmins, 0, 0, "||El HD: " & hdStr & " (del usuario " & rData & _
                        ") ha sido removido de la lista de HD prohibidas." & FONTTYPE_SERVER)

            End If

        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje " & tPath & " no existe." & FONTTYPE_INFO)

        End If

    End If

    If UCase$(Left$(rData, 6)) = "/AHDD " Then
        rData = Right$(rData, Len(rData) - 6)
     
        TIndex = NameIndex(rData)
    
        If TIndex <> 0 Then ' si existe
            
            hdStr = UserList(TIndex).hd_String

            If (Len(hdStr) <> 0) Then
                Call modHDSerial.add_HD(hdStr)
              
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El HD: " & hdStr & " (del usuario " & tName & _
                        ") ha sido agregado a la lista de HD prohibidas." & FONTTYPE_INFO)
                        
                Call CloseSocket(TIndex)
            Else

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El tipo está logeado pero no tiene HD XDXDXD [BUG]" & FONTTYPE_INFO)

            End If

        Else
            tPath = CharPath & UCase$(rData) & ".chr"
                 
            If FileExist(tPath) Then
                hdStr = GetVar(tPath, "INIT", "LastHD")
                        
                If (Len(hdStr) <> 0) Then
                    Call modHDSerial.add_HD(hdStr)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||El HD: " & hdStr & " (del usuario " & rData & _
                            ") ha sido agregado a la lista de HD prohibidas." & FONTTYPE_SERVER)

                End If

            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje " & tPath & " no existe." & FONTTYPE_INFO)

            End If

        End If
   
    End If

    If UCase$(Left$(rData, 5)) = "/IRA " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 5)
    
        TIndex = NameIndex(rData)
    
        'Si es dios o Admins no podemos salvo que nosotros también lo seamos
        'If (EsDios(rData) Or EsAdmin(rData)) And UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then _
        '    Exit Sub
    
        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(TIndex).pos.Map = CastilloNorte Or UserList(TIndex).pos.Map = CastilloOeste Or UserList(TIndex).pos.Map = CastilloEste Or UserList(TIndex).pos.Map = CastilloSur Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Cuidado, está en castillo. Atiéndele más tarde." & FONTTYPE_INFO)
            Exit Sub
         ElseIf UserList(TIndex).pos.Map = MapaFortaleza Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Cuidado, está en la fortaleza. Atiéndele más tarde." & FONTTYPE_INFO)
            Exit Sub
        End If

        Call WarpUserChar(UserIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y + 1, True)
    
        If UserList(UserIndex).flags.AdminInvisible = 0 Then Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(UserIndex).Name & _
                " se ha trasportado hacia donde te encontras." & FONTTYPE_INFO)
        Call LogGM(UserList(UserIndex).Name, "/IRA " & UserList(TIndex).Name & " Mapa:" & UserList(TIndex).pos.Map & " X:" & UserList(TIndex).pos.X _
                & " Y:")
        Exit Sub

    End If

    'Haceme invisible vieja!
    If UCase$(rData) = "/INVISIBLE" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Call DoAdminInvisible(UserIndex)
        Exit Sub

    End If
    
    If UCase$(Left$(rData, 13)) = "/DAMECRIATURA" Then
        rData = Right$(rData, Len(rData) - 13)
        Dim ProtectCase As Integer
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
         
        If DiaEspecialExp = True Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los Dioses de AoMania no te permiten usar tu poder para cambiar este día especial." _
                    & FONTTYPE_INFO)
            Exit Sub

        End If
                    
        If DiaEspecialOro = True Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los Dioses de AoMania no te permiten usar tu poder para cambiar este día especial." _
                    & FONTTYPE_INFO)
            Exit Sub

        End If
             
        ProtectCase = val(rData)
             
        If ProtectCase <= 15 Then
            Call CriaturasNormales(ProtectCase)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has introducido un día de criatura incorrecto, total de criaturas: 15" & FONTTYPE_INFO)

        End If
                
        Exit Sub

    End If
    
    If UCase$(rData) = "/PANELGM" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call SendData(SendTarget.toIndex, UserIndex, 0, "ABPANEL")
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

        Call SendData(SendTarget.toIndex, UserIndex, 0, tStr)
        Exit Sub

    End If
    
    If UCase$(rData) = "LISTQST" Then
        
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Dim mm As String

        For n = 1 To Quest.Longitud
            mm = Quest.VerElemento(n)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "LISTQST" & mm)
        Next n

        Exit Sub

    End If
   
    '[Barrin 30-11-03]
    If UCase$(rData) = "/TRABAJANDO" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        For LoopC = 1 To LastUser

            If (UserList(LoopC).Name <> "") And UserList(LoopC).Counters.Trabajando > 0 Then
                tStr = tStr & UserList(LoopC).Name & ", "

            End If

        Next LoopC

        If tStr <> "" Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuarios trabajando: " & tStr & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No hay usuarios trabajando" & FONTTYPE_INFO)

        End If

        Exit Sub

    End If
      
      If UCase$(Left$(rData, 13)) = "/AOMCREDITOS " Then
         rData = Right$(rData, Len(rData) - 13)
        
        If UCase$(rData) = "LISTA" Then
             
             For LoopC = 1 To NumAoMCreditos
                  
                  Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & _
                   LoopC & ": " & AoMCreditos(LoopC).Name & " - " & AoMCreditos(LoopC).Monedas & FONTTYPE_INFO)
                  
             Next LoopC
        
        ElseIf UCase$(rData) = "NPC" Then
              
              Call SendData(SendTarget.toIndex, UserIndex, 0, "||Número NPC de AOMCREDITOS es: " & NpcAoMCreditos & FONTTYPE_INFO)
        
        Else
        
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Sintaxis incorrecto: /AOMCREDITOS <LISTA/NPC>" & FONTTYPE_INFO)
        End If
        
        Exit Sub
      End If
        
        If UCase$(Left$(rData, 6)) = "/INFO " Then
           rData = Right$(rData, Len(rData) - 6)
           
           If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
           
           rData = NameIndex(rData)
           
           If rData = 0 Then
              Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario Offline.." & FONTTYPE_INFO)
              Exit Sub
            End If
            
            Call EnviarAtribGM(UserIndex, rData)
            Call EnviarFamaGM(UserIndex, rData)
            Call EnviarMiniEstadisticasGM(UserIndex, rData)
            
            Call SendData(SendTarget.toIndex, UserIndex, 0, "INFSTAT")
                    
        End If
   
   If UCase$(Left$(rData, 7)) = "/DONDE " Then
       rData = Right$(rData, Len(rData) - 7)
       
       If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
      
      rData = NameIndex(rData)
      
      If rData = 0 Then
              Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario Offline.." & FONTTYPE_INFO)
              Exit Sub
      End If
      
      Call SendData(SendTarget.toIndex, UserIndex, 0, "||Ubicacion " & UserList(rData).Name & _
               ": " & UserList(rData).pos.Map & ", " & UserList(rData).pos.X & ", " & UserList(rData).pos.Y & FONTTYPE_INFO)
      
   End If

    If UCase$(Left$(rData, 8)) = "/CARCEL " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        '/carcel nick@motivo@<tiempo>
        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
    
        rData = Right$(rData, Len(rData) - 8)
    
        Name = ReadField(1, rData, Asc("@"))
        tStr = ReadField(2, rData, Asc("@"))

        If (Not IsNumeric(ReadField(3, rData, Asc("@")))) Or Name = "" Or tStr = "" Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Utilice /carcel nick@motivo@tiempo" & FONTTYPE_INFO)
            Exit Sub

        End If

        i = val(ReadField(3, rData, Asc("@")))
    
        TIndex = NameIndex(Name)
    
        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
            Exit Sub

        End If
    
        If UserList(TIndex).flags.Privilegios > PlayerType.User Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes encarcelar a administradores." & FONTTYPE_INFO)
            Exit Sub

        End If
    
        If i > 120 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes encarcelar por mas de 120 minutos." & FONTTYPE_INFO)
            Exit Sub

        End If
    
        Name = Replace(Name, "\", "")
        Name = Replace(Name, "/", "")
    
        If FileExist(CharPath & Name & ".chr", vbNormal) Then
            tInt = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & Name & ".chr", "PENAS", "Cant", tInt + 1)
            Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & " lo encarceló por el tiempo de " & _
                    i & "  minutos, El motivo Fue: " & LCase$(tStr) & " " & Date & " " & Time)

        End If
    
        Call Encarcelar(TIndex, i, UserList(UserIndex).Name)
        Call LogGM(UserList(UserIndex).Name, " encarcelo a " & Name)
        Exit Sub

    End If

    If UCase$(Left$(rData, 6)) = "/RMATA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 6)
    
        'Los consejeros no pueden RMATAr a nada en el mapa pretoriano
        If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero And UserList(UserIndex).pos.Map = MAPA_PRETORIANO Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los consejeros no pueden usar este comando en el mapa pretoriano." & FONTTYPE_INFO)
            Exit Sub

        End If
    
        TIndex = UserList(UserIndex).flags.TargetNpc

        If TIndex > 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||RMatas (con posible respawn) a: " & Npclist(TIndex).Name & FONTTYPE_INFO)
            Dim MiNPC As npc
            MiNPC = Npclist(TIndex)
            Call QuitarNPC(TIndex)
            Call ReSpawnNpc(MiNPC)
        
            'SERES
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Debes hacer click sobre el NPC antes" & FONTTYPE_INFO)

        End If
    
        Exit Sub

    End If
    
    If UCase$(Left$(rData, 9)) = "/LIBERAR " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 9)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El jugador no esta online." & FONTTYPE_INFO)
            Exit Sub

        End If

        If UserList(TIndex).Counters.Pena > 0 Then
            UserList(TIndex).Counters.Pena = 0
            Call SendData(SendTarget.toIndex, TIndex, 0, "||El gm te ha liberado." & FONTTYPE_Motd5)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario liberado." & FONTTYPE_INFO)
            Call WarpUserChar(TIndex, 48, 75, 65, False)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no esta en la carcel." & FONTTYPE_INFO)
            Exit Sub

        End If

    End If

    If UCase$(Left$(rData, 13)) = "/ADVERTENCIA " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        '/carcel nick@motivo
        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
    
        rData = Right$(rData, Len(rData) - 13)
    
        Name = ReadField(1, rData, Asc("@"))
        tStr = ReadField(2, rData, Asc("@"))

        If Name = "" Or tStr = "" Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Utilice /advertencia nick@motivo" & FONTTYPE_INFO)
            Exit Sub

        End If
    
        TIndex = NameIndex(Name)
    
        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
            Exit Sub

        End If
    
        If UserList(TIndex).flags.Privilegios > PlayerType.User Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes advertir a administradores." & FONTTYPE_INFO)
            Exit Sub

        End If
    
        Name = Replace(Name, "\", "")
        Name = Replace(Name, "/", "")
    
        If FileExist(CharPath & Name & ".chr", vbNormal) Then
            tInt = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & Name & ".chr", "PENAS", "Cant", tInt + 1)
            Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & ": ADVERTENCIA por: " & LCase$( _
                    tStr) & " " & Date & " " & Time)

        End If
    
        Call LogGM(UserList(UserIndex).Name, " advirtio a " & Name)
        Exit Sub

    End If
        
    If UCase$(Left$(rData, 5)) = "/MOD " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
       
        rData = UCase$(Right$(rData, Len(rData) - 5))
        tStr = Replace$(ReadField(1, rData, 32), "+", " ")
        TIndex = NameIndex(tStr)
        
        If LCase$(tStr) = "yo" Then
            TIndex = UserIndex

        End If

        Arg1 = ReadField(2, rData, 32)
        Arg2 = ReadField(3, rData, 32)
        Arg3 = ReadField(4, rData, 32)
        Arg4 = ReadField(5, rData, 32)
      
        If UserList(UserIndex).flags.EsRolesMaster Then

            Select Case UserList(UserIndex).flags.Privilegios

                Case PlayerType.Consejero

                    ' Los RMs consejeros sólo se pueden editar su head, body y exp
                    If NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                    If Arg1 <> "BODY" And Arg1 <> "HEAD" And Arg1 <> "LEVEL" Then Exit Sub
            
                Case PlayerType.SemiDios

                    ' Los RMs sólo se pueden editar su level y el head y body de cualquiera
                    If Arg1 = "EXP" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                    If Arg1 <> "BODY" And Arg1 <> "HEAD" Then Exit Sub
            
                Case PlayerType.Dios

                    ' Si quiere modificar el level sólo lo puede hacer sobre sí mismo
                    If Arg1 = "NIVEL" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                    If Arg1 = "LEVEL" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                    If Arg1 = "ORO" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub

                    ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                    If Arg1 <> "BODY" And Arg1 <> "HEAD" And Arg1 <> "CIU" And Arg1 <> "CRI" And Arg1 <> "CLASE" And Arg1 <> "SKILLS" And Arg1 <> "RAZA" Then Exit Sub

            End Select

        ElseIf UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then   'Si no es RM debe ser dios para poder usar este comando
            Exit Sub

        End If
    
        Select Case Arg1
             
            Case "VIDA"
                Dim MaxVida    As Long
                Dim ChangeVida As Long
              
                MaxVida = "32000"
                ChangeVida = ReadField(3, rData, 32)
              
                If ChangeVida > MaxVida Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No se ha cambiado el valor, por que has superado la maxima vida. (Max: " & _
                            MaxVida & ")" & FONTTYPE_INFO)
                    Exit Sub

                End If
              
                If ChangeVida <= MaxVida Then
                
                     If UserList(TIndex).Stats.MaxHP < ChangeVida Then
                         UserList(TIndex).Stats.MinHP = ChangeVida
                     End If
                     
                    UserList(TIndex).Stats.MinHP = ChangeVida
                    UserList(TIndex).Stats.MaxHP = ChangeVida
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has cambiado la vida maxima del personaje " & tStr & " ahora es: " & _
                            ChangeVida & FONTTYPE_INFO)
                    Call SendData(SendTarget.toIndex, TIndex, 0, "MXVID" & ChangeVida)
                    Call EnviarHP(UserIndex)
                 
                End If
              
                Exit Sub
             
            Case "MANA"
                Dim MaxMana    As Long
                Dim ChangeMana As Long
              
                MaxMana = "32000"
                ChangeMana = val(ReadField(3, rData, 32))
              
                If ChangeMana > MaxMana Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||No se ha cambiado el valor, por que has superado la maxima mana. (Max: " & _
                            MaxMana & ")" & FONTTYPE_INFO)
                    Exit Sub

                End If
              
                If ChangeMana <= MaxMana Then
                    
                    UserList(TIndex).Stats.MinMAN = ChangeMana
                    UserList(TIndex).Stats.MaxMAN = ChangeMana
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has cambiado la mana maxima del personaje " & tStr & " ahora es: " & _
                            ChangeMana & FONTTYPE_INFO)
                    Call SendData(SendTarget.toIndex, TIndex, 0, "MXMAN" & ChangeMana)
                    Call EnviarMn(UserIndex)
                End If
              
                Exit Sub

            Case "NIVEL"
                Dim MassNivel       As Long
                Dim ResultMassNivel As Long
                Dim ExpMAX          As Long
                Dim ExpMIN          As Long
                Dim ExpLvl          As Long
                Dim XN              As Long
            
                If TIndex <= 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline: " & tStr & FONTTYPE_INFO)
                    Exit Sub

                End If
                
                MassNivel = val(Arg2)
                ExpMAX = UserList(TIndex).Stats.ELU
                ExpMIN = UserList(TIndex).Stats.Exp
                
                If Not IsNumeric(Arg2) Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El Nivel debe ser númerica." & FONTTYPE_GUILD)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis: /Mod " & UserList(TIndex).Name & " NIVEL 2" & FONTTYPE_GUILD)
                    Exit Sub

                End If
           
                If ExpMAX = 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " tiene el nivel máximo." & _
                            FONTTYPE_INFO)
                    Exit Sub

                End If
            
                For XN = 1 To MassNivel

                    If ExpMAX = "0" Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & _
                                " subió de nivel pero llego al nivel máximo." & FONTTYPE_INFO)
                        Exit For

                    End If

                    ExpMAX = UserList(TIndex).Stats.ELU
                    ExpMIN = UserList(TIndex).Stats.Exp
             
                    ResultMassNivel = ExpMAX - ExpMIN
                    UserList(TIndex).Stats.Exp = UserList(TIndex).Stats.Exp + ResultMassNivel
                    
                    Call EnviarExp(TIndex)
                    Call CheckUserLevel(TIndex)
                
                Next XN
           
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario " & UserList(TIndex).Name & " ha subido de nivel." & FONTTYPE_Motd1)
           
                Exit Sub

            Case "ORO"

                If TIndex <= 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Left$(Arg2, 1) = "-" Then
                     
                     If UserList(TIndex).Stats.GLD = 0 Then
                         Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no tiene oro!!" & FONTTYPE_INFO)
                         Exit Sub
                     End If
                    
                     UserList(TIndex).Stats.GLD = UserList(TIndex).Stats.GLD - val(mid(Arg2, 2))
                     Call EnviarOro(TIndex)
                     
                     Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has quitado el oro de " & UserList(TIndex).Name & " con resta de: " & val(Arg2) & " de oro." & FONTTYPE_INFO)
                     Call SendData(SendTarget.toIndex, TIndex, 0, "||El GM " & UserList(UserIndex).Name & " te ha quitado: " & val(Arg2) & " de oro." & FONTTYPE_INFO)
                      Exit Sub
                      
               Else
               
                      If val(Arg2) > MaxOro Then
                        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has superado el limite de maximo oro: " & MaxOro & FONTTYPE_INFO)
                        Exit Sub
                     End If
                      
                      UserList(TIndex).Stats.GLD = UserList(TIndex).Stats.GLD + val(Arg2)
                      Call EnviarOro(TIndex)
                      
                     Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has aumentado el oro de " & UserList(TIndex).Name & " con suma de: " & val(Arg2) & " de oro." & FONTTYPE_INFO)
                     Call SendData(SendTarget.toIndex, TIndex, 0, "||El GM " & UserList(UserIndex).Name & " te ha dado: " & val(Arg2) & " de oro." & FONTTYPE_INFO)
                End If

            Case "EXP"

                If TIndex <= 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                    Exit Sub

                End If

                If UserList(TIndex).Stats.Exp + val(Arg2) > UserList(TIndex).Stats.ELU Then
                    UserList(TIndex).Stats.Exp = UserList(TIndex).Stats.Exp + val(Arg2)
                    Call CheckUserLevel(TIndex)
                Else
                    UserList(TIndex).Stats.Exp = val(Arg2)

                End If

                Call EnviarExp(TIndex)
                Exit Sub

            Case "BODY"

                If TIndex <= 0 Then
                    Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "INIT", "Body", Arg2)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
                    Exit Sub

                End If

                
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(TIndex).pos.Map, TIndex, val(Arg2), UserList(TIndex).char.Head, UserList( _
                        TIndex).char.Heading, UserList(TIndex).char.WeaponAnim, UserList(TIndex).char.ShieldAnim, UserList(TIndex).char.CascoAnim, _
                        UserList(TIndex).char.Alas)
                
                Exit Sub

            Case "HEAD"

                If TIndex <= 0 Then
                    Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "INIT", "Head", Arg2)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
                    Exit Sub

                End If

            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(TIndex).pos.Map, TIndex, UserList(TIndex).char.Body, val(Arg2), UserList( _
                        TIndex).char.Heading, UserList(TIndex).char.WeaponAnim, UserList(TIndex).char.ShieldAnim, UserList(TIndex).char.CascoAnim, _
                        UserList(TIndex).char.Alas)
                Exit Sub

            Case "CRI"

                If TIndex <= 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                    Exit Sub

                End If
            
                UserList(TIndex).Faccion.CriminalesMatados = val(Arg2)
                Exit Sub

            Case "CIU"

                If TIndex <= 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
                    Exit Sub

                End If
            
                UserList(TIndex).Faccion.CiudadanosMatados = val(Arg2)
                Exit Sub

            Case "CLASE"

                If TIndex <= 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline: " & tStr & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(TIndex).Clase = UCase$(Left$(Arg2, 1)) & UCase$(mid$(Arg2, 2)) Then
                     Call SendData(SendTarget.toIndex, UserIndex, 0, "||La clase de: " & tStr & " no ha cambiado porque ya es: " & UCase$(Left$(Arg2, 1)) & UCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Len(Arg2) > 1 Then
                    UserList(TIndex).Clase = UCase$(Left$(Arg2, 1)) & UCase$(mid$(Arg2, 2))
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Clase cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                Else
                    UserList(TIndex).Clase = UCase$(Arg2)
                End If

            Case "RAZA"
                
                If TIndex <= 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline: " & tStr & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) = UserList(TIndex).Raza Then
                   Call SendData(SendTarget.toIndex, UserIndex, 0, "||La raza de: " & tStr & " ya no ha cambiado porque ya es: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                   Exit Sub
                End If
                
                Select Case UCase$(Arg2)
                 
                 Case "HUMANO"
                         Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                         UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                         Call DarCuerpoDesnudo(TIndex)
                     
                 Case "ENANO"
                          Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                          UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                          Call DarCuerpoDesnudo(TIndex)
                 
                 Case "HOBBIT"
                          Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                          UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                          Call DarCuerpoDesnudo(TIndex)
                
                Case "ELFO"
                          Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                          UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                          Call DarCuerpoDesnudo(TIndex)
                
                Case "ELFO OSCURO"
                          Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                          UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                          Call DarCuerpoDesnudo(TIndex)
                     
                Case "LICANTROPO"
                         Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                         UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                         Call DarCuerpoDesnudo(TIndex)
                
                Case "GNOMO"
                          Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                          UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                          Call DarCuerpoDesnudo(TIndex)
                
                Case "ORCO"
                         Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                         UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                         Call DarCuerpoDesnudo(TIndex)
                
                Case "VAMPIRO"
                         Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                         UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                         Call DarCuerpoDesnudo(TIndex)
                
                Case "CICLOPE"
                          Call SendData(SendTarget.toIndex, UserIndex, 0, "||Raza cambiada: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & FONTTYPE_INFO)
                          UserList(TIndex).Raza = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
                          Call DarCuerpoDesnudo(TIndex)
                     
                Case Else
                         Call SendData(SendTarget.toIndex, UserIndex, 0, "||La raza: " & UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2)) & " no existe." & FONTTYPE_INFO)
                End Select
                
            Case "SKILLS"

                For LoopC = 1 To NUMSKILLS
                    If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg2) Then n = LoopC
                Next LoopC

                If n = 0 Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Skill Inexistente!" & FONTTYPE_INFO)
                    Exit Sub
                End If

                If TIndex = 0 Then
                    Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "Skills", "SK" & n, Arg3)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
                Else
                    UserList(TIndex).Stats.UserSkills(n) = val(Arg3)
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has cambiado la skill de " & SkillsNames(n) & " a " & UserList(TIndex).Name & " por: " & val(Arg3) & FONTTYPE_INFO)
                    Call SendData(SendTarget.toIndex, TIndex, 0, "||GM " & UserList(UserIndex).Name & " te ha cambiado el valor de la skill " & SkillsNames(n) & " a: " & val(Arg3) & FONTTYPE_INFO)
                End If

                Exit Sub
        
            Case "SKILLSLIBRES"
               Dim SLName As String
               Dim SLSkills As Integer
               Dim SLResult As Integer
              
              If Left(Arg2, 1) = "-" Then
                      
                    If TIndex = 0 Then
                         SLName = ReadField(1, rData, 32)
                         
                         If Not FileExist(CharPath & UCase(SLName) & ".chr") Then
                             Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje: " & SLName & " no existe." & FONTTYPE_INFO)
                             Exit Sub
                         End If
                         
                         SLSkills = GetVar(CharPath & UCase$(SLName) & ".chr", "STATS", "SkillptsLibres")
    
                         SLResult = SLSkills - mid(Arg2, 2)
                       
                       If SLResult < 0 Then
                                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El cambio no se ha podido efectuar:" & SLResult & FONTTYPE_INFO)
                                Exit Sub
                       Else
                               Call WriteVar(CharPath & UCase$(SLName) & ".chr", "STATS", "SkillPtsLibres", SLResult)
                               Call SendData(SendTarget.toIndex, UserIndex, 0, "||El charfile de " & SLName & " ha cambiado su SkillsLibres de " & SLSkills & " a " & SLResult & FONTTYPE_INFO)
                               Exit Sub
                        End If
                         
                         Else
                         SLName = UserList(TIndex).Name
                         SLSkills = UserList(TIndex).Stats.SkillPts
                         
                         SLResult = SLSkills - mid(Arg2, 2)
                         
                         If SLResult < 0 Then
                             Call SendData(SendTarget.toIndex, UserIndex, 0, "||El cambio no se ha podido efectuar:" & SLResult & FONTTYPE_INFO)
                             Exit Sub
                         Else
                             UserList(TIndex).Stats.SkillPts = SLResult
                             Call SendData(SendTarget.toIndex, UserIndex, 0, "||Las skills de " & SLName & " han sido modificadas." & FONTTYPE_INFO)
                             Call EnviarSkills(TIndex)
                             Exit Sub
                         End If
                    End If
                      
               
                  Else 'Parte donde Suma
                  
                  If TIndex = 0 Then
                     
                         SLName = ReadField(1, rData, 32)
                         
                         If Not FileExist(CharPath & UCase(SLName) & ".chr") Then
                             Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje: " & SLName & " no existe." & FONTTYPE_INFO)
                             Exit Sub
                         End If
                         
                         SLSkills = GetVar(CharPath & UCase$(SLName) & ".chr", "STATS", "SkillptsLibres")
    
                         SLResult = SLSkills + Arg2
                         
                         Call WriteVar(CharPath & UCase$(SLName) & ".chr", "STATS", "SkillPtsLibres", SLResult)
                         Call SendData(SendTarget.toIndex, UserIndex, 0, "||El charfile de " & SLName & " ha cambiado su SkillsLibres de " & SLSkills & " a " & SLResult & FONTTYPE_INFO)
                         Exit Sub
                         
                         Else
                         SLName = UserList(TIndex).Name
                         SLSkills = UserList(TIndex).Stats.SkillPts
                         SLResult = SLSkills + Arg2
                         
                         UserList(TIndex).Stats.SkillPts = SLResult
                         Call SendData(SendTarget.toIndex, UserIndex, 0, "||Las skills de " & SLName & " han sido modificadas." & FONTTYPE_INFO)
                         Call EnviarSkills(TIndex)
                         Exit Sub
                         
                End If
                  
              End If

                Exit Sub
                
            Case Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Sintaxis incorrecto" & FONTTYPE_GUILD)
                Call SendData(SendTarget.toIndex, UserIndex, 0, _
                        "||Comando: /MOD <Nick/yo> <NIVEL/SKILLS/SKILLSLIBRES/ORO/CIU/CRI/EXP/BODY/HEAD> <VALOR>" & FONTTYPE_GUILD)
                Exit Sub

        End Select

        Exit Sub

    End If

    '<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
    '<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
    '<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
    If UserList(UserIndex).flags.Privilegios < PlayerType.SemiDios Then
        Exit Sub

    End If

    If UCase$(Left$(rData, 6)) = "/INFO " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
    
        rData = Right$(rData, Len(rData) - 6)
    
        TIndex = NameIndex(rData)
    
        If TIndex <= 0 Then
       
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline, Buscando en Charfile." & FONTTYPE_INFO)
            SendUserStatsTxtOFF UserIndex, rData
        Else

            If UserList(TIndex).flags.Privilegios >= PlayerType.Dios Then Exit Sub
            SendUserStatsTxt UserIndex, TIndex

        End If

        Exit Sub

    End If

    'MINISTATS DEL USER
    If UCase$(Left$(rData, 6)) = "/STAT " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        
        rData = Right$(rData, Len(rData) - 6)
        
        TIndex = NameIndex(rData)
        
        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline. Leyendo Charfile... " & FONTTYPE_INFO)
            SendUserMiniStatsTxtFromChar UserIndex, rData
        Else
            SendUserMiniStatsTxt UserIndex, TIndex

        End If
    
        Exit Sub

    End If

    If UCase$(Left$(rData, 5)) = "/BAL " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 5)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline. Leyendo charfile... " & FONTTYPE_TALK)
            SendUserOROTxtFromChar UserIndex, rData
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El usuario " & rData & " tiene " & UserList(TIndex).Stats.Banco & " en el banco" & _
                    FONTTYPE_TALK)

        End If

        Exit Sub

    End If

    'INV DEL USER
    If UCase$(Left$(rData, 5)) = "/INV " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
    
        rData = Right$(rData, Len(rData) - 5)
    
        TIndex = NameIndex(rData)
    
        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline. Leyendo del charfile..." & FONTTYPE_TALK)
            SendUserInvTxtFromChar UserIndex, rData
        Else
            SendUserInvTxt UserIndex, TIndex

        End If

        Exit Sub

    End If

    'INV DEL USER
    If UCase$(Left$(rData, 5)) = "/BOV " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
    
        rData = Right$(rData, Len(rData) - 5)
    
        TIndex = NameIndex(rData)
    
        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline. Leyendo charfile... " & FONTTYPE_TALK)
            SendUserBovedaTxtFromChar UserIndex, rData
        Else
            SendUserBovedaTxt UserIndex, TIndex

        End If

        Exit Sub

    End If

    'SKILLS DEL USER
    If UCase$(Left$(rData, 8)) = "/SKILLS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
    
        rData = Right$(rData, Len(rData) - 8)
    
        TIndex = NameIndex(rData)
    
        If TIndex <= 0 Then
            Call Replace(rData, "\", " ")
            Call Replace(rData, "/", " ")
        
            For tInt = 1 To NUMSKILLS
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| CHAR>" & SkillsNames(tInt) & " = " & GetVar(CharPath & rData & ".chr", _
                        "SKILLS", "SK" & tInt) & FONTTYPE_INFO)
            Next tInt

            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| CHAR> Libres:" & GetVar(CharPath & rData & ".chr", "STATS", "SKILLPTSLIBRES") & _
                    FONTTYPE_INFO)
            Exit Sub

        End If

        SendUserSkillsTxt UserIndex, TIndex
        Exit Sub

    End If

    If UCase$(Left$(rData, 9)) = "/REVIVIR " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 9)
        Name = rData

        If UCase$(Name) <> "YO" Then
            TIndex = NameIndex(Name)
        Else
            TIndex = UserIndex

        End If

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(TIndex).flags.Muerto = 0 Then
           Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario esta vivo." & FONTTYPE_INFO)
            Exit Sub
        End If

        UserList(TIndex).flags.Muerto = 0
        UserList(TIndex).Stats.MinHP = UserList(TIndex).Stats.MaxHP
        Call DarCuerpoDesnudo(TIndex)
   
        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(TIndex).pos.Map, val(TIndex), UserList(TIndex).char.Body, UserList(TIndex).OrigChar.Head, _
                UserList(TIndex).char.Heading, UserList(TIndex).char.WeaponAnim, UserList(TIndex).char.ShieldAnim, UserList(TIndex).char.CascoAnim, _
                UserList(TIndex).char.Alas)
            
        Call SendUserStatsBox(val(TIndex))
        Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(UserIndex).Name & " te ha resucitado." & FONTTYPE_INFO)
       
        Exit Sub

    End If

    If UCase$(rData) = "/ONLINEGM" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        For LoopC = 1 To LastUser

            'Tiene nombre? Es GM? Si es Dios o Admin, nosotros lo somos también??
            If (UserList(LoopC).Name <> "") And UserList(LoopC).flags.Privilegios > PlayerType.User And (UserList(LoopC).flags.Privilegios < _
                    PlayerType.Dios Or UserList(UserIndex).flags.Privilegios >= PlayerType.Dios) Then
                tStr = tStr & UserList(LoopC).Name & ", "

            End If

        Next LoopC

        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No hay GMs Online" & FONTTYPE_INFO)

        End If

        Exit Sub

    End If

    If UCase$(rData) = "/ONLINEMAP" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        For LoopC = 1 To LastUser

            If UserList(LoopC).Name <> "" And UserList(LoopC).pos.Map = UserList(UserIndex).pos.Map And (UserList(LoopC).flags.Privilegios < _
                    PlayerType.Dios Or UserList(UserIndex).flags.Privilegios >= PlayerType.Dios) Then
                tStr = tStr & UserList(LoopC).Name & ", "

            End If

        Next LoopC

        If Len(tStr) > 2 Then tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuarios en el mapa: " & tStr & FONTTYPE_INFO)
        Exit Sub

    End If

   If UCase$(Left$(rData, 7)) = "/PERDON" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 8)
        TIndex = NameIndex(rData)

        Call VolverCiudadano(TIndex)
        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "/ECHAR " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 7)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
            Exit Sub

        End If
    
        If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes echar a alguien con jerarquia mayor a la tuya." & FONTTYPE_INFO)
            Exit Sub

        End If
        
        Call SendData(SendTarget.toall, 0, 0, "||" & UserList(UserIndex).Name & " echo a " & UserList(TIndex).Name & "." & FONTTYPE_INFO)
        Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(TIndex).Name)
        Call CloseSocket(TIndex)
        Exit Sub

    End If

    If UCase$(Left$(rData, 4)) = "/BAN " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 4)
        tStr = ReadField(2, rData, Asc("@")) ' NICK
        TIndex = NameIndex(tStr)
        Name = ReadField(1, rData, Asc("@")) ' MOTIVO
        
        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_TALK)
        
            If FileExist(CharPath & tStr & ".chr", vbNormal) Then
                tLong = UserDarPrivilegioLevel(tStr)
            
                If tLong > UserList(UserIndex).flags.Privilegios Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estás loco??! No podés banear a alguien de mayor jerarquia que vos!" & _
                            FONTTYPE_INFO)
                    Exit Sub

                End If
            
                If GetVar(CharPath & tStr & ".chr", "FLAGS", "Ban") <> "0" Then
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "||El personaje ya ha sido baneado anteriormente." & FONTTYPE_INFO)
                    Exit Sub

                End If
            
                Call LogBanFromName(tStr, UserIndex, Name)
                Call SendData(SendTarget.ToAdmins, 0, 0, "||AOMania> El GM & " & UserList(UserIndex).Name & "baneó a " & tStr & "." & FONTTYPE_SERVER)
            
                'ponemos el flag de ban a 1
                Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
                'ponemos la pena
                tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
                Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
                Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & _
                        " Lo Baneó por el siguiente motivo: " & LCase$(Name) & " " & Date & " " & Time)
            
                If tLong > 0 Then
                    UserList(UserIndex).flags.Ban = 1
                    Call CloseSocket(UserIndex)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||" & " El gm " & UserList(UserIndex).Name & _
                            " fue baneado por el propio servidor por intentar banear a otro admin." & FONTTYPE_FIGHT)

                End If

                Call LogGM(UserList(UserIndex).Name, "BAN a " & tStr)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El pj " & tStr & " no existe." & FONTTYPE_INFO)

            End If

        Else

            If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes banear a al alguien de mayor jerarquia." & FONTTYPE_INFO)
                Exit Sub

            End If
        
            Call LogBan(TIndex, UserIndex, Name)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||AOMania> " & UserList(UserIndex).Name & " ha baneado a " & UserList(TIndex).Name & "." & _
                    FONTTYPE_SERVER)
        
            'Ponemos el flag de ban a 1
            UserList(TIndex).flags.Ban = 1
        
            If UserList(TIndex).flags.Privilegios > PlayerType.User Then
                UserList(UserIndex).flags.Ban = 1
                Call CloseSocket(UserIndex)
                Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " banned by the server por bannear un Administrador." & _
                        FONTTYPE_FIGHT)

            End If
        
            Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(TIndex).Name)
        
            'ponemos el flag de ban a 1
            Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
            'ponemos la pena
            tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & " Lo Baneó Debido a: " & LCase$( _
                    Name) & " " & Date & " " & Time)
        
            Call CloseSocket(TIndex)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "/UNBAN " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 7)
    
        rData = Replace(rData, "\", "")
        rData = Replace(rData, "/", "")
    
        If Not FileExist(CharPath & rData & ".chr", vbNormal) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Charfile inexistente (no use +)" & FONTTYPE_INFO)
            Exit Sub

        End If
    
        Call UnBan(rData)
    
        'penas
        i = val(GetVar(CharPath & rData & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & rData & ".chr", "PENAS", "Cant", i + 1)
        Call WriteVar(CharPath & rData & ".chr", "PENAS", "P" & i + 1, LCase$(UserList(UserIndex).Name) & " Lo unbaneó. " & Date & " " & Time)
    
        Call LogGM(UserList(UserIndex).Name, "/UNBAN a " & rData)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & rData & " unbanned." & FONTTYPE_INFO)

        Exit Sub

    End If

    'SEGUIR
    If UCase$(rData) = "/SEGUIR" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.TargetNpc > 0 Then
            Call DoFollow(UserList(UserIndex).flags.TargetNpc, UserList(UserIndex).Name)

        End If

        Exit Sub

    End If

    If UCase(rData) = "/BLOQ" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Blocked = 0 Then
            MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Blocked = 1
            Call Bloquear(SendTarget.ToMap, UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, _
                    UserList(UserIndex).pos.Y, 1)
        Else
            MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Blocked = 0
            Call Bloquear(SendTarget.ToMap, UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, _
                    UserList(UserIndex).pos.Y, 0)

        End If

        Exit Sub

    End If

    If UCase$(rData) = "/ACTCOM" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If ComerciarAc = True Then
            ComerciarAc = False
            Call SendData(SendTarget.toall, 0, 0, "||Comercio entre usuarios desactivados!!." & FONTTYPE_CYAN)
        Else
            Call SendData(SendTarget.toall, 0, 0, "||Comercio entre usuarios activados!!." & FONTTYPE_CYAN)
            ComerciarAc = True

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 8)) = "/AOMANIA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 8)

        Call GuerraBanda.Ban_Comienza("32")

    End If

    'Crear criatura
    If UCase$(Left$(rData, 3)) = "/CC" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Call EnviarSpawnList(UserIndex)
        Exit Sub

    End If

    'Spawn!!!!! ¿What?
    If UCase$(Left$(rData, 3)) = "SPA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 3)
    
        If val(rData) > 0 And val(rData) < UBound(SpawnList) + 1 Then Call SpawnNpc(SpawnList(val(rData)).NpcIndex, UserList(UserIndex).pos, True, _
                False)
        Call LogGM(UserList(UserIndex).Name, "Sumoneo " & SpawnList(val(rData)).NpcName)
          
        Exit Sub

    End If

    'Resetea el inventario
    If UCase$(rData) = "/RESETINV" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 9)

        If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
        Call ResetNpcInv(UserList(UserIndex).flags.TargetNpc)
        Exit Sub

    End If

    '/Clean
    If UCase$(rData) = "/LIMPIAR" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call LimpiarMundo
        Exit Sub

    End If

    'Ip del nick
    If UCase$(Left$(rData, 9)) = "/NICK2IP " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 9)
        TIndex = NameIndex(UCase$(rData))
        Call LogGM(UserList(UserIndex).Name, "NICK2IP Solicito la IP de " & rData)

        If TIndex > 0 Then
            If (UserList(UserIndex).flags.Privilegios > PlayerType.User And UserList(TIndex).flags.Privilegios = PlayerType.User) Or (UserList( _
                    UserIndex).flags.Privilegios >= PlayerType.Dios) Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El ip de " & rData & " es " & UserList(TIndex).ip & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No tienes los privilegios necesarios" & FONTTYPE_INFO)

            End If

        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No hay ningun personaje con ese nick" & FONTTYPE_INFO)

        End If

        Exit Sub

    End If
 
    'Ip del nick
    If UCase$(Left$(rData, 9)) = "/IP2NICK " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 9)

        If InStr(rData, ".") < 1 Then
            tInt = NameIndex(rData)

            If tInt < 1 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Pj Offline" & FONTTYPE_INFO)
                Exit Sub

            End If

            rData = UserList(tInt).ip

        End If

        tStr = vbNullString
        Call LogGM(UserList(UserIndex).Name, "IP2NICK Solicito los Nicks de IP " & rData)

        For LoopC = 1 To LastUser

            If UserList(LoopC).ip = rData And UserList(LoopC).Name <> "" And UserList(LoopC).flags.UserLogged Then

                If (UserList(UserIndex).flags.Privilegios > PlayerType.User And UserList(LoopC).flags.Privilegios = PlayerType.User) Or (UserList( _
                        UserIndex).flags.Privilegios >= PlayerType.Dios) Then
                    tStr = tStr & UserList(LoopC).Name & ", "

                End If

            End If

        Next LoopC
    
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los personajes con ip " & rData & " son: " & tStr & FONTTYPE_INFO)
        Exit Sub

    End If

    If UCase$(Left$(rData, 8)) = "/ONCLAN " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 8)
        tInt = GuildIndex(rData)
    
        If tInt > 0 Then
            tStr = modGuilds.m_ListaDeMiembrosOnline(UserIndex, tInt)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Clan " & UCase(rData) & ": " & tStr & FONTTYPE_GUILDMSG)

        End If

    End If

    'Crear Teleport
    If UCase(Left(rData, 5)) = "/CTP " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        '/ct mapa_dest x_dest y_dest
        rData = Right(rData, Len(rData) - 5)
        Mapa = ReadField(1, rData, 32)
        X = ReadField(2, rData, 32)
        Y = ReadField(3, rData, 32)
    
        If MapaValido(Mapa) = False Or InMapBounds(Mapa, X, Y) = False Then
            Exit Sub

        End If

        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).OBJInfo.ObjIndex > 0 Then
            Exit Sub

        End If

        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.Map > 0 Then
            Exit Sub

        End If
    
        If MapData(Mapa, X, Y).OBJInfo.ObjIndex > 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, Mapa, "||Hay un objeto en el piso en ese lugar" & FONTTYPE_INFO)
            Exit Sub

        End If
    
        Dim ET As Obj
        ET.Amount = 1
        ET.ObjIndex = 378
    
        Call MakeObj(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, ET, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList( _
                UserIndex).pos.Y - 1)
    
        MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.Map = Mapa
        MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.X = X
        MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.Y = Y
    
        Exit Sub

    End If

    'Destruir Teleport
    'toma el ultimo click
    If UCase(Left(rData, 4)) = "/DTP" Then
        '/dt
   
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
    
        Mapa = UserList(UserIndex).flags.TargetMap
        X = UserList(UserIndex).flags.TargetX
        Y = UserList(UserIndex).flags.TargetY
    
        If ObjData(MapData(Mapa, X, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otTELEPORT And MapData(Mapa, X, Y).TileExit.Map > 0 Then
            Call EraseObj(SendTarget.ToMap, 0, Mapa, MapData(Mapa, X, Y).OBJInfo.Amount, Mapa, X, Y)
            Call EraseObj(SendTarget.ToMap, 0, MapData(Mapa, X, Y).TileExit.Map, 1, MapData(Mapa, X, Y).TileExit.Map, MapData(Mapa, X, _
                    Y).TileExit.X, MapData(Mapa, X, Y).TileExit.Y)
            MapData(Mapa, X, Y).TileExit.Map = 0
            MapData(Mapa, X, Y).TileExit.X = 0
            MapData(Mapa, X, Y).TileExit.Y = 0

        End If
    
        Exit Sub

    End If

    If UCase$(rData) = "/LLUVIA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Call SecondaryAmbient
        Exit Sub

    End If

    Select Case UCase$(Left$(rData, 13))
      
        Case "/FORCEMIDIMAP"
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If Len(rData) > 13 Then
                rData = Right$(rData, Len(rData) - 14)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, _
                        "||El formato correcto de este comando es /FORCEMIDMAP MIDI MAPA, siendo el MAPA opcional" & FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Solo dioses, admins y RMS
            If UserList(UserIndex).flags.Privilegios < PlayerType.Dios And Not UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
        
            'Obtenemos el número de midi
            Arg1 = ReadField(1, rData, vbKeySpace)
            ' y el de mapa
            Arg2 = ReadField(2, rData, vbKeySpace)
        
            'Si el mapa no fue enviado tomo el actual
            If IsNumeric(Arg2) Then
                tInt = CInt(Arg2)
            Else
                tInt = UserList(UserIndex).pos.Map

            End If
        
            If IsNumeric(Arg1) Then
                If Arg1 = "0" Then
                    'Ponemos el default del mapa
                    Call SendData(SendTarget.ToMap, 0, tInt, "TM" & CStr(MapInfo(UserList(UserIndex).pos.Map).Music))
                Else
                    'Ponemos el pedido por el GM
                    Call SendData(SendTarget.ToMap, 0, tInt, "TM" & Arg1)

                End If

            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, _
                        "||El formato correcto de este comando es /FORCEMIDMAP MIDI MAPA, siendo el MAPA opcional" & FONTTYPE_INFO)

            End If

            Exit Sub
    
        Case "/FORCEWAVMAP "
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
            rData = Right$(rData, Len(rData) - 13)

            'Solo dioses, admins y RMS
            If UserList(UserIndex).flags.Privilegios < PlayerType.Dios And Not UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
        
            'Obtenemos el número de wav
            Arg1 = ReadField(1, rData, vbKeySpace)
            ' el de mapa
            Arg2 = ReadField(2, rData, vbKeySpace)
            ' el de X
            Arg3 = ReadField(3, rData, vbKeySpace)
            ' y el de Y (las coords X-Y sólo tendrán sentido al implementarse el panning en la 11.6)
            Arg4 = ReadField(4, rData, vbKeySpace)
        
            If IsNumeric(Arg2) And IsNumeric(Arg3) And IsNumeric(Arg4) Then
                tInt = CInt(Arg2)
            Else
                tInt = UserList(UserIndex).pos.Map
                Arg3 = CStr(UserList(UserIndex).pos.X)
                Arg4 = CStr(UserList(UserIndex).pos.Y)

            End If
        
            If IsNumeric(Arg1) Then
                Call SendData(SendTarget.ToMap, 0, tInt, "TW" & Arg1)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, _
                        "||El formato correcto de este comando es /FORCEWAVMAP WAV MAPA X Y, siendo la posición opcional" & FONTTYPE_INFO)

            End If

            Exit Sub

    End Select

    Select Case UCase$(Left$(rData, 8))
    
        Case "/TALKAS "
            Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

            If UserList(UserIndex).flags.Privilegios > PlayerType.SemiDios Or UserList(UserIndex).flags.EsRolesMaster Then

                If UserList(UserIndex).flags.TargetNpc > 0 Then
                    tStr = Right$(rData, Len(rData) - 8)
                
                    Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNpc, Npclist(UserList(UserIndex).flags.TargetNpc).pos.Map, _
                            "||" & vbWhite & "°" & tStr & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNpc).char.CharIndex))
                Else
                    Call SendData(SendTarget.toIndex, UserIndex, 0, _
                            "||Debes seleccionar el NPC por el que quieres hablar antes de usar este comando" & FONTTYPE_INFO)

                End If

            End If

            Exit Sub

    End Select

    '<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
    '<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
    '<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
    If UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then
        Exit Sub

    End If

    If UCase$(rData) = "/MASSEJECUTAR" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        For LoopC = 1 To LastUser

            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.UserLogged Then
                    If Not UserList(LoopC).flags.Privilegios >= 1 Then
                        If UserList(LoopC).pos.Map = UserList(UserIndex).pos.Map Then
                            Call UserDie(LoopC)
                       End If

                    End If

                End If

            End If

        Next LoopC

        Exit Sub

    End If

    '[yb]
    If UCase$(Left$(rData, 8)) = "/MASSORO" Then

        With UserList(UserIndex)
                    
            Call LogGM(.Name, "Comando: " & rData)

            For Y = .pos.Y - MinYBorder + 1 To .pos.Y + MinYBorder - 1
                For X = .pos.X - MinXBorder + 1 To .pos.X + MinXBorder - 1

                    If InMapBounds(.pos.Map, X, Y) Then
                        If MapData(.pos.Map, X, Y).OBJInfo.ObjIndex = iORO Then
                            Call EraseObj(SendTarget.ToMap, 0, .pos.Map, 10000, .pos.Map, X, Y)

                        End If

                    End If

                Next X
            Next Y

        End With

        Exit Sub
       
    End If

    If UCase$(Left$(rData, 6)) = "/PASS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 6)
        TIndex = NameIndex(rData)

        If Not FileExist(CharPath & rData & ".chr") Then Exit Sub
        Arg1 = GetVar(CharPath & rData & ".chr", "INIT", "Password")
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||la pass de " & rData & " es " & Arg1 & FONTTYPE_INFO)
        Exit Sub

    End If

    If UCase$(Left$(rData, 12)) = "/ACEPTCONSE " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 12)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline" & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toall, 0, 0, "||" & rData & " Ha sido coronado como el nuevo Rey Imperial." & FONTTYPE_CONSEJO)
            UserList(TIndex).flags.PertAlCons = 1
            Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y, False)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 16)) = "/ACEPTCONSECAOS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 16)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline" & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toall, 0, 0, "||" & rData & " Ha sido coronado como el nuevo Rey del Caos." & FONTTYPE_CONSEJOCAOS)
            UserList(TIndex).flags.PertAlConsCaos = 1
            Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y, False)

        End If

        Exit Sub

    End If

    If Left$(UCase$(rData), 13) = "/DUMPSECURITY" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Call SecurityIp.DumpTables
        Exit Sub

    End If

    If UCase$(Left$(rData, 11)) = "/KICKCONSE " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 11)
        TIndex = NameIndex(rData)

        If TIndex <= 0 Then
            If FileExist(CharPath & rData & ".chr") Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline, Echando de los consejos" & FONTTYPE_INFO)
                Call WriteVar(CharPath & UCase(rData) & ".chr", "CONSEJO", "PERTENECE", 0)
                Call WriteVar(CharPath & UCase(rData) & ".chr", "CONSEJO", "PERTENECECAOS", 0)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No se encuentra el charfile " & CharPath & rData & ".chr" & FONTTYPE_INFO)
                Exit Sub

            End If

        Else

            If UserList(TIndex).flags.PertAlCons > 0 Then
                Call SendData(SendTarget.toIndex, TIndex, 0, "||Has sido echado en el consejo de banderbill" & FONTTYPE_TALK & ENDC)
                UserList(TIndex).flags.PertAlCons = 0
                Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y)
                Call SendData(SendTarget.toall, 0, 0, "||" & rData & " fue expulsado del consejo de Banderbill" & FONTTYPE_CONSEJO)

            End If

            If UserList(TIndex).flags.PertAlConsCaos > 0 Then
                Call SendData(SendTarget.toIndex, TIndex, 0, "||Has sido echado en el consejo de la legión oscura" & FONTTYPE_TALK & ENDC)
                UserList(TIndex).flags.PertAlConsCaos = 0
                Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y)
                Call SendData(SendTarget.toall, 0, 0, "||" & rData & " fue expulsado del consejo de la Legión Oscura" & FONTTYPE_CONSEJOCAOS)

            End If

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 8)) = "/TRIGGER" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
    
        rData = Trim(Right(rData, Len(rData) - 8))
        Mapa = UserList(UserIndex).pos.Map
        X = UserList(UserIndex).pos.X
        Y = UserList(UserIndex).pos.Y

        If rData <> "" Then
            tInt = MapData(Mapa, X, Y).trigger
            MapData(Mapa, X, Y).trigger = val(rData)

        End If

        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Trigger " & MapData(Mapa, X, Y).trigger & " en mapa " & Mapa & " " & X & ", " & Y & _
                FONTTYPE_INFO)
        Exit Sub

    End If

    If UCase(rData) = "/BANIPLIST" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        tStr = "||"

        For LoopC = 1 To BanIps.Count
            tStr = tStr & BanIps.Item(LoopC) & ", "
        Next LoopC

        tStr = tStr & FONTTYPE_INFO
        Call SendData(SendTarget.toIndex, UserIndex, 0, tStr)
        Exit Sub

    End If

    If UCase(rData) = "/BANIPRELOAD" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call BanIpGuardar
        Call BanIpCargar
        Exit Sub

    End If

    If UCase(Left(rData, 14)) = "/MIEMBROSCLAN " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Trim(Right(rData, Len(rData) - 9))

        If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
            Exit Sub

        End If
    
        Call LogGM(UserList(UserIndex).Name, "MIEMBROSCLAN a " & rData)

        tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
    
        For i = 1 To tInt
            tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
  
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||" & tStr & "<" & rData & ">." & FONTTYPE_INFO)
        Next i

        Exit Sub

    End If

    If UCase(Left(rData, 9)) = "/BANCLAN " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Trim(Right(rData, Len(rData) - 9))

        If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
            Exit Sub

        End If
    
        Call SendData(SendTarget.toall, 0, 0, "|| " & UserList(UserIndex).Name & " banned al clan " & UCase$(rData) & FONTTYPE_FIGHT)
    
        Call LogGM(UserList(UserIndex).Name, "BANCLAN a " & rData)

        tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
    
        For i = 1 To tInt
            tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
            'tstr es la victima
            Call Ban(tStr, "Administracion del servidor", "Clan Banned")
            TIndex = NameIndex(tStr)

            If TIndex > 0 Then
    
                UserList(TIndex).flags.Ban = 1
                Call CloseSocket(TIndex)

            End If
        
            Call SendData(SendTarget.toall, 0, 0, "||   " & tStr & "<" & rData & "> ha sido expulsado del servidor." & FONTTYPE_FIGHT)

            Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")

            n = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", n + 1)
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & n + 1, LCase$(UserList(UserIndex).Name) & ": BAN AL CLAN: " & rData & " " & Date _
                    & " " & Time)

        Next i

        Exit Sub

    End If

    'Ban x IP
    If UCase(Left(rData, 9)) = "/BANLAIP " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Dim BanIP As String, XNick As Boolean
    
        rData = Right$(rData, Len(rData) - 9)
        tStr = Replace(ReadField(1, rData, Asc(" ")), "+", " ")
 
        TIndex = NameIndex(tStr)

        If TIndex <= 0 Then
            XNick = False
            BanIP = tStr
        Else
            XNick = True
            Call LogGM(UserList(UserIndex).Name, "/BANLAIP " & UserList(TIndex).Name & " - " & UserList(TIndex).ip)
            BanIP = UserList(TIndex).ip

        End If
    
        rData = Right$(rData, Len(rData) - Len(tStr))
    
        If BanIpBuscar(BanIP) > 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La IP " & BanIP & " ya se encuentra en la lista de bans." & FONTTYPE_INFO)
            Exit Sub

        End If
    
        Call BanIpAgrega(BanIP)
        Call SendData(SendTarget.ToAdmins, UserIndex, 0, "||" & UserList(UserIndex).Name & " Baneo la IP " & BanIP & FONTTYPE_FIGHT)
    
        If XNick = True Then
            Call LogBan(TIndex, UserIndex, "Ban por IP desde Nick por " & rData)
        
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " echo a " & UserList(TIndex).Name & "." & FONTTYPE_FIGHT)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " Banned a " & UserList(TIndex).Name & "." & FONTTYPE_FIGHT)
        
            UserList(TIndex).flags.Ban = 1
        
            Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(TIndex).Name)
            Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(TIndex).Name)
            Call CloseSocket(TIndex)

        End If
    
        Exit Sub

    End If

    If UCase(Left(rData, 9)) = "/UNBANIP " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
    
        rData = Right(rData, Len(rData) - 9)
    
        If BanIpQuita(rData) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La IP """ & rData & """ se ha quitado de la lista de bans." & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La IP """ & rData & """ NO se encuentra en la lista de bans." & FONTTYPE_INFO)

        End If
    
        Exit Sub

    End If
     
    If UCase(Left(rData, 6)) = "/UMSG " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 6)
      
        tName = ReadField(1, rData, 32)
        tMessage = Right$(rData, Len(rData) - (1 + Len(tName)))
        TIndex = NameIndex(tName)
    
        If TIndex <= 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub

        End If
    
        Call SendData(SendTarget.toIndex, TIndex, 0, "||< " & UserList(UserIndex).Name & " > te dice: " & tMessage & FONTTYPE_SERVER)
  
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Le has mandado a " & tName & " : " & tMessage & FONTTYPE_SERVER)

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
                Call SendData(SendTarget.toIndex, UserIndex, 0, "VCTS" & CID & "#" & CNameItem)
            Else

                If InStr(LCase(CNameItem), LCase(rData)) Then
            
                    Asi = Asi + 1
                    Call SendData(SendTarget.toIndex, UserIndex, 0, "VCTS" & CID & "#" & CNameItem)

                End If
            
            End If

        Next xci
        
        If Asi = 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "VITS" & Asi)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "VITS" & Asi)

        End If

        Exit Sub

    End If
    
    'Crear Item
    If UCase(Left(rData, 3)) = "/CI" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        Dim txt      As String
        Dim Cadena() As String
        Dim IdItem   As String
        Dim Cantidad As String
          
        txt = rData
          
        Cadena = Split(txt, Chr$(32))
          
        If txt = "/CI" Or UBound(Cadena) < 2 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis incorrecto." & FONTTYPE_GUILD)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis: /CI <ID ITEM> <CANTIDAD>" & FONTTYPE_GUILD)
            Exit Sub

        End If
          
        If Not IsNumeric(Cadena(1)) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El ID Item debe ser númerica." & FONTTYPE_GUILD)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis: /CI <ID ITEM> <CANTIDAD> Ejemplo: /CI 2 <CANTIDAD>." & FONTTYPE_GUILD)
            Exit Sub

        End If
          
        If Not IsNumeric(Cadena(2)) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| El Cantidad debe ser numérica." & FONTTYPE_GUILD)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Sintaxis: /CI <ID ITEM> <CANTIDAD> Ejemplo: /CI 2 10." & FONTTYPE_GUILD)
            Exit Sub

        End If
          
        If Cadena(2) > 1200 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Has superado el tope de cantidad. (Max: 1200)" & FONTTYPE_GUILD)
            Exit Sub

        End If

        IdItem = Cadena(1)
        Cantidad = Cadena(2)
    
        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).OBJInfo.ObjIndex > 0 Then
            Exit Sub

        End If

        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.Map > 0 Then
            Exit Sub

        End If
         
        If val(IdItem) < 1 Or val(IdItem) > NumObjDatas Then
            Exit Sub

        End If
    
        'Is the object not null?
        If ObjData(val(IdItem)).Name = "" Then Exit Sub
    
        Dim Objeto As Obj
        
        Objeto.Amount = val(Cantidad)
        Objeto.ObjIndex = val(IdItem)
        
        Call MeterItemEnInventario(UserIndex, Objeto)
        
        Call LogGM(UserList(UserIndex).Name, "Creo: " & Cantidad & " " & ObjData(Objeto.ObjIndex).Name)

        Exit Sub

    End If
        
    If UCase$(Left$(rData, 15)) = "/CHAUTEMPLARIO " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 15)
        Call LogGM(UserList(UserIndex).Name, "ECHO DEL TEMPLARIO A: " & rData)

        TIndex = NameIndex(rData)
        Dim tArmIndex As Integer
            
        If TIndex > 0 Then
            UserList(TIndex).Faccion.Templario = 0
            UserList(TIndex).Faccion.Reenlistadas = 200
            UserList(TIndex).Faccion.RecibioArmaduraTemplaria = 0
                
            Call PerderItemsFaccionarios(UserIndex, UserList(TIndex).Faccion.ArmaduraFaccionaria)
            
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas TEMPLARIAS y prohibida la reenlistada" & _
                    FONTTYPE_INFO)
                        
            Call SendData(SendTarget.toIndex, TIndex, 0, "|| " & UserList(UserIndex).Name & _
                    " te ha expulsado en forma definitiva de las fuerzas TEMPLARIAS." & FONTTYPE_FIGHT)
        Else

            If FileExist(CharPath & rData & ".chr") Then
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Templario", 0)
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(UserIndex).Name)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas TEMPLARIAS y prohibida la reenlistada" & _
                        FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & ".chr inexistente." & FONTTYPE_INFO)

            End If

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 13)) = "/CHAUNEMESIS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 13)
        Call LogGM(UserList(UserIndex).Name, "ECHO DEL NEMESIS A: " & rData)

        TIndex = NameIndex(rData)
    
        If TIndex > 0 Then
            UserList(TIndex).Faccion.Nemesis = 0
            UserList(TIndex).Faccion.Reenlistadas = 200
            UserList(TIndex).Faccion.RecibioArmaduraNemesis = 0
                
            Call PerderItemsFaccionarios(UserIndex, UserList(TIndex).Faccion.ArmaduraFaccionaria)
                
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas NEMESIS y prohibida la reenlistada" & _
                    FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, TIndex, 0, "|| " & UserList(UserIndex).Name & _
                    " te ha expulsado en forma definitiva de las fuerzas NEMESIS." & FONTTYPE_FIGHT)
        Else

            If FileExist(CharPath & rData & ".chr") Then
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Nemesis", 0)
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(UserIndex).Name)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas NEMESIS y prohibida la reenlistada" & _
                        FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & ".chr inexistente." & FONTTYPE_INFO)

            End If

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 10)) = "/CHAUCAOS " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 10)
        Call LogGM(UserList(UserIndex).Name, "ECHO DEL CAOS A: " & rData)

        TIndex = NameIndex(rData)
    
        If TIndex > 0 Then
            UserList(TIndex).Faccion.FuerzasCaos = 0
            UserList(TIndex).Faccion.Reenlistadas = 200
            UserList(TIndex).Faccion.RecibioArmaduraCaos = 0
                
            Call PerderItemsFaccionarios(UserIndex, UserList(TIndex).Faccion.ArmaduraFaccionaria)
                
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas del caos y prohibida la reenlistada" & _
                    FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, TIndex, 0, "|| " & UserList(UserIndex).Name & _
                    " te ha expulsado en forma definitiva de las fuerzas del caos." & FONTTYPE_FIGHT)
        Else

            If FileExist(CharPath & rData & ".chr") Then
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "EjercitoCaos", 0)
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(UserIndex).Name)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas del caos y prohibida la reenlistada" & _
                        FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & ".chr inexistente." & FONTTYPE_INFO)

            End If

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 10)) = "/CHAUREAL " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 10)
        Call LogGM(UserList(UserIndex).Name, "ECHO DE LA REAL A: " & rData)

        rData = Replace(rData, "\", "")
        rData = Replace(rData, "/", "")

        TIndex = NameIndex(rData)

        If TIndex > 0 Then
            UserList(TIndex).Faccion.ArmadaReal = 0
            UserList(TIndex).Faccion.Reenlistadas = 200
            UserList(TIndex).Faccion.RecibioArmaduraReal = 0
                
            Call PerderItemsFaccionarios(UserIndex, UserList(TIndex).Faccion.ArmaduraFaccionaria)
                
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas reales y prohibida la reenlistada" & _
                    FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, TIndex, 0, "|| " & UserList(UserIndex).Name & _
                    " te ha expulsado en forma definitiva de las fuerzas reales." & FONTTYPE_FIGHT)
        Else

            If FileExist(CharPath & rData & ".chr") Then
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "EjercitoReal", 0)
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
                Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(UserIndex).Name)
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & " expulsado de las fuerzas reales y prohibida la reenlistada" & _
                        FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toIndex, UserIndex, 0, "|| " & rData & ".chr inexistente." & FONTTYPE_INFO)

            End If

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 11)) = "/FORCEMIDI " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 11)

        If Not IsNumeric(rData) Then
            Exit Sub
        Else
            Call SendData(SendTarget.toall, 0, 0, "|| " & UserList(UserIndex).Name & " broadcast musica: " & rData & FONTTYPE_SERVER)
            Call SendData(SendTarget.toall, 0, 0, "TM" & rData)

        End If

    End If

    If UCase$(Left$(rData, 10)) = "/FORCEWAV " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 10)

        If Not IsNumeric(rData) Then
            Exit Sub
        Else
            Call SendData(SendTarget.toall, 0, 0, "TW" & rData)

        End If

    End If

    If UCase$(Left$(rData, 12)) = "/BORRARPENA " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        '/borrarpena pj pena
        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
    
        rData = Right$(rData, Len(rData) - 12)
    
        Name = ReadField(1, rData, Asc("@"))
        tStr = ReadField(2, rData, Asc("@"))
    
        If Name = "" Or tStr = "" Or Not IsNumeric(tStr) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Utilice /borrarpj Nick@NumeroDePena" & FONTTYPE_INFO)
            Exit Sub

        End If
    
        Name = Replace(Name, "\", "")
        Name = Replace(Name, "/", "")
    
        If FileExist(CharPath & Name & ".chr", vbNormal) Then
            rData = GetVar(CharPath & Name & ".chr", "PENAS", "P" & val(tStr))
            Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & val(tStr), LCase$(UserList(UserIndex).Name) & ": <Pena borrada> " & Date & " " & _
                    Time)

        End If
    
        Call LogGM(UserList(UserIndex).Name, " borro la pena: " & tStr & "-" & rData & " de " & Name)
        Exit Sub

    End If

    'Bloquear

    'Ultima ip de un char
    If UCase(Left(rData, 8)) = "/LASTIP " Then
        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right(rData, Len(rData) - 8)
    
        'No se si sea MUY necesario, pero por si las dudas... ;)
        rData = Replace(rData, "\", "")
        rData = Replace(rData, "/", "")
    
        If FileExist(CharPath & rData & ".chr", vbNormal) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La ultima IP de """ & rData & """ fue : " & GetVar(CharPath & rData & ".chr", _
                    "INIT", "LastIP") & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Charfile """ & rData & """ inexistente." & FONTTYPE_INFO)

        End If

        Exit Sub

    End If

    'Quita todos los NPCs del area
    If UCase$(rData) = "/LIMPIAR" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call LimpiarMundo
        Exit Sub

    End If

    'Mensaje del sistema
    If UCase$(Left$(rData, 6)) = "/SMSW " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 6)
        Call SendData(SendTarget.toall, 0, 0, "!!" & rData & ENDC)
        Exit Sub

    End If

    'Crear criatura, toma directamente el indice
    If UCase$(Left$(rData, 5)) = "/ACC " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 5)
        Call LogGM(UserList(UserIndex).Name, "Sumoneo a " & Npclist(val(rData)).Name & " en mapa " & UserList(UserIndex).pos.Map)
        Call SpawnNpc(val(rData), UserList(UserIndex).pos, True, False)
        Call LogGM(UserList(UserIndex).Name, " Sumoneo un " & Npclist(val(rData)).Name & " en mapa " & UserList(UserIndex).pos.Map)
        Exit Sub

    End If

    'Crear criatura con respawn, toma directamente el indice
    If UCase$(Left$(rData, 6)) = "/RACC " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)
        rData = Right$(rData, Len(rData) - 6)
        Call LogGM(UserList(UserIndex).Name, "Sumoneo con respawn " & Npclist(val(rData)).Name & " en mapa " & UserList(UserIndex).pos.Map)
        Call SpawnNpc(val(rData), UserList(UserIndex).pos, True, True)
        Call LogGM(UserList(UserIndex).Name, " Sumoneo un " & Npclist(val(rData)).Name & " en mapa " & UserList(UserIndex).pos.Map)
        Exit Sub

    End If

    'Comando para depurar la navegacion
    If UCase$(rData) = "/NAVE" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        If UserList(UserIndex).flags.Navegando = 1 Then
            UserList(UserIndex).flags.Navegando = 0
        Else
            UserList(UserIndex).flags.Navegando = 1

        End If

        Exit Sub

    End If

    If UCase$(rData) = "/QEVALGA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        If ServerSoloGMs > 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Servidor Válido para todos" & FONTTYPE_INFO)
            ServerSoloGMs = 0
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Servidor Válido solo a administradores." & FONTTYPE_INFO)
            ServerSoloGMs = 1

        End If

        Exit Sub

    End If

    'Apagamos
    If UCase$(rData) = "/OFFE" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call SendData(SendTarget.toall, UserIndex, 0, "||" & UserList(UserIndex).Name & " APAGA EL SERVIDOR!!!" & FONTTYPE_FIGHT)

        mifile = FreeFile
        Open App.Path & "\logs\Main.log" For Append Shared As #mifile
        Print #mifile, Date & " " & Time & " server apagado por " & UserList(UserIndex).Name & ". "
        Close #mifile
        Unload frmMain
        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "/CONDEN" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 8)
        TIndex = NameIndex(rData)

        If TIndex > 0 Then Call VolverCriminal(TIndex)
        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "/RAJAR " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 7)
        TIndex = NameIndex(UCase$(rData))

        If TIndex > 0 Then
            Call ResetFacciones(TIndex)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 11)) = "/RAJARCLAN " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 11)
        tInt = modGuilds.m_EcharMiembroDeClan(UserIndex, rData, False)  'me da el guildindex

        If tInt = 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| No pertenece a ningun clan o es fundador." & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Expulsado." & FONTTYPE_INFO)
            Call SendData(SendTarget.ToGuildMembers, tInt, 0, "|| " & rData & " ha sido expulsado del clan por los administradores del servidor" & _
                    FONTTYPE_GUILD)

        End If

        Exit Sub

    End If

    'lst email
    If UCase$(Left$(rData, 11)) = "/LASTEMAIL " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 11)

        If FileExist(CharPath & rData & ".chr") Then
            tStr = GetVar(CharPath & rData & ".chr", "CONTACTO", "email")
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Last email de " & rData & ":" & tStr & FONTTYPE_INFO)

        End If

        Exit Sub

    End If

    'altera email
    If UCase$(Left$(rData, 8)) = "/AIMAIL " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
       
        rData = Right$(rData, Len(rData) - 8)
        tStr = ReadField(1, rData, Asc("-"))

        If tStr = "" Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||usar /AEMAIL <pj>-<nuevomail>" & FONTTYPE_INFO)
            Exit Sub

        End If

        TIndex = NameIndex(tStr)

        If TIndex > 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El usuario esta online, no se puede si esta online" & FONTTYPE_INFO)
            Exit Sub

        End If

        Arg1 = ReadField(2, rData, Asc("-"))

        If Arg1 = "" Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||usar /AEMAIL <pj>-<nuevomail>" & FONTTYPE_INFO)
            Exit Sub

        End If

        If Not FileExist(CharPath & tStr & ".chr") Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No existe el charfile " & CharPath & tStr & ".chr" & FONTTYPE_INFO)
        Else
            Call WriteVar(CharPath & tStr & ".chr", "CONTACTO", "Email", Arg1)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Email de " & tStr & " cambiado a: " & Arg1 & FONTTYPE_INFO)

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 7)) = "/ANUER " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        
        rData = Right$(rData, Len(rData) - 7)
        tStr = ReadField(1, rData, Asc("@"))
        Arg1 = ReadField(2, rData, Asc("@"))
    
        If tStr = "" Or Arg1 = "" Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Usar: /ANAME origen@destino" & FONTTYPE_INFO)
            Exit Sub

        End If
    
        TIndex = NameIndex(tStr)

        If TIndex > 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El Pj esta online, debe salir para el cambio" & FONTTYPE_WARNING)
            Exit Sub

        End If
    
        If FileExist(CharPath & UCase(tStr) & ".chr") = False Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El pj " & tStr & " es inexistente " & FONTTYPE_INFO)
            Exit Sub

        End If
    
        Arg2 = GetVar(CharPath & UCase(tStr) & ".chr", "GUILD", "GUILDINDEX")

        If IsNumeric(Arg2) Then
            If CInt(Arg2) > 0 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||El pj " & tStr & _
                        " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido. " & FONTTYPE_INFO)
                Exit Sub

            End If

        End If
    
        If FileExist(CharPath & UCase(Arg1) & ".chr") = False Then
            FileCopy CharPath & UCase(tStr) & ".chr", CharPath & UCase(Arg1) & ".chr"
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Transferencia exitosa" & FONTTYPE_INFO)
            Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
            'ponemos la pena
            tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & ": BAN POR Cambio de nick a " & _
                    UCase$(Arg1) & " " & Date & " " & Time)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||El nick solicitado ya existe" & FONTTYPE_INFO)
            Exit Sub

        End If

        Exit Sub

    End If

    If UCase$(Left$(rData, 10)) = "/SHOWCMSG " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        rData = Right$(rData, Len(rData) - 10)
        Call modGuilds.GMEscuchaClan(UserIndex, rData)
        Exit Sub

    End If

    If UCase$(Left$(rData, 11)) = "/GUARDAMAPA" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
 
        Call GrabarMapa(UserList(UserIndex).pos.Map, App.Path & "\WorldBackUp\Mapa" & UserList(UserIndex).pos.Map)
        Exit Sub

    End If

    If UCase$(Left$(rData, 5)) = "/MAP " Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
 
        rData = Right(rData, Len(rData) - 5)

        Select Case UCase(ReadField(1, rData, 32))

            Case "PK"
                tStr = ReadField(2, rData, 32)

                If tStr <> "" Then
                    MapInfo(UserList(UserIndex).pos.Map).Pk = IIf(tStr = "0", True, False)
                    Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).pos.Map & ".dat", "Mapa" & UserList(UserIndex).pos.Map, "Pk", _
                            tStr)

                End If

                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Mapa " & UserList(UserIndex).pos.Map & " PK: " & MapInfo(UserList( _
                        UserIndex).pos.Map).Pk & FONTTYPE_INFO)

            Case "BACKUP"
                tStr = ReadField(2, rData, 32)

                If tStr <> "" Then
                    MapInfo(UserList(UserIndex).pos.Map).BackUp = CByte(tStr)
                    Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).pos.Map & ".dat", "Mapa" & UserList(UserIndex).pos.Map, _
                            "backup", tStr)

                End If
        
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Mapa " & UserList(UserIndex).pos.Map & " Backup: " & MapInfo(UserList( _
                        UserIndex).pos.Map).BackUp & FONTTYPE_INFO)

        End Select

        Exit Sub

    End If

    If UCase$(Left$(rData, 11)) = "/BORRAR SOS" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        Call Ayuda.Reset
        Exit Sub

    End If

    If UCase$(Left$(rData, 9)) = "/SHOW INT" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
   
        Call frmMain.mnuMostrar_Click
        Exit Sub

    End If


    If UCase$(rData) = "/ECHARTODOSPJSS" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
        Call EcharPjsNoPrivilegiados
        Exit Sub

    End If

    If UCase$(rData) = "/TCPESSTATS" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
  
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Los datos estan en BYTES." & FONTTYPE_INFO)

        With TCPESStats
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||IN/s: " & .BytesRecibidosXSEG & " OUT/s: " & .BytesEnviadosXSEG & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||IN/s MAX: " & .BytesRecibidosXSEGMax & " -> " & .BytesRecibidosXSEGCuando & _
                    FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||OUT/s MAX: " & .BytesEnviadosXSEGMax & " -> " & .BytesEnviadosXSEGCuando & _
                    FONTTYPE_INFO)

        End With

    End If

    If UCase$(rData) = "/RELOADNPCS" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub
            
        Call CargaNpcsDat

        Call SendData(SendTarget.toIndex, UserIndex, 0, "|| Npcs.dat y npcsHostiles.dat recargados." & FONTTYPE_INFO)
        Exit Sub

    End If

    If UCase$(rData) = "/RELOADSINI" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        Call LoadSini
        Exit Sub

    End If

    If UCase$(rData) = "/RELOADHECHIZOS" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        Call CargarHechizos
        Exit Sub

    End If

    If UCase$(rData) = "/RELOADOBJ" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)

        If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios <= PlayerType.SemiDios Then Exit Sub

        Call LoadOBJData
        Exit Sub

    End If

    If UCase$(rData) = "/REINICIAR" Then
        Call LogGM(UserList(UserIndex).Name, "Comando: " & rData)


        Call ReiniciarServidor(True)
        Exit Sub
    End If


ErrorHandler:
    Call LogError("HandleData. CadOri:" & CadenaOriginal & " Nom:" & UserList(UserIndex).Name & "UI:" & UserIndex & " N: " & Err.Number & " D: " & _
            Err.Description)
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

    Call SendData(SendTarget.toIndex, UserIndex, 0, "NOC" & IIf(DeNoche And (MapInfo(UserList(UserIndex).pos.Map).Zona = Campo Or MapInfo(UserList( _
            UserIndex).pos.Map).Zona = Ciudad), "1", "0"))
    Call SendData(SendTarget.toIndex, UserIndex, 0, "NOC" & IIf(DeNoche, "1", "0"))

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
    Call SendData(toIndex, UserIndex, 0, "NA" & tnow)
End Sub
