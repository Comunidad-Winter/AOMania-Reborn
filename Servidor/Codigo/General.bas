Attribute VB_Name = "General"
'test seth

Option Explicit

Global LeerNPCs As New clsIniManager
Global LeerNPCsHostil As New clsIniManager

Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal T As Long, ByVal r As String)

'[Pablo ToxicWaste]
Public Type ModClase

    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    AtaqueWrestling As Double
    Da�oArmas As Double
    Da�oProyectiles As Double
    Da�oWrestling As Double
    Escudo As Double

End Type

Public ModClase() As ModClase

Public Function ClaseToByte(ByVal Clase As String) As Byte

    Dim i As Long

    Clase = UCase$(Clase)

    'Modificadores de Clase
    For i = 1 To NUMCLASES

        If Clase = ListaClases(i) Then
            ClaseToByte = CByte(i)
            Exit For

        End If

    Next i

End Function

Public Function StringToClase(ByVal Class As String) As String
    Dim tStr As String

    Select Case UCase$(Class)

    Case "THESAUROS"
        tStr = "TRABAJADOR"

    Case "PESCADOR"
        tStr = "TRABAJADOR"

    Case "HERRERO"
        tStr = "TRABAJADOR"

    Case "LE�ADOR"
        tStr = "TRABAJADOR"

    Case "MINERO"
        tStr = "TRABAJADOR"

    Case "CARPINTERO"
        tStr = "TRABAJADOR"

    Case "SASTRE"
        tStr = "TRABAJADOR"

    Case "HERRERO MAGICO"
        tStr = "TRABAJADOR"

    Case Else

        tStr = Class

    End Select

    StringToClase = tStr

End Function

Public Sub LoadBalance()

    Dim ReadDat As clsIniManager
    Dim i As Long

    Set ReadDat = New clsIniManager

    ReDim ModClase(1 To NUMCLASES) As ModClase

    'Modificadores de Clase
    For i = 1 To NUMCLASES

        With ModClase(i)
            .Evasion = val(GetVar(DatPath & "Balance.dat", "MODEVASION", ListaClases(i)))
            .AtaqueArmas = val(GetVar(DatPath & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
            .AtaqueProyectiles = val(GetVar(DatPath & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
            .AtaqueWrestling = val(GetVar(DatPath & "Balance.dat", "MODATAQUEWRESTLING", ListaClases(i)))
            .Da�oArmas = val(GetVar(DatPath & "Balance.dat", "MODDA�OARMAS", ListaClases(i)))
            .Da�oProyectiles = val(GetVar(DatPath & "Balance.dat", "MODDA�OPROYECTILES", ListaClases(i)))
            .Da�oWrestling = val(GetVar(DatPath & "Balance.dat", "MODDA�OWRESTLING", ListaClases(i)))
            .Escudo = val(GetVar(DatPath & "Balance.dat", "MODESCUDO", ListaClases(i)))

        End With

    Next i

    Set ReadDat = Nothing

    Call LoadBalanceVida

End Sub

Sub CargarELU()

    Dim X As Long

    For X = 1 To STAT_MAXELV
        levelELU(X) = GetVar(DatPath & "Niveles.dat", "INIT", "Nivel" & X)
    Next X

End Sub

Public Sub LoadBalanceVida()

'Guerrero
    GCONST21MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA GUERRERO", "CONST21MINVIDA"))
    GCONST21MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA GUERRERO", "CONST21MAXVIDA"))
    GCONST20MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA GUERRERO", "CONST20MINVIDA"))
    GCONST20MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA GUERRERO", "CONST20MAXVIDA"))
    GCONST19MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA GUERRERO", "CONST19MINVIDA"))
    GCONST19MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA GUERRERO", "CONST19MAXVIDA"))
    GCONST18MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA GUERRERO", "CONST18MINVIDA"))
    GCONST18MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA GUERRERO", "CONST18MAXVIDA"))
    GCONST17MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA GUERRERO", "CONST17MINVIDA"))
    GCONST17MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA GUERRERO", "CONST17MAXVIDA"))
    GCONSTOTROMINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA GUERRERO", "CONSTOTROMINVIDA"))
    GCONSTOTROMAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA GUERRERO", "CONSTOTROMAXVIDA"))

    'Cazador
    CCONST21MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CAZADOR", "CONST21MINVIDA"))
    CCONST21MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CAZADOR", "CONST21MAXVIDA"))
    CCONST20MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CAZADOR", "CONST20MINVIDA"))
    CCONST20MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CAZADOR", "CONST20MAXVIDA"))
    CCONST19MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CAZADOR", "CONST19MINVIDA"))
    CCONST19MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CAZADOR", "CONST19MAXVIDA"))
    CCONST18MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CAZADOR", "CONST18MINVIDA"))
    CCONST18MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CAZADOR", "CONST18MAXVIDA"))
    CCONST17MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CAZADOR", "CONST17MINVIDA"))
    CCONST17MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CAZADOR", "CONST17MAXVIDA"))
    CCONSTOTROMINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CAZADOR", "CONSTOTROMINVIDA"))
    CCONSTOTROMAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CAZADOR", "CONSTOTROMAXVIDA"))

    'Paladin
    PCONST21MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA PALADIN", "CONST21MINVIDA"))
    PCONST21MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA PALADIN", "CONST21MAXVIDA"))
    PCONST20MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA PALADIN", "CONST20MINVIDA"))
    PCONST20MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA PALADIN", "CONST20MAXVIDA"))
    PCONST19MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA PALADIN", "CONST19MINVIDA"))
    PCONST19MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA PALADIN", "CONST19MAXVIDA"))
    PCONST18MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA PALADIN", "CONST18MINVIDA"))
    PCONST18MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA PALADIN", "CONST18MAXVIDA"))
    PCONST17MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA PALADIN", "CONST17MINVIDA"))
    PCONST17MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA PALADIN", "CONST17MAXVIDA"))
    PCONSTOTROMINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA PALADIN", "CONSTOTROMINVIDA"))
    PCONSTOTROMAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA PALADIN", "CONSTOTROMAXVIDA"))

    'Mago
    MCONST21MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA MAGO", "CONST21MINVIDA"))
    MCONST21MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA MAGO", "CONST21MAXVIDA"))
    MCONST20MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA MAGO", "CONST20MINVIDA"))
    MCONST20MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA MAGO", "CONST20MAXVIDA"))
    MCONST19MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA MAGO", "CONST19MINVIDA"))
    MCONST19MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA MAGO", "CONST19MAXVIDA"))
    MCONST18MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA MAGO", "CONST18MINVIDA"))
    MCONST18MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA MAGO", "CONST18MAXVIDA"))
    MCONST17MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA MAGO", "CONST17MINVIDA"))
    MCONST17MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA MAGO", "CONST17MAXVIDA"))
    MCONSTOTROMINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA MAGO", "CONSTOTROMINVIDA"))
    MCONSTOTROMAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA MAGO", "CONSTOTROMAXVIDA"))

    'Clerigo
    CLCONST21MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CLERIGO", "CONST21MINVIDA"))
    CLCONST21MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CLERIGO", "CONST21MAXVIDA"))
    CLCONST20MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CLERIGO", "CONST20MINVIDA"))
    CLCONST20MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CLERIGO", "CONST20MAXVIDA"))
    CLCONST19MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CLERIGO", "CONST19MINVIDA"))
    CLCONST19MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CLERIGO", "CONST19MAXVIDA"))
    CLCONST18MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CLERIGO", "CONST18MINVIDA"))
    CLCONST18MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CLERIGO", "CONST18MAXVIDA"))
    CLCONST17MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CLERIGO", "CONST17MINVIDA"))
    CLCONST17MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CLERIGO", "CONST17MAXVIDA"))
    CLCONSTOTROMINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CLERIGO", "CONSTOTROMINVIDA"))
    CLCONSTOTROMAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA CLERIGO", "CONSTOTROMAXVIDA"))

    'Asesino
    ACONST21MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ASESINO", "CONST21MINVIDA"))
    ACONST21MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ASESINO", "CONST21MAXVIDA"))
    ACONST20MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ASESINO", "CONST20MINVIDA"))
    ACONST20MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ASESINO", "CONST20MAXVIDA"))
    ACONST19MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ASESINO", "CONST19MINVIDA"))
    ACONST19MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ASESINO", "CONST19MAXVIDA"))
    ACONST18MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ASESINO", "CONST18MINVIDA"))
    ACONST18MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ASESINO", "CONST18MAXVIDA"))
    ACONST17MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ASESINO", "CONST17MINVIDA"))
    ACONST17MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ASESINO", "CONST17MAXVIDA"))
    ACONSTOTROMINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ASESINO", "CONSTOTROMINVIDA"))
    ACONSTOTROMAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ASESINO", "CONSTOTROMAXVIDA"))

    'Bardo
    BACONST21MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BARDO", "CONST21MINVIDA"))
    BACONST21MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BARDO", "CONST21MAXVIDA"))
    BACONST20MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BARDO", "CONST20MINVIDA"))
    BACONST20MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BARDO", "CONST20MAXVIDA"))
    BACONST19MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BARDO", "CONST19MINVIDA"))
    BACONST19MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BARDO", "CONST19MAXVIDA"))
    BACONST18MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BARDO", "CONST18MINVIDA"))
    BACONST18MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BARDO", "CONST18MAXVIDA"))
    BACONST17MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BARDO", "CONST17MINVIDA"))
    BACONST17MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BARDO", "CONST17MAXVIDA"))
    BACONSTOTROMINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BARDO", "CONSTOTROMINVIDA"))
    BACONSTOTROMAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BARDO", "CONSTOTROMAXVIDA"))

    'ladron
    LCONST21MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA LADRON", "CONST21MINVIDA"))
    LCONST21MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA LADRON", "CONST21MAXVIDA"))
    LCONST20MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA LADRON", "CONST20MINVIDA"))
    LCONST20MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA LADRON", "CONST20MAXVIDA"))
    LCONST19MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA LADRON", "CONST19MINVIDA"))
    LCONST19MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA LADRON", "CONST19MAXVIDA"))
    LCONST18MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA LADRON", "CONST18MINVIDA"))
    LCONST18MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA LADRON", "CONST18MAXVIDA"))
    LCONST17MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA LADRON", "CONST17MINVIDA"))
    LCONST17MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA LADRON", "CONST17MAXVIDA"))
    LCONSTOTROMINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA LADRON", "CONSTOTROMINVIDA"))
    LCONSTOTROMAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA LADRON", "CONSTOTROMAXVIDA"))

    'Druida
    DCONST21MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA DRUIDA", "CONST21MINVIDA"))
    DCONST21MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA DRUIDA", "CONST21MAXVIDA"))
    DCONST20MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA DRUIDA", "CONST20MINVIDA"))
    DCONST20MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA DRUIDA", "CONST20MAXVIDA"))
    DCONST19MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA DRUIDA", "CONST19MINVIDA"))
    DCONST19MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA DRUIDA", "CONST19MAXVIDA"))
    DCONST18MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA DRUIDA", "CONST18MINVIDA"))
    DCONST18MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA DRUIDA", "CONST18MAXVIDA"))
    DCONST17MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA DRUIDA", "CONST17MINVIDA"))
    DCONST17MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA DRUIDA", "CONST17MAXVIDA"))
    DCONSTOTROMINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA DRUIDA", "CONSTOTROMINVIDA"))
    DCONSTOTROMAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA DRUIDA", "CONSTOTROMAXVIDA"))

    'Trabajador
    TCONST21MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA TRABAJADOR", "CONST21MINVIDA"))
    TCONST21MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA TRABAJADOR", "CONST21MAXVIDA"))
    TCONST20MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA TRABAJADOR", "CONST20MINVIDA"))
    TCONST20MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA TRABAJADOR", "CONST20MAXVIDA"))
    TCONST19MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA TRABAJADOR", "CONST19MINVIDA"))
    TCONST19MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA TRABAJADOR", "CONST19MAXVIDA"))
    TCONST18MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA TRABAJADOR", "CONST18MINVIDA"))
    TCONST18MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA TRABAJADOR", "CONST18MAXVIDA"))
    TCONST17MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA TRABAJADOR", "CONST17MINVIDA"))
    TCONST17MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA TRABAJADOR", "CONST17MAXVIDA"))
    TCONSTOTROMINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA TRABAJADOR", "CONSTOTROMINVIDA"))
    TCONSTOTROMAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA TRABAJADOR", "CONSTOTROMAXVIDA"))

    'Brujo
    BCONST21MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BRUJO", "CONST21MINVIDA"))
    BCONST21MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BRUJO", "CONST21MAXVIDA"))
    BCONST20MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BRUJO", "CONST20MINVIDA"))
    BCONST20MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BRUJO", "CONST20MAXVIDA"))
    BCONST19MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BRUJO", "CONST19MINVIDA"))
    BCONST19MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BRUJO", "CONST19MAXVIDA"))
    BCONST18MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BRUJO", "CONST18MINVIDA"))
    BCONST18MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BRUJO", "CONST18MAXVIDA"))
    BCONST17MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BRUJO", "CONST17MINVIDA"))
    BCONST17MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BRUJO", "CONST17MAXVIDA"))
    BCONSTOTROMINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BRUJO", "CONSTOTROMINVIDA"))
    BCONSTOTROMAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA BRUJO", "CONSTOTROMAXVIDA"))

    'Arquero
    ARCONST21MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ARQUERO", "CONST21MINVIDA"))
    ARCONST21MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ARQUERO", "CONST21MAXVIDA"))
    ARCONST20MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ARQUERO", "CONST20MINVIDA"))
    ARCONST20MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ARQUERO", "CONST20MAXVIDA"))
    ARCONST19MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ARQUERO", "CONST19MINVIDA"))
    ARCONST19MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ARQUERO", "CONST19MAXVIDA"))
    ARCONST18MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ARQUERO", "CONST18MINVIDA"))
    ARCONST18MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ARQUERO", "CONST18MAXVIDA"))
    ARCONST17MINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ARQUERO", "CONST17MINVIDA"))
    ARCONST17MAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ARQUERO", "CONST17MAXVIDA"))
    ARCONSTOTROMINVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ARQUERO", "CONSTOTROMINVIDA"))
    ARCONSTOTROMAXVIDA = val(GetVar(DatPath & "Balance.dat", "MODVIDA ARQUERO", "CONSTOTROMAXVIDA"))

End Sub

Function ZonaCura(ByVal UserIndex As Integer) As Boolean

    Dim X As Integer, Y As Integer

    For Y = UserList(UserIndex).Pos.Y - MinYBorder + 1 To UserList(UserIndex).Pos.Y + MinYBorder - 1
        For X = UserList(UserIndex).Pos.X - MinXBorder + 1 To UserList(UserIndex).Pos.X + MinXBorder - 1

            If MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex > 0 Then
                If Npclist(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex).NPCtype = 1 Then
                    If Distancia(UserList(UserIndex).Pos, Npclist(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex).Pos) < 10 Then
                        ZonaCura = True
                        Exit Function

                    End If

                End If

            End If

        Next X
    Next Y

    ZonaCura = False

End Function

Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, Optional ByVal Mimetizado As Boolean = False)

    Select Case UCase$(UserList(UserIndex).Raza)

    Case "HUMANO"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 21
            Else
                UserList(UserIndex).char.Body = 21

            End If

        Case "MUJER"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 39
            Else
                UserList(UserIndex).char.Body = 39

            End If

        End Select

    Case "ELFO OSCURO"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 32
            Else
                UserList(UserIndex).char.Body = 32

            End If

        Case "MUJER"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 40
            Else
                UserList(UserIndex).char.Body = 40

            End If

        End Select

    Case "ENANO"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 53
            Else
                UserList(UserIndex).char.Body = 53

            End If

        Case "MUJER"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 60
            Else
                UserList(UserIndex).char.Body = 60

            End If

        End Select

    Case "GNOMO"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 53
            Else
                UserList(UserIndex).char.Body = 53

            End If

        Case "MUJER"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 60
            Else
                UserList(UserIndex).char.Body = 60

            End If

        End Select

    Case "HOBBIT"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 297
            Else
                UserList(UserIndex).char.Body = 297

            End If

        Case "MUJER"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 298
            Else
                UserList(UserIndex).char.Body = 298

            End If

        End Select

    Case "ORCO"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 300
            Else
                UserList(UserIndex).char.Body = 300

            End If

        Case "MUJER"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 302
            Else
                UserList(UserIndex).char.Body = 302

            End If

        End Select

    Case "LICANTROPO"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 21
            Else
                UserList(UserIndex).char.Body = 21

            End If

        Case "MUJER"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 39
            Else
                UserList(UserIndex).char.Body = 39

            End If

        End Select

    Case "VAMPIRO"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 32
            Else
                UserList(UserIndex).char.Body = 32

            End If

        Case "MUJER"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 40
            Else
                UserList(UserIndex).char.Body = 40

            End If

        End Select

    Case "CICLOPE"

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 21
            Else
                UserList(UserIndex).char.Body = 21

            End If

        Case "MUJER"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 39
            Else
                UserList(UserIndex).char.Body = 39

            End If

        End Select

    Case Else

        Select Case UCase$(UserList(UserIndex).Genero)

        Case "HOMBRE"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 21
            Else
                UserList(UserIndex).char.Body = 21

            End If

        Case "MUJER"

            If Mimetizado Then
                UserList(UserIndex).CharMimetizado.Body = 39
            Else
                UserList(UserIndex).char.Body = 39

            End If

        End Select

    End Select
    
    UserList(UserIndex).char.Head = UserList(UserIndex).char.Head

    UserList(UserIndex).flags.Desnudo = 1

End Sub

Sub Bloquear(ByVal sndRoute As Byte, _
             ByVal sndIndex As Integer, _
             ByVal sndMap As Integer, _
             Map As Integer, _
             ByVal X As Integer, _
             ByVal Y As Integer, _
           b As Byte)
'b=1 bloquea el tile en (x,y)
'b=0 desbloquea el tile indicado

    Call SendData(sndRoute, sndIndex, sndMap, "BQ" & X & "," & Y & "," & b)

End Sub

Function HayAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

    If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        If MapData(Map, X, Y).Graphic(1) >= 1505 And MapData(Map, X, Y).Graphic(1) <= 1520 And MapData(Map, X, Y).Graphic(2) = 0 Then
            HayAgua = True
        Else
            HayAgua = False

        End If

    Else
        HayAgua = False

    End If

End Function

Sub LimpiarObjs()

    Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Realizando Limpieza del Mundo" & FONTTYPE_Motd5)
    Dim i As Integer
    Dim Y As Integer
    Dim X As Integer
    Dim tInt As String

    For i = 1 To NumMaps
        For Y = YMinMapSize To YMaxMapSize
            For X = XMinMapSize To XMaxMapSize

                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(i, X, Y).OBJInfo.ObjIndex > 0 Then
                        tInt = ObjData(MapData(i, X, Y).OBJInfo.ObjIndex).ObjType

                        If ObjetosBorrable(MapData(i, X, Y).OBJInfo.ObjIndex) Then

                            Call EraseObj(ToMap, 0, i, MapData(i, X, Y).OBJInfo.Amount, i, X, Y)

                        End If

                    End If

                End If

            Next X
        Next Y
    Next i

    Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Limpieza del Mundo finalizada!" & FONTTYPE_Motd5)

End Sub

Sub LimpiarMundo()

    On Error Resume Next

    Dim i As Integer

    Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Realizando Limpieza del Mundo" & FONTTYPE_Motd5)

    For i = 1 To TrashCollector.Count
        Dim d As cGarbage
        Set d = TrashCollector(1)
        Call EraseObj(SendTarget.ToMap, 0, d.Map, 1, d.Map, d.X, d.Y)
        Call TrashCollector.Remove(1)
        Set d = Nothing
    Next i

    Call SecurityIp.IpSecurityMantenimientoLista
    Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Limpieza del Mundo finalizada!" & FONTTYPE_Motd5)

End Sub

Sub EnviarSpawnList(ByVal UserIndex As Integer)
    Dim k As Integer, SD As String
    SD = "SPL" & UBound(SpawnList) & ","

    For k = 1 To UBound(SpawnList)
        SD = SD & SpawnList(k).NpcName & ","
    Next k

    Call SendData(SendTarget.ToIndex, UserIndex, 0, SD)

End Sub

Sub Main()

    On Error Resume Next
    
    Set AodefConv = New AoDefenderConverter

    Dim f As Date

    Call SoundCarga

    ChDir App.Path
    ChDrive App.Path

    #If MYSQL = 1 Then
        Call Load_Mysql
        DoEvents
        Call Add_DataBase("0", "Status")
    #End If


    Call Load_Rank

    Call LoadMotd
    Call BanIpCargar

    Prision.Map = 48
    Libertad.Map = 48

    Prision.X = RandomNumber(67, 69)
    Prision.Y = RandomNumber(47, 52)
    Libertad.X = 75
    Libertad.Y = 65

    LastBackup = Format(Now, "Short Time")
    Minutos = Format(Now, "Short Time")

    ReDim Npclist(1 To MAXNPCS) As npc    'NPCS
    ReDim CharList(1 To MAXCHARS) As Integer
    ReDim Parties(1 To MAX_PARTIES) As clsParty
    ReDim Guilds(1 To MAX_GUILDS) As clsClan

    IniPath = App.Path & "\"
    DatPath = App.Path & "\Dat\"

    ListaRazas(1) = "Humano"
    ListaRazas(2) = "Elfo"
    ListaRazas(3) = "Elfo Oscuro"
    ListaRazas(4) = "Gnomo"
    ListaRazas(5) = "Enano"
    ListaRazas(6) = "Hobbit"
    ListaRazas(7) = "Orco"
    ListaRazas(8) = "Licantropo"
    ListaRazas(9) = "Vampiro"
    ListaRazas(10) = "Ciclope"

    Call modHDSerial.load_HDList
    Call mod_Climas.InitTimeLife

    ReDim LevelSkill(1 To STAT_MAXELV) As LevelSkill

    Dim i As Long

    For i = 1 To STAT_MAXELV

        If i * 2.51 < 100 Then
            LevelSkill(i).LevelValue = i * 2.51

            If LevelSkill(i).LevelValue Mod 10 > 5 Then LevelSkill(i).LevelValue = LevelSkill(i).LevelValue - 1
        Else
            LevelSkill(i).LevelValue = 100

        End If

    Next i

    'CP1="MAGO"
    'CP2="CLERIGO"
    'CP3="GUERRERO"
    'CP4="ASESINO"
    'CP5="LADRON"
    'CP6="BARDO"
    'CP7="DRUIDA"
    'CP8="TRABAJADOR"
    'CP9="PALADIN"
    'CP10="CAZADOR"
    'CP11="PIRATA"
    'CP12="BRUJO"
    'CP13="ARQUERO"

    ListaClases(1) = "MAGO"
    ListaClases(2) = "CLERIGO"
    ListaClases(3) = "GUERRERO"
    ListaClases(4) = "ASESINO"
    ListaClases(5) = "LADRON"
    ListaClases(6) = "BARDO"
    ListaClases(7) = "DRUIDA"
    ListaClases(8) = "TRABAJADOR"
    ListaClases(9) = "PALADIN"
    ListaClases(10) = "CAZADOR"
    ListaClases(11) = "PIRATA"
    ListaClases(12) = "BRUJO"
    ListaClases(13) = "ARQUERO"
    ListaClases(14) = "DIOS"

    SkillsNames(1) = "Suerte"
    SkillsNames(2) = "Magia"
    SkillsNames(3) = "Robar"
    SkillsNames(4) = "Tacticas de combate"
    SkillsNames(5) = "Combate con armas"
    SkillsNames(6) = "Meditar"
    SkillsNames(7) = "Apu�alar"
    SkillsNames(8) = "Ocultarse"
    SkillsNames(9) = "Supervivencia"
    SkillsNames(10) = "Talar arboles"
    SkillsNames(11) = "Comercio"
    SkillsNames(12) = "Defensa con escudos"
    SkillsNames(13) = "Resistencia Magica"
    SkillsNames(14) = "Pesca"
    SkillsNames(15) = "Mineria"
    SkillsNames(16) = "Carpinteria"
    SkillsNames(17) = "Herreria"
    SkillsNames(18) = "Liderazgo"
    SkillsNames(19) = "Domar animales"
    SkillsNames(20) = "Armas de proyectiles"
    SkillsNames(21) = "Wresterling"
    SkillsNames(22) = "Navegacion"
    SkillsNames(23) = "Sastreria"
    SkillsNames(24) = "Recolectar hierba"
    SkillsNames(25) = "Hechiceria"
    SkillsNames(26) = "Herrero M�gico"

    frmCargando.Show

    frmMain.caption = frmMain.caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    IniPath = App.Path & "\"
    CharPath = App.Path & "\Charfile\"

    'Bordes del mapa
    MinXBorder = XMinMapSize + (XWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)
    DoEvents

    frmCargando.Label1(2).caption = "Iniciando Arrays..."

    Call LoadGuildsDB

    Call CargarSpawnList
    Call CargarForbidenWords
    Call CargaPalabrasCensuradas
    Call CargarMensajesEnvio
    
    '�?�?�?�?�?�?�?� CARGAMOS DATOS DESDE ARCHIVOS �??�?�?�?�?�?�?�

    frmCargando.Label1(2).caption = "Cargando configuraciones del Servidor."
    MaxUsers = 0
    Call LoadSini
    Call LoadAntiCheat

    frmCargando.Label1(2).caption = "Cargando apuestas.."
    Call CargaApuestas

    '*************************************************
    Call CargaNpcsDat
    '*************************************************

    frmCargando.Label1(2).caption = "Cargando Objetos.."
    Call LoadOBJData
    
    frmCargando.Label1(2).caption = "Cargando Jerarquias.."
    Call CargarJerarquias
    
    frmCargando.Label1(2).caption = "Cargando Objetos Canjes Sagrados."
    Call LoadCanjeSagrados

    frmCargando.Label1(2).caption = "Cargando Hechizos.."
    Call CargarHechizos

    frmCargando.Label1(2).caption = "Cargando Zonas.."
    Call LoadZonas
    Barcos.TiempoRest = 60

    frmCargando.Label1(2).caption = "Cargando Objetos de Trabajadores.."
    Call LoadArmasHerreria
    Call LoadArmadurasHerreria
    Call LoadObjCarpintero
    Call LoadObjSastre
    Call LoadObjHechizeria
    Call LoadObjArmaHerreroMagico
    Call LoadObjArmaduraHerreroMagico

    If BootDelBackUp Then
        frmCargando.Label1(2).caption = "Cargando BackUp"
        Call CargarBackUp
    Else
        frmCargando.Label1(2).caption = "Cargando Mapas"
        Call LoadMapData
    End If
    
    Call SonidosMapas.LoadSoundMapInfo

    frmCargando.Label1(2).caption = "Cargando Balance.."
    Call LoadBalance
    
    frmCargando.Label1(2).caption = "Cargando Quests.."
    Call Load_Quest
    
    frmCargando.Label1(2).caption = "Cargando Subastas.."
    Call CargarSubastas

    'Comentado porque hay worldsave en ese mapa!
    'Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
    '�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�

    Dim LoopC As Integer

    'Resetea las conexiones de los usuarios
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
    Next LoopC

    '�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�

    With frmMain
        .AutoSave.Enabled = True
        .tPiqueteC.Enabled = True
        .GameTimer.Enabled = True
        .FX.Enabled = True
        .Auditoria.Enabled = True
        .TIMER_AI.Enabled = True
        .npcataca.Enabled = True
    End With

    centinelaActivado = True

    '�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�
    'Configuracion de los sockets

    Call SecurityIp.InitIpTables(1000)

    If LastSockListen >= 0 Then Call apiclosesocket(LastSockListen)    'Cierra el socket de escucha
    Call IniciaWsApi(frmMain.hWnd)
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

    If SockListen <> -1 Then
        Call WriteVar(IniPath & "Server.ini", "INIT", "LastSockListen", SockListen)    ' Guarda el socket escuchando
    Else
        MsgBox "Ha ocurrido un error al iniciar el socket del Servidor.", vbCritical + vbOKOnly

    End If

    If frmMain.Visible Then frmMain.txStatus.caption = "Escuchando conexiones entrantes ..."
    '�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�

    Unload frmCargando

    'Log
    Dim n As Integer
    n = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #n
    Print #n, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
    Close #n

    'Ocultar
    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)

    End If

    tInicioServer = GetTickCount() And &H7FFFFFFF
    denuncias = True

    CountTC = RandomNumber(1, 23)
    NocheLicantropo = False

    Call CargarCastillos
    Call Load_Criatura
    Call LoadGuerras
    Call Mod_Monedas.Load_Creditos
    Call Mod_Monedas.Load_Canjes
    Call LoadCasaEncantada
    Call LoadDragonAlado
    Call TimeChange
    Call LoadClima(CountTC)
    Call LoadGMs

    ReDim ValidMap(1 To NumMaps) As Byte

    For i = 1 To NumMaps

        If i <> 34 And i <> 1 And i <> 59 And i <> 84 And i <> 132 And i <> 86 Then
            ValidMap(i) = 0
        Else
            ValidMap(i) = 1
        End If

    Next i

End Sub

Function FileExist(ByVal File As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************
    FileExist = Dir$(File, FileType) <> ""

End Function

Public Function readfield(ByVal Pos As Long, ByRef Text As String, ByVal SepASCII As Long) As String
'*****************************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'Gets a field from a delimited string
'*****************************************************************

    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long

    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, ChrW$(SepASCII), vbBinaryCompare)
    Next i

    If CurrentPos = 0 Then
        readfield = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        readfield = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If

End Function

Function readfield2(ByVal iPos As Long, _
                    ByRef sText As String, _
                    ByVal CharAscii As Long) As String
    ' Mismo que anterior con los parametros formales...
 
    '
    ' @ maTih.-
     
    Dim Read_Field() As String
 
    'Creo un array temporal.
    Read_Field = Split(sText, ChrW$(CharAscii))
    ' Mismo que antes con chrW
     
    If (iPos - 1) <= UBound(Read_Field()) Then

        'devuelve
        readfield2 = (Read_Field(iPos - 1))

    End If
     
End Function

Function MapaValido(ByVal Map As Integer) As Boolean
    MapaValido = Map >= 1 And Map <= NumMaps

End Function

Public Sub LogCriticEvent(Desc As String)

    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoReal(Desc As String)

    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
    Print #nfile, Desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoCaos(Desc As String)

    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
    Print #nfile, Desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogIndex(ByVal Index As Integer, ByVal Desc As String)

    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\" & Index & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogError(Desc As String)

    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogStatic(Desc As String)

    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogTarea(Desc As String)

    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile(1)    ' obtenemos un canal
    Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogClanes(ByVal str As String)

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub

Public Sub LogIP(ByVal str As String)

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\IP.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub

Public Sub Alas(ByVal str As String)

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\ALAS.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub
Public Sub LogAsedio(ByVal str As String)

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\ASEDIO.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub

Public Sub LogDesarrollo(ByVal str As String)

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\desarrollo.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub

Public Sub LogNivel(ByVal UserIndex As Integer, ByVal str As String)

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\NewUser\" & UserList(UserIndex).Name & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub

Public Sub LogGM(nombre As String, texto As String)

    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal

    Open App.Path & "\logs\Gms\" & nombre & ".log" For Append Shared As #nfile

    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogCreditos(nombre As String, Tiene As Long, Gasto As Long, Item As String)
    On Error GoTo errhandler
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\Donaciones\" & nombre & ".log" For Append Shared As #nfile
    Print #nfile, Date & "  " & Time; " " & nombre & ": Ten�a: " & Tiene & " Gast�: " & Gasto & " Obtuvo: " & Item
    Close #nfile
    Exit Sub
errhandler:
End Sub

Public Sub LogCanjes(Opcion As Integer, nombre As String, Tiene As Long, Gasto As Long, Item As String)
    On Error GoTo errhandler
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\Canjeadores\" & nombre & ".log" For Append Shared As #nfile
    If Opcion = 1 Then
        Print #nfile, Date & "  " & Time; " " & nombre & ": Ten�a: " & Tiene & " Gast�: " & Gasto & " Obtuvo: " & Item
    ElseIf Opcion = 2 Then
        Print #nfile, Date & "  " & Time; " " & nombre & ": Ten�a: " & Tiene & " Gan�: " & Gasto & " Por el item: " & Item
    End If
    Close #nfile
    Exit Sub
errhandler:
End Sub

Public Sub LogTelepatia(Envia As String, Recibe As String, Mensaje As String)

    On Error GoTo errhandler

    Dim nfile As Integer

    nfile = FreeFile

    Open App.Path & "\logs\Telepatia\" & Envia & ".log" For Append Shared As #nfile

    Print #nfile, Date & " " & Time & " Telepatia a " & Recibe & ": " & Mensaje
    Close #nfile

errhandler:

End Sub

Public Sub LogUser(nombre As String, texto As String)

    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal

    Open App.Path & "\logs\Usuarios\" & nombre & ".log" For Append Shared As #nfile

    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogAsesinato(texto As String)

    On Error GoTo errhandler

    Dim nfile As Integer

    nfile = FreeFile    ' obtenemos un canal

    Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub logVentaCasa(ByVal texto As String)

    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal

    Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogHackAttemp(texto As String)

    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogCheating(texto As String)

    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\CH.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogCriticalHackAttemp(texto As String)

    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogAntiCheat(texto As String)

    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, ""
    Close #nfile

    Exit Sub

errhandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
    Dim Arg As String
    Dim i As Integer

    For i = 1 To 33

        Arg = readfield2(i, cad, 44)

        If Arg = "" Then Exit Function

    Next i

    ValidInputNP = True

End Function

Sub Restart()

'Se asegura de que los sockets estan cerrados e ignora cualquier err
    On Error Resume Next

    If frmMain.Visible Then frmMain.txStatus.caption = "Reiniciando."

    Dim LoopC As Integer

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)

    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

    For LoopC = 1 To MaxUsers
        Call CloseSocket(LoopC)
    Next

    ReDim UserList(1 To MaxUsers)

    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
    Next LoopC

    LastUser = 0
    NumUsers = 0

    ReDim Npclist(1 To MAXNPCS) As npc    'NPCS
    ReDim CharList(1 To MAXCHARS) As Integer

    Call LoadSini
    Call LoadOBJData
    Call LoadMapData
    Call CargarHechizos

    If frmMain.Visible Then frmMain.txStatus.caption = "Escuchando conexiones entrantes ..."

    'Log it
    Dim n As Integer
    n = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #n
    Print #n, Date & " " & Time & " servidor reiniciado."
    Close #n

    'Ocultar

    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)

    End If

End Sub

Public Function Intemperie(ByVal UserIndex As Integer) As Boolean

    With UserList(UserIndex)

        If MapInfo(.Pos.Map).Zona <> "DUNGEON" Then
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger <> 1 And MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger <> 2 And MapData(.Pos.Map, .Pos.X, _
                                                                                                                             .Pos.Y).Trigger <> 4 Then Intemperie = True
        Else
            Intemperie = False

        End If

    End With

End Function

Public Sub EfectoLluvia(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    On Error GoTo errhandler

    If UserList(UserIndex).flags.UserLogged Then
        If Intemperie(UserIndex) Then
            Call QuitarSta(UserIndex, Porcentaje(UserList(UserIndex).Stats.MaxSta, 3))
            Call EnviarSta(UserIndex)
        End If
    End If

    Exit Sub
errhandler:
    LogError ("Error en EfectoLluvia")

End Sub

Public Sub TiempoInvocacion(ByVal UserIndex As Integer)

    Dim i As Integer

    For i = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasIndex(i) > 0 Then
            If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia - 1

                If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))

            End If

        End If

    Next i

End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)

    Dim modifi As Integer

    If UserList(UserIndex).flags.Privilegios >= PlayerType.Consejero Then Exit Sub

    If UserList(UserIndex).Counters.Frio < IntervaloFrio Then
        UserList(UserIndex).Counters.Frio = UserList(UserIndex).Counters.Frio + 1
    Else

        If MapInfo(UserList(UserIndex).Pos.Map).Terreno = Nieve Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||��Estas muriendo de frio, abrigate o moriras!!." & FONTTYPE_INFO)
            modifi = Porcentaje(UserList(UserIndex).Stats.MaxHP, 5)
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - modifi

            If UserList(UserIndex).Stats.MinHP < 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||��Has muerto de frio!!." & FONTTYPE_INFO)
                UserList(UserIndex).Stats.MinHP = 0
                Call UserDie(UserIndex)

            End If

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ASH" & UserList(UserIndex).Stats.MinHP)
        Else
            modifi = Porcentaje(UserList(UserIndex).Stats.MaxSta, 5)
            Call QuitarSta(UserIndex, modifi)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ASS" & UserList(UserIndex).Stats.MinSta)

            'Call SendData(SendTarget.toIndex, UserIndex, 0, "||��Has perdido stamina, si no te abrigas rapido perderas toda!!." & FONTTYPE_INFO)

        End If

        UserList(UserIndex).Counters.Frio = 0

    End If

End Sub

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)

    If UserList(UserIndex).Counters.Mimetismo < IntervaloInvisible Then
        UserList(UserIndex).Counters.Mimetismo = UserList(UserIndex).Counters.Mimetismo + 1
        'Else
        'restore old char
        'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Recuperas tu apariencia normal." & FONTTYPE_INFO)

        'UserList(UserIndex).char.Body = UserList(UserIndex).CharMimetizado.Body
        'UserList(UserIndex).char.Head = UserList(UserIndex).CharMimetizado.Head
        'UserList(UserIndex).char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
        'UserList(UserIndex).char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
        'UserList(UserIndex).char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim

        'UserList(UserIndex).Counters.Mimetismo = 0
        'UserList(UserIndex).flags.Mimetizado = 0
        'Call ChangeUserChar(SendTarget.ToMap, UserIndex, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).char.Heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim)
    End If

End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)

    If UserList(UserIndex).Counters.Invisibilidad > 0 Then

        UserList(UserIndex).Counters.Invisibilidad = UserList(UserIndex).Counters.Invisibilidad - 1
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "INVI" & UserList(UserIndex).Counters.Invisibilidad)
    Else
        UserList(UserIndex).Counters.Invisibilidad = IntervaloInvisible
        UserList(UserIndex).flags.Invisible = 0
        UserList(UserIndex).Counters.Ocultando = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "INVI0")

        If UserList(UserIndex).flags.Oculto = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z11")
            Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).char.CharIndex & ",0," & UserList( _
                                                                            UserIndex).PartyIndex)
        End If

    End If

End Sub

Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)

    If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
        Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
    Else
        Npclist(NpcIndex).flags.Paralizado = 0
        Npclist(NpcIndex).flags.Inmovilizado = 0
    End If

End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)

    If UserList(UserIndex).Counters.Ceguera > 0 Then
        UserList(UserIndex).Counters.Ceguera = UserList(UserIndex).Counters.Ceguera - 1
    Else

        If UserList(UserIndex).flags.Ceguera = 1 Then
            UserList(UserIndex).flags.Ceguera = 0
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "NSEGUE")

        End If

        If UserList(UserIndex).flags.Estupidez = 1 Then
            UserList(UserIndex).flags.Estupidez = 0
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "NESTUP")

        End If

    End If

End Sub

Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)

    If UserList(UserIndex).Counters.Paralisis > 0 Then
        UserList(UserIndex).Counters.Paralisis = UserList(UserIndex).Counters.Paralisis - 1
    Else
        UserList(UserIndex).flags.Paralizado = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "PARADOW")
        Call Corr_ActualizarPosicion(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    End If

End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
    If UserList(UserIndex).flags.Desnudo = 0 Then

        If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Trigger = 1 And MapData(UserList( _
                                                                                                                              UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Trigger = 2 And MapData(UserList(UserIndex).Pos.Map, UserList( _
                                                                                                                                                                                                                                                             UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Trigger = 4 Then Exit Sub

        Dim massta As Integer

        If UserList(UserIndex).Stats.MinSta < UserList(UserIndex).Stats.MaxSta Then
            If UserList(UserIndex).Counters.STACounter < Intervalo Then
                UserList(UserIndex).Counters.STACounter = UserList(UserIndex).Counters.STACounter + 1
            Else
                EnviarStats = True
                UserList(UserIndex).Counters.STACounter = 0
                massta = RandomNumber(1, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5))
                UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta + massta

                If UserList(UserIndex).Stats.MinSta > UserList(UserIndex).Stats.MaxSta Then
                    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta

                End If

            End If

        End If

    End If

End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean)

    If UserList(UserIndex).Counters.Veneno < IntervaloVeneno Then
        UserList(UserIndex).Counters.Veneno = UserList(UserIndex).Counters.Veneno + 1
        '  Call SendData(SendTarget.toindex, UserIndex, 0, "ASH" & UserList(UserIndex).Stats.MinHP)
    Else

        If UserList(UserIndex).TipoVeneno = 0 Or UserList(UserIndex).TipoVeneno = 2 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & 37 & "," & 1)
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & 53 & "," & 1)

        End If

        UserList(UserIndex).Counters.Veneno = 0

        If UserList(UserIndex).flags.HechizoVeneno = 0 Then
            UserList(UserIndex).Da�oVeneno = 1
        Else
            UserList(UserIndex).Da�oVeneno = RandomNumber(6, UserList(UserIndex).AumentoVeneno)

        End If

        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - UserList(UserIndex).Da�oVeneno
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El veneno te causa " & UserList(UserIndex).Da�oVeneno & " puntos de da�o. Si no te curas moriras." & FONTTYPE_VENENO)

        If UserList(UserIndex).Stats.MinHP < 1 Then Call UserDie(UserIndex)
        Call EnviarHP(UserIndex)

    End If

End Sub

Public Sub EfectoMetamorfosis(ByVal UserIndex As Integer)

    If UserList(UserIndex).Counters.Metamorfosis > 0 Then
        UserList(UserIndex).Counters.Metamorfosis = UserList(UserIndex).Counters.Metamorfosis - 1

    End If

    If UserList(UserIndex).Counters.Metamorfosis = 0 Then
        UserList(UserIndex).char.Body = UserList(UserIndex).CharMimetizado.Body
        UserList(UserIndex).char.Head = UserList(UserIndex).CharMimetizado.Head
        UserList(UserIndex).char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
        UserList(UserIndex).char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
        UserList(UserIndex).char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
        UserList(UserIndex).char.Alas = UserList(UserIndex).CharMimetizado.Alas
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).CharMimetizado.Fuerza
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).CharMimetizado.Agilidad
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).CharMimetizado.Inteligencia

        UserList(UserIndex).Counters.Metamorfosis = 0
        UserList(UserIndex).flags.Metamorfosis = 0
        UserList(UserIndex).CharMimetizado.Fuerza = 0
        UserList(UserIndex).CharMimetizado.Agilidad = 0
        UserList(UserIndex).CharMimetizado.Inteligencia = 0

        Call ChangeUserChar(SendTarget.ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)

        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & 1 & "," & 1)

    End If

End Sub

Public Sub EfectoAttack(ByVal UserIndex As Integer)

    If UserList(UserIndex).Counters.TimerAttack > 0 Then
        UserList(UserIndex).Counters.TimerAttack = UserList(UserIndex).Counters.TimerAttack - 1
    End If

End Sub

Public Sub DuracionPociones(ByVal UserIndex As Integer)

    'Controla la duracion de las pociones
    With UserList(UserIndex)

        If .flags.DuracionEfectoAmarillas > 0 Then
            .flags.DuracionEfectoAmarillas = .flags.DuracionEfectoAmarillas - 1

            If .flags.DuracionEfectoAmarillas <= 0 Then

                .flags.TomoPocionAmarilla = False

                'volvemos el atributo al estado normal
                .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributosBackUP(eAtributos.Agilidad)

                Call EnviarAmarillas(UserIndex)
                Exit Sub

            End If

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ATG" & .flags.DuracionEfectoAmarillas)

        End If

        If .flags.DuracionEfectoVerdes > 0 Then
            .flags.DuracionEfectoVerdes = .flags.DuracionEfectoVerdes - 1

            If .flags.DuracionEfectoVerdes <= 0 Then

                .flags.TomoPocionVerde = False

                'volvemos el atributo al estado normal
                .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributosBackUP(eAtributos.Fuerza)

                Call EnviarVerdes(UserIndex)
                Exit Sub

            End If

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "VTG" & .flags.DuracionEfectoVerdes)

        End If

    End With

End Sub

Public Sub HambreYSed(ByVal UserIndex As Integer, ByRef fEnviarAyS As Boolean)

'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)

        If .flags.Privilegios > 1 Then Exit Sub

        'Sed
        If .Stats.MinAGU > 0 Then
            If .Counters.AGUACounter < IntervaloSed Then
                .Counters.AGUACounter = .Counters.AGUACounter + 1
            Else
                .Counters.AGUACounter = 0
                .Stats.MinAGU = .Stats.MinAGU - 10

                If .Stats.MinAGU <= 0 Then
                    .Stats.MinAGU = 0
                    .flags.Sed = 1

                End If

                fEnviarAyS = True

            End If

        End If

        'hambre
        If .Stats.MinHam > 0 Then
            If .Counters.COMCounter < IntervaloHambre Then
                .Counters.COMCounter = .Counters.COMCounter + 1
            Else
                .Counters.COMCounter = 0
                .Stats.MinHam = .Stats.MinHam - 10

                If .Stats.MinHam <= 0 Then
                    .Stats.MinHam = 0
                    .flags.Hambre = 1

                End If

                fEnviarAyS = True

            End If

        End If

    End With

End Sub

Public Sub Sanar(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
        '<EhHeader>
        On Error GoTo Sanar_Err
        '</EhHeader>

100     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Trigger = 1 And _
       MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Trigger = 2 And _
       MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Trigger = 4 Then Exit Sub

        Dim mashit As Integer

        'con el paso del tiempo va sanando....pero muy lentamente ;-)
102     If UserList(UserIndex).Stats.MinHP < UserList(UserIndex).Stats.MaxHP Then
104         If UserList(UserIndex).Counters.HPCounter < Intervalo Then
106             UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
            Else
108             mashit = RandomNumber(2, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5))

110             UserList(UserIndex).Counters.HPCounter = 0
112             UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + mashit

114             If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList( _
                   UserIndex).Stats.MaxHP
116             Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has sanado " & mashit & " Puntos de vida." & FONTTYPE_INFO)
118             EnviarStats = True
            
120             If UserList(UserIndex).PartyIndex > 0 Then
122                    Parties(UserList(UserIndex).PartyIndex).UpdateUserParty
                End If
            
            End If

        End If

        '<EhFooter>
        Exit Sub

Sanar_Err:
        Debug.Print err.Description & vbCrLf & _
               "in AoManiaServer.General.Sanar " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Public Sub CargaNpcsDat()

    Call LeerNPCs.Initialize(DatPath & "NPCs.dat")
    Call LeerNPCsHostil.Initialize(DatPath & "NPCs-HOSTILES.dat")

End Sub

Sub PasarSegundo()

    Dim Saturin As Integer

    Dim Pos     As WorldPos

    Dim Posa    As WorldPos

    Dim i       As Long

    For i = 1 To LastUser

        If SecondaryWeather Then Call EfectoLluvia(i)

        'Cerrar usuario
        If UserList(i).Counters.Saliendo Then
            UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1

            If UserList(i).Counters.Salir <= 0 Then
                'If NumUsers <> 0 Then NumUsers = NumUsers - 1

                Call SendData(SendTarget.ToIndex, i, 0, "||Gracias por jugar AoMania" & FONTTYPE_INFO)
                Call SendData(SendTarget.ToIndex, i, 0, "FINOC")

                Call CloseSocket(i)
                Exit Sub
            Else
                Call SendData(SendTarget.ToIndex, i, 0, "||En " & UserList(i).Counters.Salir & " segundos se cerrara el juego..." & FONTTYPE_INFO)

            End If

            'ANTIEMPOLLOS
        ElseIf UserList(i).flags.EstaEmpo = 1 Then
            UserList(i).EmpoCont = UserList(i).EmpoCont + 1

            If UserList(i).EmpoCont = 18 Then
                Call Encarcelar(i, "25", UserList(i).Name)

                UserList(i).EmpoCont = 0

                Exit Sub
            ElseIf UserList(i).EmpoCont = 3 Then
                Call SendData(SendTarget.ToIndex, i, 0, "||Est�s bloqueando el acceso a un objeto, mu�vete o ser�s encarcelado." & FONTTYPE_INFO)
            ElseIf UserList(i).EmpoCont = 8 Then
                Call SendData(SendTarget.ToIndex, i, 0, "||Est�s bloqueando el acceso a un objeto, mu�vete o ser�s encarcelado." & FONTTYPE_INFO)
            ElseIf UserList(i).EmpoCont = 13 Then
                Call SendData(SendTarget.ToIndex, i, 0, "||Est�s bloqueando el acceso a un objeto, mu�vete o ser�s encarcelado." & FONTTYPE_INFO)

            End If

        End If

    Next i

    If Encuesta.ACT = 1 Then
        Encuesta.Tiempo = Encuesta.Tiempo + 1

        If Encuesta.Tiempo = 45 Then
            Call SendData(SendTarget.ToAll, 0, 0, "||Faltan 15 segundos para terminar la encuesta." & FONTTYPE_GUILD)
        ElseIf Encuesta.Tiempo = 60 Then
            Call SendData(SendTarget.ToAll, 0, 0, "||Encuesta: Terminada con �xito" & FONTTYPE_TALK)
            Call SendData(SendTarget.ToAll, 0, 0, "||SI: " & Encuesta.EncSI & " / NO: " & Encuesta.EncNO & FONTTYPE_TALK)

            If Encuesta.EncNO < Encuesta.EncSI Then
                Call SendData(SendTarget.ToAll, 0, 0, "||Gana: SI" & FONTTYPE_GUILD)
            ElseIf Encuesta.EncSI < Encuesta.EncNO Then
                Call SendData(SendTarget.ToAll, 0, 0, "||Gana: NO" & FONTTYPE_GUILD)
            ElseIf Encuesta.EncNO = Encuesta.EncSI Then
                Call SendData(SendTarget.ToAll, 0, 0, "||Encuesta empatada." & FONTTYPE_GUILD)

            End If

            Encuesta.ACT = 0
            Encuesta.Tiempo = 0
            Encuesta.EncNO = 0
            Encuesta.EncSI = 0

            For Saturin = 1 To LastUser

                If UserList(Saturin).flags.VotEnc = True Then UserList(Saturin).flags.VotEnc = False
            Next Saturin

        End If

        Exit Sub

    End If

    If CuentaRegresiva > 0 Then
        If CuentaRegresiva > 1 Then
            Call SendData(SendTarget.ToAll, 0, 0, "||En..." & CuentaRegresiva - 1 & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToAll, 0, 0, "||YA!!!!" & FONTTYPE_GUILD)

        End If

        CuentaRegresiva = CuentaRegresiva - 1

    End If

End Sub

Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)

'commit experiencias


'Guardar Pjs
    Call GuardarUsuarios

    'WorldSave
    Call DoBackUp

    If EjecutarLauncher Then Shell (App.Path & "\AoM.exe")

    'Chauuu
    Unload frmMain

End Sub

Sub GuardarUsuarios(Optional ByVal DoBackUp As Boolean = True)

    If DoBackUp Then
        haciendoBK = True

        Call SendData(SendTarget.ToAll, 0, 0, "BKW")
        Call SendData(SendTarget.ToAll, 0, 0, "||����(_.���`����GUARDANDO PERSONAJES������`�._)����" & FONTTYPE_WorldCarga)

    End If

    Dim i As Long

    For i = 1 To LastUser

        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, CharPath & UCase$(UserList(i).Name) & ".chr")

        End If

    Next i

    If DoBackUp Then
        Call SendData(SendTarget.ToAll, 0, 0, "||����(_.���`����PERSONAJES GUARDADOS������`�._)����" & FONTTYPE_WorldSave)
        Call SendData(SendTarget.ToAll, 0, 0, "BKW")

        haciendoBK = False

    End If

End Sub

Sub ActSlot()
    Dim LoopC As Integer

    For LoopC = 1 To MaxUsers

        If UserList(LoopC).ConnID <> -1 And Not UserList(LoopC).flags.UserLogged Then
            Call CloseSocket(LoopC)
        End If

    Next LoopC
End Sub

Public Sub LoadAntiCheat()

    Dim i As Integer
    Lac_Camina = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Caminar")))
    Lac_Lanzar = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Lanzar")))
    Lac_Usar = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Usar")))
    Lac_Tirar = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Tirar")))
    Lac_Pociones = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Pociones")))
    Lac_Pegar = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Pegar")))

    For i = 1 To MaxUsers
        ResetearLac i
    Next

End Sub

Public Sub ResetearLac(UserIndex As Integer)

    With UserList(UserIndex).Lac
        .LCaminar.init Lac_Camina
        .LPociones.init Lac_Pociones
        .LUsar.init Lac_Usar
        .LPegar.init Lac_Pegar
        .LLanzar.init Lac_Lanzar
        .LTirar.init Lac_Tirar

    End With

End Sub

Public Sub CargaLac(UserIndex As Integer)

    With UserList(UserIndex).Lac
        Set .LCaminar = New Cls_InterGTC
        Set .LLanzar = New Cls_InterGTC
        Set .LPegar = New Cls_InterGTC
        Set .LPociones = New Cls_InterGTC
        Set .LTirar = New Cls_InterGTC
        Set .LUsar = New Cls_InterGTC
        .LCaminar.init Lac_Camina
        .LPociones.init Lac_Pociones
        .LUsar.init Lac_Usar
        .LPegar.init Lac_Pegar
        .LLanzar.init Lac_Lanzar
        .LTirar.init Lac_Tirar

    End With

End Sub

Public Sub DescargaLac(UserIndex As Integer)

    Exit Sub

    With UserList(UserIndex).Lac
        Set .LCaminar = Nothing
        Set .LLanzar = Nothing
        Set .LPegar = Nothing
        Set .LPociones = Nothing
        Set .LTirar = Nothing
        Set .LUsar = Nothing

    End With

End Sub

Public Sub DuracionAngelyDemonio(ByVal UserIndex As Integer)

    Dim QuitaSta As Byte

    With UserList(UserIndex)

        If .Counters.Frio < IntervaloFrio Then
            .Counters.Frio = .Counters.Frio + 1
        Else

            If .Stats.MinSta > 0 Then
                QuitaSta = RandomNumber(40, 55)
                Call QuitarSta(UserIndex, QuitaSta)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "ASS" & .Stats.MinSta)
            Else

                .char.Body = .CharMimetizado.Body
                .char.Head = .CharMimetizado.Head
                .char.CascoAnim = .CharMimetizado.CascoAnim
                .char.ShieldAnim = .CharMimetizado.ShieldAnim
                .char.WeaponAnim = .CharMimetizado.WeaponAnim
                .char.Alas = .CharMimetizado.Alas

                .Counters.Mimetismo = 0
                .flags.Mimetizado = 0

                Call ChangeUserChar(SendTarget.ToMap, UserIndex, .Pos.Map, UserIndex, _
                                    .char.Body, .char.Head, .char.heading, .char.WeaponAnim, .char.ShieldAnim, _
                                    .char.CascoAnim, .char.Alas)

                Call SendData(SendTarget.ToPCArea, UserIndex, .Pos.Map, "CFX" & _
                                                                        .char.CharIndex & "," & 1 & "," & 1)

                If .Metamorfosis.Angel = 1 Then
                    .Metamorfosis.Angel = 0
                ElseIf .Metamorfosis.Demonio = 1 Then
                    .Metamorfosis.Demonio = 0
                End If

            End If

            .Counters.Frio = 0

        End If

    End With

End Sub

Public Function MensajeCensura(ByVal Mensaje As String) As String
     
     Dim i As Integer
     Dim data As String
     
     For i = 1 To UBound(PalabrasCensuradas)
           
           If InStr(LCase(Mensaje), LCase(PalabrasCensuradas(i))) Then
              Mensaje = Replace(LCase(Mensaje), LCase$(PalabrasCensuradas(i)), "#%$!#")
           End If
           
     Next i
     
     MensajeCensura = Mensaje
     
End Function

Public Sub EnviaMensajesClient()
      
      Dim Total As Integer
      Dim Mensaje As String
      
      Total = UBound(EnvioMensajes)
      Total = RandomNumber(1, Total)
     Mensaje = EnvioMensajes(Total)
      
            
            Call SendData(ToAll, 0, 0, "EMC" & Mensaje)
      
End Sub
