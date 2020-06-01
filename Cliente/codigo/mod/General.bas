Attribute VB_Name = "Mod_General"
Option Explicit

Private Type PicMenuIco
    Diablo As StdPicture
    Mano As StdPicture
    Cruceta As StdPicture
        
    Ico_Diablo As Picture
    Ico_Mano As Picture
    Ico_Cruceta As Picture
End Type

Private Type PicMiniMapa
    Mapa As StdPicture
End Type

Private Type PicInterfaces

    FrmMain_Principal As StdPicture
    FrmMain_Inventario As StdPicture
    FrmMain_Hechizos As StdPicture
    
    Clima_Dia As StdPicture
    Clima_Noche As StdPicture
    
    FrmBanco_Principal As StdPicture
    FrmBanco_Depositar As StdPicture
    FrmBanco_Retirar As StdPicture
                
    FrmComerciar_Principal As StdPicture
    FrmComerciar_Comprar As StdPicture
    FrmComerciar_Vender As StdPicture
    
    FrmConnect_Principal As StdPicture
    
    FrmConnect_BtConectar As StdPicture
    FrmConnect_BtConectarApretado As StdPicture
    
    FrmConnect_BtCrearPj As StdPicture
    FrmConnect_BtCrearPjApretado As StdPicture
    
    FrmConnect_BtRecuperar As StdPicture
    FrmConnect_BtRecuperarApretado As StdPicture
    
    FrmCrearPersonaje_Principal As StdPicture
    
    FrmSkill_Principal As StdPicture
    FrmMapa_Principal As StdPicture

    FrmSoporteGM_Principal As StdPicture
    FrmSoporte_Principal As StdPicture
    FrmRespuesta_Principal As StdPicture
    FrmCargando_Principal As StdPicture
    FrmRecuPass_Principal As StdPicture
    
    FrmEstadisticas_Principal As StdPicture
    
    FrmCantidad_Principal As StdPicture
    FrmMayor_Principal As StdPicture
    FrmCustomKeys_Principal As StdPicture
    FrmBancoInfo_Principal As StdPicture
    FrmBancoDepositar_Principal As StdPicture
    FrmBancoRetirar_Principal As StdPicture
    FrmBancoFinal_Principal As StdPicture
    
    FrmOlvidarHechizo_Principal As StdPicture
    
    FrmCabezas_Principal As StdPicture
    
    FrmHerrero_Principal As StdPicture
    FrmCarp_Principal As StdPicture
    FrmSastre_Principal As StdPicture
    
    FrmHechiceria_Principal As StdPicture
    
    FrmQuest_SinHacer As StdPicture
    FrmQuest_Terminado As StdPicture
    
End Type

Private Enum PlayerType

    User = 0
    Consejero = 1
    SemiDios = 2
    Dios = 3
    Admin = 4

End Enum

Public Interfaces As PicInterfaces
Public MiniMapa As PicMiniMapa
Public Iconos As PicMenuIco

Public bFogata    As Boolean

''
' Retrieves the active window's hWnd for this app.
'
' @return Retrieves the active window's hWnd for this app. If this app is not in the foreground it returns 0.

Private Declare Function GetActiveWindow Lib "user32" () As Long

Public Sub LoadIconos()
    Dim Data() As Byte
    With Iconos
           
        If Get_File_Data(DirRecursos, "DIABLO.ICO", Data, ICONOS_FILE) Then
            Set .Diablo = ArrayToPicture(Data(), 0, UBound(Data) + 1)
            Set .Ico_Diablo = .Diablo
        End If
        
        Erase Data
        
        If Get_File_Data(DirRecursos, "MANO.ICO", Data, ICONOS_FILE) Then
            Set .Mano = ArrayToPicture(Data(), 0, UBound(Data) + 1)
            Set .Ico_Mano = .Mano
        End If
        
        Erase Data
        
        If Get_File_Data(DirRecursos, "CRUCETA.ICO", Data, ICONOS_FILE) Then
          Set .Cruceta = ArrayToPicture(Data(), 0, UBound(Data) + 1)
          'Set .Ico_Cruceta = .Cruceta
        End If
        
        Erase Data
           
    End With
        
End Sub

Public Sub LoadInterfaces()

    Dim Data() As Byte

    With Interfaces
              
        'Get Picture
        If Get_File_Data(DirRecursos, "PRINCIPAL.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmMain_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)

        End If

        Erase Data
     
        'Get Picture
        If Get_File_Data(DirRecursos, "HECHIZOS.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmMain_Hechizos = ArrayToPicture(Data(), 0, UBound(Data) + 1)

        End If

        Erase Data

        'Get Picture
        If Get_File_Data(DirRecursos, "INVENTARIO.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmMain_Inventario = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data

        'Get Picture
        If Get_File_Data(DirRecursos, "DIA.JPG", Data, INT_RESOURCE_FILE) Then
            Set .Clima_Dia = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data

        'Get Picture
        If Get_File_Data(DirRecursos, "NOCHE.JPG", Data, INT_RESOURCE_FILE) Then
            Set .Clima_Noche = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data
        
        'Get Picture
        If Get_File_Data(DirRecursos, "BANCO.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmBanco_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data

        'Get Picture
        If Get_File_Data(DirRecursos, "BTDEPOSITAR.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmBanco_Depositar = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data

        'Get Picture
        If Get_File_Data(DirRecursos, "BTRETIRAR.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmBanco_Retirar = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data
        
        'Get Picture
        If Get_File_Data(DirRecursos, "COMERCIO.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmComerciar_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data

        'Get Picture
        If Get_File_Data(DirRecursos, "BTCOMPRA.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmComerciar_Comprar = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data

        'Get Picture
        If Get_File_Data(DirRecursos, "BTVENDER.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmComerciar_Vender = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data
        
        'Get Picture
        If Get_File_Data(DirRecursos, "CONECTAR.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmConnect_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data
        
        'Get Picture
        If Get_File_Data(DirRecursos, "BTCONNECT.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmConnect_BtConectar = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data

        'Get Picture
        If Get_File_Data(DirRecursos, "BTCONNECT2.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmConnect_BtConectarApretado = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data
        
        'Get Picture
        If Get_File_Data(DirRecursos, "BTCREARPJ.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmConnect_BtCrearPj = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data

        'Get Picture
        If Get_File_Data(DirRecursos, "BTCREARPJ2.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmConnect_BtCrearPjApretado = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data

        'Get Picture
        If Get_File_Data(DirRecursos, "BTRECUPERAR.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmConnect_BtRecuperar = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data

        'Get Picture
        If Get_File_Data(DirRecursos, "BTRECUPERAR2.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmConnect_BtRecuperarApretado = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data
        
        'Get Picture
        If Get_File_Data(DirRecursos, "CREARPJ.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmCrearPersonaje_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data
        
        'Get Picture
        If Get_File_Data(DirRecursos, "SKILLS.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmSkill_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data

        'Get Picture
        If Get_File_Data(DirRecursos, "MAPA.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmMapa_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data

        'Get Picture
        If Get_File_Data(DirRecursos, "SOPORTEGM.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmSoporteGM_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data
        
        'Get Picture
        If Get_File_Data(DirRecursos, "SOPORTE.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmSoporte_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data

        'Get Picture
        If Get_File_Data(DirRecursos, "RESPUESTA.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmRespuesta_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data

        'Get Picture
        If Get_File_Data(DirRecursos, "CARGANDO.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmCargando_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data
        
        'Get Picture
        If Get_File_Data(DirRecursos, "RECUPASS.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmRecuPass_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
           
        End If

        Erase Data
        
        If Get_File_Data(DirRecursos, "STATS_USER.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmEstadisticas_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        End If
        
        Erase Data
        
         If Get_File_Data(DirRecursos, "TIRARI.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmCantidad_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        End If
        
        If Get_File_Data(DirRecursos, "MAYORES.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmMayor_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        End If
        
        Erase Data
        
        If Get_File_Data(DirRecursos, "TECLAS.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmCustomKeys_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        End If
        
        Erase Data
        
         If Get_File_Data(DirRecursos, "BANCOPRINCI.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmBancoInfo_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        End If
        
        Erase Data
        
         If Get_File_Data(DirRecursos, "BANCODEPO.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmBancoDepositar_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        End If
        
        Erase Data
        
         If Get_File_Data(DirRecursos, "BANCORETIRAR.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmBancoRetirar_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        End If
        
        Erase Data
        
         If Get_File_Data(DirRecursos, "BANCOFINAL.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmBancoFinal_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        End If
        
        Erase Data
        
        If Get_File_Data(DirRecursos, "OLVHECHIZO.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmOlvidarHechizo_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        End If
        
        Erase Data
        
        If Get_File_Data(DirRecursos, "CAMBCAB.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmCabezas_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        End If
        
        Erase Data
        
        If Get_File_Data(DirRecursos, "HERRERIA.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmHerrero_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        End If
        
        Erase Data
        
        If Get_File_Data(DirRecursos, "CARPINTERO.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmCarp_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        End If
        
        Erase Data
        
        If Get_File_Data(DirRecursos, "SASTRERIA.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmSastre_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        End If
        
        Erase Data
        
         If Get_File_Data(DirRecursos, "HECHIZERIA.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmHechiceria_Principal = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        End If
        
        Erase Data
        
        If Get_File_Data(DirRecursos, "QUESTSH.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmQuest_SinHacer = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        End If
        
        Erase Data
        
        If Get_File_Data(DirRecursos, "QUESTTM.JPG", Data, INT_RESOURCE_FILE) Then
            Set .FrmQuest_Terminado = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        End If
        
        Erase Data
                
    End With

End Sub

Public Sub SacarScreen()

    Dim i As Integer
    Clipboard.Clear
    DoEvents
    Call keybd_event(VK_SNAPSHOT, PS_TheScreen, 0, 0)
    DoEvents

    For i = 1 To 1000

        If Not FileExist(DirScreenshot & "foto" & i & ".jpg", vbNormal) Then Exit For
    Next
    
    SavePicture Clipboard.GetData, DirScreenshot & "foto" & i & ".jpg"
    
    Call AddtoRichTextBox(frmMain.RecTxt, "Screenshot guardada en " & DirScreenshot & "foto" & i & ".jpg !", 255, 150, 50, False, False, False)

End Sub

Public Function EsGm(ByVal charindex As Integer) As Boolean

    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 23/01/2007
    '***************************************************

    EsGm = (CharList(charindex).priv And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero))

End Function

''
' Checks if this is the active (foreground) application or not.
'
' @return   True if any of the app's windows are the foreground window, false otherwise.

Private Function IsAppActive() As Boolean
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (maraxus)
    'Last Modify Date: 03/03/2007
    'Checks if this is the active application or not
    '***************************************************
    
    IsAppActive = (GetActiveWindow <> 0)

End Function

Public Function DirScreenshot() As String

    DirScreenshot = App.Path & "\Screenshots\"

End Function

Public Function DirFotos() As String

    DirFotos = App.Path & "\Imagenes\"

End Function

Public Function DirMiniMapa() As String

    DirMiniMapa = App.Path & "\MiniMapa\"

End Function

Public Function DirFont() As String

    DirFont = App.Path & "\Libs\Font\"

End Function

Public Function DirConfiguracion() As String

    DirConfiguracion = App.Path & "\Libs\Configuracion\"

End Function

Public Function DirRecursos() As String

    DirRecursos = App.Path & "\Libs\"

End Function

Public Function DirIconos() As String

    DirIconos = App.Path & "\Libs\Iconos\"

End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long

    'Initialize randomizer
    Call Randomize(Timer)
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound

End Function

Sub CargarAnimArmas()

    Dim LooPC  As Long
    Dim arch   As String
   
    Dim Data() As Byte
    Dim handle As Integer
        
    If Not Get_File_Data(DirRecursos, "ARMAS.DAT", Data, INIT_RESOURCE_FILE) Then Exit Sub

    arch = DirRecursos & "Armas.dat"
    
    handle = FreeFile
    Open arch For Binary Access Write As handle
    Put handle, , Data
    Close handle
     
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For LooPC = 1 To NumWeaponAnims

        If LooPC <> 2 Then
            InitGrh WeaponAnimData(LooPC).WeaponWalk(1), Val(GetVar(arch, "ARMA" & LooPC, "Dir1")), 0
            InitGrh WeaponAnimData(LooPC).WeaponWalk(2), Val(GetVar(arch, "ARMA" & LooPC, "Dir2")), 0
            InitGrh WeaponAnimData(LooPC).WeaponWalk(3), Val(GetVar(arch, "ARMA" & LooPC, "Dir3")), 0
            InitGrh WeaponAnimData(LooPC).WeaponWalk(4), Val(GetVar(arch, "ARMA" & LooPC, "Dir4")), 0

        End If

    Next LooPC

    If FileExist(arch, vbArchive) Then Call Kill(arch)
       
End Sub

Sub CargarColores()

    Dim arch   As String
    Dim Data() As Byte
    Dim handle As Integer
    
    If Not Get_File_Data(DirRecursos, "COLORES.DAT", Data, INIT_RESOURCE_FILE) Then Exit Sub

    arch = DirRecursos & "COLORES.dat"
    
    handle = FreeFile
    Open arch For Binary Access Write As handle
    Put handle, , Data
    Close handle
        
    Dim i As Long

    For i = 0 To 48 '46-47-48-49-50 Reservados.
        ColoresPJ(i) = D3DColorXRGB(CInt(GetVar(arch, CStr(i), "R")), CInt(GetVar(arch, CStr(i), "G")), CInt(GetVar(arch, CStr(i), "B")))
    Next i
    
    ' Ciuda
    ColoresPJ(49) = D3DColorXRGB(CInt(GetVar(arch, "CI", "R")), CInt(GetVar(arch, "CI", "G")), CInt(GetVar(arch, "CI", "B")))

    ' Crimi
    ColoresPJ(50) = D3DColorXRGB(CInt(GetVar(arch, "CR", "R")), CInt(GetVar(arch, "CR", "G")), CInt(GetVar(arch, "CR", "B")))

    If FileExist(arch, vbArchive) Then Call Kill(arch)

End Sub

Sub CargarAnimEscudos()

    Dim LooPC  As Long
    Dim arch   As String
    Dim Data() As Byte
    Dim handle As Integer
    
    If Not Get_File_Data(DirRecursos, "ESCUDOS.DAT", Data, INIT_RESOURCE_FILE) Then Exit Sub

    arch = DirRecursos & "ESCUDOS.DAT"
    
    handle = FreeFile
    Open arch For Binary Access Write As handle
    Put handle, , Data
    Close handle
 
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For LooPC = 1 To NumEscudosAnims

        If LooPC <> 2 Then
            InitGrh ShieldAnimData(LooPC).ShieldWalk(1), Val(GetVar(arch, "ESC" & LooPC, "Dir1")), 0
            InitGrh ShieldAnimData(LooPC).ShieldWalk(2), Val(GetVar(arch, "ESC" & LooPC, "Dir2")), 0
            InitGrh ShieldAnimData(LooPC).ShieldWalk(3), Val(GetVar(arch, "ESC" & LooPC, "Dir3")), 0
            InitGrh ShieldAnimData(LooPC).ShieldWalk(4), Val(GetVar(arch, "ESC" & LooPC, "Dir4")), 0

        End If

    Next LooPC

    If FileExist(arch, vbArchive) Then Call Kill(arch)
    
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, _
    ByVal Text As String, _
    Optional ByVal Red As Integer = -1, _
    Optional ByVal Green As Integer, _
    Optional ByVal Blue As Integer, _
    Optional ByVal bold As Boolean = False, _
    Optional ByVal italic As Boolean = False, _
    Optional ByVal bCrLf As Boolean = False)

    With RichTextBox

        If Len(.Text) > 10000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf)
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
            .Text = ""

        End If
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
    End With

End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i   As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function

        End If

    Next i
    
    AsciiValidos = True

End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim LooPC     As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Dirección de email invalida")
        Exit Function

    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function

    End If
    
    For LooPC = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, LooPC, 1))

        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function

        End If

    Next LooPC
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function

    End If
    
    If Len(UserName) > 20 Then
        MsgBox ("El Nombre de tu Personaje debe tener menos de 20 letras.")
        Exit Function

    End If
    
    For LooPC = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, LooPC, 1))

        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function

        End If

    Next LooPC
    
    CheckUserData = True

End Function

Sub UnloadAllForms()
    
    Dim mifrm As Form
    
    For Each mifrm In Forms

        Unload mifrm
    Next

End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean

    '*****************************************************************
    'Only allow characters that are Win 95 filename compatible
    '*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function

    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function

    End If
    
    If KeyAscii > 126 Then
        Exit Function

    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or _
        KeyAscii = 124 Then
        Exit Function

    End If
    
    'else everything is cool
    LegalCharacter = True

End Function

Sub SetConnected()
    '*****************************************************************
    'Sets the client to "Connect" mode
    '*****************************************************************
    'Set Connected
    Connected = True
    
    'Unload the connect form
    Unload frmConnect
    
    'Load main form
    frmMain.Visible = True

End Sub

Sub MoveTo(ByVal Direccion As E_Heading)

    '***************************************************
    'Author: Alejandro Santos (AlejoLp)
    'Last Modify Date: 06/03/2006
    ' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
    '***************************************************
    Dim LegalOk  As Boolean
    
    Dim userTick As Long

    Dim now      As Long

    'now = (GetTickCount() And &H7FFFFFFF)
    now = (GetTickCount())

    userTick = TickCountServer + (now - TickCountClient)
    
    If Cartel Then Cartel = False
    
    Select Case Direccion

        Case E_Heading.NORTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y - 1)

        Case E_Heading.EAST
            LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)

        Case E_Heading.SOUTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)

        Case E_Heading.WEST
            LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)

    End Select
    
    If UserEstupido Then
          
        Select Case Direccion
          
            Case E_Heading.NORTH
                LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)

            Case E_Heading.EAST
                LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)

            Case E_Heading.SOUTH
                LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)

            Case E_Heading.WEST
                LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)
          
        End Select
    
    End If
    
    If LegalOk And UserMeditar Then Exit Sub
    
    If LegalOk And Not UserParalizado And Not UserMeditar Then

        Dim Heading As Byte

        Heading = Direccion
        Call EnviaM(Heading)
        'Call SendData("Ñ" & Direccion & "," & userTick)
            
        If Not UserDescansar And Not UserMeditar Then
        
            Call MoveCharbyHead(UserCharIndex, Direccion)
            Call MoveScreen(Direccion)

        End If

    Else

        If CharList(UserCharIndex).Heading <> Direccion Then
            If UserParalizado And UserInmovilizado Then Exit Sub
            Call SendData("CHEA" & Direccion)

        End If

    End If
    
    Call ActualizarShpUserPos
    
    Call Audio.MoveListener(UserPos.X, UserPos.Y)
    
    TextoMapa = MapInfo.Name & " (  " & UserMap & "   X: " & CharList(UserCharIndex).pos.X & " Y: " & CharList(UserCharIndex).pos.Y & ")"
    
End Sub



Sub RandomMove()
    '***************************************************
    'Author: Alejandro Santos (AlejoLp)
    'Last Modify Date: 06/03/2006
    ' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
    '***************************************************

    MoveTo RandomNumber(1, 4)
    
End Sub

Sub CheckKeys()
    
    If Comerciando Then Exit Sub
    If pausa Then Exit Sub
    
    'Static lastMovement As Long

    'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
    'If GetTickCount - lastMovement > 56 Then
    'lastMovement = GetTickCount
    'Else
    'Exit Sub
    'End If
    
    'If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyToggleMAPA)) < 0 Then
    '   If frmMain.SendTxt.Visible Then Exit Sub
    '
    '   If frmMapa.Visible = False Then frmMapa.Visible = True
    'ElseIf frmMapa.Visible Then
    '    frmMapa.Visible = False
'
'    End If
    

    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then

            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                Call MoveTo(NORTH)
                ' frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.y & ")"
                Exit Sub

            End If
       
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                Call MoveTo(EAST)
                ' frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.y & ")"
                Exit Sub

            End If
       
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                Call MoveTo(SOUTH)
                'frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.y & ")"
                Exit Sub

            End If
       
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                Call MoveTo(WEST)
                'frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.y & ")"
                Exit Sub

            End If

        Else
            Dim kp As Boolean

            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0

            If kp Then Call RandomMove

            ' frmMain.Coord.Caption = "(" & UserPos.x & "," & UserPos.y & ")"
        End If

    End If
    
    Call RefreshAllChars

End Sub


Sub SwitchMap(ByVal Map As Integer)
    '**************************************************************
    'Formato de mapas optimizado para reducir el espacio que ocupan.
    'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
    '**************************************************************
    
    Call CleanerPlus
    Call Particle_Group_Remove_All
    

    Dim Y        As Long
    Dim X        As Long
    Dim TempInt  As Integer
    Dim ByFlags  As Byte
    Dim FileBuff As clsByteBuffer
    Dim ii       As Long
    Set FileBuff = New clsByteBuffer
    
    Dim Data() As Byte

    If Not Get_File_Data(DirRecursos, "MAPA" & Map & ".MAP", Data, MAPAS_RESOURCE_FILE) Then Exit Sub
   
    FileBuff.initializeReader Data
   
    'map Header
    MapInfo.MapVersion = FileBuff.getInteger
   
    MiCabecera.desc = FileBuff.getString(Len(MiCabecera.desc))
    MiCabecera.CRC = FileBuff.getLong
    MiCabecera.MagicWord = FileBuff.getLong

    FileBuff.getDouble
   
    'Load arrays

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            'Get handle, , ByFlags
            ByFlags = FileBuff.getByte()
           
            MapData(X, Y).Blocked = (ByFlags And 1)
           
            'Get handle, , MapData(X, Y).Graphic(1).GrhIndex
            MapData(X, Y).Graphic(1).GrhIndex = FileBuff.getLong()
            InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
           
            'Layer 2 used?

            If ByFlags And 2 Then
                'Get handle, , MapData(X, Y).Graphic(2).GrhIndex
                MapData(X, Y).Graphic(2).GrhIndex = FileBuff.getLong()
                InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
            Else
                MapData(X, Y).Graphic(2).GrhIndex = 0

            End If
               
            'Layer 3 used?

            If ByFlags And 4 Then
                'Get handle, , MapData(X, Y).Graphic(3).GrhIndex
                MapData(X, Y).Graphic(3).GrhIndex = FileBuff.getLong()
                InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
            Else
                MapData(X, Y).Graphic(3).GrhIndex = 0

            End If
               
            'Layer 4 used?

            If ByFlags And 8 Then
                'Get handle, , MapData(X, Y).Graphic(4).GrhIndex
                MapData(X, Y).Graphic(4).GrhIndex = FileBuff.getLong()
                InitGrh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex
            Else
                MapData(X, Y).Graphic(4).GrhIndex = 0

            End If
           
            'Trigger used?

            If ByFlags And 16 Then
                'Get handle, , MapData(X, Y).Trigger
                MapData(X, Y).Trigger = FileBuff.getInteger()
            Else
                MapData(X, Y).Trigger = 0

            End If
            
            For ii = 1 To 5 'inicialiamos los grhs de la sangre
                Call InitGrh(MapData(X, Y).Sangre(ii).grhSangre, 17355, 0)
                
            Next ii
           
            'Erase NPCs

            If MapData(X, Y).charindex > 0 Then
                Call EraseChar(MapData(X, Y).charindex)

            End If
           
            'Erase OBJs
            MapData(X, Y).ObjGrh.GrhIndex = 0

        Next X
    Next Y

    'Close handle
     
     'CRAW; 13/03/2020 --> HARDCODED
     If (Map = 1) Then
     
        
        MapData(50, 50).Particle_Group = 1
        Call Particle_Create(1, 45, 46, -1)
     
     End If
     
     
    Set FileBuff = Nothing
      
    MapInfo.Name = vbNullString
    MapInfo.Music = vbNullString
   
    CurMap = Map
    
    Call CargarMiniMapa

End Sub

'TODO : Reemplazar por la nueva versión, esta apesta!!!
Public Function ReadField(ByVal pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String

    '*****************************************************************
    'Gets a field from a string
    '*****************************************************************
    Dim i         As Integer
    Dim lastPos   As Integer
    Dim CurChar   As String * 1
    Dim FieldNum  As Integer
    Dim Seperator As String
    
    Seperator = Chr$(SepASCII)
    lastPos = 0
    FieldNum = 0
    
    For i = 1 To Len(Text)
        CurChar = mid$(Text, i, 1)

        If CurChar = Seperator Then
            FieldNum = FieldNum + 1

            If FieldNum = pos Then
                ReadField = mid$(Text, lastPos + 1, (InStr(lastPos + 1, Text, Seperator, vbTextCompare) - 1) - (lastPos))
                Exit Function

            End If

            lastPos = i

        End If

    Next i

    FieldNum = FieldNum + 1
    
    If FieldNum = pos Then
        ReadField = mid$(Text, lastPos + 1)

    End If

End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (LenB(Dir$(File, FileType)) <> 0)

End Function

Sub Main()
            
    #If Launcher = 1 Then
          
          Call LoadDataLauncher
          
          If PlayLauncher = 0 Then
               MsgBox "¡Para ejecutar AoMania.exe, debes pasar primero por el Launcher!", vbInformation
               End
          
          ElseIf PlayLauncher = 1 Then
                PlayLauncher = 0
                Call SaveDataLauncher
                
          End If
          
     #End If
     
    Set AodefConv = New AoDefenderConverter
    
    'If AoDefDebugger Then
    '   Call AoDefAntiDebugger
    '   End
   'End If
        
    Call InitializeCompression
    Call LoadInterfaces
    Call LoadClientSetup
    Call LoadIconos
      
    'Set resolution BEFORE the loading form is displayed, therefore it will be centered.
    Call Resolution.SetResolution
    
    'If FindPreviousInstance Then
    '    Call MsgBox("Ya está siendo ejecutado AOMania!.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
    '    End
    'End If

    Call LeerLineaComandos
    Call Disco
   
    Dim EstaBloqueado As Byte
    EstaBloqueado = Val(GetVar(DirConfiguracion & "sinfo.dat", "s10", "PJ"))

    If EstaBloqueado = 1 Then
        Call MsgBox("Tu Cliente ha sido Bloqueado, Consulta a un Game Master para Solucionarlo", vbCritical + vbOKOnly)
        End

    End If
            
    ' Load constants, classes, flags, graphics..
    LoadInitialConfig
    Call LoadClientSetup
    Call LoadSoundConfig
    
    frmMain.Socket1.Startup
    frmConnect.Visible = True
        
    Dim LooPC          As Long
    
    Dim F              As Boolean
    Dim ulttick        As Long, esttick As Long
    Dim timers(1 To 2) As Long

    If AoSetup.bMover = 0 Then
        DragPantalla = False
    Else
        DragPantalla = True

    End If
    
    'Inicialización de variables globales
    prgRun = True
    pausa = False
        
    Do While prgRun
         
        Call RefreshAllChars

        Call Directx_Renderer
                                
        'TODO : Sería mejor comparar el tiempo desde la última vez que se hizo hasta el actual SOLO cuando se precisa. Además evitás el corte de intervalos con 2 golpes seguidos.
        'Sistema de timers renovado:
        esttick = GetTickCount

        For LooPC = 1 To UBound(timers)
            timers(LooPC) = timers(LooPC) + (esttick - ulttick)

            'Timer de trabajo
            If timers(1) >= tUs Then
                timers(1) = 0
                NoPuedeUsar = False

            End If

            'timer de attaque (77)
            If timers(2) >= tAt Then
                timers(2) = 0
                UserCanAttack = 1
                UserPuedeRefrescar = True

            End If

        Next LooPC

        ulttick = GetTickCount
        
        DoEvents
    Loop
    
    Call CloseClient
    
End Sub

Private Sub LoadSoundConfig()
   VOLUMEN_FX = GetVar(DirConfiguracion & "Opciones.opc", "CONFIG", "Vol_fx")
    VOLUMEN_MUSICA = GetVar(DirConfiguracion & "Opciones.opc", "CONFIG", "Vol_music")
     
    Audio.SoundVolume = (10 ^ ((VOLUMEN_FX + 900) / 1000 + 1))
    Audio.MusicVolume = (VOLUMEN_MUSICA)
End Sub

Private Sub LoadInitialConfig()

    frmCargando.Show
    frmCargando.Refresh
    
    frmConnect.version = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    
    '###########
    ' CONSTANTES
    Call AddtoRichTextBox(frmCargando.status, "Iniciando constantes...", 0, 0, 0, 0, 0, 1)
    
    Call InicializarNombres
    UserMap = 1
                  
    Call AddtoRichTextBox(frmCargando.status, "Hecho", , , , 1)
    
    '#######
    ' CLASES
    Call AddtoRichTextBox(frmCargando.status, "Instanciando clases... ", 255, 255, 255, True, False, True)
   
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    '##############
    ' MOTOR GRÁFICO
    Call AddtoRichTextBox(frmCargando.status, "Iniciando motor gráfico... ", 255, 255, 255, True, False, True)
    
    If Not InitTileEngine(frmMain.hwnd, 32, 32, Round(frmMain.MainViewPic.Height / 32), Round(frmMain.MainViewPic.Width / 32)) Then
        Call CloseClient

    End If

    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    '###################
    ' ANIMACIONES EXTRAS
    Call AddtoRichTextBox(frmCargando.status, "Creando animaciones extra... ", 255, 255, 255, True, False, True)
    
    Call CargarParticulas
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call Load_Quest
    
    If Not Directx_Initialize(0) Then
        Call CloseClient

    End If
    
    Call CargarColores
    Call Ambient_LoadColor
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    'Inicializamos el sonido
    Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectSound....", 255, 255, 255, True, False, True)
    Call Audio.Initialize(frmMain.hwnd, DirRecursos, DirRecursos)
 
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(frmMain.picInv, MAX_INVENTORY_SLOTS)
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.status, "                    ¡Bienvenido a AOMania!", 255, 255, 255, True, False, True)
          
    'Give the user enough time to read the welcome text
    Call Sleep(500)
    Unload frmCargando

End Sub

Public Sub CloseClient()

    ' Allow new instances of the client to be opened
    Call PrevInstance.ReleaseInstance

    EngineRun = False
    frmCargando.Show
    AddtoRichTextBox frmCargando.status, "Liberando recursos...", 0, 0, 0, 0, 0, 1

    Call Resolution.ResetResolution

    'Stop tile engine
    Call Directx_DeInitialize

    'Destruimos los objetos públicos creados
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    
    Call UnloadAllForms
    
    End
    
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
    '*****************************************************************
    'Writes a var to a text file
    '*****************************************************************
    writeprivateprofilestring Main, Var, Value, File

End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
    '*****************************************************************
    'Gets a Var from a text file
    '*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean

    On Error GoTo errHnd

    Dim lPos As Long
    Dim lX   As Long
    Dim iAsc As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")

    If (lPos <> 0) Then

        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1

            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))

                If Not CMSValidateChar_(iAsc) Then Exit Function

            End If

        Next lX
        
        'Finale
        CheckMailString = True

    End If

errHnd:

End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean

    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or ( _
        iAsc = 46)

End Function

'TODO : como todo lorelativo a mapas, no tiene anda que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean

    HayAgua = MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520 And MapData(X, Y).Graphic(2).GrhIndex = 0

End Function
    
Public Sub LeerLineaComandos()

    Dim T() As String
    Dim i   As Long
    
    'Parseo los comandos
    T = Split(Command, " ")
    
    For i = LBound(T) To UBound(T)

        Select Case UCase$(T(i))

            Case "/NORES" 'no cambiar la resolucion
                NoRes = True

        End Select

    Next i

End Sub

Public Sub LoadClientSetup()

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/27/2005
    '
    '**************************************************************
    Dim fHandle As Integer
    
    DoEvents
    fHandle = FreeFile
    Open App.Path & "\AOM.cfg" For Binary As fHandle
    Get fHandle, , AoSetup
    Close fHandle
    DoEvents
    
    If AoSetup.bMusica = 0 Then
        Audio.MusicActivated = False
    Else
        Audio.MusicActivated = True

    End If
    
    If AoSetup.bSonido = 0 Then
        Audio.SoundActivated = False
    Else
        Audio.SoundActivated = True

    End If
    
    If AoSetup.bPajaritos = 0 Then
        SoundPajaritos = False
    Else
        SoundPajaritos = True

    End If

End Sub

Private Sub InicializarNombres()

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/27/2005
    'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
    '**************************************************************
    Ciudades(1) = "Ullathorpe"
    Ciudades(2) = "Nix"
    Ciudades(3) = "Banderbill"

    CityDesc(1) = _
        "Ullathorpe está establecida en el medio de los grandes bosques de AOMania, es principalmente un pueblo de campesinos y leñadores. Su ubicación hace de Ullathorpe un punto de paso obligado para todos los aventureros ya que se encuentra cerca de los lugares más legendarios de este mundo."
    CityDesc(2) = "Nix es una gran ciudad. Edificada sobre la costa oeste del principal continente de AOMania."
    CityDesc(3) = "Banderbill se encuentra al norte de Ullathorpe y Nix, es una de las ciudades más importantes de todo el imperio."

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

    SkillsNames(Skills.Suerte) = "Suerte"
    SkillsNames(Skills.Magia) = "Magia"
    SkillsNames(Skills.Robar) = "Robar"
    SkillsNames(Skills.Tacticas) = "Tacticas de combate"
    SkillsNames(Skills.Armas) = "Combate con armas"
    SkillsNames(Skills.Meditar) = "Meditar"
    SkillsNames(Skills.Apuñalar) = "Apuñalar"
    SkillsNames(Skills.Ocultarse) = "Ocultarse"
    SkillsNames(Skills.Supervivencia) = "Supervivencia"
    SkillsNames(Skills.Talar) = "Talar árboles"
    SkillsNames(Skills.Comerciar) = "Comercio"
    SkillsNames(Skills.Defensa) = "Defensa con escudos"
    SkillsNames(Skills.Resistencia) = "Resistencia Magica"
    SkillsNames(Skills.Pesca) = "Pesca"
    SkillsNames(Skills.Mineria) = "Mineria"
    SkillsNames(Skills.Carpinteria) = "Carpinteria"
    SkillsNames(Skills.Herreria) = "Herreria"
    SkillsNames(Skills.Liderazgo) = "Liderazgo"
    SkillsNames(Skills.Domar) = "Domar animales"
    SkillsNames(Skills.Proyectiles) = "Armas de proyectiles"
    SkillsNames(Skills.Wresterling) = "Wresterling"
    SkillsNames(Skills.Navegacion) = "Navegacion"
    SkillsNames(Skills.Sastreria) = "Sastreria"
    SkillsNames(Skills.Recolectar) = "Recolectar hierba"
    SkillsNames(Skills.Hechiceria) = "Hechiceria"
    SkillsNames(Skills.Herrero) = "Herrero Mágico"

    AtributosNames(1) = "Fuerza"
    AtributosNames(2) = "Agilidad"
    AtributosNames(3) = "Inteligencia"
    AtributosNames(4) = "Carisma"
    AtributosNames(5) = "Constitucion"

End Sub

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long

    '*****************************************************************
    'Gets the number of fields in a delimited string
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 07/29/2007
    '*****************************************************************
    Dim Count     As Long
    Dim curPos    As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count

End Function

Public Sub ForeColorToNivel(ByVal Nivel As Byte)

    Dim tColor As Long

    Select Case Nivel

        Case Is < 55
            tColor = RGB(255, 255, 255)

        Case Else
            tColor = RGB(0, 51, 255)

    End Select

    frmMain.LvlLbl.ForeColor = tColor

End Sub

Public Function LONGTORGBDX8(ByVal lColor As Long) As Long

    Dim r As Long, B As Long, G As Long
    
    'Convert LONG to RGB:
    B = lColor \ 65536
    G = (lColor - B * 65536) \ 256
    r = lColor - B * 65536 - G * 256
    
    LONGTORGBDX8 = D3DColorXRGB(r, G, B)

End Function

Public Sub ActualizarShpUserPos()
    
    'frmMain.PicMiniMapa.Cls
    'frmMain.PicMiniMapa.PSet (UserPos.X, UserPos.Y), vbRed
    
    frmMain.MiniUserPos.Left = UserPos.X - 2
    frmMain.MiniUserPos.Top = UserPos.Y - 2

End Sub

Public Sub ActualizarShpClanPos()
    Dim i As Integer
    
    For i = 1 To 10
        If ClanPos(i).X > 0 And ClanPos(i).Y > 0 Then
            frmMain.UserClanPos(i).Left = ClanPos(i).X - 2
            frmMain.UserClanPos(i).Top = ClanPos(i).Y - 2
            frmMain.UserClanPos(i).Visible = True
            Else
            frmMain.UserClanPos(i).Visible = False
        End If
    Next i

End Sub

Public Sub CargarMiniMapa()
    Dim Data() As Byte
    Dim picMapa As Picture
    
    'Get Picture
    If Get_File_Data(DirRecursos, "Mapa" & UserMap & ".JPG", Data, MINIMAPA_FILE) Then
        Set MiniMapa.Mapa = ArrayToPicture(Data(), 0, UBound(Data) + 1)
        Set picMapa = MiniMapa.Mapa
    Else
        If Get_File_Data(DirRecursos, "Mapa0.JPG", Data, MINIMAPA_FILE) Then
            Set MiniMapa.Mapa = ArrayToPicture(Data(), 0, UBound(Data) + 1)
            Set picMapa = MiniMapa.Mapa
        End If
    End If
        
         

    ' If FileExist(DirMiniMapa & "Mapa" & UserMap & ".BMP", vbArchive) Then
    '     Set picMapa = LoadPicture(DirMiniMapa & "Mapa" & UserMap & ".BMP")
    ' Else
    '     Set picMapa = LoadPicture(DirMiniMapa & "Mapa0.BMP")

    ' End If

    frmMain.picMiniMap.Picture = picMapa
    
    Set picMapa = Nothing

End Sub

Sub Disco()

    Dim Fso As New Scripting.FileSystemObject
    Dim DR  As Scripting.Drive
    Set DR = Fso.GetDrive("c:")
    HDD = Abs(DR.SerialNumber)

End Sub


Public Function getTagPosition(ByVal Nick As String) As Integer
Dim buf As Integer
    buf = InStr(Nick, "<")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
    buf = InStr(Nick, "[")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
     getTagPosition = Len(Nick) + 2
End Function

Sub DayNameChange(ByVal Hora As Byte)

        If Hora >= 1 And Hora <= 7 Then NameDay = "Noche"
        If Hora >= 8 And Hora <= 12 Then NameDay = "Amanecer"
        If Hora >= 13 And Hora <= 18 Then NameDay = "Día"
        If Hora >= 19 And Hora <= 21 Then NameDay = "Tarde"
        If Hora >= 22 And Hora <= 24 Then NameDay = "Noche"
       
End Sub

Sub LoadDataLauncher()
       
       PlayLauncher = Val(GetVar(DirConfiguracion & "Launcher.dat", "CONFIG", "Play"))
       
End Sub

Sub SaveDataLauncher()
      
      Call WriteVar(DirConfiguracion & "Launcher.dat", "CONFIG", "Play", PlayLauncher)
      
End Sub

Sub EnviaM(sentido As Byte)
    Call SendData("Ñ" & sentido & CodigoCorreccion & "*" & CharList(UserCharIndex).pos.X & "*" & CharList(UserCharIndex).pos.Y)
End Sub

Sub RefreshAllChars()
   '*****************************************************************
   'Goes through the charlist and replots all the characters on the map
   'Used to make sure everyone is visible
   '*****************************************************************
   
   Dim LooPC As Integer
   
   For LooPC = 1 To LastChar
      If CharList(LooPC).Active = 1 Then
         MapData(CharList(LooPC).pos.X, CharList(LooPC).pos.Y).charindex = LooPC
      End If
   Next LooPC
   
End Sub
