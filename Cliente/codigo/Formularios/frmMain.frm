VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AoMania 2019"
   ClientHeight    =   8970
   ClientLeft      =   1950
   ClientTop       =   1815
   ClientWidth     =   5190
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8970
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Frame Multiplicadores 
      Caption         =   "Multiplicadores"
      Height          =   1455
      Left            =   120
      TabIndex        =   19
      Top             =   7200
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CommandButton BotomM 
         Caption         =   "Ok"
         Height          =   360
         Left            =   1200
         TabIndex        =   26
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox CantMTexto 
         Height          =   315
         Left            =   650
         TabIndex        =   25
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox CantMOro 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   600
         TabIndex        =   24
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox CantMExp 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   600
         TabIndex        =   23
         Top             =   240
         Width           =   495
      End
      Begin VB.Label LabelMTexto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto:"
         Height          =   210
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   510
      End
      Begin VB.Label LabelOro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Oro:"
         Height          =   210
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   345
      End
      Begin VB.Label LabelMExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp:"
         Height          =   210
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   330
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Actualizar Lista"
      Height          =   360
      Left            =   1680
      TabIndex        =   17
      Top             =   6720
      Width           =   1830
   End
   Begin VB.CommandButton Command5 
      Caption         =   "User Hablan"
      Height          =   360
      Left            =   240
      TabIndex        =   16
      Top             =   6720
      Width           =   1350
   End
   Begin VB.Frame Frame3 
      Caption         =   "Gm's Conectados"
      Height          =   1815
      Left            =   120
      TabIndex        =   14
      Top             =   4800
      Width           =   4935
      Begin VB.ListBox Gms 
         Height          =   1320
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Timer bandaymedusa 
      Interval        =   60000
      Left            =   0
      Top             =   1800
   End
   Begin VB.Timer deat 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   0
      Top             =   2160
   End
   Begin VB.Timer Mascotas 
      Interval        =   60000
      Left            =   4680
      Top             =   2160
   End
   Begin VB.Timer torneos 
      Interval        =   60000
      Left            =   4680
      Top             =   3240
   End
   Begin VB.CheckBox SUPERLOG 
      Caption         =   "log"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton CMDDUMP 
      Caption         =   "dump"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer FX 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   0
      Top             =   3600
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4680
      Top             =   2520
   End
   Begin VB.Timer GameTimer 
      Interval        =   40
      Left            =   0
      Top             =   2880
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   0
      Top             =   3240
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4680
      Top             =   1800
   End
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   0
      Top             =   1440
   End
   Begin VB.Timer KillLog 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4680
      Top             =   2880
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4680
      Top             =   3600
   End
   Begin VB.Frame Estadisticas 
      Caption         =   "Estadisticas"
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4935
      Begin VB.Timer tGranPoder 
         Interval        =   60000
         Left            =   4560
         Top             =   720
      End
      Begin VB.Timer TNosfeSagrada 
         Interval        =   1000
         Left            =   4560
         Top             =   1080
      End
      Begin VB.Label CantNumGM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número de gms: 0"
         Height          =   210
         Left            =   240
         TabIndex        =   31
         Top             =   480
         Width           =   1515
      End
      Begin VB.Label CantTimer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         Height          =   210
         Left            =   3600
         TabIndex        =   30
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label CantOnDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Días Online: 0"
         Height          =   210
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label CantOnHor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horas Online: 0"
         Height          =   210
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label CantOnMin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minutos Online: 0"
         Height          =   210
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label CantUsuarios 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número de usuarios: 0"
         Height          =   210
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1875
      End
   End
   Begin VB.ListBox User 
      Height          =   1320
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Caption         =   "BroadCast"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   4935
      Begin VB.Timer TBarcos 
         Interval        =   60000
         Left            =   -120
         Top             =   840
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Enviar SMSG a los GM's"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Enviar RMSG a los GM's"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Enviar SMSG a TODOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Enviar RMSG a TODOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ArgentuMania (Mod del Argentum Online Mod) v11.5"
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   360
      TabIndex        =   18
      Top             =   8640
      Width           =   4290
   End
   Begin VB.Label Escuch 
      Caption         =   "Label2"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "AoMania"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "Opciones"
      Begin VB.Menu mnuLluvia 
         Caption         =   "Lluvia On/Off"
      End
      Begin VB.Menu mnucSagrados 
         Caption         =   "Valor Curar Sagrados"
      End
      Begin VB.Menu mnucReyes 
         Caption         =   "Curar Reyes"
      End
      Begin VB.Menu mnuMultiplicadores 
         Caption         =   "Multiplicadores"
      End
      Begin VB.Menu mnuRMGms 
         Caption         =   "Recargar Mc Gms"
      End
      Begin VB.Menu mnuREParty 
         Caption         =   "Recargar Exp Party"
      End
      Begin VB.Menu mnuRMClan 
         Caption         =   "Recargar Miembros Clan"
      End
      Begin VB.Menu mnuRMClientes 
         Caption         =   "Recargar Md5 Clientes"
      End
      Begin VB.Menu mnuModificaciones 
         Caption         =   "Recargar Modificaciones"
      End
   End
   Begin VB.Menu mnuResetear 
      Caption         =   "Resetear"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA

    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64

End Type

Const NIM_ADD = 0
Const NIM_MODIFY = 1
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, Id As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA

    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = Id
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp

End Function

'Sub CheckIdleUser()
'    Dim iUserIndex As Long
'
'    For iUserIndex = 1 To MaxUsers
'
'        With UserList(iUserIndex)
'
'            'Conexion activa? y es un usuario loggeado?
'            If .ConnID <> -1 And .flags.UserLogged Then
'
'                'Actualiza el contador de inactividad
'                .Counters.IdleCount = .Counters.IdleCount + 1
'
'                If .Counters.IdleCount >= IdleLimit Then
'                    Call SendData(SendTarget.toindex, iUserIndex, 0, _
'                            "Has sido desconectado por permanecer mas de 30 minutos inactivo.")
'
'                    'mato los comercios seguros
'                    If .ComUsu.DestUsu > 0 Then
'                        If UserList(.ComUsu.DestUsu).flags.UserLogged Then
'                            If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
'                                Call SendData(SendTarget.toindex, .ComUsu.DestUsu, 0, _
'                                        "||Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
'                                Call FinComerciarUsu(.ComUsu.DestUsu)
'
'                            End If
'
'                        End If
'
'                        Call FinComerciarUsu(iUserIndex)
'
'                    End If
'
'                    Call Cerrar_Usuario(iUserIndex)
'
'                End If
'
'            End If
'
'        End With
'
'    Next iUserIndex
'
'End Sub

Private Sub Auditoria_Timer()

    On Error GoTo errhand

    Call PasarSegundo    'sistema de desconexion de 10 segs

    Call ActualizaStatsES

    Exit Sub

errhand:
    Call LogError("Error en Timer Auditoria. Err: " & Err.Description & " - " & Err.Number)

End Sub

Private Sub AutoSave_Timer()

    On Error GoTo errhandler

    'fired every minute
    Static Minutos          As Long
    Static MinutosLatsClean As Long
    Static MinsPjesSave     As Long

    Dim i                   As Long

    Minutos = Minutos + 1
    MinsPjesSave = MinsPjesSave + 1
    MinutosLatsClean = MinutosLatsClean + 1

    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    Call ModAreas.AreasOptimizacion
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    Call SpendTime
    Call RecompensaCastillos

    'If MinsPjesSave = MinutosGuardarUsuarios - 1 Then
    'Call SendData(SendTarget.toall, 0, 0, "||CharSave en 1 minuto ..." & FONTTYPE_VENENO)
    'Else
    If MinsPjesSave >= MinutosGuardarUsuarios Then
        Call mdParty.ActualizaExperiencias
        Call GuardarUsuarios(False)
        MinsPjesSave = 0

    End If

    If Minutos = MinutosWs - 1 Then
        Call SendData(SendTarget.toall, 0, 0, "||Worldsave y Limpieza en 1 minuto ..." & FONTTYPE_VENENO)

    ElseIf Minutos >= MinutosWs Then
        Call aClon.VaciarColeccion
        Call mdParty.ActualizaExperiencias
        Call GuardarUsuarios
        'WorldSave
        Call DoBackUp
        'Call SendData(SendTarget.toall, 0, 0, _
         "||WORLDSAVE TERMINADO CORRECTAMENTE. YA PUEDES SEGUIR JUGANDO!." & FONTTYPE_SERVER)
        Minutos = 0

    End If

    If MinutosLatsClean = MinutosLimpia - 1 Then
        Call SendData(SendTarget.toall, 0, 0, "||Limpieza del mundo en 1 minuto ..." & FONTTYPE_VENENO)

    ElseIf MinutosLatsClean >= MinutosLimpia Then
        MinutosLatsClean = 0
        Call ReSpawnOrigPosNpcs    'respawn de los guardias en las pos originales
        Call LimpiarMundo

    End If

    Call PurgarPenas
    'Call CheckIdleUser

    '<<<<<-------- Log the number of users online ------>>>
    Dim n As Integer
    n = FreeFile()
    Open App.Path & "\logs\numusers.log" For Output Shared As n
    Print #n, NumUsers
    Close #n
    '<<<<<-------- Log the number of users online ------>>>

    Exit Sub
errhandler:
    Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.Description)

End Sub

Private Sub bandaymedusa_Timer()

    On Error GoTo errordm:

    Call Timer_GuerradeBanda

errordm:

End Sub

Private Sub BotomM_Click()
    Multexp = CantMExp.Text
    MultOro = CantMOro.Text
    MultMsg = CantMTexto.Text
    Call WriteVar(DatPath & "Ini\Multipli.ini", "Multiplicadores", "Exp", Multexp)
    Call WriteVar(DatPath & "Ini\Multipli.ini", "Multiplicadores", "Oro", MultOro)
    Call WriteVar(DatPath & "Ini\Multipli.ini", "Multiplicadores", "Msg", MultMsg)
    Multiplicadores.Visible = False

End Sub

Private Sub CMDDUMP_Click()

    On Error Resume Next

    Dim i As Integer

    For i = 1 To MaxUsers
        Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).Name & _
            " UserLogged: " & UserList(i).flags.UserLogged)
    Next i

    Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub Command1_Click()
    Call SendData(SendTarget.toall, 0, 0, "||AoMania> " & BroadMsg.Text & FONTTYPE_SERVER)

End Sub

Public Sub InitMain(ByVal f As Byte)

    If f = 1 Then
        Call mnuSystray_Click
    Else
        frmMain.Show

    End If

End Sub

Private Sub Command2_Click()
    Call SendData(SendTarget.toall, 0, 0, "!!" & BroadMsg.Text & ENDC)

End Sub

Private Sub Command3_Click()
    Call SendData(SendTarget.ToAdmins, 0, 0, "||AoMania> " & BroadMsg.Text & FONTTYPE_GUILD)

End Sub

Private Sub Command4_Click()
    Call SendData(SendTarget.ToAdmins, 0, 0, "!!" & BroadMsg.Text & ENDC)

End Sub

Private Sub Command5_Click()
    FrmUserhablan.Show

End Sub

Private Sub Command6_Click()
    Call TCP_HandleData2.ActUser
    Call TCP_HandleData2.ActGM

End Sub

Private Sub deat_Timer()

    On Error GoTo errordm:

    tukiql = tukiql + 1

    Select Case tukiql

        Case 53
            Call SendData(SendTarget.toall, 0, 0, "||DeathMatch> En 10 minutos se realizará un deathmatch automatico." & FONTTYPE_GUILD)

        Case 58
            Call SendData(SendTarget.toall, 0, 0, "||DeathMatch> En 5 minutos se realizará un deathmatch automatico." & FONTTYPE_GUILD)

        Case 62
            Call SendData(SendTarget.toall, 0, 0, "||DeathMatch> En 1 minutos se realizará un deathmatch automatico." & FONTTYPE_GUILD)

        Case 63
            Call death_comienza(RandomNumber(8, 16))

        Case 65

            If deathesp = True Then
                Call Deathauto_Cancela
                tukiql = 2
            Else
                tukiql = 2

            End If

    End Select

errordm:

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next

    If Not Visible Then

        Select Case X \ Screen.TwipsPerPixelX

            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess

            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp

                If hHook Then
                    UnhookWindowsHookEx hHook
                    hHook = 0
                End If


        End Select

    End If

End Sub

Private Sub QuitarIconoSystray()

    On Error Resume Next

    'Borramos el icono del systray
    Dim i   As Integer
    Dim nid As NOTIFYICONDATA

    nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

    i = Shell_NotifyIconA(NIM_DELETE, nid)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
#If MYSQL = 1 Then
    Call Add_DataBase("0", "Status")
#End If
End Sub

Private Sub Form_Terminate()
#If MYSQL = 1 Then
    Call Add_DataBase("0", "Status")
#End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next

    Call QuitarIconoSystray

    Call LimpiaWsApi(frmMain.hWnd)

    Dim loopc As Integer

    For loopc = 1 To MaxUsers

        If UserList(loopc).ConnID <> -1 Then Call CloseSocket(loopc)
    Next

    'Log
    Dim n As Integer
    n = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #n
    Print #n, Date & " " & Time & " server cerrado."
    Close #n

    End

    Set SonidosMapas = Nothing
    
#If MYSQL = 1 Then
    Call Add_DataBase("0", "Status")
#End If

End Sub

Private Sub FX_Timer()

    On Error GoTo hayerror

    Call SonidosMapas.ReproducirSonidosDeMapas

    Exit Sub
hayerror:

End Sub

Private Sub GameTimer_Timer()

    Dim iUserIndex   As Long
    Dim bEnviarStats As Boolean
    Dim bEnviarAyS   As Boolean

    On Error GoTo hayerror

    '<<<<<< Procesa eventos de los usuarios >>>>>>
    For iUserIndex = 1 To LastUser

        With UserList(iUserIndex)

            'Conexion activa?
            If .ConnID <> -1 Then

                '¿User valido?
                If .ConnIDValida And .flags.UserLogged Then

                    '[Alejo-18-5]
                    bEnviarStats = False
                    bEnviarAyS = False

                    Call DoTileEvents(iUserIndex, .pos.Map, .pos.X, .pos.Y)

                    If .flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
                    If .flags.Ceguera = 1 Or .flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)

                    If .flags.Muerto = 0 Then

                        '[Consejeros]
                        If .flags.Desnudo <> 0 And (.flags.Privilegios = PlayerType.User) <> 0 Then Call EfectoFrio(iUserIndex)

                        If .flags.Meditando Then Call DoMeditar(iUserIndex)

                        If .flags.Envenenado <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoVeneno(iUserIndex, bEnviarStats)

                        If .flags.AdminInvisible <> 1 Then
                            If .flags.Invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
                            If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)

                        End If

                        If .flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)

                        Call DuracionPociones(iUserIndex)

                        Call HambreYSed(iUserIndex, bEnviarAyS)

                        If .flags.Hambre = 0 And .flags.Sed = 0 Then
                            If SecondaryWeather Then
                                If Not Intemperie(iUserIndex) Then

                                    If Not .flags.Descansar Then
                                        'No esta descansando

                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)

                                        If bEnviarStats Then
                                            Call SendData(SendTarget.toIndex, iUserIndex, 0, "ASH" & .Stats.MinHP)
                                            bEnviarStats = False

                                        End If

                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)

                                        If bEnviarStats Then
                                            Call SendData(SendTarget.toIndex, iUserIndex, 0, "ASS" & .Stats.MinSta)
                                            bEnviarStats = False

                                        End If

                                    Else
                                        'esta descansando

                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)

                                        If bEnviarStats Then
                                            Call SendData(SendTarget.toIndex, iUserIndex, 0, "ASH" & .Stats.MinHP)
                                            bEnviarStats = False

                                        End If

                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)

                                        If bEnviarStats Then
                                            Call SendData(SendTarget.toIndex, iUserIndex, 0, "ASS" & .Stats.MinSta)
                                            bEnviarStats = False

                                        End If

                                        'termina de descansar automaticamente
                                        If .Stats.MaxHP = .Stats.MinHP And .Stats.MaxSta = .Stats.MinSta Then
                                            Call SendData(SendTarget.toIndex, iUserIndex, 0, "DOK")
                                            Call SendData(SendTarget.toIndex, iUserIndex, 0, "||Has terminado de descansar." & FONTTYPE_INFO)
                                            .flags.Descansar = False

                                        End If

                                    End If

                                End If

                            Else

                                If Not .flags.Descansar Then
                                    'No esta descansando

                                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)

                                    If bEnviarStats Then
                                        Call SendData(SendTarget.toIndex, iUserIndex, 0, "ASH" & .Stats.MinHP)
                                        bEnviarStats = False

                                    End If

                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)

                                    If bEnviarStats Then
                                        Call SendData(SendTarget.toIndex, iUserIndex, 0, "ASS" & .Stats.MinSta)
                                        bEnviarStats = False

                                    End If

                                Else
                                    'esta descansando

                                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)

                                    If bEnviarStats Then
                                        Call SendData(SendTarget.toIndex, iUserIndex, 0, "ASH" & .Stats.MinHP)
                                        bEnviarStats = False

                                    End If

                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)

                                    If bEnviarStats Then
                                        Call SendData(SendTarget.toIndex, iUserIndex, 0, "ASS" & .Stats.MinSta)
                                        bEnviarStats = False

                                    End If

                                    'termina de descansar automaticamente
                                    If .Stats.MaxHP = .Stats.MinHP And .Stats.MaxSta = .Stats.MinSta Then
                                        Call SendData(SendTarget.toIndex, iUserIndex, 0, "DOK")
                                        Call SendData(SendTarget.toIndex, iUserIndex, 0, "||Has terminado de descansar." & FONTTYPE_INFO)
                                        .flags.Descansar = False

                                    End If

                                End If    'Not .Flags.Descansar And (.Flags.Hambre = 0 And .Flags.Sed = 0)

                            End If

                        End If

                        If bEnviarAyS Then Call EnviarHambreYsed(iUserIndex)

                        If .NroMacotas > 0 Then Call TiempoInvocacion(iUserIndex)

                    End If    'Muerto

                    'Else    'no esta logeado?
                    '.Counters.IdleCount = 0
                    '[Gonzalo]: deshabilitado para el nuevo sistema de tiraje
                    'de dados :)

                    ' .Counters.IdleCount = .Counters.IdleCount + 1

                    'If .Counters.IdleCount > IntervaloParaConexion Then
                    '.Counters.IdleCount = 0
                    'Call CloseSocket(iUserIndex)

                    'End If

                End If    'UserLogged

            End If
            
            If .GuildIndex > 0 Then
                If .Clan.Timer > 0 Then
                    Call modGuilds.TimerPuntosClan(iUserIndex)
                End If
            End If
            
        End With

    Next iUserIndex

    Exit Sub

hayerror:
    LogError ("Error en GameTimer: " & Err.Description & " UserIndex = " & iUserIndex)

End Sub

Private Sub Mascotas_Timer()

    Dim Npc1    As Integer
    Dim Npc1Pos As WorldPos
    Npc1 = RandomNumber(924, 939)
    Npc1Pos.Map = 30
    Npc1Pos.X = 61
    Npc1Pos.Y = 38

    mariano = mariano + 1

    Select Case mariano

        Case 475
            Call SendData(SendTarget.toall, 0, 0, "||AoMania> En 5 minutos se invocara un Domador." & FONTTYPE_GUILD)

        Case 479
            Call SendData(SendTarget.toall, 0, 0, "||AoMania> En 1 minuto se invocara un Domador." & FONTTYPE_GUILD)

        Case 480
            Call SendData(SendTarget.toall, 0, 0, "||AoMania> Se ha invocado un domador en el mapa 30." & FONTTYPE_GUILD)
            Call SendData(SendTarget.toall, 0, 0, "TW122")
            Call SpawnNpc(val(Npc1), Npc1Pos, True, False)
            mariano = 0

    End Select

End Sub

Private Sub mnuCerrar_Click()

    If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
        Dim f

        For Each f In Forms

            Unload f
        Next

    End If

End Sub

Private Sub mnuLluvia_Click()

    Call SecondaryAmbient
    Exit Sub

End Sub

Private Sub mnuMultiplicadores_Click()
    Multiplicadores.Visible = True
    CantMExp.Text = Multexp
    CantMOro.Text = MultOro
    CantMTexto.Text = MultMsg

End Sub

Private Sub mnusalir_Click()

    Call mnuCerrar_Click

End Sub

Public Sub mnuMostrar_Click()

    On Error Resume Next

    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0

End Sub

Private Sub KillLog_Timer()

    On Error Resume Next

    If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
    If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
    If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
    If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
    If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
    If Not FileExist(App.Path & "\logs\nokillwsapi.txt") Then
        If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"

    End If

End Sub

Private Sub mnuServidor_Click()

    frmServidor.Visible = True

End Sub

Private Sub mnuSystray_Click()

    Dim i   As Integer
    Dim s   As String
    Dim nid As NOTIFYICONDATA

    s = "ARGENTUM-ONLINE"
    nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, s)
    i = Shell_NotifyIconA(NIM_ADD, nid)

    If WindowState <> vbMinimized Then WindowState = vbMinimized
    Visible = False

End Sub

Private Sub mnuResetear_Click()

    If MsgBox("¿Estás seguro de que quieres resetear el servidor?", vbYesNo) = vbYes Then
        Call Resetear
    Else
        Exit Sub

    End If

End Sub

Private Sub Resetear()
    Dim CharFile       As String
    Dim Guilds         As String
    Dim Logs           As String
    Dim Logsa          As String
    Dim Gms            As String
    Dim Users          As String
    Dim ShowSos        As String
    Dim ShowDenuncia   As String
    Dim ShowBug        As String
    Dim ShowSugerencia As String
    Dim Consultas      As String
    Dim Restard        As Double
  
    CharFile = FileExist(App.Path & "\Charfile\*.*", vbNormal)
    Guilds = FileExist(App.Path & "\Guilds\*.*", vbNormal)
    Logs = FileExist(App.Path & "\Logs\*.*", vbNormal)
    Logsa = FileExist(App.Path & "\Logsa\*.*", vbNormal)
    Gms = FileExist(App.Path & "\Logs\Gms\*.*", vbNormal)
    Users = FileExist(App.Path & "\Logs\Usuarios\*.*", vbNormal)
    ShowSos = FileExist(App.Path & "\Logs\Show\Sos\*.*", vbNormal)
    ShowDenuncia = FileExist(App.Path & "\Logs\Show\Denuncia\*.*", vbNormal)
    ShowBug = FileExist(App.Path & "\Logs\Show\Bug\*.*", vbNormal)
    ShowSugerencia = FileExist(App.Path & "\Logs\Show\Sugerencia\*.*", vbNormal)
    Consultas = FileExist(App.Path & "\Logs\Consultas\*.*", vbNormal)
  
    If CharFile = True Then
        Kill (App.Path & "\Charfile\*.*")
    Else

    End If
  
    If Guilds = True Then
        Kill (App.Path & "\Guilds\*.*")
    Else

    End If
  
    If Logs = True Then
        Kill (App.Path & "\Logs\*.*")
    Else

    End If
  
    If Logsa = True Then
        Kill (App.Path & "\Logsa\*.*")
    Else

    End If
  
    If Gms = True Then
        Kill (App.Path & "\Logs\Gms\*.*")
    Else

    End If
  
    If Users = True Then
        Kill (App.Path & "\Logs\Usuarios\*.*")
    Else

    End If
  
    If ShowSos = True Then
        Kill (App.Path & "\Logs\Show\Sos\*.*")
    Else

    End If
  
    If ShowDenuncia = True Then
        Kill (App.Path & "\Logs\Show\Denuncia\*.*")
    Else

    End If
  
    If ShowBug = True Then
        Kill (App.Path & "\Logs\Show\Bug\*.*")
    Else

    End If
  
    If ShowSugerencia = True Then
        Kill (App.Path & "\Logs\Show\Sugerencia\*.*")
    Else

    End If
  
    Call WriteVar(App.Path & "\Server.Ini", "INIT", "Record", "0")
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Este", vbNullString)
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraEste", vbNullString)
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Norte", vbNullString)
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraNorte", vbNullString)
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Oeste", vbNullString)
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraOeste", vbNullString)
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Sur", vbNullString)
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraSur", vbNullString)
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Fortaleza", vbNullString)
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraForta", vbNullString)
    Call WriteVar(App.Path & "\Dat\Ini\Config.ini", "NOTOKAR", "NOPSD", "0")
    Call WriteVar(App.Path & "\Dat\Ini\Config.ini", "NOTOKAR", "USUARIO", vbNullString)
    Call WriteVar(DirINI & "Ranking.ini", "Ranking", "NvMaxCiudadano", "0")
    Call WriteVar(DirINI & "Ranking.ini", "Ranking", "MaxCiudadano", vbNullString)
    Call WriteVar(DirINI & "Ranking.ini", "Ranking", "NvMaxCriminal", "0")
    Call WriteVar(DirINI & "Ranking.ini", "Ranking", "MaxCriminal", vbNullString)
    Call WriteVar(DirINI & "Ranking.ini", "Ranking", "MonedaOroMax", "0")
    Call WriteVar(DirINI & "Ranking.ini", "Ranking", "MaxOro", vbNullString)
    
#If MYSQL = 1 Then
    Call Reset_Mysql
#End If
    
    Restard = Shell(App.Path & "\AoM.exe", vbNormalFocus)
    
    End
  
End Sub

Private Sub npcataca_Timer()
    Dim npc As Integer

    For npc = 1 To LastNPC
        Npclist(npc).CanAttack = 1
    Next npc

End Sub

Private Sub TBarcos_Timer()

    On Error Resume Next

    Barcos.TiempoRest = Barcos.TiempoRest - 1

    If Barcos.TiempoRest < 11 And Barcos.TiempoRest > 0 Then
        Call SendData(SendTarget.toall, 0, 0, "||Les anunciamos a todos los viajantes a " & Zonas(Barcos.Zona).nombre & " que queda " & _
            Barcos.TiempoRest & " minutos antes de sarpar." & FONTTYPE_INFO)

    End If

    If Barcos.TiempoRest = 0 Then
        Call SendData(SendTarget.toall, 0, 0, "||La embarcacion a " & Zonas(Barcos.Zona).nombre & " ya a partido." & FONTTYPE_INFO)

    End If

    If Barcos.TiempoRest <= TIEMPO_LLEGADA Then
        Barcos.TiempoRest = 60

        If NumZonas > 0 Then
            Call SendData(SendTarget.toall, 0, 0, "||La embarcacion a " & Zonas(Barcos.Zona).nombre & _
                " ya a llegado. En 1 hora partira la proxima embarcacion a " & Zonas(Barcos.Zona + IIf((Barcos.Zona >= NumZonas), -(NumZonas), _
                1)).nombre & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toall, 0, 0, "||La embarcacion a " & Zonas(Barcos.Zona).nombre & _
                " ya a llegado. En 1 hora partira la proxima embarcacion a " & Zonas(Barcos.Zona).nombre & FONTTYPE_INFO)

        End If

        Dim loopc  As Integer
        Dim PosNNN As WorldPos
        PosNNN.Map = Zonas(Barcos.Zona).Map
        PosNNN.Y = Zonas(Barcos.Zona).Y
        PosNNN.X = Zonas(Barcos.Zona).X
        Barcos.Pasajeros = 0

        For loopc = 1 To LastUser

            If UserList(loopc).flags.Embarcado = 1 Then
                Call WarpUserChar(loopc, PosNNN.Map, PosNNN.X, PosNNN.Y, False)
                UserList(loopc).flags.Embarcado = 0

            End If

        Next loopc

        Barcos.Zona = Barcos.Zona + 1

        If Barcos.Zona > NumZonas Then Barcos.Zona = 0

    End If

End Sub

Private Sub tGranPoder_Timer()

    GranPoder.Timer = GranPoder.Timer + 1
    
    If GranPoder.Timer >= 5 Then
        GranPoder.Timer = 0
        If GranPoder.Status = 0 Then
            Call ActSlot
            Call mod_GranPoder.DarGranPoder(0)
        Else
            Call SendData(SendTarget.toall, 0, 0, "||" & UserList(GranPoder.UserIndex).Name & " poseedor del Aura de los Heroes!!!. En el mapa " & UserList(GranPoder.UserIndex).pos.Map & "." & FONTTYPE_GUERRA)
            Call SendData(SendTarget.ToPCArea, GranPoder.UserIndex, UserList(GranPoder.UserIndex).pos.Map, "CFX" & UserList(GranPoder.UserIndex).char.CharIndex & "," & FX_Poder & "," & 1)
            Call SendData(SendTarget.toIndex, GranPoder.UserIndex, 0, "TW" & Sound_Poder)
        End If
    End If
    
    
End Sub

Private Sub TIMER_AI_Timer()

    On Error GoTo ErrorHandler

    Dim NpcIndex As Integer
    Dim X        As Integer
    Dim Y        As Integer
    Dim UseAI    As Integer
    Dim Mapa     As Integer
    Dim e_p      As Integer

    'Barrin 29/9/03
    If Not haciendoBK And Not EnPausa Then

        'Update NPCs
        For NpcIndex = 1 To LastNPC

            If Npclist(NpcIndex).flags.NPCActive Then    'Nos aseguramos que sea INTELIGENTE!
                e_p = esPretoriano(NpcIndex)

                If e_p > 0 Then
                    If Npclist(NpcIndex).flags.Paralizado = 1 Then Call EfectoParalisisNpc(NpcIndex)

                    Select Case e_p

                        Case 1  ''clerigo
                            Call PRCLER_AI(NpcIndex)

                        Case 2  ''mago
                            Call PRMAGO_AI(NpcIndex)

                        Case 3  ''cazador
                            Call PRCAZA_AI(NpcIndex)

                        Case 4  ''rey
                            Call PRREY_AI(NpcIndex)

                        Case 5  ''guerre
                            Call PRGUER_AI(NpcIndex)

                    End Select

                Else
                    Call CuraRey(NpcIndex)

                    ''ia comun
                    If Npclist(NpcIndex).flags.Paralizado = 1 Then
                        Call EfectoParalisisNpc(NpcIndex)
                    Else

                        'Usamos AI si hay algun user en el mapa
                        If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
                            Call EfectoParalisisNpc(NpcIndex)

                        End If

                        Mapa = Npclist(NpcIndex).pos.Map

                        If Mapa > 0 Then
                            If MapInfo(Mapa).NumUsers > 0 Then
                                If Npclist(NpcIndex).Movement <> TipoAI.ESTATICO Then
                                    Call NPCAI(NpcIndex)

                                End If

                            End If

                        End If

                    End If

                End If

            End If

        Next NpcIndex

    End If

    Exit Sub

ErrorHandler:
    Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).pos.Map)
    Call MuereNpc(NpcIndex, 0)

End Sub

Private Sub TNosfeSagrada_Timer()
    Call TCP_HandleData2.RegUser
    
    Call TCP_HandleData2.RegGM
    frmMain.CantTimer.caption = Time
    
    Static MinutesOn As Long
    
    MinutesOn = MinutesOn + 1
    
    If MinutesOn = 60 Then
      
        If OnMin <= 60 Then
         
            If OnMin < 60 Then
         
                OnMin = OnMin + 1
             
                If OnMin = 60 Then
               
                    If OnHor < 24 Then
              
                        OnHor = OnHor + 1

                        If OnHor = 24 Then
                  
                            OnDay = OnDay + 1
                  
                            OnHor = 0
                  
                        End If
               
                    End If
              
                    OnMin = 0
             
                End If
         
            End If
         
        End If
      
        MinutesOn = 0
        Call MostrarTimeOnline

    End If
   
    'Sistema de Criaturas
    Call SistemaCriatura.Timer_SistemaCriatura
    
    'Comienzo timer Nosfe
    Static Minutos As Long
    
    Minutos = Minutos + 1
        
    If MataNosfe = True Then
        
        Call SendData(SendTarget.toall, 0, 0, "||El usuario " & UserList(NickMataNosfe).Name & " ha matado al Nosferatu." & FONTTYPE_GUILD)
        
        Call SendData(SendTarget.toIndex, NickMataNosfe, 0, "||Has ganado " & ExpMataNosfe & " experencia extra por matar a Nosferatu." & _
            FONTTYPE_FIGHT)
        
        UserList(NickMataNosfe).Stats.Exp = UserList(NickMataNosfe).Stats.Exp + ExpMataNosfe
                 
        Call CheckUserLevel(NickMataNosfe)
        Call EnviarExp(NickMataNosfe)
        
        Minutos = 0
        StatusNosfe = False
        MataNosfe = False
        
    End If
        
    If StatusNosfe = True Then
        If Minutos > IntervaloMsjNosfe Then
            Call SendData(SendTarget.toall, 0, 0, "||Nosferatu esta haciendo estragos en el mapa" & MapaNosfe & FONTTYPE_GUILD)
            Minutos = 0

        End If

    End If
    
    If StatusNosfe = False Then
        
        If RepiteInvoNosfe = True Then
           
            If Minutos > 10 Then
                Call InvocaNosfe
                Minutos = 0

            End If
          
        End If
        
        If AvisoNosfe = True Then
           
            If Minutos = TimeAvisoNosfe Then
                Call SendData(SendTarget.toall, 0, 0, "||El Nosferatu saldra en 10 Minutos." & FONTTYPE_GUILD)

            End If
           
        End If
        
        If Minutos > IntervaloNosfe Then
            Call InvocaNosfe
            Minutos = 0

        End If
       
    End If
    
    'Finish Time Nosfe
    
    'Comienzo Time Sagradas
    Static SecYetiOscura As Long
    
    SecYetiOscura = SecYetiOscura + 1
    
    If MataYetiOscura = True Then
         
        Call SendData(SendTarget.toall, 0, 0, "||El Yeti Sagrado Oscuro Regreso al otro Mundo." & FONTTYPE_GUILD)
         
        StatusYetiOscura = False
        SecYetiOscura = 0
        MataYetiOscura = False

    End If
    
    If StatusYetiOscura = False Then
          
        If RepiteInvoYetiOscura = True Then
              
            If SecYetiOscura > 10 Then
                Call SpawnSagrada("YetiOscura")
                SecYetiOscura = 0

            End If
              
        End If
         
        If SecYetiOscura > IntervaloSagrada Then
            Call SpawnSagrada("YetiOscura")
            SecYetiOscura = 0

        End If

    End If
    
    Static SecYeti As Long
    
    SecYeti = SecYeti + 1
    
    If MataYeti = True Then
         
        Call SendData(SendTarget.toall, 0, 0, "||El Yeti Sagrado Regreso al otro Mundo." & FONTTYPE_GUILD)
         
        StatusYeti = False
        SecYeti = 0
        MataYeti = False

    End If
    
    If StatusYeti = False Then
          
        If RepiteInvoYeti = True Then
              
            If SecYeti > 10 Then
                Call SpawnSagrada("Yeti")
                SecYetiOscura = 0

            End If
              
        End If
         
        If SecYeti > IntervaloSagrada Then
            Call SpawnSagrada("Yeti")
            SecYeti = 0

        End If

    End If
    
    Static SecCleopatra As Long
    
    SecCleopatra = SecCleopatra + 1
    
    If MataCleopatra = True Then
         
        Call SendData(SendTarget.toall, 0, 0, "||Cleopatra Sagrada Regreso al otro Mundo." & FONTTYPE_GUILD)
         
        StatusCleopatra = False
        SecCleopatra = 0
        MataCleopatra = False

    End If
    
    If StatusCleopatra = False Then
          
        If RepiteInvoCleopatra = True Then
              
            If SecCleopatra > 10 Then
                Call SpawnSagrada("Cleopatra")
                SecCleopatra = 0

            End If
              
        End If
         
        If SecCleopatra > IntervaloSagrada Then
            Call SpawnSagrada("Cleopatra")
            SecCleopatra = 0

        End If

    End If
    
    Static SecReyScorpion As Long
    
    SecReyScorpion = SecReyScorpion + 1
    
    If MataReyScorpion = True Then
         
        Call SendData(SendTarget.toall, 0, 0, "||El Rey Scorpion Sagrado Regreso al otro Mundo." & FONTTYPE_GUILD)
         
        StatusReyScorpion = False
        SecReyScorpion = 0
        MataReyScorpion = False

    End If
    
    If StatusReyScorpion = False Then
          
        If RepiteInvoReyScorpion = True Then
              
            If SecReyScorpion > 10 Then
                Call SpawnSagrada("ReyScorpion")
                SecCleopatra = 0

            End If
              
        End If
         
        If SecReyScorpion > IntervaloSagrada Then
            Call SpawnSagrada("ReyScorpion")
            SecReyScorpion = 0

        End If

    End If
    
    Static SecDarkSeth As Long
    
    SecDarkSeth = SecDarkSeth + 1
    
    If MataDarkSeth = True Then
         
        Call SendData(SendTarget.toall, 0, 0, "||El Dark Seth Sagrado Regreso al otro Mundo." & FONTTYPE_GUILD)
         
        StatusDarkSeth = False
        SecDarkSeth = 0
        MataDarkSeth = False

    End If
    
    If StatusDarkSeth = False Then
          
        If RepiteInvoDarkSeth = True Then
              
            If SecDarkSeth > 10 Then
                Call SpawnSagrada("DarkSeth")
                SecCleopatra = 0

            End If
              
        End If
         
        If SecDarkSeth > IntervaloSagrada Then
            Call SpawnSagrada("DarkSeth")
            SecDarkSeth = 0

        End If

    End If
    
    Static SecElfica As Long
    
    SecElfica = SecElfica + 1
    
    If MataElfica = True Then
         
        Call SendData(SendTarget.toall, 0, 0, "||La hada Elfica Sagrada Regreso al otro Mundo." & FONTTYPE_GUILD)
         
        StatusElfica = False
        SecElfica = 0
        MataElfica = False

    End If
    
    If StatusElfica = False Then
          
        If RepiteInvoElfica = True Then
              
            If SecElfica > 10 Then
                Call SpawnSagrada("Elfica")
                SecElfica = 0

            End If
              
        End If
         
        If SecElfica > IntervaloSagrada Then
            Call SpawnSagrada("Elfica")
            SecElfica = 0

        End If

    End If
    
    Static SecGranDragonRojo As Long
    
    SecGranDragonRojo = SecGranDragonRojo + 1
    
    If MataGranDragonRojo = True Then
         
        Call SendData(SendTarget.toall, 0, 0, "||El Gran Dragon Rojo Sagrado Regreso al otro Mundo." & FONTTYPE_GUILD)
         
        StatusGranDragonRojo = False
        SecGranDragonRojo = 0
        MataGranDragonRojo = False

    End If
    
    If StatusGranDragonRojo = False Then
          
        If RepiteInvoGranDragonRojo = True Then
              
            If SecGranDragonRojo > 10 Then
                Call SpawnSagrada("GranDragonRojo")
                SecGranDragonRojo = 0

            End If
              
        End If
         
        If SecGranDragonRojo > IntervaloSagrada Then
            Call SpawnSagrada("GranDragonRojo")
            SecGranDragonRojo = 0

        End If

    End If
    
    Static SecTiburonBlanco As Long
    
    SecTiburonBlanco = SecTiburonBlanco + 1
    
    If MataTiburonBlanco = True Then
         
        Call SendData(SendTarget.toall, 0, 0, "||El Tiburon Blanco Sagrado Regreso al otro Mundo." & FONTTYPE_GUILD)
         
        StatusTiburonBlanco = False
        SecTiburonBlanco = 0
        MataTiburonBlanco = False

    End If
    
    If StatusTiburonBlanco = False Then
          
        If RepiteInvoTiburonBlanco = True Then
              
            If SecTiburonBlanco > 10 Then
                Call SpawnSagrada("TiburonBlanco")
                SecTiburonBlanco = 0

            End If
              
        End If
         
        If SecTiburonBlanco > IntervaloSagrada Then
            Call SpawnSagrada("TiburonBlanco")
            SecTiburonBlanco = 0

        End If

    End If

    'Finish time sagrada
   
    Static TimeInvo As Long
   
    TimeInvo = TimeInvo + 1
   
    If StatusInvo = False Then
        If ConfInvo = 0 Then
            TimeInvo = 0
            Call SendData(SendTarget.ToMap, "0", "96", "||Hasta dentro de 5 minutos no podréis Invocar otra Criatura." & FONTTYPE_GUILD)
            ConfInvo = 1

        End If
      
        If ConfInvo = 1 Then
            If TimeInvo = 300 Then
                Call SendData(SendTarget.toall, "0", "0", "||¡Ya se puede invocar en la sala de invocaciones.!" & FONTTYPE_GUILD)
                StatusInvo = True

            End If

        End If

    End If
       
End Sub

Private Sub torneos_Timer()

    xao = xao + 1

    Select Case xao

        Case 84
            Call SendData(SendTarget.toall, 0, 0, "||Torneo> En 10 minutos se realizará un torneo automatico." & FONTTYPE_GUILD)

        Case 89
            Call SendData(SendTarget.toall, 0, 0, "||Torneo> En 5 minutos se realizará un torneo automatico." & FONTTYPE_GUILD)

        Case 93
            Call SendData(SendTarget.toall, 0, 0, "||Torneo> En 1 minutos se realizará un torneo automatico." & FONTTYPE_GUILD)

        Case 94
            Call torneos_auto(3)

        Case 96

            If Torneo_Esperando = True Then
                Call Torneoauto_Cancela
                xao = 2
            Else
                xao = 2

            End If

    End Select

End Sub

Private Sub tPiqueteC_Timer()

    On Error GoTo errhandler

    Static Segundos As Integer
    Dim NuevaA      As Boolean
    Dim NuevoL      As Boolean
    Dim GI          As Integer

    Segundos = Segundos + 6

    Dim i As Integer

    For i = 1 To LastUser

        If UserList(i).flags.UserLogged Then

            If MapData(UserList(i).pos.Map, UserList(i).pos.X, UserList(i).pos.Y).trigger = eTrigger.ANTIPIQUETE Then
                UserList(i).Counters.PiqueteC = UserList(i).Counters.PiqueteC + 1
                Call SendData(SendTarget.toIndex, i, 0, "Z39")

                If UserList(i).Counters.PiqueteC > 23 Then
                    UserList(i).Counters.PiqueteC = 0
                    Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)

                End If

            Else

                If UserList(i).Counters.PiqueteC > 0 Then UserList(i).Counters.PiqueteC = 0

            End If



            If Segundos >= 18 Then

                '                Dim nfile As Integer
                '                nfile = FreeFile ' obtenemos un canal
                '                Open App.Path & "\logs\maxpasos.log" For Append Shared As #nfile
                '                Print #nfile, UserList(i).Counters.Pasos
                '                Close #nfile
                If Segundos >= 18 Then UserList(i).Counters.Pasos = 0

            End If

        End If

    Next i

    If Segundos >= 18 Then Segundos = 0

    Exit Sub

errhandler:
    Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.Description)

End Sub

