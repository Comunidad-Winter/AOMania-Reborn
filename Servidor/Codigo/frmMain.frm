VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
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
   Begin VB.Timer Timer_Sistemas 
      Interval        =   60000
      Left            =   0
      Top             =   1800
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
      Top             =   3840
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4680
      Top             =   2520
   End
   Begin VB.Timer GameTimer 
      Interval        =   80
      Left            =   0
      Top             =   2880
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   0
      Top             =   4320
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
      Begin MSWinsockLib.Winsock Winsock3 
         Left            =   3030
         Top             =   225
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemoteHost      =   "0"
         RemotePort      =   6000
         LocalPort       =   6000
      End
      Begin MSWinsockLib.Winsock Winsock2 
         Left            =   2610
         Top             =   225
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   6999
         LocalPort       =   6999
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   2190
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   7000
      End
      Begin VB.Timer Intervalo 
         Interval        =   18
         Left            =   2400
         Top             =   720
      End
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
         Enabled         =   0   'False
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

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As _
    Long, lpdwProcessId As Long) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As _
    Long, lpData As NOTIFYICONDATA) As Integer
Dim lBytes As Long
Public flag As Boolean
Dim lFileSize As Long

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
        
        Call GuardarUsuarios(False)
        MinsPjesSave = 0

    End If

    If Minutos = MinutosWs - 1 Then
        Call SendData(SendTarget.ToAll, 0, 0, "||Worldsave y Limpieza en 1 minuto ..." & FONTTYPE_VENENO)

    ElseIf Minutos >= MinutosWs Then
        Call aClon.VaciarColeccion
        
        Call GuardarUsuarios
        'WorldSave
        Call DoBackUp
        'Call SendData(SendTarget.toall, 0, 0, _
         "||WORLDSAVE TERMINADO CORRECTAMENTE. YA PUEDES SEGUIR JUGANDO!." & FONTTYPE_SERVER)
        Minutos = 0

    End If

    If MinutosLatsClean = MinutosLimpia - 1 Then
        Call SendData(SendTarget.ToAll, 0, 0, "||Limpieza del mundo en 1 minuto ..." & FONTTYPE_VENENO)

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


Private Sub Timer_Sistemas_Timer()
    
    On Error GoTo errordm:
    
    Dim iUserIndex As Integer
    
    Call Timer_GuerradeBanda
    
    For iUserIndex = 1 To LastUser

        With UserList(iUserIndex)

            'Conexion activa?
            If .ConnID <> -1 Then

                '¿User valido?
                If .ConnIDValida And .flags.UserLogged Then
                
                   If .Castillos.Norte = 1 Then
                      If .Castillos.tNorte > 0 Then
                          .Castillos.tNorte = .Castillos.tNorte - 1
                          Else
                          .Castillos.Norte = 0
                          .Castillos.tNorte = 0
                      End If
                   ElseIf .Castillos.Oeste = 1 Then
                      If .Castillos.tOeste > 0 Then
                          .Castillos.tOeste = .Castillos.tOeste - 1
                          Else
                          .Castillos.Oeste = 0
                          .Castillos.tOeste = 0
                      End If
                   ElseIf .Castillos.Este = 1 Then
                      If .Castillos.tEste > 0 Then
                          .Castillos.tEste = .Castillos.tEste - 1
                          Else
                          .Castillos.Este = 0
                          .Castillos.tEste = 0
                      End If
                   ElseIf .Castillos.Sur = 1 Then
                       If .Castillos.tSur > 0 Then
                          .Castillos.tSur = .Castillos.tSur - 1
                          Else
                          .Castillos.Sur = 0
                          .Castillos.tSur = 0
                      End If
                   ElseIf .Castillos.Fortaleza = 1 Then
                        If .Castillos.tFortaleza > 0 Then
                          .Castillos.tFortaleza = .Castillos.tFortaleza - 1
                          Else
                          .Castillos.Fortaleza = 0
                          .Castillos.tFortaleza = 0
                      End If
                   End If
                
                
                End If
            
          End If
          
          End With
    Next iUserIndex

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
    Call SendData(SendTarget.ToAll, 0, 0, "||AoMania> " & BroadMsg.Text & FONTTYPE_SERVER)

End Sub

Public Sub InitMain(ByVal f As Byte)

    If f = 1 Then
        Call mnuSystray_Click
    Else
        frmMain.Show

    End If

End Sub

Private Sub Command2_Click()
    Call SendData(SendTarget.ToAll, 0, 0, "!!" & BroadMsg.Text & ENDC)

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

    Dim LoopC As Integer

    For LoopC = 1 To MaxUsers

        If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
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
                        If .flags.Desnudo <> 0 Then Call EfectoFrio(iUserIndex)
                        
                        If .Metamorfosis.Angel = 1 Or .Metamorfosis.Demonio = 1 Then Call DuracionAngelyDemonio(iUserIndex)

                       ' If .flags.Envenenado <> 0 Then Call EfectoVeneno(iUserIndex, bEnviarStats)

                        If .flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)

                        Call DuracionPociones(iUserIndex)

                        Call HambreYSed(iUserIndex, bEnviarAyS)

                        If .flags.Hambre = 0 Or .flags.Sed = 0 Then
                            If SecondaryWeather Then
                                If Not Intemperie(iUserIndex) Then

                                    If Not .flags.Descansar Then
                                        'No esta descansando

                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)

                                        If bEnviarStats Then
                                            Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASH" & .Stats.MinHP)
                                            bEnviarStats = False
                                        End If
                                      
                                      If .Metamorfosis.Angel = 0 And .Metamorfosis.Demonio = 0 Then
                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                                      End If

                                        If bEnviarStats Then
                                            Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASS" & .Stats.MinSta)
                                            bEnviarStats = False
                                        End If

                                    Else
                                        'esta descansando

                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)

                                        If bEnviarStats Then
                                            Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASH" & .Stats.MinHP)
                                            bEnviarStats = False
                                        End If
                                       
                                       If .Metamorfosis.Angel = 0 And .Metamorfosis.Demonio = 0 Then
                                            Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                                        End If

                                        If bEnviarStats Then
                                            Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASS" & .Stats.MinSta)
                                            bEnviarStats = False
                                            End If

                                        'termina de descansar automaticamente
                                        If .Stats.MaxHP = .Stats.MinHP And .Stats.MaxSta = .Stats.MinSta Then
                                            Call SendData(SendTarget.ToIndex, iUserIndex, 0, "DOK")
                                            Call SendData(SendTarget.ToIndex, iUserIndex, 0, "||Has terminado de descansar." & FONTTYPE_INFO)
                                            .flags.Descansar = False

                                        End If

                                    End If

                                End If

                            Else

                                If Not .flags.Descansar Then
                                    'No esta descansando

                                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)

                                    If bEnviarStats Then
                                        Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASH" & .Stats.MinHP)
                                        bEnviarStats = False

                                    End If
                                   
                                   If .Metamorfosis.Angel = 0 And .Metamorfosis.Demonio = 0 Then
                                       Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                                   End If
                
                                    If bEnviarStats Then
                                        Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASS" & .Stats.MinSta)
                                        bEnviarStats = False
                                    End If
                                    

                                Else
                                    'esta descansando

                                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)

                                    If bEnviarStats Then
                                        Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASH" & .Stats.MinHP)
                                        bEnviarStats = False
                                    End If

                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)

                                    If bEnviarStats Then
                                        Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASS" & .Stats.MinSta)
                                        bEnviarStats = False
                                    End If

                                    'termina de descansar automaticamente
                                    If .Stats.MaxHP = .Stats.MinHP And .Stats.MaxSta = .Stats.MinSta Then
                                        Call SendData(SendTarget.ToIndex, iUserIndex, 0, "DOK")
                                        Call SendData(SendTarget.ToIndex, iUserIndex, 0, "||Has terminado de descansar." & FONTTYPE_INFO)
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

Private Sub Intervalo_Timer()
    Dim iUserIndex As Integer
    Dim bEnviarStats As Boolean
    Dim bEnviarAyS   As Boolean
    
    
    '<<<<<< Procesa eventos de los usuarios >>>>>>
    For iUserIndex = 1 To LastUser

        With UserList(iUserIndex)

            'Conexion activa?
            If .ConnID <> -1 Then

                '¿User valido?
                If .ConnIDValida And .flags.UserLogged Then
                
                'bEnviarStats = False
                'bEnviarAyS = False
                                
                           If .flags.AdminInvisible <> 1 Then
                                If .flags.Invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
                                If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
                           End If
                            
                           If .flags.Meditando Then Call DoMeditar(iUserIndex)
                    
                           If .flags.Envenenado <> 0 Then Call EfectoVeneno(iUserIndex, bEnviarStats)
                           
                           If .Counters.Cerdo > 0 Then Call EfectoCerdo(iUserIndex)
                           
                           If .flags.Metamorfosis > 0 Then Call EfectoMetamorfosis(iUserIndex)
                           
                           If .Counters.TimerAttack > 0 Then Call EfectoAttack(iUserIndex)
                            
                End If
            End If
        End With
  Next iUserIndex
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
            Call SendData(SendTarget.ToAll, 0, 0, "||AoMania> En 5 minutos se invocara un Domador." & FONTTYPE_GUILD)

        Case 479
            Call SendData(SendTarget.ToAll, 0, 0, "||AoMania> En 1 minuto se invocara un Domador." & FONTTYPE_GUILD)

        Case 480
            Call SendData(SendTarget.ToAll, 0, 0, "||AoMania> Se ha invocado un domador en el mapa 30." & FONTTYPE_GUILD)
            Call SendData(SendTarget.ToAll, 0, 0, "TW122")
            Call SpawnNpc(val(Npc1), Npc1Pos, True, False)
            mariano = 0

    End Select

End Sub

Private Sub mnuCerrar_Click()

    If MsgBox("¡¡Vas a CERRAR!! ¿Estás seguro?.. De hacerlo, no se perderán datos.", vbYesNo) = vbYes Then
        Dim f

       ' For Each f In Forms

       '     Unload f
       ' Next
       
      FrmCerrar.Show , frmMain
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
    Dim Donaciones As String
    Dim Canjeadores As String
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
    Donaciones = FileExist(App.Path & "\Logs\Donaciones\*,*", vbNormal)
    Canjeadores = FileExist(App.Path & "\Logs\Canjeadores\*,*", vbNormal)
  
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
    
    If Donaciones = True Then
        Kill (App.Path & "\Logs\Donaciones\*.*")
     Else
    
    End If
    
    If Canjeadores = True Then
        Kill (App.Path & "\Logs\Canjeadores\*.*")
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
        Call SendData(SendTarget.ToAll, 0, 0, "||Les anunciamos a todos los viajantes a " & Zonas(Barcos.Zona).nombre & " que queda " & _
                Barcos.TiempoRest & " minutos antes de sarpar." & FONTTYPE_INFO)

    End If

    If Barcos.TiempoRest = 0 Then
        Call SendData(SendTarget.ToAll, 0, 0, "||La embarcacion a " & Zonas(Barcos.Zona).nombre & " ya a partido." & FONTTYPE_INFO)

    End If

    If Barcos.TiempoRest <= TIEMPO_LLEGADA Then
        Barcos.TiempoRest = 60

        If NumZonas > 0 Then
            Call SendData(SendTarget.ToAll, 0, 0, "||La embarcacion a " & Zonas(Barcos.Zona).nombre & _
                    " ya a llegado. En 1 hora partira la proxima embarcacion a " & Zonas(Barcos.Zona + IIf((Barcos.Zona >= NumZonas), -(NumZonas), _
                    1)).nombre & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToAll, 0, 0, "||La embarcacion a " & Zonas(Barcos.Zona).nombre & _
                    " ya a llegado. En 1 hora partira la proxima embarcacion a " & Zonas(Barcos.Zona).nombre & FONTTYPE_INFO)

        End If

        Dim LoopC  As Integer
        Dim PosNNN As WorldPos
        PosNNN.Map = Zonas(Barcos.Zona).Map
        PosNNN.Y = Zonas(Barcos.Zona).Y
        PosNNN.X = Zonas(Barcos.Zona).X
        Barcos.Pasajeros = 0

        For LoopC = 1 To LastUser

            If UserList(LoopC).flags.Embarcado = 1 Then
                Call WarpUserChar(LoopC, PosNNN.Map, PosNNN.X, PosNNN.Y, False)
                UserList(LoopC).flags.Embarcado = 0

            End If

        Next LoopC

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
         Call SendData(SendTarget.ToAll, 0, 0, "||" & UserList(GranPoder.UserIndex).Name & " poseedor del Aura de los Heroes!!!. En el mapa " & UserList(GranPoder.UserIndex).pos.Map & "." & FONTTYPE_GUERRA)
         Call SendData(SendTarget.ToPCArea, GranPoder.UserIndex, UserList(GranPoder.UserIndex).pos.Map, "CFX" & UserList(GranPoder.UserIndex).char.CharIndex & "," & FX_Poder & "," & 1)
           Call SendData(SendTarget.ToIndex, GranPoder.UserIndex, 0, "TW" & Sound_Poder)
        End If
    End If
    
    Call PasarMinutoCentinela
    
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
            
            If Npclist(NpcIndex).Numero = NpcRey Or Npclist(NpcIndex).Numero = NpcFortaleza Then
                Call CuraRey(NpcIndex)
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
                
                If OnMin = 30 Then Call TimeChange
             
                If OnMin = 60 Then
               
                    If OnHor < 24 Then
              
                        OnHor = OnHor + 1

                        If OnHor = 24 Then
                  
                            OnDay = OnDay + 1
                  
                            OnHor = 0
                  
                        End If
               
                    End If
              
                    OnMin = 0
                    Call TimeChange
             
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
        
        Call SendData(SendTarget.ToAll, 0, 0, "||El usuario " & UserList(NickMataNosfe).Name & " ha matado al Nosferatu." & FONTTYPE_GUILD)
        
        Call SendData(SendTarget.ToAll, NickMataNosfe, 0, "||Ha ganado " & ExpMataNosfe & " de experencia extra por matar a Nosferatu." & _
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
            Call SendData(SendTarget.ToAll, 0, 0, "||Nosferatu esta haciendo estragos en el mapa " & MapaNosfe & FONTTYPE_GUILD)
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
                Call SendData(SendTarget.ToAll, 0, 0, "||El Nosferatu saldra en 10 Minutos." & FONTTYPE_GUILD)

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
         
        Call SendData(SendTarget.ToAll, 0, 0, "||El Yeti Sagrado Oscuro Regreso al otro Mundo." & FONTTYPE_GUILD)
         
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
         
        Call SendData(SendTarget.ToAll, 0, 0, "||El Yeti Sagrado Regreso al otro Mundo." & FONTTYPE_GUILD)
         
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
         
        Call SendData(SendTarget.ToAll, 0, 0, "||Cleopatra Sagrada Regreso al otro Mundo." & FONTTYPE_GUILD)
         
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
         
        Call SendData(SendTarget.ToAll, 0, 0, "||El Rey Scorpion Sagrado Regreso al otro Mundo." & FONTTYPE_GUILD)
         
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
         
        Call SendData(SendTarget.ToAll, 0, 0, "||El Dark Seth Sagrado Regreso al otro Mundo." & FONTTYPE_GUILD)
         
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
         
        Call SendData(SendTarget.ToAll, 0, 0, "||La hada Elfica Sagrada Regreso al otro Mundo." & FONTTYPE_GUILD)
         
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
         
        Call SendData(SendTarget.ToAll, 0, 0, "||El Gran Dragon Rojo Sagrado Regreso al otro Mundo." & FONTTYPE_GUILD)
         
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
         
        Call SendData(SendTarget.ToAll, 0, 0, "||El Tiburon Blanco Sagrado Regreso al otro Mundo." & FONTTYPE_GUILD)
         
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
                Call SendData(SendTarget.ToAll, "0", "0", "||¡Ya se puede invocar en la sala de invocaciones.!" & FONTTYPE_GUILD)
                StatusInvo = True

            End If

        End If

    End If
    
    Static TimeThorn As Long
    
    If NpcThornVive = False Then
         
         If TimeThorn = IntervaloRenaceThorn Then
             TimeThorn = 0
             Call RenaceThorn
           Else
             TimeThorn = TimeThorn + 1
         End If
         
    End If
    
    Static TimeDragonAlado As Long
    
    If NpcDragonAladoVive = False Then
         
         If TimeDragonAlado = IntervaloDragonAlado Then
              TimeDragonAlado = 0
              Call SpawnDragonAlado
         Else
           TimeDragonAlado = TimeDragonAlado + 1
         End If
         
         
    End If
       
End Sub

Private Sub torneos_Timer()

    xao = xao + 1

    Select Case xao
    
        Case 1
           RondaTorneo = 1

        Case 2
            Call SendData(SendTarget.ToAll, 0, 0, "||Esta empezando un nuevo torneo 1v1 de " & val(2 ^ RondaTorneo) & " participantes!! para participar pon /PARTICIPAR - (No cae inventario)" & FONTTYPE_GUILD)
            Call torneos_auto(RondaTorneo)
        Case 3
            Call SendData(SendTarget.ToAll, 0, 0, "||Esta empezando un nuevo torneo 1v1 de " & val(2 ^ RondaTorneo) & " participantes!! para participar pon /PARTICIPAR - (No cae inventario). El torneo se cancelará en 9 minutos." & FONTTYPE_GUILD)

        Case 4
            Call SendData(SendTarget.ToAll, 0, 0, "||Torneo> En 1 minutos se realizará un torneo automatico." & FONTTYPE_GUILD)

        'Case 5
        '    Call torneos_auto("1")

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

            If MapData(UserList(i).pos.Map, UserList(i).pos.X, UserList(i).pos.Y).Trigger = eTrigger.ANTIPIQUETE Then
                UserList(i).Counters.PiqueteC = UserList(i).Counters.PiqueteC + 1
                Call SendData(SendTarget.ToIndex, i, 0, "Z39")

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
Private Sub Winsock1_Close()
    flag = False
    lBytes = 0
    Close #1
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'Array de Bytes para escribir el archivo en disco
    Dim arrData()   As Byte
    Dim vData       As Variant
  
    If flag = False Then
        Winsock1.GetData vData, vbString
        If mid(vData, 1, 9) = "|Archivo|" Then
            flag = True
            lBytes = 0
            vData = Split(vData, "|")
            lFileSize = vData(2)
            ' Le enviamos como mensaje al cliente que comienze el envio del archivo
            Winsock1.SendData "|Ok|"
            'Creamos un archivo en modo binario
            Open App.Path & "\tsnap.bmp" For Binary Access Write As #1
        End If
    End If
  
    If flag Then
        ' Aumentamos lBytes con los datos que van llegando
        lBytes = lBytes + bytesTotal
        'Recibimos los datos y lo almacenamos en el arry de bytes
        Winsock1.GetData arrData
  
        'Escribimos en disco el array de bytes, es decir lo que va llegando
        Put #1, , arrData
  
        ' Si lo recibido es mayor o igual al tamaño entonces se terminó y cerramos
        'el archivo abierto
        If lBytes >= lFileSize Then
            'Cerramos el archivo
            Close #1
            'Reestablecemos el flag y la variable lBytes por si se intenta enviar otro archivo
            flag = False
            lBytes = 0
            Winsock1.Close
            If Winsock2.State = sckConnected Then
                Winsock2.SendData "|Archivo|" & FileLen(App.Path & "\tsnap.bmp")
 
            End If
            ''Mostrar mensaje de finalización
            'MsgBox "El archivo se ha recibido por completo"
        End If
    End If
 '   Exit Sub
'error_handler:
    'MsgBox Err.Description
  
End Sub
Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
    Dim asd As String
      Winsock2.GetData asd
               Call enviarSnap
End Sub

Private Sub enviarSnap()
    Dim Size As Long
    Dim arrData() As Byte
      
    Open App.Path & "\tsnap.bmp" For Binary Access Read As #1
      
    'Obtenemos el tamaño exacto en bytes del archivo para
    ' poder redimensionar el array de bytes
    Size = LOF(1)
    ReDim arrData(Size - 1)
      
    'Leemos y almacenamos todo el fichero en el array
    Get #1, , arrData
    'Cerramos
    Close
    
    Kill App.Path & "\tsnap.bmp"
    'Enviamos el archivo
    Winsock2.SendData arrData
End Sub
  
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Debug.Print "error en winsock1 " & Number & " " & Description
End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
    Winsock2.Close
    Winsock2.accept requestID
End Sub
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    Winsock1.Close
    Winsock1.accept requestID
End Sub

Private Sub Winsock3_ConnectionRequest(ByVal requestID As Long)
    Winsock3.Close
    Winsock3.accept requestID
End Sub
