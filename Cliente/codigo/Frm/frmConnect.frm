VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "AoMania"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   Picture         =   "frmConnect.frx":08CA
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   240
      Left            =   9585
      MaskColor       =   &H00000000&
      TabIndex        =   4
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   240
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1830
      Top             =   1305
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   6000
      LocalPort       =   6000
   End
   Begin VB.TextBox PasswordTXT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   5295
      MouseIcon       =   "frmConnect.frx":A827D
      MousePointer    =   99  'Custom
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   6045
      Width           =   4755
   End
   Begin VB.TextBox NombreTXT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   5295
      MouseIcon       =   "frmConnect.frx":A8F47
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4680
      Width           =   4755
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   975
      TabIndex        =   3
      Top             =   2475
      Width           =   75
   End
   Begin VB.Image Exit 
      Height          =   525
      Left            =   14715
      MouseIcon       =   "frmConnect.frx":A9C11
      MousePointer    =   99  'Custom
      Top             =   105
      Width           =   525
   End
   Begin VB.Image Image6 
      Height          =   345
      Left            =   8565
      MouseIcon       =   "frmConnect.frx":AA8DB
      MousePointer    =   99  'Custom
      Top             =   10560
      Width           =   1650
   End
   Begin VB.Image Image5 
      Height          =   345
      Left            =   6840
      MouseIcon       =   "frmConnect.frx":AB5A5
      MousePointer    =   99  'Custom
      Top             =   10560
      Width           =   1650
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   5130
      MouseIcon       =   "frmConnect.frx":AC26F
      MousePointer    =   99  'Custom
      Top             =   10560
      Width           =   1635
   End
   Begin VB.Image Image3 
      Height          =   810
      Left            =   5820
      MouseIcon       =   "frmConnect.frx":ACF39
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   3810
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Image2 
      Height          =   810
      Left            =   5820
      MouseIcon       =   "frmConnect.frx":ADC03
      MousePointer    =   99  'Custom
      Top             =   6750
      Width           =   3780
   End
   Begin VB.Image Image1 
      Height          =   810
      Left            =   5820
      MouseIcon       =   "frmConnect.frx":AE8CD
      MousePointer    =   99  'Custom
      Top             =   7620
      Width           =   3780
   End
   Begin VB.Image FONDO 
      Height          =   11520
      Left            =   0
      MouseIcon       =   "frmConnect.frx":AF597
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Connect As String = "173.WAV"

Private Sub Exit_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call UnloadAllForms

End Sub

Private Sub FONDO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Set Image2.Picture = Interfaces.FrmConnect_BtConectar
    Set Image1.Picture = Interfaces.FrmConnect_BtCrearPj
    Set Image3.Picture = Interfaces.FrmConnect_BtRecuperar

End Sub

Private Sub Form_Activate()
    
    Dim i As Integer
    
    NombreTXT.SetFocus
    
    If i = 0 Then
        i = i + 1

        Call Audio.PlayWave("173.wav")
    End If
 
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then prgRun = False
    If KeyCode = 13 Then Call Image2_Click

End Sub


Private Sub Form_Load()
    '[CODE 002]:MatuX
    EngineRun = False
    '[END]
    
    version.Caption = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    
    Set FONDO.Picture = Interfaces.FrmConnect_Principal
    Set Image2.Picture = Interfaces.FrmConnect_BtConectar
    Set Image1.Picture = Interfaces.FrmConnect_BtCrearPj
    Set Image3.Picture = Interfaces.FrmConnect_BtRecuperar
    Set FONDO.MouseIcon = Iconos.Ico_Diablo

    PasswordTXT.Text = GetSetting(App.exeName, "textos", "Pasword", vbNullString)
    NombreTXT.Text = GetSetting(App.exeName, "textos", "Cuenta", vbNullString)
    Check1.Value = Val(GetSetting(App.exeName, "textos", "Check", vbNullString))

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Check1.Value = 1 Then
        SaveSetting App.exeName, "textos", "Pasword", PasswordTXT.Text
        SaveSetting App.exeName, "textos", "Cuenta", NombreTXT.Text
        SaveSetting App.exeName, "textos", "Check", Check1.Value
    ElseIf Check1.Value = 0 Then
        SaveSetting App.exeName, "textos", "Pasword", ""
        SaveSetting App.exeName, "textos", "Cuenta", ""
        SaveSetting App.exeName, "textos", "Check", Check1.Value

    End If

End Sub

Private Sub Image1_Click()
    
    Call Audio.PlayWave(SND_CLICK)

    frmCrearPersonaje.Show vbModal

End Sub

Private Sub Image1_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
                             
    Set Image1.Picture = Interfaces.FrmConnect_BtCrearPjApretado

End Sub

Private Sub Image2_Click()

    Call Audio.PlayWave(SND_CLICK)
           
    UserName = NombreTXT.Text
   
    UserPassword = MD5String(PasswordTXT.Text)

    EstadoLogin = E_MODO.Normal
    Me.MousePointer = 11
    
    If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
    
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
    
    If Check1.Value = 1 Then
        SaveSetting App.exeName, "textos", "Pasword", PasswordTXT.Text
        SaveSetting App.exeName, "textos", "Cuenta", NombreTXT.Text
        SaveSetting App.exeName, "textos", "Check", Check1.Value
    ElseIf Check1.Value = 0 Then
        SaveSetting App.exeName, "textos", "Pasword", ""
        SaveSetting App.exeName, "textos", "Cuenta", ""
        SaveSetting App.exeName, "textos", "Check", Check1.Value

    End If

End Sub

Private Sub Image2_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
                             
    Set Image2.Picture = Interfaces.FrmConnect_BtConectarApretado

End Sub

Private Sub Image3_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmRecuPass.Show

End Sub

Private Sub Image3_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    Set Image3.Picture = Interfaces.FrmConnect_BtRecuperarApretado

End Sub

Private Sub Image4_Click()
    Call Audio.PlayWave(SND_CLICK)
    Shell "explorer " & "https:\\www.AoMania.net", vbNormalFocus

End Sub

Private Sub Image5_Click()
    Call Audio.PlayWave(SND_CLICK)
    Shell "explorer " & "https://discord.gg/fc6rxaR", vbNormalFocus

End Sub

Private Sub Image6_Click()
    Call Audio.PlayWave(SND_CLICK)
    Shell "explorer " & "https://www.facebook.com/AOMania.net/", vbNormalFocus

End Sub

'Private Sub PasswordTxt_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyReturn Then Call Image2_Click
'
'End Sub
