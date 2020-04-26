VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "AoMania"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
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
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   4140
      MouseIcon       =   "frmConnect.frx":08CA
      MousePointer    =   99  'Custom
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4755
      Width           =   3705
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
      Height          =   255
      Left            =   4140
      MouseIcon       =   "frmConnect.frx":1594
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3690
      Width           =   3705
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
      Height          =   420
      Left            =   11505
      MouseIcon       =   "frmConnect.frx":225E
      MousePointer    =   99  'Custom
      Top             =   75
      Width           =   420
   End
   Begin VB.Image Image6 
      Height          =   270
      Left            =   6690
      MouseIcon       =   "frmConnect.frx":2F28
      MousePointer    =   99  'Custom
      Top             =   8250
      Width           =   1290
   End
   Begin VB.Image Image5 
      Height          =   270
      Left            =   5340
      MouseIcon       =   "frmConnect.frx":3BF2
      MousePointer    =   99  'Custom
      Top             =   8250
      Width           =   1290
   End
   Begin VB.Image Image4 
      Height          =   270
      Left            =   3990
      MouseIcon       =   "frmConnect.frx":48BC
      MousePointer    =   99  'Custom
      Top             =   8250
      Width           =   1290
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   4545
      MouseIcon       =   "frmConnect.frx":5586
      MousePointer    =   99  'Custom
      Top             =   6660
      Width           =   2985
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
      Height          =   615
      Left            =   4545
      MouseIcon       =   "frmConnect.frx":6250
      MousePointer    =   99  'Custom
      Top             =   5280
      Width           =   2985
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   4545
      MouseIcon       =   "frmConnect.frx":6F1A
      MousePointer    =   99  'Custom
      Top             =   5955
      Width           =   2985
   End
   Begin VB.Image FONDO 
      Height          =   9000
      Left            =   0
      MouseIcon       =   "frmConnect.frx":7BE4
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   12000
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
    Winsock1.Close
     Winsock1.RemoteHost = CurServerIp
     Winsock1.RemotePort = 6000
    'Winsock1.Connect
 
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then prgRun = False

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
     Call Image2_Click
   End If
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

End Sub

Private Sub Image1_Click()
    
    Call Audio.PlayWave(SND_CLICK)

    frmCrearPersonaje.Show vbModal

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
                             
    Set Image1.Picture = Interfaces.FrmConnect_BtCrearPjApretado

End Sub

Private Sub Image2_Click()

    Call Audio.PlayWave(SND_CLICK)
  
    'If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect

   'If frmConnect.MousePointer = 11 Then
      '  Exit Sub
    'End If
           
    UserName = NombreTXT.Text
   
    UserPassword = MD5String(PasswordTXT.Text)

    If CheckUserData(False) = True Then
        'SendNewChar = False
        EstadoLogin = E_MODO.Normal
        Me.MousePointer = 11
        
        frmMain.Socket1.HostName = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect

    End If
    
    SaveSetting App.exeName, "textos", "Pasword", PasswordTXT.Text
    SaveSetting App.exeName, "textos", "Cuenta", NombreTXT.Text

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
                             
    Set Image2.Picture = Interfaces.FrmConnect_BtConectarApretado

End Sub

Private Sub Image3_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmRecuPass.Show

End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Set Image3.Picture = Interfaces.FrmConnect_BtRecuperarApretado

End Sub

Private Sub Image4_Click()
    Call Audio.PlayWave(SND_CLICK)
    Shell "explorer " & "http:\\www.AoMania.net", vbNormalFocus

End Sub

Private Sub Image5_Click()
    Call Audio.PlayWave(SND_CLICK)
    Shell "explorer " & "https://discord.gg/fc6rxaR", vbNormalFocus

End Sub

Private Sub Image6_Click()
    Call Audio.PlayWave(SND_CLICK)
    Shell "explorer " & "https://www.facebook.com/AOMania.net/", vbNormalFocus

End Sub

Private Sub PasswordTxt_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then Call Image2_Click

End Sub

Private Sub Winsock1_Connect()
       lblStatus = "Servidor online"
       lblStatus.ForeColor = &H8000&
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
      lblStatus = "Servidor cerrado"
      lblStatus.ForeColor = &HC0&
End Sub
