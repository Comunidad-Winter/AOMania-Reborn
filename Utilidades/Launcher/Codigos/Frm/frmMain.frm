VERSION 5.00
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Launcher AoMania Reborn"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   Picture         =   "frmMain.frx":08CA
   ScaleHeight     =   520
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   3645
      Top             =   60
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2715
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3180
      Top             =   60
   End
   Begin vbalProgBarLib6.vbalProgressBar ProgressBar1 
      Height          =   405
      Left            =   2115
      TabIndex        =   1
      Top             =   7290
      Visible         =   0   'False
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   714
      Picture         =   "frmMain.frx":386E0
      ForeColor       =   16777215
      Appearance      =   0
      BarPicture      =   "frmMain.frx":386FC
      ShowText        =   -1  'True
      Text            =   "[0% Completado]"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2100
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image picNotice 
      Height          =   6015
      Left            =   3180
      Top             =   630
      Width           =   8775
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1530
      TabIndex        =   4
      Top             =   2490
      Width           =   120
   End
   Begin VB.Label lblOro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2520
      TabIndex        =   3
      Top             =   2955
      Width           =   120
   End
   Begin VB.Label LblExp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   885
      TabIndex        =   2
      Top             =   2955
      Width           =   120
   End
   Begin VB.Image PicStatus 
      Height          =   510
      Left            =   690
      Top             =   1425
      Width           =   1755
   End
   Begin VB.Image cmdFacebook 
      Height          =   480
      Left            =   1050
      Top             =   75
      Width           =   390
   End
   Begin VB.Image cmdDiscord 
      Height          =   480
      Left            =   525
      Top             =   75
      Width           =   480
   End
   Begin VB.Image cmdWeb 
      Height          =   480
      Left            =   75
      Top             =   75
      Width           =   420
   End
   Begin VB.Image cmdConf 
      Height          =   420
      Left            =   10590
      Top             =   90
      Width           =   435
   End
   Begin VB.Image cmdPlay 
      Height          =   1095
      Left            =   45
      Top             =   6660
      Width           =   1935
   End
   Begin VB.Image PicMarcosUpdate 
      Height          =   570
      Left            =   1995
      Top             =   7185
      Visible         =   0   'False
      Width           =   9960
   End
   Begin VB.Image cmdMinimizar 
      Height          =   420
      Left            =   11055
      MousePointer    =   99  'Custom
      Top             =   90
      Width           =   435
   End
   Begin VB.Label txtUpdate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Info"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2100
      TabIndex        =   0
      Top             =   6825
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdCerrar 
      Height          =   420
      Left            =   11505
      Top             =   90
      Width           =   435
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' LauncherAoM 1.0.0
' By Bassinger [www.AoMania.Net]

Option Explicit

Private Sub cmdCerrar_Click()
  Call UnloadAllForms
End Sub

Private Sub cmdCerrar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
     With frmMain
         .MouseIcon = Iconos.Ico_Mano
     End With
     
End Sub

Private Sub cmdConf_Click()
      Call ShellExecute(Me.hWnd, "Open", App.path & "\AoManiaSetup.exe", 0, 0, 1)
End Sub

Private Sub cmdConf_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     With frmMain
         .MouseIcon = Iconos.Ico_Mano
     End With
End Sub

Private Sub cmdDiscord_Click()
      Shell "explorer " & "https://discord.gg/BVJBfC5", vbNormalFocus
End Sub

Private Sub cmdDiscord_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      With frmMain
         .MouseIcon = Iconos.Ico_Mano
     End With
End Sub

Private Sub cmdFacebook_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     With frmMain
         .MouseIcon = Iconos.Ico_Mano
     End With
End Sub

Private Sub cmdMinimizar_Click()
    If WindowState <> vbMinimized Then WindowState = vbMinimized
    Visible = True
End Sub

Private Sub cmdMinimizar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    With frmMain
         .MouseIcon = Iconos.Ico_Mano
     End With
End Sub

Private Sub cmdPlay_Click()
        
    Select Case Launcher.Play
               
        Case 0
            Exit Sub
               
        Case 1
               Call ShellExecute(Me.hWnd, "Open", App.path & "\AoMania.exe", 0, 0, 1)
               Unload Me
            Exit Sub
               
    End Select
        
End Sub

Private Sub cmdPlay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    With frmMain
         .MouseIcon = Iconos.Ico_Mano
     End With
End Sub

Private Sub cmdWeb_Click()
     Shell "explorer " & "http:\\www.AoMania.net", vbNormalFocus
End Sub

Private Sub cmdWeb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   With frmMain
         .MouseIcon = Iconos.Ico_Mano
     End With
End Sub

Private Sub Form_Load()
   
   ProximaIMG = 2
   
   With frmMain
        .MouseIcon = Iconos.Ico_Diablo
        .Picture = Interfaces.Fondo_Principal
        .PicMarcosUpdate.Picture = Interfaces.MBUpdate
        .ProgressBar1.Picture = Interfaces.BVacia
        .ProgressBar1.BarPicture = Interfaces.BLlena
        .cmdPlay.Picture = Interfaces.NoPlay
        .picNotice = Interfaces.Notice1
    End With

   Call LoadServer
   Winsock1.Connect
      
   If Launcher.Use = 0 Then
       txtUpdate.Caption = "Comprobando y registrando (dll/ocx)"
       Call RevDlls
   End If
   
   If Launcher.Play = 1 Then
        Launcher.Play = 0
   End If

    ProgressBar1.Value = 0
    ProgressBar1.Text = ""
    txtUpdate.Caption = ""
    Timer1.Enabled = True
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     
     With frmMain
         .MouseIcon = Iconos.Ico_Diablo
     End With
     
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     Call UnloadAllForms
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    
    Select Case State
        
        Case icError
            PicMarcosUpdate.Visible = False
            ProgressBar1.Visible = False
            SetUpdate = "1"
            txtUpdate.Caption = "Error en la conexión, descarga abortada."
            bDone = True
            dError = True
            Exit Sub
            
        Case icResponseCompleted
            Dim vtData As Variant
            Dim tempArray() As Byte
            Dim FileSize As Long

            
            FileSize = Inet1.GetHeader("Content-length")
            ProgressBar1.max = FileSize
            ProgressBar1.BarPicture = Interfaces.BLlena
            
            Open Directory For Binary Access Write As #1
                vtData = Inet1.GetChunk(1024, icByteArray)
                DoEvents
                
                
                Do While Not Len(vtData) = 0
                    tempArray = vtData
                    Put #1, , tempArray
            
                vtData = Inet1.GetChunk(1024, icByteArray)
                    
                    ProgressBar1.Value = ProgressBar1.Value + Len(vtData) * 2
                    txtUpdate.Caption = "Descargando: " & CLng((ProgressBar1.Value + Len(vtData) * 2) / 1000000) & " MBs de " & CLng((FileSize / 1000000)) & " MBs"
                    ProgressBar1.Text = Round(CDbl(ProgressBar1.Value) * CDbl(100) / CDbl(ProgressBar1.max), 2) _
                            & "%"
                    
                    DoEvents
                    
                Loop
            Close #1
            
            txtUpdate.Caption = "¡Ok! Actualización finalizada."
            
            ProgressBar1.Value = 0
            
            bDone = True
    End Select
    
End Sub

Private Sub Timer1_Timer()

    Static TimerUpdater As Long

    TimerUpdater = TimerUpdater + "1"
   
    If TimerUpdater = "80" Then
        Call Analizar
        'SetUpdate = "1"
        'TimerUpdater = "0"
        Timer1.Enabled = False

    End If

End Sub

Private Sub Timer2_Timer()
          
          Select Case ProximaIMG
                
                Case 1
                   frmMain.picNotice.Picture = Interfaces.Notice1
                   ProximaIMG = ProximaIMG + 1
                
                Case 2
                   frmMain.picNotice.Picture = Interfaces.Notice2
                   ProximaIMG = ProximaIMG + 1
                Case 3
                   frmMain.picNotice.Picture = Interfaces.Notice3
                   ProximaIMG = 1
                
          End Select
          
End Sub

Private Sub Winsock1_Connect()
     Call ChangeStatus(eStatus.Online)
     Winsock1.Close
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
      Call ChangeStatus(eStatus.Offline)
      Winsock1.Close
End Sub
