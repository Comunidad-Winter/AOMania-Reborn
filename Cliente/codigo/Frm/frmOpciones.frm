VERSION 5.00
Begin VB.Form frmOpciones 
   BorderStyle     =   0  'None
   Caption         =   "Opciones"
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4770
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmOpciones.frx":0152
   MousePointer    =   99  'Custom
   Picture         =   "frmOpciones.frx":0E1C
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ChkCarteles 
      Caption         =   "Chk carteles"
      Height          =   195
      Left            =   4080
      TabIndex        =   10
      Top             =   2025
      Width           =   180
   End
   Begin VB.CheckBox ChkNpc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1515
      TabIndex        =   9
      Top             =   1995
      Width           =   195
   End
   Begin VB.CheckBox ChkPajaros 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1560
      MaskColor       =   &H00000000&
      TabIndex        =   7
      Top             =   1515
      Width           =   180
   End
   Begin VB.CheckBox ChkTransparencia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4170
      TabIndex        =   6
      Top             =   1230
      Width           =   180
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   300
      Max             =   60
      Min             =   30
      TabIndex        =   5
      Top             =   3495
      Value           =   30
      Width           =   4155
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   300
      Max             =   100
      Min             =   1
      TabIndex        =   4
      Top             =   2895
      Value           =   1
      Width           =   4155
   End
   Begin VB.CheckBox ChkPantalla 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4170
      TabIndex        =   3
      Top             =   1515
      Width           =   180
   End
   Begin VB.CheckBox ChkMusic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4170
      TabIndex        =   2
      Top             =   960
      Width           =   180
   End
   Begin VB.CheckBox ChkSound 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   180
   End
   Begin VB.CommandButton Command3 
      Caption         =   "a"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   9120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Borrar carteles"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2325
      TabIndex        =   11
      Top             =   2025
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Npc's"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   375
      TabIndex        =   8
      Top             =   1965
      Width           =   975
   End
   Begin VB.Image ChkTeclas 
      Height          =   255
      Left            =   1560
      Top             =   1200
      Width           =   255
   End
   Begin VB.Image GuardarConfig 
      Height          =   390
      Left            =   4410
      MouseIcon       =   "frmOpciones.frx":21B6D
      MousePointer    =   99  'Custom
      Top             =   -45
      Width           =   300
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkMusic_Click()

    Select Case ChkMusic.Value
     
        Case 0
            Audio.MusicActivated = False
            ChkMusic.Caption = "Musica Desactivada"
            Audio.StopMidi

        Case 1
            Audio.MusicActivated = True
            ChkMusic.Caption = "Musica Activada"
            Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
       
    End Select

End Sub

Private Sub ChkPajaros_Click()

    Select Case ChkPajaros.Value

        Case 0
            ChkPajaros.Caption = "Pajaritos Desactivado"
            SoundPajaritos = False

        Case 1
            ChkPajaros.Caption = "Pajaritos Activado"
            SoundPajaritos = True

    End Select

End Sub

Private Sub ChkPantalla_Click()

    Select Case ChkPantalla.Value

        Case 0
            frmMain.InitDrawMain False
            ChkPantalla.Caption = "Mover Pantalla Desactivada"
    
        Case 1
            frmMain.InitDrawMain True
            ChkPantalla.Caption = "Mover Pantalla Activada"
      
    End Select

End Sub

Private Sub ChkSound_Click()

    Select Case ChkSound.Value
      
        Case 0
            Audio.SoundActivated = False
            ChkSound.Caption = "Sonidos Desactivados"
            Call Audio.StopWave
            RainBufferIndex = 0
            IsPlaying = PlayLoop.plNone
        
        Case 1
            Audio.SoundActivated = True
            ChkSound.Caption = "Sonidos Activados"
         
    End Select

End Sub



Private Sub ChkTeclas_Click()
    
    Call frmCustomKeys.Show(vbModeless, frmMain)

End Sub

Private Sub ChkTransparencia_Click()

    Select Case ChkTransparencia.Value

        Case 0
            ChkTransparencia.Caption = "Con Transparencia"
            AoSetup.bTransparencia = 0
          
        Case 1
            ChkTransparencia.Caption = "Sin Transparencia"
            AoSetup.bTransparencia = 1

    End Select

End Sub

Private Sub Form_Load()

    'Valores máximos y mínimos para el ScrollBar
  
    If Audio.MusicActivated Then
        ChkMusic.Caption = "Musica Activada"
        ChkMusic.Value = 1
    Else
        ChkMusic.Caption = "Musica Desactivada"
        ChkMusic.Value = 0

    End If
    
    If Audio.SoundActivated Then
        ChkSound.Caption = "Sonidos Activados"
        ChkSound.Value = 1
    Else
        ChkSound.Caption = "Sonidos Desactivada"
        ChkSound.Value = 0

    End If
    
    If DragPantalla Then
        ChkPantalla.Caption = "Mover Pantalla Activada"
        ChkPantalla.Value = 1
    Else
        ChkPantalla.Caption = "Mover Pantalla Desactivada"
        ChkPantalla.Value = 0

    End If
    
    If AoSetup.bTransparencia Then
        ChkTransparencia.Caption = "Sin Transparencia"
        ChkTransparencia.Value = 1
    Else
        ChkTransparencia.Caption = "Con Transparencia"
        ChkTransparencia.Value = 0

    End If
    
    If SoundPajaritos Then
        ChkPajaros.Caption = "Pajaritos Activado"
        ChkPajaros.Value = 1
    Else
        ChkPajaros.Caption = "Pajaritos Desactivado"
        ChkPajaros.Value = 0

    End If
    
    If AoSetup.bNombreNpc Then
        ChkNpc.Value = 1
        Else
        ChkNpc.Value = 0
     End If
     
     If AoSetup.bCarteles Then
         ChkCarteles.Value = 1
     Else
          ChkCarteles.Value = 0
     End If
    
    ChkTeclas.ToolTipText = "Configurar Teclas"
    
    If (VOLUMEN_FX < HScroll1.min) Or (VOLUMEN_FX > HScroll1.max) Then
        VOLUMEN_FX = HScroll1.min

    End If

    HScroll1.Value = VOLUMEN_FX
 
    If (VOLUMEN_MUSICA < HScroll2.min) Or (VOLUMEN_MUSICA > HScroll2.max) Then
        VOLUMEN_MUSICA = HScroll2.min

    End If

    HScroll2.Value = VOLUMEN_MUSICA

End Sub

Private Sub GuardarConfig_Click()
    AoSetup.bTransparencia = ChkTransparencia.Value
    AoSetup.bMusica = ChkMusic.Value
    AoSetup.bSonido = ChkSound.Value
    AoSetup.bMover = ChkPantalla.Value
    AoSetup.bPajaritos = ChkPajaros.Value
    AoSetup.bResolucion = AoSetup.bResolucion
    AoSetup.bEjecutar = AoSetup.bEjecutar
    AoSetup.bNombreNpc = ChkNpc.Value
    AoSetup.bCarteles = ChkCarteles.Value
    
    DoEvents
   
    Dim handle As Integer
    handle = FreeFile
    Open App.Path & "\AOM.cfg" For Binary As handle
    Put handle, , AoSetup
    Close handle
    
    Call HScroll1_Change
    Call HScroll2_Change
   
    DoEvents
   
    Unload Me
   
End Sub

Private Sub HScroll1_Change()
    'FX
    Dim s As Integer

    If (HScroll1.Value < 0) Or (HScroll1.Value > 100) Then Exit Sub
    
    VOLUMEN_FX = HScroll1.Value
    s = 10 ^ ((HScroll1.Value + 900) / 1000 + 1)
 
    Audio.SoundVolume = (s)
    Call WriteVar(DirConfiguracion & "Opciones.opc", "CONFIG", "Vol_fx", str(HScroll1.Value))
 
End Sub
 
Private Sub HScroll2_Change()
    'musica, distinto control que fx en valores
    Dim s As Integer
 
    If (HScroll2.Value < 0) Or (HScroll2.Value > 100) Then Exit Sub
    
    VOLUMEN_MUSICA = HScroll2.Value
    Audio.MusicVolume = (HScroll2.Value)
    Call WriteVar(DirConfiguracion & "Opciones.opc", "CONFIG", "Vol_music", str(HScroll2.Value))
 
End Sub
