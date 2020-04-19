VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   0  'None
   Caption         =   "Configuración de AoMania"
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2850
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmmain.frx":31C32A
   ScaleHeight     =   3300
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkEjecutar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ejecutar cliente"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   420
      TabIndex        =   4
      Top             =   2100
      Width           =   195
   End
   Begin VB.CheckBox ChkPantalla 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cambiar resolución"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   420
      TabIndex        =   3
      Top             =   1770
      Width           =   195
   End
   Begin VB.CheckBox ChkSonidos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sonidos activados"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   420
      TabIndex        =   2
      Top             =   1440
      Width           =   195
   End
   Begin VB.CheckBox ChkMusic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Musica activada"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   420
      TabIndex        =   1
      Top             =   1125
      Width           =   195
   End
   Begin VB.CheckBox ChkTransparencia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   420
      MaskColor       =   &H80000005&
      TabIndex        =   0
      Top             =   795
      Width           =   195
   End
   Begin VB.Image ImgBat 
      Height          =   375
      Left            =   480
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Image ImgSalir 
      Height          =   375
      Left            =   1440
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image ImgGuardar 
      Height          =   375
      Left            =   360
      Top             =   2400
      Width           =   975
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ImgGuardar_Click()
   AoSetup.bTransparencia = ChkTransparencia.value
   AoSetup.bMusica = ChkMusic.value
   AoSetup.bSonido = ChkSonidos.value
   AoSetup.bResolucion = ChkPantalla.value
   AoSetup.bEjecutar = ChkEjecutar.value
   
   DoEvents
   
   Dim handle As Integer
   handle = FreeFile
   Open App.Path & "\AOM.cfg" For Binary As handle
    Put handle, , AoSetup
   Close handle
    
    
End Sub

Private Sub ImgSalir_Click()
   If ChkEjecutar.value = 1 Then
       Call Shell(App.Path & "\AoMania.exe")
   End If
   Unload Me
End Sub

Private Sub ImgBat_Click()
   Call Shell(App.Path & "\Solucionador.bat")
End Sub

Private Sub Form_Load()
   Call Mod_General.LeerSetup
End Sub
