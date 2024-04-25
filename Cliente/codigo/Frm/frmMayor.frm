VERSION 5.00
Begin VB.Form frmMayor 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   -60
   ClientTop       =   -60
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   5175
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image salir 
      Height          =   255
      Left            =   4320
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Label NameOroOffline 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   7
      Top             =   4120
      Width           =   405
   End
   Begin VB.Label NameOroOnline 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      Top             =   3410
      Width           =   405
   End
   Begin VB.Label CntCriOn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3405
      TabIndex        =   5
      Top             =   4595
      Width           =   90
   End
   Begin VB.Label CntCiuOn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3405
      TabIndex        =   4
      Top             =   4895
      Width           =   90
   End
   Begin VB.Label NameCriMax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   3
      Top             =   2720
      Width           =   405
   End
   Begin VB.Label NameCiuMax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   1980
      Width           =   405
   End
   Begin VB.Label NameCriMaxLvl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   1
      Top             =   1310
      Width           =   405
   End
   Begin VB.Label NameCiuMaxLvl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   405
   End
End
Attribute VB_Name = "frmMayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Set frmMayor.Picture = Interfaces.FrmMayor_Principal
    Set Me.MouseIcon = Iconos.Ico_Diablo
   
    NameCiuMaxLvl = Mayores.CiudadanoMaxNivel
    NameCriMaxLvl = Mayores.CriminalMaxNivel
    NameCiuMax = Mayores.MaxCiudadano
    NameCriMax = Mayores.MaxCriminal
    CntCiuOn = Mayores.OnlineCiudadano
    CntCriOn = Mayores.OnlineCriminal
    NameOroOnline = Mayores.MaxOroOnline
    NameOroOffline = Mayores.MaxOro
   
End Sub
Private Sub salir_click()

    Unload Me

End Sub

Private Sub salir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)


Set salir.MouseIcon = Iconos.Ico_Mano

End Sub
