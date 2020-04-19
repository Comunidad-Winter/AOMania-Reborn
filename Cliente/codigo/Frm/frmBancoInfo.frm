VERSION 5.00
Begin VB.Form frmBancoInfo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image cmdDepositar 
      Height          =   330
      Left            =   30
      MousePointer    =   99  'Custom
      Top             =   3555
      Width           =   3735
   End
   Begin VB.Image cmdRetirar 
      Height          =   255
      Left            =   30
      MousePointer    =   99  'Custom
      Top             =   3900
      Width           =   3420
   End
   Begin VB.Image cmdObj 
      Height          =   330
      Left            =   30
      MousePointer    =   99  'Custom
      Top             =   4155
      Width           =   4005
   End
   Begin VB.Image cmdCerrar 
      Height          =   360
      Left            =   5010
      MousePointer    =   99  'Custom
      Top             =   4140
      Width           =   990
   End
   Begin VB.Label LblObj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   195
      Left            =   1125
      TabIndex        =   2
      Top             =   1995
      Width           =   120
   End
   Begin VB.Label LblOro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   195
      Left            =   1125
      TabIndex        =   1
      Top             =   1530
      Width           =   120
   End
   Begin VB.Label lblBanco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   195
      Left            =   1125
      TabIndex        =   0
      Top             =   1035
      Width           =   120
   End
End
Attribute VB_Name = "frmBancoInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCerrar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Unload frmBancoInfo
End Sub

Private Sub cmdCerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdCerrar.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub cmdDepositar_Click()
      Call Audio.PlayWave(SND_CLICK)
      Call SendData("BANKDEP")
      Unload frmBancoInfo
End Sub

Private Sub cmdDepositar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     cmdDepositar.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub cmdObj_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("BANKOBJ")
    Unload frmBancoInfo
End Sub

Private Sub cmdObj_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdObj.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub cmdRetirar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("BANKRET")
    Unload frmBancoInfo
End Sub

Private Sub cmdRetirar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdRetirar.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Load()
    frmBancoInfo.Picture = Interfaces.FrmBancoInfo_Principal
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmBancoInfo.MouseIcon = Iconos.Ico_Diablo
End Sub
