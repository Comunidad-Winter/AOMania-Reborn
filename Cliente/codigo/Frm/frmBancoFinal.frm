VERSION 5.00
Begin VB.Form frmBancoFinal 
   BorderStyle     =   0  'None
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
   MousePointer    =   99  'Custom
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image cmdNo 
      Height          =   600
      Left            =   4245
      MousePointer    =   99  'Custom
      Top             =   3735
      Width           =   840
   End
   Begin VB.Image cmdSi 
      Height          =   585
      Left            =   1200
      MousePointer    =   99  'Custom
      Top             =   3735
      Width           =   495
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
   Begin VB.Label LblBanco 
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
Attribute VB_Name = "frmBancoFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNo_Click()
    Call Audio.PlayWave(SND_CLICK)
    Unload frmBancoFinal
End Sub

Private Sub cmdNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdNo.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub cmdSi_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("BANKVOL")
    Unload frmBancoFinal
End Sub

Private Sub cmdSi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  cmdSi.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Load()
    Set frmBancoFinal.Picture = Interfaces.FrmBancoFinal_Principal
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Set frmBancoFinal.MouseIcon = Iconos.Ico_Diablo
End Sub
