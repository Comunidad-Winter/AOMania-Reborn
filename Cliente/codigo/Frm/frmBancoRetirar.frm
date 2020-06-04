VERSION 5.00
Begin VB.Form frmBancoRetirar 
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
   MousePointer    =   99  'Custom
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   375
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3945
      Width           =   3255
   End
   Begin VB.Image cmdRetirar 
      Height          =   270
      Left            =   3795
      MousePointer    =   99  'Custom
      Top             =   3930
      Width           =   1080
   End
   Begin VB.Image cmdCerrar 
      Height          =   360
      Left            =   5010
      MousePointer    =   99  'Custom
      Top             =   4140
      Width           =   990
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
Attribute VB_Name = "frmBancoRetirar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCerrar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Unload frmBancoRetirar
End Sub

Private Sub cmdCerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdCerrar.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub cmdRetirar_Click()
     Call Audio.PlayWave(SND_CLICK)
     
     If Text1 = "" Or Val(Text1) = 0 Then
         MsgBox "Debes especificar una cantidad valida.", vbInformation
         Exit Sub
     ElseIf Text1 > lblBanco.Caption Then
         MsgBox "No tienes suficiente oro para retirar.", vbInformation
         Exit Sub
     End If
     
     Call SendData("RETBANK" & Text1)
     Unload frmBancoRetirar
     
End Sub

Private Sub cmdRetirar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     cmdRetirar.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Load()
   Set frmBancoRetirar.Picture = Interfaces.FrmBancoRetirar_Principal
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmBancoRetirar.MouseIcon = Iconos.Ico_Diablo
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.MouseIcon = Iconos.Ico_Mano
End Sub
