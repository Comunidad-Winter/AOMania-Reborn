VERSION 5.00
Begin VB.Form frmSeguridadPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4410
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
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   294
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Aceptar"
      Height          =   300
      Left            =   1320
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2160
      Width           =   1800
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Respuesta secreta"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña Banco"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¡IMPORTANTE! Seguridad del personaje."
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4380
   End
End
Attribute VB_Name = "frmSeguridadPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnviar_Click()
   Call Audio.PlayWave(SND_CLICK)
    
    If Text1 = "" Then
        MsgBox "Faltan datos en contraseña banco.", vbInformation
        Exit Sub
     ElseIf Text2 = "" Then
         MsgBox "Faltan datos en recuperar contraseña.", vbInformation
         Exit Sub
    End If
    
    Call SendData("/SEGPJ" & Text1 & "," & Text2)
    
    Unload frmSeguridadPersonaje

End Sub

Private Sub cmdEnviar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      cmdEnviar.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmSeguridadPersonaje.MouseIcon = Iconos.Ico_Diablo
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text2.MouseIcon = Iconos.Ico_Mano
End Sub
