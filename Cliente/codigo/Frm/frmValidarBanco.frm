VERSION 5.00
Begin VB.Form frmValidarBanco 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   248
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Enviar"
      Height          =   300
      Left            =   2400
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1440
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1440
      Width           =   1080
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1680
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Palabra Secreta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmValidarBanco.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmValidarBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
        Call Audio.PlayWave(SND_CLICK)
        Unload frmValidarBanco
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Command1.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command2_Click()
        Call Audio.PlayWave(SND_CLICK)
        Call SendData("VALIDBANK" & Text1)
        Unload frmValidarBanco
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Command2.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Load()
    frmValidarBanco.Caption = "Validar Banco de " & CharList(UserCharIndex).Nombre
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     frmValidarBanco.MouseIcon = Iconos.Ico_Diablo
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.MouseIcon = Iconos.Ico_Mano
End Sub
