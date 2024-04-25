VERSION 5.00
Begin VB.Form frmExpulsarClan 
   Caption         =   "Expulsar miembro clan"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5340
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
   ScaleHeight     =   4305
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Contraseña de Clan"
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   5055
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
         Height          =   360
         Left            =   3840
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   960
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancelar"
         Height          =   360
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   960
         Width           =   990
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "usuario. Debe ser exacta minuscula y mayusculas."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deberás introducir aquí la contraseña del Clan para poder echar al"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4755
      End
   End
   Begin VB.Frame FrameGeneral 
      Caption         =   "General"
      Height          =   2415
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.Label StatusData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Left            =   840
         TabIndex        =   22
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label BancoData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         Height          =   195
         Left            =   840
         TabIndex        =   21
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label OroData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Oro"
         Height          =   195
         Left            =   600
         TabIndex        =   20
         Top             =   1560
         Width           =   270
      End
      Begin VB.Label NivelData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel"
         Height          =   195
         Left            =   720
         TabIndex        =   19
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label GeneroData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Genero"
         Height          =   195
         Left            =   840
         TabIndex        =   18
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label ClaseData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clase"
         Height          =   195
         Left            =   720
         TabIndex        =   17
         Top             =   840
         Width           =   390
      End
      Begin VB.Label RazaData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Raza"
         Height          =   195
         Left            =   720
         TabIndex        =   16
         Top             =   600
         Width           =   360
      End
      Begin VB.Label NombreData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   195
         Left            =   960
         TabIndex        =   15
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Oro:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Genero:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label Class 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clase:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Raze 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Raza:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Name 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmExpulsarClan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ParseExpulsarClan(Nombre As String, Raza As String, Clase As String, Genero As String, _
    Nivel As String, Oro As String, Banco As String, status As String)

    NombreData.Caption = readfield2(2, Nombre, 58)
    RazaData.Caption = readfield2(2, Raza, 58)
    ClaseData.Caption = readfield2(2, Clase, 58)
    GeneroData.Caption = readfield2(2, Genero, 58)
    NivelData.Caption = readfield2(2, Nivel, 58)
    OroData.Caption = readfield2(2, Oro, 58)
    BancoData.Caption = readfield2(2, Banco, 58)
    If UCase$(status) = " (CIUDADANO)" Then
        StatusData.Caption = "Ciudadano"
        StatusData.ForeColor = vbBlue
    ElseIf UCase$(status) = " (CRIMINAL)" Then
        StatusData.Caption = "Criminal"
        StatusData.ForeColor = vbRed
    End If
            
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command1.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command2_Click()
 
    If txtPassword.Text = "" Then
        MsgBox "Debes introducir una contraseña."
        Exit Sub
    End If
 
   
    Call SendData("ECHARCLA" & txtPassword.Text & "," & NombreData.Caption)
    frmCharInfo.frmmiembros = False
    frmCharInfo.frmsolicitudes = False
    Unload frmCharInfo
    Unload frmGuildLeader
    Call SendData("GLINFO")
    Unload Me
   
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command2.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Load()
    Me.MouseIcon = Iconos.Ico_Diablo
End Sub
