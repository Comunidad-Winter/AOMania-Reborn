VERSION 5.00
Begin VB.Form frmGuildBrief 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalles del Clan"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   7455
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
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
   ScaleHeight     =   7290
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ofrecer Paz"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MouseIcon       =   "frmGuildBrief.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton aliado 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ofrecer Alianza"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MouseIcon       =   "frmGuildBrief.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Guerra 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Declarar Guerra"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      MouseIcon       =   "frmGuildBrief.frx":02A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Solicitar Ingreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      MouseIcon       =   "frmGuildBrief.frx":03F6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmGuildBrief.frx":0548
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   5400
      Width           =   7215
      Begin VB.TextBox Desc 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Codex"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      TabIndex        =   9
      Top             =   2970
      Width           =   7215
      Begin VB.Label Codex 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   6735
      End
      Begin VB.Label Codex 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   6735
      End
      Begin VB.Label Codex 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   6735
      End
      Begin VB.Label Codex 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   6735
      End
      Begin VB.Label Codex 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   6735
      End
      Begin VB.Label Codex 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   6735
      End
      Begin VB.Label Codex 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   6735
      End
      Begin VB.Label Codex 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   6735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Info del clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2940
      Left            =   120
      TabIndex        =   0
      Top             =   -15
      Width           =   7215
      Begin VB.Label antifaccion 
         BackColor       =   &H00404040&
         Caption         =   "Puntos Antifaccion:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2640
         Width           =   6975
      End
      Begin VB.Label Aliados 
         BackColor       =   &H00404040&
         Caption         =   "Clanes Aliados:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   6975
      End
      Begin VB.Label Enemigos 
         BackColor       =   &H00404040&
         Caption         =   "Clanes Enemigos:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   6975
      End
      Begin VB.Label lblAlineacion 
         BackColor       =   &H00404040&
         Caption         =   "Alineacion:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   6975
      End
      Begin VB.Label eleccion 
         BackColor       =   &H00404040&
         Caption         =   "Elecciones:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   6975
      End
      Begin VB.Label Miembros 
         BackColor       =   &H00404040&
         Caption         =   "Miembros:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   6975
      End
      Begin VB.Label web 
         BackColor       =   &H00404040&
         Caption         =   "Web site:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   6975
      End
      Begin VB.Label lider 
         BackColor       =   &H00404040&
         Caption         =   "Lider:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   6975
      End
      Begin VB.Label creacion 
         BackColor       =   &H00404040&
         Caption         =   "Fecha de creacion:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   6975
      End
      Begin VB.Label fundador 
         BackColor       =   &H00404040&
         Caption         =   "Fundador:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   6975
      End
      Begin VB.Label nombre 
         BackColor       =   &H00404040&
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6975
      End
   End
End
Attribute VB_Name = "frmGuildBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EsLeader As Boolean

Public Sub ParseGuildInfo(ByVal buffer As String)

    If Not EsLeader Then
        guerra.Visible = False
        aliado.Visible = False
        Command3.Visible = False
    Else
        guerra.Visible = True
        aliado.Visible = True
        Command3.Visible = True

    End If

    nombre.Caption = "Nombre:" & ReadField(1, buffer, Asc("¬"))
    fundador.Caption = "Fundador:" & ReadField(2, buffer, Asc("¬"))
    creacion.Caption = "Fecha de creacion:" & ReadField(3, buffer, Asc("¬"))
    lider.Caption = "Lider:" & ReadField(4, buffer, Asc("¬"))
    web.Caption = "Web site:" & ReadField(5, buffer, Asc("¬"))
    Miembros.Caption = "Miembros:" & ReadField(6, buffer, Asc("¬"))
    eleccion.Caption = "Dias para proxima eleccion de lider:" & ReadField(7, buffer, Asc("¬"))
    'Oro.Caption = "Oro:" & ReadField(8, Buffer, Asc("¬"))
    Enemigos.Caption = "Clanes enemigos:" & ReadField(8, buffer, Asc("¬"))
    aliados.Caption = "Clanes aliados:" & ReadField(9, buffer, Asc("¬"))
    antifaccion.Caption = "Puntos Antifaccion: " & ReadField(10, buffer, Asc("¬"))

    Dim T As Long

    For T = 1 To 8
        Codex(T - 1).Caption = ReadField(11 + T, buffer, Asc("¬"))
    Next T

    Dim des As String

    des = ReadField(20, buffer, Asc("¬"))
    Desc.Text = Replace(des, "º", vbCrLf)

    Me.Show vbModeless, frmMain

End Sub

Private Sub aliado_Click()

    frmCommet.nombre = Right(nombre.Caption, Len(nombre.Caption) - 7)
    frmCommet.T = ALIANZA
    frmCommet.Caption = "Ingrese propuesta de alianza"
    Call frmCommet.Show(, frmGuildBrief)

    'Call SendData("OFRECALI" & Right(Nombre, Len(Nombre) - 7))
    'Unload Me
End Sub

Private Sub aliado_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set aliado.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set Command1.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command2_Click()

    Call frmGuildSol.RecieveSolicitud(Right$(nombre, Len(nombre) - 7))
    Call frmGuildSol.Show(, frmGuildBrief)
    'Unload Me

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set Command2.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command3_Click()

    frmCommet.nombre = Right(nombre.Caption, Len(nombre.Caption) - 7)
    frmCommet.T = PAZ
    frmCommet.Caption = "Ingrese propuesta de paz"
    Call frmCommet.Show(, frmGuildBrief)

    'Unload Me
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set Command3.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Load()

    Set Me.MouseIcon = Iconos.Ico_Diablo
  
End Sub

Private Sub Guerra_Click()
     
    frmMain.SetFocus

    Call SendData("DECGUERR" & Right(nombre.Caption, Len(nombre.Caption) - 7))

    Unload Me

End Sub

Private Sub Guerra_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set guerra.MouseIcon = Iconos.Ico_Mano
End Sub
