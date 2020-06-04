VERSION 5.00
Begin VB.Form frmGuildLeader 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Administración del Clan"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   5880
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   7065
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
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
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5970
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Propuestas de paz"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5400
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Editar URL de la web del clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":02A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Editar Codex o Descripcion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":03F6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4290
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2655
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command9 
         Caption         =   "Clasificación por Puntos"
         Height          =   360
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   2280
         Width           =   2655
      End
      Begin VB.ListBox guildslist 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000004&
         Height          =   1395
         ItemData        =   "frmGuildLeader.frx":0548
         Left            =   120
         List            =   "frmGuildLeader.frx":054A
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detalles"
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
         MouseIcon       =   "frmGuildLeader.frx":054C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1800
         Width           =   2655
      End
   End
   Begin VB.Frame txtnews 
      BackColor       =   &H00000000&
      Caption         =   "GuildNews"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1575
      Left            =   0
      TabIndex        =   6
      Top             =   2640
      Width           =   5775
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Actualizar"
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
         MouseIcon       =   "frmGuildLeader.frx":069E
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txtguildnews 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000004&
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Miembros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2295
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detalles"
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
         MouseIcon       =   "frmGuildLeader.frx":07F0
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ListBox members 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000004&
         Height          =   1395
         ItemData        =   "frmGuildLeader.frx":0942
         Left            =   120
         List            =   "frmGuildLeader.frx":0944
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Solicitudes de ingreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Width           =   2895
      Begin VB.CommandButton cmdElecciones 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Abrir elecciones"
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
         MouseIcon       =   "frmGuildLeader.frx":0946
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1935
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detalles"
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
         MouseIcon       =   "frmGuildLeader.frx":0A98
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1170
         Width           =   2655
      End
      Begin VB.ListBox solicitudes 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000005&
         Height          =   810
         ItemData        =   "frmGuildLeader.frx":0BEA
         Left            =   120
         List            =   "frmGuildLeader.frx":0BEC
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Miembros 
         BackColor       =   &H00404040&
         Caption         =   "El clan cuenta con x miembros"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1620
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmGuildLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Private Sub cmdElecciones_Click()

'    Call SendData("ABREELEC")
'    Unload Me

'End Sub

Private Sub Command1_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmCharInfo.frmsolicitudes = True
    Call SendData("1HRINFO<" & solicitudes.List(solicitudes.ListIndex))

    'Unload Me

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set Command1.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command2_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmCharInfo.frmmiembros = True
    Call SendData("1HRINFO<" & members.List(members.ListIndex))

    'Unload Me

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set Command2.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command3_Click()
    Call Audio.PlayWave(SND_CLICK)
    Dim k$

    k$ = Replace(txtguildnews, vbCrLf, "º")

    Call SendData("ACTGNEWS" & k$)

End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set Command3.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command4_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmGuildBrief.EsLeader = True
    Call SendData("CLANDETAILS" & guildslist.List(guildslist.ListIndex))

    'Unload Me

End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set Command4.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command5_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call frmGuildDetails.Show(vbModal, frmGuildLeader)

    'Unload Me

End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set Command5.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command6_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call frmGuildURL.Show(vbModeless, frmGuildLeader)

    'Unload Me
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set Command6.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command7_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("ENVPROPP")

End Sub


Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set Command7.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command8_Click()
    Call Audio.PlayWave(SND_CLICK)
    Unload Me
    frmMain.SetFocus

End Sub

Public Sub ParseLeaderInfo(ByVal Data As String)

    If Me.Visible Then Exit Sub

    Dim r%, T%

    r% = Val(ReadField(1, Data, Asc("¬")))

    For T% = 1 To r%
        guildslist.AddItem ReadField(1 + T%, Data, Asc("¬"))
    Next T%

    r% = Val(ReadField(T% + 1, Data, Asc("¬")))
    Miembros.Caption = "El clan cuenta con " & r% & " miembros."

    Dim k%

    For k% = 1 To r%
        members.AddItem ReadField(T% + 1 + k%, Data, Asc("¬"))
    Next k%

    txtguildnews = Replace(ReadField(T% + k% + 1, Data, Asc("¬")), "º", vbCrLf)

    T% = T% + k% + 2

    r% = Val(ReadField(T%, Data, Asc("¬")))

    For k% = 1 To r%
        solicitudes.AddItem ReadField(T% + k%, Data, Asc("¬"))
    Next k%

    Me.Show vbModeless, frmMain

End Sub

Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set Command8.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command9_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call frmListClanes.Show(vbModeless, frmMain)
    Call frmListClanes.ParseListClanes
End Sub

Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set Command9.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Deactivate()

    'Me.SetFocus
End Sub

Private Sub Form_Load()
    Set Me.MouseIcon = Iconos.Ico_Diablo
End Sub

