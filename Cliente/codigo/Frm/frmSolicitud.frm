VERSION 5.00
Begin VB.Form frmGuildSol 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ingreso"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   4305
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
   ScaleHeight     =   3675
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
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
      Left            =   240
      MouseIcon       =   "frmSolicitud.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enviar"
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
      Left            =   3240
      MouseIcon       =   "frmSolicitud.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000004&
      Height          =   1215
      Left            =   240
      MaxLength       =   400
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   $"frmSolicitud.frx":02A4
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
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmGuildSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim CName As String

Private Sub Command1_Click()

    Dim f$

    f$ = "SOLICITUD" & CName
    f$ = f$ & "," & Replace(Replace(Text1.Text, ",", ";"), vbCrLf, "º")

    Call SendData(f$)

    Unload Me

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set Command1.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command2_Click()

    Unload Me

End Sub

Public Sub RecieveSolicitud(ByVal GuildName As String)

    CName = GuildName

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set Command2.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Load()
    Set Me.MouseIcon = Iconos.Ico_Diablo
End Sub
