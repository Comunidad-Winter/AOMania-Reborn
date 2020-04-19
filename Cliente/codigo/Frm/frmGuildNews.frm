VERSION 5.00
Begin VB.Form frmGuildNews 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "GuildNews"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   6405
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Clanes aliados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   4575
      Begin VB.ListBox aliados 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H8000000E&
         Height          =   1005
         ItemData        =   "frmGuildNews.frx":0000
         Left            =   120
         List            =   "frmGuildNews.frx":0002
         TabIndex        =   6
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Clanes con los que estamos en guerra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   4575
      Begin VB.ListBox guerra 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H8000000E&
         Height          =   1005
         ItemData        =   "frmGuildNews.frx":0004
         Left            =   120
         List            =   "frmGuildNews.frx":0006
         TabIndex        =   4
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "GuildNews"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      Begin VB.TextBox news 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000009&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
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
      Left            =   1080
      MouseIcon       =   "frmGuildNews.frx":0008
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   2895
   End
End
Attribute VB_Name = "frmGuildNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
    Call Audio.PlayWave(SND_CLICK)
    Unload Me
    frmMain.SetFocus

End Sub

Public Sub ParseGuildNews(ByVal s As String)

    news = Replace(ReadField(1, s, Asc("¬")), "º", vbCrLf)

    Dim h%, j%

    h% = Val(ReadField(2, s, Asc("¬")))

    For j% = 1 To h%
    
        guerra.AddItem ReadField(j% + 2, s, Asc("¬"))
    
    Next j%

    j% = j% + 2

    h% = Val(ReadField(j%, s, Asc("¬")))

    For j% = j% + 1 To j% + h%
    
        aliados.AddItem ReadField(j%, s, Asc("¬"))
    
    Next j%

    Me.Show , frmMain

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set Command1.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Load()
    Set Me.MouseIcon = Iconos.Ico_Diablo
End Sub
