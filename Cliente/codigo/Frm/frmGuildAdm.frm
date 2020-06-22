VERSION 5.00
Begin VB.Form frmGuildAdm 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Lista de Clanes Registrados"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   3990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   3435
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
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
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmGuildAdm.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lista Clanes"
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
      Left            =   1200
      MouseIcon       =   "frmGuildAdm.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
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
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "frmGuildAdm.frx":02A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Clanes"
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
      Height          =   2655
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.ListBox GuildsList 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   1950
         ItemData        =   "frmGuildAdm.frx":03F6
         Left            =   240
         List            =   "frmGuildAdm.frx":03F8
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmGuildAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()

    'If GuildsList.ListIndex = 0 Then Exit Sub
    Call SendData("CLANDETAILS" & guildslist.List(guildslist.ListIndex))

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Command1.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command2_Click()
    Unload Me
    Call frmListClanes.Show(vbModeless, frmMain)
    Call frmListClanes.ParseListClanes
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command2.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command3_Click()

    Unload Me

End Sub

Public Sub ParseGuildList(ByVal rData As String)

    Dim j As Integer, k As Integer

    For j = 0 To guildslist.ListCount - 1
        Me.guildslist.RemoveItem 0
    Next j

    k = CInt(readfield2(1, rData, 44))

    For j = 1 To k
        guildslist.AddItem readfield2(1 + j, rData, 44)
    Next j

    Me.Show , frmMain

End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Command3.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Load()
    Set Me.MouseIcon = Iconos.Ico_Diablo
End Sub

Private Sub GuildsList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    guildslist.MouseIcon = Iconos.Ico_Diablo
End Sub
