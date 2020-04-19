VERSION 5.00
Begin VB.Form frmHerrero 
   BorderStyle     =   0  'None
   Caption         =   "Herrero"
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   263
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   298
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstArmas 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1620
      Left            =   360
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1260
      Width           =   3960
   End
   Begin VB.ListBox lstArmaduras 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   360
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1260
      Width           =   3960
   End
   Begin VB.Image Command3 
      Height          =   405
      Left            =   2205
      MousePointer    =   99  'Custom
      Top             =   3375
      Width           =   1845
   End
   Begin VB.Image Command4 
      Height          =   390
      Left            =   510
      MousePointer    =   99  'Custom
      Top             =   3375
      Width           =   960
   End
   Begin VB.Image Command2 
      Height          =   360
      Left            =   2175
      MousePointer    =   99  'Custom
      Top             =   390
      Width           =   2040
   End
   Begin VB.Image Command1 
      Height          =   360
      Left            =   405
      MousePointer    =   99  'Custom
      Top             =   390
      Width           =   1140
   End
End
Attribute VB_Name = "frmHerrero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()

    lstArmaduras.Visible = False
    lstArmas.Visible = True

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command1.MouseIcon = Iconos.Mano
End Sub

Private Sub Command2_Click()

    lstArmaduras.Visible = True
    lstArmas.Visible = False

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Command2.MouseIcon = Iconos.Mano
End Sub

Private Sub Command3_Click()

    If lstArmas.ListIndex = -1 Then
        Unload Me
        Exit Sub
    End If

    If lstArmas.Visible Then
        Call SendData("CNS" & ArmasHerrero(lstArmas.ListIndex))
    Else
        Call SendData("CNS" & ArmadurasHerrero(lstArmaduras.ListIndex))

    End If

    Unload Me

End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Command3.MouseIcon = Iconos.Mano
End Sub

Private Sub Command4_Click()

    Unload Me

End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command4.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Deactivate()

    'Me.SetFocus
End Sub

Private Sub Form_Load()
    frmHerrero.Picture = Interfaces.FrmHerrero_Principal
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmHerrero.MouseIcon = Iconos.Ico_Cruceta
End Sub

Private Sub lstArmaduras_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        lstArmaduras.MouseIcon = Iconos.Ico_Mano
End Sub


Private Sub lstArmas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      lstArmas.MouseIcon = Iconos.Ico_Mano
End Sub
