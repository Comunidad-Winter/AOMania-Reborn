VERSION 5.00
Begin VB.Form frmOlvidarHechizo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
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
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   242
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3705
      Width           =   3015
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
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
      Height          =   3180
      Left            =   285
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   300
      Width           =   3135
   End
   Begin VB.Image cmdOlvidar 
      Height          =   435
      Left            =   90
      MousePointer    =   99  'Custom
      Top             =   4065
      Width           =   1440
   End
   Begin VB.Image cmdCancelar 
      Height          =   435
      Left            =   1830
      MousePointer    =   99  'Custom
      Top             =   4065
      Width           =   1770
   End
End
Attribute VB_Name = "frmOlvidarHechizo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
      Unload frmOlvidarHechizo
End Sub

Private Sub cmdCancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    List1.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub cmdOlvidar_Click()
    Dim Slot As Integer
    
    If List1.ListIndex = -1 Then
        Call AddtoRichTextBox(frmMain.RecTxt, "¡Primero selecciona el hechizo!", 65, 190, 156, False, False, False)
        Exit Sub
    ElseIf Text1 = "" Then
        MsgBox "Debe poner la respuesta correcta", vbInformation
        Exit Sub
    End If
    
    Slot = List1.ListIndex + 1
    
    Call SendData("OLVHECA" & Slot & "," & Text1)
    
    Unload frmOlvidarHechizo
    
End Sub

Private Sub cmdOlvidar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    List1.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Load()
     frmOlvidarHechizo.Picture = Interfaces.FrmOlvidarHechizo_Principal
     Call SendData("ENVHECA")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     frmOlvidarHechizo.MouseIcon = Iconos.Ico_Diablo
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     List1.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Text1.MouseIcon = Iconos.Ico_Mano
End Sub
