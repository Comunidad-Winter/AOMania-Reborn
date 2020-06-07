VERSION 5.00
Begin VB.Form frmCarp 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   0  'None
   Caption         =   "Carpintero"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   ControlBox      =   0   'False
   FillColor       =   &H0000FFFF&
   FillStyle       =   6  'Cross
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Text            =   "1"
      Top             =   2520
      Width           =   855
   End
   Begin VB.ListBox lstArmas 
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
      ForeColor       =   &H0000FFFF&
      Height          =   2010
      Left            =   360
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   270
      Width           =   3975
   End
   Begin VB.Image Command4 
      Height          =   525
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   2820
      Width           =   1290
   End
   Begin VB.Image Command3 
      Height          =   405
      Left            =   2400
      MousePointer    =   99  'Custom
      Top             =   2865
      Width           =   1875
   End
End
Attribute VB_Name = "frmCarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command3_Click()
    
    If lstArmas.ListIndex = -1 Then
        Unload Me
        Exit Sub
    End If
    
    If Text1 = 0 Then
        Unload Me
        Exit Sub
    End If
    
    Call SendData("CNC" & ObjCarpintero(lstArmas.ListIndex) & "," & Text1)

    Unload Me

End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Command3.MouseIcon = Iconos.Ico_Mano
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
       frmCarp.Picture = Interfaces.FrmCarp_Principal
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       frmCarp.MouseIcon = Iconos.Ico_Diablo
End Sub

Private Sub lstArmas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      lstArmas.MouseIcon = Iconos.Ico_Mano
End Sub
