VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Tirar Item"
   ClientHeight    =   1710
   ClientLeft      =   1575
   ClientTop       =   4275
   ClientWidth     =   4470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   1710
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
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
      Height          =   375
      Left            =   220
      TabIndex        =   0
      Top             =   650
      Width           =   4020
   End
   Begin VB.Image salir 
      Height          =   255
      Left            =   4200
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   600
      MousePointer    =   99  'Custom
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Image Command2 
      Height          =   255
      Left            =   2280
      MousePointer    =   99  'Custom
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Call Audio.PlayWave(SND_CLICK)

    frmCantidad.Visible = False
    SendData "OH" & Inventario.SelectedItem & "," & frmCantidad.Text1.Text
    frmCantidad.Text1.Text = "0"

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Command1.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command2_Click()
    
    Call Audio.PlayWave(SND_CLICK)
    

    frmCantidad.Visible = False

    If Inventario.SelectedItem <> FLAGORO Then
        SendData "OH" & Inventario.SelectedItem & "," & Inventario.Amount(Inventario.SelectedItem)
    Else
        SendData "OH" & Inventario.SelectedItem & "," & UserGLD

    End If

    frmCantidad.Text1.Text = "0"

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Command2.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Deactivate()
     'Unload Me
End Sub

Private Sub salir_Click()

    Unload Me
    
End Sub

Private Sub salir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Set salir.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub text1_Change()

    On Error GoTo ErrHandler

    If Val(Text1.Text) < 0 Then
        Text1.Text = MAX_INVENTORY_OBJS

    End If
    
    If Val(Text1.Text) > MAX_INVENTORY_OBJS Then
        If Inventario.SelectedItem <> FLAGORO Or Val(Text1.Text) > UserGLD Then
            Text1.Text = "1"

        End If

    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    Text1.Text = "1"

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0

        End If

    End If

End Sub

Private Sub Form_Load()
    Set frmCantidad.Picture = Interfaces.FrmCantidad_Principal
    Set Me.MouseIcon = Iconos.Ico_Diablo
End Sub
