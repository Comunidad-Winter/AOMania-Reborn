VERSION 5.00
Begin VB.Form frmCabezas 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5850
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
   ScaleHeight     =   291
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   480
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   555
      Width           =   2715
   End
   Begin VB.PictureBox PicHead 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4950
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   255
   End
   Begin VB.Image cmdCambiar 
      Height          =   465
      Left            =   180
      MousePointer    =   99  'Custom
      Top             =   3705
      Width           =   2790
   End
   Begin VB.Image cmdCerrar 
      Height          =   465
      Left            =   4005
      MousePointer    =   99  'Custom
      Top             =   3705
      Width           =   1380
   End
End
Attribute VB_Name = "frmCabezas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private EligeHead As Integer

Private Sub cmdCambiar_Click()
   If EligeHead = 0 Then
       MsgBox "No has elegido cabeza.", vbInformation
       Exit Sub
   End If
   
   Call SendData("CHAHEAD" & EligeHead)
   
   Unload frmCabezas
End Sub

Private Sub cmdCambiar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      cmdCambiar.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub cmdCerrar_Click()
      Unload frmCabezas
End Sub

Private Sub DrawHead(ByVal Head As Integer)

    PicHead.Picture = Nothing
    PicHead.AutoRedraw = True
    
    Dim DR As RECT
    Dim Grh As Long

    Grh = HeadData(Head).Head(3).GrhIndex

    With GrhData(Grh)
        DR.Left = 0
        DR.Top = 0
        DR.Right = 32
        DR.bottom = 32
        
        Call DrawGrhtoHdc(PicHead.hdc, Grh, DR)
    End With
    
End Sub

Private Sub cmdCerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       cmdCerrar.MouseIcon = Iconos.Mano
End Sub

Private Sub Form_Load()
        Call DrawHead(2)
        PicHead.Picture = Nothing
        List1.MouseIcon = Iconos.Ico_Mano
        frmCabezas.Picture = Interfaces.FrmCabezas_Principal
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     frmCabezas.MouseIcon = Iconos.Ico_Diablo
End Sub

Private Sub List1_Change()
      Dim Index As Integer
    
    Index = List1.ListIndex + 1
    EligeHead = Heads(Index)
   
   Call DrawHead(EligeHead)
End Sub

Private Sub List1_Click()
    Dim Index As Integer
    
    Index = List1.ListIndex + 1
    EligeHead = Heads(Index)
   
   Call DrawHead(EligeHead)
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
     List1.MouseIcon = Iconos.Ico_Mano

End Sub
