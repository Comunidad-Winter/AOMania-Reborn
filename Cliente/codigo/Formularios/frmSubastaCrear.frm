VERSION 5.00
Begin VB.Form frmSubastaCrear 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
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
   Picture         =   "frmSubastaCrear.frx":0000
   ScaleHeight     =   273
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox listDuration 
      Height          =   315
      Left            =   2985
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Text            =   "listDuration"
      Top             =   2565
      Width           =   3015
   End
   Begin VB.TextBox txtPrecioInicial 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   3015
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1560
      Width           =   2670
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   960
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   675
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   210
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   450
      Width           =   480
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2790
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Image cmdCancelar 
      Height          =   450
      Left            =   2925
      MousePointer    =   99  'Custom
      Top             =   3615
      Width           =   1320
   End
   Begin VB.Image cmdAceptar 
      Height          =   450
      Left            =   4620
      MousePointer    =   99  'Custom
      Top             =   3615
      Width           =   1320
   End
End
Attribute VB_Name = "frmSubastaCrear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
        
    Call Audio.PlayWave(SND_CLICK)
        
    Dim i As Integer
        
    i = List1.ListIndex + 1
         
    Call SendData("CRSUB" & Inventario.ObjIndex(i) & "," & txtCantidad.Text & "," & txtPrecioInicial.Text & "," & listDuration.Text)
         
End Sub

Private Sub cmdAceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     cmdAceptar.MouseIcon = Iconos.Mano
End Sub

Private Sub cmdCancelar_Click()
      Call Audio.PlayWave(SND_CLICK)
      Unload Me
End Sub

Private Sub cmdCancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdCancelar.MouseIcon = Iconos.Mano
End Sub

Private Sub Form_Load()

    Dim i As Long
    
    i = 1

    Do While i <= MAX_INVENTORY_SLOTS

        If Inventario.ObjIndex(i) <> 0 Then

            List1.AddItem Inventario.ItemName(i)
        Else
            List1.AddItem "Nada"

        End If

        i = i + 1
    Loop
    
    listDuration.AddItem "1"
    listDuration.AddItem "2"
    listDuration.AddItem "3"
    listDuration.AddItem "6"
    listDuration.AddItem "8"
    listDuration.AddItem "12"
    listDuration.AddItem "24"
    listDuration.AddItem "48"
    listDuration.ListIndex = 0
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Me.MouseIcon = Iconos.Diablo
End Sub

Private Sub List1_Click()
    
    Dim SR As RECT, DR As RECT

    SR.Left = 0
    SR.Top = 0
    SR.Right = 32
    SR.Bottom = 32

    DR.Left = 0
    DR.Top = 0
    DR.Right = 32
    DR.Bottom = 32
    
    Call DrawGrhtoHdc(Picture1.hdc, Inventario.GrhIndex(List1.ListIndex + 1), DR)
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      List1.MouseIcon = Iconos.Diablo
End Sub

Private Sub txtCantidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     txtCantidad.MouseIcon = Iconos.Mano
End Sub

Private Sub txtPrecioInicial_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      txtPrecioInicial.MouseIcon = Iconos.Mano
End Sub
