VERSION 5.00
Begin VB.Form frmHerreroMagico 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4470
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
   ScaleHeight     =   263
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   298
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
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
   Begin VB.Image CmdConstruir 
      Height          =   405
      Left            =   2205
      MousePointer    =   99  'Custom
      Top             =   3375
      Width           =   1845
   End
   Begin VB.Image CmdSalir 
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
Attribute VB_Name = "frmHerreroMagico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdConstruir_Click()
   Dim m As Integer
   
   If List1.ListIndex = -1 Then
        Unload Me
        Exit Sub
   End If
   
   Call SendData("COMHERM" & ObjHerreroMagico(List1.ListIndex))
   
   For m = 0 To UBound(ObjHerreroMagico)
                ObjHerreroMagico(m) = 0
   Next m
     
   NumHerrero = 0
     
   Unload frmHerreroMagico
   
End Sub

Private Sub CmdConstruir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdConstruir.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub CmdSalir_Click()
  Dim m As Integer
  
 For m = 0 To UBound(ObjHerreroMagico)
                ObjHerreroMagico(m) = 0
   Next m
     
   NumHerrero = 0
     
   Unload frmHerreroMagico
End Sub

Private Sub CmdSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdSalir.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command1_Click()
     Dim m As Integer
  
   For m = 0 To UBound(ObjHerreroMagico)
                ObjHerreroMagico(m) = 0
   Next m
     
   NumHerrero = 0
   List1.Clear
   
   Call SendData("ACTOBHW")
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command1.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command2_Click()
   Dim m As Integer
  
   For m = 0 To UBound(ObjHerreroMagico)
                ObjHerreroMagico(m) = 0
   Next m
     
   NumHerrero = 0
   List1.Clear
   
   Call SendData("ACTOBHA")
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Command2.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Load()
    frmHerreroMagico.Picture = Interfaces.FrmHerrero_Principal
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   List1.MouseIcon = Iconos.Ico_Mano
End Sub
