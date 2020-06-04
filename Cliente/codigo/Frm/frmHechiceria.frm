VERSION 5.00
Begin VB.Form frmHechiceria 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
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
   ScaleHeight     =   229
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
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
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Text            =   "1"
      Top             =   2550
      Width           =   855
   End
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
      Height          =   2205
      Left            =   270
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   240
      Width           =   4080
   End
   Begin VB.Image CmdFabricar 
      Height          =   450
      Left            =   2520
      MousePointer    =   99  'Custom
      Top             =   2865
      Width           =   1980
   End
   Begin VB.Image CmdSalir 
      Height          =   540
      Left            =   480
      MousePointer    =   99  'Custom
      Top             =   2805
      Width           =   1275
   End
End
Attribute VB_Name = "frmHechiceria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdFabricar_Click()
   Dim m As Integer
   
   If List1.ListIndex = -1 Then
        Unload Me
        Exit Sub
   End If
   
   Call SendData("COMHECH" & ObjHechizeria(List1.ListIndex) & "," & Text1)
   
   For m = 0 To UBound(ObjSastre)
                ObjHechizeria(m) = 0
   Next m
     
   NumHechizeria = 0
     
   Unload frmHechiceria

End Sub

Private Sub CmdFabricar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     CmdFabricar.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub CmdSalir_Click()
     Dim m As Integer
  
 For m = 0 To UBound(ObjHechizeria)
                ObjHechizeria(m) = 0
   Next m
     
   NumHechizeria = 0
     
   Unload frmHechiceria
End Sub

Private Sub CmdSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdSalir.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Load()
    frmHechiceria.Picture = Interfaces.FrmHechiceria_Principal
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmHechiceria.MouseIcon = Iconos.Ico_Diablo
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     List1.MouseIcon = Iconos.Ico_Mano
End Sub
