VERSION 5.00
Begin VB.Form frmSastre 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11520
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
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   768
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      ForeColor       =   &H0000FFFF&
      Height          =   2205
      Left            =   360
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   255
      Width           =   10890
   End
   Begin VB.Image CmdSalir 
      Height          =   525
      Left            =   645
      MousePointer    =   99  'Custom
      Top             =   2610
      Width           =   1275
   End
   Begin VB.Image CmdTejer 
      Height          =   525
      Left            =   9825
      MousePointer    =   99  'Custom
      Top             =   2625
      Width           =   1275
   End
End
Attribute VB_Name = "frmSastre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSalir_Click()
  Dim m As Integer
  
 For m = 0 To UBound(ObjSastre)
                ObjSastre(m) = 0
   Next m
     
   NumSastre = 0
     
   Unload frmSastre
End Sub

Private Sub CmdSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdSalir.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub CmdTejer_Click()
    Dim m As Integer
   
   If List1.ListIndex = -1 Then
        Unload Me
        Exit Sub
   End If
   
   Call SendData("COMSAST" & ObjSastre(List1.ListIndex))
   
   For m = 0 To UBound(ObjSastre)
                ObjSastre(m) = 0
   Next m
     
   NumSastre = 0
     
   Unload frmSastre
End Sub

Private Sub CmdTejer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     CmdTejer.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Load()
     frmSastre.Picture = Interfaces.FrmSastre_Principal
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmSastre.MouseIcon = Iconos.Ico_Diablo
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    List1.MouseIcon = Iconos.Ico_Mano
End Sub
