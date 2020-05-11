VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Launcher AoMania Reborn"
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   676
   StartUpPosition =   3  'Windows Default
   Begin VB.Image cmdCerrar 
      Height          =   555
      Left            =   9345
      Top             =   270
      Width           =   555
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' LauncherAoM 1.0.0
' By Bassinger [www.AoMania.Net]

Option Explicit

Private Sub cmdCerrar_Click()
    UnloadAllForms
End Sub

Private Sub cmdCerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
     With frmMain
         .MouseIcon = Iconos.Ico_Mano
     End With
     
End Sub

Private Sub Form_Load()
     
     With frmMain
         .MouseIcon = Iconos.Ico_Diablo
     End With
     
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
     With frmMain
         .MouseIcon = Iconos.Ico_Diablo
     End With
     
End Sub
