VERSION 5.00
Begin VB.Form frmEsc 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Exit 
      Height          =   375
      Left            =   4800
      MouseIcon       =   "frmEsc.frx":0000
      MousePointer    =   99  'Custom
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Options 
      Height          =   735
      Left            =   960
      MouseIcon       =   "frmEsc.frx":0CCA
      MousePointer    =   99  'Custom
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Image ExitGame 
      Height          =   735
      Left            =   960
      MouseIcon       =   "frmEsc.frx":1994
      MousePointer    =   99  'Custom
      Top             =   960
      Width           =   3375
   End
End
Attribute VB_Name = "frmEsc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub Exit_Click()
    Unload frmEsc
    Set clsFormulario = Nothing
    frmMain.StatusESC = 0

End Sub

Private Sub Exit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    frmEsc.MouseIcon = LoadPicture(App.Path & "\interfaces\mano.ico")
    frmEsc.MousePointer = vbCustom

End Sub

Private Sub ExitGame_Click()
    Unload frmEsc
    frmMain.StatusESC = 0
    Set clsFormulario = Nothing
    Call SendData("/SALIR")

End Sub

Private Sub ExitGame_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
                               
    frmEsc.MouseIcon = LoadPicture(App.Path & "\interfaces\mano.ico")
    frmEsc.MousePointer = vbCustom

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'frmEsc.MouseIcon = LoadPicture(App.Path & "\interfaces\diablo.ico")
    'frmEsc.MousePointer = vbCustom

End Sub

Private Sub Options_Click()
    Unload frmEsc
    frmMain.StatusESC = 0
    Set clsFormulario = Nothing
    frmOpciones.Show

End Sub

Private Sub Options_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    frmEsc.MouseIcon = LoadPicture(App.Path & "\interfaces\mano.ico")
    frmEsc.MousePointer = vbCustom

End Sub

