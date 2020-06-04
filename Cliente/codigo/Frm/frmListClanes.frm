VERSION 5.00
Begin VB.Form frmListClanes 
   Caption         =   "Mejores Clanes AoMania"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "frmListClanes"
   MousePointer    =   99  'Custom
   ScaleHeight     =   279
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   386
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox ListClan 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   720
      Width           =   3585
   End
   Begin VB.Image ImgEsc 
      Height          =   615
      Left            =   4320
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "frmListClanes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ParseListClanes()
    Call SendData("/RANKCLAN")
End Sub

Private Sub Form_Load()
    Set Me.MouseIcon = Iconos.Ico_Diablo
End Sub

Private Sub ImgEsc_Click()
    Unload Me
End Sub

Private Sub ImgEsc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set ImgEsc.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub ListClan_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ListClan.MouseIcon = Iconos.Ico_Diablo
End Sub
