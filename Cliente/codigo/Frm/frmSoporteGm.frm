VERSION 5.00
Begin VB.Form frmSoporteGm 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   360
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Escribe aqui tu respuesta."
      Top             =   3840
      Width           =   6855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   720
      Width           =   6855
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1320
      MouseIcon       =   "frmSoporteGm.frx":0000
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3960
      MouseIcon       =   "frmSoporteGm.frx":0CCA
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   630
   End
End
Attribute VB_Name = "frmSoporteGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Set frmSoporteGm.Picture = Interfaces.FrmSoporteGM_Principal
    Me.Top = 0
    Me.Left = 0

End Sub

Private Sub Image1_Click()

    Unload frmSoporteGm

End Sub

Private Sub Image2_Click()

    SendData ("SOSDONE" & Label1.Caption)
    Unload frmSoporteGm

End Sub

