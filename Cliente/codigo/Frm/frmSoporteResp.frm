VERSION 5.00
Begin VB.Form frmSoporteResp 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Height          =   2055
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1200
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   2760
      MouseIcon       =   "frmSoporteResp.frx":0000
      MousePointer    =   99  'Custom
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STAFF"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "frmSoporteResp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    Set frmSoporteResp.Picture = Interfaces.FrmSoporte_Principal
    Me.Top = 0
    Me.Left = 0

End Sub

Private Sub Image1_Click()

    SendData "/RESETSOP"
    frmMain.Label10.Visible = False
    Unload frmSoporteResp

End Sub
