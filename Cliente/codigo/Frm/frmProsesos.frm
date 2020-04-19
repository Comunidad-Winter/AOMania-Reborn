VERSION 5.00
Begin VB.Form frmProsesos 
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Foto"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   6240
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   6240
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   5715
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "frmProsesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    Call SacarScreen

End Sub

