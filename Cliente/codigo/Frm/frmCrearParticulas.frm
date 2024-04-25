VERSION 5.00
Begin VB.Form frmCrearParticulas 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "frmCrearParticulas"
   ClientHeight    =   3600
   ClientLeft      =   12255
   ClientTop       =   5955
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Salir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Crear 
      Caption         =   "Crear"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Height          =   735
      Index           =   2
      Left            =   2640
      TabIndex        =   2
      Text            =   "Coord Y"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text 
      Height          =   735
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Text            =   "Coord X"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text 
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Text            =   "Nº Particula"
      Top             =   2760
      Width           =   975
   End
End
Attribute VB_Name = "frmCrearParticulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Crear_Click()

Particle_Create Val(frmCrearParticulas.Text(0)), Val(frmCrearParticulas.Text(1)), Val(frmCrearParticulas.Text(2))

End Sub

Private Sub Salir_Click()

    Unload Me

End Sub
