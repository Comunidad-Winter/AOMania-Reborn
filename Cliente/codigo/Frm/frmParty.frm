VERSION 5.00
Begin VB.Form frmParty 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Party"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3630
   ControlBox      =   0   'False
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
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   242
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir de la Party"
      Height          =   390
      Left            =   120
      TabIndex        =   32
      Top             =   5520
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Cerrar"
      Height          =   390
      Left            =   2040
      TabIndex        =   1
      Top             =   5520
      Width           =   1470
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      Height          =   195
      Index           =   10
      Left            =   1950
      TabIndex        =   30
      Top             =   5070
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   29
      Top             =   4830
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      Height          =   195
      Index           =   9
      Left            =   1950
      TabIndex        =   27
      Top             =   4560
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   26
      Top             =   4320
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      Height          =   195
      Index           =   8
      Left            =   1950
      TabIndex        =   24
      Top             =   4050
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   23
      Top             =   3810
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      Height          =   195
      Index           =   7
      Left            =   1950
      TabIndex        =   21
      Top             =   3540
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   20
      Top             =   3300
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      Height          =   195
      Index           =   6
      Left            =   1950
      TabIndex        =   18
      Top             =   3030
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   2790
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      Height          =   195
      Index           =   5
      Left            =   1950
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      Height          =   195
      Index           =   4
      Left            =   1950
      TabIndex        =   12
      Top             =   2010
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   1770
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      Height          =   195
      Index           =   3
      Left            =   1950
      TabIndex        =   10
      Top             =   1500
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1260
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      Height          =   195
      Index           =   2
      Left            =   1950
      TabIndex        =   5
      Top             =   990
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      Height          =   195
      Index           =   1
      Left            =   1950
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   180
      Index           =   1
      Left            =   120
      Top             =   480
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   750
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmParty.frx":0000
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      Height          =   180
      Index           =   10
      Left            =   120
      TabIndex        =   31
      Top             =   5070
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      Height          =   180
      Index           =   9
      Left            =   120
      TabIndex        =   28
      Top             =   4560
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      Height          =   180
      Index           =   8
      Left            =   120
      TabIndex        =   25
      Top             =   4050
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      Height          =   180
      Index           =   7
      Left            =   120
      TabIndex        =   22
      Top             =   3540
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   3030
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   2010
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1500
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   990
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   180
      Index           =   2
      Left            =   120
      Top             =   990
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   180
      Index           =   3
      Left            =   120
      Top             =   1500
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   180
      Index           =   4
      Left            =   120
      Top             =   2010
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   180
      Index           =   5
      Left            =   120
      Top             =   2520
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   180
      Index           =   6
      Left            =   120
      Top             =   3030
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   180
      Index           =   7
      Left            =   120
      Top             =   3540
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   180
      Index           =   8
      Left            =   120
      Top             =   4050
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   180
      Index           =   9
      Left            =   120
      Top             =   4560
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   180
      Index           =   10
      Left            =   120
      Top             =   5070
      Visible         =   0   'False
      Width           =   1500
   End
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdExit_Click()
      Call Audio.PlayWave(SND_CLICK)
      Unload Me
End Sub

Private Sub CmdSalir_Click()
      Call Audio.PlayWave(SND_CLICK)
      Call SendData("/SalirParty")
      Unload Me
End Sub
