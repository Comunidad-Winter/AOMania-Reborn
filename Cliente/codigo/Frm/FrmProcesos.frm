VERSION 5.00
Begin VB.Form FrmProcesos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Matar Proceso"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Foto"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   615
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3540
      ItemData        =   "FrmProcesos.frx":0000
      Left            =   240
      List            =   "FrmProcesos.frx":0002
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   4800
      Width           =   2535
   End
End
Attribute VB_Name = "FrmProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    SacarScreen

End Sub

Private Sub Command3_Click()

    Dim Nombree As String
    Dim Proceso As String
    Nombree = UCase(Label1.Caption)
    Proceso = UCase(Label2.Caption)
    Call SendData("/MATARPROCESO " & Nombree & "," & Proceso)

End Sub

Private Sub Form_Load()

    'V
End Sub

Private Sub List1_Click()

    Label2.Caption = UCase(List1.Text)

End Sub

