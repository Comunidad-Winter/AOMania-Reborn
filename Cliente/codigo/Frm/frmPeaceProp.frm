VERSION 5.00
Begin VB.Form frmPeaceProp 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ofertas de paz"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   4875
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
   ScaleHeight     =   2715
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rechazar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MouseIcon       =   "frmPeaceProp.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      MouseIcon       =   "frmPeaceProp.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      MouseIcon       =   "frmPeaceProp.frx":02A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmPeaceProp.frx":03F6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.ListBox lista 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   1980
      ItemData        =   "frmPeaceProp.frx":0548
      Left            =   120
      List            =   "frmPeaceProp.frx":054A
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmPeaceProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private tipoprop As TIPO_PROPUESTA

Private Enum TIPO_PROPUESTA

    ALIANZA = 1
    PAZ = 2

End Enum

Private Sub Command1_Click()

    Unload Me

End Sub

Public Sub ParsePeaceOffers(ByVal s As String)

    Dim T%, r%

    T% = Val(readfield2(1, s, 44))

    For r% = 1 To T%
        Call lista.AddItem(readfield2(r% + 1, s, 44))
    Next r%

    tipoprop = PAZ

    Me.Show vbModeless, frmMain

End Sub

Public Sub ParseAllieOffers(ByVal s As String)

    Dim T%, r%

    T% = Val(readfield2(1, s, 44))

    For r% = 1 To T%
        Call lista.AddItem(readfield2(r% + 1, s, 44))
    Next r%

    tipoprop = ALIANZA
    Me.Show vbModeless, frmMain

End Sub

Private Sub Command2_Click()

    'Me.Visible = False
    If tipoprop = PAZ Then
        Call SendData("PEACEDET" & lista.List(lista.ListIndex))
    Else
        Call SendData("ALLIEDET" & lista.List(lista.ListIndex))

    End If

End Sub

Private Sub Command3_Click()
    
    frmMain.SetFocus
    
    'Me.Visible = False
    If tipoprop = PAZ Then
        Call SendData("ACEPPEAT" & lista.List(lista.ListIndex))
    Else
        Call SendData("ACEPALIA" & lista.List(lista.ListIndex))

    End If

    Me.Hide
    Unload Me

End Sub

Private Sub Command4_Click()

    If tipoprop = PAZ Then
        Call SendData("RECPPEAT" & lista.List(lista.ListIndex))
    Else
        Call SendData("RECPALIA" & lista.List(lista.ListIndex))

    End If

    Me.Hide
    Unload Me

End Sub

