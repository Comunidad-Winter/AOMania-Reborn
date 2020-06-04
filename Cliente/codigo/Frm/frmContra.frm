VERSION 5.00
Begin VB.Form frmContra 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambiar Contaseña"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cambiar Contraseña"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   $"frmContra.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "Estado:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Escriba aqui su nueva contraseña:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   4335
   End
End
Attribute VB_Name = "frmContra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    If Text1.Text = "" Then
        Label1.Caption = "Estado: Debe escribir una contraseña!"
        Exit Sub

    End If

    Call SendData("/PASSWD " & Text1.Text)
    Label1.Caption = "Estado: Contraseña cambiada correctamente."
    Text1.Text = ""
    Unload frmContra

End Sub

