VERSION 5.00
Begin VB.Form frmAdmin 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Administración del servidor"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Personajes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   495
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   720
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Echar todos los PJS no privilegiados"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   2280
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "R"
         Height          =   315
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.ComboBox cboPjs 
         Height          =   315
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Echar"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   1800
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim TIndex As Long

    TIndex = NameIndex(cboPjs.Text)

    If TIndex > 0 Then
        Call SendData(SendTarget.toall, 0, 0, "||AOMania> " & UserList(TIndex).Name & " ha sido hechado. " & FONTTYPE_SERVER)
        Call CloseSocket(TIndex)

    End If

End Sub

Public Sub ActualizaListaPjs()

    Dim loopc As Long

    With cboPjs
        .Clear
    
        For loopc = 1 To LastUser

            If UserList(loopc).flags.UserLogged And UserList(loopc).ConnID >= 0 And UserList(loopc).ConnIDValida Then

                If UserList(loopc).flags.Privilegios = PlayerType.User Then
                    .AddItem UserList(loopc).Name
                    .ItemData(.NewIndex) = loopc

                End If

            End If

        Next loopc

    End With

End Sub

Private Sub Command3_Click()

    Call EcharPjsNoPrivilegiados

End Sub

Private Sub Label1_Click()

End Sub
