VERSION 5.00
Begin VB.Form frmGuildDetails 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Detalles del Clan"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   6735
   ClipControls    =   0   'False
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
   MousePointer    =   99  'Custom
   ScaleHeight     =   6825
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Contraseña"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   6495
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " roban el personaje no puedan expulsar miembros de tu clan sin tu consentimiento."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   5955
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deberás introducir aquí una contraseña inventada. De esta forma evitarás que si te"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.CommandButton Command1 
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
      Index           =   1
      Left            =   5160
      MouseIcon       =   "frmGuildDetails.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
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
      Index           =   0
      Left            =   120
      MouseIcon       =   "frmGuildDetails.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Codex"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   6495
      Begin VB.TextBox txtCodex1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   360
         TabIndex        =   10
         Top             =   3720
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   360
         TabIndex        =   9
         Top             =   3360
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   360
         TabIndex        =   8
         Top             =   3000
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   7
         Top             =   1800
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   5655
      End
   End
   Begin VB.Frame frmDesc 
      BackColor       =   &H00000000&
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000009&
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   6255
      End
   End
End
Attribute VB_Name = "frmGuildDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click(Index As Integer)

    Select Case Index

        Case 0
            Unload Me

        Case 1
            Dim fdesc$
            fdesc$ = Replace(txtDesc, vbCrLf, "º", , , vbBinaryCompare)
    
            '    If Not AsciiValidos(fdesc$) Then
            '        MsgBox "La descripcion contiene caracteres invalidos"
            '        Exit Sub
            '    End If
    
            Dim k    As Integer
            Dim Cont As Integer
            Cont = 0

            For k = 0 To txtCodex1.UBound

                '        If Not AsciiValidos(txtCodex1(k)) Then
                '            MsgBox "El codex tiene invalidos"
                '            Exit Sub
                '        End If
                If Len(txtCodex1(k).Text) > 0 Then Cont = Cont + 1
            Next k

            If Cont < 4 Then
                MsgBox "Debes definir al menos cuatro mandamientos."
                Exit Sub
            End If
            
            If txtPassword.Text = "" Then
                MsgBox "Debes introducir una contraseña."
                Exit Sub
            End If
    
            Dim chunk$
    
            If CreandoClan Then
                chunk$ = "CIG" & fdesc$
                chunk$ = chunk$ & "¬" & ClanName & "¬" & Site & "¬" & Cont & "¬" & txtPassword.Text
            Else
                chunk$ = "DESCOD" & fdesc$ & "¬" & Cont

            End If
    
            For k = 0 To txtCodex1.UBound
                chunk$ = chunk$ & "¬" & txtCodex1(k)
            Next k
    
            Call SendData(chunk$)
    
            CreandoClan = False
    
            Unload Me
    
    End Select

End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Set Command1(0).MouseIcon = Iconos.Ico_Mano
    Set Command1(1).MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Deactivate()
     
    'If Not frmGuildLeader.Visible Then
    '    Me.SetFocus
    'Else
    '    'Unload Me
    'End If
    '
End Sub

Private Sub Form_Load()
    Set Me.MouseIcon = Iconos.Diablo
End Sub
