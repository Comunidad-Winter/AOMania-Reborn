VERSION 5.00
Begin VB.Form frmComerciarUsu 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   483
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   6600
      Width           =   2235
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ofrecer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5475
      Left            =   3360
      TabIndex        =   6
      Top             =   600
      Width           =   3135
      Begin VB.OptionButton optQue 
         Caption         =   "Oro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   14
         Top             =   240
         Width           =   915
      End
      Begin VB.OptionButton optQue 
         Caption         =   "Objeto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.TextBox txtCant 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Text            =   "1"
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton cmdOfrecer 
         Caption         =   "Ofrecer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   8
         Top             =   4980
         Width           =   2490
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3960
         Left            =   180
         TabIndex        =   7
         Top             =   480
         Width           =   2490
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   4610
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Respuesta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5475
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3135
      Begin VB.CommandButton cmdRechazar 
         Caption         =   "Rechazar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   4980
         Width           =   1230
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   4
         Top             =   4980
         Width           =   1230
      End
      Begin VB.ListBox List2 
         Height          =   3960
         Left            =   180
         TabIndex        =   2
         Top             =   480
         Width           =   2490
      End
      Begin VB.Label Label3 
         Caption         =   "Cantidad: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   4620
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   540
      Left            =   1080
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   60
      Width           =   540
   End
   Begin VB.Label lblEstadoResp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Esperando respuesta..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1762
      TabIndex        =   5
      Top             =   180
      Visible         =   0   'False
      Width           =   2490
   End
End
Attribute VB_Name = "frmComerciarUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()

    Call SendData("COMUSUOK")

End Sub

Private Sub cmdOfrecer_Click()

    If optQue(0).Value = True Then
        If List1.ListIndex < 0 Then Exit Sub
        If List1.ItemData(List1.ListIndex) <= 0 Then Exit Sub
    
        '    If Val(txtCant.Text) > List1.ItemData(List1.ListIndex) Or _
        '        Val(txtCant.Text) <= 0 Then Exit Sub
    ElseIf optQue(1).Value = True Then

        '    If Val(txtCant.Text) > UserGLD Then
        '        Exit Sub
        '    End If
    End If

    If optQue(0).Value = True Then
        Call SendData("OFRECER" & List1.ListIndex + 1 & "," & Trim(Val(txtCant.Text)))
    ElseIf optQue(1).Value = True Then
        Call SendData("OFRECER" & FLAGORO & "," & Trim(Val(txtCant.Text)))
    Else
        Exit Sub

    End If

    lblEstadoResp.Visible = True

End Sub

Private Sub cmdRechazar_Click()

    Call SendData("COMUSUNO")

End Sub

Private Sub Command2_Click()

    Call SendData("FINCOMUSU")

End Sub

Private Sub Form_Deactivate()

    'Me.SetFocus
    'Picture1.SetFocus

End Sub

Private Sub Form_Load()

    'Valores m�ximos y m�nimos para el ScrollBar
   
    'Carga las imagenes...?
    lblEstadoResp.Visible = False

End Sub

Private Sub Form_LostFocus()

    Me.SetFocus
    Picture1.SetFocus

End Sub

Private Sub List1_Click()

    DibujaGrh Inventario.GrhIndex(List1.ListIndex + 1)

End Sub

Public Sub DibujaGrh(Grh As Long)

    Dim SR As RECT, DR As RECT

    SR.Left = 0
    SR.Top = 0
    SR.Right = 32
    SR.Bottom = 32

    DR.Left = 0
    DR.Top = 0
    DR.Right = 32
    DR.Bottom = 32

    Call DrawGrhtoHdc(Picture1.hdc, Grh, DR)

End Sub

Private Sub List2_Click()

    If List2.ListIndex >= 0 Then
        DibujaGrh OtroInventario(List2.ListIndex + 1).GrhIndex
        Label3.Caption = "Cantidad: " & List2.ItemData(List2.ListIndex)
        cmdAceptar.Enabled = True
        cmdRechazar.Enabled = True
    Else
        cmdAceptar.Enabled = False
        cmdRechazar.Enabled = False

    End If

End Sub

Private Sub optQue_Click(Index As Integer)

    Select Case Index

        Case 0
            List1.Enabled = True

        Case 1
            List1.Enabled = False

    End Select

End Sub

Private Sub txtCant_Change()

    If Val(txtCant.Text) < 1 Then txtCant.Text = "1"

End Sub

Private Sub txtCant_KeyDown(KeyCode As Integer, Shift As Integer)

    If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Or (KeyCode >= 37 And KeyCode <= 40)) Then
        'txtCant = KeyCode
        KeyCode = 0

    End If

End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)

    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
        'txtCant = KeyCode
        KeyAscii = 0

    End If

End Sub

'[/Alejo]

