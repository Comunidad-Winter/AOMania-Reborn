VERSION 5.00
Begin VB.Form frmComerciar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox cantidad 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3105
      TabIndex        =   9
      Text            =   "1"
      Top             =   6690
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   1050
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   960
      Width           =   480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6195
      MouseIcon       =   "frmComerciar.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6780
      Width           =   465
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
      Height          =   3765
      Index           =   1
      Left            =   3930
      TabIndex        =   1
      Top             =   1920
      Width           =   2310
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
      Height          =   3765
      Index           =   0
      Left            =   690
      TabIndex        =   0
      Top             =   1920
      Width           =   2310
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2265
      TabIndex        =   10
      Top             =   6750
      Width           =   645
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   1
      Left            =   3855
      MouseIcon       =   "frmComerciar.frx":0152
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6165
      Width           =   2460
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   0
      Left            =   615
      MouseIcon       =   "frmComerciar.frx":02A4
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6150
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   3990
      TabIndex        =   8
      Top             =   975
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   3990
      TabIndex        =   7
      Top             =   630
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   3585
      TabIndex        =   6
      Top             =   1380
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   3120
      TabIndex        =   5
      Top             =   960
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   1125
      TabIndex        =   4
      Top             =   450
      Width           =   45
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^ y
'   le puse borde a la ventana.
'
'[END]'

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
Public LastIndex1 As Integer
Public LastIndex2 As Integer

Private Sub cantidad_Change()

    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1

    End If
    
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = 1

    End If

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)

    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0

        End If

    End If

End Sub

Private Sub Command2_Click()

    SendData ("FINCOM")

End Sub

Private Sub Form_Deactivate()

    'Me.SetFocus
End Sub

Private Sub Form_Load()

    'Valores máximos y mínimos para el ScrollBar
   
    'Cargamos la interfase
    Set Me.Picture = Interfaces.FrmComerciar_Principal
    Set Image1(0).Picture = Interfaces.FrmComerciar_Comprar
    Set Image1(1).Picture = Interfaces.FrmComerciar_Vender

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Image1(0).Tag = 0 Then
        Set Image1(0).Picture = Interfaces.FrmComerciar_Comprar
        Image1(0).Tag = 1

    End If

    If Image1(1).Tag = 0 Then
        Set Image1(1).Picture = Interfaces.FrmComerciar_Vender
        Image1(1).Tag = 1

    End If

End Sub

Private Sub Image1_Click(Index As Integer)

    Call Audio.PlayWave(SND_CLICK)

    If List1(Index).List(List1(Index).ListIndex) = "Nada" Or List1(Index).ListIndex < 0 Then Exit Sub

    Select Case Index

        Case 0
            frmComerciar.List1(0).SetFocus
            LastIndex1 = List1(0).ListIndex

            If UserGLD >= NPCInventory(List1(0).ListIndex + 1).Valor * Val(cantidad) Then
                SendData ("COMP" & "," & List1(0).ListIndex + 1 & "," & cantidad.Text)
            Else
                AddtoRichTextBox frmMain.RecTxt, "No tienes suficiente oro.", 0, 0, 174, 1, 1
                Exit Sub

            End If

        Case 1
            LastIndex2 = List1(1).ListIndex

            If Not Inventario.Equipped(List1(1).ListIndex + 1) Then
                SendData ("VEND" & "," & List1(1).ListIndex + 1 & "," & cantidad.Text)
            Else
                AddtoRichTextBox frmMain.RecTxt, "No podes vender el item porque lo estas usando.", 2, 51, 223, 1, 1
                Exit Sub

            End If
                
    End Select

    List1(0).Clear
    List1(1).Clear

    NPCInvDim = 0

End Sub

Private Sub List1_Click(Index As Integer)

    Dim SR As RECT, dr As RECT

    SR.Left = 0
    SR.Top = 0
    SR.Right = 32
    SR.bottom = 32

    dr.Left = 0
    dr.Top = 0
    dr.Right = 32
    dr.bottom = 32

    Select Case Index

        Case 0
            Label1(0).Caption = NPCInventory(List1(0).ListIndex + 1).Name
            Label1(1).Caption = NPCInventory(List1(0).ListIndex + 1).Valor
            Label1(2).Caption = NPCInventory(List1(0).ListIndex + 1).Amount

            Select Case NPCInventory(List1(0).ListIndex + 1).ObjType

                Case 2
                    Label1(3).Caption = "Max Golpe:" & NPCInventory(List1(0).ListIndex + 1).MaxHit
                    Label1(4).Caption = "Min Golpe:" & NPCInventory(List1(0).ListIndex + 1).MinHit
                    Label1(3).Visible = True
                    Label1(4).Visible = True

                Case 3
                    Label1(3).Visible = False
                    Label1(4).Caption = "Defensa:" & NPCInventory(List1(0).ListIndex + 1).Def
                    Label1(4).Visible = True

                Case 16
                    Label1(3).Visible = False
                    Label1(4).Caption = "Defensa:" & NPCInventory(List1(0).ListIndex + 1).Def
                    Label1(4).Visible = True

                Case 17
                    Label1(3).Visible = False
                    Label1(4).Caption = "Defensa:" & NPCInventory(List1(0).ListIndex + 1).Def
                    Label1(4).Visible = True

            End Select

            Call DrawGrhtoHdc(Picture1.hdc, NPCInventory(List1(0).ListIndex + 1).GrhIndex, dr)

        Case 1
            Label1(0).Caption = Inventario.ItemName(List1(1).ListIndex + 1)
            Label1(1).Caption = Inventario.Valor(List1(1).ListIndex + 1)
            Label1(2).Caption = Inventario.Amount(List1(1).ListIndex + 1)

            Select Case Inventario.ObjType(List1(1).ListIndex + 1)

                Case 2
                    Label1(3).Caption = "Max Golpe:" & Inventario.MaxHit(List1(1).ListIndex + 1)
                    Label1(4).Caption = "Min Golpe:" & Inventario.MinHit(List1(1).ListIndex + 1)
                    Label1(3).Visible = True
                    Label1(4).Visible = True

                Case 3
                    Label1(3).Visible = False
                    Label1(4).Caption = "Defensa:" & Inventario.MaxDef(List1(1).ListIndex + 1)
                    Label1(4).Visible = True

                Case 16
                    Label1(3).Visible = False
                    Label1(4).Caption = "Defensa:" & Inventario.MaxDef(List1(1).ListIndex + 1)
                    Label1(4).Visible = True

                Case 17
                    Label1(3).Visible = False
                    Label1(4).Caption = "Defensa:" & Inventario.MaxDef(List1(1).ListIndex + 1)
                    Label1(4).Visible = True

            End Select

            Call DrawGrhtoHdc(Picture1.hdc, Inventario.GrhIndex(List1(1).ListIndex + 1), dr)

    End Select

    Picture1.Refresh

End Sub

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Image1(0).Tag = 0 Then
        Set Image1(0).Picture = Interfaces.FrmComerciar_Comprar
        Image1(0).Tag = 1

    End If

    If Image1(1).Tag = 0 Then
        Set Image1(1).Picture = Interfaces.FrmComerciar_Vender
        Image1(1).Tag = 1

    End If

End Sub

