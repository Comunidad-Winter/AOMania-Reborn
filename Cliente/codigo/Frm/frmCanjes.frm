VERSION 5.00
Begin VB.Form frmCanjes 
   BorderStyle     =   0  'None
   Caption         =   "Venta de Canjes"
   ClientHeight    =   6075
   ClientLeft      =   -60
   ClientTop       =   -60
   ClientWidth     =   6210
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   414
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Cantidad 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Text            =   "1"
      Top             =   5370
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   480
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Index           =   1
      Left            =   3720
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Image SellItem 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   3600
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      Height          =   195
      Left            =   1800
      TabIndex        =   10
      Top             =   5400
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CantidadCanjes"
      Height          =   195
      Index           =   5
      Left            =   2280
      TabIndex        =   8
      Top             =   5760
      Width           =   1110
   End
   Begin VB.Image BuyItem 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   360
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stats1"
      Height          =   195
      Index           =   4
      Left            =   4320
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stats1"
      Height          =   195
      Index           =   3
      Left            =   4320
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      Height          =   195
      Index           =   2
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monedas/Precio"
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Exit 
      Height          =   735
      Left            =   4800
      Top             =   5280
      Width           =   1335
   End
End
Attribute VB_Name = "frmCanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BuyItem_Click()
    Dim Index As Integer
    Dim UseItem As Integer
    Dim Monedas As Long
    
    Index = 0
    
    Call Audio.PlayWave(SND_CLICK)
    
    If List1(Index).List(List1(Index).ListIndex) = "Nada" Or List1(Index).ListIndex < 0 Then Exit Sub
    
    UseItem = List1(Index).ListIndex + 1
    Monedas = CANJInventory(UseItem).Monedas * Cantidad.Text
    
    If Monedas = 0 Then Exit Sub
    
    If UserCanjes >= CANJInventory(UseItem).Monedas Then
        SendData ("COAJ" & "," & UseItem & "," & Cantidad.Text)
    Else
        AddtoRichTextBox frmMain.RecTxt, "No tienes suficiente AoMCanjes.", 0, 0, 174, 1, 1
        Exit Sub
    End If

End Sub

Private Sub Exit_Click()
    SendData ("FINCOC")
End Sub

Private Sub Form_Load()
    Label1(5).Caption = "Tienes " & UserCanjes & " AoMCanjes."
End Sub

Private Sub List1_Click(Index As Integer)
    Dim SR As RECT, DR As RECT
    Dim UseItem As Integer

    SR.Left = 0
    SR.Top = 0
    SR.Right = 32
    SR.Bottom = 32

    DR.Left = 0
    DR.Top = 0
    DR.Right = 32
    DR.Bottom = 32
    
    If Index = 0 Then
       
        UseItem = List1(Index).ListIndex + 1
       
        If CANJInventory(UseItem).Name = "Nada" Then
            Label1(0).Visible = False
            Label1(1).Visible = False
            Label1(2).Visible = False
            Picture1.Refresh
        Else
            Label1(0).Visible = True
            Label1(1).Visible = True
            Label1(2).Visible = True
        End If
      
        Label1(0).Caption = CANJInventory(UseItem).Name
        Label1(1).Caption = CANJInventory(UseItem).Monedas
        Label1(2).Caption = CANJInventory(UseItem).Cantidad
       
        Select Case CANJInventory(UseItem).ObjType
            Case 2
                Label1(3).Caption = "Max Golpe:" & CANJInventory(UseItem).MaxHit
                Label1(4).Caption = "Min Golpe:" & CANJInventory(UseItem).MinHit
                Label1(3).Visible = True
                Label1(4).Visible = True

            Case 3
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & CANJInventory(UseItem).Def
                Label1(4).Visible = True

            Case 16
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & CANJInventory(UseItem).Def
                Label1(4).Visible = True

            Case 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & CANJInventory(UseItem).Def
                Label1(4).Visible = True
                
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
       
        Call DrawGrhtoHdc(Picture1.hdc, CANJInventory(UseItem).GrhIndex, DR)
    
    ElseIf Index = 1 Then
        
        UseItem = List1(Index).ListIndex + 1
        
        If Inventario.ItemName(UseItem) = "(None)" Then
            Label1(0).Visible = False
            Label1(1).Visible = False
            Label1(2).Visible = False
            Picture1.Refresh
        Else
            Label1(0).Visible = True
            Label1(1).Visible = True
            Label1(2).Visible = True
        End If
      
        Label1(0).Caption = Inventario.ItemName(UseItem)
        Label1(1).Caption = Inventario.Valor(UseItem)
        Label1(2).Caption = Inventario.Amount(UseItem)
       
        Select Case Inventario.ObjType(UseItem)

            Case 2
                Label1(3).Caption = "Max Golpe:" & Inventario.MaxHit(UseItem)
                Label1(4).Caption = "Min Golpe:" & Inventario.MinHit(UseItem)
                Label1(3).Visible = True
                Label1(4).Visible = True

            Case 3
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & Inventario.MaxDef(UseItem)
                Label1(4).Visible = True

            Case 16
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & Inventario.MaxDef(UseItem)
                Label1(4).Visible = True

            Case 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & Inventario.MaxDef(UseItem)
                Label1(4).Visible = True
                    
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False

        End Select

        
        Call DrawGrhtoHdc(Picture1.hdc, Inventario.GrhIndex(UseItem), DR)
    End If
    
End Sub

Private Sub SellItem_Click()
    Dim Index As Integer
    Dim UseItem As Integer
    Dim Monedas As Long
    
    Index = 1
    
    Call Audio.PlayWave(SND_CLICK)
    
    If List1(Index).List(List1(Index).ListIndex) = "Nada" Or List1(Index).ListIndex < 0 Then Exit Sub
    
    UseItem = List1(Index).ListIndex + 1
    
    Call SendData("VEAJ" & "," & UseItem & "," & Cantidad.Text)
End Sub
