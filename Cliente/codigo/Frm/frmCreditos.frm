VERSION 5.00
Begin VB.Form frmCreditos 
   BorderStyle     =   0  'None
   Caption         =   "Comercio AoMCreditos"
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   421
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   398
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   765
      ScaleHeight     =   36.571
      ScaleMode       =   0  'User
      ScaleWidth      =   36.571
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   705
      Width           =   480
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Index           =   1
      Left            =   3120
      TabIndex        =   1
      Top             =   1560
      Width           =   2550
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Index           =   0
      Left            =   285
      TabIndex        =   0
      Top             =   1560
      Width           =   2550
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad de AoMCreditos"
      Height          =   195
      Index           =   5
      Left            =   1920
      TabIndex        =   8
      Top             =   6000
      Width           =   1800
   End
   Begin VB.Image Exit 
      Height          =   255
      Left            =   4800
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image BuyItem 
      Height          =   375
      Left            =   240
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stats2"
      Height          =   195
      Index           =   4
      Left            =   3960
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stats1"
      Height          =   195
      Index           =   3
      Left            =   3960
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      Height          =   195
      Index           =   2
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monedas o Precio"
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   555
   End
End
Attribute VB_Name = "frmCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BuyItem_Click()
    Dim Index As Integer
    Dim UseItem As Integer
    
    Index = 0
    
    Call Audio.PlayWave(SND_CLICK)
    
    If List1(Index).List(List1(Index).ListIndex) = "Nada" Or List1(Index).ListIndex < 0 Then Exit Sub
    
    UseItem = List1(Index).ListIndex + 1
    
    If UserCreditos >= CREDInventory(UseItem).Monedas Then
        SendData ("COAC" & "," & CREDInventory(UseItem).ObjIndex & "," & CREDInventory(UseItem).Monedas)
    Else
        AddtoRichTextBox frmMain.RecTxt, "No tienes suficiente AoMCreditos.", 0, 0, 174, 1, 1
        Exit Sub
    End If
End Sub

Private Sub Exit_Click()
    SendData ("FINCOA")
End Sub

Private Sub Form_Load()
    Label1(5).Caption = "Tienes " & UserCreditos & " AoMCreditos."
End Sub

Private Sub List1_Click(Index As Integer)

    Dim SR As RECT, dr As RECT
    Dim UseItem As Integer

    SR.Left = 0
    SR.Top = 0
    SR.Right = 32
    SR.bottom = 32

    dr.Left = 0
    dr.Top = 0
    dr.Right = 32
    dr.bottom = 32
    
    If Index = 0 Then
      
        UseItem = List1(Index).ListIndex + 1
      
        If CREDInventory(UseItem).Name = "Nada" Then
            Label1(0).Visible = False
            Label1(1).Visible = False
            Label1(2).Visible = False
            Picture1.Refresh
        Else
            Label1(0).Visible = True
            Label1(1).Visible = True
            Label1(2).Visible = True
        End If
      
        Label1(0).Caption = CREDInventory(UseItem).Name
        Label1(1).Caption = CREDInventory(UseItem).Monedas
        Label1(2).Caption = "Ilimitado"
       
        Select Case CREDInventory(UseItem).ObjType
            Case 2
                Label1(3).Caption = "Max Golpe:" & CREDInventory(UseItem).MaxHit
                Label1(4).Caption = "Min Golpe:" & CREDInventory(UseItem).MinHit
                Label1(3).Visible = True
                Label1(4).Visible = True

            Case 3
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & CREDInventory(UseItem).Def
                Label1(4).Visible = True

            Case 16
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & CREDInventory(UseItem).Def
                Label1(4).Visible = True

            Case 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & CREDInventory(UseItem).Def
                Label1(4).Visible = True
                
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
       
       
      
        Call DrawGrhtoHdc(Picture1.hdc, CREDInventory(UseItem).GrhIndex, dr)
 
    
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

      
        Call DrawGrhtoHdc(Picture1.hdc, Inventario.GrhIndex(UseItem), dr)
    End If

End Sub
