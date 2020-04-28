VERSION 5.00
Begin VB.Form frmQuest 
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8700
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
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   480
      Left            =   3825
      TabIndex        =   5
      Top             =   5520
      Width           =   4590
   End
   Begin VB.CommandButton cmdEntregar 
      Caption         =   "Entregar misión"
      Height          =   480
      Left            =   3825
      TabIndex        =   4
      Top             =   4890
      Width           =   4590
   End
   Begin VB.CommandButton cmdIniciar 
      Caption         =   "Iniciar"
      Height          =   480
      Left            =   3825
      TabIndex        =   3
      Top             =   4290
      Width           =   4590
   End
   Begin VB.TextBox InfoQuest 
      Height          =   2985
      Left            =   3705
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   4740
   End
   Begin VB.ListBox ListQuest 
      Height          =   5910
      Left            =   135
      TabIndex        =   1
      Top             =   120
      Width           =   3390
   End
   Begin VB.PictureBox PicQuest 
      BackColor       =   &H00000000&
      Height          =   960
      Left            =   3705
      ScaleHeight     =   900
      ScaleWidth      =   900
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   960
   End
End
Attribute VB_Name = "frmQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdIniciar_Click()
     
     Dim Index As Integer
     
     Index = ListQuest.ListIndex + 1
     
     Call SendData("INIQUEST" & Index)
     
End Sub

Private Sub Form_Load()
     
     Dim LooPC As Integer
     
    ListQuest.Clear
     
     For LooPC = 1 To NumQuests
            ListQuest.AddItem QuestList(LooPC).Nombre
     Next LooPC
     
End Sub

Private Sub ListQuest_Click()
      
      Dim Index As Integer
      Dim Datos As String
      Dim LooPC As Integer
      
      Index = ListQuest.ListIndex + 1
      
      Datos = "NOMBRE: " & vbCrLf _
                     & QuestList(Index).Nombre & vbCrLf
                     
      Datos = Datos & vbCrLf _
                     & "DESCRIPCION: " & vbCrLf _
                     & QuestList(Index).Descripcion & vbCrLf
                     
      Datos = Datos & vbCrLf _
                     & "RECOMPENSA: "
                     
      If QuestList(Index).RecompensaOro > 0 Then
           Datos = Datos & vbCrLf _
                     & "Oro: " & QuestList(Index).RecompensaOro
      End If
      
      If QuestList(Index).RecompensaExp > 0 Then
           Datos = Datos & vbCrLf _
                    & "Experencia: " & QuestList(Index).RecompensaExp
      End If
      
      If QuestList(Index).RecompensaItem > 0 Then
          For LooPC = 1 To QuestList(Index).RecompensaItem
                 
          Datos = Datos & vbCrLf _
                   & "Objeto: " & QuestList(Index).RecompensaObjeto(LooPC).ObjIndex & " x" & QuestList(Index).RecompensaObjeto(LooPC).Amount
                 
          Next LooPC
      End If
      
      InfoQuest.Text = Datos
      
      If Quest.InfoUser.UserQuest(Index) = 0 Then
           Set PicQuest.Picture = Interfaces.FrmQuest_SinHacer
      ElseIf Quest.InfoUser.UserQuest(Index) = 1 Then
           Set PicQuest.Picture = Interfaces.FrmQuest_Terminado
      End If
      
End Sub
