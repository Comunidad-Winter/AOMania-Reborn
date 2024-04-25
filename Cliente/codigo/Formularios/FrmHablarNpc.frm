VERSION 5.00
Begin VB.Form FrmHablarNpc 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5820
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
   ScaleHeight     =   267
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   388
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFinalizar 
      Caption         =   "Finalizar"
      Height          =   360
      Left            =   2355
      TabIndex        =   3
      Top             =   3300
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "<"
      Height          =   360
      Left            =   225
      TabIndex        =   2
      Top             =   3300
      Width           =   990
   End
   Begin VB.CommandButton cmdSiguiente 
      Caption         =   ">"
      Height          =   360
      Left            =   4605
      TabIndex        =   1
      Top             =   3300
      Width           =   990
   End
   Begin VB.TextBox Text1 
      Height          =   2940
      Left            =   75
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FrmHablarNpc.frx":0000
      Top             =   75
      Width           =   5640
   End
End
Attribute VB_Name = "FrmHablarNpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAnterior_Click()
          
      If HablarQuest.Proceso = 1 Then
          Exit Sub
      End If
       
      HablarQuest.Proceso = HablarQuest.Proceso - 1
      
      Text1.Text = HablarQuest.Mensaje(HablarQuest.Proceso)
      
End Sub

Private Sub cmdFinalizar_Click()
    Call SendData("HBNF")
    Unload Me
End Sub

Private Sub cmdSiguiente_Click()
      
      If HablarQuest.NumMsj = 1 Then
          Exit Sub
      End If
      
      If HablarQuest.NumMsj = HablarQuest.Proceso Then
          Exit Sub
      End If
      
      HablarQuest.Proceso = HablarQuest.Proceso + 1
      
      If HablarQuest.NumMsj = HablarQuest.Proceso Then
          Text1.Text = HablarQuest.Mensaje(HablarQuest.Proceso)
          cmdFinalizar.Visible = True
          Exit Sub
      End If
      
      Text1.Text = HablarQuest.Mensaje(HablarQuest.Proceso)
      
End Sub

Private Sub Form_Load()
    
    Dim Proceso As Byte
    
    HablarQuest.Proceso = 1
    
    Proceso = HablarQuest.Proceso
    
    Text1.Text = HablarQuest.Mensaje(Proceso)
    
    If HablarQuest.NumMsj = Proceso Then
        cmdFinalizar.Visible = True
    End If
    
End Sub


