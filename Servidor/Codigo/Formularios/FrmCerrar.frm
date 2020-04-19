VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "COMCTL32.OCX"
Begin VB.Form FrmCerrar 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4590
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
   ScaleHeight     =   3135
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "FrmCerrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
      Dim i As Integer
      
      For i = 1 To LastNPC
          
          Select Case Npclist(i).Numero
          
                 Case NpcYetiOscura, NpcYeti, NpcCleopatra, NpcReyScorpion, NpcDarkSeth, NpcTiburonBlanco, _
                          NpcElfica, NpcGranDragonRojo, NpcNosfe, NpcThorn, NpcDragonAlado, NPC_CENTINELA_TIERRA
                          Call QuitarNPC(i)
                  
                  Case NpcBruja
                          If Npclist(i).pos.Map = MapaCasaAbandonada1 Then
                             Call QuitarNPC(i)
                          End If
                  
          End Select
          
      Next i
      
      DoEvents
      ProgressBar1.value = 20
      
      Call GuardarUsuarios
      DoEvents
      ProgressBar1.value = 40
      
      Call DoBackUp
      
      DoEvents
      ProgressBar1.value = 99
      
      Dim f
      
    For Each f In Forms

         Unload f
     Next
      
      
End Sub
