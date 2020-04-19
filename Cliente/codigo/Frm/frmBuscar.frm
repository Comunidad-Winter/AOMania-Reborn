VERSION 5.00
Begin VB.Form frmBuscar 
   BorderStyle     =   0  'None
   ClientHeight    =   6345
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6120
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
   ScaleHeight     =   6345
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin AoManiaClienteGM.ChameleonBtn Cerrar 
      Height          =   405
      Left            =   3975
      TabIndex        =   19
      Top             =   5535
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   65535
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBuscar.frx":0000
      PICN            =   "frmBuscar.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AoManiaClienteGM.ChameleonBtn Command4 
      Height          =   405
      Left            =   2355
      TabIndex        =   18
      Top             =   5535
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Salir Buscador"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   65535
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBuscar.frx":04B6
      PICN            =   "frmBuscar.frx":04D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AoManiaClienteGM.ChameleonBtn Baneos 
      Height          =   405
      Left            =   1230
      TabIndex        =   17
      Top             =   5535
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Baneos"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   65535
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBuscar.frx":096C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AoManiaClienteGM.ChameleonBtn Command3 
      Height          =   345
      Left            =   1965
      TabIndex        =   16
      Top             =   2130
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Buscar NPCs Hostiles"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   65535
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBuscar.frx":0988
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AoManiaClienteGM.ChameleonBtn Command2 
      Height          =   330
      Left            =   1980
      TabIndex        =   15
      Top             =   1425
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "Buscar NPCs"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   65535
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBuscar.frx":09A4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AoManiaClienteGM.ChameleonBtn Command1 
      Height          =   315
      Left            =   1950
      TabIndex        =   14
      Top             =   705
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Buscar Objeto"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   65535
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBuscar.frx":09C0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox NpchData 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00000080&
      Height          =   2595
      Left            =   1440
      TabIndex        =   13
      Top             =   2805
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.ListBox NpchId 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00000080&
      Height          =   2595
      Left            =   600
      TabIndex        =   12
      Top             =   2805
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox NpchText 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.ListBox NPCID 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00000080&
      Height          =   2595
      Left            =   600
      TabIndex        =   10
      Top             =   2805
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox NPCData 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00000080&
      Height          =   2595
      Left            =   1440
      TabIndex        =   9
      Top             =   2805
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox NPCText 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.ListBox OBJData 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   2595
      Left            =   1440
      TabIndex        =   7
      Top             =   2805
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.ListBox OBJIDs 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   2595
      Left            =   600
      TabIndex        =   6
      Top             =   2805
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox OBJText 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   600
      X2              =   5280
      Y1              =   5415
      Y2              =   5415
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   600
      X2              =   5280
      Y1              =   1365
      Y2              =   1365
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   600
      X2              =   5280
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   600
      X2              =   5280
      Y1              =   2085
      Y2              =   2085
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Crear"
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BUSCADOR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2160
      TabIndex        =   0
      Top             =   0
      Width           =   1620
   End
   Begin VB.Menu Menu_Item 
      Caption         =   "Crear Item"
      Visible         =   0   'False
      Begin VB.Menu mnuCI 
         Caption         =   "Crear objeto"
      End
   End
   Begin VB.Menu Menu_Npc 
      Caption         =   "Crear NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuCrearNpc 
         Caption         =   "Sacar Npc"
      End
   End
   Begin VB.Menu Menu_Npch 
      Caption         =   "Crear Npc"
      Visible         =   0   'False
      Begin VB.Menu mnuCrearNpch 
         Caption         =   "Sacar Npc"
      End
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tmp As Long

Private Sub Baneos_Click()
    Unload Me
    Unload frmPaneldeGM
    
    frmBaneos.Show vbModal, frmMain

End Sub

Private Sub Cerrar_Click()
    Unload frmPaneldeGM.fChild
    Unload frmPaneldeGM

End Sub

Private Sub Command1_Click()
     
    If NpchText.Visible = True Then
        NpchText.Visible = False

    End If
     
    If NpchId.Visible = True Then
        NpchId.Visible = False

    End If
     
    If NpchData.Visible = True Then
        NpchData.Visible = False

    End If
     
    If NPCText.Visible = True Then
        NPCText.Visible = False

    End If
     
    If NPCID.Visible = True Then
        NPCID.Visible = False

    End If
     
    If NPCData.Visible = True Then
        NPCData.Visible = False

    End If
     
    If Label2.Visible = False Then
        Label2.Visible = True

    End If
     
    If OBJText.Visible = False Then
        OBJText.Visible = True

    End If
     
    If OBJIDs.Visible = False Then
        OBJIDs.Visible = True

    End If
     
    If OBJData.Visible = False Then
        OBJData.Visible = True

    End If
     
    If Label2.Visible = True And OBJText.Visible = True And OBJIDs.Visible = True And OBJData.Visible = True Then
        OBJText.Text = " "
        OBJIDs.Clear
        OBJData.Clear
         
        Call SendData("/SEARCHITEMS " & Text1.Text)

    End If

End Sub

Private Sub Command2_Click()
      
    If NpchText.Visible = True Then
        NpchText.Visible = False

    End If
     
    If NpchId.Visible = True Then
        NpchId.Visible = False

    End If
     
    If NpchData.Visible = True Then
        NpchData.Visible = False

    End If
     
    If OBJText.Visible = True Then
        OBJText.Visible = False

    End If
    
    If OBJIDs.Visible = True Then
        OBJIDs.Visible = False

    End If
    
    If OBJData.Visible = True Then
        OBJIDs.Visible = False

    End If
    
    If Label2.Visible = False Then
        Label2.Visible = True

    End If
      
    If NPCText.Visible = False Then
        NPCText.Visible = True

    End If
      
    If NPCID.Visible = False Then
        NPCID.Visible = True

    End If
      
    If NPCData.Visible = False Then
        NPCData.Visible = True

    End If
      
    If NPCText.Visible = True And NPCID.Visible = True And NPCData.Visible = True Then
        NPCText.Text = " "
        NPCID.Clear
        NPCData.Clear
        Call SendData("/SEARCHNPCS " & Text2.Text)

    End If
      
End Sub

Private Sub Command3_Click()

    If NPCText.Visible = True Then
        NPCText.Visible = False

    End If
     
    If NPCID.Visible = True Then
        NPCID.Visible = False

    End If
     
    If NPCData.Visible = True Then
        NPCData.Visible = False

    End If
     
    If OBJText.Visible = True Then
        OBJText.Visible = False

    End If
    
    If OBJIDs.Visible = True Then
        OBJIDs.Visible = False

    End If
    
    If OBJData.Visible = True Then
        OBJIDs.Visible = False

    End If
    
    If Label2.Visible = False Then
        Label2.Visible = True

    End If
      
    If NpchText.Visible = False Then
        NpchText.Visible = True

    End If
      
    If NpchId.Visible = False Then
        NpchId.Visible = True

    End If
      
    If NpchData.Visible = False Then
        NpchData.Visible = True

    End If
      
    If NpchText.Visible = True And NpchId.Visible = True And NpchData.Visible = True Then
        NpchText.Text = " "
        NpchId.Clear
        NpchData.Clear
        Call SendData("/SEARCHNPCSH " & Text3.Text)

    End If

End Sub

Private Sub Command4_Click()
    'Unload frmPaneldeGM.fChild
    Unload frmPaneldeGM.fChild

End Sub

Private Sub mnuCI_Click()
    tmp = InputBox("Ingrese la Cantidad de Objetos Max 1200", "")
    Call SendData("/CI " & OBJIDs.Text & " " & tmp)

End Sub

Private Sub mnuCrearNpc_Click()
    Call SendData("/ACC " & NPCID.Text)

End Sub

Private Sub mnuCrearNpch_Click()
    Call SendData("/ACC " & NpchId.Text)

End Sub

Private Sub NpchID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NpchData.TopIndex = NpchId.TopIndex

End Sub

Private Sub NpchData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NpchId.TopIndex = NpchData.TopIndex

End Sub

Private Sub NPChID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
      
        'Mostramos el menú popup
        PopupMenu Menu_Npch
  
    End If

End Sub

Private Sub NPCID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
      
        'Mostramos el menú popup
        PopupMenu Menu_Npc
  
    End If

End Sub

Private Sub NpcID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NPCData.TopIndex = NPCID.TopIndex

End Sub

Private Sub NpcData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NPCID.TopIndex = NPCData.TopIndex

End Sub

Private Sub NpcData_Scroll()
    NPCID.TopIndex = NPCData.TopIndex

End Sub

Private Sub NpcId_Scroll()
    NPCData.TopIndex = NPCID.TopIndex

End Sub

Private Sub NpchData_Scroll()
    NpchId.TopIndex = NpchData.TopIndex

End Sub

Private Sub NpchId_Scroll()
    NpchData.TopIndex = NpchId.TopIndex

End Sub

Private Sub NPCID_Click()
    Static Estoy As Boolean
   
    If Not Estoy Then
       
        NPCData.ListIndex = NPCID.ListIndex
        NPCData.TopIndex = NPCID.TopIndex
       
    End If

End Sub

Private Sub NPCDATA_Click()
    Static Estoy As Boolean
   
    If Not Estoy Then
       
        NPCID.ListIndex = NPCData.ListIndex
        NPCID.TopIndex = NPCData.TopIndex
       
    End If

End Sub

Private Sub NPCHID_Click()
    Static Estoy As Boolean
   
    If Not Estoy Then
       
        NpchData.ListIndex = NpchId.ListIndex
        NpchData.TopIndex = NpchId.TopIndex
       
    End If

End Sub

Private Sub NPCHDATA_Click()
    Static Estoy As Boolean
   
    If Not Estoy Then
       
        NpchId.ListIndex = NpchData.ListIndex
        NpchId.TopIndex = NpchData.TopIndex
       
    End If

End Sub

Private Sub OBJData_Click()
    Static Estoy As Boolean
   
    If Not Estoy Then
       
        OBJIDs.ListIndex = OBJData.ListIndex
        OBJIDs.TopIndex = OBJData.TopIndex
       
    End If

End Sub

Private Sub OBJData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OBJIDs.TopIndex = OBJData.TopIndex

End Sub

Private Sub OBJData_Scroll()
    OBJIDs.TopIndex = OBJData.TopIndex

End Sub

Private Sub OBJIDs_Click()
    Static Estoy As Boolean
   
    If Not Estoy Then
       
        OBJData.ListIndex = OBJIDs.ListIndex
        OBJData.TopIndex = OBJIDs.TopIndex
       
    End If

End Sub

Private Sub OBJIDs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
      
        'Mostramos el menú popup
        PopupMenu Menu_Item
  
    End If

End Sub

Private Sub OBJIDs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OBJData.TopIndex = OBJIDs.TopIndex

End Sub

Private Sub OBJIDs_Scroll()
    OBJData.TopIndex = OBJIDs.TopIndex

End Sub

