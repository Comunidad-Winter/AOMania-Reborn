VERSION 5.00
Object = "{B389CD47-E20E-4D96-A4EC-576F2B1F43BF}#1.0#0"; "hook-menu-2.ocx"
Begin VB.Form frmPanelGm 
   BackColor       =   &H80000007&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Panel GM"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   5520
      Top             =   420
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   4
      Bmp:1           =   "frmPanelGm.frx":0000
      Key:1           =   "#mnuborramensajeshow"
      Bmp:2           =   "frmPanelGm.frx":0808
      Key:2           =   "#mnuborrarmensajequest"
      Bmp:3           =   "frmPanelGm.frx":1010
      Key:3           =   "#mnuborrarmensajeconsultas"
      Bmp:4           =   "frmPanelGm.frx":1818
      Key:4           =   "#mnuiralusuarioshow"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Ver Procesos"
      Height          =   315
      Index           =   22
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Cliente"
      Height          =   315
      Index           =   21
      Left            =   1320
      TabIndex        =   22
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Blok cliente"
      Height          =   375
      Index           =   20
      Left            =   2280
      TabIndex        =   21
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Ver oro en banco"
      Height          =   315
      Index           =   19
      Left            =   120
      TabIndex        =   20
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Show SOS"
      Height          =   315
      Index           =   18
      Left            =   1200
      TabIndex        =   19
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Boveda"
      Height          =   315
      Index           =   17
      Left            =   2280
      TabIndex        =   18
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Ban X ip"
      Height          =   315
      Index           =   16
      Left            =   3360
      TabIndex        =   17
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Penas"
      Height          =   315
      Index           =   15
      Left            =   3360
      TabIndex        =   16
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "IP 2 NICK"
      Height          =   315
      Index           =   14
      Left            =   3360
      TabIndex        =   15
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "NICK 2 IP"
      Height          =   315
      Index           =   13
      Left            =   2280
      TabIndex        =   14
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "UNBAN"
      Height          =   315
      Index           =   12
      Left            =   1200
      TabIndex        =   13
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "CARCEL"
      Height          =   315
      Index           =   11
      Left            =   3360
      TabIndex        =   12
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "SKILLS"
      Height          =   315
      Index           =   10
      Left            =   1200
      TabIndex        =   11
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "INV"
      Height          =   315
      Index           =   9
      Left            =   1200
      TabIndex        =   10
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "INFO"
      Height          =   315
      Index           =   8
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "DONDE"
      Height          =   315
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "HORA"
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "IRA"
      Height          =   315
      Index           =   3
      Left            =   1200
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "SUM"
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "BAN"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "ECHAR"
      Height          =   315
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "Actualiza"
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   2950
      Width           =   315
   End
   Begin VB.ComboBox cboListaUsus 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3435
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "< -OJO con ESTE"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4440
      TabIndex        =   28
      Top             =   1750
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "<-- Acciones Auxiliares"
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   2280
      TabIndex        =   27
      Top             =   2570
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "<-- Verificaciones de CHITS"
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   2280
      TabIndex        =   26
      Top             =   3070
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "<-- Acciones hacia el user"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "<-- Cerrar"
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   5040
      TabIndex        =   24
      Top             =   3070
      Width           =   1335
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   120
      X2              =   120
      Y1              =   540
      Y2              =   1380
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   4440
      X2              =   4440
      Y1              =   540
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2280
      X2              =   2280
      Y1              =   960
      Y2              =   1380
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2280
      X2              =   4440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   4440
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   2280
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Menu Menu_Show 
      Caption         =   "Show"
      Begin VB.Menu mnuborramensajeshow 
         Caption         =   "Borrar mensaje"
      End
      Begin VB.Menu mnuiralusuarioshow 
         Caption         =   "Ir al usuario"
      End
      Begin VB.Menu mnutraeralusuarioshow 
         Caption         =   "Traer al usuario"
      End
      Begin VB.Menu mnuinvalidashow 
         Caption         =   "Inválida"
      End
      Begin VB.Menu mnumanualshow 
         Caption         =   "Manual/FAQ"
      End
      Begin VB.Menu ViewMsgSOS 
         Caption         =   "Ver mensaje"
      End
   End
   Begin VB.Menu Menu_Quest 
      Caption         =   "Quests"
      Begin VB.Menu mnuborrarmensajequest 
         Caption         =   "Borrar mensaje"
      End
      Begin VB.Menu mnutraerusuarioquest 
         Caption         =   "Traer al usuario"
      End
      Begin VB.Menu mnullevaranixquest 
         Caption         =   "Llevar a nix"
      End
   End
   Begin VB.Menu Menu_Consultas 
      Caption         =   "Consultas"
      Begin VB.Menu mnuborrarmensajeconsultas 
         Caption         =   "Borrar mensaje"
      End
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccion_Click(Index As Integer)

    Dim Ok   As Boolean, tmp As String, tmp2 As String
    Dim Nick As String

    Nick = cboListaUsus.Text

    Select Case Index

        Case 0 '/ECHAR nick
            Call SendData("/ECHHH " & Nick)

        Case 1 '/ban motivo@nick
            tmp = InputBox("Motivo ?", "")

            If MsgBox("Esta seguro que desea banear al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
                Call SendData("/BEEEH " & tmp & "@" & Nick)

            End If

        Case 2 '/sum nick
            Call SendData("/SUM " & Nick)

        Case 3 '/ira nick
            Call SendData("/IRA " & Nick)

        Case 4 '/rem
            tmp = InputBox("Comentario ?", "")
            Call SendData("/REM " & tmp)

        Case 5 '/hora
            Call SendData("/HORA")

        Case 6 '/donde nick
            Call SendData("/DONDE " & Nick)

        Case 7 '/nene
            tmp = InputBox("Mapa ?", "")
            Call SendData("/NENE " & Trim(tmp))

        Case 8 '/info nick
            Call SendData("/INFO " & Nick)

        Case 9 '/inv nick
            Call SendData("/INV " & cboListaUsus.Text)

        Case 10 '/skills nick
            Call SendData("/SKILLS " & Nick)

        Case 11 '/carcel minutos nick
            tmp = InputBox("Minutos ? (hasta 30)", "")
            tmp2 = InputBox("Razon ?", "")

            If MsgBox("Esta seguro que desea encarcelar al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
                Call SendData("/CARCEL " & Nick & "@" & tmp2 & "@" & tmp)

            End If

        Case 12 '/unban nick

            If MsgBox("Esta seguro que desea removerle el ban al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
                Call SendData("/UNNEEEH " & Nick)

            End If

        Case 13 '/nick2ip nick
            Call SendData("/NICK2IP " & Nick)

        Case 14 '/ip2nick ip
            Call SendData("/IP2NICK " & Nick)

        Case 15 '/penas
            Call SendData("/PENAS " & cboListaUsus.Text)

        Case 16 'Ban X ip

            If MsgBox("Esta seguro que desea banear el (ip o personaje) " & Nick & "Por IP?", vbYesNo) = vbYes Then
                Nick = Replace(Nick, " ", "+")
                Call SendData("/BANLAIP " & Nick)

            End If

        Case 17 ' MUESTA BOBEDA
            Call SendData("/BOV " & Nick)

        Case 18 ' Sos
            Call SendData("/SHOW SOS")

        Case 19 ' Balance
            Call SendData("/BAL " & cboListaUsus.Text)

        Case 20 ' Blokeo de cliente
            Call SendData("/BLOKK " & Nick)

        Case 21 ' Verificacion del cliente
            Call SendData("/SO33 " & Nick)

        Case 22 'Ver procesos del usuario
            Call SendData("/VERPROCESOS " & Nick)
    
    End Select

End Sub

Private Sub cmdActualiza_Click()

    Call SendData("LISTUSU")

End Sub

Private Sub cmdCerrar_Click()

    Unload Me

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

    'Valores máximos y mínimos para el ScrollBar
    
    Call cmdActualiza_Click

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Unload Me

End Sub

'Añadir al nuevo codigo miqueas

Sub mnuborramensajeshow_Click()

    Call SendData("/DROPSOS " & frmPaneldeGM.ListShow.Text)

    Dim X      As Integer
    Dim Count  As Integer
    Dim result As Integer

    Count = frmPaneldeGM.ListShow.ListCount + "1"
    result = 0
     
    For X = 1 To Count
     
        If (frmPaneldeGM.ListShow.List(result) = frmPaneldeGM.ListShow.Text) Then
             
            If frmPaneldeGM.ListShow.Text = "" Then
                Exit Sub
            Else
             
                frmPaneldeGM.ListShow.RemoveItem result
             
            End If
             
            Exit Sub

        End If

        result = result + "1"
    Next X
        
End Sub

Sub ViewMsgSOS_Click()
      
    Dim msg As String
     
    msg = frmPaneldeGM.ListShow.Text
     
    Dim MsgArray() As String
     
    MsgArray() = Split(msg, Chr(64))
     
    Dim Name    As String
    Dim Section As String
    Dim Message As String
     
    If msg = "" Then
        Exit Sub

    End If
     
    Name = MsgArray(0)
    Section = MsgArray(1)
    Message = MsgArray(2)
     
    Call MsgBox(Section & " de " & Name & " : " & Message, vbInformation)
      
End Sub

Sub mnuborrarmensajeconsultas_Click()

    Call SendData("/DROPGM " & frmPaneldeGM.ListConsultas.Text)

    Dim X      As Integer
    Dim Count  As Integer
    Dim result As Integer

    Count = frmPaneldeGM.ListConsultas.ListCount + "1"
    result = 0
     
    For X = 1 To Count
     
        If (frmPaneldeGM.ListConsultas.List(result) = frmPaneldeGM.ListConsultas.Text) Then
            
            If frmPaneldeGM.ListConsultas.Text = "" Then
                Exit Sub
            Else
                frmPaneldeGM.ListConsultas.RemoveItem result

            End If
             
            Exit Sub

        End If

        result = result + "1"
    Next X
        
End Sub

Sub mnuborrarmensajequest_Click()

    Call SendData("/DROPQUEST " & frmPaneldeGM.ListQuest.Text)

    Dim X      As Integer
    Dim Count  As Integer
    Dim result As Integer

    Count = frmPaneldeGM.ListQuest.ListCount + "1"
    result = 0
     
    For X = 1 To Count
     
        If (frmPaneldeGM.ListQuest.List(result) = frmPaneldeGM.ListQuest.Text) Then
             
            If frmPaneldeGM.ListQuest.Text = "" Then
                Exit Sub
            Else
                frmPaneldeGM.ListQuest.RemoveItem result

            End If
             
            Exit Sub

        End If

        result = result + "1"
    Next X
        
End Sub

Sub mnuiralusuarioshow_Click()
    
    Dim msg As String
     
    msg = frmPaneldeGM.ListShow.Text
     
    Dim MsgArray() As String
     
    MsgArray() = Split(msg, Chr(64))
     
    Dim Name    As String
    Dim Section As String
    Dim Message As String
     
    If msg = "" Then
        Exit Sub

    End If
     
    Name = MsgArray(0)
    Section = MsgArray(1)
    Message = MsgArray(2)
    
    Call SendData("/Ira " & Name)
    
End Sub

Sub mnutraeralusuarioshow_Click()
    
    Dim msg As String
     
    msg = frmPaneldeGM.ListShow.Text
     
    Dim MsgArray() As String
     
    MsgArray() = Split(msg, Chr(64))
     
    Dim Name    As String
    Dim Section As String
    Dim Message As String
     
    If msg = "" Then
        Exit Sub

    End If
     
    Name = MsgArray(0)
    Section = MsgArray(1)
    Message = MsgArray(2)
    
    Call SendData("/Sum " & Name)
    
End Sub

Sub mnutraerusuarioquest_Click()
    
    Dim msg As String
     
    msg = frmPaneldeGM.ListQuest.Text
         
    Dim Name As String
     
    If msg = "" Then
        Exit Sub

    End If
     
    Name = frmPaneldeGM.ListQuest.Text
    
    Call SendData("/Sum " & Name)
    
End Sub

Sub mnullevaranixquest_Click()
    
    Dim msg As String
     
    msg = frmPaneldeGM.ListQuest.Text
         
    Dim Name As String
     
    If msg = "" Then
        Exit Sub

    End If
     
    Name = frmPaneldeGM.ListQuest.Text
    
    Call SendData("/Telep " & Name & " 34" & " 40" & " 49")
    
End Sub
