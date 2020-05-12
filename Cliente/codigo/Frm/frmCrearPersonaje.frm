VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "AoMania"
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H00E0E0E0&
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCrearPersonaje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmCrearPersonaje.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPersonaje 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   390
      TabIndex        =   20
      Top             =   5355
      Width           =   3570
   End
   Begin VB.TextBox txtBanco 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   390
      TabIndex        =   19
      Top             =   4515
      Width           =   3570
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   390
      MaxLength       =   20
      MouseIcon       =   "frmCrearPersonaje.frx":1594
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1425
      Width           =   3570
   End
   Begin VB.TextBox txtPasswdCheck 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   390
      MaxLength       =   25
      MouseIcon       =   "frmCrearPersonaje.frx":225E
      MousePointer    =   99  'Custom
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2970
      Width           =   3570
   End
   Begin VB.TextBox txtPasswd 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   390
      MaxLength       =   25
      MouseIcon       =   "frmCrearPersonaje.frx":2F28
      MousePointer    =   99  'Custom
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2145
      Width           =   3570
   End
   Begin VB.TextBox txtCorreo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   390
      MaxLength       =   50
      MouseIcon       =   "frmCrearPersonaje.frx":3BF2
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3720
      Width           =   3570
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      ItemData        =   "frmCrearPersonaje.frx":48BC
      Left            =   1680
      List            =   "frmCrearPersonaje.frx":48F3
      MouseIcon       =   "frmCrearPersonaje.frx":498D
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   6345
      Width           =   1905
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      ItemData        =   "frmCrearPersonaje.frx":5657
      Left            =   1680
      List            =   "frmCrearPersonaje.frx":5661
      MouseIcon       =   "frmCrearPersonaje.frx":5674
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   7350
      Width           =   1905
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      ItemData        =   "frmCrearPersonaje.frx":633E
      Left            =   1680
      List            =   "frmCrearPersonaje.frx":6360
      MouseIcon       =   "frmCrearPersonaje.frx":63B9
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   6855
      Width           =   1905
   End
   Begin VB.Label Lbltotalconstitucion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   11490
      TabIndex        =   25
      Top             =   3780
      Width           =   120
   End
   Begin VB.Label Lbltotalcarisma 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   11490
      TabIndex        =   24
      Top             =   3270
      Width           =   120
   End
   Begin VB.Label Lbltotalinteligencia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   11490
      TabIndex        =   23
      Top             =   2775
      Width           =   120
   End
   Begin VB.Label Lbltotalagilidad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   11490
      TabIndex        =   22
      Top             =   2265
      Width           =   120
   End
   Begin VB.Label Lbltotalfuerza 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   11490
      TabIndex        =   21
      Top             =   1785
      Width           =   120
   End
   Begin VB.Label lblPlusFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   10905
      TabIndex        =   18
      Top             =   1785
      Width           =   285
   End
   Begin VB.Label lblPlusConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   10905
      TabIndex        =   17
      Top             =   3780
      Width           =   285
   End
   Begin VB.Label lblPlusCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   10905
      TabIndex        =   16
      Top             =   3270
      Width           =   285
   End
   Begin VB.Label lblPlusInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   10905
      TabIndex        =   15
      Top             =   2775
      Width           =   285
   End
   Begin VB.Label lblPlusAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   10905
      TabIndex        =   14
      Top             =   2265
      Width           =   285
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "+3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   5400
      TabIndex        =   13
      Top             =   8640
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   8760
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image boton 
      Height          =   1785
      Index           =   2
      Left            =   9195
      MouseIcon       =   "frmCrearPersonaje.frx":7083
      MousePointer    =   99  'Custom
      Top             =   5370
      Width           =   1785
   End
   Begin VB.Image boton 
      Height          =   480
      Index           =   1
      Left            =   270
      MouseIcon       =   "frmCrearPersonaje.frx":71D5
      MousePointer    =   99  'Custom
      Top             =   8265
      Width           =   1605
   End
   Begin VB.Image boton 
      Height          =   525
      Index           =   0
      Left            =   8070
      MouseIcon       =   "frmCrearPersonaje.frx":7327
      MousePointer    =   99  'Custom
      Top             =   8235
      Width           =   3705
   End
   Begin VB.Image imgHogar 
      Height          =   2850
      Left            =   6120
      Picture         =   "frmCrearPersonaje.frx":7479
      Top             =   9000
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.Label lbCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   10380
      TabIndex        =   12
      Top             =   3255
      Width           =   285
   End
   Begin VB.Label lbSabiduria 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   8640
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   10380
      TabIndex        =   10
      Top             =   2760
      Width           =   285
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   10380
      TabIndex        =   9
      Top             =   3765
      Width           =   285
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   10380
      TabIndex        =   8
      Top             =   2250
      Width           =   285
   End
   Begin VB.Label lbFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   10380
      TabIndex        =   1
      Top             =   1770
      Width           =   285
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SkillPoints As Byte
Private DadoStatus As Byte

Function CheckData() As Boolean

    If UserRaza = "" Then
        MsgBox "Seleccione la raza del personaje."
        Exit Function

    End If

    If UserSexo = "" Then
        MsgBox "Seleccione el sexo del personaje."
        Exit Function

    End If

    If UserClase = "" Then
        MsgBox "Seleccione la clase del personaje."
        Exit Function

    End If

    If UserHogar = "" Then
        MsgBox "Seleccione el hogar del personaje."
        Exit Function

    End If

    If SkillPoints > 0 Then
        MsgBox "Asigne los skillpoints del personaje."
        Exit Function
    End If

    Dim i As Integer

    For i = 1 To NUMATRIBUTOS

        If UserAtributos(i) = 0 Then
            MsgBox "Los atributos del personaje son invalidos."
            Exit Function

        End If

    Next i

    CheckData = True

End Function

Private Sub boton_Click(Index As Integer)
     
    Call Audio.PlayWave(SND_CLICK)
    
    Select Case Index

        Case 0
            
            UserName = txtNombre.Text
            UserPassword = txtPasswd.Text
            UserEmail = txtCorreo.Text
        
            UserRaza = lstRaza.List(lstRaza.ListIndex)
            UserSexo = lstGenero.List(lstGenero.ListIndex)
            UserClase = lstProfesion.List(lstProfesion.ListIndex)
            UserBanco = txtBanco
            UserPersonaje = txtPersonaje
            
            UserFuerza = Val(lbFuerza.Caption)
            UserAgilidad = Val(lbAgilidad.Caption)
            UserInteligencia = Val(lbInteligencia.Caption)
            UserCarisma = Val(lbCarisma.Caption)
            UserConstitucion = Val(lbConstitucion.Caption)
        
            UserAtributos(1) = Val(lbFuerza.Caption)
            UserAtributos(2) = Val(lbInteligencia.Caption)
            UserAtributos(3) = Val(lbAgilidad.Caption)
            UserAtributos(4) = Val(lbCarisma.Caption)
            UserAtributos(5) = Val(lbConstitucion.Caption)
                        
        
            If CheckDatos() Then
    

                EstadoLogin = CrearNuevoPj
        
                If Not frmMain.Socket1.Connected Then
                
                    frmMain.Socket1.Disconnect
                    frmMain.Socket1.Cleanup
              
                    frmMain.Socket1.HostName = CurServerIp
                    frmMain.Socket1.RemotePort = CurServerPort
                     frmMain.Socket1.Connect

                  EstadoLogin = E_MODO.CrearNuevoPj
                  
                  Call login
        
                Else
                    
                  frmMain.Socket1.HostName = CurServerIp
                  frmMain.Socket1.RemotePort = CurServerPort
                  frmMain.Socket1.Connect

                  EstadoLogin = E_MODO.CrearNuevoPj
                
                    Call login

                End If

            End If
        
        Case 1
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup

            frmConnect.MousePointer = 1
            
            Audio.StopWave
        
            Set frmConnect.FONDO.Picture = Interfaces.FrmConnect_Principal
            Me.Visible = False
            
            frmMain.Socket1.Disconnect
            AoDefResult = 0
            
        Case 2
           
            DadoStatus = 1
           
            Call Audio.PlayWave(SND_DICE)
            Call TirarDados
      
    End Select

End Sub

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long

    'Initialize randomizer
    Call Randomize(Timer)
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound

End Function

Function intNumeroaleatorio() As Integer
    Dim r       As String, s As Integer, T As Integer, seacabo As Boolean
    Dim gletras As String
    Dim gMaxNum As Integer
    seacabo = False

    Do While seacabo = False
        r = CStr(Timer)
        s = Len(r)
        T = mid$(r, s, 1)
        intNumeroaleatorio = (T * Int(gletras * Rnd))
        r = CStr(intNumeroaleatorio)
        s = Len(r)
        T = mid$(r, s, 1)
        intNumeroaleatorio = T

        If intNumeroaleatorio >= 0 And intNumeroaleatorio < gMaxNum Then
            seacabo = True

        End If

    Loop

End Function

Private Sub TirarDados()

    lbFuerza.Caption = RandomNumber(15, 18)
    lbInteligencia.Caption = RandomNumber(15, 18)
    lbAgilidad.Caption = RandomNumber(15, 18)
    lbCarisma.Caption = RandomNumber(15, 18)
    lbConstitucion.Caption = RandomNumber(15, 18)
    
    Call TotalAtributos

End Sub

Private Sub boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set boton(2).MouseIcon = Iconos.Ico_Mano
    Set boton(1).MouseIcon = Iconos.Ico_Mano
    Set boton(0).MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Command1_Click(Index As Integer)

    Call Audio.PlayWave(SND_CLICK)

End Sub

Private Sub Form_Activate()
    Call Audio.StopWave
    Call Audio.PlayWave("174.wav")
                

End Sub

Private Sub Form_Load()

    Set Me.Picture = Interfaces.FrmCrearPersonaje_Principal
    Set Me.MouseIcon = Iconos.Ico_Diablo

    Dim i As Integer
    lstProfesion.Clear

    For i = LBound(ListaClases) To UBound(ListaClases)
        lstProfesion.AddItem ListaClases(i)
    Next i

    lstProfesion.ListIndex = 0

    Call TirarDados

End Sub

Private Sub TotalAtributos()
       
     If Left(lblPlusFuerza, 1) = "+" Then
             Lbltotalfuerza = (Val(lbFuerza) + Val(Right(lblPlusFuerza, Len(lblPlusFuerza) - 1)))
     Else
             Lbltotalfuerza = (Val(lbFuerza) - Val(Right(lblPlusFuerza, Len(lblPlusFuerza) - 1)))
          End If

    
     If Left(lblPlusAgilidad, 1) = "+" Then
             Lbltotalagilidad = (Val(lbAgilidad) + Val(Right(lblPlusAgilidad, Len(lblPlusAgilidad) - 1)))
      Else
          Lbltotalagilidad = (Val(lbAgilidad) - Val(Right(lblPlusAgilidad, Len(lblPlusAgilidad) - 1)))
      End If
    
     If Left(lblPlusInteligencia, 1) = "+" Then
          Lbltotalinteligencia = (Val(lbInteligencia) + Val(Right(lblPlusInteligencia, Len(lblPlusInteligencia) - 1)))
      Else
          Lbltotalinteligencia = (Val(lbInteligencia) - Val(Right(lblPlusInteligencia, Len(lblPlusInteligencia) - 1)))
    End If

   
     If Left(lblPlusCarisma, 1) = "+" Then
          Lbltotalcarisma = (Val(lbCarisma) + Val(Right(lblPlusCarisma, Len(lblPlusCarisma) - 1)))
      Else
          Lbltotalcarisma = (Val(lbCarisma) - Val(Right(lblPlusCarisma, Len(lblPlusCarisma) - 1)))
      End If

     If Left(lblPlusConstitucion, 1) = "+" Then
          Lbltotalconstitucion = (Val(lbConstitucion) + Val(Right(lblPlusConstitucion, Len(lblPlusConstitucion) - 1)))
      Else
          Lbltotalconstitucion = (Val(lbConstitucion) - Val(Right(lblPlusConstitucion, Len(lblPlusConstitucion) - 1)))
      End If

    
End Sub

Private Sub lstRaza_Click()

    Select Case UCase(lstRaza.List(lstRaza.ListIndex))

        Case "HUMANO"
            lblPlusFuerza = "+1"
            lblPlusAgilidad = "+1"
            lblPlusConstitucion = "+2"
            lblPlusInteligencia = "+1"
            lblPlusCarisma = "+0"

        Case "ELFO"
            lblPlusAgilidad = "+2"
            lblPlusInteligencia = "+2"
            lblPlusCarisma = "+2"
            lblPlusConstitucion = "+1"
            lblPlusFuerza = "+1"

        Case "ELFO OSCURO"
            lblPlusFuerza = "+2"
            lblPlusAgilidad = "+4"
            lblPlusInteligencia = "+3"
            lblPlusCarisma = "-2"
            lblPlusConstitucion = "+0"

        Case "ENANO"
            lblPlusFuerza = "+3"
            lblPlusConstitucion = "+3"
            lblPlusInteligencia = "-3"
            lblPlusAgilidad = "+1"
            lblPlusCarisma = "-2"

        Case "GNOMO"
            lblPlusFuerza = "-1"
            lblPlusInteligencia = "+2"
            lblPlusAgilidad = "+3"
            lblPlusCarisma = "+1"
            lblPlusConstitucion = "+1"
            
        Case "HOBBIT"
            lblPlusFuerza = "-5"
            lblPlusInteligencia = "+4"
            lblPlusAgilidad = "+6"
            lblPlusCarisma = "+3"
            lblPlusConstitucion = "-1"
            
        Case "ORCO"
            lblPlusFuerza = "+5"
            lblPlusInteligencia = "-5"
            lblPlusAgilidad = "-6"
            lblPlusCarisma = "-3"
            lblPlusConstitucion = "+3"
            
        Case "LICANTROPO"
            lblPlusFuerza = "+0"
            lblPlusInteligencia = "-0"
            lblPlusAgilidad = "+0"
            lblPlusCarisma = "+0"
            lblPlusConstitucion = "+0"
            
        Case "VAMPIRO"
            lblPlusFuerza = "+2"
            lblPlusInteligencia = "+2"
            lblPlusAgilidad = "+1"
            lblPlusCarisma = "+1"
            lblPlusConstitucion = "+0"
            
        Case "CICLOPE"
            lblPlusFuerza = "+3"
            lblPlusInteligencia = "+0"
            lblPlusAgilidad = "+1"
            lblPlusCarisma = "+0"
            lblPlusConstitucion = "+2"

    End Select
    
    Call TotalAtributos

End Sub

Private Sub txtNombre_Change()

    txtNombre.Text = LTrim$(txtNombre.Text)

End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Function CheckDatos() As Boolean
             
              If txtNombre.Text = "" Then
                  MsgBox "Introduzca un nombre valido."
                  Exit Function
             End If
              
               If Len(txtNombre.Text) < 4 Then
                    MsgBox "El nombre debe tener más de 4 letras."
                    Exit Function
                End If
                
                If Right$(txtNombre.Text, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
                Exit Function
                End If
                
                If txtPasswd.Text = "" Then
                   MsgBox "Debes introducir una contraseña."
                   Exit Function
                End If

                If txtPasswdCheck.Text = "" Then
                   MsgBox "Debes repetir la contraseña repetida."
                   Exit Function
                End If
                
                 If txtPasswd.Text <> txtPasswdCheck.Text Then
                     MsgBox "Los passwords que tipeo no coinciden, por favor vuelva a ingresarlos."
                     Exit Function
                 End If
    
                If Not CheckMailString(UserEmail) Then
                    MsgBox "Direccion de mail invalida."
                    Exit Function
                End If
                
                
                If lstProfesion = "" Then
                    MsgBox "Seleccione la clase del personaje."
                    Exit Function

                End If
                
                If lstRaza = "" Then
                    MsgBox "Seleccione la raza del personaje."
                    Exit Function

                End If
    
                If lstGenero = "" Then
                    MsgBox "Seleccione el sexo del personaje."
                    Exit Function

                End If
                
                'If DadoStatus < 1 Then
                '    MsgBox "Debes tirar los dados antes!!"
                '    Exit Function
                'End If
                
                If UserBanco = "" Then
                   MsgBox "Elige una contraseña para el Banco."
                   Exit Function
                End If

              If UserPersonaje = "" Then
                  MsgBox "Pon una Contraseña para poder recuperar el personaje en caso de robo."
                  Exit Function
                End If

              If UserBanco = UserPersonaje Then
               MsgBox "Pon una Contraseña diferenta a la palabrasecreta."
               Exit Function
             End If

    CheckDatos = True

End Function

