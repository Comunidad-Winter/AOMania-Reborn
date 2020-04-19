VERSION 5.00
Begin VB.Form FrmSos 
   Caption         =   "Formulario de mensaje a administradores"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5730
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
   ScaleHeight     =   6765
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar sin Enviar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   720
      TabIndex        =   11
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3240
      TabIndex        =   10
      Top             =   5880
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "FmSos.frx":0000
      Top             =   2520
      Width           =   4935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Sugerencias"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   3720
      TabIndex        =   8
      Top             =   2040
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Reporte de bug"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   7
      Top             =   2040
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Denuncias"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   6
      Top             =   1680
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Enviar SOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "almacenado. Todo mensaje formado será eliminado."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   720
      TabIndex        =   14
      Top             =   5280
      Width           =   4410
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "no hay ningún administrador conectado, quedará"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   13
      Top             =   5040
      Width           =   4050
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tu mensaje será guardado en nuestra base de datos. Si"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   12
      Top             =   4800
      Width           =   4620
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "junto con el manuel de juego."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Width           =   2445
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "antes de llamar, leer la ayuda que se encuentra en opciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   4965
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "está bien formulada, no se recibirá el mensaje. Recuerda,"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   4770
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "llamado indebido será gravemente penado. Si tu consulta no"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Por favor, escribe el mensaje para el administrador. Un"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   4515
   End
End
Attribute VB_Name = "FrmSos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    If Option1(0) = True Then
        Call SendData("/SHOW_SOS " & Text1.Text)
    ElseIf Option1(1) = True Then
        Call SendData("/SHOW_DENUNCIA " & Text1.Text)
    ElseIf Option1(2) = True Then
        Call SendData("/SHOW_BUG " & Text1.Text)
    ElseIf Option1(3) = True Then
        Call SendData("/SHOW_SUGERENCIA " & Text1.Text)
    Else
        Call MsgBox("No has seleccionado que tipo de consulta deseas realizar", vbInformation)

    End If

End Sub

Private Sub Command2_Click()
    Unload Me

End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = Option1.LBound To Option1.UBound
     
        Option1(i).value = False
      
    Next i

End Sub

Private Sub Option1_Click(Index As Integer)
     
    If Index = 0 Then 'Enviar Sos
        Text1.Locked = False
        Label6.Caption = "¡Por favor explique correctamente el motivo de su"
        Label7.Caption = "consulta!"
        Label8.Caption = ""

    End If
     
    If Index = 1 Then 'Denuncia
        Text1.Locked = False
        Label6.Caption = "¡Por favor explique correctamente el motivo de su"
        Label7.Caption = "Denuncia lo mas detalladamente posible!"
        Label8.Caption = " "

    End If
     
    If Index = 2 Then 'Reporte de bug
        Text1.Locked = False
        Label6.Caption = "Se dará prioridad a su consulta enviando un mensaje a los"
        Label7.Caption = "administradores conectados, por favor utilize ésta opción"
        Label8.Caption = "responsablemente."

    End If
     
    If Index = 3 Then 'Sugerencias
        Text1.Locked = False
        Label6.Caption = "Su sugerencia SERÁ leída por un miembro del staff, y será"
        Label7.Caption = "tomada en cuenta para futuros cambios."
        Label8.Caption = " "

    End If
     
End Sub

Private Sub Text1_Click()

    If Option1(0) = True Then
        If Text1.Text = "Escriba Aquí su Mensaje." Then
            Text1.Text = ""

        End If

    ElseIf Option1(1) = True Then

        If Text1.Text = "Escriba Aquí su Mensaje." Then
            Text1.Text = ""

        End If

    ElseIf Option1(2) = True Then

        If Text1.Text = "Escriba Aquí su Mensaje." Then
            Text1.Text = ""

        End If

    ElseIf Option1(3) = True Then

        If Text1.Text = "Escriba Aquí su Mensaje." Then
            Text1.Text = ""

        End If

    End If

End Sub
