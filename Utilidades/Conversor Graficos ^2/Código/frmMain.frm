VERSION 5.00
Object = "{DA729E34-689F-49EA-A856-B57046630B73}#1.0#0"; "progressbar-xp.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AO- Conversor de graficos a ^2"
   ClientHeight    =   3465
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   9465
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   231
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   631
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pProgreso 
      BackColor       =   &H00404040&
      ForeColor       =   &H00404040&
      Height          =   1575
      Left            =   3120
      ScaleHeight     =   1515
      ScaleWidth      =   5955
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   6015
      Begin Proyecto2.XP_ProgressBar PB1 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   4210752
         Scrolling       =   9
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 / 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   5655
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Convirtiendo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.TextBox txtDirFinal 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   8895
   End
   Begin VB.TextBox txtDirInicial 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   8895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convertir graficos"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Número del último gráfico:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Directorio de los gráficos convertidos a ^2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Directorio de los gráficos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.Menu CONVERSOR 
      Caption         =   "Conversor"
      Begin VB.Menu CONVERTIR 
         Caption         =   "Convertir"
         Shortcut        =   ^C
      End
      Begin VB.Menu SP1 
         Caption         =   "-"
      End
      Begin VB.Menu SALIR 
         Caption         =   "Salir"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu AYUDA 
      Caption         =   "Ayuda"
      Begin VB.Menu INS 
         Caption         =   "Instrucciones"
      End
      Begin VB.Menu AD 
         Caption         =   "Acerca de..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AD_Click()
MsgBox "Conversor de gráficos de Argentum Online a ^2, para su correcto funcionamiento cuando se usa DirectX8 o OpenGL como motor gráfico." & vbNewLine & vbNewLine & "By Thusing", vbInformation, "Conversor"
End Sub

Private Sub Command1_Click()
Call Conversion
End Sub

Private Sub CONVERTIR_Click()
Call Conversion
End Sub

Private Sub Form_Load()
txtDirInicial.Text = App.Path & "\Graficos"
txtDirFinal.Text = App.Path & "\Graficos en ^2"
End Sub

Sub Conversion()
Dim i As Long
Dim numgraficos As String
Dim imagen As IPictureDisp
Dim alto As Integer, ancho As Integer
Dim newdimension As Integer
Dim porcentaje As Integer

Unload frmIns

numgraficos = Val(frmMain.Text1.Text)

If Not FileExist(txtDirInicial, vbDirectory) Then
MsgBox "La carpeta en donde se encuentran los graficos no existe." & vbNewLine & "Asegurese de tener este ejecutable en la carpeta del cliente, y revisar bien el directorio de los graficos.", vbCritical, "Error"
Exit Sub
End If

If Not FileExist(txtDirFinal, vbDirectory) Then
MsgBox "La carpeta en donde se convierten los graficos a ^2 no existe, se creará.", vbInformation, "Carpeta de conversion"
MkDir txtDirFinal
End If

If Val(numgraficos) <= 0 Or numgraficos = "" Then
MsgBox "Ingrese el número de grafico mas alto.", vbCritical, "Error"
Exit Sub
End If

MsgBox "Este proceso puede tardar varios minutos, dependiendo de la cantidad de graficos a convertir" & vbNewLine & "Por favor, espere.", vbInformation, "Conversion"


frmMain.Label5.Caption = "0%"
frmMain.PB1.Min = 1
frmMain.PB1.Max = 100
frmMain.PB1.Value = 0

frmMain.txtDirInicial.Enabled = False
frmMain.txtDirFinal.Enabled = False
frmMain.Text1.Enabled = False
frmMain.Command1.Enabled = False
frmMain.SALIR.Enabled = False
frmMain.CONVERTIR.Enabled = False
frmMain.AD.Enabled = False
frmMain.INS.Enabled = False

frmMain.pProgreso.Visible = True


' Deshabilitar el botón de cerrar el formulario
Dim hMenu As Long
'
hMenu = GetSystemMenu(frmMain.hWnd, 0)
Call ModifyMenu(hMenu, SC_CLOSE, MF_BYCOMMAND Or MF_GRAYED, -10, "Close")
'
Call DrawMenuBar(frmMain.hWnd)

For i = 1 To numgraficos
    
    porcentaje = (i / numgraficos) * 100
    PB1.Value = porcentaje
    frmMain.Label5.Caption = porcentaje & "%"
    frmMain.Label6.Caption = i & " / " & numgraficos
    DoEvents
    
    If FileExist(txtDirInicial & "\" & i & ".bmp", vbArchive) Then
        If Not Tamaño(i) = 0 Then
        Set imagen = LoadPicture(txtDirInicial & "\" & i & ".bmp")
        alto = ScaleY(imagen.Height, vbHimetric, vbPixels)
        ancho = ScaleX(imagen.Width, vbHimetric, vbPixels)
        'If i = 20 Then Stop
        If alto > ancho Then
            newdimension = ObtenerDimension(alto)
        ElseIf ancho > alto Then
            newdimension = ObtenerDimension(ancho)
        ElseIf ancho = alto Then
            newdimension = ObtenerDimension(ancho)
        Else
            'Son iguales no importa cual usemos
            If EsPotenciaDos(ancho) Then
                newdimension = ancho + 4
            Else
                ObtenerDimension (ancho)
            End If
        End If
        frmMain.Picture1.Width = newdimension
        frmMain.Picture1.Height = newdimension
        Call frmMain.Picture1.Cls
        Call frmMain.Picture1.PaintPicture(imagen, 0, 0, ancho, alto)
        Call SavePicture(frmMain.Picture1.Image, txtDirFinal & "\" & i & ".bmp")
    End If
    End If
    DoEvents

Next i

' Habilitar el botón de cerrar el formulario
hMenu = GetSystemMenu(frmMain.hWnd, 0)
Call ModifyMenu(hMenu, -10, MF_BYCOMMAND Or MF_ENABLED, SC_CLOSE, "Close")
Call DrawMenuBar(frmMain.hWnd)

frmMain.txtDirInicial.Enabled = True
frmMain.txtDirFinal.Enabled = True
frmMain.Text1.Enabled = True
frmMain.Command1.Enabled = True
frmMain.SALIR.Enabled = True
frmMain.CONVERTIR.Enabled = True
frmMain.AD.Enabled = True
frmMain.INS.Enabled = True

Text1.Text = vbNullString

frmMain.pProgreso.Visible = False

MsgBox "Gráficos convertidos a ^2 correctamente.", vbInformation, "Conversion completa"
End Sub

Private Sub INS_Click()
frmIns.Show , frmMain
End Sub

Private Sub SALIR_Click()
End
End Sub
