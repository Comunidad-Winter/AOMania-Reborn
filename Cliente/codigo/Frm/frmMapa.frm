VERSION 5.00
Begin VB.Form frmMapa 
   BackColor       =   &H0000C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   7845
   ClientLeft      =   1080
   ClientTop       =   2250
   ClientWidth     =   8805
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H80000007&
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   7215
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Funciones del Api
'-------------------------------------------------------------
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long, _
    ByVal X3 As Long, _
    ByVal Y3 As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" (ByVal HWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Sub Command1_Click()

    Dim ret        As Long
    Dim L          As Long
    Dim Ancho_form As Long
    Dim Alto_form  As Long
    Dim OldScale   As Integer
    
    ' guarda el scale del form
    OldScale = ScaleMode
    
    ' cambia la escala ya que el api trabaja con pixeles
    ScaleMode = vbPixels
    
    'Ancho y alto del form en pixeles
    Ancho_form = Me.ScaleWidth
    Alto_form = Me.ScaleHeight
    
    'Crea la región
    ret = CreateRoundRectRgn(10, 35, Ancho_form, Alto_form + 25, 0, 0)
    
    'Aplica la nueva región al formulario
    L = SetWindowRgn(Me.HWnd, ret, True)
    ' reestablece la escala que tenia el formulario
    ScaleMode = OldScale

End Sub

Private Sub cmdCerrar_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    'Valores máximos y mínimos para el ScrollBar
    Me.Left = 0
    Me.Top = 0
    Set frmMapa.Picture = Interfaces.FrmMapa_Principal

End Sub
 
