VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{7C1A7C04-1571-4390-8302-BA83F3FA717E}#1.0#0"; "mail.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
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
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   1320
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin mail.sendmail sendmail1 
      Height          =   480
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   If FindPreviousInstance Then
    End
   Else
     Call RevisaRegistro
   End If
End Sub

Private Sub sendmail1_SendSuccesful()
    ValReg = ValReg + 1
    Call DatosRegistro
End Sub

Private Sub sendmail1_Progress(lPercentCompete As Long)
    'Visualiza el porcentaje del progreso del envío en el Label
    'lblProgress = lPercentCompete & "% completado"

End Sub

Private Sub sendmail1_SendFailed(Explanation As String)
    'En caso de fallar el envío se dispara este evento con la descripción del error
    'MsgBox ("El envío del Email falló por las posibles razones:: " & vbCrLf & Explanation)
   ' lblProgress = ""
    Screen.MousePointer = vbDefault
  ' cmdSend.Enabled = True
    End
End Sub


Sub EnviaMail()
    
    Timer1.Enabled = False
    ' NewRegistro.Enabled = False
    'lstStatus.Clear
    Screen.MousePointer = vbHourglass

    With sendmail1

        'Valida (opcional)
        .SMTPHostValidacion = VALIDATE_HOST_NONE
        'Valida la sintaxis de l cuenta (opcional)
        .ValidarEmail = VALIDATE_SYNTAX
        'Opcional
        .Delimitador = ";"
        'Texto  para visualizar en el campo De (opcional)
        .FromDisplayName = "Registro AoMania"
        'Requerido (Nombre del servidor SMTP)
        .SMTPHost = "mail.argentumania.es"
        'Requerido
        .Remitente = "Soporte@Argentumania.es"
        'Requerido
        .Destinatario = Email
        'Asunto del mensaje
        .Asunto = "Registro del Nick '" & Nombre & "'en AoMania."
        'Cuerpodel mensaje
        .Mensaje = "Hola <b>" & Nombre & "</b>." & vbCrLf & vbCrLf & _
                            "Tu personaje se ha creado correctamente. A continuación te dejamos aquí los datos de tu personaje:" & vbCrLf & vbCrLf & _
                            "<b>Nick</b>: " & Nombre & "." & vbCrLf & _
                            "<b>Contraseña</b>: " & Password & "." & vbCrLf & _
                            "<b>Clase</b>: " & Clase & "." & vbCrLf & _
                            "<b>Raza</b>: " & Raza & "." & vbCrLf & vbCrLf & _
                            "¡<b>RECUERDA NO DAR SU CONTRASEÑA A NADIE</b>!" & vbCrLf & vbCrLf & _
                            "Un saludo," & vbCrLf & _
                            "El Staff de AoMania."
        
        'Adjunto (opcional)
      '  .Adjunto = Trim(txtAttach.Text)
        
        'Opcional (Prioridad del mensaje)
        .Prioridad = Baja
        'Opcional (si requiere autentificación)
        .UsarLoginSMTP = True
        'Requerido si requiere autentificación
        .Usuario = "Soporte@Argentumania.es"
        .Password = "Loleitor1@"
        
       ' txtServer.Text = .SMTPHost
       'Opcional (por defectoutiliza el Tipo MIME)
       .Codificacion = MIME_ENCODE
       
       'Envia el Mail
       .EnviarEmail
    
    End With
    Screen.MousePointer = vbDefault
    'cmdSend.Enabled = True

End Sub

Private Sub Timer1_Timer()
    Static Timer As Long
    Timer = Timer + 1
        
    If Timer = "5" Then
      Call EnviaRegistro
      Timer = 0
    End If
    
End Sub
