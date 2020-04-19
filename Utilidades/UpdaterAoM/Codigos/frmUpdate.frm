VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form frmUpdate 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Update AoMania"
   ClientHeight    =   3450
   ClientLeft      =   -60
   ClientTop       =   -165
   ClientWidth     =   9060
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Update AoMania"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "frmUpdate.frx":31C32A
   ScaleHeight     =   3450
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Ejecutador 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7800
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6960
      Top             =   120
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin vbalProgBarLib6.vbalProgressBar ProgressBar1 
      Height          =   810
      Left            =   1440
      TabIndex        =   0
      Top             =   1440
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   1429
      Picture         =   "frmUpdate.frx":3437A5
      BackColor       =   4194368
      ForeColor       =   16777215
      BorderStyle     =   0
      BarPicture      =   "frmUpdate.frx":3493C6
      BarPictureMode  =   0
      BackPictureMode =   0
      ShowText        =   -1  'True
      Text            =   "[0% Completado]"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label btCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3360
      TabIndex        =   3
      Top             =   2520
      Width           =   2205
   End
   Begin VB.Label LSize 
      BackStyle       =   0  'Transparent
      Caption         =   "0 MBs de 0 MBs"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label txtUpdate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ACTUALIZACION OK"
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
      Height          =   210
      Left            =   3480
      TabIndex        =   2
      Top             =   1680
      Width           =   1950
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20&
Dim Directory As String, bDone As Boolean, dError As Boolean, F As Integer
Rem Programado por Shedark

Public SetUpdateChange As Long
Public SetUpdate As Long
Public TimerOn As Byte


Private Sub Analizar()

    Dim i As Integer, iX As Integer, tX As Integer, DifX As Integer, dNum As String
    
    'lEstado.Caption = "Obteniendo datos..."
    
    iX = Inet1.OpenURL("http://argentumania.es/cosas/parches/VEREXE.TXT") 'Host
    tX = LeerInt(FileUpdate)
    
    DifX = iX - tX
    
    If Not (DifX = 0) Then
      ProgressBar1.Visible = True

       For i = 1 To DifX
            Inet1.AccessType = icUseDefault
            dNum = i + tX
            
            #If BuscarLinks Then 'Buscamos el link en el host (1)
                Inet1.URL = Inet1.OpenURL("http://argentumania.es/cosas/parches/Parche" & dNum & ".zip") 'Host
            #Else                'Generamos Link por defecto (0)
                Inet1.URL = "http://argentumania.es/cosas/parches/Parche" & dNum & ".zip" 'Host
            #End If
            
            Directory = App.Path & "\Libs\Configuracion\Parche" & dNum & ".zip"
            bDone = False
            dError = False
            
            'lURL.Caption = Inet1.URL
            'lName.Caption = "Parche" & dNum & ".zip"
            'lDirectorio.Caption = App.Path & "\"
                
            frmUpdate.Inet1.Execute , "GET"
            
            Do While bDone = False
            DoEvents
            Loop
            
            If dError Then Exit Sub
            
            UnZip Directory, App.Path & "\"
            Kill Directory
        Next i
    End If

    Call GuardarInt(FileUpdate, iX)
    SaveSetting "AoMania", "Updater", "Status", "1"
    

    ProgressBar1.Value = 0

   ProgressBar1.Visible = False
   
   txtUpdate.Visible = True
   TimerOn = 1
   SetUpdate = "1"
   txtUpdate.Left = 3480
   Ejecutador.Enabled = True
   txtUpdate.Caption = "Actualización OK"
   SetUpdateChange = "1"
   

End Sub

Private Sub btCancel_Click()
SaveSetting "AoMania", "Updater", "Status", "0"
Call UnloadAllForms
End Sub

Private Sub btCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      btCancel.MousePointer = vbCustom
      btCancel.MouseIcon = Iconos.Mano
End Sub

Private Sub Ejecutador_Timer()
     Static Timer As Long
     Timer = Timer + 1
     If Timer = 2 Then
       Unload Me
      Call ShellExecute(Me.hWnd, "Open", App.Path & "\AoMania.exe", 0, 0, 1)
     End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Me.MousePointer = vbCustom
       Me.MouseIcon = Iconos.Ico_Diablo
End Sub

Private Sub Form_Load()

Set ProgressBar1.Picture = Interfaces.BVacia

'frmUpdate.Picture = LoadPicture(App.Path & "\Interfaces\update.jpg")

Me.MousePointer = vbCustom
Set Me.MouseIcon = Iconos.Ico_Diablo
ProgressBar1.Value = 0
'ProgressBar1.Height = 0
LSize.Caption = ""
ProgressBar1.Text = ""

End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    Select Case State
        Case icError
            ProgressBar1.Visible = False
            SetUpdate = "1"
            txtUpdate.Left = 2320
            txtUpdate.Caption = "Error en la conexión, descarga abortada."
            bDone = True
            dError = True
            Exit Sub
        Case icResponseCompleted
            Dim vtData As Variant
            Dim tempArray() As Byte
            Dim FileSize As Long

            
            FileSize = Inet1.GetHeader("Content-length")
            ProgressBar1.max = FileSize
            
            
            Open Directory For Binary Access Write As #1
                vtData = Inet1.GetChunk(1024, icByteArray)
                DoEvents
                
                
                Do While Not Len(vtData) = 0
                    tempArray = vtData
                    Put #1, , tempArray
                        
                ProgressBar1.BarPicture = Interfaces.BLlena
            
                vtData = Inet1.GetChunk(1024, icByteArray)
                    
                    ProgressBar1.Value = ProgressBar1.Value + Len(vtData) * 2
                    LSize.Caption = (ProgressBar1.Value + Len(vtData) * 2) / 1000000 & "MBs de " & (FileSize / 1000000) & "MBs"
                    ProgressBar1.Text = Round(CDbl(ProgressBar1.Value) * CDbl(100) / CDbl(ProgressBar1.max), 2) _
                            & "%"
                    DoEvents
                Loop
            Close #1
            
            LSize.Caption = FileSize & "bytes"
            ProgressBar1.Value = 0
            
            bDone = True
    End Select
End Sub

Private Function LeerInt(ByVal Ruta As String) As Integer
    F = FreeFile
    Open Ruta For Input As F
    LeerInt = Input$(LOF(F), #F)
    Close #F
End Function

Private Sub GuardarInt(ByVal Ruta As String, ByVal Data As Integer)
    F = FreeFile
    Open Ruta For Output As F
    Print #F, Data
    Close #F
End Sub


Private Sub Timer1_Timer()
Static TimerUpdater As Long

If TimerOn = 1 Then
    TimerUpdater = 0
End If

 TimerUpdater = TimerUpdater + "1"
 
 If SetUpdateChange = 0 Then
 If SetUpdate = 0 Then
 If TimerUpdater = "80" Then
  Call Analizar
  SetUpdate = "1"
  TimerUpdater = "0"
  Timer1.Enabled = False
 End If
 End If
End If
 If SetUpdateChange = "1" Then
     TimerUpdater = "0"
     SetUpdateChange = "0"
     'Timer1.Enabled = False
End If

If SetUpdate = 1 Then
      If TimerUpdater = "120" Then
       Unload Me
       'frmdeclaraciones.Visible = True
       'frmdeclaraciones.StatusCondi = "1"
 End If
 End If
  
End Sub
