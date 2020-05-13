VERSION 5.00
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Launcher AoMania Reborn"
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   676
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   720
      Top             =   330
   End
   Begin vbalProgBarLib6.vbalProgressBar ProgressBar1 
      Height          =   195
      Left            =   1785
      TabIndex        =   1
      Top             =   3540
      Visible         =   0   'False
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   344
      Picture         =   "frmMain.frx":08CA
      ForeColor       =   8421504
      BarForeColor    =   8421504
      BarPicture      =   "frmMain.frx":08E6
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
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label LSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L size"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   3615
      TabIndex        =   2
      Top             =   5220
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label txtUpdate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Info"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2490
      TabIndex        =   0
      Top             =   5220
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdCerrar 
      Height          =   555
      Left            =   9345
      Top             =   270
      Width           =   555
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' LauncherAoM 1.0.0
' By Bassinger [www.AoMania.Net]

Option Explicit

Private Sub cmdCerrar_Click()
    UnloadAllForms
End Sub

Private Sub cmdCerrar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
     With frmMain
         .MouseIcon = Iconos.Ico_Mano
     End With
     
End Sub

Private Sub Form_Load()

   If Launcher.Play = 1 Then
        Launcher.Play = 0
   End If
     
    With frmMain
        .MouseIcon = Iconos.Ico_Diablo

    End With

    ProgressBar1.Value = 0
    LSize.Caption = ""
    ProgressBar1.Text = ""
    txtUpdate.Caption = ""
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     
     With frmMain
         .MouseIcon = Iconos.Ico_Diablo
     End With
     
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
                vtData = Inet1.GetChunk((1048576 / 2), icByteArray)
                DoEvents
                
                
                Do While Not Len(vtData) = 0
                    tempArray = vtData
                    Put #1, , tempArray
                        
                ProgressBar1.BarPicture = Interfaces.BLlena
                txtUpdate.Caption = "Descargando: "
            
                vtData = Inet1.GetChunk((1048576 / 2), icByteArray)
                    
                    ProgressBar1.Value = ProgressBar1.Value + Len(vtData) * 2
                    LSize.Caption = CLng((ProgressBar1.Value + Len(vtData) * 2) / 1000000) & " MBs de " & CLng((FileSize / 1000000)) & " MBs"
                    ProgressBar1.Text = Round(CDbl(ProgressBar1.Value) * CDbl(100) / CDbl(ProgressBar1.max), 2) _
                            & "%"
   
                  '  DoEvents
                Loop
            Close #1
            
            'LSize.Caption = FileSize & " bytes"
            LSize.Visible = False
            
            txtUpdate.Caption = "¡Ok! Actualización finalizada."
            
            ProgressBar1.Value = 0
            
            bDone = True
    End Select
    
End Sub

Private Sub Timer1_Timer()

    Static TimerUpdater As Long

    TimerUpdater = TimerUpdater + "1"
   
    If TimerUpdater = "80" Then
        Call Analizar
        'SetUpdate = "1"
        'TimerUpdater = "0"
        Timer1.Enabled = False

    End If

End Sub
