VERSION 5.00
Begin VB.Form DeclaracionesFrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7365
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7380
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
   Icon            =   "frmDeclaraciones.frx":0000
   LinkTopic       =   "AoMania"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDeclaraciones.frx":000C
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   2760
      Top             =   480
   End
   Begin VB.Label btCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   4080
      TabIndex        =   1
      Top             =   6600
      Width           =   2325
   End
   Begin VB.Label btAccept 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   1080
      TabIndex        =   0
      Top             =   6600
      Width           =   2445
   End
End
Attribute VB_Name = "DeclaracionesFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Updater As Byte

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public StatusCondi As Long

'Declaramos el Api GetAsyncKeyState
Private Declare Function GetAsyncKeyState _
    Lib "user32" ( _
    ByVal vKey As Long) As Integer
  
Private Sub btAccept_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    btAccept.MousePointer = vbCustom
    btAccept.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub btCancel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    btCancel.MousePointer = vbCustom
    btCancel.MouseIcon = Iconos.Ico_Mano
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub Form_Terminate()
    If Updater = 1 Then
        SaveSetting App.exeName, "Updater", "Status", "0"
        Unload Me
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
      
    For i = 0 To 255
        'Consultamos el valor de la tecla mediante el Api. _
         Si se presionó devuelve -32767 y mostramos el valor de i
        If GetAsyncKeyState(i) = -32767 Then
  
            If i = "13" Then
                Call btAccept_Click
            End If
             
            If i = "27" Then
                Call btCancel_Click
            End If
  
        End If
    Next
End Sub


  
Private Sub Form_Load()
       
    Call InitializeCompression
    Call Mod_General.LoadInterfaces
    Call Mod_General.LoadClientSetup
    Call Mod_General.LoadIconos

    Updater = Val(GetSetting(App.exeName, "Updater", "Status", ""))
    '       Me.Picture = LoadPicture(App.Path & "/Graficos/cliente/condition.jpg")
    Me.MousePointer = vbCustom
    Me.MouseIcon = Iconos.Ico_Diablo
    Timer1.Interval = 50
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.MousePointer = vbCustom
    Me.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ReleaseCapture
    Call SendMessage(Me.HWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub btAccept_Click()
    
    Call Audio.PlayWave(SND_CLICK)
    
    If Updater = "1" Then
        SaveSetting App.exeName, "Updater", "Status", "0"
        Unload Me
        Call Mod_General.Main
    End If
    
    If Updater = "0" Then
       
        Call ShellExecute(Me.HWnd, "Open", App.Path & "\Update.exe", 0, 0, 1)
 
        SaveSetting App.exeName, "Updater", "Status", "1"
        Unload Me
       
    End If
End Sub



Private Sub btCancel_Click()
    Call Audio.PlayWave(SND_CLICK)
    If Updater = 1 Then
        SaveSetting App.exeName, "Updater", "Status", "0"
    End If
    Unload Me
End Sub
