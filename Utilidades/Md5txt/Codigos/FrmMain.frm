VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFF00&
   Caption         =   "FrmMain"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Md5txt"
   ScaleHeight     =   137
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   548
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TEXT a MD5"
      Height          =   360
      Left            =   3120
      TabIndex        =   4
      Top             =   1560
      Width           =   1290
   End
   Begin VB.TextBox TxtMD5 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   6495
   End
   Begin VB.TextBox TxtEncrypt 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MD5:"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Texto:"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
     TxtMD5.Text = MD5String(TxtEncrypt.Text)
End Sub

'Private Sub Form_KeyPress(sender As Object, e As KeyPressEventArgs)
'
'    If e.Control And e.Alt And e.KeyCode = Keys.S Then
'      MsgBox ("ahora")
'    End If
'  End Sub

Private Sub Form_Load()
Call Disco
End Sub

Private Sub Timer1_Timer()
    'Boton ESC uso
   
    Dim i As Integer
   
    For i = 1 To 250
        
            If GetAsyncKeyState(i) = -32767 Then
             If frmMain.Visible = True Then
            
                If i = 164 Then
                frmMain.Visible = False
                 ViewHDD.Show
                 
               End If
              End If
            End If

     
    Next

End Sub

