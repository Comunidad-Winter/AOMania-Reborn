VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Cambios de mana para Brujos!"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7410
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   158
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   494
   StartUpPosition =   2  'CenterScreen
   Begin ChangeManaBrujos.ChameleonBtn ChangeMana 
      Height          =   450
      Left            =   2445
      TabIndex        =   2
      Top             =   1800
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   794
      BTYPE           =   3
      TX              =   "Change Mana"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   65535
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":08CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   45
      TabIndex        =   1
      Top             =   105
      Width           =   7305
   End
   Begin ChangeManaBrujos.ChameleonBtn LoadFile 
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   1785
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "Load file"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   65535
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":08E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChangeMana_Click()
    
    Dim AMana As Integer
    Dim NMana As Integer
    Dim Atributos As Byte
    Dim Arch As String
    Dim Name As String
    Dim Nivel As Byte
    
    
    Dim N As Integer
     
     If NumBrujos = 0 Then
        MsgBox "Primero debes de cargar lista de charfile.", vbCritical
        Exit Sub
     End If
     
     For N = 1 To NumBrujos
         
         Arch = App.Path & "\charfile\" & Brujo(N).Name & ".chr"
         
         Name = Brujo(N).Name
         
         Atributos = Brujo(N).Inteligencia
         Nivel = Brujo(N).Nivel
         AMana = Brujo(N).Mana
         
         NMana = ((Atributos * 2.65) * (Nivel - 1)) + 100
         
         Call WriteVar(Arch, "STATS", "MaxMAN", NMana)
         
         Call LogCambios(Name, Nivel, Atributos, AMana, NMana)
     Next N
     
     MsgBox "Todos los charfiles fueron modificados.", vbInformation
     
End Sub

Private Sub LoadFile_Click()
     
     Dim Name As String
     
     Dim i As Long, nombre As String
        nombre = Dir(App.Path & "\charfile\*.chr")
        Do While nombre <> ""
         
        i = i + 1
         nombre = Dir
         Name = ReadField(1, nombre, 46)
         
         If UCase$(GetVar(App.Path & "\charfile\" & nombre, "INIT", "Clase")) = "BRUJO" Then
             NumBrujos = NumBrujos + 1
             Brujo(NumBrujos).Name = Name
             Brujo(NumBrujos).Inteligencia = GetVar(App.Path & "\charfile\" & nombre, "ATRIBUTOS", "AT3")
             Brujo(NumBrujos).Mana = GetVar(App.Path & "\charfile\" & nombre, "STATS", "MaxMAN")
             Brujo(NumBrujos).Nivel = GetVar(App.Path & "\charfile\" & nombre, "STATS", "ELV")
             List1.AddItem Name
         End If
         
        Loop
        MsgBox "Hay " & CStr(NumBrujos) & " arhivos"
        
End Sub
