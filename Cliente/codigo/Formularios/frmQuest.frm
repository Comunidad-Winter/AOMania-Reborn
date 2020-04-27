VERSION 5.00
Begin VB.Form frmQuest 
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8700
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
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin AoManiaClienteGM.ChameleonBtn Salir 
      Height          =   405
      Left            =   3675
      TabIndex        =   5
      Top             =   5445
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Salir"
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
      FCOLO           =   16576
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmQuest.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AoManiaClienteGM.ChameleonBtn cmdIniciar 
      Height          =   405
      Left            =   3675
      TabIndex        =   3
      Top             =   4245
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Iniciar misión"
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
      FCOLO           =   16576
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmQuest.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AoManiaClienteGM.ChameleonBtn cmdEntregar 
      Height          =   420
      Left            =   3675
      TabIndex        =   4
      Top             =   4845
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   741
      BTYPE           =   3
      TX              =   "Entregar misión"
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
      FCOLO           =   16576
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmQuest.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Height          =   675
      Left            =   5340
      ScaleHeight     =   615
      ScaleWidth      =   750
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   225
      Width           =   810
   End
   Begin VB.ListBox ListQuest 
      Height          =   2595
      Left            =   3630
      TabIndex        =   2
      Top             =   1215
      Width           =   4800
   End
   Begin VB.TextBox InfoQuest 
      Height          =   5775
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   3270
   End
End
Attribute VB_Name = "frmQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

