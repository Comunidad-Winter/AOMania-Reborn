VERSION 5.00
Begin VB.Form frmIns 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Instrucciones"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "6. Borrar la carpeta que contenia graficos convertidos en ^2."
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
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   6120
      Width           =   8115
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmIns.frx":0000
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
      Height          =   720
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   8175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmIns.frx":00C9
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
      Height          =   600
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   8115
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "3. Poner el numero de gráfico mas alto que tengan en la carpeta de los graficos."
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
      Height          =   480
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   8115
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmIns.frx":0163
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
      Height          =   720
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   8115
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Es recomendable poner este ejecutable y la libreria: ""progressbar-xp.ocx"" en la carpeta de su cliente."
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
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   8055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "¿Como usarlo?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Introduccion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmIns.frx":020D
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7575
   End
End
Attribute VB_Name = "frmIns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

