VERSION 5.00
Begin VB.Form frmSkills3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7890
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   7575
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmSkills3.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   526
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Command1 
      Height          =   375
      Index           =   51
      Left            =   4440
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image Command1 
      Height          =   375
      Index           =   50
      Left            =   5160
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   4920
      TabIndex        =   48
      Top             =   2640
      Width           =   90
   End
   Begin VB.Image Command1 
      Height          =   375
      Index           =   49
      Left            =   4440
      Top             =   2040
      Width           =   375
   End
   Begin VB.Image Command1 
      Height          =   375
      Index           =   48
      Left            =   5280
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   4920
      TabIndex        =   47
      Top             =   2040
      Width           =   90
   End
   Begin VB.Image Command1 
      Height          =   375
      Index           =   47
      Left            =   4440
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image Command1 
      Height          =   375
      Index           =   46
      Left            =   5160
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image Command1 
      Height          =   255
      Index           =   45
      Left            =   4320
      Top             =   960
      Width           =   375
   End
   Begin VB.Image Command1 
      Height          =   255
      Index           =   44
      Left            =   5160
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   4920
      TabIndex        =   46
      Top             =   1440
      Width           =   90
   End
   Begin VB.Label Text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   4920
      TabIndex        =   45
      Top             =   960
      Width           =   90
   End
   Begin VB.Image Command2 
      Height          =   495
      Left            =   5280
      MouseIcon       =   "frmSkills3.frx":0CCA
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   42
      Left            =   7020
      Top             =   6270
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   43
      Left            =   6270
      Top             =   6360
      Width           =   180
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   22
      Left            =   6510
      TabIndex        =   44
      Top             =   6330
      Width           =   450
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   21
      Left            =   600
      TabIndex        =   43
      Top             =   7560
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   0
      Left            =   630
      TabIndex        =   42
      Top             =   420
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   1
      Left            =   630
      TabIndex        =   41
      Top             =   765
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   2
      Left            =   630
      TabIndex        =   40
      Top             =   1110
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   3
      Left            =   630
      TabIndex        =   39
      Top             =   1455
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   4
      Left            =   630
      TabIndex        =   38
      Top             =   1800
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   5
      Left            =   630
      TabIndex        =   37
      Top             =   2145
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   6
      Left            =   630
      TabIndex        =   36
      Top             =   2445
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   7
      Left            =   630
      TabIndex        =   35
      Top             =   2835
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   8
      Left            =   630
      TabIndex        =   34
      Top             =   3195
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   9
      Left            =   630
      TabIndex        =   33
      Top             =   3540
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   10
      Left            =   630
      TabIndex        =   32
      Top             =   3885
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   11
      Left            =   630
      TabIndex        =   31
      Top             =   4230
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   270
      Index           =   1
      Left            =   3105
      TabIndex        =   30
      Top             =   855
      Width           =   450
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   2
      Left            =   3105
      TabIndex        =   29
      Top             =   1410
      Width           =   450
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   3
      Left            =   3105
      TabIndex        =   28
      Top             =   1950
      Width           =   450
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   4
      Left            =   3105
      TabIndex        =   27
      Top             =   2505
      Width           =   450
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   5
      Left            =   3105
      TabIndex        =   26
      Top             =   3045
      Width           =   450
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   6
      Left            =   3105
      TabIndex        =   25
      Top             =   3615
      Width           =   450
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   7
      Left            =   3105
      TabIndex        =   24
      Top             =   4140
      Width           =   450
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   8
      Left            =   3105
      TabIndex        =   23
      Top             =   4695
      Width           =   450
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   9
      Left            =   3105
      TabIndex        =   22
      Top             =   5235
      Width           =   450
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   10
      Left            =   3105
      TabIndex        =   21
      Top             =   5790
      Width           =   450
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   11
      Left            =   3105
      TabIndex        =   20
      Top             =   6330
      Width           =   450
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   12
      Left            =   6510
      TabIndex        =   19
      Top             =   855
      Width           =   450
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   0
      Left            =   3585
      Top             =   810
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   2
      Left            =   3585
      Top             =   1380
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   3
      Left            =   2835
      Top             =   1470
      Width           =   180
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   4
      Left            =   3585
      Top             =   1905
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   5
      Left            =   2835
      Top             =   1995
      Width           =   180
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   6
      Left            =   3585
      Top             =   2460
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   7
      Left            =   2835
      Top             =   2550
      Width           =   180
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   8
      Left            =   3585
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   9
      Left            =   2835
      Top             =   3090
      Width           =   180
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   10
      Left            =   3585
      Top             =   3555
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   11
      Left            =   2835
      Top             =   3645
      Width           =   180
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   12
      Left            =   3585
      Top             =   4065
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   13
      Left            =   2835
      Top             =   4170
      Width           =   180
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   14
      Left            =   3585
      Top             =   4635
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   15
      Left            =   2835
      Top             =   4725
      Width           =   180
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   16
      Left            =   3585
      Top             =   5175
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   17
      Left            =   2835
      Top             =   5265
      Width           =   180
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   18
      Left            =   3585
      Top             =   5730
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   19
      Left            =   2835
      Top             =   5820
      Width           =   180
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   20
      Left            =   3585
      Top             =   6270
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   21
      Left            =   2835
      Top             =   6360
      Width           =   180
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   22
      Left            =   7020
      Top             =   810
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   23
      Left            =   6270
      Top             =   900
      Width           =   180
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   24
      Left            =   7020
      Top             =   1380
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   25
      Left            =   6270
      Top             =   1470
      Width           =   180
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   13
      Left            =   6510
      TabIndex        =   18
      Top             =   1410
      Width           =   450
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   12
      Left            =   630
      TabIndex        =   17
      Top             =   4575
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   26
      Left            =   7020
      Top             =   1905
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   27
      Left            =   6270
      Top             =   1995
      Width           =   180
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   14
      Left            =   6510
      TabIndex        =   16
      Top             =   1950
      Width           =   450
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   13
      Left            =   630
      TabIndex        =   15
      Top             =   4920
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   28
      Left            =   7020
      Top             =   2460
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   29
      Left            =   6270
      Top             =   2550
      Width           =   180
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   15
      Left            =   6510
      TabIndex        =   14
      Top             =   2505
      Width           =   450
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   14
      Left            =   630
      TabIndex        =   13
      Top             =   5265
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   30
      Left            =   7020
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   31
      Left            =   6270
      Top             =   3090
      Width           =   180
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   16
      Left            =   6510
      TabIndex        =   12
      Top             =   3045
      Width           =   450
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   15
      Left            =   630
      TabIndex        =   11
      Top             =   5610
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   32
      Left            =   7020
      Top             =   3555
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   33
      Left            =   6270
      Top             =   3645
      Width           =   180
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   17
      Left            =   6510
      TabIndex        =   10
      Top             =   3615
      Width           =   450
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   16
      Left            =   630
      TabIndex        =   9
      Top             =   5955
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   34
      Left            =   7020
      Top             =   4080
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   35
      Left            =   6270
      Top             =   4170
      Width           =   180
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   18
      Left            =   6510
      TabIndex        =   8
      Top             =   4140
      Width           =   450
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   17
      Left            =   630
      TabIndex        =   7
      Top             =   6300
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   1
      Left            =   2835
      Top             =   900
      Width           =   180
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   18
      Left            =   630
      TabIndex        =   6
      Top             =   6645
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   19
      Left            =   6510
      TabIndex        =   5
      Top             =   4695
      Width           =   450
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   36
      Left            =   7020
      Top             =   4635
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   37
      Left            =   6270
      Top             =   4725
      Width           =   180
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   19
      Left            =   630
      TabIndex        =   4
      Top             =   6990
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   20
      Left            =   6510
      TabIndex        =   3
      Top             =   5235
      Width           =   450
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   38
      Left            =   7020
      Top             =   5175
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   39
      Left            =   6270
      Top             =   5265
      Width           =   180
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
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
      Height          =   240
      Index           =   20
      Left            =   630
      TabIndex        =   2
      Top             =   7350
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   21
      Left            =   6510
      TabIndex        =   1
      Top             =   5790
      Width           =   450
   End
   Begin VB.Image Command1 
      Height          =   300
      Index           =   40
      Left            =   7020
      Top             =   5730
      Width           =   270
   End
   Begin VB.Image Command1 
      Height          =   105
      Index           =   41
      Left            =   6270
      Top             =   5820
      Width           =   180
   End
   Begin VB.Label puntos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2790
      TabIndex        =   0
      Top             =   6990
      Width           =   450
   End
End
Attribute VB_Name = "frmSkills3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click(Index As Integer)

    Call Audio.PlayWave(SND_CLICK)

    Dim indice

    If Index Mod 2 = 0 Then
        If Alocados > 0 Then
            indice = Index \ 2 + 1

            If indice > NUMSKILLS Then indice = NUMSKILLS
            If Val(Text1(indice).Caption) < MAXSKILLPOINTS Then
                Text1(indice).Caption = Val(Text1(indice).Caption) + 1
                Flags(indice) = Flags(indice) + 1
                Alocados = Alocados - 1

            End If
            
        End If

    Else

        If Alocados < SkillPoints Then
        
            indice = Index \ 2 + 1

            If Val(Text1(indice).Caption) > 0 And Flags(indice) > 0 Then
                Text1(indice).Caption = Val(Text1(indice).Caption) - 1
                Flags(indice) = Flags(indice) - 1
                Alocados = Alocados + 1

            End If

        End If

    End If

    Puntos.Caption = Alocados
    
    If Puntos.Caption = 0 Then
        Puntos.ForeColor = &HFFFFFF
    Else
        Puntos.ForeColor = &H8000&

    End If

End Sub

Private Sub Command2_Click()

    Dim i   As Integer
    Dim cad As String
    
    For i = 1 To NUMSKILLS
        cad = cad & Flags(i) & ","
        'Actualizamos nuestros datos locales
        UserSkills(i) = Val(Text1(i).Caption)
    Next i
    
    SendData "SKSE" & cad

    If Alocados = 0 Then frmMain.imgSkillpts.Visible = False
    SkillPoints = Alocados
    Unload Me

End Sub

Private Sub Form_Activate()

    If SkillPoints = 0 Then
        Puntos.ForeColor = &HFFFFFF
    Else
        Puntos.ForeColor = &H8000&

    End If

End Sub

Private Sub Form_Load()

    'Valores máximos y mínimos para el ScrollBar
   
    'Nombres de los skills

    Dim L
    Dim i As Integer
    i = 1

    For Each L In Label2

        L.Caption = SkillsNames(i)
        L.AutoSize = True
        i = i + 1
    Next
    i = 0

    'Flags para saber que skills se modificaron
    ReDim Flags(1 To NUMSKILLS)
   
    'Alocados = SkillPoints
    
    Set Me.Picture = Interfaces.FrmSkill_Principal

End Sub

