VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBaneos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Baneos en AoMania"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Tipos de Baneos en AoMania"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3630
      Left            =   300
      TabIndex        =   0
      Top             =   120
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   6403
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Banear"
      TabPicture(0)   =   "frmBaneos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ChameleonBtn3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Banear por Tiempo"
      TabPicture(1)   =   "frmBaneos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Ban desconectado"
      TabPicture(2)   =   "frmBaneos.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Command6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Ban de Ip"
      TabPicture(3)   =   "frmBaneos.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Image4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame4"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Command5"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Ban ide Unico"
      TabPicture(4)   =   "frmBaneos.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Image5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame5"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Command1"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Ban Clan"
      TabPicture(5)   =   "frmBaneos.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Command2"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame6"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Image6"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Castigos"
      TabPicture(6)   =   "frmBaneos.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Command3"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Frame7"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "logo"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Ban Formulario"
      TabPicture(7)   =   "frmBaneos.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Command4"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Frame8"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Image7"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).ControlCount=   3
      Begin AoManiaClienteGM.ChameleonBtn ChameleonBtn3 
         Height          =   315
         Left            =   7935
         TabIndex        =   37
         Top             =   3090
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
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
         FCOLO           =   65535
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmBaneos.frx":00E0
         PICN            =   "frmBaneos.frx":00FC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Salir"
         Height          =   360
         Left            =   -64800
         TabIndex        =   34
         Top             =   3120
         Width           =   990
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Salir"
         Height          =   360
         Left            =   -64560
         TabIndex        =   33
         Top             =   3120
         Width           =   990
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Salir"
         Height          =   360
         Left            =   -64680
         TabIndex        =   32
         Top             =   3120
         Width           =   990
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Salir"
         Height          =   360
         Left            =   -64680
         TabIndex        =   31
         Top             =   3120
         Width           =   990
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Salir"
         Height          =   360
         Left            =   -64680
         TabIndex        =   30
         Top             =   3120
         Width           =   990
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   360
         Left            =   -64440
         TabIndex        =   29
         Top             =   3120
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Salir"
         Height          =   360
         Left            =   -64560
         TabIndex        =   28
         Top             =   3120
         Width           =   990
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   1935
         Left            =   -74640
         TabIndex        =   12
         Top             =   720
         Width           =   4215
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo del Baneo:"
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   1080
            Width           =   1290
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banear Formulario:"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   1365
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   1935
         Left            =   -74760
         TabIndex        =   11
         Top             =   720
         Width           =   3735
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Días de Castigo:"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   960
            Width           =   1170
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   360
            TabIndex        =   23
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame6"
         Height          =   2055
         Left            =   -74640
         TabIndex        =   10
         Top             =   720
         Width           =   3855
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre del Clan:"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   1230
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   1815
         Left            =   -74640
         TabIndex        =   9
         Top             =   840
         Width           =   3495
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo del Baneo:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   1290
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   360
            TabIndex        =   20
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banear iD Unica:"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   480
            Width           =   1185
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   2055
         Left            =   -74640
         TabIndex        =   8
         Top             =   720
         Width           =   3735
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   480
            TabIndex        =   18
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   2055
         Left            =   -74640
         TabIndex        =   7
         Top             =   720
         Width           =   4095
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo del Baneo:"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   960
            Width           =   1290
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   360
            TabIndex        =   16
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2175
         Left            =   -74640
         TabIndex        =   6
         Top             =   720
         Width           =   4455
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dia de Baneo:"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo del Baneo"
            Height          =   195
            Left            =   360
            TabIndex        =   14
            Top             =   960
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            Height          =   195
            Left            =   360
            TabIndex        =   13
            Top             =   480
            Width           =   555
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2175
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   810
         Width           =   4455
         Begin AoManiaClienteGM.ChameleonBtn ChameleonBtn2 
            Height          =   465
            Left            =   3210
            TabIndex        =   36
            Top             =   1620
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   820
            BTYPE           =   3
            TX              =   "Banear"
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
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   0
            MICON           =   "frmBaneos.frx":1386
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin AoManiaClienteGM.ChameleonBtn ChameleonBtn1 
            Height          =   450
            Left            =   120
            TabIndex        =   35
            Top             =   1635
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   794
            BTYPE           =   3
            TX              =   "VerBaneos"
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
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   0
            MICON           =   "frmBaneos.frx":13A2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox MotivoBan 
            Height          =   285
            Left            =   1560
            TabIndex        =   5
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox NombreBan 
            Height          =   300
            Left            =   1560
            TabIndex        =   4
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo de Baneo:"
            Height          =   195
            Left            =   240
            TabIndex        =   3
            Top             =   840
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   840
            TabIndex        =   2
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Image Image7 
         Height          =   2445
         Left            =   -66000
         Picture         =   "frmBaneos.frx":13BE
         Top             =   600
         Width           =   2445
      End
      Begin VB.Image logo 
         Height          =   2445
         Left            =   -66000
         Picture         =   "frmBaneos.frx":DA9D
         Top             =   600
         Width           =   2445
      End
      Begin VB.Image Image6 
         Height          =   2445
         Left            =   -66000
         Picture         =   "frmBaneos.frx":1A17C
         Top             =   600
         Width           =   2445
      End
      Begin VB.Image Image5 
         Height          =   2445
         Left            =   -66000
         Picture         =   "frmBaneos.frx":2685B
         Top             =   600
         Width           =   2445
      End
      Begin VB.Image Image4 
         Height          =   2445
         Left            =   -66000
         Picture         =   "frmBaneos.frx":32F3A
         Top             =   600
         Width           =   2445
      End
      Begin VB.Image Image3 
         Height          =   2445
         Left            =   -66000
         Picture         =   "frmBaneos.frx":3F619
         Top             =   600
         Width           =   2445
      End
      Begin VB.Image Image2 
         Height          =   2445
         Left            =   -66000
         Picture         =   "frmBaneos.frx":4BCF8
         Top             =   600
         Width           =   2445
      End
      Begin VB.Image Image1 
         Height          =   2445
         Left            =   7155
         Picture         =   "frmBaneos.frx":583D7
         Top             =   585
         Width           =   2445
      End
   End
End
Attribute VB_Name = "frmBaneos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command8_Click()

End Sub
