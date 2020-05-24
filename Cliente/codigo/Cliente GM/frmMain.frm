VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{B370EF78-425C-11D1-9A28-004033CA9316}#2.0#0"; "Captura.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   11490
   ClientLeft      =   360
   ClientTop       =   -1110
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMain.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   766
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   240
      Top             =   240
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   16384
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   9999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer TimerMsj 
      Interval        =   100
      Left            =   3525
      Top             =   300
   End
   Begin VB.Timer TimerCarteles 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   2970
      Top             =   240
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   10680
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsScreen 
      Left            =   10440
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Particulas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Particulas"
      Height          =   495
      Left            =   9120
      MaskColor       =   &H000000FF&
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Clickeado 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   2070
      Top             =   240
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8520
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2.04000e5
      Width           =   555
   End
   Begin Captura.wndCaptura Foto 
      Left            =   360
      Top             =   840
      _ExtentX        =   688
      _ExtentY        =   688
   End
   Begin VB.Timer AntiFLOOD 
      Interval        =   1000
      Left            =   1620
      Top             =   240
   End
   Begin RichTextLib.RichTextBox SendCMSTXT 
      Height          =   225
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   397
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":1594
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picMiniMap 
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   6750
      ScaleHeight     =   100
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   255
      Width           =   1500
      Begin VB.Shape UserClanPos 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C00000&
         FillColor       =   &H00C00000&
         Height          =   75
         Index           =   9
         Left            =   990
         Shape           =   3  'Circle
         Top             =   180
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape UserClanPos 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C00000&
         FillColor       =   &H00C00000&
         Height          =   75
         Index           =   8
         Left            =   885
         Shape           =   3  'Circle
         Top             =   180
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape UserClanPos 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C00000&
         FillColor       =   &H00C00000&
         Height          =   75
         Index           =   7
         Left            =   780
         Shape           =   3  'Circle
         Top             =   180
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape UserClanPos 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C00000&
         Height          =   75
         Index           =   6
         Left            =   675
         Shape           =   3  'Circle
         Top             =   180
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape UserClanPos 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C00000&
         FillColor       =   &H00C00000&
         Height          =   75
         Index           =   5
         Left            =   570
         Shape           =   3  'Circle
         Top             =   180
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape UserClanPos 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C00000&
         FillColor       =   &H00C00000&
         Height          =   75
         Index           =   4
         Left            =   450
         Shape           =   3  'Circle
         Top             =   180
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape UserClanPos 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C00000&
         FillColor       =   &H00C00000&
         Height          =   75
         Index           =   3
         Left            =   345
         Shape           =   3  'Circle
         Top             =   180
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape UserClanPos 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C00000&
         FillColor       =   &H00C00000&
         Height          =   75
         Index           =   2
         Left            =   240
         Shape           =   3  'Circle
         Top             =   180
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape UserClanPos 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C00000&
         FillColor       =   &H00C00000&
         Height          =   75
         Index           =   1
         Left            =   135
         Shape           =   3  'Circle
         Top             =   180
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape UserClanPos 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C00000&
         FillColor       =   &H00C00000&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   10
         Left            =   1095
         Shape           =   3  'Circle
         Top             =   180
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape MiniUserPos 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   75
         Left            =   750
         Shape           =   3  'Circle
         Top             =   750
         Width           =   75
      End
   End
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9120
      Left            =   165
      ScaleHeight     =   608
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   736
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2190
      Width           =   11040
      Begin VB.PictureBox picScreen 
         AutoRedraw      =   -1  'True
         Height          =   735
         Left            =   7800
         ScaleHeight     =   675
         ScaleWidth      =   1755
         TabIndex        =   22
         Top             =   2520
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   12615
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   14
      Top             =   2940
      Width           =   2400
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2370
      Left            =   12660
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.CommandButton DespInv 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   9120
      MouseIcon       =   "frmMain.frx":1612
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   4920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   150
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   8115
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   720
      Top             =   240
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   1170
      Top             =   240
   End
   Begin RichTextLib.RichTextBox RecTxt 
      CausesValidation=   0   'False
      Height          =   1575
      Left            =   150
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   180
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox EnvioMsj 
      Height          =   225
      Left            =   120
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1800
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   397
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   0   'False
      HideSelection   =   0   'False
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":17E9
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
   Begin VB.Image Clima 
      Height          =   495
      Left            =   3195
      Top             =   8505
      Width           =   1950
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   0
      ImageWidth      =   32
      ImageHeight     =   32
      _Version        =   327682
   End
   Begin VB.Image Image8 
      Height          =   780
      Left            =   10530
      MouseIcon       =   "frmMain.frx":186D
      MousePointer    =   99  'Custom
      Top             =   7545
      Width           =   1110
   End
   Begin VB.Image Image7 
      Height          =   270
      Left            =   11400
      MouseIcon       =   "frmMain.frx":2537
      MousePointer    =   99  'Custom
      Top             =   8370
      Width           =   285
   End
   Begin VB.Image Image6 
      Height          =   270
      Left            =   11055
      MouseIcon       =   "frmMain.frx":3201
      MousePointer    =   99  'Custom
      Top             =   8370
      Width           =   285
   End
   Begin VB.Image Image5 
      Height          =   270
      Left            =   10710
      MouseIcon       =   "frmMain.frx":3ECB
      MousePointer    =   99  'Custom
      Top             =   8370
      Width           =   270
   End
   Begin VB.Image norte 
      Height          =   270
      Left            =   10365
      MouseIcon       =   "frmMain.frx":4B95
      MousePointer    =   99  'Custom
      Top             =   8370
      Width           =   270
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   10500
      MouseIcon       =   "frmMain.frx":585F
      MousePointer    =   99  'Custom
      Top             =   6945
      Width           =   1140
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   10680
      MouseIcon       =   "frmMain.frx":6529
      MousePointer    =   99  'Custom
      Top             =   6645
      Width           =   795
   End
   Begin VB.Image bOnline 
      Height          =   285
      Left            =   5190
      Top             =   8565
      Width           =   1215
   End
   Begin VB.Image bOnlineClan 
      Height          =   285
      Left            =   1650
      MouseIcon       =   "frmMain.frx":71F3
      MousePointer    =   99  'Custom
      Top             =   8565
      Width           =   1575
   End
   Begin VB.Image bGM 
      Height          =   285
      Left            =   90
      MouseIcon       =   "frmMain.frx":7EBD
      MousePointer    =   99  'Custom
      Top             =   8565
      Width           =   495
   End
   Begin VB.Image Donaciones 
      Height          =   270
      Left            =   8685
      MousePointer    =   99  'Custom
      Top             =   45
      Width           =   1680
   End
   Begin VB.Label exp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9015
      TabIndex        =   16
      Top             =   495
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblPorcLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9015
      TabIndex        =   15
      Top             =   495
      Width           =   2295
   End
   Begin VB.Label GldLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000000"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   315
      Left            =   12600
      TabIndex        =   12
      Top             =   5520
      Width           =   2355
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   10950
      TabIndex        =   11
      Top             =   1110
      Width           =   225
   End
   Begin VB.Image imgSkillpts 
      Height          =   405
      Left            =   11250
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":8B87
      Top             =   1350
      Width           =   405
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "Bassinger"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   8715
      TabIndex        =   9
      Top             =   1080
      Width           =   1665
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   375
      Index           =   1
      Left            =   11520
      MouseIcon       =   "frmMain.frx":BF66
      MousePointer    =   99  'Custom
      Top             =   3660
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   375
      Index           =   0
      Left            =   11520
      MouseIcon       =   "frmMain.frx":CC30
      MousePointer    =   99  'Custom
      Top             =   3240
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgExp 
      Height          =   225
      Left            =   9015
      Picture         =   "frmMain.frx":D8FA
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   570
      Left            =   12540
      MouseIcon       =   "frmMain.frx":11072
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   2130
      Width           =   990
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   570
      Left            =   14040
      MouseIcon       =   "frmMain.frx":11D3C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2160
      Width           =   990
   End
   Begin VB.Label lblVidaBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8730
      TabIndex        =   6
      Top             =   6390
      Width           =   1515
   End
   Begin VB.Label lblManaBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8730
      TabIndex        =   5
      Top             =   6975
      Width           =   1515
   End
   Begin VB.Label lblStaBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8730
      TabIndex        =   4
      Top             =   7515
      Width           =   1515
   End
   Begin VB.Label lblHamBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8805
      TabIndex        =   3
      Top             =   8040
      Width           =   1395
   End
   Begin VB.Label lblSedBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8805
      TabIndex        =   2
      Top             =   8535
      Width           =   1395
   End
   Begin VB.Image imgSed 
      Height          =   180
      Left            =   8805
      Picture         =   "frmMain.frx":12A06
      Stretch         =   -1  'True
      Top             =   8565
      Width           =   1395
   End
   Begin VB.Image imgComida 
      Height          =   180
      Left            =   8805
      Picture         =   "frmMain.frx":15C6C
      Stretch         =   -1  'True
      Top             =   8055
      Width           =   1395
   End
   Begin VB.Image imgEnergia 
      Height          =   150
      Left            =   8730
      Picture         =   "frmMain.frx":1A848
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1515
   End
   Begin VB.Image imgMana 
      Height          =   150
      Left            =   8730
      Picture         =   "frmMain.frx":1DCFB
      Stretch         =   -1  'True
      Top             =   7005
      Width           =   1515
   End
   Begin VB.Image imgVida 
      Height          =   150
      Left            =   8730
      Picture         =   "frmMain.frx":2112E
      Stretch         =   -1  'True
      Top             =   6420
      Width           =   1515
   End
   Begin VB.Image CmdLanzar 
      Height          =   315
      Left            =   12645
      MouseIcon       =   "frmMain.frx":245E3
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":252AD
      Top             =   5550
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image CmdInfo 
      Height          =   315
      Left            =   14310
      MouseIcon       =   "frmMain.frx":29593
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":2A25D
      Top             =   5520
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image InvEqu 
      Height          =   585
      Left            =   8850
      Picture         =   "frmMain.frx":2DC6D
      Top             =   2085
      Width           =   2595
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   0
      Left            =   10440
      MouseIcon       =   "frmMain.frx":3364F
      MousePointer    =   99  'Custom
      Top             =   7260
      Width           =   1290
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   1
      Left            =   6450
      MouseIcon       =   "frmMain.frx":34319
      MousePointer    =   99  'Custom
      Top             =   8565
      Width           =   1860
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   2
      Left            =   645
      MouseIcon       =   "frmMain.frx":34FE3
      MousePointer    =   99  'Custom
      Top             =   8565
      Width           =   945
   End
   Begin VB.Image Image3 
      Height          =   315
      Index           =   0
      Left            =   9120
      MouseIcon       =   "frmMain.frx":35CAD
      MousePointer    =   99  'Custom
      Top             =   5760
      Width           =   2355
   End
   Begin VB.Image Image13 
      Height          =   255
      Left            =   10500
      MouseIcon       =   "frmMain.frx":36977
      MousePointer    =   99  'Custom
      Top             =   6300
      Width           =   1155
   End
   Begin VB.Image Image11 
      Height          =   300
      Left            =   11655
      MousePointer    =   99  'Custom
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image12 
      Height          =   300
      Left            =   11310
      MousePointer    =   99  'Custom
      Top             =   15
      Width           =   300
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public UsandoDrag As Boolean
Public UsabaDrag  As Boolean
Public MoverHechizo As Boolean

Dim HechizoMove   As Integer

'Uso de Botones Key.
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public StatusCondi As Long

'Declaramos el Api GetAsyncKeyState
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'Quitar Bordes del ListBox
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, _
    ByVal hSrcRgn1 As Long, _
    ByVal hSrcRgn2 As Long, _
    ByVal nCombineMode As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'Transparencia ListBox
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As Any) As Long

Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long

Private Const Transparent        As Long = 1
Private Const WM_CTLCOLORLISTBOX As Long = &H134
Private Const WM_CTLCOLORSTATIC  As Long = &H138
Private Const WM_VSCROLL         As Long = &H115

'Consola transparente.
Const GWL_EXSTYLE = (-20)
Const WS_EX_TRANSPARENT = &H20&

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Transparencia ListBox
Dim WithEvents wndProc As clsTrickSubclass
Attribute wndProc.VB_VarHelpID = -1
Dim WithEvents lstProc As clsTrickSubclass
Attribute lstProc.VB_VarHelpID = -1

Dim hBackBrush         As Long

Private clsFormulario  As clsFormMovementManager

Public StatusESC       As Integer

' NUNCA OLVIDAR, TAMAÑO DE VISION 545 415

Public TX              As Byte
Public TY              As Byte

Public MouseX          As Long
Public MouseY          As Long
Public MouseBoton      As Long
Public MouseShift      As Long

Public SelM            As Integer
Public MapMapa         As Integer
Dim lBytes As Long, lFileSize As Long
Public flag As Boolean

Private Const LB_GETITEMHEIGHT = &H1A1
 
Private Sub AntiFLOOD_Timer()
    If FloodStats > 0 Then
        FloodStats = FloodStats - 1
    End If
End Sub

Private Sub bGM_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("/GM")
End Sub

Private Sub bOnline_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("/ONLINE")
End Sub

Private Sub bOnlineClan_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("/ONLINECLAN")
End Sub

Private Sub Clickeado_Timer()

If TiempoEst > 0 Then
TiempoEst = TiempoEst - 1
If TiempoEst = 0 Then
Estadisticas = False
Clickeado.Enabled = False
End If
End If
End Sub

Private Sub CmdLanzar_Click()
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub CmdLanzar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If hlst.List(hlst.ListIndex) <> "(Vacío)" Then
        Call SendData("VB" & hlst.ListIndex + 1)
        Call SendData("UK" & Magia)
    End If
End Sub

Private Sub Donaciones_Click()
    Call Audio.PlayWave(SND_CLICK)
    Shell "explorer " & "https://AoMania.net/Donaciones/", vbNormalFocus

End Sub

Private Sub Donaciones_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Donaciones.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Form_Activate()
    Call Audio.StopWave
    Call Audio.StopMidi

End Sub

Private Sub GldLbl_Click()
    
    Call Audio.PlayWave(SND_CLICK)

    Inventario.SelectGold

    If UserGLD > 0 Then frmCantidad.Show , frmMain

End Sub

Private Sub Image11_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("/Salir")
End Sub

Private Sub hlst_Click()
     
    If IsSeguroHechizos Then
        MoverHechizo = False
        Exit Sub
    End If
   
    If HechizoMove <> -1 And MoverHechizo = True Then
    
        Dim NewIndex   As Integer
        Dim NewHechizo As String
        
        'here I set the str one to the text to be moved
        NewHechizo = hlst.List(HechizoMove)
    
        'set the new index for str1 to be moved
        NewIndex = hlst.ListIndex
        
        If NewIndex < 0 Then Exit Sub 'subir
        If HechizoMove = NewIndex Then Exit Sub
        If NewIndex > hlst.ListCount Then Exit Sub  'bajar
        
        Call SendData("DESPHE" & HechizoMove + 1 & "," & NewIndex + 1)
        MoverHechizo = False
    End If
    
End Sub

Private Sub hlst_DblClick()
     
    If IsSeguroHechizos Then
       MoverHechizo = False
       Exit Sub
    End If
    
    If MoverHechizo = False Then
        HechizoMove = hlst.ListIndex
        MoverHechizo = True
    End If
     
End Sub

Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image11.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Image12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image12.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub Image13_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("/Mayor")
End Sub

Private Sub Image14_Click()

End Sub

Private Sub Image2_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("/Party")
End Sub

Private Sub Image4_Click()
    Call Audio.PlayWave(SND_CLICK)
    Shell "explorer " & "https://AoMania.net/Manual/", vbNormalFocus

End Sub

Private Sub Image5_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("/Castillo Sur")
End Sub

Private Sub Image6_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("/Castillo Oeste")
End Sub

Private Sub Image7_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("/Castillo Este")
End Sub

Private Sub Image8_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("/Fortaleza")
End Sub

Private Sub imgSkillpts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgSkillpts.MouseIcon = Iconos.Ico_Mano
End Sub

Private Sub lblPorcLvl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblPorcLvl.Caption = UserExp & "/" & UserPasarNivel
    
    If UserPasarNivel = 0 Then
        lblPorcLvl.Caption = "¡Nivel máximo!"

    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '    Macros.ClickRatonDown
    MouseBoton = Button
    MouseShift = Shift

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Else
    'Call AddtoRichTextBox(frmMain.RecTxt, "Mouse->No se permiten macros externos", 255, 255, 255, False, False, False)
    '  Exit Sub

    ' End If

    MouseBoton = Button
    MouseShift = Shift

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If prgRun = True Then
        prgRun = False
        Cancel = 1

    End If

End Sub

Private Sub Image12_Click()

    Call Audio.PlayWave(SND_CLICK)
    If WindowState <> vbMinimized Then WindowState = vbMinimized
    Visible = True

End Sub

Private Sub imgSkillpts_Click()

    Dim i As Integer
    
    Call Audio.PlayWave(SND_CLICK)
    
    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i

    Alocados = SkillPoints
    frmSkills3.puntos.Caption = SkillPoints
    frmSkills3.Show , frmMain

End Sub



Private Sub Label10_Click()

    Call Audio.PlayWave(SND_CLICK)
    SendData "/VERS"

End Sub


Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Macros.ClickRatonDown

End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call Audio.PlayWave(SND_CLICK)

    Set InvEqu.Picture = Interfaces.FrmMain_Hechizos
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    'DespInv(0).Visible = False
    'DespInv(1).Visible = False
    picInv.Visible = False
    GldLbl.Visible = False
    Image3(0).Visible = False
        
    hlst.Visible = True
    CmdInfo.Visible = True
    CmdLanzar.Visible = True
 
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True

End Sub

Private Sub Lemu_Click()

    Call Audio.PlayWave(SND_CLICK)
    Call SendData("/DEFLEMU")

End Sub

Private Sub MainViewPic_DblClick()

    Form_DblClick

End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If SendTxt.Visible Then SendTxt.SetFocus
    MouseBoton = Button
    MouseShift = Shift

End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseX = X
    MouseY = Y
    
    'Get new target positions
    ConvertCPtoTP MouseX, MouseY, TX, TY

    If InMapBounds(TX, TY) Then

        With MapData(TX, TY)

            If UsandoDrag = False Then   ' Utiliza Drag
                frmMain.picInv.MousePointer = vbDefault
            Else

                'Drag de items a posiciones. [maTih.-]
                Dim selInvSlot As Byte

                'Get the selected slot of the inventory.
                selInvSlot = Inventario.SelectedItem

                'Not selected item?
                If Not selInvSlot <> 0 Then Exit Sub

                'There is invalid position?.
                If .Blocked <> 0 Then
                    Call AddtoRichTextBox(frmMain.RecTxt, "Posición inválida", 255, 255, 255, True, , True)
                    Call StopDragInv
                    Exit Sub

                End If

                ' Not Drop on ilegal position; Standelf
                Dim IS_VALID_POS As Boolean

                IS_VALID_POS = LegalPos(TX + 1, TY) = False And LegalPos(TX - 1, TY) = False And LegalPos(TX, TY - 1) = False And LegalPos(TX, TY + _
                    1) = False

                If IS_VALID_POS Then
           
                    Call AddtoRichTextBox(frmMain.RecTxt, "La posición donde desea tirar el ítem es ilegal.", 255, 255, 255, True, , True)
                    Call StopDragInv
                    Exit Sub

                End If

                'There is already an object in that position?.
                If Not .charindex <> 0 Then
                    If .ObjGrh.GrhIndex <> 0 Then
                        
                        Call AddtoRichTextBox(frmMain.RecTxt, "Hay un objeto en esa posición!", 255, 255, 255, True, , True)
                
                        Call StopDragInv
                        Exit Sub

                    End If

                End If
                
                Dim Amount As Integer
                Amount = 1

                If Shift = 1 Then
                    Amount = Val(InputBox("Ingresar la cantidad a tirar."))

                    Do While Amount < 0 And Not IsNumeric(Amount)
                        Amount = Val(InputBox("Ingresar la cantidad a tirar."))
                    
                    Loop

                End If

                'Send the package.
                Call SendData("DRO" & selInvSlot & "," & TX & "," & TY & "," & Amount)

                'Reset the flag.
                Call StopDragInv

            End If

        End With

    End If

End Sub

Private Sub StopDragInv()

    ' GSZAO
    UsabaDrag = False
    UsandoDrag = False

    frmMain.picInv.MousePointer = vbDefault
    
End Sub

Private Sub MainViewPic_Click()

    If Cartel Then Cartel = False

    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, TX, TY)

        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                If UsingSkill = 0 Then
                    SendData "LC" & TX & "," & TY
                    If MapData(TX, TY).charindex > 0 Then
                        UserClicado = MapData(TX, TY).charindex
                    End If
                    frmMain.MousePointer = vbCustom
                    Set frmMain.MouseIcon = Iconos.Ico_Diablo
                Else
                    frmMain.MousePointer = vbCustom
                    Set frmMain.MouseIcon = Iconos.Ico_Diablo
                    
                    If (UsingSkill = Magia Or UsingSkill = Proyectiles) And UserCanAttack = 0 Then Exit Sub
                    SendData "WLC" & TX & "," & TY & "," & UsingSkill

                    If UsingSkill = Magia Or UsingSkill = Proyectiles Then UserCanAttack = 0
                    UsingSkill = 0
                    
                    If AoSetup.bCarteles Then
                        
                        TimerCarteles.Enabled = True
                        
                    End If
                    
                End If

            End If

        ElseIf (MouseShift And 1) = 1 Then

            If MouseShift = vbLeftButton Then
                Call SendData("/TELEP YO " & UserMap & " " & TX & " " & TY)

            End If

        End If

    End If
    
End Sub

Private Sub mnuEquipar_Click()
    
    Call EquiparItem

End Sub

Private Sub mnuNPCComerciar_Click()

    SendData "LC" & TX & "," & TY
    SendData "/COMERCIAR"

End Sub

Private Sub mnuNpcDesc_Click()

    SendData "LC" & TX & "," & TY

End Sub

Private Sub mnuTirar_Click()

    Call TirarItem

End Sub

Private Sub mnuUsar_Click()

    Call UsarItem

End Sub

Private Sub Nix_Click()

    Call Audio.PlayWave(SND_CLICK)
    Call SendData("/DEFNIX")

End Sub

Private Sub norte_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("/Castillo Norte")
End Sub

Private Sub Particulas_Click()

    Call frmCrearParticulas.Show(vbModeless, frmMain)

End Sub

Private Sub picInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
       
        Dim Last_I As Long
        Last_I = Inventario.SelectedItem

        If Last_I > 0 And Last_I <= MAX_INVENTORY_SLOTS Then
            Dim GrhIndex As Long

            GrhIndex = Inventario.GrhIndex(Last_I)
       
            If GrhIndex > 0 Then
                Dim Poss As Long
                
                Poss = BuscarI(GrhIndex)
                    
                If Poss = 0 Then
             
                    'Dim Data()  As Byte
                    'Dim handle  As Integer
                    'Dim BmpData As StdPicture

                    'If Get_Image(DirGraficos, CStr(GrhData(GrhIndex).FileNum), Data, True) Then
                    '   Set BmpData = ArrayToPicture(Data(), 0, UBound(Data) + 1) ' GSZAO
                    '   frmMain.ImageList1.ListImages.Add , CStr("g" & GrhIndex), Picture:=BmpData
                    '   Poss = frmMain.ImageList1.ListImages.Count
                    '   Set BmpData = Nothing
                    '
                    'End If
                    Dim DR As RECT

                    DR.Left = 0
                    DR.Top = 0
                    DR.Right = 32
                    DR.Bottom = 32
                    
                    ' Esto tiene que ir si o si
                    Call DrawGrhtoHdc(Me.Picture5.hdc, GrhIndex, DR)
               
                    'Set Me.picMiniMap.Picture = BmpData
                    frmMain.ImageList1.ListImages.Add , CStr("g" & GrhIndex), Picture:=Picture5.Image
                    Poss = frmMain.ImageList1.ListImages.Count
                 
                End If
                    
                UsandoDrag = True

                If frmMain.ImageList1.ListImages.Count > 0 Then
                    Set picInv.MouseIcon = frmMain.ImageList1.ListImages(Poss).ExtractIcon

                End If

                frmMain.picInv.MousePointer = vbCustom
                Exit Sub

            End If

        End If

    End If

End Sub

Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
                             
    If Not UsandoDrag Then picInv.MousePointer = vbDefault

End Sub

Private Sub picMiniMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

          If Button = 2 Then
             Call SendData("/TELEP YO " & UserMap & " " & X & " " & Y)
          End If


End Sub

Private Sub Second_Timer()

    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer

End Sub

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()

    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then

        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            SendData "OH" & Inventario.SelectedItem & "," & 1
        Else

            If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                frmCantidad.Show , frmMain

            End If

        End If

    End If

End Sub

Private Sub AgarrarItem()

    SendData "AG"

End Sub

Private Sub UsarItem()

    SendData "HDP"

    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then SendData "USA" & Inventario.SelectedItem

End Sub

Private Sub EquiparItem()

    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then SendData "EQUI" & Inventario.SelectedItem
        
End Sub

Private Sub CmdInfo_Click()
    
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("INFS" & hlst.ListIndex + 1)

End Sub

''''''''''''''''''''''''''''''''''''''
'     OTROS                          '
''''''''''''''''''''''''''''''''''''''

Private Sub Form_DblClick()

    If Not frmForo.Visible Then
        SendData "RC" & TX & "," & TY
        Call SendData("/MOV")

    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
       
        '[CUSTOM KEYS]
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then

            Select Case KeyCode

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)

                    If Not Audio.PlayingMusic Then
                        Audio.MusicActivated = True
                        Audio.PlayMIDI CStr(currentMidi) & ".mid"
                    Else
                        Audio.MusicActivated = False
                        Audio.StopMidi

                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem

                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres

                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    Call SendData("UK" & Domar)

                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    Call SendData("UK" & Robar)

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    Call SendData("/SEG")

                Case CustomKeys.BindedKey(eKeyType.mKeySafeGuild)
                    Call SendData("/SEGCLAN")

                Case CustomKeys.BindedKey(eKeyType.mKeyCombate)
                    Call SendData("/SEGCMBT")

                Case CustomKeys.BindedKey(eKeyType.mKeyObjetos)
                    Call SendData("/SEGOBJT")

                Case CustomKeys.BindedKey(eKeyType.mKeyHechizos)
                    Call SendData("/SEGHZS")

                Case vbKeyZ:
                    frmMain.RecTxt.Text = vbNullString

                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    Call SendData("UK" & Ocultarse)

                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem

                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)

                    If Not NoPuedeUsar Then
                        NoPuedeUsar = True
                        Call UsarItem

                    End If
        
                Case vbKeyP:

                    If Not NoPuedeUsar Then
                        NoPuedeUsar = True
                        Call UsarItem

                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)

                    If UserPuedeRefrescar Then
                        Call SendData("RPU")
                        UserPuedeRefrescar = False
                        Beep
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySeguroCvc)
                     'para Seth: Aquí lo que hace el COMANDO!

            End Select

        End If

    End If
        
    Select Case KeyCode
               
        Case CustomKeys.BindedKey(eKeyType.mKeyOnline)
            Call SendData("/ONLINE")

        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
            Call SendData("/SALIR")

        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)

            If SendTxt.Visible Then Exit Sub
            If (Not UserDescansar) And (Not UserMeditar) Then
                SendData "KC"

            End If

        Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)

            If Timer1.Enabled = True Then
                Call AddtoRichTextBox(frmMain.RecTxt, "Has desactivado el macro", 255, 0, 0, False, False, False)
                Timer1.Enabled = False
                Exit Sub
            End If
      
            If Not Inventario.ObjType(Inventario.SelectedItem) = 18 Then
                Call AddtoRichTextBox(frmMain.RecTxt, "Debes tener seleccionada la herramienta.", 255, 0, 0, False, False, False)
                Exit Sub
            End If

            If Inventario.Equipped(Inventario.SelectedItem) = False Then
                Call AddtoRichTextBox(frmMain.RecTxt, "Equipate antes la herramienta.", 255, 0, 0, False, False, False)
                Exit Sub

            End If

            If Timer1.Enabled = False Then
                Timer1.Enabled = True
                Call AddtoRichTextBox(frmMain.RecTxt, "Has activado el macro", 255, 0, 0, False, False, False)
            Else
                Timer1.Enabled = False
                Call AddtoRichTextBox(frmMain.RecTxt, "Has desactivado el macro", 255, 0, 0, False, False, False)
            End If

        Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            Call frmOpciones.Show(, frmMain)

        Case CustomKeys.BindedKey(eKeyType.mKeyToggleMAPA)
            If frmMain.SendTxt.Visible Then Exit Sub
            
            If VerMapa Then
                 VerMapa = False
                Exit Sub
            End If
            
            If Not VerMapa Then
                VerMapa = True
                Exit Sub
            End If

        Case CustomKeys.BindedKey(eKeyType.mKeyBank)
            Call SendData("/BOVEDA")

        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
            Call SendData("/MEDITAR")

        Case CustomKeys.BindedKey(eKeyType.mKeyTrade)
            Call SendData("/COMERCIAR")

        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
            Dim X As Long
            Foto.Area = Ventana
            Foto.Captura

            For X = 1 To 1000

                If Not FileExist(DirFotos & X & ".jpg", vbNormal) Then Exit For
            Next
            Call Convertir(Foto.Imagen, DirFotos & X & ".jpg")
            Call AddtoRichTextBox(frmMain.RecTxt, "Foto guardada en " & DirFotos & X & ".jpg", 255, 128, 69, False, False, False)
                         
    End Select

    Select Case KeyCode

        Case CustomKeys.BindedKey(eKeyType.mKeyTalkGuild)

            If SendTxt.Visible Then Exit Sub
            
            If Not Comerciando And Not frmCantidad.Visible Then
                SendCMSTXT.Visible = True
                SendCMSTXT.SetFocus

            End If
            
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)

            If SendCMSTXT.Visible Then Exit Sub
            
            If Not Comerciando And Not frmCantidad.Visible Then
                SendTxt.Visible = True
                SendTxt.SetFocus

            End If

    End Select
 
End Sub

Public Sub InitDrawMain(ByVal Drag As Boolean)

    DragPantalla = Drag

    If DragPantalla Then
    
        If NoRes Then
            ' Handles Form movement (drag and drop).
            Set clsFormulario = New clsFormMovementManager
            clsFormulario.Initialize Me, 120

        End If

    Else
   
        If Not clsFormulario Is Nothing Then
      
            Set clsFormulario = Nothing

        End If

    End If
    
End Sub

Private Sub Form_Load()
 
    'Consola transparente
    Dim result As Long
   
    result = SetWindowLong(RecTxt.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    result = SetWindowLong(SendCMSTXT.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    result = SetWindowLong(EnvioMsj.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    
    Set InvEqu.Picture = Interfaces.FrmMain_Inventario
    Set frmMain.Picture = Interfaces.FrmMain_Principal
     
    'ListBox
    hBackBrush = CreatePatternBrush(Me.Picture.handle)
    
    Set wndProc = New clsTrickSubclass
    Set lstProc = New clsTrickSubclass
    
    wndProc.Hook Me.hwnd
    lstProc.Hook hlst.hwnd
   
    'Esta funcion no es necesaria usarla.
    ' Do While hlst.ListCount < 35
    '     hlst.AddItem Format(hlst.ListCount, "ITE\M 00")
    ' Loop
     
    'Borde ListBox
    Dim rgn1 As Long
    rgn1 = CreateRectRgn(1, 1, hlst.Width - 1, hlst.Height - 1)
    SetWindowRgn hlst.hwnd, rgn1, True
     
    Detectar RecTxt.hwnd, Me.hwnd
    
    InitDrawMain DragPantalla

    SendTxt.Visible = False
    SendCMSTXT.Visible = False
    lblUserName.Caption = UserName
    LvlLbl.Caption = UserLvl
    Call ForeColorToNivel(CByte(UserLvl))
       
  
    wsScreen.RemoteHost = CurServerIp
    'wsScreen.RemoteHostIP
    Me.Left = 0
    Me.Top = 0
    Me.Height = 11520
    Me.Width = 15360
           
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseX = X
    MouseY = Y
    
    If UserPasarNivel = 0 Then
        lblPorcLvl.Caption = "¡Nivel máximo!"
    Else

        If UserExp <> 0 And UserPasarNivel <> 0 Then
            frmMain.lblPorcLvl.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"

        End If

    End If

End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyCode = 0

End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)

    KeyAscii = 0

End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyCode = 0

End Sub

Private Sub Image1_Click(Index As Integer)
   
    Call Audio.PlayWave(SND_CLICK)

    Select Case Index

        Case 0
            '[MatuX] : 01 de Abril del 2002
            Call frmOpciones.Show(vbModeless, frmMain)
             
            '[END]
        Case 1
            
            If FloodStats > 0 Then
                Call AddtoRichTextBox(frmMain.RecTxt, "AVISO: Debes esperar 15 segundos entre cada petición de estadísticas..", 68, 147, 66, 0, 0)
                Exit Sub
            End If
            
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            SendData "ATRI"
            SendData "ESKI"
            SendData "FEST"
            SendData "FAMA"

            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            FloodStats = 15

        Case 2

            If Not frmGuildLeader.Visible Then Call SendData("GLINFO")

    End Select

End Sub

Private Sub Image3_Click(Index As Integer)

    Select Case Index

        Case 0
            Inventario.SelectGold

            If UserGLD > 0 Then
                Call FrmTransferir.Show(vbModeless, frmMain)

            End If

    End Select

End Sub

Private Sub Label4_Click()

    Call Audio.PlayWave(SND_CLICK)

    Set InvEqu.Picture = Interfaces.FrmMain_Inventario

    'DespInv(0).Visible = True
    'DespInv(1).Visible = True
    picInv.Visible = True
    GldLbl.Visible = True
    Image3(0).Visible = True
    
    hlst.Visible = False
    CmdInfo.Visible = False
    CmdLanzar.Visible = False
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True

End Sub

Private Sub picInv_DblClick()

    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
        Dim ObjType As Integer
        ObjType = Inventario.ObjType(Inventario.SelectedItem)

        If ObjType = eObjType.otAlas Or ObjType = eObjType.otArmadura Or ObjType = eObjType.otCASCO Or ObjType = eObjType.otESCUDO Or ObjType = _
            eObjType.otWeapon Then
            Call EquiparItem
        Else
  
            Call UsarItem

        End If

    End If

End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call Audio.PlayWave(SND_CLICK)
   
End Sub

Private Sub RecTxt_Change()
        
    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar

    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf Me.SendCMSTXT.Visible Then
        SendCMSTXT.SetFocus
    ElseIf (Not frmComerciar.Visible) And (Not frmSkills3.Visible) And (Not frmForo.Visible) And (Not frmEstadisticas.Visible) And (Not _
        frmCantidad.Visible) Then

        If picInv.Visible Then
            picInv.SetFocus
        ElseIf hlst.Visible Then
            hlst.SetFocus

        End If

    End If

End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Sin comentar lo que hace perder el foco del Rectxt, no funciona nada xd
        If Not (KeyCode = vbKeyControl Or KeyCode = vbKeyC) Then  'KeyCode = 0  'copy (ctrl + c) Then
    If picInv.Visible Then
    picInv.SetFocus
    Else
    hlst.SetFocus

    End If
        End If

End Sub

Private Sub SendTxt_Change()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 3/06/2006
    '3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
    '**************************************************************

    'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
    Dim i         As Long
    Dim tempstr   As String
    Dim CharAscii As Integer
        
    For i = 1 To Len(SendTxt.Text)
        CharAscii = Asc(mid$(SendTxt.Text, i, 1))

        If CharAscii >= vbKeySpace And CharAscii <= 250 Then
            tempstr = tempstr & Chr$(CharAscii)

        End If

    Next i
        
    If tempstr <> SendTxt.Text Then
        SendTxt.Text = tempstr

    End If
        
    stxtbuffer = SendTxt.Text

End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0

End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

    'Send text
    If KeyCode = vbKeyReturn Then
        If Left$(stxtbuffer, 1) = "/" Then
            If UCase$(Left$(stxtbuffer, 8)) = "/PASSWD " Then
                stxtbuffer = "/PASSWD " & Right$(stxtbuffer, Len(stxtbuffer) - 8)
            ElseIf UCase$(stxtbuffer) = "/FUNDARCLAN" Then
                stxtbuffer = "/FUNDARCLAN NEUTRO"

            End If

            Call SendData(stxtbuffer)
    
            'Shout
        ElseIf Left$(stxtbuffer, 1) = "-" Then
            Call SendData("-" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

            'Whisper
        ElseIf Left$(stxtbuffer, 1) = "\" Then
            Call SendData("\" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

            'Say
        ElseIf Len(stxtbuffer) <> 0 Then
            Call SendData(";" & stxtbuffer)

        End If

        stxtbuffer = vbNullString
        KeyCode = 0
        SendTxt.Text = vbNullString
        SendTxt.Visible = False
        
        If picInv.Visible Then
            picInv.SetFocus
        Else
            hlst.SetFocus

        End If

    End If

End Sub

Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)

    'Send text
    If KeyCode = vbKeyReturn Then

        'Say
        If Len(stxtbuffercmsg) <> 0 Then
            Call SendData("/CMSG " & stxtbuffercmsg)
            
        End If

        stxtbuffercmsg = vbNullString
        KeyCode = 0
        SendCMSTXT.Text = vbNullString
        SendCMSTXT.Visible = False
        
        If picInv.Visible Then
            picInv.SetFocus
        Else
            hlst.SetFocus

        End If

    End If

End Sub

Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0

End Sub

Private Sub SendCMSTXT_Change()

    'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
    Dim i         As Long

    Dim tempstr   As String

    Dim CharAscii As Integer
        
    For i = 1 To Len(SendCMSTXT.Text)
        CharAscii = Asc(mid$(SendCMSTXT.Text, i, 1))

        If CharAscii >= vbKeySpace And CharAscii <= 250 Then
            tempstr = tempstr & Chr$(CharAscii)

        End If

    Next i
        
    If tempstr <> SendCMSTXT.Text Then
        'We only set it if it's different, otherwise the event will be raised
        'constantly and the client will crush
        SendCMSTXT.Text = tempstr

    End If

    stxtbuffercmsg = SendCMSTXT.Text

End Sub

''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''

Private Sub Socket1_Connect()

    Second.Enabled = True

    'If frmCrearPersonaje.Visible Then
    If EstadoLogin = E_MODO.CrearNuevoPj Then
        Call login
            
    ElseIf EstadoLogin = E_MODO.Normal Then
        Call login
   
    ElseIf EstadoLogin = E_MODO.Dados Then
        Call login
  
    ElseIf EstadoLogin = E_MODO.RecuperarPass Then
        Dim cmd As String
        cmd = "PASSRECO" & frmRecuperar.txtNombre.Text & "~" & frmRecuperar.txtCorreo
        frmMain.Socket1.Write cmd, Len(cmd)

    End If

End Sub

Private Sub Socket1_Disconnect()

    Dim i As Long
    
    Second.Enabled = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    
    If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next

    For i = 1 To Forms.Count

        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name Then
            Unload Forms(i)

        End If

    Next i

    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = vbNullString
    UserSexo = vbNullString
    UserRaza = vbNullString
    UserEmail = vbNullString
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.RemoveAllDialogs
    Inventario.ClearAllSlots
    
    AoDefResult = 0

End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)

    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    If ErrorCode = 24060 Or ErrorCode = 24061 Then
        Call MsgBox("No se puede conectar. Puede que el servidor esté OFF o no tengas conexión a internet.", vbApplicationModal + vbInformation + _
        vbOKOnly + vbDefaultButton1, "Error")
        Response = 0
        frmMain.Socket1.Disconnect
        Exit Sub
    End If
    
    frmConnect.MousePointer = 1
    Response = 0
    
     'Second.Enabled = False

    frmMain.Socket1.Disconnect
    AoDefResult = 0
    
    If Not frmCrearPersonaje.Visible Then
        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
        End If

    Else
        frmCrearPersonaje.MousePointer = 0

    End If

End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    
    On Error Resume Next
    
    Dim LooPC             As Long
    Dim RD                As String
    Dim rBuffer(1 To 500) As String

    Static TempString     As String

    Dim CR                As Integer
    Dim tChar             As String
    Dim sChar             As Integer
    Dim Echar             As Integer
    Dim Lenght            As Integer
    
    Call Socket1.Read(RD, DataLength)
    
    'Check for previous broken data and add to current data
    If Len(TempString) <> 0 Then
        RD = TempString & RD
        TempString = vbNullString

    End If

    'Check for more than one line
    sChar = 1

    Lenght = Len(RD)

    For LooPC = 1 To Lenght

        tChar = mid$(RD, LooPC, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = LooPC - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = LooPC + 1

        End If

    Next LooPC

    'Check for broken line and save for next time
    If Lenght - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Lenght)

    End If

    'Send buffer to Handle data
    For LooPC = 1 To CR
        Call HandleData(rBuffer(LooPC))
    Next LooPC

End Sub

Private Sub Tale_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("/DEFTALE")
End Sub



Private Sub Timer1_Timer()

    If Inventario.ObjType(Inventario.SelectedItem) = 18 Then
        
       If frmMain.hwnd <> GetActiveWindow And (GetForegroundWindow <> frmMain.hwnd) Then
         Call AddtoRichTextBox(frmMain.RecTxt, "Has desactivado el macro", 255, 0, 0, False, False, False)
         Timer1.Enabled = False
       End If
       
        Call UsarItem

        'Form_click
        If Cartel Then Cartel = False

        If Not Comerciando Then
            Call ConvertCPtoTP(MouseX, MouseY, TX, TY)

            If MouseShift = 0 Then
                If MouseBoton <> vbRightButton Then
                    If UsingSkill = 0 Then
                        SendData "LC" & TX & "," & TY
                    Else
                        frmMain.MousePointer = vbCustom

                        If (UsingSkill = Magia Or UsingSkill = Proyectiles) And UserCanAttack = 0 Then Exit Sub
                        SendData "WLC" & TX & "," & TY & "," & UsingSkill

                        If UsingSkill = Magia Or UsingSkill = Proyectiles Then UserCanAttack = 0
                        UsingSkill = 0

                    End If

                End If

            ElseIf (MouseShift And 1) = 1 Then

                If MouseShift = vbLeftButton Then
                    Call SendData("/TELEP YO " & UserMap & " " & TX & " " & TY)
                End If

            End If

        End If

    Else
        Call AddtoRichTextBox(frmMain.RecTxt, "Debes equiparte y seleccionar la herramienta!", 255, 255, 255, False, False, False)

    End If

End Sub

Private Sub Timer2_Timer()

    If VidaVerde > 0 And StatusVerde Then
        StatusVerde = False
    Else
        StatusVerde = True

    End If
      
    If VidaAmarilla > 0 And StatusAmarilla Then
        StatusAmarilla = False
    Else
        StatusAmarilla = True

    End If

End Sub

Private Sub Ulla_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("/DEFULLA")
End Sub

Private Sub lstProc_wndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ret As Long, DefCall As Boolean)
    
    Select Case msg
        
        Case WM_VSCROLL
    
            InvalidateRect hwnd, ByVal 0&, 0

    End Select
   
    DefCall = True
    
End Sub

Private Sub TimerCarteles_Timer()
      
          Call SendData(";" & " ")
          TimerCarteles = False
      
End Sub

Private Sub TimerMsj_Timer()
        CountMEC = CountMEC - 1
        
        MensajeEnvio = mid(MensajeEnvio, 2) & Left(MensajeEnvio, 1)
        EnvioMsj = MensajeEnvio
        
        If CountMEC = 0 Then TimerMsj.Enabled = False
End Sub

Private Sub wndProc_WndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ret As Long, DefCall As Boolean)
    
    Select Case msg

        Case WM_CTLCOLORSTATIC, WM_CTLCOLORLISTBOX
            Dim pts(1) As Long
    
            MapWindowPoints lParam, Me.hwnd, pts(0), 1
     
            SetBrushOrgEx wParam, -pts(0), -pts(1), ByVal 0&
 
            If lParam = hlst.hwnd Then
                SetBkMode wParam, Transparent
                SetTextColor wParam, vbWhite
        
            End If

            ret = hBackBrush
        
        Case Else
            DefCall = True

    End Select
    
End Sub

Private Function BuscarI(ByVal Gh As Long) As Long

    Dim i As Long

    For i = 1 To frmMain.ImageList1.ListImages.Count

        If frmMain.ImageList1.ListImages(i).key = "g" & CStr(Gh) Then
            BuscarI = i
            Exit For

        End If

    Next i

End Function

Private Sub wsScreen_Connect()
    wsScreen.SendData "|Archivo|" & FileLen(capturaPath)
End Sub

Private Sub wsScreen_DataArrival(ByVal bytesTotal As Long)
    Dim bData As String
    
    'Winsock1.GetData bData, vbString
    'If bData = "|Okkkkkkkkkkkk|" Then
    Call Enviar_Archivo
    
End Sub

Private Sub wsScreen_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Debug.Print "eRROR AL CONECTAR EL WSSCREEN" & Number & Description
End Sub


Private Sub Enviar_Archivo()
    Dim Size As Long
    Dim arrData() As Byte
      
    Open capturaPath For Binary Access Read As #1
      
    'Obtenemos el tamaño exacto en bytes del archivo para
    ' poder redimensionar el array de bytes
    Size = LOF(1)
    ReDim arrData(Size - 1)
      
    'Leemos y almacenamos todo el fichero en el array
    Get #1, , arrData
    'Cerramos
    Close
    
    Kill capturaPath
    'Enviamos el archivo
    wsScreen.SendData arrData
  
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error GoTo errh
      Dim arrData()   As Byte
      Dim vData       As Variant
    
1      If flag = False Then
2          Winsock1.GetData vData, vbString
3          If mid(vData, 1, 9) = "|Archivo|" Then
4              flag = True
5              lBytes = 0
6              vData = Split(vData, "|")
7              lFileSize = vData(2)
8              ' Le enviamos como mensaje al cliente que comienze el envio del archivo
                Winsock1.SendData "Ok"
9              'Creamos un archivo en modo binario
10              Open App.Path & "\tsnap.bmp" For Binary Access Write As #1
11          End If
12      End If
    
13      If flag Then
14          ' Aumentamos lBytes con los datos que van llegando
15          lBytes = lBytes + bytesTotal
16          'Recibimos los datos y lo almacenamos en el arry de bytes
17          Winsock1.GetData arrData
    
          'Escribimos en disco el array de bytes, es decir lo que va llegando
18          Put #1, , arrData
    
          ' Si lo recibido es mayor o igual al tamaño entonces se terminó y cerramos
          'el archivo abierto
19          If lBytes >= lFileSize Then
              'Cerramos el archivo
20              Close #1
              'Reestablecemos el flag y la variable lBytes por si se intenta enviar otro archivo
21              flag = False
22              lBytes = 0
23              frmCapturas.picScreen.Picture = LoadPicture(App.Path & "\tsnap.bmp")
                frmCapturas.Width = (frmCapturas.picScreen.Width * 15) + 200
                frmCapturas.Height = (frmCapturas.picScreen.Height * 15) + 500
24              frmCapturas.Show , frmMain
                Kill App.Path & "\tsnap.bmp"
25              Winsock1.Close
              'MsgBox "El archivo se ha recibido por completo"
          End If
      End If
Exit Sub
errh:
    Debug.Print "Error en winsock1 dataarrival " & Err.Number & Err.Description & Erl

    
    
End Sub
