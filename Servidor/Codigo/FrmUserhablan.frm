VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "COMCTL32.OCX"
Begin VB.Form FrmUserhablan 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9300
   FillColor       =   &H0000FF00&
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
   ScaleHeight     =   6015
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabHeight       =   520
      BackColor       =   16744576
      TabCaption(0)   =   "User Hablan"
      TabPicture(0)   =   "FrmUserhablan.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "User"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Hablan Clan"
      TabPicture(1)   =   "FrmUserhablan.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Clan"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Privado hablan"
      TabPicture(2)   =   "FrmUserhablan.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Privado"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Party Hablan"
      TabPicture(3)   =   "FrmUserhablan.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Party"
      Tab(3).ControlCount=   1
      Begin VB.ListBox Privado 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   4515
         Left            =   -74880
         TabIndex        =   6
         Top             =   720
         Width           =   8895
      End
      Begin VB.ListBox Party 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   4350
         Left            =   -74880
         TabIndex        =   5
         Top             =   840
         Width           =   8775
      End
      Begin VB.ListBox Clan 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   4350
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   8775
      End
      Begin VB.ListBox User 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   4350
         Left            =   -74880
         TabIndex        =   3
         Top             =   780
         Width           =   8895
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar Ventana"
      Height          =   360
      Left            =   480
      TabIndex        =   1
      Top             =   5520
      Width           =   1470
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9128
      MultiRow        =   -1  'True
      Style           =   1
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Usuarios Hablan"
            Object.Tag             =   "Tab 1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Hablan por Clan"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Hablan en Privado"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Party Hablan"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "FrmUserhablan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long

Private Const LB_SETHORIZONTALEXTENT = &H225

Private Sub AddHScroll(LB As ListBox)

    Dim i As Long
    Dim lLength As Long
    Dim lWdith As Long

    'see what the longest entry is
    For i = 0 To LB.ListCount - 1

        If Len(LB.List(i)) = Len(LB.List(lLength)) Then
            lLength = i

        End If

    Next i

    'add a little space
    lWdith = LB.Parent.TextWidth(LB.List(lLength) + Space(5))

    'Convert to Pixels
    lWdith = lWdith \ Screen.TwipsPerPixelX

    'Use api to add scrollbar
    Call SendMessage(LB.hWnd, LB_SETHORIZONTALEXTENT, lWdith, ByVal 0&)

End Sub

Private Sub Command1_Click()
    FrmUserhablan.Hide

End Sub

Public Sub hUser(ByVal texto As String)
    Dim CadenaOne() As String
    Dim CountLen As Long
    Dim LenMax As Long

    CadenaOne = Split(texto, Chr(62))
    User.AddItem (CadenaOne(0))

    CountLen = Len(CadenaOne(1))
    LenMax = 128

    If CountLen < 100 Then
        User.AddItem (mid(LCase(CadenaOne(1)), 1, 100))

    End If

    If CountLen > 100 And CountLen < 200 Then
        User.AddItem (mid(LCase(CadenaOne(1)), 1, "100"))
        User.AddItem (mid(LCase(CadenaOne(1)), 101, "100"))

    End If

    If CountLen > 200 Then
        User.AddItem (mid(LCase(CadenaOne(1)), 1, 100))
        User.AddItem (mid(LCase(CadenaOne(1)), 101, 100))
        User.AddItem (mid(LCase(CadenaOne(1)), 201, 55))

    End If

End Sub

Public Sub hClan(ByVal texto As String)

    Dim CadenaOne() As String

    Dim CountLen    As Long

    Dim LenMax      As Long

    CadenaOne = Split(texto, Chr(62))
    Clan.AddItem (CadenaOne(0))

    CountLen = Len(CadenaOne(1))
    LenMax = 128

    If CountLen < 100 Then

        Clan.AddItem (mid(LCase(CadenaOne(1)), 1, 100))

    End If

    If CountLen > 100 And CountLen < 200 Then

        Clan.AddItem (mid(LCase(CadenaOne(1)), 1, "100"))
        Clan.AddItem (mid(LCase(CadenaOne(1)), 101, "100"))

    End If

    If CountLen > 200 Then

        Clan.AddItem (mid(LCase(CadenaOne(1)), 1, 100))
        Clan.AddItem (mid(LCase(CadenaOne(1)), 101, 100))
        Clan.AddItem (mid(LCase(CadenaOne(1)), 201, 55))

    End If

End Sub

Public Sub hPrivado(ByVal texto As String)
    Dim CadenaOne() As String
    Dim CountLen As Long
    Dim LenMax As Long

    CadenaOne = Split(texto, Chr(62))
    Privado.AddItem (CadenaOne(0))

    CountLen = Len(CadenaOne(1))
    LenMax = 128

    If CountLen < 100 Then
        Privado.AddItem (mid(LCase(CadenaOne(1)), 1, 100))

    End If

    If CountLen > 100 And CountLen < 200 Then
        Privado.AddItem (mid(LCase(CadenaOne(1)), 1, "100"))
        Privado.AddItem (mid(LCase(CadenaOne(1)), 101, "100"))

    End If

    If CountLen > 200 Then
        Privado.AddItem (mid(LCase(CadenaOne(1)), 1, 100))
        Privado.AddItem (mid(LCase(CadenaOne(1)), 101, 100))
        Privado.AddItem (mid(LCase(CadenaOne(1)), 201, 55))

    End If

End Sub

Public Sub hParty(ByVal texto As String)
    Dim CadenaOne() As String
    Dim CountLen As Long
    Dim LenMax As Long

    CadenaOne = Split(texto, Chr(62))
    Party.AddItem (CadenaOne(0))

    CountLen = Len(CadenaOne(1))
    LenMax = 128

    If CountLen < 100 Then
        Party.AddItem (mid(LCase(CadenaOne(1)), 1, 100))

    End If

    If CountLen > 100 And CountLen < 200 Then
        Party.AddItem (mid(LCase(CadenaOne(1)), 1, "100"))
        Party.AddItem (mid(LCase(CadenaOne(1)), 101, "100"))

    End If

    If CountLen > 200 Then
        Party.AddItem (mid(LCase(CadenaOne(1)), 1, 100))
        Party.AddItem (mid(LCase(CadenaOne(1)), 101, 100))
        Party.AddItem (mid(LCase(CadenaOne(1)), 201, 55))

    End If

End Sub

