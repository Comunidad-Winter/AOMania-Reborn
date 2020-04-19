VERSION 5.00
Begin VB.Form ViewHDD 
   BackColor       =   &H00FFFF00&
   Caption         =   "MD5Txt"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ViewHDD.frx":0000
   LinkTopic       =   "ViewHDD"
   ScaleHeight     =   2385
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H0000FFFF&
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   6375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Leer"
      Height          =   360
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   360
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   990
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "ViewHDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   frmMain.Visible = True
   Unload Me
End Sub

Private Sub Command2_Click()
   Text1.Text = fdpc
   Text1.Text = MD5String(Text1.Text)
   Text2.Text = "1 - Editar el Fichero gmsmac.dat" & vbCrLf & _
                            "2 - Introducir el dato en MAC=" & Text1.Text & vbCrLf & vbCrLf & _
                            "By Bassinger :-)"
End Sub

Private Sub Form_Load()
 Call Disco
End Sub
