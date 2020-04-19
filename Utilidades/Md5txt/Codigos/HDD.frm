VERSION 5.00
Begin VB.Form ViewHDD 
   Caption         =   "MD5Txt"
   ClientHeight    =   3255
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
   LinkTopic       =   "ViewHDD"
   ScaleHeight     =   3255
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   1575
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
      TabIndex        =   0
      Text            =   "Text1"
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

Private Sub Command2_Click()
   Text1.Text = MD5String(HDD)
   Text2.Text = "1 - Editar el Fichero gmsmac.dat" & vbCrLf & _
                            "2 - Introducir el dato en MAC=" & Text1.Text & vbCrLf & vbCrLf & _
                            "By Bassinger :-)"
End Sub

