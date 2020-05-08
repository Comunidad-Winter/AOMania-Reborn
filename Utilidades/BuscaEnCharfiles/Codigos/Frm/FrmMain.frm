VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "BuscaCharfiles By Bassinger"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
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
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   105
      TabIndex        =   2
      Top             =   1290
      Width           =   4245
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   360
      Left            =   1845
      TabIndex        =   1
      Top             =   780
      Width           =   990
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1260
      TabIndex        =   0
      Text            =   "Bloque=Variable=Cantidad"
      Top             =   225
      Width           =   2130
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuscar_Click()
          
       Dim Bloque As String
       Dim Variable As String
       Dim Cantidad As String
       
       Bloque = ReadField(1, Text1.Text, 61)
       Variable = ReadField(2, Text1.Text, 61)
       Cantidad = ReadField(3, Text1.Text, 61)
       
       Dim i As Integer, Nombre As String
       Dim Leer As New clsIniManager
       
       Nombre = Dir(DirChar & "*.chr")
       
       Do While Nombre <> ""
          
          Call Leer.Initialize(DirChar & Nombre)
          
          If Leer.GetValue(Bloque, Variable) = Cantidad Then
              i = i + 1
              List1.AddItem Nombre
          End If
          
          Nombre = Dir
          
       Loop
       
       MsgBox "Has encontrado " & i & " resultados.", vbInformation
          
End Sub
