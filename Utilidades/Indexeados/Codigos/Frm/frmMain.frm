VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Ver indexeados! By Bassinger"
   ClientHeight    =   3330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   222
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Alas"
      Height          =   360
      Left            =   405
      TabIndex        =   6
      Top             =   2835
      Width           =   990
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ropa"
      Height          =   360
      Left            =   420
      TabIndex        =   5
      Top             =   2400
      Width           =   990
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Casco"
      Height          =   360
      Left            =   435
      TabIndex        =   4
      Top             =   1980
      Width           =   990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Escudos"
      Height          =   360
      Left            =   450
      TabIndex        =   3
      Top             =   1575
      Width           =   990
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   1710
      TabIndex        =   2
      Top             =   255
      Width           =   2805
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   765
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Armas"
      Height          =   360
      Left            =   450
      TabIndex        =   0
      Top             =   1110
      Width           =   990
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
                
         Dim i As Long
         
         List1.Clear
         
         For i = 1 To NumObjDatas
             
             If ObjData(i).ObjType = eObjType.Arma Then
                 List1.AddItem ObjData(i).NumObj
             End If
           
         Next i
                
End Sub

Private Sub Command2_Click()
         
          Dim i As Long
         
         List1.Clear
         
         For i = 1 To NumObjDatas
             
             If ObjData(i).ObjType = eObjType.Escudo Then
                 List1.AddItem ObjData(i).NumObj
             End If
           
         Next i
         
End Sub

Private Sub Command3_Click()
Dim i As Long
         
         List1.Clear
         
         For i = 1 To NumObjDatas
             
             If ObjData(i).ObjType = eObjType.Casco Then
                 List1.AddItem ObjData(i).NumObj
             End If
           
         Next i
End Sub

Private Sub Command4_Click()
           Dim i As Long
         
         List1.Clear
         
         For i = 1 To NumObjDatas
             
             If ObjData(i).ObjType = eObjType.Armadura Then
                 List1.AddItem ObjData(i).NumObj
             End If
           
         Next i
End Sub

Private Sub Command5_Click()
Dim i As Long
         
         List1.Clear
         
         For i = 1 To NumObjDatas
             
             If ObjData(i).ObjType = eObjType.Alas Then
                 List1.AddItem ObjData(i).NumObj
             End If
           
         Next i
End Sub

Private Sub List1_Click()
     Call Grafico(ObjData(List1.Text).GrhIndex)
End Sub
