VERSION 5.00
Begin VB.Form frmCapturas 
   Caption         =   "Capturas de pantalla"
   ClientHeight    =   10005
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   ScaleHeight     =   667
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   802
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picScreen1 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   5400
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disminuir tamaño"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aumentar tamaño"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Image picScreen 
      Height          =   8775
      Left            =   0
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmCapturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowSnap()
    Dim pHeight As Integer, pWidth As Integer

    GetDimensions pHeight, pWidth
   ' picScreen.PaintPicture picScreen.Image, 0, 0 ', (pWidth * 9) / 15, (pHeight * 9) / 15
End Sub

