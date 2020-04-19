VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9930
   ControlBox      =   0   'False
   Icon            =   "frmCargando.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   662
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox LOGO 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   663
      TabIndex        =   0
      Top             =   0
      Width           =   9945
      Begin RichTextLib.RichTextBox Status 
         Height          =   2775
         Left            =   2400
         TabIndex        =   1
         Top             =   2760
         Visible         =   0   'False
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   4895
         _Version        =   393217
         Enabled         =   -1  'True
         MousePointer    =   99
         Appearance      =   0
         TextRTF         =   $"frmCargando.frx":08CA
      End
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Set logo.Picture = Interfaces.FrmCargando_Principal
    Set Me.MouseIcon = Iconos.Ico_Diablo
    Set logo.MouseIcon = Iconos.Ico_Diablo
    Set status.MouseIcon = Iconos.Ico_Diablo
End Sub

