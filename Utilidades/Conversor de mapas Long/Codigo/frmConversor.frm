VERSION 5.00
Begin VB.Form frmConversor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Argentum noline"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   1800
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   1800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Convertir"
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmConversor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCommand1_Click()

    Dim i       As Long
    Dim numMaps As Integer
    
    numMaps = CInt(InputBox("Cantidad de mapas"))

    Do While numMaps < 0

        numMaps = CInt(InputBox("Cantidad de mapas"))

    Loop
    
    ReDim MapData(1 To numMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To numMaps) As Integer
          
    For i = 1 To numMaps
        Call CargarMapaOLD(i, App.Path & "\Mapas Viejos\Mapa" & i)
        DoEvents
    Next i

    For i = 1 To numMaps
        Call GrabarMapa(i, App.Path & "\Mapas Nuevos\Mapa" & i)
        DoEvents
    Next i

    If MsgBox("Terminado", vbYesNo) = vbYes Then
        End

    End If

End Sub
