VERSION 5.00
Begin VB.Form frmSubasta 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmSubasta.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   3045
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TextBox2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   240
      Left            =   3480
      TabIndex        =   9
      Text            =   "1"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox TextBox1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   240
      Left            =   3480
      TabIndex        =   8
      Text            =   "1"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   2175
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   385
      Width           =   1845
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   2370
      ScaleHeight     =   510
      ScaleWidth      =   540
      TabIndex        =   0
      Top             =   480
      Width           =   570
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   4800
      MouseIcon       =   "frmSubasta.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "frmSubasta.frx":1994
      Top             =   100
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   810
      TabIndex        =   7
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   1485
      TabIndex        =   6
      Top             =   420
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   3600
      TabIndex        =   5
      Top             =   1155
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   3435
      TabIndex        =   4
      Top             =   1140
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   3435
      TabIndex        =   3
      Top             =   1485
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   0
      Left            =   120
      MouseIcon       =   "frmSubasta.frx":1EC2
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   2640
      Width           =   2460
   End
   Begin VB.Image Image1 
      Height          =   570
      Index           =   1
      Left            =   2400
      MouseIcon       =   "frmSubasta.frx":2014
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   2040
      Width           =   2460
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1950
      TabIndex        =   2
      Top             =   6420
      Width           =   645
   End
End
Attribute VB_Name = "frmSubasta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez



Option Explicit

'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^ y
'   le puse borde a la ventana.
'
'[END]'

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
Public LastIndex1 As Integer


Private Sub Form_Deactivate()
'frmMain.SetFocus
End Sub


Private Sub Form_Load()
'Cargamos la interfase
Me.Picture = LoadPicture(App.path & "\Graficos\Subasta.bmp")
'Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónComprar.jpg")
'Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botónvender.jpg")

End Sub

Private Sub Image1_Click(Index As Integer)

Call PlayWaveDS(SND_CLICK, SNDCHANNEL_CLICK)

If List1(Index).List(List1(Index).ListIndex) = "Nada" Or _
   List1(Index).ListIndex < 0 Then Exit Sub

Select Case Index
   Case 1
        
        If UserInventory(List1(1).ListIndex + 1).Equipped = 0 Then
           Call SendData("SUBA" & "," & List1(1).ListIndex + 1 & "," & TextBox1.Text & "," & TextBox2.Text)
        Else
            AddtoRichTextBox frmMain.RecTxt, "No podes vender el item porque lo estas usando.", 2, 51, 223, 1, 1
            Exit Sub
        End If
                
End Select

List1(1).Clear

frmMain.SetFocus
Unload Me

NPCInvDim = 0
End Sub

Private Sub Image2_Click()
SendData ("FINSUB")
End Sub

Private Sub list1_Click(Index As Integer)
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.Bottom = 32

If UserInventory(List1(1).ListIndex + 1).grhindex > 0 Then
       ' Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hdc, Inventario.grhindex(List1(1).ListIndex + 1), SR, DR)
         Call ExtractData(App.path & "\init\Graficos.Aom", str(GrhData(UserInventory(List1(1).ListIndex + 1).grhindex).FileNum).where)
            StretchDIBits Picture1.hdc, 0, 0, Bmx.gudtBMPInfo.bmiHeader.biWidth, Bmx.gudtBMPInfo.bmiHeader.biHeight, 0, 0, Bmx.gudtBMPInfo.bmiHeader.biWidth, Bmx.gudtBMPInfo.bmiHeader.biHeight, Bmx.gudtBMPData(0), Bmx.gudtBMPInfo, DIB_RGB_COLORS, SRCCOPY
 Else
            Picture1.Cls
          End If
          Picture1.Refresh

End Sub

Private Sub TextBox1_Change()
If Val(TextBox1.Text) < 1 Then
        TextBox1.Text = 1
    End If
    
    If Val(TextBox1.Text) > MAX_INVENTORY_OBJS Then
        TextBox1.Text = 1
    End If
End Sub


Private Sub TextBox2_Change()
If Val(TextBox2.Text) > 3000000 Then
        MsgBox ("No puedes pedir mas de 3 millones de monedas de oro")
        TextBox2.Text = 3000000
    End If
End Sub
