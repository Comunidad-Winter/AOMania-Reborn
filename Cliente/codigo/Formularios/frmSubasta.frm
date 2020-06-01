VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "COMCTL32.OCX"
Begin VB.Form frmSubasta 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
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
   MousePointer    =   99  'Custom
   Picture         =   "frmSubasta.frx":0000
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   673
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Height          =   210
      Left            =   6885
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   8205
      Width           =   210
   End
   Begin VB.TextBox txtOferta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   2235
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Text            =   "0"
      Top             =   8085
      Width           =   1275
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   6270
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1560
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   11060
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   65535
      BackColor       =   0
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   300
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   780
      Width           =   480
   End
   Begin VB.Image cmdUpdate 
      Height          =   465
      Left            =   5505
      MousePointer    =   99  'Custom
      Top             =   930
      Width           =   1500
   End
   Begin VB.Image cmdOfrecer 
      Height          =   420
      Left            =   75
      MousePointer    =   99  'Custom
      Top             =   8040
      Width           =   2040
   End
   Begin VB.Image cmdCrearSubasta 
      Height          =   525
      Left            =   7245
      MousePointer    =   99  'Custom
      Top             =   900
      Width           =   2220
   End
   Begin VB.Image cmdCerrar 
      Height          =   420
      Left            =   8910
      MousePointer    =   99  'Custom
      Top             =   135
      Width           =   990
   End
End
Attribute VB_Name = "frmSubasta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    
    Dim i    As Integer

    Dim Item As ListItem
    
    ListView1.ListItems.Clear
    
    If Check1.Value = 0 Then
        If NumSubasta > 0 Then
       
            For i = 1 To NumSubasta
              
                Set Item = ListView1.ListItems.Add(, , Subasta(i).Objeto)
                Item.SubItems(1) = Subasta(i).Cantidad 'Cantidad
                Item.SubItems(2) = Subasta(i).Valor 'Valor
                Item.SubItems(3) = Subasta(i).Subastador 'Subastador
                Item.SubItems(4) = Subasta(i).Timer 'Tiempo
                Item.SubItems(5) = Subasta(i).Comprador 'Comprador
           
            Next i
                  
        End If

    ElseIf Check1.Value = 1 Then
       
        If NumSubasta > 0 Then
    
            For i = 1 To NumSubasta

                If UCase$(UserName) = UCase$(Subasta(i).Subastador) Then

                    Set Item = ListView1.ListItems.Add(, , Subasta(i).Objeto)
                    Item.SubItems(1) = Subasta(i).Cantidad 'Cantidad
                    Item.SubItems(2) = Subasta(i).Valor 'Valor
                    Item.SubItems(3) = Subasta(i).Subastador 'Subastador
                    Item.SubItems(4) = Subasta(i).Timer 'Tiempo
                    Item.SubItems(5) = Subasta(i).Comprador 'Comprador

                End If

            Next i
                  
        End If
           
    End If
       
End Sub

Private Sub Check1_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    Check1.MouseIcon = Iconos.Mano

End Sub

Private Sub cmdCerrar_Click()

    Call Audio.PlayWave(SND_CLICK)
      
    Unload Me
      
End Sub

Private Sub cmdCerrar_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    cmdCerrar.MouseIcon = Iconos.Ico_Mano

End Sub

Private Sub cmdCrearSubasta_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmSubastaCrear.Show , frmMain

End Sub

Private Sub cmdCrearSubasta_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    cmdCrearSubasta.MouseIcon = Iconos.Mano

End Sub

Private Sub cmdOfrecer_Click()
        
    Call Audio.PlayWave(SND_CLICK)
        
    Dim i As Integer
    
    i = ListView1.SelectedItem.Index
         
    If txtOferta.Text = 0 Then

        Call AddtoRichTextBox(frmMain.RecTxt, "La oferta a la subasta no puede ser a valor 0.", 65, 190, 156, False, , False)
        Exit Sub

    End If
     
    If txtOferta.Text <= Subasta(i).Valor Then

        Call AddtoRichTextBox(frmMain.RecTxt, "La oferta a la subasta debe ser superior de " & Subasta(i).Valor & " oro.", 65, 190, 156, False, , False)
        Exit Sub

    End If
     
    Call SendData("OFSUB" & i & "," & Subasta(i).Subastador & "," & Subasta(i).IdObjeto & "," & txtOferta.Text)
     
End Sub

Private Sub cmdOfrecer_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    cmdOfrecer.MouseIcon = Iconos.Mano

End Sub

Private Sub cmdUpdate_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call SendData("RLSUB")

End Sub

Private Sub cmdUpdate_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    cmdUpdate.MouseIcon = Iconos.Mano

End Sub

Private Sub Form_Load()
        
    With ListView1
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Objeto"
        .ColumnHeaders.Add , , "Cantidad"
        .ColumnHeaders.Add , , "Oferta"
        .ColumnHeaders.Add , , "Subastador"
        .ColumnHeaders.Add , , "Tiempo restante (En horas)"
        .ColumnHeaders.Add , , "Comprador"
          
        .ColumnHeaders(1).Width = 95
        .ColumnHeaders(2).Width = 67
        .ColumnHeaders(3).Width = 100
        .ColumnHeaders(4).Width = 60
        .ColumnHeaders(5).Width = 130
        .ColumnHeaders(6).Width = 67
            
    End With
    
    Call ListDataItem
        
End Sub

Private Sub ListDataItem()
       
    Dim Item As ListItem

    Dim i    As Byte
       
    If NumSubasta > 0 Then
       
        For i = 1 To NumSubasta
              
            Set Item = ListView1.ListItems.Add(, , Subasta(i).Objeto)
            Item.SubItems(1) = Subasta(i).Cantidad 'Cantidad
            Item.SubItems(2) = Subasta(i).Valor 'Valor
            Item.SubItems(3) = Subasta(i).Subastador 'Subastador
            Item.SubItems(4) = Subasta(i).Timer 'Tiempo
            Item.SubItems(5) = Subasta(i).Comprador 'Comprador
           
        Next i
                  
    End If
       
End Sub

Sub ReloadVentanaSubasta()
    
    With ListView1
        .ListItems.Clear

    End With
    
    Call ListDataItem
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MouseIcon = Iconos.Diablo

End Sub

Private Sub ListView1_Click()

    Dim SR As RECT, DR As RECT

    SR.Left = 0
    SR.Top = 0
    SR.Right = 32
    SR.Bottom = 32

    DR.Left = 0
    DR.Top = 0
    DR.Right = 32
    DR.Bottom = 32
    
    If NumSubasta > 0 Then
    
        Call DrawGrhtoHdc(Picture1.hdc, Subasta(ListView1.SelectedItem.Index).GrhIndex, DR)
    
        txtOferta.Text = Subasta(ListView1.SelectedItem.Index).Valor + 100
    
    End If
    
End Sub

Private Sub ListView1_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    ListView1.MouseIcon = Iconos.Diablo

End Sub

Private Sub txtOferta_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    txtOferta.MouseIcon = Iconos.Mano

End Sub
