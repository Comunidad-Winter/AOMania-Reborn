VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Indexador by Columdrum"
   ClientHeight    =   7830
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10950
   Icon            =   "Encode.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmbResourceFile 
      Caption         =   "Limpiar Memoria"
      Height          =   255
      Left            =   960
      TabIndex        =   48
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CheckBox Checkcabeza 
      Caption         =   "cabeza"
      Height          =   195
      Left            =   3840
      TabIndex        =   47
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "f. verde"
      Height          =   315
      Left            =   4200
      TabIndex        =   46
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Ver Graficos"
      Height          =   495
      Index           =   9
      Left            =   0
      TabIndex        =   44
      Top             =   5040
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Fx"
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   43
      Top             =   4560
      Width           =   855
   End
   Begin VB.ComboBox CDibujarWalk 
      Height          =   315
      ItemData        =   "Encode.frx":030A
      Left            =   3840
      List            =   "Encode.frx":031D
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Botas"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   41
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Capas"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   40
      Top             =   4080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Armas"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   39
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Escudos"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   38
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Cascos"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   37
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Cabezas"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   36
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Bodys"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   35
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Graficos"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   34
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton BotonBorrrar 
      Caption         =   "borrar"
      Height          =   255
      Left            =   4440
      TabIndex        =   32
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton NuevoGhr 
      Caption         =   "Nuevo/buscar"
      Height          =   255
      Left            =   840
      TabIndex        =   29
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3240
      TabIndex        =   26
      Text            =   "Graficos"
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      TabIndex        =   25
      Text            =   "Ini"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Guardar Graficos.ind"
      Height          =   375
      Left            =   2520
      TabIndex        =   24
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton BotonGuardar 
      Caption         =   "Guardar"
      Height          =   255
      Left            =   2640
      TabIndex        =   23
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Timer Dibujado 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   240
      Top             =   120
   End
   Begin VB.TextBox TextDatos 
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   3840
      TabIndex        =   12
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   3840
      TabIndex        =   11
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   7
      Left            =   3840
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   6
      Left            =   3840
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   5
      Left            =   3840
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   4
      Left            =   3840
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   0
      Width           =   7575
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   0
      Left            =   3840
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.ListBox Lista 
      Height          =   4740
      ItemData        =   "Encode.frx":0341
      Left            =   840
      List            =   "Encode.frx":0343
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.PictureBox Visor 
      AutoRedraw      =   -1  'True
      Height          =   6375
      Left            =   5160
      ScaleHeight     =   6315
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   360
      Width           =   5775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Encode Mapas"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label DescripcionAyuda 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   2400
      TabIndex        =   45
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label LUlitError 
      BackColor       =   &H80000004&
      Height          =   975
      Left            =   0
      TabIndex        =   33
      Top             =   6840
      Width           =   10935
   End
   Begin VB.Label LGHRnumeroA 
      Caption         =   "0"
      Height          =   255
      Left            =   3240
      TabIndex        =   31
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label LNumActual 
      Caption         =   "Ghr:"
      Height          =   255
      Left            =   2640
      TabIndex        =   30
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Carpeta de graficos"
      Height          =   255
      Left            =   1800
      TabIndex        =   28
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Carpeta Inds"
      Height          =   255
      Left            =   2280
      TabIndex        =   27
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label LTexto 
      Caption         =   "Ancho Titles:"
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   22
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label LTexto 
      Caption         =   "Alto Titles:"
      Height          =   255
      Index           =   8
      Left            =   2520
      TabIndex        =   21
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label LTexto 
      Caption         =   "PosicionY:"
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   20
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label LTexto 
      Caption         =   "PosicionX:"
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   19
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label LTexto 
      Caption         =   "Velocidad:"
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   18
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label LTexto 
      Caption         =   "Ancho:"
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   17
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label LTexto 
      Caption         =   "Alto:"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   16
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label LTexto 
      Caption         =   "Numero Frames:"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   15
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label LTexto 
      Caption         =   "Frames:"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   14
      Top             =   0
      Width           =   735
   End
   Begin VB.Label LTexto 
      Caption         =   "Numero BMP:"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   13
      Top             =   720
      Width           =   1215
   End
   Begin VB.Menu MenuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MenuArchivoGuardar 
         Caption         =   "Guardar"
         Shortcut        =   ^S
      End
      Begin VB.Menu MenuBotonGuardarP 
         Caption         =   "Guardar..."
      End
      Begin VB.Menu MenuArchivoCargar 
         Caption         =   "Cargar"
         Shortcut        =   ^O
      End
      Begin VB.Menu MenuBotonCargarP 
         Caption         =   "Cargar..."
      End
      Begin VB.Menu MenuArchivoGuardado 
         Caption         =   "Indice de guardado"
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "Sobreescribir"
            Index           =   0
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "1"
            Index           =   1
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "2"
            Index           =   2
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "3"
            Index           =   3
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "4"
            Index           =   4
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "5"
            Index           =   5
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "6"
            Index           =   6
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "7"
            Index           =   7
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "8"
            Index           =   8
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "9"
            Index           =   9
         End
      End
   End
   Begin VB.Menu Medicion 
      Caption         =   "Edicion"
      Begin VB.Menu MenuEdicionNuevo 
         Caption         =   "Nuevo/Ir A"
         Shortcut        =   ^F
      End
      Begin VB.Menu menuEdicionMover 
         Caption         =   "Mover"
      End
      Begin VB.Menu MenuEdicionCopiar 
         Caption         =   "Copiar"
      End
      Begin VB.Menu MenuEdicionBorrar 
         Caption         =   "Borrar"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu MenuEdicionClonar 
         Caption         =   "Clonar..."
      End
      Begin VB.Menu menuEdicionColor 
         Caption         =   "Color de fondo..."
      End
   End
   Begin VB.Menu MenuHerramientas 
      Caption         =   "Herramientas"
      Begin VB.Menu MenuHerramientasBG 
         Caption         =   "Buscar Grh Con bmp..."
      End
      Begin VB.Menu MenuHerramientasNI 
         Caption         =   "Buscar Bmps sin indexar (Cry)"
      End
      Begin VB.Menu MenuHerramientasBN 
         Caption         =   "Buscar siguiente"
         Shortcut        =   {F3}
         Visible         =   0   'False
      End
      Begin VB.Menu MenuHerramientasAAnim 
         Caption         =   "Autoindexador"
      End
      Begin VB.Menu MenuHerramientasBR 
         Caption         =   "Buscar Grh Repetidos"
      End
   End
   Begin VB.Menu MenuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu MenuAcercaDe 
         Caption         =   "Acerca de..."
      End
   End
   Begin VB.Menu mnuautoI 
      Caption         =   "Indexar como..."
      Visible         =   0   'False
      Begin VB.Menu IAnim 
         Caption         =   "indexar como Animacion"
      End
      Begin VB.Menu mnIgeneral 
         Caption         =   "indexar como grafico individual"
      End
      Begin VB.Menu mnuibody 
         Caption         =   "indexar como body"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DibujarGHRVisor(ByVal GhrIndex As Long)

    On Error Resume Next

    If Not GrHCambiando Then
        If GhrIndex <= 0 Then Exit Sub
        Call dibujarGrh2(GrhData(GhrIndex))
        frmMain.Visor.Refresh
    Else
        Call dibujarGrh2(TempGrh)
        frmMain.Visor.Refresh

    End If

End Sub

Private Sub DibujarBMPVisor(ByVal GhrIndex As Long)
    Dim SR    As RECT, DR As RECT
    Dim Alto  As Long
    Dim Ancho As Long
    frmMain.Visor.Cls
    Dim dummy As Integer

    If GhrIndex <= 0 Then Exit Sub
    '    Call GetTamañoBMP(GhrIndex, Alto, Ancho, dummy)
    '    SR.Left = 0
    '    SR.Top = 0
    '    SR.Right = Ancho
    '    SR.Bottom = Alto
    '
    '    DR.Left = 0
    '    DR.Top = 0
    '    DR.Right = SR.Right
    '    DR.Bottom = SR.Bottom
    '    Call DrawGrhtoHdc(frmMain.Visor.hWnd, frmMain.Visor.hDC, GhrIndex, SR, DR)
    Call dibujarBMP2(GhrIndex)
    frmMain.Visor.Refresh

End Sub

Private Sub DibujarDataIndex(ByRef DataIndex As BodyData, _
                             Optional ByVal Frame As Integer = 1, _
                             Optional ByVal Index As Byte = 0)

    On Error Resume Next

    Dim SR                 As RECT, DR As RECT
    Dim r                  As RECT
    Dim sourceRect         As RECT, destRect As RECT

    Dim i                  As Long
    Dim curX               As Long
    Dim curY               As Long
    Dim GhrIndex(1 To 4)   As Grh
    Dim Posiciones(1 To 4) As Position
    Dim tGrhIndex          As Long
    curX = 0
    curY = 0

    If EstadoIndexador = e_EstadoIndexador.Fx Then
        Index = 1

    End If

    With sourceRect
        .Bottom = 500
        .Left = 0
        .Right = 500
        .Top = 0

    End With

    If (Index > 0 And Index < 5) Or EstadoIndexador = e_EstadoIndexador.Fx Then
        If DataIndex.Walk(Index).GrhIndex <= 0 Then DataIndex.Walk(Index).GrhIndex = 1
        If GrhData(DataIndex.Walk(Index).GrhIndex).NumFrames > 1 Then
            tGrhIndex = GrhData(DataIndex.Walk(Index).GrhIndex).Frames(Frame)
        Else
            tGrhIndex = DataIndex.Walk(Index).GrhIndex

        End If

        If tGrhIndex <= 0 Then Exit Sub
        
        '        SR.Left = Grhdata(tGrhIndex).sX
        '        SR.Top = Grhdata(tGrhIndex).sY
        '        SR.Right = Grhdata(tGrhIndex).sX + Grhdata(tGrhIndex).pixelWidth
        '        SR.Bottom = Grhdata(tGrhIndex).sY + Grhdata(tGrhIndex).pixelHeight
        '
        '        DR.Left = CurX
        '        DR.Top = CurY
        '        DR.Right = CurX + Grhdata(tGrhIndex).pixelWidth
        '        DR.Bottom = CurY + Grhdata(tGrhIndex).pixelHeight
        Call dibujarGrh2(GrhData(tGrhIndex))
        'Call DrawGrhtoHdc(frmMain.Visor.hWnd, frmMain.Visor.hDC, Grhdata(tGrhIndex).FileNum, SR, DR)
        frmMain.Visor.Refresh
    Else

        If DibujarFondo Then
            BackBufferSurface.BltColorFill r, ColorFondo
        Else
            BackBufferSurface.BltColorFill r, 0

        End If

        Call CalcularPosiciones(DataIndex, Posiciones)

        For i = 1 To 4

            If DataIndex.Walk(i).GrhIndex <= 0 Then DataIndex.Walk(i).GrhIndex = 1
            If GrhData(DataIndex.Walk(i).GrhIndex).NumFrames > 1 Then
                tGrhIndex = GrhData(DataIndex.Walk(i).GrhIndex).Frames(Frame)
            Else
                tGrhIndex = DataIndex.Walk(i).GrhIndex

            End If

            If tGrhIndex <= 0 Then Exit Sub
        
            SR.Left = GrhData(tGrhIndex).sX
            SR.Top = GrhData(tGrhIndex).sY
            SR.Right = GrhData(tGrhIndex).sX + GrhData(tGrhIndex).pixelWidth
            SR.Bottom = GrhData(tGrhIndex).sY + GrhData(tGrhIndex).pixelHeight
        
            DR.Left = Posiciones(i).X
            DR.Top = Posiciones(i).Y
            DR.Right = Posiciones(i).X + GrhData(tGrhIndex).pixelWidth
            DR.Bottom = Posiciones(i).Y + GrhData(tGrhIndex).pixelHeight

            Call dibujapjESpecial(BackBufferSurface, GrhData(tGrhIndex), DR.Left, DR.Top)
        
        Next i

        If EstadoIndexador = e_EstadoIndexador.Body And cabezaActual <> 0 Then
            If cabezaActual > 0 And cabezaActual <= grhCount Then
                Call dibujapjESpecial(BackBufferSurface, GrhData(cabezaActual), Posiciones(3).X + (GrhData( _
                        GrhData(DataIndex.Walk(3).GrhIndex).Frames(Frame)).pixelWidth / 2) - (GrhData( _
                        cabezaActual).pixelWidth / 2) + DataIndex.HeadOffset.X, Posiciones(3).Y + GrhData( _
                        GrhData(DataIndex.Walk(3).GrhIndex).Frames(Frame)).pixelHeight - GrhData( _
                        cabezaActual).pixelHeight + DataIndex.HeadOffset.Y - 1)

            End If

        End If

        If EstadoIndexador = e_EstadoIndexador.Cabezas Then
            If frmMain.Checkcabeza.value = vbChecked Then
                cabezaActual = DataIndex.Walk(3).GrhIndex

            End If

        End If

        'SecundaryClipper.SetHWnd frmMain.Visor.hWnd
        If DataIndex.Walk(4).GrhIndex > 0 Then
            sourceRect.Right = Posiciones(2).X + GrhData(DataIndex.Walk(4).GrhIndex).pixelWidth
        Else
            sourceRect.Right = Posiciones(2).X * 2

        End If

        If DataIndex.Walk(3).GrhIndex > 0 Then
            sourceRect.Bottom = Posiciones(3).Y + GrhData(DataIndex.Walk(3).GrhIndex).pixelHeight
        Else
            sourceRect.Bottom = Posiciones(3).Y * 2

        End If

        destRect = sourceRect
        BackBufferSurface.BltToDC frmMain.Visor.hDC, sourceRect, destRect
    
        frmMain.Visor.Refresh

    End If

End Sub

Private Sub DibujarTempGHRVisor(ByVal loopAnim As Integer)

    On Error Resume Next

    Dim SR       As RECT, DR As RECT
    Dim GhrIndex As Integer
    GhrIndex = loopAnim

    If GhrIndex <= 0 Then Exit Sub
    
    SR.Left = GrhData(GhrIndex).sX
    SR.Top = GrhData(GhrIndex).sY
    SR.Right = GrhData(GhrIndex).sX + GrhData(GhrIndex).pixelWidth
    SR.Bottom = GrhData(GhrIndex).sY + GrhData(GhrIndex).pixelHeight
    
    DR.Left = 0
    DR.Top = 0
    DR.Right = GrhData(GhrIndex).pixelWidth
    DR.Bottom = GrhData(GhrIndex).pixelHeight
    Call DrawGrhtoHdc(frmMain.Visor.hWnd, frmMain.Visor.hDC, GrhData(GhrIndex).FileNum, SR, DR)
    frmMain.Visor.Refresh

End Sub

Private Sub GetInfoGHR(ByVal GrhIndex As Long)

    If GrhIndex <= 0 Then Exit Sub
    LoadingNew = True
    Dim i           As Long
    Dim Ancho       As Long
    Dim Alto        As Long
    Dim PrimerAlto  As Long
    Dim PrimerAncho As Long
    Dim dumYin      As Integer

    TextDatos(0).Text = GrhData(GrhIndex).FileNum
    TextDatos(1).Text = ""
    TextDatos(2).Text = GrhData(GrhIndex).NumFrames

    TextDatos(3).Text = GrhData(GrhIndex).pixelHeight
    TextDatos(4).Text = GrhData(GrhIndex).pixelWidth
    TextDatos(5).Text = GrhData(GrhIndex).Speed
    TextDatos(6).Text = GrhData(GrhIndex).sX
    TextDatos(7).Text = GrhData(GrhIndex).sY
    TextDatos(8).Text = GrhData(GrhIndex).TileHeight
    TextDatos(9).Text = GrhData(GrhIndex).TileWidth
    LUlitError.Caption = ""

    If GrhData(GrhIndex).NumFrames = 1 Then
        TextDatos(1).BackColor = vbWhite
        TextDatos(1).Text = GrhData(GrhIndex).Frames(1)
        Call GetTamañoBMP(GrhData(GrhIndex).FileNum, Alto, Ancho, dumYin)
        frmMain.Dibujado.Enabled = False
        TextDatos(1).Enabled = False

        For i = 3 To 4
            TextDatos(i).Enabled = True
        Next i

        TextDatos(5).Enabled = False

        For i = 6 To 7
            TextDatos(i).Enabled = True
        Next i

    Else
        TextDatos(1).BackColor = vbWhite

        For i = 1 To GrhData(GrhIndex).NumFrames

            If i = 1 Then
                TextDatos(1).Text = GrhData(GrhIndex).Frames(i)
            Else
                TextDatos(1).Text = TextDatos(1).Text & "-" & GrhData(GrhIndex).Frames(i)

            End If

        Next i

        If GrhData(GrhIndex).Speed > 0 Then ' pervenimos division por 0
            frmMain.Dibujado.Interval = 50 '* GrhData(GrhIndex).Speed
        Else
            frmMain.Dibujado.Interval = 100

        End If

        frmMain.Dibujado.Enabled = True
        TextDatos(1).Enabled = True

        For i = 3 To 4
            TextDatos(i).Enabled = False
            TextDatos(i).BackColor = vbWhite
        Next i

        TextDatos(5).Enabled = True

        For i = 6 To 7
            TextDatos(i).Enabled = False
            TextDatos(i).BackColor = vbWhite
        Next i

    End If

    GrHCambiando = False
    LNumActual.Caption = "Ghr:"
    BotonGuardar.Visible = False
    LoadingNew = False

End Sub

Private Sub GetInfoBmp(ByVal GrhIndex As Long)

    If GrhIndex <= 0 Then Exit Sub
    Dim i             As Long
    Dim Ancho         As Long
    Dim Alto          As Long
    Dim PrimerAlto    As Long
    Dim PrimerAncho   As Long
    Dim BitCount      As Integer
    Dim existenciaBMP As Byte
    Dim ResourceS     As String

    existenciaBMP = ExisteBMP(GrhIndex)

    If existenciaBMP = 0 Then Exit Sub
    If existenciaBMP = 1 And ResourceFile = 3 Then
        If GrhIndex > 0 And GrhIndex <= UBound(ResourceF.graficos) Then
            If ResourceF.graficos(GrhIndex).tamaño > 0 Then ResourceS = "+ ResF"

        End If

    End If

    Call GetTamañoBMP(GrhIndex, Alto, Ancho, BitCount)

    If existenciaBMP = 2 Then TextDatos(0).Text = ResourceF.graficos(GrhIndex).tamaño
    TextDatos(1).Text = ""
    TextDatos(2).Text = Alto

    TextDatos(3).Text = Ancho
    TextDatos(4).Text = BitCount
    TextDatos(5).Text = StringRecurso(existenciaBMP)

    If ResourceS <> vbNullString Then TextDatos(5).Text = TextDatos(5).Text & ResourceS

    LNumActual.Caption = "BMP:"
    BotonGuardar.Visible = False

End Sub

Private Sub GetInfoDataIndex(ByVal DataIndex As Long)

    If DataIndex <= 0 Then Exit Sub
    Dim i           As Long
    Dim Ancho       As Long
    Dim Alto        As Long
    Dim PrimerAlto  As Long
    Dim PrimerAncho As Long

    LoadingNew = True
    Dim GhrIndex(1 To 4) As Long
    Dim tGrhIndex        As Long
    TextDatos(5).Visible = False
    LTexto(5).Visible = False
    TextDatos(5).Text = ""
    LUlitError.Caption = ""

    For i = 1 To 4

        If EstadoIndexador = e_EstadoIndexador.Body Then
            GhrIndex(i) = BodyData(DataIndex).Walk(i).GrhIndex
            tempDataIndex.Walk(i).GrhIndex = BodyData(DataIndex).Walk(i).GrhIndex
        ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
            GhrIndex(i) = HeadData(DataIndex).Head(i).GrhIndex
            tempDataIndex.Walk(i).GrhIndex = HeadData(DataIndex).Head(i).GrhIndex
        ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
            GhrIndex(i) = CascoAnimData(DataIndex).Head(i).GrhIndex
            tempDataIndex.Walk(i).GrhIndex = CascoAnimData(DataIndex).Head(i).GrhIndex
        ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
            GhrIndex(i) = ShieldAnimData(DataIndex).ShieldWalk(i).GrhIndex
            tempDataIndex.Walk(i).GrhIndex = ShieldAnimData(DataIndex).ShieldWalk(i).GrhIndex
        ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
            GhrIndex(i) = WeaponAnimData(DataIndex).WeaponWalk(i).GrhIndex
            tempDataIndex.Walk(i).GrhIndex = WeaponAnimData(DataIndex).WeaponWalk(i).GrhIndex
        ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
            GhrIndex(i) = BotasAnimData(DataIndex).Head(i).GrhIndex
            tempDataIndex.Walk(i).GrhIndex = BotasAnimData(DataIndex).Head(i).GrhIndex
        ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
            GhrIndex(i) = EspaldaAnimData(DataIndex).Head(i).GrhIndex
            tempDataIndex.Walk(i).GrhIndex = EspaldaAnimData(DataIndex).Head(i).GrhIndex
        ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
            GhrIndex(i) = FxData(DataIndex).Fx.GrhIndex
            tempDataIndex.Walk(i).GrhIndex = FxData(DataIndex).Fx.GrhIndex

        End If

    Next i

    TextDatos(0).Text = GhrIndex(1)
    TextDatos(2).Text = GhrIndex(2)

    TextDatos(3).Text = GhrIndex(3)
    TextDatos(4).Text = GhrIndex(4)

    If EstadoIndexador = e_EstadoIndexador.Body Then
        TextDatos(5).Text = BodyData(DataIndex).HeadOffset.Y & "º" & BodyData(DataIndex).HeadOffset.X
        tempDataIndex.HeadOffset.X = BodyData(DataIndex).HeadOffset.X
        tempDataIndex.HeadOffset.Y = BodyData(DataIndex).HeadOffset.Y
        TextDatos(5).Visible = True
        LTexto(5).Visible = True
    ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
        TextDatos(2).Text = FxData(DataIndex).OffsetY & "º" & FxData(DataIndex).OffsetX
        tempDataIndex.HeadOffset.X = FxData(DataIndex).OffsetX
        tempDataIndex.HeadOffset.Y = FxData(DataIndex).OffsetY
        TextDatos(2).Visible = True
        LTexto(2).Visible = True

    End If

    GrHCambiando = False
    BotonGuardar.Visible = False
    
    LoadingNew = False

End Sub

Private Sub BotonBorrrar_Click()
    Call SBotonBorrrar

End Sub

Public Sub CambiarEstado(ByVal Index As Integer)
    ' Cambia el estado del indexador entre las distintas secciones. Oculta/cambia labels

    Dim i As Long
    EstadoIndexador = Index
    Dibujado.Enabled = False
    Visor.Cls
    Lista.Clear
    GrHCambiando = False
    CDibujarWalk.Visible = False
    LUlitError.Caption = ""
    MenuEdicionClonar.Visible = False
    MenuHerramientas.Visible = False
    Command10.Visible = True
    BotonBorrrar.Visible = True
    DescripcionAyuda.Visible = False
    Checkcabeza.Visible = False

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh
            Call RenuevaListaGrH   'mostramos lista de grhs

            For i = 0 To 9
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                LTexto(i).Visible = True
            Next i

            MenuHerramientas.Visible = True
            MenuHerramientasBN.Visible = False
            MenuEdicionClonar.Visible = True
            LNumActual.Caption = "Grh: "
            LTexto(0).Caption = "Numero BMP:"
            LTexto(1).Caption = "Frames:"
            LTexto(2).Caption = "Numero Frames:"
            LTexto(3).Caption = "Alto:"
            LTexto(4).Caption = "Ancho:"
            LTexto(5).Caption = "Velocidad:"
            LTexto(6).Caption = "PosicionX:"
            LTexto(7).Caption = "PosicionY:"
            LTexto(8).Caption = "Alto Titles:"
            LTexto(9).Caption = "Ancho Titles:"
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Body
            Checkcabeza.Visible = True
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaBodys

            For i = 0 To 5
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 6 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Body: "
            LTexto(0).Caption = "Subir:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "izquierda:"
            LTexto(5).Caption = "Offset"
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Cabezas
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaCabezas

            For i = 0 To 4
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 5 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Cabeza: "
            LTexto(0).Caption = "Subir:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "izquierda:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Cascos
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaCascos

            For i = 0 To 4
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 5 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Casco: "
            LTexto(0).Caption = "Subir:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "izquierda:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Escudos
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaEscudos

            For i = 0 To 4
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 5 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Escudo: "
            LTexto(0).Caption = "Subir:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "izquierda:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Armas
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaArmas

            For i = 0 To 4
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 5 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Armas: "
            LTexto(0).Caption = "Subir:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "izquierda:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Botas
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaBotas

            For i = 0 To 4
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 5 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Botas: "
            LTexto(0).Caption = "Subir:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "izquierda:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Capas
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaCapas

            For i = 0 To 4
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 5 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Capa: "
            LTexto(0).Caption = "Subir:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "izquierda:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Fx
            Call RenuevaListaFX

            For i = 0 To 2
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 3 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Fx: "
            LTexto(0).Caption = "NumGrh:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "Offset:"
            LTexto(3).Caption = ""
            LTexto(4).Caption = ""
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Resource
            Call RenuevaListaResource

            For i = 0 To 5
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 6 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Crypt: "
            LTexto(0).Caption = "Tamaño:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "Alto:"
            LTexto(3).Caption = "Ancho:"
            LTexto(4).Caption = "Bits:"
            LTexto(5).Caption = "Situacion:"
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""
            Command10.Visible = False
            BotonBorrrar.Visible = False
            DescripcionAyuda.Visible = True
            DescripcionAyuda.Caption = "N:Si estan disponible el BMP y el archivo de recursos, se usa el bmp"

    End Select

    Call CambiarcaptionCommand10

End Sub

Private Sub MoverGrh(ByVal numGRH As Integer, _
                     ByVal OrigenGRH As Integer, _
                     ByVal BorrarOriginal As Boolean)
    Dim tempLong  As Long
    Dim cadena    As String
    Dim respuesta As Byte
    Dim GrhVacio  As GrhData
    Dim looPero   As Long

    tempLong = ListaindexGrH(OrigenGRH)

    If tempLong <= 0 Then
        LUlitError.Caption = "grafico incorrecto"
        Exit Sub

    End If

    tempLong = ListaindexGrH(numGRH)

    If tempLong > 0 Then
        respuesta = MsgBox("El grafico " & numGRH & " ya existe, ¿deseas sobreescribirlo?", 4, "Aviso")

        If respuesta = vbYes Then
            GrhData(numGRH) = GrhData(OrigenGRH)

            If BorrarOriginal Then
                GrhData(OrigenGRH) = GrhVacio

            End If

            GRHActual = Val(numGRH)
            LOOPActual = 1
            'frmMain.Visor.Cls
            'Call DibujarGHRVisor(GRHActual)
            'Call GetInfoGHR(GRHActual)
            LGHRnumeroA.Caption = GRHActual
            tempLong = ListaindexGrH(GRHActual)
            frmMain.Lista.listIndex = tempLong
            EstadoNoGuardado(e_EstadoIndexador.Grh) = True

        End If

    Else
        GrhData(numGRH) = GrhData(OrigenGRH)

        If BorrarOriginal Then
            GrhData(OrigenGRH) = GrhVacio

        End If

        GRHActual = numGRH
        LOOPActual = 1
        'frmMain.Visor.Cls
        'Call DibujarGHRVisor(GRHActual)
        'Call GetInfoGHR(GRHActual)
        LGHRnumeroA.Caption = GRHActual
        tempLong = ListaindexGrH(GRHActual)
        frmMain.Lista.listIndex = tempLong
        EstadoNoGuardado(e_EstadoIndexador.Grh) = True

    End If
    
End Sub

Private Sub SBotonMover(ByVal BorrarOriginal As Boolean, _
                        Optional ByVal CantidadM As Integer = 1)
    Dim tempLong  As Long
    Dim cadena    As String
    Dim respuesta As Byte
    Dim GrhVacio  As GrhData
    Dim LooPer    As Long

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh

            If GrHCambiando Then
                GrHCambiando = False
                LNumActual.Caption = "Ghr:"
                BotonGuardar.Visible = False

            End If

            cadena = InputBox("Introduzca número de GHR al que quieres mover el grafico " & GRHActual, _
                    "Mover Grafico")

            If IsNumeric(cadena) Then
                If Val(cadena) > 0 And Val(cadena) < grhCount Then
                    Call MoverGrh(Val(cadena), GRHActual, BorrarOriginal)
                    Call RenuevaListaGrH
                    tempLong = ListaindexGrH(Val(cadena))
                    frmMain.Lista.listIndex = tempLong
                Else
                    LUlitError.Caption = "introduzca un numero correcto"

                End If

            Else
                LUlitError.Caption = "introduzca un numero"

            End If

        Case Else
            Dim StringCaso As String

            If EstadoIndexador = Body Then
                StringCaso = "Body"
            ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
                StringCaso = "Cabeza"
            ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
                StringCaso = "Casco"
            ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
                StringCaso = "Escudo"
            ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
                StringCaso = "Arma"
            ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
                StringCaso = "Bota"
            ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
                StringCaso = "Capa"
            ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
                StringCaso = "Fx"
            ElseIf EstadoIndexador = e_EstadoIndexador.Resource Then
                Exit Sub

            End If

            cadena = InputBox("Introduzca numero de " & StringCaso & " al que quieres mover", "Mover " & _
                    StringCaso)

            If IsNumeric(cadena) And (Val(cadena) < grhCount) Then
                If EstadoIndexador = e_EstadoIndexador.Body Then
                    Call mueveBody(Val(cadena), DataIndexActual, BorrarOriginal)
                ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
                    Call MueveCabeza(Val(cadena), DataIndexActual, BorrarOriginal)
                ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
                    Call MueveCasco(Val(cadena), DataIndexActual, BorrarOriginal)
                ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
                    Call MueveEscudo(Val(cadena), DataIndexActual, BorrarOriginal)
                ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
                    Call MueveArma(Val(cadena), DataIndexActual, BorrarOriginal)
                ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
                    Call MueveBota(Val(cadena), DataIndexActual, BorrarOriginal)
                ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
                    Call MueveCapa(Val(cadena), DataIndexActual, BorrarOriginal)
                ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
                    Call MueveFX(Val(cadena), DataIndexActual, BorrarOriginal)

                End If

                DataIndexActual = Val(cadena)
                LOOPActual = 1
                frmMain.Visor.Cls
                Call GetInfoDataIndex(DataIndexActual)
                Dibujado.Interval = 100
                Dibujado.Enabled = True
                LGHRnumeroA.Caption = DataIndexActual
                tempLong = ListaindexGrH(DataIndexActual)
                frmMain.Lista.listIndex = tempLong
                EstadoNoGuardado(EstadoIndexador) = True
            Else
                LUlitError.Caption = "introduzca un numero valido"

            End If

    End Select

End Sub

Private Sub BotonI_Click(Index As Integer)
    Call CambiarEstado(Index)

    Call ComprobarIndexLista

End Sub

Private Sub CDibujarWalk_Click()
    DibujarWalk = CDibujarWalk.listIndex
    Visor.Cls

End Sub

Private Sub Checkcabeza_Click()

    If Checkcabeza.value = vbChecked Then
        cabezaActual = 3008
    Else
        cabezaActual = 0

    End If

End Sub

Private Sub CmbResourceFile_Click()
    'Borramos todas las surfaces de la memoria. Sirve por si se hacen cambios en los BMPs y se necesita obligar a recargarlos
    
    If Not IniciadoTodo Then
        Call SurfaceDB.BorrarTodo
    Else
        IniciadoTodo = False

    End If
    
    Call CambiarEstado(EstadoIndexador)
    
    Call ComprobarIndexLista

End Sub

Private Sub SBotonBorrrar()
    Dim respuesta As Byte
    Dim tempLong  As Long

    tempLong = frmMain.Lista.listIndex

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh

            If GRHActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el Grh " & GRHActual & " ?", 4, _
                    "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                GrhData(GRHActual).NumFrames = 0
                'Call RenuevaListaGrH
                frmMain.Lista.RemoveItem tempLong

            End If

        Case e_EstadoIndexador.Body

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el body " & DataIndexActual & " ?", 4, _
                    "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                BodyData(DataIndexActual).Walk(1).GrhIndex = 0
                BodyData(DataIndexActual).Walk(2).GrhIndex = 0
                BodyData(DataIndexActual).Walk(3).GrhIndex = 0
                BodyData(DataIndexActual).Walk(4).GrhIndex = 0
                'Call RenuevaListaBodys
                frmMain.Lista.RemoveItem tempLong

            End If

        Case e_EstadoIndexador.Cabezas

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar la Cabeza " & DataIndexActual & " ?", 4, _
                    "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                HeadData(DataIndexActual).Head(1).GrhIndex = 0
                HeadData(DataIndexActual).Head(2).GrhIndex = 0
                HeadData(DataIndexActual).Head(3).GrhIndex = 0
                HeadData(DataIndexActual).Head(4).GrhIndex = 0
                'Call RenuevaListaCabezas
                frmMain.Lista.RemoveItem tempLong

            End If

        Case e_EstadoIndexador.Cascos

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el casco " & DataIndexActual & " ?", 4, _
                    "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                CascoAnimData(DataIndexActual).Head(1).GrhIndex = 0
                CascoAnimData(DataIndexActual).Head(2).GrhIndex = 0
                CascoAnimData(DataIndexActual).Head(3).GrhIndex = 0
                CascoAnimData(DataIndexActual).Head(4).GrhIndex = 0
                'Call RenuevaListaCascos
                frmMain.Lista.RemoveItem tempLong

            End If

        Case e_EstadoIndexador.Escudos

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el escudo " & DataIndexActual & " ?", 4, _
                    "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                ShieldAnimData(DataIndexActual).ShieldWalk(1).GrhIndex = 0
                ShieldAnimData(DataIndexActual).ShieldWalk(2).GrhIndex = 0
                ShieldAnimData(DataIndexActual).ShieldWalk(3).GrhIndex = 0
                ShieldAnimData(DataIndexActual).ShieldWalk(4).GrhIndex = 0
                frmMain.Lista.RemoveItem tempLong

                'Call RenuevaListaEscudos
            End If

        Case e_EstadoIndexador.Armas

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el arma " & DataIndexActual & " ?", 4, _
                    "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                WeaponAnimData(DataIndexActual).WeaponWalk(1).GrhIndex = 0
                WeaponAnimData(DataIndexActual).WeaponWalk(2).GrhIndex = 0
                WeaponAnimData(DataIndexActual).WeaponWalk(3).GrhIndex = 0
                WeaponAnimData(DataIndexActual).WeaponWalk(4).GrhIndex = 0
                'Call RenuevaListaArmas
                frmMain.Lista.RemoveItem tempLong

            End If

        Case e_EstadoIndexador.Botas

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar la bota " & DataIndexActual & " ?", 4, _
                    "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                BotasAnimData(DataIndexActual).Head(1).GrhIndex = 0
                BotasAnimData(DataIndexActual).Head(2).GrhIndex = 0
                BotasAnimData(DataIndexActual).Head(3).GrhIndex = 0
                BotasAnimData(DataIndexActual).Head(4).GrhIndex = 0
                'Call RenuevaListaBotas
                frmMain.Lista.RemoveItem tempLong

            End If

        Case e_EstadoIndexador.Capas

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el capa " & DataIndexActual & " ?", 4, _
                    "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                EspaldaAnimData(DataIndexActual).Head(1).GrhIndex = 0
                EspaldaAnimData(DataIndexActual).Head(2).GrhIndex = 0
                EspaldaAnimData(DataIndexActual).Head(3).GrhIndex = 0
                EspaldaAnimData(DataIndexActual).Head(4).GrhIndex = 0
                'Call RenuevaListaCapas
                frmMain.Lista.RemoveItem tempLong

            End If

        Case e_EstadoIndexador.Fx

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el FX " & DataIndexActual & " ?", 4, _
                    "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                FxData(DataIndexActual).Fx.GrhIndex = 0
                FxData(DataIndexActual).OffsetX = 0
                FxData(DataIndexActual).OffsetY = 0
                'Call RenuevaListaFX
                frmMain.Lista.RemoveItem tempLong

            End If

    End Select

    If tempLong < frmMain.Lista.ListCount Then
        frmMain.Lista.listIndex = tempLong
    Else
        frmMain.Lista.listIndex = frmMain.Lista.ListCount - 1

    End If

End Sub

Public Function StringGuardadoActual(PEstado As e_EstadoIndexador) As String
    Dim elq As String

    Select Case PEstado

        Case e_EstadoIndexador.Grh

            If SavePath = 0 Then
                elq = "Graficos"
            Else
                elq = "Graficos" & SavePath

            End If

        Case e_EstadoIndexador.Body

            If SavePath = 0 Then
                elq = "personajes"
            Else
                elq = "personajes" & SavePath

            End If

        Case e_EstadoIndexador.Cabezas

            If SavePath = 0 Then
                elq = "cabezas"
            Else
                elq = "cabezas" & SavePath

            End If

        Case e_EstadoIndexador.Cascos

            If SavePath = 0 Then
                elq = "cascos"
            Else
                elq = "cascos" & SavePath

            End If

        Case e_EstadoIndexador.Escudos

            If SavePath = 0 Then
                elq = "escudos"
            Else
                elq = "escudos" & SavePath

            End If

        Case e_EstadoIndexador.Armas

            If SavePath = 0 Then
                elq = "armas"
            Else
                elq = "armas" & SavePath

            End If

        Case e_EstadoIndexador.Botas

            If SavePath = 0 Then
                elq = "botas"
            Else
                elq = "botas" & SavePath

            End If

        Case e_EstadoIndexador.Capas

            If SavePath = 0 Then
                elq = "capas"
            Else
                elq = "capas" & SavePath

            End If

        Case e_EstadoIndexador.Fx

            If SavePath = 0 Then
                elq = "fxs"
            Else
                elq = "fxs" & SavePath

            End If

        Case e_EstadoIndexador.Resource
            elq = ""

    End Select

    StringGuardadoActual = elq

End Function

Private Sub CambiarcaptionCommand10()
    Command10.Caption = "Guardar " & StringGuardadoActual(EstadoIndexador)
    MenuArchivoGuardar.Caption = "Guardar " & StringGuardadoActual(EstadoIndexador)
    MenuArchivoCargar.Caption = "Cargar " & StringGuardadoActual(EstadoIndexador)

    If EstadoIndexador = e_EstadoIndexador.Escudos Or EstadoIndexador = e_EstadoIndexador.Armas Then
        Command10.Caption = Command10.Caption & ".dat"
        MenuArchivoGuardar.Caption = MenuArchivoGuardar.Caption & ".dat"
        MenuArchivoCargar.Caption = MenuArchivoCargar.Caption & ".dat"
    Else
        Command10.Caption = Command10.Caption & ".ind"
        MenuArchivoGuardar.Caption = MenuArchivoGuardar.Caption & ".ind"
        MenuArchivoCargar.Caption = MenuArchivoCargar.Caption & ".ind"

    End If

End Sub

Private Sub Command1_Click()
    DibujarFondo = Not DibujarFondo
    Call ClickEnLista

End Sub

Private Sub Command10_Click()
    'Boton de guardado en disco
    Call BotonGuardado

End Sub

Private Sub Command4_Click()

    On Error GoTo ErrHandler

    Call CargarTips

    Dim N As Integer, i As Integer
    N = FreeFile

    Open App.Path & "\" & CarpetaDeInis & "\Tips.ayu" For Binary As #N
    'Escribimos la cabecera
    Put #N, , MiCabecera
    'Guardamos las cabezas
    Put #N, , NumTips

    For i = 1 To NumTips
        Put #N, , Tips(i)
    Next i

    Close #N
    Call MsgBox("Listo, encode ok!!")

    Exit Sub
ErrHandler:
    Call MsgBox("Error en tip " & i)

End Sub

Private Sub Command6_Click()

    On Error GoTo ErrHandler

    Call CargarMapas

    Dim N As Integer, i As Integer
    N = FreeFile
    Open App.Path & "\" & CarpetaDeInis & "\FK.ind" For Binary As #N
    'Escribimos la cabecera
    Put #N, , MiCabecera
    'Guardamos las cabezas
    Put #N, , NumMapas

    For i = 1 To NumMapas
        Put #N, , Mapas(i)
    Next i

    Close #N

    Call MsgBox("Listo, encode ok!!")

    Exit Sub

ErrHandler:
    Call MsgBox("Error en casco " & i)

End Sub

Private Sub BotonGuardar_Click()
    'boton de guardado en memoria
    Dim i As Long

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh
            'guardando un grafico
        
            If GRHActual = 0 Then Exit Sub
    
            If Val(TextDatos(2).Text) <= 0 Then ' numframes = 0
                MsgBox "numero de frames incorrecto"
                Exit Sub

            End If
    
            If Val(TextDatos(2).Text) = 1 Then ' si no es animacion se comprueba si existe el BMP
                If ExisteBMP(Val(TextDatos(0).Text)) = ResourceFile Or (ResourceFile And ExisteBMP(Val( _
                        TextDatos(0).Text)) > 0) Then
                Else
                    LUlitError.Caption = "No existe el archivo del grafico"
                    Exit Sub

                End If

            End If
        
            GrhData(GRHActual).FileNum = Val(TextDatos(0).Text)
            GrhData(GRHActual).NumFrames = Val(TextDatos(2).Text)

            If GrhData(GRHActual).NumFrames = 1 Then
                GrhData(GRHActual).Frames(1) = GRHActual
                GrhData(GRHActual).pixelHeight = Val(TextDatos(3).Text)
                GrhData(GRHActual).pixelWidth = Val(TextDatos(4).Text)
                GrhData(GRHActual).Speed = Val(TextDatos(5).Text)
                GrhData(GRHActual).sX = Val(TextDatos(6).Text)
                GrhData(GRHActual).sY = Val(TextDatos(7).Text)
        
                GrhData(GRHActual).TileHeight = GrhData(GRHActual).pixelHeight / TilePixelHeight
                GrhData(GRHActual).TileWidth = GrhData(GRHActual).pixelWidth / TilePixelWidth
            Else

                For i = 1 To GrhData(GRHActual).NumFrames

                    If Val(ReadField(i, TextDatos(1).Text, Asc("-"))) < 32000 Then
                        GrhData(GRHActual).Frames(i) = Val(ReadField(i, TextDatos(1).Text, Asc("-")))

                    End If

                Next i

                GrhData(GRHActual).Speed = Val(TextDatos(5).Text)

                If GrhData(GRHActual).Frames(1) > 0 Then
                    GrhData(GRHActual).pixelHeight = GrhData(GrhData(GRHActual).Frames(1)).pixelHeight
                    GrhData(GRHActual).pixelWidth = GrhData(GrhData(GRHActual).Frames(1)).pixelWidth
                    GrhData(GRHActual).sX = GrhData(GrhData(GRHActual).Frames(1)).sX
                    GrhData(GRHActual).sY = GrhData(GrhData(GRHActual).Frames(1)).sY
                    GrhData(GRHActual).TileHeight = GrhData(GrhData(GRHActual).Frames(1)).TileHeight
                    GrhData(GRHActual).TileWidth = GrhData(GrhData(GRHActual).Frames(1)).TileWidth
                Else
                    GrhData(GRHActual).pixelHeight = Val(TextDatos(3).Text)
                    GrhData(GRHActual).pixelWidth = Val(TextDatos(4).Text)
                    GrhData(GRHActual).sX = Val(TextDatos(6).Text)
                    GrhData(GRHActual).sY = Val(TextDatos(7).Text)
                    GrhData(GRHActual).TileHeight = GrhData(GRHActual).pixelHeight / TilePixelHeight
                    GrhData(GRHActual).TileWidth = GrhData(GRHActual).pixelWidth / TilePixelWidth

                End If

            End If
        
            Call GetInfoGHR(GRHActual)
            frmMain.Visor.Cls
            Call DibujarGHRVisor(GRHActual)

        Case e_EstadoIndexador.Body
            BodyData(DataIndexActual).HeadOffset.Y = Val(ReadField(1, TextDatos(5).Text, Asc("º")))
            BodyData(DataIndexActual).HeadOffset.X = Val(ReadField(2, TextDatos(5).Text, Asc("º")))
            BodyData(DataIndexActual).Walk(1).GrhIndex = Val(TextDatos(0).Text)
            BodyData(DataIndexActual).Walk(2).GrhIndex = Val(TextDatos(2).Text)
            BodyData(DataIndexActual).Walk(3).GrhIndex = Val(TextDatos(3).Text)
            BodyData(DataIndexActual).Walk(4).GrhIndex = Val(TextDatos(4).Text)

        Case e_EstadoIndexador.Cabezas
            HeadData(DataIndexActual).Head(1).GrhIndex = Val(TextDatos(0).Text)
            HeadData(DataIndexActual).Head(2).GrhIndex = Val(TextDatos(2).Text)
            HeadData(DataIndexActual).Head(3).GrhIndex = Val(TextDatos(3).Text)
            HeadData(DataIndexActual).Head(4).GrhIndex = Val(TextDatos(4).Text)

        Case e_EstadoIndexador.Cascos
            CascoAnimData(DataIndexActual).Head(1).GrhIndex = Val(TextDatos(0).Text)
            CascoAnimData(DataIndexActual).Head(2).GrhIndex = Val(TextDatos(2).Text)
            CascoAnimData(DataIndexActual).Head(3).GrhIndex = Val(TextDatos(3).Text)
            CascoAnimData(DataIndexActual).Head(4).GrhIndex = Val(TextDatos(4).Text)

        Case e_EstadoIndexador.Armas
            WeaponAnimData(DataIndexActual).WeaponWalk(1).GrhIndex = Val(TextDatos(0).Text)
            WeaponAnimData(DataIndexActual).WeaponWalk(2).GrhIndex = Val(TextDatos(2).Text)
            WeaponAnimData(DataIndexActual).WeaponWalk(3).GrhIndex = Val(TextDatos(3).Text)
            WeaponAnimData(DataIndexActual).WeaponWalk(4).GrhIndex = Val(TextDatos(4).Text)

        Case e_EstadoIndexador.Escudos
            ShieldAnimData(DataIndexActual).ShieldWalk(1).GrhIndex = Val(TextDatos(0).Text)
            ShieldAnimData(DataIndexActual).ShieldWalk(2).GrhIndex = Val(TextDatos(2).Text)
            ShieldAnimData(DataIndexActual).ShieldWalk(3).GrhIndex = Val(TextDatos(3).Text)
            ShieldAnimData(DataIndexActual).ShieldWalk(4).GrhIndex = Val(TextDatos(4).Text)

        Case e_EstadoIndexador.Botas
            BotasAnimData(DataIndexActual).Head(1).GrhIndex = Val(TextDatos(0).Text)
            BotasAnimData(DataIndexActual).Head(2).GrhIndex = Val(TextDatos(2).Text)
            BotasAnimData(DataIndexActual).Head(3).GrhIndex = Val(TextDatos(3).Text)
            BotasAnimData(DataIndexActual).Head(4).GrhIndex = Val(TextDatos(4).Text)

        Case e_EstadoIndexador.Capas
            EspaldaAnimData(DataIndexActual).Head(1).GrhIndex = Val(TextDatos(0).Text)
            EspaldaAnimData(DataIndexActual).Head(2).GrhIndex = Val(TextDatos(2).Text)
            EspaldaAnimData(DataIndexActual).Head(3).GrhIndex = Val(TextDatos(3).Text)
            EspaldaAnimData(DataIndexActual).Head(4).GrhIndex = Val(TextDatos(4).Text)

        Case e_EstadoIndexador.Fx
            FxData(DataIndexActual).Fx.GrhIndex = Val(TextDatos(0).Text)
            FxData(DataIndexActual).OffsetX = Val(ReadField(2, TextDatos(2).Text, Asc("º")))
            FxData(DataIndexActual).OffsetY = Val(ReadField(1, TextDatos(2).Text, Asc("º")))

    End Select

    If EstadoIndexador <> e_EstadoIndexador.Grh Then
        Call GetInfoDataIndex(DataIndexActual)

    End If

    EstadoNoGuardado(EstadoIndexador) = True

End Sub

Private Sub Dibujado_Timer()

    On Error Resume Next

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh

            If Not GrHCambiando Then
                If GRHActual <= 0 Then Exit Sub
                If LOOPActual > GrhData(GRHActual).NumFrames Then LOOPActual = 1
                Call DibujarGHRVisor(GrhData(GRHActual).Frames(LOOPActual))
                LOOPActual = LOOPActual + 1
            Else

                If LOOPActual > TempGrh.NumFrames Then LOOPActual = 1
                Call DibujarTempGHRVisor(TempGrh.Frames(LOOPActual))
                LOOPActual = LOOPActual + 1

            End If

        Case e_EstadoIndexador.Resource

        Case Else

            If DataIndexActual <= 0 Then Exit Sub
            If tempDataIndex.Walk(1).GrhIndex = 0 Then Exit Sub
            If LOOPActual > GrhData(tempDataIndex.Walk(1).GrhIndex).NumFrames Then LOOPActual = 1
            Call DibujarDataIndex(tempDataIndex, LOOPActual, DibujarWalk)
            LOOPActual = LOOPActual + 1

    End Select

End Sub

Private Sub Form_close()
    End

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i          As Long
    Dim HayCambios As Boolean
    Dim respuesta  As Byte
    Dim Tstr       As String

Inicio:
    HayCambios = False

    For i = e_EstadoIndexador.Grh To Resource

        If EstadoNoGuardado(i) Then
            HayCambios = True
            Tstr = Tstr & StringGuardadoActual(i) & vbCrLf

        End If

    Next i

    If HayCambios Then
        respuesta = MsgBox("Hay cambios sin Guardar en :" & vbCrLf & vbCrLf & Tstr & vbCrLf & _
                "¿Quieres GUARDAR los cambios antes de salir?" & vbCrLf & _
                "(Si pulsas NO se perderan estos cambios)" & vbCrLf, 3, "Aviso")

        If respuesta = vbCancel Then
            Cancel = 1 ' cancelamos la salida
            Exit Sub
        ElseIf respuesta = vbYes Then

            For i = e_EstadoIndexador.Grh To e_EstadoIndexador.Resource ' guardamos

                If EstadoNoGuardado(i) Then
                    EstadoIndexador = i
                    Call BotonGuardado

                End If

            Next i

            Tstr = vbNullString
            GoTo Inicio ' weno el goto es el alien d la programacion estructurada pero paso d romperme la cabeza xD asi se ve mejor

            ' volvemos a comprobar si algo no se guardo
        End If
        
    End If

End Sub

Private Sub Form_resize()
    Visor.Height = Abs(frmMain.Height - Visor.Top - LUlitError.Height - 810)
    Visor.Width = Abs(frmMain.Width - Visor.Left - 120)
    LUlitError.Top = Abs(frmMain.Height - LUlitError.Height - 705)
    LUlitError.Width = Abs(frmMain.Width - 155)
    Call ClickEnLista

End Sub

Private Sub Form_Load()

    'configuracion inicial:
    SavePath = 0
    LoadingNew = False ' variable que evita redibujado excesibo
    IniciadoTodo = True
    ColorFondo = vbGreen
    CarpetaDeInis = GetVar(App.Path & "\Conf.ini", "Config", "CarpetaDeInis")
    CarpetaGraficos = GetVar(App.Path & "\Conf.ini", "Config", "CarpetaGraficos")
    ResourceFile = 1 ' siempre cargamos lo bmps, esta deshabilitado el archivo de recursos.
    
    If ResourceFile <= 0 Then ResourceFile = 1
    If CarpetaDeInis = vbNullString Then CarpetaDeInis = "INIT"
    If CarpetaGraficos = vbNullString Then CarpetaGraficos = "graficos"
    
    Text1.Text = CarpetaDeInis
    Text2.Text = CarpetaGraficos
    Call IniciarCabecera(MiCabecera)
    
    Call IniciarObjetosDirectX
    Set SurfaceDB = New clsSurfaceManDyn
    Call InitTileEngine(frmMain.hWnd, 155, 16, 32, 32, 13, 17, 9)
    

    Call CargarTips
    
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos

    EstadoIndexador = e_EstadoIndexador.Grh
    Dim Lister As Long

    For Lister = 0 To 9
        MenuIndiceGuardado(Lister).Checked = False
    Next Lister

    MenuIndiceGuardado(0).Checked = True
    Call CambiarcaptionCommand10
    
End Sub

Private Sub CargarMapas()
    Dim loopc As Integer

    NumMapas = Val(GetVar(App.Path & "\encode\mapas.dat", "INIT", "NumMaps"))

    ReDim Mapas(0 To NumMapas + 1) As Byte

    For loopc = 1 To NumMapas
        Mapas(loopc) = Val(GetVar(App.Path & "\encode\mapas.dat", "Map" & loopc, "Lluvia"))
    Next loopc

End Sub

Private Sub CargarTips()
    Dim loopc As Integer
    NumTips = Val(GetVar(App.Path & "\encode\tips.dat", "INIT", "Tips"))

    ReDim Tips(0 To NumTips + 1) As String * 255

    For loopc = 1 To NumTips
        Tips(loopc) = GetVar(App.Path & "\encode\tips.dat", "Tip" & loopc, "Tip")
    Next loopc

End Sub

Private Sub IAnim_Click()
    Dim textoActual As String
    Dim bmpstring   As String
    Dim BMPBuscado  As Long

    bmpstring = ReadField(1, Lista.List(Lista.listIndex), Asc(" "))

    If IsNumeric(bmpstring) Then
        BMPBuscado = Val(bmpstring)

        If BMPBuscado > 0 And BMPBuscado <= 32000 Then
            If FormAuto.Visible Then
                FormAuto.SetFocus
            Else
                FormAuto.Show , frmMain

            End If

            FormAuto.FrameAnim(0).Visible = True
            FormAuto.FrameAnim(1).Visible = False
            FormAuto.FrameAnim(2).Visible = False
            FormAuto.TextDatos(4).Text = BMPBuscado

        End If

    End If

End Sub

Private Sub Lista_Click()
    Call ClickEnLista

End Sub

Public Sub ClickEnLista()

    On Error Resume Next

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh
            GRHActual = Val(ReadField(1, Lista.List(Lista.listIndex), Asc(" ")))
            LOOPActual = 1
            Call GetInfoGHR(GRHActual)
            Call DibujarGHRVisor(GRHActual)
            LGHRnumeroA.Caption = GRHActual

        Case e_EstadoIndexador.Resource
            GRHActual = Val(Lista.List(Lista.listIndex))
            frmMain.Visor.Cls

            If ExisteBMP(GRHActual) = 0 Then Exit Sub
            Call GetInfoBmp(GRHActual)
            Call DibujarBMPVisor(GRHActual)
            LGHRnumeroA.Caption = GRHActual

        Case Else
            frmMain.Visor.Cls
            DataIndexActual = Val(Lista.List(Lista.listIndex))
            LOOPActual = 1
            Call GetInfoDataIndex(DataIndexActual)
        
            LGHRnumeroA.Caption = DataIndexActual

    End Select

    UltimoindexE(EstadoIndexador) = Lista.listIndex

End Sub

Private Sub Lista_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton And EstadoIndexador = Resource Then
        Call Me.PopupMenu(Me.mnuautoI)

    End If

End Sub

Private Sub MenuAcercaDe_Click()
    MsgBox "Indexador Creado por columdrum" & vbCrLf & vbCrLf & "Email: columdrum@gmail.com" & vbCrLf & _
            vbCrLf & "Version: " & VERSION_ACTUAL

End Sub

Private Sub MenuArchivoCargar_Click()
    Call BotonCargado

End Sub

Private Sub MenuArchivoGuardar_Click()
    Call BotonGuardado

End Sub

Private Sub MenuBotonCargarP_Click()

    On Error GoTo Cancelar

    With CommonDialog1
        .Filter = "Binario(.ind)|*.ind|Archivo de texto DAT(*.dat)|*.dat"

        If EstadoIndexador = e_EstadoIndexador.Armas Or EstadoIndexador = e_EstadoIndexador.Escudos Then
            .Filter = "Archivo de texto DAT(*.dat)|*.dat"

        End If

        .CancelError = True
        .flags = cdlOFNFileMustExist
        .FileName = StringGuardadoActual(EstadoIndexador)
        .ShowOpen

    End With
 
    Select Case UCase$(Right$(CommonDialog1.FileName, 3))

        Case "DAT"
            Call BotonCargadoDat(CommonDialog1.FileName)

        Case "IND"
            Call BotonCargado(CommonDialog1.FileName)

        Case Else
            Exit Sub

    End Select

    Exit Sub
Cancelar:

End Sub

Private Sub MenuBotonGuardarP_Click()

    On Error GoTo Cancelar

    With CommonDialog1
        .Filter = "Binario(.ind)|*.ind|Archivo de texto DAT(*.dat)|*.dat"

        If EstadoIndexador = e_EstadoIndexador.Armas Or EstadoIndexador = e_EstadoIndexador.Escudos Then
            .Filter = "Archivo de texto DAT(*.dat)|*.dat"

        End If

        .CancelError = True
        .flags = cdlOFNOverwritePrompt
        .FileName = StringGuardadoActual(EstadoIndexador)
        .ShowSave

    End With
 
    Select Case UCase$(Right$(CommonDialog1.FileName, 3))

        Case "DAT"
            Call BotonGuardadoDat(CommonDialog1.FileName)

        Case "IND"
            Call BotonGuardado(CommonDialog1.FileName)

        Case Else
            Exit Sub

    End Select

    Exit Sub
Cancelar:

End Sub

Private Sub MenuEdicionBorrar_Click()
    Call SBotonBorrrar

End Sub

Private Sub MenuEdicionClonar_Click()
    Call SbotonClonar

End Sub

Private Sub menuEdicionColor_Click()

    With CommonDialog1
        .DialogTitle = "Seleccionar color para el fondo"
        .ShowColor

    End With

    ColorFondo = CommonDialog1.Color
    Call ClickEnLista

End Sub

Private Sub MenuEdicionCopiar_Click()
    Call SBotonMover(False)

End Sub

Private Sub menuEdicionMover_Click()
    Call SBotonMover(True)

End Sub

Public Sub SbotonClonar()
    Dim tempLong  As Long
    Dim cadena    As String
    Dim respuesta As Byte
    Dim LooPer    As Long
    Dim Inicial   As Long
    Dim Final     As Long
    Dim Origen    As Long

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh

            If GrHCambiando Then
                GrHCambiando = False
                LNumActual.Caption = "Ghr:"
                BotonGuardar.Visible = False

            End If

            If GRHActual < 0 Or GRHActual > grhCount Then Exit Sub
            cadena = InputBox("Introduzca el primer Numero de GHR al que quieres mover el grafico " & _
                    GRHActual, "Clonar Grafico")

            If IsNumeric(cadena) Then
                Inicial = Val(cadena)

                If Inicial > 0 And Inicial < grhCount Then
                    cadena = InputBox("Introduzca Cantidad de veces que quieres clonar el grafico " & _
                            GRHActual & " a partir de la posicion: " & Inicial, "Clonar Grafico")

                    If IsNumeric(cadena) Then
                        Final = Val(cadena) + Inicial

                        If Final > 0 And Final < grhCount Then
                            Origen = GRHActual

                            For LooPer = Inicial To Final
                                Call MoverGrh(LooPer, Origen, False)
                            Next LooPer

                            Call RenuevaListaGrH
                            tempLong = ListaindexGrH(Inicial)
                            frmMain.Lista.listIndex = tempLong
                            EstadoNoGuardado(e_EstadoIndexador.Grh) = True
                        Else
                            MsgBox "Fuera de los limites"

                        End If

                    End If

                Else
                    MsgBox "numero incorrecto"

                End If

            End If

            '    Case Else
            '        Dim StringCaso As String
            '        If EstadoIndexador = Body Then
            '            StringCaso = "Body"
            '        ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
            '            StringCaso = "Cabeza"
            '        ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
            '            StringCaso = "Casco"
            '        ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
            '            StringCaso = "Escudo"
            '        ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
            '            StringCaso = "Arma"
            '        ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
            '            StringCaso = "Bota"
            '        ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
            '            StringCaso = "Capa"
            '        ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
            '            StringCaso = "Fx"
            '        End If
            '        cadena = InputBox("Introduzca numero de " & StringCaso & " al que quieres mover", "Mover " & StringCaso)
            '        If IsNumeric(cadena) And (Val(cadena) < 31000) Then
            '            If EstadoIndexador = e_EstadoIndexador.Body Then
            '                Call mueveBody(Val(cadena), DataIndexActual, False)
            '            ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
            '                Call MueveCabeza(Val(cadena), DataIndexActual, False)
            '            ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
            '                Call MueveCasco(Val(cadena), DataIndexActual, False)
            '            ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
            '                Call MueveEscudo(Val(cadena), DataIndexActual, False)
            '            ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
            '                Call MueveArma(Val(cadena), DataIndexActual, False)
            '            ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
            '                Call MueveBota(Val(cadena), DataIndexActual, False)
            '            ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
            '                Call MueveCapa(Val(cadena), DataIndexActual, False)
            '            ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
            '                Call MueveFX(Val(cadena), DataIndexActual, False)
            '            End If
            '                DataIndexActual = Val(cadena)
            '                LOOPActual = 1
            '                frmMain.Visor.Cls
            '                Call GetInfoDataIndex(DataIndexActual)
            '                Dibujado.Interval = 200
            '                Dibujado.Enabled = True
            '                LGHRnumeroA.Caption = DataIndexActual
            '                templong = ListaindexGrH(DataIndexActual)
            '                frmMain.Lista.listIndex = templong
            '        Else
            '            MsgBox "introduzca un numero valido"
            '        End If
    End Select

End Sub

Private Sub MenuEdicionNuevo_Click()
    Call BotonNuevoGRH

End Sub

Private Sub MenuHerramientasAAnim_Click()

    If FormAuto.Visible Then
        FormAuto.SetFocus
    Else
        FormAuto.Show , frmMain

    End If

End Sub

Private Sub MenuHerramientasBG_Click()
    Dim i      As Long
    Dim cadena As String

    LastFound = 0
    cadena = InputBox("Introduzca número de Bmp a buscar", "Nuevo Grafico")

    If Val(cadena) > 0 And Val(cadena) <= grhCount Then
        BMPBuscado = Val(cadena)

        For i = 1 To grhCount

            If GrhData(i).FileNum = BMPBuscado Then
                Call BuscarNuevoF(i)
                LastFound = i
                MenuHerramientasBN.Visible = True
                LUlitError.Caption = "F3 para continuar la busqueda"
                Exit Sub

            End If

        Next i

        LUlitError.Caption = "BMP no encontrado"
        MenuHerramientasBN.Visible = False

    End If

End Sub

Private Sub MenuHerramientasBN_Click()
    Dim i As Long

    If LastFound = 0 Or BMPBuscado = 0 Then Exit Sub

    For i = LastFound + 1 To grhCount

        If GrhData(i).FileNum = BMPBuscado Then
            Call BuscarNuevoF(i)
            LastFound = i
            LUlitError.Caption = "F3 para continuar la busqueda"
            Exit Sub

        End If

    Next i

    LUlitError.Caption = " Se termino la busqueda"
    MenuHerramientasBN.Visible = False
    LastFound = 0
    BMPBuscado = 0

End Sub

Private Sub MenuHerramientasBR_Click()

    If FrmSearch.Visible Then
        FrmSearch.SetFocus
    Else
        FrmSearch.Show , frmMain

    End If

    Call FrmSearch.HacerBusquedaR

End Sub

Private Sub MenuHerramientasNI_Click()

    If FrmSearch.Visible Then
        FrmSearch.SetFocus
    Else
        FrmSearch.Show , frmMain

    End If

    Call FrmSearch.HacerBusquedaNI

End Sub

Private Sub MenuIndiceGuardado_Click(Index As Integer)
    Dim i As Long

    For i = 0 To 9
        MenuIndiceGuardado(i).Checked = False
    Next i

    MenuIndiceGuardado(Index).Checked = True
    SavePath = Index
    Call CambiarcaptionCommand10

End Sub

Private Sub mnIgeneral_Click()
    Dim textoActual As String
    Dim bmpstring   As String
    Dim BMPBuscado  As Long

    bmpstring = ReadField(1, Lista.List(Lista.listIndex), Asc(" "))

    If IsNumeric(bmpstring) Then
        BMPBuscado = Val(bmpstring)

        If BMPBuscado > 0 And BMPBuscado <= 32000 Then
            If FormAuto.Visible Then
                FormAuto.SetFocus
            Else
                FormAuto.Show , frmMain

            End If

            FormAuto.FrameAnim(0).Visible = False
            FormAuto.FrameAnim(1).Visible = False
            FormAuto.FrameAnim(2).Visible = True
            FormAuto.TextDatos3(4).Text = BMPBuscado
            FormAuto.TextDatos3(5).Text = BuscarGrHlibres(1)

        End If

    End If

End Sub

Private Sub mnuibody_Click()
    Dim textoActual As String
    Dim bmpstring   As String
    Dim BMPBuscado  As Long

    bmpstring = ReadField(1, Lista.List(Lista.listIndex), Asc(" "))

    If IsNumeric(bmpstring) Then
        BMPBuscado = Val(bmpstring)

        If BMPBuscado > 0 And BMPBuscado <= 32000 Then
            If FormAuto.Visible Then
                FormAuto.SetFocus
            Else
                FormAuto.Show , frmMain

            End If

            FormAuto.FrameAnim(1).Visible = True
            FormAuto.FrameAnim(0).Visible = False
            FormAuto.FrameAnim(2).Visible = False
            FormAuto.Combo2.Visible = False
            FormAuto.Labelbody.Visible = False
            FormAuto.Labelbody1.Visible = False
            FormAuto.Labelbody2.Visible = False
            FormAuto.Loff.Visible = False
            FormAuto.Loffx.Visible = False
            FormAuto.Loffy.Visible = False
            FormAuto.TextDatos2(7).Visible = False
            FormAuto.TextDatos2(8).Visible = False
            FormAuto.TextDatos2(0).Enabled = False
            FormAuto.TextDatos2(1).Enabled = False
            FormAuto.TextDatos2(6).Enabled = False
            FormAuto.Text1.Visible = False
            FormAuto.Text2.Visible = False
            FormAuto.CheckAuto.Visible = False
            FormAuto.Optiondimension(0).Visible = False
            FormAuto.Optiondimension(1).Visible = False
            FormAuto.Optiondimension(2).Visible = False
            FormAuto.Label5.Visible = False
            FormAuto.Label6.Visible = False
            FormAuto.TextDatos2(0).Text = 6
            FormAuto.TextDatos2(1).Text = 4
            FormAuto.TextDatos2(4).Text = BMPBuscado
            FormAuto.TextDatos2(6).Text = 22
            FormAuto.TextDatos2(2).Text = 46
            FormAuto.TextDatos2(3).Text = 26
            FormAuto.Optiondimension(0).value = True
            FormAuto.Labelbody.Visible = True
            FormAuto.Labelbody1.Visible = True
            FormAuto.Labelbody2.Visible = True
            FormAuto.Text1.Visible = True
            FormAuto.Text1.Enabled = False
            FormAuto.Text2.Visible = True
            FormAuto.Text2.Enabled = False
            FormAuto.CheckAuto.Visible = True
            FormAuto.CheckAuto.value = vbUnchecked
            FormAuto.Text1.Text = UBound(BodyData) + 1
            FormAuto.Text2.Text = "-38º0"
            FormAuto.Combo2.Visible = True
            FormAuto.Combo2.listIndex = 0
            FormAuto.Optiondimension(0).Visible = True
            FormAuto.Optiondimension(1).Visible = True
            FormAuto.Optiondimension(2).Visible = True
            FormAuto.Label5.Visible = True
            FormAuto.Label6.Visible = True

        End If

    End If

End Sub

Private Sub NuevoGhr_Click()
    Call BotonNuevoGRH

End Sub

Public Sub BotonNuevoGRH()
    Dim cadena    As String
    Dim respuesta As Byte
    Dim tempLong  As Long

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh

            If GrHCambiando Then
                GrHCambiando = False
                LNumActual.Caption = "Ghr:"
                BotonGuardar.Visible = False

            End If

            cadena = InputBox("Introduzca el número de GHR (0 Para encontrar un hueco libre)", _
                    "Nuevo Grafico")

            If IsNumeric(cadena) Then
                Call BuscarNuevoF(Val(cadena))
            Else
                LUlitError.Caption = "introduzca un numero"

            End If

        Case e_EstadoIndexador.Resource
            cadena = InputBox("Introduzca el número de BMP", "Nuevo Grafico")

            If IsNumeric(cadena) Then
                Call BuscarNuevoF(Val(cadena))
            Else
                LUlitError.Caption = "introduzca un numero"

            End If

            Exit Sub

        Case Else
            Dim StringCaso As String

            If EstadoIndexador = Body Then
                StringCaso = "Body"
            ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
                StringCaso = "Cabeza"
            ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
                StringCaso = "Casco"
            ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
                StringCaso = "Escudo"
            ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
                StringCaso = "Arma"
            ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
                StringCaso = "Bota"
            ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
                StringCaso = "Capa"
            ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
                StringCaso = "Fx"

            End If
        
            cadena = InputBox("Introduzca " & StringCaso & " (0 Para encontrar un hueco libre)", "Nuevo " & _
                    StringCaso & "/buscar")

            If IsNumeric(cadena) Then
                Call BuscarNuevoF(Val(cadena))
            Else
                LUlitError.Caption = "introduzca un numero"

            End If

    End Select

End Sub

Public Sub BuscarNuevoF(ByVal Index As Long)
    Dim cadena    As String
    Dim respuesta As Byte
    Dim tempLong  As Long

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh

            If GrHCambiando Then
                GrHCambiando = False
                LNumActual.Caption = "Ghr:"
                BotonGuardar.Visible = False

            End If

            If Index > 0 And Index < grhCount Then
                tempLong = ListaindexGrH(Index)

                If tempLong >= 0 Then
                    GRHActual = Index
                    LOOPActual = 1
                    frmMain.Visor.Cls
                    'Call DibujarGHRVisor(GRHActual)
                    'Call GetInfoGHR(GRHActual)
                    LGHRnumeroA.Caption = GRHActual
                    frmMain.Lista.listIndex = tempLong
                Else
                    respuesta = MsgBox("El grafico no existe,¿ quieres crearlo? ", 4, "Aviso")

                    If respuesta = vbYes Then
                        GRHActual = Index
                        'GrhData(GRHActual).NumFrames = 0
                        Call AgregaGrH(GRHActual)
                        LOOPActual = 1
                        frmMain.Visor.Cls
                        'Call DibujarGHRVisor(GRHActual)
                        Call GetInfoGHR(GRHActual)
                        LGHRnumeroA.Caption = GRHActual
                        EstadoNoGuardado(e_EstadoIndexador.Grh) = True

                    End If

                End If

            ElseIf Index = 0 Then
                GRHActual = BuscarGrHlibre()

                If GRHActual > 0 And GRHActual <= grhCount Then
                    Call AgregaGrH(GRHActual)
                    LOOPActual = 1
                    frmMain.Visor.Cls
                    'Call DibujarGHRVisor(GRHActual)
                    Call GetInfoGHR(GRHActual)
                    LGHRnumeroA.Caption = GRHActual
                Else
                    LUlitError.Caption = "No Se encontro hueco"

                End If

            Else
                LUlitError.Caption = "Valor no valido"

            End If

        Case e_EstadoIndexador.Resource

            If Index > 0 And Index <= 32000 Then
                tempLong = ListaindexGrH(Index)

                If tempLong >= 0 Then
                    GRHActual = Index
                    LOOPActual = 1
                    frmMain.Visor.Cls
                    'Call DibujarBMPVisor(GRHActual)
                    Call GetInfoBmp(GRHActual)
                    Call DibujarBMPVisor(GRHActual)
                    LGHRnumeroA.Caption = GRHActual
                    frmMain.Lista.listIndex = tempLong
                Else
                    LUlitError.Caption = "Bmp no existe"

                End If

            Else
                LUlitError.Caption = "Valor no valido"

            End If

            Exit Sub

        Case Else
            Dim StringCaso As String

            If EstadoIndexador = Body Then
                StringCaso = "Body"
            ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
                StringCaso = "Cabeza"
            ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
                StringCaso = "Casco"
            ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
                StringCaso = "Escudo"
            ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
                StringCaso = "Arma"
            ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
                StringCaso = "Bota"
            ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
                StringCaso = "Capa"
            ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
                StringCaso = "Fx"

            End If

            If Index > 0 And Index < grhCount Then
                tempLong = ListaindexGrH(Index)

                If tempLong >= 0 Then
                    DataIndexActual = Index
                    LOOPActual = 1
                    frmMain.Visor.Cls
                    Call GetInfoDataIndex(DataIndexActual)
                    Dibujado.Interval = 100
                    Dibujado.Enabled = True
                    LGHRnumeroA.Caption = DataIndexActual
                    Lista.listIndex = tempLong
                Else
                    respuesta = MsgBox("El " & StringCaso & " no existe,¿ quieres crearlo? ", 4, "Aviso")

                    If respuesta = vbYes Then
                        DataIndexActual = Index
                        LOOPActual = 1
                        frmMain.Visor.Cls
                        Dibujado.Interval = 100
                        Dibujado.Enabled = True
                        LGHRnumeroA.Caption = DataIndexActual

                        If EstadoIndexador = e_EstadoIndexador.Body Then
                            Call AgregaBody(DataIndexActual)
                        ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
                            Call AgregaCabeza(DataIndexActual)
                        ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
                            Call AgregaCasco(DataIndexActual)
                        ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
                            Call AgregaEscudo(DataIndexActual)
                        ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
                            Call AgregaArma(DataIndexActual)
                        ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
                            Call AgregaBota(DataIndexActual)
                        ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
                            Call AgregaCapa(DataIndexActual)
                        ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
                            Call AgregaFx(DataIndexActual)

                        End If

                        Call GetInfoDataIndex(DataIndexActual)
                        EstadoNoGuardado(EstadoIndexador) = True

                    End If

                End If

                '            ElseIf Val(cadena) = 0 Then
                '                GRHActual = BuscarGrHlibre()
                '                If GRHActual > 0 And GRHActual <= MAXGrH Then
                '                    Call AgregaGrH(GRHActual)
                '                    LOOPActual = 1
                '                    frmMain.Visor.Cls
                '                    Call DibujarGHRVisor(GRHActual)
                '                    Call GetInfoGHR(GRHActual)
                '                    LGHRnumeroA.Caption = GRHActual
                '                Else
                '                    MsgBox "No Se encontro hueco"
                '                End If
            Else
                LUlitError.Caption = "Valor no valido"

            End If

    End Select

End Sub

Private Sub Text1_Change()
    CarpetaDeInis = Text1.Text
    Call WriteVar(App.Path & "\Conf.ini", "Config", "CarpetaDeInis", CarpetaDeInis)

End Sub

Private Sub Text2_Change()
    CarpetaGraficos = Text2.Text
    Call WriteVar(App.Path & "\Conf.ini", "Config", "CarpetaGraficos", Text2.Text)

End Sub

Private Sub TextDatos_DblClick(Index As Integer)

    If EstadoIndexador = e_EstadoIndexador.Grh Or Index > 4 Or (EstadoIndexador = e_EstadoIndexador.Fx) And _
            Index > 0 Then Exit Sub

    If Val(TextDatos(Index).Text) > 0 And Val(TextDatos(Index).Text) < grhCount Then

        If EstadoIndexador <> e_EstadoIndexador.Grh Then Call CambiarEstado(e_EstadoIndexador.Grh)
        Call BuscarNuevoF(TextDatos(Index).Text)

    End If

End Sub

Private Sub TextDatos_Change(Index As Integer)
    'Comprueba que los datos introducidos son correctos

    Dim Ancho       As Long
    Dim Alto        As Long
    Dim PrimerAncho As Long
    Dim PrimerAlto  As Long
    Dim i           As Long
    Dim Algun_Error As Boolean
    Dim ErroresGrh  As ErroresGrh
    Dim tdouble1    As Double, tdouble2 As Double

    If EstadoIndexador = e_EstadoIndexador.Resource Then Exit Sub

2   For i = 0 To 7

        If i <> 1 And ((i <> 5) Or EstadoIndexador <> Body) And ((i <> 2) Or EstadoIndexador <> Fx) Then ' el 1 son los frames y el 5 se usa para offset
            If Val(TextDatos(i).Text) > grhCount Then
                TextDatos(i).Text = grhCount

            End If

        ElseIf ((i = 5) And EstadoIndexador = Body) Or ((i = 2) And EstadoIndexador = Fx) Then
            tdouble1 = Val(ReadField(1, TextDatos(i).Text, Asc("º")))
            tdouble2 = Val(ReadField(2, TextDatos(i).Text, Asc("º")))

            If tdouble1 < -32000 Or tdouble1 > grhCount Then
                TextDatos(i).Text = "0º" & tdouble2
                tdouble1 = 0

            End If
        
            If tdouble2 < -32000 Or tdouble2 > grhCount Then
                TextDatos(i).Text = tdouble1 & "º0"

            End If

        End If

        ErroresGrh.colores(i) = vbWhite
    Next i

    ErroresGrh.colores(8) = vbWhite
    ErroresGrh.colores(9) = vbWhite

    LUlitError.Caption = ""
    Dim resul        As Long
    Dim MensageError As String

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh

            If Not GrHCambiando Then
                GrHCambiando = True
                TempGrh = GrhData(GRHActual)
                LNumActual.Caption = "**Ghr:"
                BotonGuardar.Visible = True

            End If

            If Val(TextDatos(5).Text) > grhCount Then
                TextDatos(5).Text = grhCount

            End If

            If Val(TextDatos(2).Text) > 25 Then 'numframes > 25
                TextDatos(2).Text = 25
            ElseIf Val(TextDatos(2).Text) < 1 Then 'numframes < 1
                TextDatos(2).Text = 1

            End If
            
            If Val(TextDatos(2).Text) = 1 Then ' Es grh normal
                TextDatos(1).Enabled = False

                For i = 3 To 4
                    TextDatos(i).Enabled = True
                Next i

                TextDatos(5).Enabled = False

                For i = 6 To 7
                    TextDatos(i).Enabled = True
                Next i

            ElseIf Val(TextDatos(2).Text) > 1 Then ' es animacion
                TextDatos(1).Enabled = True

                For i = 3 To 4
                    TextDatos(i).Enabled = False
                Next i

                TextDatos(5).Enabled = True

                For i = 6 To 7
                    TextDatos(i).Enabled = False
                Next i

            End If

            TempGrh.FileNum = Val(TextDatos(0).Text)
            TempGrh.NumFrames = Val(TextDatos(2).Text)

            If TempGrh.NumFrames = 1 Then
                TempGrh.Frames(1) = Val(ReadField(1, TextDatos(1).Text, Asc("-")))
            Else
            
                ReDim TempGrh.Frames(1 To TempGrh.NumFrames) As Long

                For i = 1 To TempGrh.NumFrames

                    If Val(ReadField(i, TextDatos(1).Text, Asc("-"))) < grhCount Then
                        TempGrh.Frames(i) = Val(ReadField(i, TextDatos(1).Text, Asc("-")))

                    End If

                Next i

            End If

            TempGrh.pixelHeight = Val(TextDatos(3).Text)
            TempGrh.pixelWidth = Val(TextDatos(4).Text)
            TempGrh.Speed = Val(TextDatos(5).Text)
            TempGrh.sX = Val(TextDatos(6).Text)
            TempGrh.sY = Val(TextDatos(7).Text)
            TempGrh.TileHeight = GrhData(GRHActual).pixelHeight / TilePixelHeight
            TextDatos(8).Text = TempGrh.TileHeight
            TempGrh.TileWidth = GrhData(GRHActual).pixelWidth / TilePixelWidth
            TextDatos(9).Text = TempGrh.TileWidth

            resul = GrhCorrecto(TempGrh, MensageError, ErroresGrh)
            LUlitError.Caption = MensageError
            
            For i = 0 To 9
                TextDatos(i).BackColor = ErroresGrh.colores(i)
            Next i
            
            If ErroresGrh.ErrorCritico Then
                BotonGuardar.Visible = False
                Exit Sub
            Else
                BotonGuardar.Visible = True

            End If
            
            frmMain.Visor.Cls

            If Not LoadingNew Then Call DibujarGHRVisor(GRHActual)
            If TempGrh.NumFrames = 1 Then
                frmMain.Dibujado.Enabled = False
            ElseIf TempGrh.NumFrames > 1 Then

                If TempGrh.Speed > 0 Then ' pervenimos division por 0
                    frmMain.Dibujado.Interval = 50 '* TempGrh.Speed
                Else
                    frmMain.Dibujado.Interval = 100

                End If

                frmMain.Dibujado.Enabled = True
            Else
                frmMain.Dibujado.Enabled = False

            End If

        Case Else

            If Not GrHCambiando Then
                GrHCambiando = True
                BotonGuardar.Visible = True

            End If
        
            If Not LoadingNew Then frmMain.Visor.Cls ' Si no estamos cargando limpiamos
            
            Dibujado.Interval = 100
            Dibujado.Enabled = True

            If EstadoIndexador = e_EstadoIndexador.Body Then
                tempDataIndex.HeadOffset.Y = Val(ReadField(1, TextDatos(5).Text, Asc("º")))
                tempDataIndex.HeadOffset.X = Val(ReadField(2, TextDatos(5).Text, Asc("º")))

            End If

            Dim III  As Long
            Dim Tstr As String
            Algun_Error = False

            For i = 1 To 4

                If i = 1 Then
                    III = 0
                Else
                    III = i

                End If

                If i = 1 Then
                    If EstadoIndexador = e_EstadoIndexador.Fx Then
                        Tstr = "FX"
                    Else
                        Tstr = "Subir"

                    End If

                ElseIf i = 2 Then
                    Tstr = "Derecha"
                ElseIf i = 3 Then
                    Tstr = "Abajo"
                ElseIf i = 4 Then
                    Tstr = "Izquierda"

                End If

                If i = 1 Or EstadoIndexador <> e_EstadoIndexador.Fx Then
                    tempDataIndex.Walk(i).GrhIndex = Val(TextDatos(III).Text)

                    If tempDataIndex.Walk(i).GrhIndex > 1 Then
                        MensageError = ""
                        resul = GrhCorrecto(GrhData(tempDataIndex.Walk(i).GrhIndex), MensageError, ErroresGrh)

                        If ErroresGrh.ErrorCritico Then
                            Algun_Error = True
                            TextDatos(III).BackColor = vbRed
                            LUlitError.Caption = LUlitError.Caption & "(" & Tstr & ") " & MensageError & _
                                    vbCrLf
                        Else

                            If EstadoIndexador = e_EstadoIndexador.Cabezas Or EstadoIndexador = _
                                    e_EstadoIndexador.Cascos Then

                                If ErroresGrh.EsAnimacion Then
                                    TextDatos(III).BackColor = vbYellow
                                    LUlitError.Caption = LUlitError.Caption & "(" & Tstr & _
                                            ") Es una animacion" & vbCrLf
                                Else
                                    TextDatos(III).BackColor = vbWhite

                                End If

                            Else

                                If Not ErroresGrh.EsAnimacion Then
                                    TextDatos(III).BackColor = vbYellow
                                    LUlitError.Caption = LUlitError.Caption & "(" & Tstr & _
                                            ") No es una animacion" & vbCrLf
                                Else
                                    TextDatos(III).BackColor = vbWhite

                                End If

                            End If

                        End If

                    End If

                End If

            Next i

            If Algun_Error Then
                BotonGuardar.Visible = False
            Else
                BotonGuardar.Visible = True

            End If
        
    End Select

End Sub

