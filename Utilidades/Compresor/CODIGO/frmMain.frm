VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Compresor de recursos graficos"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3900
   ClipControls    =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extraer Iconos"
      Height          =   300
      Index           =   7
      Left            =   1935
      TabIndex        =   21
      Top             =   3420
      Width           =   1800
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Comprimir Iconos"
      Height          =   300
      Index           =   7
      Left            =   120
      TabIndex        =   20
      Top             =   3420
      Width           =   1815
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extraer MiniMapa"
      Height          =   300
      Index           =   6
      Left            =   1920
      TabIndex        =   19
      Top             =   3120
      Width           =   1770
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Comprimir MiniMapa"
      Height          =   315
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extraer Interfaces"
      Height          =   375
      Index           =   5
      Left            =   1920
      TabIndex        =   17
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Comprimir Interfaces"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Comprimir MIDI"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extraer MIDI"
      Height          =   375
      Index           =   4
      Left            =   1920
      TabIndex        =   13
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extraer WAV"
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   12
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Comprimir WAV"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Comprimir MAPAS"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extraer MAPAS"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   9
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extraer INIT"
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Comprimir INIT"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Frame StatusFrame 
      Caption         =   "StatusFrame"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   3615
      Begin MSComctlLib.ProgressBar StatusBar 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.CommandButton cmdPatch 
      Caption         =   "Parchear"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extraer"
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtVersion 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Comprimir"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Working Version :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eFolderPath

    Graficos = 0
    Mapas = 1
    Init = 2
    Wav = 3
    Midi = 4
    Interfaces = 5
    MiniMapa = 6
    Iconos = 7

End Enum

Private OutPutNameFile()      As String
Private FileExtension()       As String

Private Const GRAPHIC_PATH    As String = "\GRAFICOS\"
Private Const INIT_PATH       As String = "\INIT\"
Private Const MAP_PATH        As String = "\MAPAS\"
Private Const WAV_PATH        As String = "\WAV\"
Private Const MIDI_PATH       As String = "\MIDI\"
Private Const INTER_PATH      As String = "\INTERFACES\"
Private Const MINI_PATH        As String = "\MINIMAPA\"
Private Const ICON_PATH       As String = "\ICONOS\"

Private Const COMPRESS_PATH   As String = "\COMPRIMIR"
Private Const DECOMPRESS_PATH As String = "\DESCOMPRIMIR"

Private Const RESOURCE_PATH   As String = "\RECURSOS\"
Private Const PATCH_PATH      As String = "\PARCHES\"
Private Const EXTRACT_PATH    As String = "\EXTRACCIONES\"

Private Sub Form_Load()

    ReDim OutPutNameFile(eFolderPath.Graficos To eFolderPath.Iconos) As String
    ReDim FileExtension(eFolderPath.Graficos To eFolderPath.Iconos) As String

    OutPutNameFile(eFolderPath.Graficos) = modCompression.GRH_RESOURCE_FILE
    OutPutNameFile(eFolderPath.Mapas) = modCompression.MAPAS_RESOURCE_FILE
    OutPutNameFile(eFolderPath.Init) = modCompression.INIT_RESOURCE_FILE
    OutPutNameFile(eFolderPath.Wav) = modCompression.WAV_RESOURCE_FILE
    OutPutNameFile(eFolderPath.Midi) = modCompression.MIDI_RESOURCE_FILE
    OutPutNameFile(eFolderPath.Interfaces) = modCompression.INT_RESOURCE_FILE
    OutPutNameFile(eFolderPath.MiniMapa) = modCompression.MINIMAPA_FILE
    OutPutNameFile(eFolderPath.Iconos) = modCompression.ICON_FILE
    
    FileExtension(eFolderPath.Graficos) = ".BMP"
    FileExtension(eFolderPath.Mapas) = ".MAP"
    FileExtension(eFolderPath.Init) = ".*"
    FileExtension(eFolderPath.Wav) = ".WAV"
    FileExtension(eFolderPath.Midi) = ".MID"
    FileExtension(eFolderPath.Interfaces) = ".JPG"
    FileExtension(eFolderPath.MiniMapa) = ".BMP"
    FileExtension(eFolderPath.Iconos) = ".ICO"
    
    Call InitializeCompression

End Sub

Private Function GetSourcePathByIndex(ByVal Index As Integer) As String

    Select Case Index

        Case eFolderPath.Graficos
            GetSourcePathByIndex = GRAPHIC_PATH

        Case eFolderPath.Mapas
            GetSourcePathByIndex = MAP_PATH

        Case eFolderPath.Init
            GetSourcePathByIndex = INIT_PATH

        Case eFolderPath.Wav
            GetSourcePathByIndex = WAV_PATH

        Case eFolderPath.Midi
            GetSourcePathByIndex = MIDI_PATH

        Case eFolderPath.Interfaces
            GetSourcePathByIndex = INTER_PATH
        
        Case eFolderPath.MiniMapa
            GetSourcePathByIndex = MINI_PATH
        
        Case eFolderPath.Iconos
            GetSourcePathByIndex = ICON_PATH
            
    End Select

End Function

Private Sub cmdCompress_Click(Index As Integer)

    Dim SourcePath As String
    Dim OutputPath As String

    SourcePath = App.Path & COMPRESS_PATH & GetSourcePathByIndex(Index)
    OutputPath = App.Path & COMPRESS_PATH & RESOURCE_PATH & txtVersion.Text & "\"

    'Check if the version already exists
    If FileExist(OutputPath & OutPutNameFile(Index), vbNormal) Then
        If MsgBox("La versión ya se encuentra comprimida. ¿Desea reemplazarla?", vbYesNo) = vbNo Then Exit Sub
    Else

        'Create this version folder
        If Not FileExist(OutputPath, vbDirectory) Then Call MkDir(OutputPath)

    End If

    'Show status
    StatusFrame.Caption = "Comprimiendo..."

    'Compress!
    If Compress_Files(SourcePath, OutputPath, txtVersion.Text, StatusBar, OutPutNameFile(Index), FileExtension(Index)) Then
        'Show we finished
        Call MsgBox("Operación terminada con éxito")
    Else
        'Show we finished
        Call MsgBox("Operación abortada")

    End If

End Sub

Private Sub cmdExtract_Click(Index As Integer)

    Dim ResourcePath As String
    Dim OutputPath   As String

    ResourcePath = App.Path & DECOMPRESS_PATH & RESOURCE_PATH & txtVersion.Text & "\"
    OutputPath = App.Path & DECOMPRESS_PATH & EXTRACT_PATH & txtVersion.Text

    'Create this version folder
    If Not FileExist(OutputPath, vbDirectory) Then Call MkDir(OutputPath)
    
    OutputPath = OutputPath & GetSourcePathByIndex(Index)
    
    Debug.Print OutputPath

    'Check if the resource file exists
    If Not FileExist(ResourcePath & OutPutNameFile(Index), vbNormal) Then
        MsgBox "No se encontraron los recursos a extraer." & vbCrLf & ResourcePath, , "Error"
        Exit Sub

    End If

    'Check if the version is already extracted
    If FileExist(OutputPath, vbDirectory) Then
        If MsgBox("La versión ya se encuentra extraida. ¿Desea reextraerla?", vbYesNo, "Atencion") = vbNo Then Exit Sub
    Else
        'Create this version folder
        Call MkDir(OutputPath)

    End If

    'Show the status bar
    StatusFrame.Caption = "Extrayendo..."

    'Extract!
    If Extract_Files(ResourcePath, OutputPath, StatusBar, OutPutNameFile(Index)) Then
        'Show we finished
        Call MsgBox("Operación terminada con éxito")
    Else
        'Show we finished
        Call MsgBox("Operación abortada")

    End If

End Sub

Private Sub Command9_Click()
    '**************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 12/02/2007
    'Loads all surfaces in random order and then sorts them
    '**************************************************************
    
    Dim SurfaceIndex As Long
    Dim bmpInfo      As BITMAPINFO
    Dim Data()       As Byte
    Dim I            As Long
    Dim ResourcePath As String
    
    ResourcePath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"

    While GetNext_Bitmap(ResourcePath, I, bmpInfo, Data(), SurfaceIndex)

        DoEvents
        
    Wend

    Debug.Print "Listo."

End Sub

'Private Sub cmdPatch_Click()
'    Dim NewResourcePath As String
'    Dim OldResourcePath As String
'    Dim OutputPath      As String
'
'    Dim NewVersion      As Long
'    Dim OldVersion      As Long
'
'    NewVersion = CLng(txtVersion.Text)
'    OldVersion = NewVersion - 1 'we patch from the last version
'
'    NewResourcePath = App.Path & RESOURCE_PATH & NewVersion & "\"
'    OldResourcePath = App.Path & RESOURCE_PATH & OldVersion & "\"
'    OutputPath = App.Path & PATCH_PATH & OldVersion & " to " & NewVersion & "\"
'
'    'Check if the new resource file exists
'    If Not FileExist(NewResourcePath & GRH_RESOURCE_FILE, vbNormal) Then
'        MsgBox "No se encontraron los recursos de la version actual." & vbCrLf & NewResourcePath, , "Error"
'        Exit Sub
'
'    End If
'
'    'Check if the old resource file exists
'    If Not FileExist(OldResourcePath & GRH_RESOURCE_FILE, vbNormal) Then
'        MsgBox "No se encontraron los recursos de la version anterior." & vbCrLf & OldResourcePath, , "Error"
'        Exit Sub
'
'    End If
'
'    'Check if the version is already extracted
'    If FileExist(OutputPath, vbDirectory) Then
'        If MsgBox("El parche ya se encuentra realizado. ¿Desea reparchear?", vbYesNo, "Atencion") = vbNo Then Exit Sub
'    Else
'        'Create this version folder
'        MkDir OutputPath
'
'    End If
'
'    'Show the status bar
'    Me.Height = 2880
'    StatusFrame.Caption = "Armando el parche de " & OldVersion & " a " & NewVersion
'
'    'Patch!
'    If Make_Patch(NewResourcePath, OldResourcePath, OutputPath, StatusBar) Then
'        'Show we finished
'        MsgBox "Operación terminada con éxito"
'    Else
'        'Show we finished
'        MsgBox "Operación abortada"
'
'    End If
'
'    'Hide status
'    Me.Height = 2055
'
'End Sub
'
'Private Sub cmdCompress_Click()
'    Dim SourcePath As String
'    Dim OutputPath As String
'
'    SourcePath = App.Path & INIT_PATH
'    OutputPath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
'
'    'Check if the version already exists
'    If FileExist(OutputPath & INIT_RESOURCE_FILE, vbNormal) Then
'        If MsgBox("La versión ya se encuentra comprimida. ¿Desea reemplazarla?", vbYesNo, "Atencion") = vbNo Then Exit Sub
'    Else
'
'        If Not FileExist(OutputPath, vbDirectory) Then
'            'Create this version folder
'            MkDir OutputPath
'
'        End If
'
'    End If
'
'    'Show status
'    Me.Height = 2880
'    StatusFrame.Caption = "Comprimiendo..."
'
'    'Compress!
'    If Compress_Files(SourcePath, OutputPath, txtVersion.Text, StatusBar, INIT_RESOURCE_FILE, ".*") Then
'        'Show we finished
'        MsgBox "Operación terminada con éxito"
'    Else
'        'Show we finished
'        MsgBox "Operación abortada"
'
'    End If
'
'    'Hide status
'    Me.Height = 2055
'
'End Sub
'
'Private Sub cmdExtract_Click()
'    Dim ResourcePath As String
'    Dim OutputPath   As String
'
'    ResourcePath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
'    OutputPath = App.Path & EXTRACT_PATH & txtVersion.Text & "-Init\"
'
'    'Check if the resource file exists
'    If Not FileExist(ResourcePath & INIT_RESOURCE_FILE, vbNormal) Then
'        MsgBox "No se encontraron los recursos a extraer." & vbCrLf & ResourcePath, , "Error"
'        Exit Sub
'
'    End If
'
'    'Check if the version is already extracted
'    If FileExist(OutputPath, vbDirectory) Then
'        If MsgBox("La versión ya se encuentra extraida. ¿Desea reextraerla?", vbYesNo, "Atencion") = vbNo Then Exit Sub
'    Else
'        'Create this version folder
'        MkDir OutputPath
'
'    End If
'
'    'Show the status bar
'    Me.Height = 2880
'    StatusFrame.Caption = "Extrayendo..."
'
'    'Extract!
'    If Extract_Files(ResourcePath, OutputPath, StatusBar, INIT_RESOURCE_FILE) Then
'        'Show we finished
'        MsgBox "Operación terminada con éxito"
'    Else
'        'Show we finished
'        MsgBox "Operación abortada"
'
'    End If
'
'    'Hide status
'    Me.Height = 2055
'
'End Sub
'
'Private Sub cmdExtract_Click()
'    Dim ResourcePath As String
'    Dim OutputPath   As String
'
'    ResourcePath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
'    OutputPath = App.Path & EXTRACT_PATH & txtVersion.Text & "-Mapas\"
'
'    'Check if the resource file exists
'    If Not FileExist(ResourcePath & MIDI_RESOURCE_FILE, vbNormal) Then
'        MsgBox "No se encontraron los recursos a extraer." & vbCrLf & ResourcePath, , "Error"
'        Exit Sub
'
'    End If
'
'    'Check if the version is already extracted
'    If FileExist(OutputPath, vbDirectory) Then
'        If MsgBox("La versión ya se encuentra extraida. ¿Desea reextraerla?", vbYesNo, "Atencion") = vbNo Then Exit Sub
'    Else
'        'Create this version folder
'        MkDir OutputPath
'
'    End If
'
'    'Show the status bar
'    Me.Height = 2880
'    StatusFrame.Caption = "Extrayendo..."
'
'    'Extract!
'    If Extract_Files(ResourcePath, OutputPath, StatusBar, MAPAS_RESOURCE_FILE) Then
'        'Show we finished
'        MsgBox "Operación terminada con éxito"
'    Else
'        'Show we finished
'        MsgBox "Operación abortada"
'
'    End If
'
'    'Hide status
'    Me.Height = 2055
'
'End Sub
'
'Private Sub cmdCompress_Click()
'    Dim SourcePath As String
'    Dim OutputPath As String
'
'    SourcePath = App.Path & MAP_PATH
'    OutputPath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
'
'    'Check if the version already exists
'    If FileExist(OutputPath & MAPAS_RESOURCE_FILE, vbNormal) Then
'        If MsgBox("La versión ya se encuentra comprimida. ¿Desea reemplazarla?", vbYesNo, "Atencion") = vbNo Then Exit Sub
'    Else
'
'        If Not FileExist(OutputPath, vbDirectory) Then
'            'Create this version folder
'            MkDir OutputPath
'
'        End If
'
'    End If
'
'    'Show status
'    Me.Height = 2880
'    StatusFrame.Caption = "Comprimiendo..."
'
'    'Compress!
'    If Compress_Files(SourcePath, OutputPath, txtVersion.Text, StatusBar, MAPAS_RESOURCE_FILE, ".MAP") Then
'        'Show we finished
'        MsgBox "Operación terminada con éxito"
'    Else
'        'Show we finished
'        MsgBox "Operación abortada"
'
'    End If
'
'    'Hide status
'    Me.Height = 2055
'
'End Sub
'
'Private Sub cmdCompress_Click()
'    Dim SourcePath As String
'    Dim OutputPath As String
'
'    SourcePath = App.Path & WAV_PATH
'    OutputPath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
'
'    'Check if the version already exists
'    If FileExist(OutputPath & WAV_RESOURCE_FILE, vbNormal) Then
'        If MsgBox("La versión ya se encuentra comprimida. ¿Desea reemplazarla?", vbYesNo, "Atencion") = vbNo Then Exit Sub
'    Else
'
'        If Not FileExist(OutputPath, vbDirectory) Then
'            'Create this version folder
'            MkDir OutputPath
'
'        End If
'
'    End If
'
'    'Show status
'    Me.Height = 2880
'    StatusFrame.Caption = "Comprimiendo..."
'
'    'Compress!
'    If Compress_Files(SourcePath, OutputPath, txtVersion.Text, StatusBar, WAV_RESOURCE_FILE, ".WAV") Then
'        'Show we finished
'        MsgBox "Operación terminada con éxito"
'    Else
'        'Show we finished
'        MsgBox "Operación abortada"
'
'    End If
'
'    'Hide status
'    Me.Height = 2055
'
'End Sub
'
'Private Sub cmdExtract_Click()
'    Dim ResourcePath As String
'    Dim OutputPath   As String
'
'    ResourcePath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
'    OutputPath = App.Path & EXTRACT_PATH & txtVersion.Text & "-WAV\"
'
'    'Check if the resource file exists
'    If Not FileExist(ResourcePath & WAV_RESOURCE_FILE, vbNormal) Then
'        MsgBox "No se encontraron los recursos a extraer." & vbCrLf & ResourcePath, , "Error"
'        Exit Sub
'
'    End If
'
'    'Check if the version is already extracted
'    If FileExist(OutputPath, vbDirectory) Then
'        If MsgBox("La versión ya se encuentra extraida. ¿Desea reextraerla?", vbYesNo, "Atencion") = vbNo Then Exit Sub
'    Else
'        'Create this version folder
'        MkDir OutputPath
'
'    End If
'
'    'Show the status bar
'    Me.Height = 2880
'    StatusFrame.Caption = "Extrayendo..."
'
'    'Extract!
'    If Extract_Files(ResourcePath, OutputPath, StatusBar, WAV_RESOURCE_FILE) Then
'        'Show we finished
'        MsgBox "Operación terminada con éxito"
'    Else
'        'Show we finished
'        MsgBox "Operación abortada"
'
'    End If
'
'    'Hide status
'    Me.Height = 2055
'
'End Sub
'
'Private Sub cmdExtract_Click()
'    Dim ResourcePath As String
'    Dim OutputPath   As String
'
'    ResourcePath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
'    OutputPath = App.Path & EXTRACT_PATH & txtVersion.Text & "-Midi\"
'
'    'Check if the resource file exists
'    If Not FileExist(ResourcePath & MIDI_RESOURCE_FILE, vbNormal) Then
'        MsgBox "No se encontraron los recursos a extraer." & vbCrLf & ResourcePath, , "Error"
'        Exit Sub
'
'    End If
'
'    'Check if the version is already extracted
'    If FileExist(OutputPath, vbDirectory) Then
'        If MsgBox("La versión ya se encuentra extraida. ¿Desea reextraerla?", vbYesNo, "Atencion") = vbNo Then Exit Sub
'    Else
'        'Create this version folder
'        MkDir OutputPath
'
'    End If
'
'    'Show the status bar
'    Me.Height = 2880
'    StatusFrame.Caption = "Extrayendo..."
'
'    'Extract!
'    If Extract_Files(ResourcePath, OutputPath, StatusBar, MIDI_RESOURCE_FILE) Then
'        'Show we finished
'        MsgBox "Operación terminada con éxito"
'    Else
'        'Show we finished
'        MsgBox "Operación abortada"
'
'    End If
'
'    'Hide status
'    Me.Height = 2055
'
'End Sub
'
'Private Sub cmdCompress_Click()
'    Dim SourcePath As String
'    Dim OutputPath As String
'
'    SourcePath = App.Path & MIDI_PATH
'    OutputPath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
'
'    'Check if the version already exists
'    If FileExist(OutputPath & MIDI_RESOURCE_FILE, vbNormal) Then
'        If MsgBox("La versión ya se encuentra comprimida. ¿Desea reemplazarla?", vbYesNo, "Atencion") = vbNo Then Exit Sub
'    Else
'
'        If Not FileExist(OutputPath, vbDirectory) Then
'            'Create this version folder
'            MkDir OutputPath
'
'        End If
'
'    End If
'
'    'Show status
'    Me.Height = 2880
'    StatusFrame.Caption = "Comprimiendo..."
'
'    'Compress!
'    If Compress_Files(SourcePath, OutputPath, txtVersion.Text, StatusBar, MIDI_RESOURCE_FILE, ".mid") Then
'        'Show we finished
'        MsgBox "Operación terminada con éxito"
'    Else
'        'Show we finished
'        MsgBox "Operación abortada"
'
'    End If
'
'    'Hide status
'    Me.Height = 2055
'
'End Sub
'

