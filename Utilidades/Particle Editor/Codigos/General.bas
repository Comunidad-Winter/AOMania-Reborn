Attribute VB_Name = "Mod_General"
Option Explicit

Public Function DirGraficos() As String
    DirGraficos = App.Path & "\GRAFICOS\"

End Function

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, _
                     ByVal Text As String, _
                     Optional ByVal Red As Integer = -1, _
                     Optional ByVal Green As Integer, _
                     Optional ByVal Blue As Integer, _
                     Optional ByVal Bold As Boolean = False, _
                     Optional ByVal Italic As Boolean = False, _
                     Optional ByVal bCrLf As Boolean = False)

    '******************************************
    'Adds text to a Richtext box at the bottom.
    'Automatically scrolls to new text.
    'Text box MUST be multiline and have a 3D
    'apperance!
    '******************************************
    With RichTextBox

        If (Len(.Text)) > 10000 Then .Text = ""
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        
        .SelBold = Bold
        .SelItalic = Italic
        
        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
        RichTextBox.Refresh

    End With

End Sub

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound

End Function

Sub UnloadAllForms()

    On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms

        Unload mifrm
    Next

End Sub

Sub Main()

    On Error Resume Next

    ChDrive App.Path
    ChDir App.Path
    
    AddtoRichTextBox frmCargando.Status, "Cargando Engine Grafico...."
    
    'Por default usamos el dinámico
    Set SurfaceDB = New clsSurfaceManDynDX8
    
    frmCargando.Show
    frmCargando.Refresh
    
    AddtoRichTextBox frmCargando.Status, "Engine Grafico OK"
    LoadGrhData
    Call engine.Engine_Init
    Call CargarParticulas
    AddtoRichTextBox frmCargando.Status, "¡¡Bienvenido al Editor de Particulas NeoAO!!"
    Unload frmCargando
                   
    frmMain.Show

    'Inicialización de variables globales
    prgRun = True
    Dim pausa
    pausa = False
    
    engine.Start
    
    Exit Sub

End Sub

Public Function General_Particle_Create(ByVal ParticulaInd As Long, _
                                        ByVal x As Integer, _
                                        ByVal Y As Integer, _
                                        Optional ByVal particle_life As Long = 0) As Long

    Dim rgb_list(0 To 3) As Long
    rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, _
            StreamData(ParticulaInd).colortint(0).B)
    rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, _
            StreamData(ParticulaInd).colortint(1).B)
    rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, _
            StreamData(ParticulaInd).colortint(2).B)
    rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, _
            StreamData(ParticulaInd).colortint(3).B)

    General_Particle_Create = engine.Particle_Group_Create(x, Y, StreamData(ParticulaInd).grh_list, _
            rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, StreamData( _
            ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, _
            particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1, StreamData( _
            ParticulaInd).y1, StreamData(ParticulaInd).angle, StreamData(ParticulaInd).vecx1, StreamData( _
            ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, StreamData( _
            ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, _
            StreamData(ParticulaInd).spin_speedL, StreamData(ParticulaInd).gravity, StreamData( _
            ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData( _
            ParticulaInd).x2, StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData( _
            ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
            StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData( _
            ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin)

End Function

Sub CargarParticulas()
    Dim StreamFile As String
    Dim loopc      As Long
    Dim i          As Long
    Dim GrhListing As String
    Dim TempSet    As String
    Dim ColorSet   As Long
    
    StreamFile = App.Path & "\INIT\Particles.ini"
    TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))

    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream

    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).Name = General_Var_Get(StreamFile, Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = General_Var_Get(StreamFile, Val(loopc), "NumOfParticles")
        StreamData(loopc).x1 = General_Var_Get(StreamFile, Val(loopc), "X1")
        StreamData(loopc).y1 = General_Var_Get(StreamFile, Val(loopc), "Y1")
        StreamData(loopc).x2 = General_Var_Get(StreamFile, Val(loopc), "X2")
        StreamData(loopc).y2 = General_Var_Get(StreamFile, Val(loopc), "Y2")
        StreamData(loopc).angle = General_Var_Get(StreamFile, Val(loopc), "Angle")
        StreamData(loopc).vecx1 = General_Var_Get(StreamFile, Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = General_Var_Get(StreamFile, Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = General_Var_Get(StreamFile, Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = General_Var_Get(StreamFile, Val(loopc), "VecY2")
        StreamData(loopc).life1 = General_Var_Get(StreamFile, Val(loopc), "Life1")
        StreamData(loopc).life2 = General_Var_Get(StreamFile, Val(loopc), "Life2")
        StreamData(loopc).friction = General_Var_Get(StreamFile, Val(loopc), "Friction")
        StreamData(loopc).spin = General_Var_Get(StreamFile, Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = General_Var_Get(StreamFile, Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = General_Var_Get(StreamFile, Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = General_Var_Get(StreamFile, Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = General_Var_Get(StreamFile, Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = General_Var_Get(StreamFile, Val(loopc), "XMove")
        StreamData(loopc).YMove = General_Var_Get(StreamFile, Val(loopc), "YMove")
        StreamData(loopc).move_x1 = General_Var_Get(StreamFile, Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = General_Var_Get(StreamFile, Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = General_Var_Get(StreamFile, Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = General_Var_Get(StreamFile, Val(loopc), "move_y2")
        StreamData(loopc).life_counter = General_Var_Get(StreamFile, Val(loopc), "life_counter")
        StreamData(loopc).speed = Val(General_Var_Get(StreamFile, Val(loopc), "Speed"))
        StreamData(loopc).grh_resize = Val(General_Var_Get(StreamFile, Val(loopc), "resize"))
        StreamData(loopc).grh_resizex = Val(General_Var_Get(StreamFile, Val(loopc), "rx"))
        StreamData(loopc).grh_resizey = Val(General_Var_Get(StreamFile, Val(loopc), "ry"))
        StreamData(loopc).NumGrhs = General_Var_Get(StreamFile, Val(loopc), "NumGrhs")
        
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = General_Field_Read(Str(i), GrhListing, 44)
        Next i

        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)

        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).g = General_Field_Read(2, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).B = General_Field_Read(3, TempSet, 44)
        Next ColorSet

        frmMain.List2.AddItem loopc & " - " & StreamData(loopc).Name
    Next loopc

End Sub

Public Function General_Random_Number(ByVal LowerBound As Long, _
                                      ByVal UpperBound As Long) As Single
    '*****************************************************************
    'Author: Aaron Perkins
    'Find a Random number between a range
    '*****************************************************************
    Randomize Timer
    General_Random_Number = (UpperBound - LowerBound + 1) * Rnd + LowerBound

End Function

Public Sub General_Var_Write(ByVal file As String, _
                             ByVal Main As String, _
                             ByVal var As String, _
                             ByVal value As String)
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Writes a var to a text file
    '*****************************************************************
    writeprivateprofilestring Main, var, value, file

End Sub

Public Function General_Var_Get(ByVal file As String, _
                                ByVal Main As String, _
                                ByVal var As String) As String
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Get a var to from a text file
    '*****************************************************************
    Dim l        As Long
    Dim Char     As String
    Dim sSpaces  As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
    
    szReturn = ""
    
    sSpaces = Space$(5000)
    
    getprivateprofilestring Main, var, szReturn, sSpaces, Len(sSpaces), file
    
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = Left$(General_Var_Get, Len(General_Var_Get) - 1)

End Function

Public Function General_Field_Read(ByVal field_pos As Long, _
                                   ByVal Text As String, _
                                   ByVal delimiter As Byte) As String
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Gets a field from a delimited string
    '*****************************************************************
    Dim i        As Long
    Dim LastPos  As Long
    Dim FieldNum As Long
    
    LastPos = 0
    FieldNum = 0

    For i = 1 To Len(Text)

        If delimiter = CByte(Asc(Mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1

            If FieldNum = field_pos Then
                General_Field_Read = Mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr$(delimiter), _
                        vbTextCompare) - 1) - (LastPos))
                Exit Function

            End If

            LastPos = i

        End If

    Next i

    FieldNum = FieldNum + 1

    If FieldNum = field_pos Then
        General_Field_Read = Mid$(Text, LastPos + 1)

    End If

End Function

Public Function General_File_Exists(ByVal file_path As String, _
                                    ByVal file_type As VbFileAttribute) As Boolean

    If Dir(file_path, file_type) = "" Then
        General_File_Exists = False
    Else
        General_File_Exists = True

    End If

End Function

Public Sub HookSurfaceHwnd(pic As Form)
    Call ReleaseCapture
    Call SendMessage(pic.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

End Sub

Function LoadGrhData() As Boolean

    On Error GoTo ErrorHandler

    Dim Grh         As Long
    Dim Frame       As Long
    Dim GrhCount    As Long
    Dim handle      As Integer
    Dim FileVersion As Long
    
    'Open files
    handle = FreeFile()
    
    Open App.Path & "\INIT\Graficos.ind" For Binary Access Read As handle
    Seek #1, 1
    
    'Get file version
    Get handle, , FileVersion
    
    'Get number of grhs
    Get handle, , GrhCount
    
    'Resize arrays
    ReDim GrhData(1 To GrhCount) As GrhData
    
    While Not EOF(handle)

        Get handle, , Grh
        
        If Grh <> 0 Then

            With GrhData(Grh)
                'Get number of frames
                Get handle, , .NumFrames

                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                ReDim .Frames(1 To GrhData(Grh).NumFrames)
                
                If .NumFrames > 1 Then

                    'Read a animation GRH set
                    For Frame = 1 To .NumFrames
                        Get handle, , .Frames(Frame)

                        If .Frames(Frame) <= 0 Or .Frames(Frame) > GrhCount Then
                            GoTo ErrorHandler

                        End If

                    Next Frame
                    
                    Get handle, , .speed
                    
                    If .speed <= 0 Then GoTo ErrorHandler
                    
                    'Compute width and height
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight

                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth

                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = GrhData(.Frames(1)).TileWidth

                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileHeight = GrhData(.Frames(1)).TileHeight

                    If .TileHeight <= 0 Then GoTo ErrorHandler
                Else
                    'Read in normal GRH data
                    Get handle, , .FileNum

                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , GrhData(Grh).sX

                    If .sX < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .sY

                    If .sY < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelWidth

                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelHeight

                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    'Compute width and height
                    .TileWidth = .pixelWidth / 32
                    .TileHeight = .pixelHeight / 32
                    
                    .Frames(1) = Grh
                    frmMain.lstGrhs.AddItem Grh

                End If

            End With

        End If

    Wend
    
    Close handle
    
    LoadGrhData = True
    Exit Function

ErrorHandler:
    LoadGrhData = False

End Function

