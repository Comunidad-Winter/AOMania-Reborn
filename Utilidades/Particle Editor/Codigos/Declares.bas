Attribute VB_Name = "Mod_Declaraciones"
Option Explicit

'drag&drop cosas
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage _
               Lib "user32" _
               Alias "SendMessageA" (ByVal hWnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Public DataChanged  As Boolean

'Objetos públicos
Public SurfaceDB    As clsSurfaceManDynDX8   'No va new porque es unainterfaz, el new se pone al decidir que clase de objeto es
Public engine       As New clsDX8Engine

'Particle Groups
Public TotalStreams As Integer
Public StreamData() As Stream
Public Particula()  As Stream

'RGB Type
Public Type RGB

    r As Long
    g As Long
    B As Long

End Type

Public Type Stream

    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    
    speed As Single
    life_counter As Long
    
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer

End Type

Public UserMap As Integer

'Control
Public prgRun  As Boolean 'When true the program ends

'
'********** FUNCIONES API ***********
'

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring _
               Lib "kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                   ByVal lpKeyname As Any, _
                                                   ByVal lpString As String, _
                                                   ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring _
               Lib "kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                 ByVal lpKeyname As Any, _
                                                 ByVal lpdefault As String, _
                                                 ByVal lpreturnedstring As String, _
                                                 ByVal nsize As Long, _
                                                 ByVal lpfilename As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Old fashion BitBlt function
Public Declare Function BitBlt _
               Lib "gdi32" (ByVal hDestDC As Long, _
                            ByVal x As Long, _
                            ByVal Y As Long, _
                            ByVal nWidth As Long, _
                            ByVal nHeight As Long, _
                            ByVal hSrcDC As Long, _
                            ByVal xSrc As Long, _
                            ByVal ySrc As Long, _
                            ByVal dwRop As Long) As Long

Public Declare Function SelectObject _
               Lib "gdi32" (ByVal hdc As Long, _
                            ByVal hObject As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'Added by Juan Martín Sotuyo Dodero

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function SetPixel _
               Lib "gdi32" (ByVal hdc As Long, _
                            ByVal x As Long, _
                            ByVal Y As Long, _
                            ByVal crColor As Long) As Long
                            
Public Declare Function GetPixel _
               Lib "gdi32" (ByVal hdc As Long, _
                            ByVal x As Long, _
                            ByVal Y As Long) As Long

