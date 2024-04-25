Attribute VB_Name = "Resolution"
Option Explicit
     
Private Const CCDEVICENAME          As Long = 32
Private Const CCFORMNAME            As Long = 32
Private Const DM_BITSPERPEL         As Long = &H40000
Private Const DM_PELSWIDTH          As Long = &H80000
Private Const DM_PELSHEIGHT         As Long = &H100000
Private Const DM_DISPLAYFREQUENCY   As Long = &H400000
Private Const CDS_TEST              As Long = &H4
Private Const ENUM_CURRENT_SETTINGS As Long = -1
     
Private Type typDevMODE

    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long

End Type
     
Private oldResHeight As Long
Private oldResWidth  As Long
Private oldDepth     As Integer
Private oldFrequency As Long
Private bNoResChange As Boolean
     
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, _
    ByVal iModeNum As Long, _
    lptypDevMode As Any) As Boolean

Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long
          
Public Sub SetResolution()

    Dim lRes              As Long
    Dim MidevM            As typDevMODE
    Dim CambiarResolucion As Boolean
    Dim MainWidth         As Integer
    Dim MainHeight        As Integer
       
    lRes = EnumDisplaySettings(0, 0, MidevM)
    
    'MainWidth = frmain.ScaleWidth
   ' MainHeight = frmMain.ScaleHeight
    MainWidth = 1024
    MainHeight = 768
       
    oldResWidth = Screen.Width \ Screen.TwipsPerPixelX
    oldResHeight = Screen.Height \ Screen.TwipsPerPixelY
       
    If NoRes Then
        CambiarResolucion = (oldResWidth < MainWidth Or oldResHeight < MainHeight)
    Else
        CambiarResolucion = (oldResWidth <> MainWidth Or oldResHeight <> MainHeight)

    End If
       
    If AoSetup.bResolucion = 1 Then

        frmMain.WindowState = vbMaximized
       
        NoRes = False

        With MidevM
               
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
            .dmPelsWidth = MainWidth
            .dmPelsHeight = MainHeight
            '.dmBitsPerPel = 32
  
        End With
           
        lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
        
    End If

End Sub
     
Public Sub ResetResolution()

    '***************************************************
    'Autor: Unknown
    'Last Modification: 03/29/08
    'Changes the display resolution if needed.
    'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
    ' 03/29/2008: Maraxus - Properly restores display depth and frequency.
    '***************************************************

    Dim typDevM As typDevMODE

    Dim lRes    As Long
       
    If AoSetup.bResolucion = 1 Then
       
        lRes = EnumDisplaySettings(0, 0, typDevM)
           
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Or DM_DISPLAYFREQUENCY Or DM_BITSPERPEL
            .dmPelsWidth = oldResWidth
            .dmPelsHeight = oldResHeight
            .dmBitsPerPel = oldDepth
        End With
           
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)

    End If

End Sub

