Attribute VB_Name = "modCapturas"
Option Explicit
Const ENUM_CURRENT_SETTINGS As Long = -1&
Const CCDEVICENAME = 32
Const CCFORMNAME = 32


Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type


Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As tPicBmp, RefIID As tGUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Private Type tPicBmp
   lSize As Long
   lType As Long
   lhBmp As Long
   lhPal As Long
   lReserved As Long
End Type

Private Type tGUID
   lData1 As Long
   lData2 As Integer
   lData3 As Integer
   abData4(7) As Byte
End Type

'Byte array <<<<<<
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

Public BitsCaptura() As Byte
Public enviarCaptura As Boolean
Public capturaPath As String
'byte array >>>>>>>>><
    

'Purpose     :  Captures a screen shot
'Inputs      :  sSaveToPath             The path to save the image to
'Outputs     :  Returns a True if successful

Public Sub GetDimensions(ByRef pHeight As Integer, ByRef pWidth As Integer)
Dim DevM As DEVMODE
Call EnumDisplaySettings(0&, ENUM_CURRENT_SETTINGS, DevM)
 pWidth = DevM.dmPelsWidth
 pHeight = DevM.dmPelsHeight
'Debug.Print "Current color depth: " & DevM.dmBitsPerPel & " bits/pixel"
End Sub


Public Function ScreenSnapshot() As Boolean
    Dim lImageWidth As Long, lImageHeight As Long
    Dim lhDCMemory As Long, lhWndSrc As Long
    Dim lhDCSrc As Long, lhwndBmp As Long
    Dim lhwndBmpPrev As Long, lRetVal As Long
    Dim tScreenShot As tPicBmp
    Dim IPic As IPicture  'OR USE IPictureDisp is this doesn't compile (depending on which VB your using)
    Dim tIDispatch As tGUID
    Dim pHeight As Integer, pWidth As Integer
    Const SM_CXSCREEN = 0, SM_CYSCREEN = 1
    
    On Error GoTo ErrFailed
    

    GetDimensions pHeight, pWidth
    lImageWidth = pWidth
    lImageHeight = pHeight
    'lImageWidth = GetSystemMetrics(SM_CXSCREEN)
    'lImageHeight = GetSystemMetrics(SM_CYSCREEN)
    'Get a handle to the desktop window and get the proper device context
    lhWndSrc = GetDesktopWindow()
    lhDCSrc = GetWindowDC(lhWndSrc)
    
    'Create a memory device context for the copy process
    lhDCMemory = CreateCompatibleDC(lhDCSrc)
   
    'Create a bitmap and place it in the memory DC
    lhwndBmp = CreateCompatibleBitmap(lhDCSrc, lImageWidth, lImageHeight)
    lhwndBmpPrev = SelectObject(lhDCMemory, lhwndBmp)
     
    'Copy the screen image to the memory
    Call BitBlt(lhDCMemory, 0, 0, lImageWidth, lImageHeight, lhDCSrc, 0, 0, 13369376)
    
    'Remove the new copy of the the on-screen image
    lhwndBmp = SelectObject(lhDCMemory, lhwndBmpPrev)
   
    'Release the DC resources
    Call DeleteDC(lhDCMemory)
    Call ReleaseDC(lhWndSrc, lhDCSrc)
   
    'Populate OLE IDispatch Interface ID
    With tIDispatch
      .lData1 = &H20400
      .abData4(0) = &HC0
      .abData4(7) = &H46
    End With
   
    With tScreenShot
      .lSize = Len(tScreenShot)     'Length of structure
      .lType = vbPicTypeBitmap                 'Type of Picture (bitmap vbPicTypeBitmap)
      .lhBmp = lhwndBmp             'Handle to bitmap
      .lhPal = 0&                    'Handle to palette (may be null)
    End With
   
    'Create OLE Picture object
    Call OleCreatePictureIndirect(tScreenShot, tIDispatch, 1, IPic)
   
   
    'Return the new Picture object
    'SavePicture IPic, capturaPath
    
    'picScreen.Width = pWidth * 5 '/ 1.5 '* 12
    'picScreen.Height = pHeight * 5 '/ 1.5  '* 12
    'picScreen.PaintPicture IPic, 0, 0, lImageWidth * 15, lImageHeight * 15 ', (pWidth * 5) / 15, (pHeight * 5) / 15
    'frmCapturas.Show , frmMain
    frmMain.picScreen.Width = pWidth / 1.5 '/ 1.5 '* 12
    frmMain.picScreen.Height = pHeight / 1.5 '/ 1.5  '* 12
    frmMain.picScreen.PaintPicture IPic, 0, 0, (pWidth / 1.5) * 15, (pHeight / 1.5) * 15
    
    Call SavePicture(frmMain.picScreen.Image, capturaPath)
    
    ScreenSnapshot = True
    
    Exit Function

ErrFailed:
    'Error occurred
    ScreenSnapshot = False
End Function

'Returns the handle of the desktop
Function GetDesktopHwnd() As Long
    Static slGetDesktopHwnd As Long    'Cache value for speed
    If slGetDesktopHwnd = 0 Then
        slGetDesktopHwnd = GetDesktopHwnd
    End If
    GetDesktopHwnd = slGetDesktopHwnd
End Function

Public Sub HandleCapturarPantalla()
    
    
End Sub
Sub showSnapshot(ByRef picBits() As Byte)
On Error Resume Next
    frmCapturas.Show , frmMain
    If UBound(picBits) = 0 Then Exit Sub
    SetBitmapBits frmCapturas.picScreen1.Image, UBound(picBits), picBits(1)
    Call frmCapturas.ShowSnap
End Sub

Private Function GetByteArray(ByRef sourcePic As PictureBox, ByRef byteArray() As Byte)
    Dim picBits() As Byte, PicInfo As BITMAP, Cnt As Long
    'Get information (such as height and width) about the picturebox
    GetObject sourcePic.Image, Len(PicInfo), PicInfo
    'reallocate storage space
    ReDim picBits(1 To PicInfo.bmWidth * PicInfo.bmHeight * 1) As Byte
    'Copy the bitmapbits to the array
    GetBitmapBits sourcePic.Image, UBound(picBits), picBits(1)
    ReDim byteArray(1 To UBound(picBits))
     byteArray = picBits
    'Set the bits back to the picture SE USA EN EL CLIENTE GM <<<>>>
    '>>><<<
    'SetBitmapBits frmMain.picScreen.Image, UBound(PicBits), PicBits(1)
    'frmMain.picScreen.Refresh
    
End Function
