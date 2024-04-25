VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubclassListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Adaptado para subclasificar Listview emulando un progresvar por
'Name............cSubclassListView
'Author.........Leandro Ascierto leandroascierto@hotmail.com
'Date...........30/01/2008
'------------------------------------------------------
'modulo original de clase
'Name.......... cSuperClass
'Author........ Paul_Caton@hotmail.com
'Date.......... June, 13th 2002

Option Explicit


Private Enum eMsgWhen                                                   'When to callback
    MSG_BEFORE = 1                                                        'Callback before the original WndProc
    MSG_AFTER = 2                                                         'Callback after the original WndProc
    MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                            'Callback before and after the original WndProc
End Enum

Private Enum eThunkType
    SubclassThunk = 0
    HookThunk = 1
    CallbackThunk = 2
End Enum

Private z_IDEflag           As Long         'Flag indicating we are in IDE
Private z_ScMem             As Long         'Thunk base address
Private z_scFunk            As Collection   'hWnd/thunk-address collection
Private z_hkFunk            As Collection   'hook/thunk-address collection
Private z_cbFunk            As Collection   'callback/thunk-address collection
Private Const IDX_INDEX     As Long = 2     'index of the subclassed hWnd OR hook type
Private Const IDX_CALLBACKORDINAL As Long = 22 ' Ubound(callback thunkdata)+1, index of the callback

Private Const IDX_WNDPROC   As Long = 9     'Thunk data index of the original WndProc
Private Const IDX_BTABLE    As Long = 11    'Thunk data index of the Before table
Private Const IDX_ATABLE    As Long = 12    'Thunk data index of the After table
Private Const IDX_PARM_USER As Long = 13    'Thunk data index of the User-defined callback parameter data index
Private Const IDX_UNICODE   As Long = 75    'Must be Ubound(subclass thunkdata)+1; index for unicode support
Private Const ALL_MESSAGES  As Long = -1    'All messages callback
Private Const MSG_ENTRIES   As Long = 32    'Number of msg table entries. Set to 1 if using ALL_MESSAGES for all subclassed windows

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpFn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SendMessageA Lib "user32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessageW Lib "user32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'-------------------------**********************-----------------------
Private Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
'Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
'Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function InvalidateRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bErase As Long) As Long
Private Declare Function ValidateRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function ValidateRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
'------------------*************************--------------------------
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal himl As Long, ByRef cx As Long, ByRef cy As Long) As Long
Private Declare Function CreateIC Lib "gdi32.dll" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByRef lpInitData As Any) As Long
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI)
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

' Constantes para SetWindowPos
Const SWP_FRAMECHANGED = &H20
Const SWP_NOMOVE = &H2
Const SWP_NOZORDER = &H4
Const SWP_NOSIZE = &H1
Const SWP_Flags = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE
  
' Api SetWindowPos
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Any, ByVal Y As Any, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    


'constantes para usar con SendMessage
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PAINTSTRUCT
    hdc                     As Long
    fErase                  As Long
    rcPaint                 As RECT
    fRestore                As Long
    fIncUpdate              As Long
    rgbReserved(1 To 32)    As Byte
End Type

Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    State As Long
    stateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Enum eModeAutoSize
    [item] = -1
    [Header] = -2
End Enum

Private Type HDHITTESTINFO
    pt As POINTAPI
    flags As Long
    iItem As Long
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Private Type HD_ITEM
    mask As Long
    cxy As Long
    pszText As String
    hbm As Long
    cchTextMax As Long
    fmt As Long
    lParam As Long
    iImage As Long
    iOrder As Long
End Type

Private Type TLoHiLong
   Lo As Integer
   Hi As Integer
End Type

Private Type TAllLong
   All As Long
End Type

Private Type LV_COLUMN
    mask As Long
    fmt As Long
    cx As Long
    pszText As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
End Type

Enum tTypeDataColumn
    numero = 0
    FECHA = 1
    Cadena = 2
End Enum

Private Const WM_GETFONT As Long = &H31

Private Const LVM_GETITEMCOUNT As Long = (LVM_FIRST + 4)
Private Const LVM_GETITEMTEXTA As Long = (LVM_FIRST + 45)
Private Const LVM_GETSUBITEMRECT As Long = (LVM_FIRST + 56)
Private Const LVM_GETVIEWRECT As Long = (LVM_FIRST + 34)
Private Const LVM_GETWORKAREAS As Long = (LVM_FIRST + 70)
Private Const LVM_GETTOPINDEX As Long = (LVM_FIRST + 39)
Private Const LVM_GETHEADER  As Long = (LVM_FIRST + 31)
Private Const LVM_GETCOLUMN = (LVM_FIRST + 25)

Private Const LVCF_TEXT As Long = &H4
Private Const LVCF_ORDER As Long = &H20
Private Const LVCF_FMT As Long = &H1
Private Const LVCF_IMAGE As Long = &H10

Private Const LVIF_TEXT = &H1
Private Const LVIR_LABEL = 2

Private Const WM_DESTROY        As Long = &H2
Private Const WM_CLOSE          As Long = &H10
Private Const WM_NCLBUTTONDOWN  As Long = &HA1
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_LBUTTONDBLCLK  As Long = &H203
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const WM_LBUTTONUP      As Long = &H202
Private Const WM_MBUTTONDBLCLK  As Long = &H209
Private Const WM_MBUTTONDOWN    As Long = &H207
Private Const WM_MBUTTONUP      As Long = &H208
Private Const WM_RBUTTONDBLCLK  As Long = &H206
Private Const WM_RBUTTONDOWN    As Long = &H204
Private Const WM_RBUTTONUP      As Long = &H205
Private Const WM_MOVING         As Long = &H216
Private Const WM_PAINT          As Long = &HF&
Private Const WM_SETFONT        As Long = &H30
Private Const WM_ERASEBKGND     As Long = &H14

Private Const HDM_FIRST = &H1200
Private Const HDM_CLEARFILTER As Long = HDM_FIRST + 24
Private Const HDM_CREATEDRAGIMAGE As Long = HDM_FIRST + 16
Private Const HDM_DELETEITEM As Long = HDM_FIRST + 2
Private Const HDM_EDITFILTER As Long = HDM_FIRST + 23
Private Const HDM_GETBITMAPMARGIN As Long = (HDM_FIRST + 21)
Private Const HDM_GETIMAGELIST As Long = (HDM_FIRST + 9)
Private Const HDM_GETITEMA As Long = (HDM_FIRST + 3)
Private Const HDM_GETITEMCOUNT As Long = (HDM_FIRST + 0)
Private Const HDM_GETITEMW As Long = (HDM_FIRST + 11)
Private Const HDM_GETITEMRECT As Long = (HDM_FIRST + 7)
Private Const HDM_GETORDERARRAY As Long = (HDM_FIRST + 17)
'Private Const HDM_GETUNICODEFORMAT As Long = CCM_GETUNICODEFORMAT
Private Const HDM_HITTEST As Long = (HDM_FIRST + 6)
Private Const HDM_INSERTITEMA As Long = (HDM_FIRST + 1)
Private Const HDM_INSERTITEMW As Long = (HDM_FIRST + 10)
Private Const HDM_LAYOUT As Long = (HDM_FIRST + 5)
Private Const HDM_ORDERTOINDEX As Long = (HDM_FIRST + 15)
Private Const HDM_SETBITMAPMARGIN As Long = (HDM_FIRST + 20)
Private Const HDM_SETFILTERCHANGETIMEOUT As Long = (HDM_FIRST + 22)
Private Const HDM_SETHOTDIVIDER As Long = (HDM_FIRST + 19)
Private Const HDM_SETIMAGELIST As Long = (HDM_FIRST + 8)
Private Const HDM_SETITEMA As Long = (HDM_FIRST + 4)
Private Const HDM_SETITEMW As Long = (HDM_FIRST + 12)
Private Const HDM_SETORDERARRAY As Long = (HDM_FIRST + 18)

Private Const HDI_TEXT As Long = &H2
Private Const HDI_ORDER As Long = &H80
Private Const HDI_LPARAM = &H8
Private Const HDI_IMAGE As Long = &H20
Private Const HDI_HIDDEN As Long = &H1
Private Const HDI_WIDTH As Long = &H1
Private Const HDI_HEIGHT As Long = HDI_WIDTH
Private Const HDI_FORMAT As Long = &H4
Private Const HDI_FILTER As Long = &H100
Private Const HDI_DI_SETITEM As Long = &H40
Private Const HDI_BITMAP As Long = &H10

Private Const HDF_BITMAP As Long = &H2000
Private Const HDF_BITMAP_ON_RIGHT As Long = &H1000
Private Const HDF_CENTER As Long = 2
Private Const HDF_IMAGE As Long = &H800
Private Const HDF_JUSTIFYMASK As Long = &H3
Private Const HDF_LEFT As Long = 0
Private Const HDF_OWNERDRAW As Long = &H8000
Private Const HDF_RIGHT As Long = 1
Private Const HDF_RTLREADING As Long = 4
Private Const HDF_SORTDOWN As Long = &H200
Private Const HDF_SORTUP As Long = &H400
Private Const HDF_STRING As Long = &H4000


Private Const ScrCopy = &HCC0020
Private Const GWL_WNDPROC = (-4)
Private Const DT_CENTER = &H1
Private Const DT_WORDBREAK = &H10
Private Const TRANSPARENT = 1
Private Const LF_FACESIZE = 32
Private Const ILD_TRANSPARENT = &H1
Private Const ILD_SELECTED As Long = &H2
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Const WM_SETREDRAW As Long = &HB&

Private Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName(LF_FACESIZE) As Byte
End Type
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long


Enum eAlignTextValue
    [eLeft] = 0
    [eCenter] = 1
    [eRight] = 2
End Enum

'---------------------------------------------

Private Type PropertyItems
    
    SubItemsIndex                 As Long
    BackColor                     As Long
    BorderColor                   As Long
    ForeColor                     As Long
    DisplayValue                  As Boolean
    DisplayPercent                As Boolean
    MaxValue                      As Currency
    MinValue                      As Currency
    ForeColorText                 As Long
    FieldCurrency                 As Boolean
    NumDigistsAfterDecimal        As Long
    displayGraph                  As Boolean
    AlignTextValue                As eAlignTextValue

End Type

Dim DataSubItems()          As PropertyItems
Dim ColItem                 As HD_ITEM

Dim m_Hwnd_LV               As Long
Dim m_Hwnd_CH               As Long

Dim IsDrag                  As Boolean
Dim MouseDown               As Boolean
Dim Indice                  As Long
Dim TimerEnabled            As Boolean
Dim m_IconAlingmentRight    As Boolean
Dim m_TextNormalColor       As Long
Dim m_TextResalteColor      As Long
Dim m_HedersFontBlod        As Boolean
Dim m_SkinPicture           As StdPicture
Dim m_hBmp                  As Long
Dim m_SkinHeight            As Long
Dim m_SkinWidth             As Long
Dim RenderWidth             As Long

Private m_Redraw            As Boolean

Private ViejaFuente         As Long

Private mUseSkinColumnHeader As Boolean




' añade y configura un nuevo item para las columnas gráficas
Public Function ModifyItem( _
    ByVal indexColListView As Long, _
    Optional ByVal DisplayValue As Boolean = True, _
    Optional ByVal DisplayPercent As Boolean = False, _
    Optional ByVal ForeColorText As Long = vbBlack, _
    Optional ByVal BackColor As Long = vbRed, _
    Optional ByVal BorderColor As Long = vbBlue, _
    Optional ByVal ForeColorGraph As Long = vbGreen, _
    Optional ByVal bFieldCurrency As Boolean = False, _
    Optional NumDigistsAfterDecimal As Long = 0, _
    Optional bdisplayGraph As Boolean = True, _
    Optional AlignTextValue As eAlignTextValue = eCenter) As Boolean
    
    
    If m_Hwnd_LV = 0 Then
       showError
       Exit Function
    End If
        
    Dim i As Integer
    
    For i = 0 To UBound(DataSubItems)
        If DataSubItems(i).SubItemsIndex = indexColListView Then
            With DataSubItems(i)
                .SubItemsIndex = indexColListView
                .BackColor = BackColor
                .BorderColor = BorderColor
                .ForeColor = ForeColorGraph
                .DisplayValue = DisplayValue
                .DisplayPercent = DisplayPercent
                .ForeColorText = ForeColorText
                .FieldCurrency = bFieldCurrency
                .NumDigistsAfterDecimal = NumDigistsAfterDecimal
                .displayGraph = bdisplayGraph
                .AlignTextValue = AlignTextValue
            End With
            Exit For
        End If
    Next
        
    
    RedrawWindow m_Hwnd_LV, ByVal 0&, ByVal 0&, &H1
    
    ModifyItem = True

End Function


' añade y configura un nuevo item para las columnas gráficas
Public Function AddSubItem( _
    ByVal indexColListView As Long, _
    Optional ByVal DisplayValue As Boolean = True, _
    Optional ByVal DisplayPercent As Boolean = False, _
    Optional ByVal ForeColorText As Long = vbBlack, _
    Optional ByVal BackColor As Long = vbRed, _
    Optional ByVal BorderColor As Long = vbBlue, _
    Optional ByVal ForeColorGraph As Long = vbGreen, _
    Optional ByVal bFieldCurrency As Boolean = False, _
    Optional NumDigistsAfterDecimal As Long = 0, _
    Optional bdisplayGraph As Boolean = True, _
    Optional AlignTextValue As eAlignTextValue = eCenter) As Boolean
    
    
    If m_Hwnd_LV = 0 Then
       showError
       Exit Function
    End If
    
    Dim lIndex As Long
    Dim Count As Long
    Dim i As Integer
    
    Count = UBound(DataSubItems)
    For i = 1 To Count
        If DataSubItems(i).SubItemsIndex = indexColListView Then Exit Function
        
    Next

    lIndex = UBound(DataSubItems) + 1
    ReDim Preserve DataSubItems(lIndex)
    With DataSubItems(lIndex)
        .SubItemsIndex = indexColListView
        .BackColor = BackColor
        .BorderColor = BorderColor
        .ForeColor = ForeColorGraph
        .DisplayValue = DisplayValue
        .DisplayPercent = DisplayPercent
        .ForeColorText = ForeColorText
        .MaxValue = GetValueMaxColumn(indexColListView)
        .MinValue = GetValueMinColumn(indexColListView)
        .FieldCurrency = bFieldCurrency
        .NumDigistsAfterDecimal = NumDigistsAfterDecimal
        .displayGraph = bdisplayGraph
        .AlignTextValue = AlignTextValue
    End With
    
    
    RedrawWindow m_Hwnd_LV, ByVal 0&, ByVal 0&, &H1
    
    AddSubItem = True

End Function

' Oculta la gráfica de todas las columnas o de una especifica

Sub HideShowColumnGraph( _
    value As Boolean, _
    Optional IcolIndex As Integer = -1)

On Error GoTo err_Sub

    If IcolIndex = -1 Then
       
       Dim i As Integer
       For i = LBound(DataSubItems) To UBound(DataSubItems)
           DataSubItems(i).displayGraph = value
       Next
    Else
       For i = LBound(DataSubItems) To UBound(DataSubItems)
           If DataSubItems(i).SubItemsIndex = IcolIndex Then
              DataSubItems(i).displayGraph = value
           End If
       Next
    End If
    
    RedrawWindow m_Hwnd_LV, ByVal 0&, ByVal 0&, &H1
    
Exit Sub
err_Sub:
    MsgBox Err.Description, vbCritical, "error al oculta o mostrar el gráfico"
    
End Sub

' refresca el área de datos del ListView cuando se añaden o eliminan nuevos valores

Function RefreshData()
    
    Dim c As Integer
    
    For c = 1 To UBound(DataSubItems)
        DataSubItems(c).MaxValue = GetValueMaxColumn(DataSubItems(c).SubItemsIndex)
        DataSubItems(c).MinValue = GetValueMinColumn(DataSubItems(c).SubItemsIndex)
    Next
    
    RedrawWindow m_Hwnd_LV, ByVal 0&, ByVal 0&, &H1
    
End Function

' refresca el ColumnHeader cuando se cambia el skin
Private Sub RefreshColData()
    Call SendMessage(m_Hwnd_LV, WM_SETREDRAW, 0&, 0&)
    
    RedrawWindow m_Hwnd_LV, ByVal 0&, ByVal 0&, &H1
    
    Call SendMessage(m_Hwnd_LV, WM_SETREDRAW, 1&, 0&)
End Sub

Function ExportListViewToExcel( _
    OuputPathXls As String, _
    ListView As ListView, _
    Optional Progressbar As Progressbar) As Boolean


    On Error GoTo err_fun
    
    
    If Len(Dir(OuputPathXls)) <> 0 Then
       MsgBox "el archivo ya existe, elija otro", vbCritical
       Exit Function
    End If
    
    'CREAR EL OBJETO (INSTANCIAR)CON EL OBJETO APLICACION (obj_Excel)
    Dim obj_Excel As Object
    Dim obj_Libro As Object
    
    ' Nueva referencia a Excel y nuevo referencia al Libro
    Set obj_Excel = CreateObject("Excel.Application")
    
    ' si es la versión de Excel 97, asigna la hoja activa ( ActiveSheet )
    If Val(obj_Excel.Application.Version) >= 8 Then
        Set obj_Libro = obj_Excel.Workbooks.Add
    Else
        Set obj_Libro = obj_Excel.Workbooks.Add(OuputPathXls)
    End If
    
    
    
    'Variables para las columnas y filas
    Dim Col As Integer, Fila As Integer
    
    With obj_Excel
        
        'crea el libro
        'Set obj_Libro = .Workbooks.Add(App.Path & "\libro1.xls")

    End With

    With obj_Libro
    
        'Asignamos El valor Maximo del Progress teniendo _
         como dato la cantidad de items en el ListView
        If Not Progressbar Is Nothing Then
            Progressbar.Max = ListView.ListItems.Count
        End If
    
        'Referencia a la hoja con índice 1
        With .Sheets(1)
        
            For Col = 1 To ListView.ColumnHeaders.Count - 1
                .Cells(1, Col) = CStr(ListView.ColumnHeaders(Col).Text)
            Next
        
            'Recorremos la cantidad de items del ListView
            For Fila = 1 To ListView.ListItems.Count
                Col = 1
                'Asignamos EL Item actual en la celda
                .Cells(Fila + 1, Col) = CStr(ListView.ListItems.item(Fila))
                
                'Asignamos EL SubitemItem actual en la celda
                 For Col = 1 To ListView.ColumnHeaders.Count - 1
                     .Cells(Fila + 1, Col + 1) = CStr(ListView.ListItems(Fila).SubItems(Col))
                 Next

                 If Not Progressbar Is Nothing Then
                     'Aumentamos en 1 la propiedad value
                     Progressbar.value = Progressbar.value + 1
                 End If
            Next
        End With
    End With
    
    obj_Excel.ActiveWorkbook.SaveAs OuputPathXls
    
    obj_Excel.ActiveWorkbook.Close
    
    'Destruimos las variables de objeto
    
    Set obj_Libro = Nothing
    obj_Excel.quit
    Set obj_Excel = Nothing
    'Ok
    ExportListViewToExcel = True
    
    If Not Progressbar Is Nothing Then
       Progressbar.value = 0
    End If
    
    
Exit Function

err_fun:

ExportListViewToExcel = False

MsgBox Err.Description, vbCritical

On Error Resume Next
Set obj_Libro = Nothing
Set obj_Excel = Nothing
If Not Progressbar Is Nothing Then Progressbar.value = 0
End Function


Function ExportRecordsetAdoToExcel( _
    ExcelPath As String, _
    Sql As String, _
    NameHoja As String, _
    cn As Object) As Boolean
    
    On Error GoTo errSub
    
    Dim rec As Object
    
    Set rec = CreateObject("ADODB.Recordset")
    
    Dim oXls As Object, oLibro As Object, Hoja As Object

    Dim recArreglo As Variant
    
    Dim i_Field As Integer, iRec As Long
    Dim Columna As Integer, Fila As Integer
        
    'Abre el Recordset pasandole la cadena sql
    rec.Open Sql, cn, 3, 1
    
    'Crea los objetos para utilizar el Excel
    Set oXls = CreateObject("Excel.Application")
    Set oLibro = oXls.Workbooks.Add
    
    'Hace referencia a la hoja
    Set Hoja = oLibro.Worksheets(NameHoja)
  
    oXls.UserControl = True

    i_Field = rec.Fields.Count
    
    For Columna = 1 To i_Field
        Hoja.Cells(1, Columna).value = rec.Fields(Columna - 1).Name
    Next
        

    If Val(Mid(oXls.Version, 1, InStr(1, oXls.Version, ".") - 1)) > 8 Then

        Hoja.Cells(2, 1).CopyFromRecordset rec
    
    Else

        recArreglo = rec.GetRows

        iRec = UBound(recArreglo, 2) + 1
        
        For Columna = 0 To i_Field - 1
            For Fila = 0 To iRec - 1

                If IsDate(recArreglo(Columna, Fila)) Then
                    recArreglo(Columna, Fila) = Format(recArreglo(Columna, Fila))

                ElseIf IsArray(recArreglo(Columna, Fila)) Then
                    recArreglo(Columna, Fila) = "Array Field"
                End If
            Next Fila
        Next Columna
            
        'traspasa los datos a la hoja de Excel
        Hoja.Cells(2, 1).Resize(iRec, i_Field).value = Pasar(recArreglo)
    End If

    oXls.Selection.CurrentRegion.Columns.AutoFit
    oXls.Selection.CurrentRegion.Rows.AutoFit

    rec.Close
    
    Set rec = Nothing
    
    Set Hoja = Nothing
    Set oLibro = Nothing
    oXls.Visible = True
    Set oXls = Nothing
    
    ExportRecordsetAdoToExcel = True
    
    Exit Function
errSub:
    MsgBox Err.Description, vbCritical, "Error"

End Function

Private Function Pasar(v As Variant) As Variant
    Dim X As Long, Y As Long, xMax As Long, yMax As Long, T As Variant
    xMax = UBound(v, 2): yMax = UBound(v, 1)
    ReDim T(xMax, yMax)
    For X = 0 To xMax
        For Y = 0 To yMax
            T(X, Y) = v(Y, X)
        Next Y
    Next X
    Pasar = T
End Function


Function ExportCsv( _
    ListView As ListView, _
    OuputPath As String) As Boolean
    
    
    
    On Error GoTo errSub
    
    Dim Linea As String
    Dim X As Integer
    Dim i As Integer
    Dim dato As String
    
    'Crea un archivo csv
    Open OuputPath For Output As #1
    
    'Recorre los encabezados de columna
    For i = 1 To ListView.ColumnHeaders.Count
        ' caption de la columna
        dato = ListView.ColumnHeaders(i).Text
        '.. si no tiene valor
        If dato = vbNullString Then
           dato = ""
        End If
        ' Le agrega el separador
        Linea = Linea & dato & ";"
    Next
    
    If Linea <> vbNullString Then
        ' quita el último delimitador
        Linea = Left(Linea, Len(Linea) - 1)
        ' Guarda la línea en el archivo
        Print #1, Linea
    End If
    
    'recorre los Item
    For i = 1 To ListView.ListItems.Count
        'valor del Item
        dato = ListView.ListItems(i)
        If dato = vbNullString Then
           dato = ""
        End If
        ' agrega el delimitador
        Linea = dato & ";"
        'recorre los SubItems
        For X = 1 To ListView.ColumnHeaders.Count - 1
            'dato  de los SubItems
            dato = ListView.ListItems.item(i).SubItems(X)
            If dato = vbNullString Then
               dato = ""
            End If
            ' agrega el delimitador
            Linea = Linea & dato & ";"
        Next
    ' quita el último ( ; )
    Linea = Left(Linea, Len(Linea) - 1)
    'Guarda la linea
    Print #1, Linea
    Next
    
    'Cierra el archivo
    Close
    ExportCsv = True

Exit Function
'Error
errSub:
MsgBox Err.Description, vbCritical

End Function



Function FindItem( _
    ListView As ListView, _
    sValue As String, _
    nCol As Integer, _
    bFraseCompleta As Boolean) As Long
    
    Dim i As Long
    Dim iStart As Long
    Dim oItem As ListItem
    
    
    
    'If ListView.SelectedItem Is Nothing Then
    '   i = 1
    If Not ListView.SelectedItem Is Nothing Then
       iStart = ListView.SelectedItem.Index + 1
    Else
       iStart = 1
    End If
    
    With ListView
    For i = iStart To ListView.ListItems.Count
        
        Set oItem = ListView.ListItems(i)
        
        Dim sItem As String
        
        If nCol = 0 Then
            sItem = ListView.ListItems(i)
        Else
            sItem = oItem.SubItems(nCol)
        End If
        If bFraseCompleta = False Then
           Dim nPos As Integer
           
           nPos = InStr(LCase(sItem), LCase(sValue))
           If nPos <> 0 Then
              FindItem = oItem.Index
              Exit For
           End If
        ElseIf bFraseCompleta = True Then
             If LCase(sItem) = LCase(sValue) Then
              FindItem = oItem.Index
              Exit For
             End If
        End If
    Next
    End With
    
    ListView.SetFocus
    
End Function



Function ExportHtml( _
    Path_Html As String, _
    ListView As ListView, _
    Optional TEXTO_ENCABEZADO As String = "", _
    Optional TEXTO_PIE As String = "", _
    Optional SIZE_BORDE_TABLA As String = "1", _
    Optional COLOR_BORDE_TABLA As String = "white", _
    Optional COLOR_FONDO_TABLA As String = "white", _
    Optional COLOR_FONDO_ENCABEZADO As String = "white", _
    Optional ALIGN_ENCABEZADO As String = "left", _
    Optional FONT_SIZE_ENCABEZDO As String = "2", _
    Optional FONT_NAME_ENCABEZDO As String = "Verdana", _
    Optional COLOR_FUENTE_ENCABEZADO As String = "black", _
    Optional COLOR_FONDO_CELDA As String = "white", _
    Optional ALIGN_FUENTE_CELDA As String = "Left", _
    Optional COLOR_FUENTE_CELDA As String = "black", _
    Optional FONT_SIZE_CELDA As String = "1", _
    Optional FONT_NAME_CELDAS As String = "Verdana", _
    Optional COLOR_FONDO_PAGINA As String = "white", _
    Optional FONT_SIZE_TITULO As String = "4", _
    Optional COLOR_FUENTE_TITULO As String = "black", _
    Optional FONT_SIZE_PIE As String = "2") As Boolean

Dim codigo_Html As String
Dim Fila As Integer
Dim Columna As Integer
Dim fname As String
Dim F As Integer
Dim item As ListItem

    Screen.MousePointer = vbHourglass
    F = FreeFile
    Open Path_Html For Output As F

    'Etiqueta de inicio de la tabla, Color de la Tabla y el borde
    codigo_Html = "<TABLE width=100% BORDER=" & SIZE_BORDE_TABLA & _
                  " Bordercolor=" & COLOR_BORDE_TABLA & " bgcolor=" & _
                    COLOR_FONDO_TABLA & ">" & vbCrLf

    'Crea la etiqueta de apertura para los encabezados de columna
    codigo_Html = codigo_Html & "<TR bgcolor=" & _
    COLOR_FONDO_ENCABEZADO & "> " & vbCrLf

    'Recorre las columnas del ListView ..si es que tiene
    For Columna = 1 To ListView.ColumnHeaders.Count

        codigo_Html = codigo_Html & " <TH><div align=" & _
                      ALIGN_ENCABEZADO & "><font size=" & _
                      FONT_SIZE_ENCABEZDO & " face=" & _
                      FONT_NAME_ENCABEZDO & _
                      " color=" & COLOR_FUENTE_ENCABEZADO & ">" & _
                      ListView.ColumnHeaders(Columna).Text & _
                      "</font></div></TH>" & vbCrLf

    Next Columna

    codigo_Html = codigo_Html & "</TR>" & vbCrLf

    Print #F, codigo_Html
    codigo_Html = ""

    ' Recorre las filas y las columnas del listview y genera el _
    etiquetado. ( lo va almacenando en la variable " codigo_Html "

    For Fila = 1 To ListView.ListItems.Count

        'Etiqueta de apertura de la fila
        codigo_Html = codigo_Html & "<TR bgcolor=" & _
                                    COLOR_FONDO_CELDA & " > " & vbCrLf

        'Etiqueta de apertura, propiedades de la celda y el valor del item actual
        codigo_Html = codigo_Html & " <td><div align=" & _
                                    ALIGN_FUENTE_CELDA & "><font color=" & _
                                    COLOR_FUENTE_CELDA & " size=" & _
                                    FONT_SIZE_CELDA & " face=" & _
                                    FONT_NAME_CELDAS & ">" & _
                                    ListView.ListItems(Fila).Text & _
                                    "</font></div></td>" & vbCrLf

        'Recorre los subitems ...si es que hay
        For Columna = 2 To ListView.ColumnHeaders.Count

            codigo_Html = codigo_Html & " <td><div align=" & _
                          ALIGN_FUENTE_CELDA & "><font color=" & _
                          COLOR_FUENTE_CELDA & " size=" & _
                          FONT_SIZE_CELDA & " face=" & _
                          FONT_NAME_CELDAS & ">" & _
                          ListView.ListItems(Fila).SubItems(Columna - 1) & _
                          "</font></div></td>" & vbCrLf

        Next

        ' Cierra la etiqueta de la fila actual
        codigo_Html = codigo_Html & "</TR>" & vbCrLf
        Print #F, codigo_Html
        codigo_Html = ""

    Next
    
    ' Cierra la etiqueta HTML de la Tabla
    codigo_Html = "</table>" & vbCrLf

    'Agrega las etiquetas restantes y el pie de página
    codigo_Html = "<HTML><HEAD></HEAD><BODY BGCOLOR=" _
                   & COLOR_FONDO_PAGINA & ">" & vbCrLf & _
                   "<p><font face=verdana size=" & _
                   FONT_SIZE_TITULO & " color=" & _
                   COLOR_FUENTE_TITULO & ">" & _
                   TEXTO_ENCABEZADO & _
                   "</font></p><HR>" & codigo_Html & _
                   "<HR><font face=verdana size=" & _
                   FONT_SIZE_PIE & " color=" & _
                   COLOR_FUENTE_TITULO & _
                   ">" & TEXTO_PIE & "</font></BODY></HTML>"

    
    'Abre y Crea el archivo Html

    'Escribe los datos
    Print #F, codigo_Html

    Close

    Screen.MousePointer = vbNormal
    
    ExportHtml = True

Exit Function

MsgBox Err.Description

End Function


' carga un recordset en el listview


Sub LoadListViewOfRecordset( _
        ListView As ListView, _
        Sql As String, _
        ActiveConnection As Object, _
        Optional SmallIconIndex As Integer = 0)

    Dim Campo As Integer

    On Error GoTo errSub
    
    Dim item As ListItem
    Dim i As Long
    Dim rst As Object
    Set rst = CreateObject("ADODB.Recordset")

    rst.Open Sql, ActiveConnection, , 3
    
    With ListView
    
        .View = lvwReport
    
        .ListItems.Clear
        .ColumnHeaders.Clear
    End With

    Screen.MousePointer = vbHourglass

    'Agrega los nombres campo junto con los encabezados de columna para el ListView
    For Campo = 0 To rst.Fields.Count - 1
        ListView.ColumnHeaders.Add , , rst.Fields(Campo).Name & "   "
    Next

    ' Recorre todos los registros del Recordset
    While Not rst.EOF
        
        
        i = 1
        
        'Agrega el Item
        Set item = ListView.ListItems.Add(, , rst.Fields(0), , SmallIconIndex)
        'Agrega los SubItem
        For Campo = 1 To rst.Fields.Count - 1
            
            If Not IsNull(rst.Fields(Campo)) Then
                item.SubItems(i) = CStr(rst.Fields(Campo))
            End If
            i = i + 1
        Next
    
        'Siguiente registro
        rst.MoveNext
    Wend
    If Not rst Is Nothing Then
       If rst.State = 1 Then rst.Close
       Set rst = Nothing
    End If
    Screen.MousePointer = vbDefault
    DoEvents

Exit Sub
'Error
errSub:

    MsgBox Err.Description, vbCritical, "Error"
    Screen.MousePointer = vbDefault
End Sub

' retorna el mínimo de una columna
Function GetValueMinColumn(ByVal ICol As Integer) As Currency
    Dim itemsCount As Long
    itemsCount = SendMessage(m_Hwnd_LV, LVM_GETITEMCOUNT, 0, 0)
    
    Dim valueMin As Currency
    Dim value As Currency
    Dim i As Integer
    
    valueMin = 99999999999999#
    
    For i = 0 To itemsCount - 1
    
        value = GetListItem(m_Hwnd_LV, i, ICol)
        
        If value < valueMin Then
           valueMin = value
        End If
    Next
    
    If valueMin = 99999999999999# Then
        GetValueMinColumn = 0
    Else
        GetValueMinColumn = valueMin
    End If
    
End Function

' retorna el máximo de una columna
Function GetValueMaxColumn(ByVal ICol As Integer) As Currency
    Dim itemsCount As Long
    itemsCount = SendMessage(m_Hwnd_LV, LVM_GETITEMCOUNT, 0, 0)
    
    Dim ValueMax As Currency
    Dim value As Currency
    Dim i As Integer
    
    ValueMax = -99999999999999#
    
    For i = 0 To itemsCount - 1
    
        value = GetListItem(m_Hwnd_LV, i, ICol)
        
        If value > ValueMax Then
           ValueMax = value
        End If
    Next
    
    If ValueMax = -99999999999999# Then
       GetValueMaxColumn = 0
    Else
       GetValueMaxColumn = ValueMax
    End If
    
End Function

' retorna la suma de una columna
Function GetSumColumn(ByVal ICol As Integer) As Currency
    
    Dim itemsCount As Long
    itemsCount = SendMessage(m_Hwnd_LV, LVM_GETITEMCOUNT, 0, 0)
    
    Dim Total As Currency
    
    Dim i As Integer
    For i = 0 To itemsCount - 1
        Total = Total + CCur(GetListItem(m_Hwnd_LV, i, ICol))
    Next
    
    GetSumColumn = CCur(Total)
    
End Function

' retorna el promedio de una columna
Function GetAverageColumn(ByVal ICol As Integer) As Currency
    
    Dim itemsCount As Long
    itemsCount = SendMessage(m_Hwnd_LV, LVM_GETITEMCOUNT, 0, 0)
    
    Dim Total As Currency

    Total = GetSumColumn(ICol)

    If itemsCount > 0 Then
        GetAverageColumn = CCur(Total) / itemsCount
    End If
    
End Function


' retorna el valor del Item x
''''''''''''''''''''''''''''''''''''''''
Private Function GetListItem( _
    hwnd As Long, _
    ByVal iItem As Long, _
    ByVal iSubItem As Long) As Currency
 
 On Error Resume Next
    Dim objItem As LV_ITEM
    Dim baBuffer(32) As Byte
    Dim lngLength As Long
    Dim stempValue As String
    
    objItem.mask = LVIF_TEXT
    objItem.iSubItem = iSubItem
    objItem.pszText = VarPtr(baBuffer(0))
    objItem.cchTextMax = UBound(baBuffer)
    lngLength = SendMessageLong(hwnd, LVM_GETITEMTEXTA, iItem, VarPtr(objItem))
    
    stempValue = Left$(StrConv(baBuffer, vbUnicode), lngLength)
    
    If Not IsNumeric(stempValue) Then
       GetListItem = 0
    Else
       GetListItem = CCur(stempValue)
    End If
    
End Function

' Eliminar todos los SubItems agregados

Public Function DeleteAllSubItems() As Boolean
    On Error Resume Next
    ReDim DataSubItems(0)
    If Err.Number = 0 Then
        DeleteAllSubItems = True
        RefreshColData
    End If
    
End Function



Public Function DeleteSubItems(ByVal Index As Integer) As Boolean
On Error GoTo Fallo
    Dim TempDSI As PropertyItems
    Dim Count As Integer
    Dim i As Integer
    Dim Encontrado As Boolean
    
    Count = UBound(DataSubItems)
        'busca el items
        For i = 1 To Count
            If DataSubItems(i).SubItemsIndex = Index Then
                Index = i
                Encontrado = True
                Exit For
            End If
        Next
        
        If Not Encontrado Then Exit Function
        'Redimenciona la matriz
        For i = Index To Count - 1
            TempDSI = DataSubItems(i + 1)
            DataSubItems(i) = TempDSI
        Next
   
    
    ReDim Preserve DataSubItems(Count - 1)
    DeleteSubItems = True
Fallo:
 
End Function

Public Sub AutoSizeHeader( _
    ByVal La_Columna As Long, _
    ByVal Modo_De_Ajuste As eModeAutoSize)

   Call SendMessage(m_Hwnd_LV, LVM_SETCOLUMNWIDTH, La_Columna - 1, ByVal Modo_De_Ajuste)
   'Call SendMessage(m_Hwnd_LV, LVM_SETCOLUMNWIDTH, La_Columna - 1, ByVal 15)
End Sub


Sub SortColumn( _
    ListView As ListView, _
    Indice As Long, _
    TypeDataColumn As tTypeDataColumn)

    On Error Resume Next
    
    Dim tag As String
    
    Select Case TypeDataColumn
        Case 0: tag = "NUMBER"
        Case 1: tag = "DATE"
        Case 2: tag = "STRING"
    End Select
    
    
    With ListView
        
        ListView.ColumnHeaders(Indice).tag = tag
        
        Dim i As Long
        Dim Formato As String
        Dim strData() As String
        
        Dim Columna As Long
        
        Call SendMessage(ListView.hwnd, WM_SETREDRAW, 0&, 0&)
        
        
        Columna = Indice
        
        '''''''''''''''''''''''''''''''''''''''''''''
        ' Tipo de dato a ordenar
        ''''''''''''''''''''''''''''''''''''''''''''''
        
        Select Case UCase$(tag)
    
        
        ' Fecha
        '''''''''''''''''''''''''''''''''''''''''''''
        Case "DATE"
        
            Formato = "YYYYMMDDHhNnSs"
        
            ' Ordena alfabéticamente la columna con Fechas _
              ( es la columna que tiene en el tag el valor DATE )
        
            With .ListItems
                If (Columna > 0) Then
                    For i = 1 To .Count
                        With .item(i).ListSubItems(Columna)
                            .tag = .Text & Chr$(0) & .tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    Formato)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .item(i)
                            .tag = .Text & Chr$(0) & .tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    Formato)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                End If
            End With
            
            ' Ordena alfabéticamente
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = Indice
            .Sorted = True
            
            With .ListItems
                If (Columna > 0) Then
                    For i = 1 To .Count
                        With .item(i).ListSubItems(Columna)
                            strData = Split(.tag, Chr$(0))
                            .Text = strData(0)
                            .tag = strData(1)
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .item(i)
                            strData = Split(.tag, Chr$(0))
                            .Text = strData(0)
                            .tag = strData(1)
                        End With
                    Next i
                End If
            End With
            
        ' Datos de numéricos
        '''''''''''''''''''''''''''''''''''''''''''''
        Case "NUMBER"
        
            ' Ordena alfabéticamente la columna con números _
              ( es la columna que tiene en el tag el valor NUMBER )
        
            Formato = String(30, "0") & "." & String(30, "0")
                
            With .ListItems
                If (Columna > 0) Then
                    For i = 1 To .Count
                        With .item(i).ListSubItems(Columna)
                            .tag = .Text & Chr$(0) & .tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        Formato)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        Formato))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .item(i)
                            .tag = .Text & Chr$(0) & .tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        Formato)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        Formato))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                End If
            End With
            
            ' Ordena alfabéticamente
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = Indice
            .Sorted = True
            
            With .ListItems
                If (Columna > 0) Then
                    For i = 1 To .Count
                        With .item(i).ListSubItems(Columna)
                            strData = Split(.tag, Chr$(0))
                            .Text = strData(0)
                            .tag = strData(1)
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .item(i)
                            strData = Split(.tag, Chr$(0))
                            .Text = strData(0)
                            .tag = strData(1)
                        End With
                    Next i
                End If
            End With
        
        Case Else
                    
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = Indice
            .Sorted = True
            
        End Select
    
    End With
    
    Call SendMessage(ListView.hwnd, WM_SETREDRAW, 1&, 0&)
    'ListView1.Refresh
    
End Sub

Private Function InvNumber(ByVal Number As String) As String
    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function



'*********---------------*********------------**********------------**********----------
'*********---------------*********------------**********------------**********----------
'*********---------------*********------------**********------------**********----------
Public Property Let SkinPicture(ByVal value As StdPicture)

    If m_Hwnd_LV = 0 Then
       showError
       Exit Property
    End If

    DeleteObject m_hBmp ' Old
    Set m_SkinPicture = value
    m_hBmp = CreateCompatibleDC(0)
    Call SelectObject(m_hBmp, m_SkinPicture.handle)
    m_SkinHeight = ConvertPixelHimetric(m_SkinPicture.Height, True, True)
    m_SkinWidth = ConvertPixelHimetric(m_SkinPicture.Width, True, True)
    
    RenderWidth = m_SkinWidth / 9
    
    If m_Hwnd_CH <> 0 And mUseSkinColumnHeader = True Then
        Dim LF As LOGFONT
        Dim hHeaderFont As Long

        LF.lfHeight = m_SkinHeight - 4
    
        hHeaderFont = CreateFontIndirect(LF)
        SendMessageLong m_Hwnd_CH, WM_SETFONT, hHeaderFont, True
        DeleteObject hHeaderFont
    End If
    
    Call RefreshColData
    
End Property

Public Property Get SkinPicture() As StdPicture
    Set SkinPicture = m_SkinPicture
End Property

Public Property Let IconAlingmentRight(ByVal New_Value As Boolean)
    m_IconAlingmentRight = New_Value
    RefreshColData
End Property

Public Property Get IconAlingmentRight() As Boolean
    IconAlingmentRight = m_IconAlingmentRight
End Property

Public Property Let TextNormalColor(ByVal New_Color As Long)
    m_TextNormalColor = New_Color
    RefreshColData
End Property

Public Property Get TextNormalColor() As Long
    TextNormalColor = m_TextNormalColor
End Property
Public Property Let TextResalteColor(ByVal New_Color As Long)
    m_TextResalteColor = New_Color
    RefreshColData
End Property

Public Property Get TextResalteColor() As Long
    TextResalteColor = m_TextResalteColor
End Property

Public Property Get HedersFontBlod() As Boolean
    HedersFontBlod = m_HedersFontBlod
End Property

Public Property Let HedersFontBlod(ByVal New_Value As Boolean)
  m_HedersFontBlod = New_Value
  RefreshColData
End Property

' deshabilita y habilita el redraw del Listview
Public Property Get Redraw() As Boolean
  Redraw = m_Redraw
End Property

Public Property Let Redraw(ByVal New_Value As Boolean)
  
  m_Redraw = New_Value
  
  If m_Hwnd_LV = 0 Then Exit Property
  
  If New_Value Then
     Call SendMessage(m_Hwnd_LV, WM_SETREDRAW, 1&, 0&)
  Else
     Call SendMessage(m_Hwnd_LV, WM_SETREDRAW, 0&, 0&)
  End If
  
End Property



'*********---------------*********------------**********------------**********----------
'*********---------------*********------------**********------------**********----------
'*********---------------*********------------**********------------**********----------

' asigna los dos Hwnd ( ColumnHeader y LV) e inicia el SubClass

Public Function SubClassListView(ListView As ListView) As Boolean
    
    'If m_Hwnd_LV <> 0 Then
    '   MsgBox "ya se ha ejecutado --->>   SubClassListView", vbInformation
    '   Exit Function
    'End If
    
    m_Hwnd_LV = ListView.hwnd
    ListView.View = lvwReport
    
    If ssc_Subclass(m_Hwnd_LV) Then
        ssc_AddMsg m_Hwnd_LV, WM_PAINT, MSG_BEFORE
        ssc_AddMsg m_Hwnd_LV, WM_LBUTTONUP, MSG_BEFORE
        ssc_AddMsg m_Hwnd_LV, 133, MSG_BEFORE
        
        
        Hook_ListView m_Hwnd_LV
        
        SubClassListView = True
        RefreshColData
    End If
        
End Function



Private Sub showError()
    MsgBox "Error. Antes de cualquier configuración se debe ejecutar el método SubClassListView", vbInformation
    End
End Sub



Property Get UseSkinColumnHeader() As Boolean
    UseSkinColumnHeader = mUseSkinColumnHeader
End Property

Property Let UseSkinColumnHeader(value As Boolean)
        
    mUseSkinColumnHeader = value
    
    If m_Hwnd_LV = 0 Then Exit Property
    
    If value = True Then
       setFontHeader m_Hwnd_LV
    Else
        SendMessageLong m_Hwnd_CH, WM_SETFONT, ViejaFuente, True
    End If
    Call RefreshColData
End Property

Private Sub setFontHeader(handle As Long)
    Dim LF As LOGFONT
    Dim hHeaderFont As Long
    
    m_Hwnd_CH = SendMessage(handle, LVM_GETHEADER, 0, 0)

    If ViejaFuente = 0 Then
        ViejaFuente = SendMessage(m_Hwnd_CH, WM_GETFONT, 0, 0)
    End If
    
    LF.lfHeight = m_SkinHeight - 4
    
    hHeaderFont = CreateFontIndirect(LF)
    SendMessageLong m_Hwnd_CH, WM_SETFONT, hHeaderFont, True
    DeleteObject hHeaderFont

    Indice = -1
End Sub


' función privada que asigna el handle del columnheader y le estblece los los mensajes a usar

Private Sub Hook_ListView(ByVal handle As Long)
    
    setFontHeader handle
    
    If ssc_Subclass(m_Hwnd_CH) Then
        ssc_AddMsg m_Hwnd_CH, WM_PAINT, MSG_BEFORE
        ssc_AddMsg m_Hwnd_CH, WM_MOUSEMOVE, MSG_BEFORE
        ssc_AddMsg m_Hwnd_CH, WM_LBUTTONUP, MSG_BEFORE
        ssc_AddMsg m_Hwnd_CH, WM_LBUTTONDOWN, MSG_BEFORE
        ssc_AddMsg m_Hwnd_CH, 132, MSG_BEFORE
    End If

    
End Sub


Private Sub Class_Initialize()
    UseSkinColumnHeader = True
    ReDim DataSubItems(0)
End Sub


Private Sub Class_Terminate()
    ssc_Terminate
    If TimerEnabled Then Call KillTimer(m_Hwnd_CH, ObjPtr(Me) + 1)
    scb_TerminateCallbacks
    
    SendMessageLong m_Hwnd_CH, WM_SETFONT, ViejaFuente, True
    
End Sub

' Función que termina el Subclassing llamando a Class_Terminate y refrescando el control
Function UnSubClassListView() As Boolean
    Call Class_Terminate
    Call RefreshColData
End Function



'-The following routines are exclusively for the ssc_subclass routines----------------------------
Private Function ssc_Subclass(ByVal lng_hWnd As Long, _
       Optional ByVal lParamUser As Long = 0, _
       Optional ByVal nOrdinal As Long = 1, _
       Optional ByVal oCallback As Object = Nothing, _
       Optional ByVal bIdeSafety As Boolean = True, _
       Optional ByVal bUnicode As Boolean = False) As Boolean 'Subclass the specified window handle

    '*************************************************************************************************
    '* lng_hWnd   - Handle of the window to subclass
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
    '* bUnicode - Optional, if True, Unicode API calls will be made to the window vs ANSI calls
    '*************************************************************************************************
    '* cSelfSub - self-subclassing class template
    '* Paul_Caton@hotmail.com
    '* Copyright free, use and abuse as you see fit.
    '*
    '* v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code............ 20060322
    '* v1.1 VirtualAlloc memory to prevent Data Execution Prevention faults on Win64......... 20060324
    '* v1.2 Thunk redesigned to handle unsubclassing and memory release...................... 20060325
    '* v1.3 Data array scrapped in favour of property accessors.............................. 20060405
    '* v1.4 Optional IDE protection added
    '*      User-defined callback parameter added
    '*      All user routines that pass in a hWnd get additional validation
    '*      End removed from zError.......................................................... 20060411
    '* v1.5 Added nOrdinal parameter to ssc_Subclass
    '*      Switched machine-code array from Currency to Long................................ 20060412
    '* v1.6 Added an optional callback target object
    '*      Added an IsBadCodePtr on the callback address in the thunk prior to callback..... 20060413
    '*************************************************************************************************
    ' Subclassing procedure must be declared identical to the one at the end of this class (Sample at Ordinal #1)

    ' \\LaVolpe - reworked routine a bit, revised the ASM to allow auto-unsubclass on WM_DESTROY
    Dim z_Sc(0 To IDX_UNICODE) As Long                 'Thunk machine-code initialised here
    Const CODE_LEN      As Long = 4 * IDX_UNICODE      'Thunk length in bytes
    
    Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES))  'Bytes to allocate per thunk, data + code + msg tables
    Const PAGE_RWX      As Long = &H40&                'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&              'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&              'Release allocated memory flag
    Const IDX_EBMODE    As Long = 3                    'Thunk data index of the EbMode function address
    Const IDX_CWP       As Long = 4                    'Thunk data index of the CallWindowProc function address
    Const IDX_SWL       As Long = 5                    'Thunk data index of the SetWindowsLong function address
    Const IDX_FREE      As Long = 6                    'Thunk data index of the VirtualFree function address
    Const IDX_BADPTR    As Long = 7                    'Thunk data index of the IsBadCodePtr function address
    Const IDX_OWNER     As Long = 8                    'Thunk data index of the Owner object's vTable address
    Const IDX_CALLBACK  As Long = 10                   'Thunk data index of the callback method address
    Const IDX_EBX       As Long = 16                   'Thunk code patch index of the thunk data
    Const GWL_WNDPROC   As Long = -4                   'SetWindowsLong WndProc index
    Const WNDPROC_OFF   As Long = &H38                 'Thunk offset to the WndProc execution address
    Const SUB_NAME      As String = "ssc_Subclass"     'This routine's name
    
    Dim nAddr         As Long
    Dim nID           As Long
    Dim nMyID         As Long

    If IsWindow(lng_hWnd) = 0 Then                      'Ensure the window handle is valid
        zError SUB_NAME, "Invalid window handle"
        Exit Function
    End If
    
    nMyID = GetCurrentProcessId                         'Get this process's ID
    GetWindowThreadProcessId lng_hWnd, nID              'Get the process ID associated with the window handle
    If nID <> nMyID Then                                'Ensure that the window handle doesn't belong to another process
        zError SUB_NAME, "Window handle belongs to another process"
        Exit Function
    End If
      
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    
    nAddr = zAddressOf(oCallback, nOrdinal)             'Get the address of the specified ordinal method
    If nAddr = 0 Then                                   'Ensure that we've found the ordinal method
        zError SUB_NAME, "Callback method not found"
        Exit Function
    End If
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    
    If z_ScMem <> 0 Then                                  'Ensure the allocation succeeded
  
        If z_scFunk Is Nothing Then Set z_scFunk = New Collection 'If this is the first time through, do the one-time initialization
        On Error GoTo CatchDoubleSub                              'Catch double subclassing
        z_scFunk.Add z_ScMem, "h" & lng_hWnd                    'Add the hWnd/thunk-address to the collection
        On Error GoTo 0
        
        ' \\Tai Chi Minh Ralph Eastwood - fixed bug where the MSG_AFTER was not being honored
        ' \\LaVolpe - modified thunks to allow auto-unsubclassing when WM_DESTROY received
        z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(16) = &H12345678: z_Sc(17) = &HF63103FF: z_Sc(18) = &H750C4339: z_Sc(19) = &H7B8B4A38: z_Sc(20) = &H95E82C: z_Sc(21) = &H7D810000: z_Sc(22) = &H228&: z_Sc(23) = &HC70C7500: z_Sc(24) = &H20443: z_Sc(25) = &H5E90000: z_Sc(26) = &H39000000: z_Sc(27) = &HF751475: z_Sc(28) = &H25E8&: z_Sc(29) = &H8BD23100: z_Sc(30) = &H6CE8307B: z_Sc(31) = &HFF000000: z_Sc(32) = &H10C2610B: z_Sc(33) = &HC53FF00: z_Sc(34) = &H13D&: z_Sc(35) = &H85BE7400: z_Sc(36) = &HE82A74C0: z_Sc(37) = &H2&: z_Sc(38) = &H75FFE5EB: z_Sc(39) = &H2C75FF30: z_Sc(40) = &HFF2875FF: z_Sc(41) = &H73FF2475: z_Sc(42) = &H1053FF24: z_Sc(43) = &H811C4589: z_Sc(44) = &H13B&: z_Sc(45) = &H39727500:
        z_Sc(46) = &H6D740473: z_Sc(47) = &H2473FF58: z_Sc(48) = &HFFFFFC68: z_Sc(49) = &H873FFFF: z_Sc(50) = &H891453FF: z_Sc(51) = &H7589285D: z_Sc(52) = &H3045C72C: z_Sc(53) = &H8000&: z_Sc(54) = &H8920458B: z_Sc(55) = &H4589145D: z_Sc(56) = &HC4816124: z_Sc(57) = &H4&: z_Sc(58) = &H8B1862FF: z_Sc(59) = &H853AE30F: z_Sc(60) = &H810D78C9: z_Sc(61) = &H4C7&: z_Sc(62) = &H28458B00: z_Sc(63) = &H2975AFF2: z_Sc(64) = &H2873FF52: z_Sc(65) = &H5A1C53FF: z_Sc(66) = &H438D1F75: z_Sc(67) = &H144D8D34: z_Sc(68) = &H1C458D50: z_Sc(69) = &HFF3075FF: z_Sc(70) = &H75FF2C75: z_Sc(71) = &H873FF28: z_Sc(72) = &HFF525150: z_Sc(73) = &H53FF2073: z_Sc(74) = &HC328C328
        
        z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
        z_Sc(IDX_INDEX) = lng_hWnd                                               'Store the window handle in the thunk data
        z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
        z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
        z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
        z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
        z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
        
        ' \\LaVolpe - validate unicode request & cache unicode usage
        If bUnicode Then bUnicode = (IsWindowUnicode(lng_hWnd) <> 0&)
        z_Sc(IDX_UNICODE) = bUnicode                                            'Store whether the window is using unicode calls or not
        
        ' \\LaVolpe - added extra parameter "bUnicode" to the zFnAddr calls
        z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree", bUnicode)           'Store the VirtualFree function address in the thunk data
        z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)        'Store the IsBadCodePtr function address in the thunk data
        
        Debug.Assert zInIDE
        If bIdeSafety = True And z_IDEflag = 1 Then                             'If the user wants IDE protection
            z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode)                'Store the EbMode function address in the thunk data
        End If
    
        ' \\LaVolpe - use ANSI for non-unicode usage, else use WideChar calls
        If bUnicode Then
            z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcW", bUnicode)          'Store CallWindowProc function address in the thunk data
            z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongW", bUnicode)           'Store the SetWindowLong function address in the thunk data
            z_Sc(IDX_UNICODE) = 1
            RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory
            nAddr = SetWindowLongW(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
        Else
            z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA", bUnicode)          'Store CallWindowProc function address in the thunk data
            z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA", bUnicode)           'Store the SetWindowLong function address in the thunk data
            RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory
            nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
        End If
        If nAddr = 0 Then                                                           'Ensure the new WndProc was set correctly
            zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
            GoTo ReleaseMemory
        End If
        'Store the original WndProc address in the thunk data
        RtlMoveMemory z_ScMem + IDX_WNDPROC * 4, VarPtr(nAddr), 4&              ' z_Sc(IDX_WNDPROC) = nAddr
        ssc_Subclass = True                                                     'Indicate success
    Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
    End If

    Exit Function                                                             'Exit ssc_Subclass
    
CatchDoubleSub:
    zError SUB_NAME, "Window handle is already subclassed"
      
ReleaseMemory:
    VirtualFree z_ScMem, 0, MEM_RELEASE                                       'ssc_Subclass has failed after memory allocation, so release the memory
End Function

'Terminate all subclassing
Private Sub ssc_Terminate()
    ' can be made public. Releases all subclassing
    ' can be removed and zTerminateThunks can be called directly
    zTerminateThunks SubclassThunk
End Sub

'UnSubclass the specified window handle
Private Sub ssc_UnSubclass(ByVal lng_hWnd As Long)
    ' can be made public. Releases a specific subclass
    ' can be removed and zUnThunk can be called directly
    zUnThunk lng_hWnd, SubclassThunk
End Sub

'Add the message value to the window handle's specified callback table
Private Sub ssc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    ' Note: can be removed if not needed and zAddMsg can be called directly
    If IsBadCodePtr(zMap_VFunction(lng_hWnd, SubclassThunk)) = 0 Then                 'Ensure that the thunk hasn't already released its memory
        If When And MSG_BEFORE Then                                             'If the message is to be added to the before original WndProc table...
            zAddMsg uMsg, IDX_BTABLE                                              'Add the message to the before table
        End If
        If When And MSG_AFTER Then                                              'If message is to be added to the after original WndProc table...
            zAddMsg uMsg, IDX_ATABLE                                              'Add the message to the after table
        End If
    End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub ssc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    ' Note: can be removed if not needed and zDelMsg can be called directly
    If IsBadCodePtr(zMap_VFunction(lng_hWnd, SubclassThunk)) = 0 Then                'Ensure that the thunk hasn't already released its memory
        If When And MSG_BEFORE Then                                             'If the message is to be deleted from the before original WndProc table...
            zDelMsg uMsg, IDX_BTABLE                                              'Delete the message from the before table
        End If
        If When And MSG_AFTER Then                                              'If the message is to be deleted from the after original WndProc table...
            zDelMsg uMsg, IDX_ATABLE                                              'Delete the message from the after table
        End If
    End If
End Sub

'Call the original WndProc
Private Function ssc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' Note: can be removed if you do not use this function inside of your window procedure
    If IsBadCodePtr(zMap_VFunction(lng_hWnd, SubclassThunk)) = 0 Then            'Ensure that the thunk hasn't already released its memory
        If zData(IDX_UNICODE) Then
            ssc_CallOrigWndProc = CallWindowProcW(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        Else
            ssc_CallOrigWndProc = CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        End If
    End If
End Function

'Get the subclasser lParamUser callback parameter
Private Function zGet_lParamUser(ByVal hWnd_Hook_ID As Long, vType As eThunkType) As Long
    'Note: can be removed if you never need to retrieve/update your user-defined paramter. See ssc_Subclass
    If vType <> CallbackThunk Then
        If IsBadCodePtr(zMap_VFunction(hWnd_Hook_ID, vType)) = 0 Then        'Ensure that the thunk hasn't already released its memory
            zGet_lParamUser = zData(IDX_PARM_USER)                                'Get the lParamUser callback parameter
        End If
    End If
End Function

'Let the subclasser lParamUser callback parameter
Private Sub zSet_lParamUser(ByVal hWnd_Hook_ID As Long, vType As eThunkType, newValue As Long)
    'Note: can be removed if you never need to retrieve/update your user-defined paramter. See ssc_Subclass
    If vType <> CallbackThunk Then
        If IsBadCodePtr(zMap_VFunction(hWnd_Hook_ID, vType)) = 0 Then          'Ensure that the thunk hasn't already released its memory
            zData(IDX_PARM_USER) = newValue                                         'Set the lParamUser callback parameter
        End If
    End If
End Sub

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)
    Dim nCount As Long                                                        'Table entry count
    Dim nBase  As Long                                                        'Remember z_ScMem
    Dim i      As Long                                                        'Loop index
    
    nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                                   'Map zData() to the specified table
    
    If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
    Else
        nCount = zData(0)                                                       'Get the current table entry count
        If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
            zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
            GoTo Bail
        End If
    
        For i = 1 To nCount                                                     'Loop through the table entries
            If zData(i) = 0 Then                                                  'If the element is free...
                zData(i) = uMsg                                                     'Use this element
                GoTo Bail                                                           'Bail
            ElseIf zData(i) = uMsg Then                                           'If the message is already in the table...
                GoTo Bail                                                           'Bail
            End If
        Next i                                                                  'Next message table entry
    
        nCount = i                                                              'On drop through: i = nCount + 1, the new table entry count
        zData(nCount) = uMsg                                                    'Store the message in the appended table entry
    End If
    
    zData(0) = nCount                                                         'Store the new table entry count
Bail:
    z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)
    Dim nCount As Long                                                        'Table entry count
    Dim nBase  As Long                                                        'Remember z_ScMem
    Dim i      As Long                                                        'Loop index
    
    nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                                   'Map zData() to the specified table
    
    If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
        zData(0) = 0                                                            'Zero the table entry count
    Else
        nCount = zData(0)                                                       'Get the table entry count
        
        For i = 1 To nCount                                                     'Loop through the table entries
            If zData(i) = uMsg Then                                               'If the message is found...
                zData(i) = 0                                                        'Null the msg value -- also frees the element for re-use
                GoTo Bail                                                           'Bail
            End If
        Next i                                                                  'Next message table entry
        
        zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
    End If
      
Bail:
    z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'-SelfCallback code------------------------------------------------------------------------------------
'-The following routines are exclusively for the scb_SetCallbackAddr routines----------------------------
Private Function scb_SetCallbackAddr(ByVal nParamCount As Long, _
       Optional ByVal nOrdinal As Long = 1, _
       Optional ByVal oCallback As Object = Nothing, _
       Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk
    '*************************************************************************************************
    '* nParamCount  - The number of parameters that will callback
    '* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
    '* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety   - Optional, set to false to disable IDE protection.
    '*************************************************************************************************
    ' Callback procedure must return a Long even if, per MSDN, the callback procedure is a Sub vs Function
    ' The number of parameters are dependent on the individual callback procedures
    
    Const MEM_LEN     As Long = IDX_CALLBACKORDINAL * 4 + 4     'Memory bytes required for the callback thunk
    Const PAGE_RWX    As Long = &H40&                           'Allocate executable memory
    Const MEM_COMMIT  As Long = &H1000&                         'Commit allocated memory
    Const SUB_NAME      As String = "scb_SetCallbackAddr"       'This routine's name
    Const INDX_OWNER    As Long = 0
    Const INDX_CALLBACK As Long = 1
    Const INDX_EBMODE   As Long = 2
    Const INDX_BADPTR   As Long = 3
    Const INDX_EBX      As Long = 5
    Const INDX_PARAMS   As Long = 12
    Const INDX_PARAMLEN As Long = 17

    Dim z_Cb()    As Long    'Callback thunk array
    Dim nCallback As Long
      
    If z_cbFunk Is Nothing Then
        Set z_cbFunk = New Collection           'If this is the first time through, do the one-time initialization
    Else
        On Error Resume Next                    'Catch already initialized?
        z_ScMem = z_cbFunk.item("h" & nOrdinal) 'Test it
        If Err = 0 Then
            scb_SetCallbackAddr = z_ScMem + 16  'we had this one, just reference it
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    If nParamCount < 0 Then                     ' validate parameters
        zError SUB_NAME, "Invalid Parameter count"
        Exit Function
    End If
    
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    nCallback = zAddressOf(oCallback, nOrdinal)         'Get the callback address of the specified ordinal
    If nCallback = 0 Then
        zError SUB_NAME, "Callback address not found."
        Exit Function
    End If
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
        
    If z_ScMem = 0& Then
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError  ' oops
        Exit Function
    End If
    z_cbFunk.Add z_ScMem, "h" & nOrdinal                  'Add the callback/thunk-address to the collection
        
    ReDim z_Cb(0 To IDX_CALLBACKORDINAL) As Long          'Allocate for the machine-code array
    
    ' Create machine-code array
    z_Cb(4) = &HBB60E089: z_Cb(6) = &H73FFC589: z_Cb(7) = &HC53FF04: z_Cb(8) = &H7B831F75: z_Cb(9) = &H20750008: z_Cb(10) = &HE883E889: z_Cb(11) = &HB9905004: z_Cb(13) = &H74FF06E3: z_Cb(14) = &HFAE2008D: z_Cb(15) = &H53FF33FF: z_Cb(16) = &HC2906104: z_Cb(18) = &H830853FF: z_Cb(19) = &HD87401F8: z_Cb(20) = &H4589C031: z_Cb(21) = &HEAEBFC
    
    z_Cb(INDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", False)
    z_Cb(INDX_OWNER) = ObjPtr(oCallback)                  'Set the Owner
    z_Cb(INDX_CALLBACK) = nCallback                       'Set the callback address
    z_Cb(IDX_CALLBACKORDINAL) = nOrdinal                  'Cache ordinal used for zTerminateThunks
      
    Debug.Assert zInIDE
    If bIdeSafety = True And z_IDEflag = 1 Then             'If the user wants IDE protection
        z_Cb(INDX_EBMODE) = zFnAddr("vba6", "EbMode", False)  'EbMode Address
    End If
        
    z_Cb(INDX_PARAMS) = nParamCount                         'Set the parameter count
    z_Cb(INDX_PARAMLEN) = nParamCount * 4                   'Set the number of stck bytes to release on thunk return
      
    '\\LaVolpe - redirect address to proper location in virtual memory. Was: z_Cb(INDX_EBX) = VarPtr(z_Cb(INDX_OWNER))
    z_Cb(INDX_EBX) = z_ScMem                                'Set the data address relative to virtual memory pointer
      
    RtlMoveMemory z_ScMem, VarPtr(z_Cb(INDX_OWNER)), MEM_LEN 'Copy thunk code to executable memory
    scb_SetCallbackAddr = z_ScMem + 16                       'Thunk code start address
    
End Function

Private Sub scb_ReleaseCallback(ByVal nOrdinal As Long)
    ' can be made public. Releases a specific callback
    ' can be removed and zUnThunk can be called directly
    zUnThunk nOrdinal, CallbackThunk
End Sub
Private Sub scb_TerminateCallbacks()
    ' can be made public. Releases all callbacks
    ' can be removed and zTerminateThunks can be called directly
    zTerminateThunks CallbackThunk
End Sub


'========================================================================
' COMMON USE ROUTINES
'-The following routines are used for each of the three types of thunks
'========================================================================

'Map zData() to the thunk address for the specified window handle
Private Function zMap_VFunction(ByVal vFuncTarget As Long, vType As eThunkType) As Long
    
    ' vFuncTarget is one of the following, depending on vType
    '   - Subclassing:  the hWnd of the window subclassed
    '   - Hooking:      the hook type created
    '   - Callbacks:    the ordinal of the callback
    
    Dim thunkCol As Collection
    
    If vType = CallbackThunk Then
        Set thunkCol = z_cbFunk
    ElseIf vType = HookThunk Then
        Set thunkCol = z_hkFunk
    ElseIf vType = SubclassThunk Then
        Set thunkCol = z_scFunk
    Else
        zError "zMap_Vfunction", "Invalid thunk type passed"
        Exit Function
    End If
    
    If thunkCol Is Nothing Then
        zError "zMap_VFunction", "Thunk hasn't been initialized"
    Else
        On Error GoTo Catch
        z_ScMem = thunkCol("h" & vFuncTarget)                    'Get the thunk address
        zMap_VFunction = z_ScMem
    End If
    Exit Function                                               'Exit returning the thunk address
    
Catch:
    zError "zMap_VFunction", "Thunk type for ID of " & vFuncTarget & " does not exist"
End Function

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
    ' \\LaVolpe -  Note. These two lines can be rem'd out if you so desire. But don't remove the routine
    App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
    MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String, ByVal asUnicode As Boolean) As Long
    If asUnicode Then
        zFnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)         'Get the specified procedure address
    Else
        zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                 'Get the specified procedure address
    End If
    Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
    ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    ' Note: used both in subclassing and hooking routines
    Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
    Dim bVal  As Byte
    Dim nAddr As Long                                                         'Address of the vTable
    Dim i     As Long                                                         'Loop index
    Dim j     As Long                                                         'Loop limit
  
    RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
    If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
        If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
            ' \\LaVolpe - Added propertypage offset
            If Not zProbe(nAddr + &H710, i, bSub) Then                            'Probe for a PropertyPage method
                If Not zProbe(nAddr + &H7A4, i, bSub) Then                          'Probe for a UserControl method
                    Exit Function                                                   'Bail...
                End If
            End If
        End If
    End If
  
    i = i + 4                                                                 'Bump to the next entry
    j = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
    Do While i < j
        RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
        If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
            Exit Do                                                               'Bad method signature, quit loop
        End If

        RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
        If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
            Exit Do                                                               'Bad method signature, quit loop
        End If
    
        i = i + 4                                                               'Next vTable entry
    Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
    Dim bVal    As Byte
    Dim nAddr   As Long
    Dim nLimit  As Long
    Dim nEntry  As Long
  
    nAddr = nStart                                                            'Start address
    nLimit = nAddr + 32                                                       'Probe eight entries
    Do While nAddr < nLimit                                                   'While we've not reached our probe depth
        RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
        If nEntry <> 0 Then                                                     'If not an implemented interface
            RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
            If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
                nMethod = nAddr                                                     'Store the vTable entry
                bSub = bVal                                                         'Store the found method signature
                zProbe = True                                                       'Indicate success
                Exit Do                                                             'Return
            End If
        End If
    
        nAddr = nAddr + 4                                                       'Next vTable entry
    Loop
End Function

Private Function zInIDE() As Long
    ' This is only run in IDE; it is never run when compiled
    z_IDEflag = 1
    zInIDE = z_IDEflag
End Function

Private Property Get zData(ByVal nIndex As Long) As Long
    ' retrieves stored value from virtual function's memory location
    RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
    ' sets value in virtual function's memory location
    RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property

Private Sub zUnThunk(ByVal thunkID As Long, ByVal vType As eThunkType)
    ' Releases a specific subclass, hook or callback
    ' thunkID depends on vType:
    '   - Subclassing:  the hWnd of the window subclassed
    '   - Hooking:      the hook type created
    '   - Callbacks:    the ordinal of the callback

    Const IDX_SHUTDOWN  As Long = 1
    Const MEM_RELEASE As Long = &H8000&                                'Release allocated memory flag
    
    If zMap_VFunction(thunkID, vType) Then
        Select Case vType
            Case SubclassThunk
                If IsBadCodePtr(z_ScMem) = 0 Then       'Ensure that the thunk hasn't already released its memory
                    zData(IDX_SHUTDOWN) = 1             'Set the shutdown indicator
                    zDelMsg ALL_MESSAGES, IDX_BTABLE    'Delete all before messages
                    zDelMsg ALL_MESSAGES, IDX_ATABLE    'Delete all after messages
                    '\\LaVolpe - Force thunks to replace original window procedure handle. Without this, app can crash when a window is subclassed multiple times simultaneously
                    If zData(IDX_UNICODE) Then          'Force window procedure handle to be replaced
                        SendMessageW thunkID, 0&, 0&, ByVal 0&
                    Else
                        SendMessageA thunkID, 0&, 0&, ByVal 0&
                    End If
                End If
                z_scFunk.Remove "h" & thunkID           'Remove the specified thunk from the collection
            Case HookThunk
                If IsBadCodePtr(z_ScMem) = 0 Then       'Ensure that the thunk hasn't already released its memory
                    zData(IDX_SHUTDOWN) = 1             'Set the shutdown indicator
                    zData(IDX_ATABLE) = 0               ' want no more After messages
                    zData(IDX_BTABLE) = 0               ' want no more Before messages
                End If
                z_hkFunk.Remove "h" & thunkID           'Remove the specified thunk from the collection
            Case CallbackThunk
                If IsBadCodePtr(z_ScMem) = 0 Then       'Ensure that the thunk hasn't already released its memory
                    VirtualFree z_ScMem, 0, MEM_RELEASE 'Release allocated memory
                End If
                z_cbFunk.Remove "h" & thunkID           'Remove the specified thunk from the collection
        End Select
    End If

End Sub

Private Sub zTerminateThunks(ByVal vType As eThunkType)
    ' Removes all thunks of a specific type: subclassing, hooking or callbacks
    Dim i As Long
    Dim thunkCol As Collection
    
    Select Case vType
        Case SubclassThunk
            Set thunkCol = z_scFunk
        Case HookThunk
            Set thunkCol = z_hkFunk
        Case CallbackThunk
            Set thunkCol = z_cbFunk
        Case Else
            Exit Sub
    End Select
    
    If Not (thunkCol Is Nothing) Then                 'Ensure that hooking has been started
        With thunkCol
            For i = .Count To 1 Step -1                   'Loop through the collection of hook types in reverse order
                z_ScMem = .item(i)                          'Get the thunk address
                If IsBadCodePtr(z_ScMem) = 0 Then           'Ensure that the thunk hasn't already released its memory
                    Select Case vType
                        Case SubclassThunk
                            zUnThunk zData(IDX_INDEX), SubclassThunk     'Unsubclass
                        Case HookThunk
                            zUnThunk zData(IDX_INDEX), HookThunk             'Unhook
                        Case CallbackThunk
                            zUnThunk zData(IDX_CALLBACKORDINAL), CallbackThunk ' release callback
                    End Select
                End If
            Next i                                        'Next member of the collection
        End With
        Set thunkCol = Nothing                         'Destroy the hook/thunk-address collection
    End If

End Sub


' Función que redibuja los gráficos para cada columna
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub DrawGraphColumn(hwnd As Long)
    Dim rec As RECT
    Dim PS As PAINTSTRUCT
    Dim r As RECT
    Dim DC As Long
    Dim Count As Long
    Dim TopIndex As Long
    Dim hRPen As Long
    Dim hBrush As Long
    Dim ItemValue As Currency
    Dim value As Long
    Dim LB As LOGBRUSH
    Dim tm As TEXTMETRIC
    Dim rec2 As RECT
    Dim j As Long
    Dim i As Long
    
    Count = SendMessage(hwnd, LVM_GETITEMCOUNT, 0, 0)
    GetClientRect hwnd, r
    
    Call BeginPaint(hwnd, PS)
    Call InvalidateRect(hwnd, r, 0)
    
    DC = PS.hdc
    
    TopIndex = SendMessage(hwnd, LVM_GETTOPINDEX, 0, 0)
    
    For j = 1 To UBound(DataSubItems)
        For i = TopIndex To Count - 1
            
            If (DataSubItems(j).MaxValue = -99999999999999#) Or _
               (DataSubItems(j).MinValue = 99999999999999#) Then
               Exit Sub
            End If
            
            If DataSubItems(j).displayGraph Then
            
            rec.Top = DataSubItems(j).SubItemsIndex
            rec.Left = LVIR_LABEL
            
            SendMessage hwnd, LVM_GETSUBITEMRECT, i, rec
            
            With rec2
                .Left = rec.Left
                .Top = rec.Top + 1
                .Right = rec.Right
                .Bottom = rec.Bottom - 1
            End With
            
            ValidateRect hwnd, rec2
            
            With LB
                .lbColor = DataSubItems(j).BackColor
                .lbStyle = 2
                .lbHatch = 6
            End With
                
            hBrush = CreateBrushIndirect(LB)
            hRPen = CreatePen(0, 1, DataSubItems(j).BorderColor)
                
            DeleteObject SelectObject(DC, hBrush)
            DeleteObject SelectObject(DC, hRPen)
            Rectangle DC, rec.Left, rec.Top + 1, rec.Right, rec.Bottom - 1
                
            DeleteObject hBrush
            DeleteObject hRPen
                
            ItemValue = GetListItem(hwnd, i, DataSubItems(j).SubItemsIndex)
            
            Dim ItemValue2 As String
            ItemValue2 = ItemValue
            
            If (DataSubItems(j).MaxValue = DataSubItems(j).MinValue) Then
               ItemValue = 100
            Else
                ItemValue = Round((ItemValue - DataSubItems(j).MinValue) * 100 / (DataSubItems(j).MaxValue - DataSubItems(j).MinValue))
            End If
                
            If ItemValue > 0 Then
               With LB
                  .lbColor = DataSubItems(j).ForeColor
                  .lbStyle = 2
                  .lbHatch = 6
               End With
                    
               hBrush = CreateBrushIndirect(LB)
               hRPen = CreatePen(0, 0, DataSubItems(j).ForeColor)
                   
               DeleteObject SelectObject(DC, hBrush)
               DeleteObject SelectObject(DC, hRPen)
            
                    
               value = (ItemValue * (rec.Right - rec.Left)) / 100
                    
               If ItemValue < 100 Then
                  Rectangle DC, rec.Left + 1, rec.Top + 2, rec.Left + value, rec.Bottom - 2
               Else
                  Rectangle DC, rec.Left + 1, rec.Top + 2, rec.Right - 1, rec.Bottom - 2
               End If
                    
               DeleteObject hBrush
               DeleteObject hRPen
            End If
                
            If DataSubItems(j).DisplayValue Then
               
               GetTextMetrics DC, tm
               
               rec.Top = (rec.Top - (rec.Top - rec.Bottom) / 2) - ((tm.tmHeight) / 2)
                    
               SetBkMode DC, TRANSPARENT
               SetTextColor DC, DataSubItems(j).ForeColorText
                    
                    
               Dim sItemValue As String
               
               If DataSubItems(j).DisplayPercent Then
                  sItemValue = CStr(Round(ItemValue)) & " %"
               ElseIf DataSubItems(j).FieldCurrency Then
                  sItemValue = FormatCurrency(ItemValue2, DataSubItems(j).NumDigistsAfterDecimal)
               Else
                  sItemValue = FormatNumber(ItemValue2, DataSubItems(j).NumDigistsAfterDecimal)
               End If
               
               
               Select Case DataSubItems(j).AlignTextValue
                Case 0
                  sItemValue = " " & sItemValue
                Case 1
                  sItemValue = Trim(sItemValue)
                Case 2
                  sItemValue = sItemValue & " "
               End Select
               
               DrawTextEx DC, sItemValue, Len(sItemValue), rec, DT_WORDBREAK Or DataSubItems(j).AlignTextValue, ByVal 0&
            
            End If
                
            If r.Bottom < rec.Bottom Then
                Exit For
            End If

        End If
        Next i
    Next j
    DeleteDC DC
    Call EndPaint(hwnd, PS)
End Sub


' Funcio´n que redibuja el Skin en los headers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DrawColumHeader( _
    hwnd As Long, _
    ByVal cHITTES As Integer)

Dim PS As PAINTSTRUCT
Dim rec As RECT
Dim Count As Integer
Dim i As Integer
Dim TextAlingment As Long
Dim mStr As String
Dim hdc As Long
Dim LF As LOGFONT
Dim hImageList As Long
Dim IconSize As Size
Dim SkinH As Long
Dim tSize As Size
Dim iSize As Size
Dim HavIcon As Boolean
Dim HedersFomat As Integer


Dim hHeaderFont As Long
Dim hDCMemory As Long
Dim hBmp As Long
Dim DC As Long

    
    If mUseSkinColumnHeader = False Then Exit Sub
    
    SkinH = m_SkinHeight
    
    Call BeginPaint(hwnd, PS)

    hImageList = SendMessageLong(hwnd, HDM_GETIMAGELIST, 0, 0)
    Call ImageList_GetIconSize(hImageList, IconSize.cx, IconSize.cy)


    DC = GetDC(0)
    hDCMemory = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(DC, PS.rcPaint.Right, PS.rcPaint.Bottom)
    Call SelectObject(hDCMemory, hBmp)


    StretchBlt hDCMemory, 0, 0, PS.rcPaint.Right, PS.rcPaint.Bottom, m_hBmp, 1, 0, RenderWidth, SkinH, ScrCopy
    Count = SendMessage(hwnd, HDM_GETITEMCOUNT, 0, 0)
    SetBkMode hDCMemory, TRANSPARENT

    If Not m_HedersFontBlod = True Then
        LF.lfHeight = 14
        LF.lfWeight = FW_NORMAL
        hHeaderFont = CreateFontIndirect(LF)
        Call SelectObject(hDCMemory, hHeaderFont)
        DeleteObject hHeaderFont
    End If

    For i = 0 To Count - 1
        SendMessage hwnd, HDM_GETITEMRECT, i, rec
        mStr = GetListviewColumn(hwnd, i)
    
        Select Case ColItem.fmt
            Case 0
                HavIcon = False: HedersFomat = 0
            Case HDF_STRING + HDF_IMAGE + HDF_LEFT
                HavIcon = True: HedersFomat = 1
            Case HDF_STRING + HDF_IMAGE + HDF_CENTER
                HavIcon = True: HedersFomat = 2
            Case HDF_STRING + HDF_IMAGE + HDF_RIGHT
                HavIcon = True: HedersFomat = 3
            Case HDF_STRING + HDF_LEFT
                HavIcon = False: HedersFomat = 4
            Case HDF_STRING + HDF_CENTER
                HavIcon = False: HedersFomat = 5
            Case HDF_STRING + HDF_RIGHT
                HavIcon = False: HedersFomat = 6
        End Select
    
    
        If HavIcon Then
            mStr = GetTextShort(PS.hdc, mStr, rec.Right - rec.Left - 18 - IconSize.cx)
        Else
            mStr = GetTextShort(PS.hdc, mStr, rec.Right - rec.Left - 18)
        End If
    
        
    
    
    
        GetTextExtentPoint32 hDCMemory, mStr, Len(mStr), tSize

        
        Select Case HedersFomat
            Case 0
                tSize.cx = rec.Left + 9
                
            Case 1
                If m_IconAlingmentRight Then
                    iSize.cx = rec.Left + tSize.cx + 9
                    tSize.cx = rec.Left + 9
                Else
                    tSize.cx = rec.Left + 9 + IconSize.cx + 9
                    iSize.cx = rec.Left + 9
                End If
            Case 2
                If m_IconAlingmentRight Then
                    iSize.cx = (rec.Left + (rec.Right - rec.Left) / 2) + (tSize.cx / 2) - 9
                    tSize.cx = (rec.Left + (rec.Right - rec.Left) / 2) - (tSize.cx / 2) - 9
                Else
                    tSize.cx = (rec.Left + (rec.Right - rec.Left) / 2) - (tSize.cx / 2) + IconSize.cx - 9
                    iSize.cx = tSize.cx - IconSize.cx - 9
                End If
            Case 3
                If m_IconAlingmentRight Then
                    iSize.cx = rec.Right - IconSize.cx - 9
                    tSize.cx = rec.Right - tSize.cx - IconSize.cx - 9
                Else
                    tSize.cx = rec.Right - tSize.cx - 9
                    iSize.cx = tSize.cx - IconSize.cx - 9
                End If
            Case 4

                tSize.cx = rec.Left + 9
            
            Case 5
                tSize.cx = (rec.Left + (rec.Right - rec.Left) / 2) - (tSize.cx / 2)
            
            Case 6

                tSize.cx = rec.Right - tSize.cx - 9
        
        End Select
        
        tSize.cy = rec.Bottom / 2 - tSize.cy / 2
        iSize.cy = rec.Bottom / 2 - IconSize.cy / 2
        
        
        If HavIcon And iSize.cx < rec.Left Or IconSize.cx > (rec.Right - rec.Left) - 5 Then
            HavIcon = False
        End If
        
        If i = cHITTES And IsDrag = False Then
    
            SetTextColor hDCMemory, m_TextResalteColor
        
            If MouseDown = True Then
                StretchBlt hDCMemory, rec.Left + RenderWidth, 0, rec.Right - rec.Left - (RenderWidth * 2), rec.Bottom, m_hBmp, 7 * RenderWidth, 0, RenderWidth, SkinH, ScrCopy
                If HavIcon Then ImageList_Draw hImageList, ColItem.iImage, hDCMemory, iSize.cx + 1, iSize.cy + 1, ILD_TRANSPARENT Or ILD_SELECTED
                TextOut hDCMemory, tSize.cx + 1, tSize.cy + 1, mStr, Len(mStr)
                StretchBlt hDCMemory, rec.Left, 0, RenderWidth, rec.Bottom, m_hBmp, 6 * RenderWidth, 0, RenderWidth, SkinH, ScrCopy
                StretchBlt hDCMemory, rec.Right - RenderWidth, 0, RenderWidth, rec.Bottom, m_hBmp, 8 * RenderWidth, 0, RenderWidth, SkinH, ScrCopy
            
            Else
                StretchBlt hDCMemory, rec.Left + RenderWidth, 0, rec.Right - rec.Left - (RenderWidth * 2), rec.Bottom, m_hBmp, 4 * RenderWidth, 0, RenderWidth, SkinH, ScrCopy
                If HavIcon Then ImageList_Draw hImageList, ColItem.iImage, hDCMemory, iSize.cx, iSize.cy, ILD_TRANSPARENT
                TextOut hDCMemory, tSize.cx, tSize.cy, mStr, Len(mStr)
                StretchBlt hDCMemory, rec.Left, 0, RenderWidth, rec.Bottom, m_hBmp, 3 * RenderWidth, 0, RenderWidth, SkinH, ScrCopy
                StretchBlt hDCMemory, rec.Right - RenderWidth, 0, RenderWidth, rec.Bottom, m_hBmp, 5 * RenderWidth, 0, RenderWidth, SkinH, ScrCopy
            End If
        
        Else
            SetTextColor hDCMemory, m_TextNormalColor
            If HavIcon Then ImageList_Draw hImageList, ColItem.iImage, hDCMemory, iSize.cx, iSize.cy, ILD_TRANSPARENT
            TextOut hDCMemory, tSize.cx, tSize.cy, mStr, Len(mStr)
            StretchBlt hDCMemory, rec.Left, 0, RenderWidth, rec.Bottom, m_hBmp, 0, 0, RenderWidth, SkinH, ScrCopy
            StretchBlt hDCMemory, rec.Right - RenderWidth, 0, RenderWidth, rec.Bottom, m_hBmp, 2 * RenderWidth, 0, RenderWidth, SkinH, ScrCopy

        End If
    
    Next
    
    StretchBlt PS.hdc, 0, 0, PS.rcPaint.Right, PS.rcPaint.Bottom, _
               hDCMemory, 0, 0, PS.rcPaint.Right, PS.rcPaint.Bottom, ScrCopy
    
    Call EndPaint(hwnd, PS)
    
    DeleteObject hBmp
    DeleteDC DC
    DeleteDC hDCMemory
End Sub

' Función que acorta el texto de las columnas
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function GetTextShort( _
    hdc As Long, _
    Text As String, _
    AreaWidth As Long) As String

    Dim tSize As Size
    
    Dim Chr As Integer
    Dim TresPuntos As Long
    
    GetTextExtentPoint32 hdc, "...", 3, tSize
    TresPuntos = tSize.cx
    
    GetTextExtentPoint32 hdc, Text, Len(Text), tSize
    
    If tSize.cx > AreaWidth Then
        Do While tSize.cx > AreaWidth - TresPuntos
            If Len(Text) = 1 Then Exit Function
            Text = Left(Text, Len(Text) - 1)
            
            GetTextExtentPoint32 hdc, Text, Len(Text), tSize
        Loop
        GetTextShort = Text & "..."
    Else
        GetTextShort = Text
    End If

End Function

' Función que retorna el caption de los columnHeaders
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function GetListviewColumn( _
    ByVal hwnd As Long, _
    ByVal pColumn As Long) As String
    
    On Error Resume Next
    Dim result As Long
    
    Dim strColText As String
    
    strColText = Space(32)
    ColItem.mask = HDI_TEXT Or HDI_ORDER Or HDI_FORMAT Or HDI_IMAGE Or HDI_BITMAP Or HDI_LPARAM
    
    ColItem.pszText = strColText
    ColItem.cchTextMax = Len(strColText)
    
    result = SendMessage(hwnd, HDM_GETITEMA, pColumn, ColItem)
    
    If result > 0 Then
        GetListviewColumn = Left(ColItem.pszText, InStr(ColItem.pszText, vbNullChar) - 1)
    End If

End Function


Private Function ConvertPixelHimetric( _
    ByVal inValue As Long, _
    ByVal ToPix As Boolean, _
    inXAxis As Boolean) As Long
    
    Dim TempIC As Long, GDCFlag As Long

    Const HimetricInch As Long = 2540
    
    TempIC = CreateIC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
    
    If (TempIC) Then
        If (inXAxis) Then GDCFlag = LOGPIXELSX Else GDCFlag = LOGPIXELSY
        
        If (ToPix) Then
            ConvertPixelHimetric = MulDiv(inValue, GetDeviceCaps(TempIC, GDCFlag), HimetricInch)
        Else
            ConvertPixelHimetric = MulDiv(inValue, HimetricInch, GetDeviceCaps(TempIC, GDCFlag))
        End If
        Call DeleteDC(TempIC)
    End If
End Function


' timer para quitar el skin cuando sale el mouse fuera del listview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TimerProc( _
       ByVal hwnd As Long, _
       ByVal tMsg As Long, _
       ByVal TimerID As Long, _
       ByVal tickCount As Long) As Long

    
    Dim pt As POINTAPI
    GetCursorPos pt
    If WindowFromPoint(pt.X, pt.Y) <> hwnd Then
        Indice = -1
        IsDrag = False
        SendMessage GetParent(hwnd), WM_PAINT, 0&, 0&
        Call KillTimer(hwnd, ObjPtr(Me) + 1)
        TimerEnabled = False
    End If
End Function


' Mensajes
''''''''''''''''''''''''''''''''''''''''''''''

'- ordinal #1
Private Sub WndProc(ByVal bBefore As Boolean, _
       ByRef bHandled As Boolean, _
       ByRef lReturn As Long, _
       ByVal hwnd As Long, _
       ByVal uMsg As Long, _
       ByVal wParam As Long, _
       ByVal lParam As Long, _
       ByRef lParamUser As Long)

    Dim hti As HDHITTESTINFO
    Dim mLH As TLoHiLong, mAL As TAllLong
    
Dim PS As PAINTSTRUCT
Dim rec As RECT
    
'Case ListView
'------------------------------
If hwnd = m_Hwnd_LV Then
    
    If uMsg = WM_PAINT Then
        Call DrawGraphColumn(hwnd)
    End If

    If uMsg = 133 Then
        SendMessage hwnd, WM_PAINT, 0&, 0&
    End If
    
End If

'Case ColumHeader
'------------------------------
If hwnd = m_Hwnd_CH Then
 
    If uMsg = WM_LBUTTONDOWN Then
        MouseDown = True
    End If
    
    If uMsg = WM_LBUTTONUP Then
        MouseDown = False
        IsDrag = False
    End If
    
    If uMsg = 132 Then
        If MouseDown Then
            IsDrag = True
        End If
    End If
    
    If uMsg = WM_MOUSEMOVE Then
    
        If TimerEnabled = False Then
            Call SetTimer(hwnd, ObjPtr(Me) + 1, 10, scb_SetCallbackAddr(4, 2))
            TimerEnabled = True
        End If
        
        mAL.All = lParam
        LSet mLH = mAL
        hti.pt.X = mLH.Lo
        hti.pt.Y = mLH.Hi
        
        SendMessage hwnd, HDM_HITTEST, 0&, hti
        
        If hti.iItem <> Indice Then
            Indice = hti.iItem
            SendMessage GetParent(hwnd), WM_PAINT, 0&, 0&
        End If
    End If
    
    If uMsg = WM_PAINT Then
        DrawColumHeader hwnd, Indice
    End If

End If
    
End Sub

