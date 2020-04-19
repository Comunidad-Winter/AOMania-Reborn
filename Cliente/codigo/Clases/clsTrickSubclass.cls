VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrickSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Класс clsTrickSubclass.cls - для сабклассинга в VB6
' © Кривоус Анатолий Анатольевич (The trick), 2014
' Версия 1.1

Private Type PROCESS_HEAP_ENTRY

    lpData              As Long
    cbData              As Long
    cbOverhead          As Byte
    iRegionIndex        As Byte
    wFlags              As Integer
    dwCommittedSize     As Long
    dwUnCommittedSize   As Long
    lpFirstBlock        As Long
    lpLastBlock         As Long

End Type

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, _
    ByVal HWnd As Long, _
    ByVal msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long

Private Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long

Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

Private Declare Function HeapWalk Lib "kernel32" (ByVal hHeap As Long, ByRef lpEntry As PROCESS_HEAP_ENTRY) As Long

Private Declare Function HeapLock Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function HeapUnlock Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpValue As Long) As Long

Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableW" (ByVal lpName As Long, _
    ByVal lpBuffer As Long, _
    ByVal nSize As Long) As Long

Private Declare Function GetMem4 Lib "msvbvm60" (pSrc As Any, pDst As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal HWnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal HWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WM_CREATE                  As Long = &H1
Private Const WM_DESTROY                 As Long = &H2
Private Const GCL_WNDPROC                As Long = (-24)
Private Const GWL_WNDPROC                As Long = (-4)
Private Const HEAP_CREATE_ENABLE_EXECUTE As Long = &H40000
Private Const HEAP_NO_SERIALIZE          As Long = &H1
Private Const HEAP_ZERO_MEMORY           As Long = &H8
Private Const PROCESS_HEAP_ENTRY_BUSY    As Long = &H4
Private Const WNDPROCINDEX               As Long = 12
Private Const EnvName                    As String = "TrickSubclass"

' Событие, возникающее при вызове процедуры окна. Ret - возвращаемое значение, DefCall - вызвать предыдущую процедуру
Public Event wndProc(ByVal HWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ret As Long, DefCall As Boolean)

Private mIsSubclassed As Boolean
Private mhWnd         As Long
Private mIsClass      As Boolean
Private mIsPaused     As Boolean

Dim lpPrev            As Long
Dim hHeap             As Long
Dim lpAsm             As Long
Dim ClsCount          As Long

' Возвращает дескриптор окна
Public Property Get HWnd() As Long
    HWnd = mhWnd

End Property

' Возвращает адрес предыдущей процедуры
Public Property Get Previous() As Long

    If mIsSubclassed Then Previous = lpPrev

End Property

' Возвращает адрес текущей процедуры
Public Property Get Current() As Long

    If mIsSubclassed Then Current = lpAsm

End Property

' Если был сабклассинг, то True
Public Property Get IsSubclassed() As Boolean
    IsSubclassed = mIsSubclassed

End Property

' Приостановить сабклассинг
Public Function PauseSubclass() As Boolean

    If mIsSubclassed And Not mIsPaused Then
        mIsPaused = True
        PauseSubclass = True
    End If


End Function

' Возобновить сабклассинг
Public Function ResumeSubclass() As Boolean

    If mIsSubclassed And mIsPaused Then
        mIsPaused = False
        ResumeSubclass = True
    End If


End Function

' Если сабклассинг класса, то  True
Public Property Get IsClass() As Boolean
    IsClass = mIsClass

End Property

' Если на паузе то True
Public Property Get IsPaused() As Boolean
    IsPaused = mIsPaused

End Property

' Установить хук на класс (если хук стоит, то он будет снят)
' Действует на все последующие экземпляры окон
Public Function HookClass(ByVal HWnd As Long) As Boolean

    If mIsSubclassed Then
        If Not UnHook Then Exit Function

    End If

    lpPrev = GetClassLong(HWnd, GCL_WNDPROC)
    mIsClass = True

    If CreateAsm Then
        If SetClassLong(HWnd, GCL_WNDPROC, lpAsm) Then
            mhWnd = HWnd
            mIsSubclassed = True
            HookClass = True
            mIsPaused = True
        Else:
        mIsClass = False

        End If

    Else:
    mIsClass = False

    End If

End Function

' Установить хук на окно (если хук стоит, то он будет снят)
Public Function Hook(ByVal HWnd As Long) As Boolean

    If mIsSubclassed Then
        If Not UnHook Then Exit Function

    End If

    lpPrev = GetWindowLong(HWnd, GWL_WNDPROC)
    mIsClass = False

    If CreateAsm Then
        If SetWindowLong(HWnd, GWL_WNDPROC, lpAsm) Then
            mhWnd = HWnd
            mIsSubclassed = True
            Hook = True

        End If

    End If

End Function

' Снять хук
Public Function UnHook() As Boolean

    If Not mIsSubclassed Then Exit Function
    If mIsClass Then
        UnHook = SetClassLong(mhWnd, GCL_WNDPROC, lpPrev)
    Else:
    UnHook = SetWindowLong(mhWnd, GWL_WNDPROC, lpPrev)

    End If

    If UnHook Then
        mhWnd = 0
        mIsSubclassed = False
    End If


End Function

' Вызвать процедуру по умолчанию для окна
Public Function CallDef(ByVal HWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef status As Boolean) As Long

    If Not mIsSubclassed Then Exit Function
    CallDef = CallWindowProc(lpPrev, HWnd, msg, wParam, lParam)
    status = True

End Function

' --------------------------------------------------------------------------------------------------------------------------------------
Private Function mWndProc(ByVal HWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If mIsPaused Then
        mWndProc = CallWindowProc(lpPrev, HWnd, msg, wParam, lParam)
    Else
        Dim DefCall As Boolean
        RaiseEvent wndProc(HWnd, msg, wParam, lParam, mWndProc, DefCall)

        If DefCall Then mWndProc = CallWindowProc(lpPrev, HWnd, msg, wParam, lParam)

    End If

    If mIsClass And mWndProc = 0 Then

        Select Case msg

            Case WM_CREATE
                ClsCount = ClsCount + 1

            Case WM_DESTROY
                ClsCount = ClsCount - 1

                If ClsCount = 0 Then
                    mhWnd = HWnd
                    UnHook
                End If


        End Select

    End If

End Function

Private Sub Class_Terminate()

    If hHeap = 0 Then Exit Sub
    UnHook

    If CountHooks = 1 Then
        HeapDestroy hHeap
        hHeap = 0
        SaveCurHeap
    Else
        HeapFree hHeap, HEAP_NO_SERIALIZE, ByVal lpAsm

    End If

End Sub

Private Function CreateAsm() As Boolean
    Dim inIDE   As Boolean
    Dim AsmSize As Long
    Dim Ptr     As Long
    Dim isFirst As Boolean

    'Debug.Assert MakeTrue(inIDE)
    
    If lpAsm = 0 Then
        If inIDE Then AsmSize = &H51 Else AsmSize = &H1D
        hHeap = GetPrevHeap()
        
        If hHeap Then
            If inIDE Then
                Dim flag As Long
                Ptr = GetFlagPointer()
                GetMem4 ByVal Ptr, flag

                If flag Then
                    HeapDestroy hHeap
                    isFirst = True

                End If

            End If

        Else:
        isFirst = True

        End If
        
        If isFirst Then
            hHeap = HeapCreate(HEAP_CREATE_ENABLE_EXECUTE Or HEAP_NO_SERIALIZE, 0, 0)

            If hHeap = 0 Then
                Err.Raise 7
                Exit Function
            End If

            If Not SaveCurHeap() Then
                HeapDestroy hHeap
                hHeap = 0
                Err.Raise 7
                Exit Function
            End If

            AsmSize = AsmSize + &H4

        End If
        
        lpAsm = HeapAlloc(hHeap, HEAP_NO_SERIALIZE Or HEAP_ZERO_MEMORY, AsmSize)
        
        If lpAsm = 0 Then
            If isFirst Then HeapDestroy hHeap
            hHeap = 0
            Err.Raise 7
            Exit Function

        End If
        
        Dim prv As Long
        Dim i   As Long
        
        If inIDE Then
            If isFirst Then
                GetMem4 0&, ByVal lpAsm
                lpAsm = lpAsm + 4

            End If

        End If
        
    End If

    Ptr = lpAsm

    If inIDE Then
        CreateIDEStub (Ptr)
        Ptr = Ptr + &H34

    End If

    CreateStackConv Ptr
    CreateAsm = True
    
End Function

Private Function GetFlagPointer() As Long
    Dim he As PROCESS_HEAP_ENTRY
    HeapLock hHeap

    Do While HeapWalk(hHeap, he)

        If he.wFlags And PROCESS_HEAP_ENTRY_BUSY Then
            GetFlagPointer = he.lpData
            Exit Function
        End If

    Loop
    HeapUnlock hHeap

End Function

Private Function CountHooks() As Long
    Dim he As PROCESS_HEAP_ENTRY
    HeapLock hHeap

    Do While HeapWalk(hHeap, he)

        If he.wFlags And PROCESS_HEAP_ENTRY_BUSY Then CountHooks = CountHooks + 1
    Loop
    HeapUnlock hHeap

End Function

Private Function SaveCurHeap() As Boolean
    Dim i   As Long
    Dim out As String
    out = Hex(hHeap)

    For i = Len(out) + 1 To 8
        out = "0" & out
    Next
    SaveCurHeap = SetEnvironmentVariable(StrPtr(EnvName), StrPtr(out))

End Function

Private Function GetPrevHeap() As Long
    Dim out As String
    out = Space(&H8)

    If GetEnvironmentVariable(StrPtr(EnvName), StrPtr(out), LenB(out)) Then GetPrevHeap = Val("&H" & out)

End Function

Private Function CreateStackConv(ByVal Ptr As Long) As Boolean
    Dim lpMeth As Long
    Dim vTable As Long
    
    GetMem4 ByVal ObjPtr(Me), vTable
    GetMem4 ByVal vTable + WNDPROCINDEX * 4 + &H1C, lpMeth
    
    GetMem4 &H5450C031, ByVal Ptr + &H0
    GetMem4 &H488DE409, ByVal Ptr + &H4
    GetMem4 &H2474FF04, ByVal Ptr + &H8
    GetMem4 &H68FAE018, ByVal Ptr + &HC
    GetMem4 &H0, ByVal Ptr + &H10
    GetMem4 &HE8, ByVal Ptr + &H14
    GetMem4 &H10C25800, ByVal Ptr + &H18
    GetMem4 &H9000, ByVal Ptr + &H1C
    
    GetMem4 ObjPtr(Me), ByVal Ptr + &H10                    ' Push Me
    GetMem4 lpMeth - (Ptr + &H14) - 5, ByVal Ptr + &H14 + 1 ' Call WndProc
    
End Function
 
Private Function CreateIDEStub(ByVal Ptr As Long) As Boolean
    Dim hInstVB6    As Long
    Dim lpEbMode    As Long
    Dim hInstUser32 As Long
    Dim lpCallProc  As Long
    Dim lpSetLong   As Long
    Dim dwIndex     As Long
    
    hInstVB6 = GetModuleHandle(StrPtr("vba6"))

    If hInstVB6 = 0 Then Exit Function
    hInstUser32 = GetModuleHandle(StrPtr("user32"))

    If hInstUser32 = 0 Then Exit Function
    
    lpEbMode = GetProcAddress(hInstVB6, "EbMode")

    If lpEbMode = 0 Then Exit Function
    lpCallProc = GetProcAddress(hInstUser32, "CallWindowProcW")

    If lpCallProc = 0 Then Exit Function
    If mIsClass Then
        lpSetLong = GetProcAddress(hInstUser32, "SetClassLongA")
        dwIndex = GCL_WNDPROC
    Else
        lpSetLong = GetProcAddress(hInstUser32, "SetWindowLongA")
        dwIndex = GWL_WNDPROC

    End If

    If lpSetLong = 0 Then Exit Function
    
    GetMem4 &HFFFFFBE8, ByVal Ptr + &H0
    GetMem4 &H74C084FF, ByVal Ptr + &H4
    GetMem4 &H74013C10, ByVal Ptr + &H8
    GetMem4 &H685827, ByVal Ptr + &HC
    GetMem4 &H50000000, ByVal Ptr + &H10
    GetMem4 &HFFFFE7E9, ByVal Ptr + &H14
    GetMem4 &HDFFFF, ByVal Ptr + &H18
    GetMem4 &H68000000, ByVal Ptr + &H1C
    GetMem4 lpPrev, ByVal Ptr + &H20
    GetMem4 &H68, ByVal Ptr + &H24
    GetMem4 &HE474FF98, ByVal Ptr + &H28
    GetMem4 &HFFCEE80C, ByVal Ptr + &H2C
    GetMem4 &HD9EBFFFF, ByVal Ptr + &H30
    
    GetMem4 lpEbMode - Ptr - 5, ByVal Ptr + 1 + 0               ' Call EbMode
    GetMem4 lpPrev, ByVal Ptr + &HF                             ' Push PrevProc
    GetMem4 lpCallProc - (Ptr + &H14) - 5, ByVal Ptr + 1 + &H14 ' Jmp CallWindowProcA
    GetMem4 dwIndex, ByVal Ptr + &H25                           ' Push nIndex
    GetMem4 lpSetLong - (Ptr + &H2D) - 5, ByVal Ptr + 1 + &H2D  ' Call SetWindowLong/SetClassLong
    GetMem4 Ptr - 4, ByVal Ptr + &H1B                           ' dec dword ptr [Flag]
    
    CreateIDEStub = True

End Function

Private Function MakeTrue(value As Boolean) As Boolean
    value = True
    MakeTrue = True
End Function

