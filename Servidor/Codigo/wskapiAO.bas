Attribute VB_Name = "wskapiAO"
Option Explicit

''
' Modulo para manejar Winsock
'

'Si la variable esta en TRUE , al iniciar el WsApi se crea
'una ventana LABEL para recibir los mensajes. Al detenerlo,
'se destruye.
'Si es FALSE, los mensajes se envian al form frmMain (o el
'que sea).
#Const WSAPI_CREAR_LABEL = True

Private Const SD_RECEIVE As Long = &H0
Private Const SD_SEND As Long = &H1
Private Const SD_BOTH As Long = &H2

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                             ByVal hWnd As Long, _
                                                                             ByVal Msg As Long, _
                                                                             ByVal wParam As Long, _
                                                                             ByVal lParam As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
                                                                              ByVal lpClassName As String, _
                                                                              ByVal lpWindowName As String, _
                                                                              ByVal dwStyle As Long, _
                                                                              ByVal X As Long, _
                                                                              ByVal Y As Long, _
                                                                              ByVal nWidth As Long, _
                                                                              ByVal nHeight As Long, _
                                                                              ByVal hwndParent As Long, _
                                                                              ByVal hMenu As Long, _
                                                                              ByVal hInstance As Long, _
                                                                              lpParam As Any) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const WS_CHILD = &H40000000
Public Const GWL_WNDPROC = (-4)

Private Const SIZE_RCVBUF As Long = 8192
Private Const SIZE_SNDBUF As Long = 8192

''
'Esto es para agilizar la busqueda del slot a partir de un socket dado,
'sino, la funcion BuscaSlotSock se nos come todo el uso del CPU.
'
' @param Sock sock
' @param slot slot
'
Public Type tSockCache

    Sock As Long
    Slot As Long

End Type

Public WSAPISock2Usr As New Collection

' ====================================================================================
' ====================================================================================

Public OldWProc As Long
Public ActualWProc As Long
Public hWndMsg As Long

' ====================================================================================
' ====================================================================================

Public SockListen As Long
Public LastSockListen As Long

' ====================================================================================
' ====================================================================================

Public Sub IniciaWsApi(ByVal hwndParent As Long)

    Call LogApiSock("IniciaWsApi")
    Debug.Print "IniciaWsApi"

    #If WSAPI_CREAR_LABEL Then
        hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", WS_CHILD, 0, 0, 0, 0, hwndParent, 0, App.hInstance, ByVal 0&)
    #Else
        hWndMsg = hwndParent
    #End If    'WSAPI_CREAR_LABEL

    OldWProc = SetWindowLong(hWndMsg, GWL_WNDPROC, AddressOf WndProc)
    ActualWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)

    Dim Desc As String
    Call StartWinsock(Desc)

End Sub

Public Sub LimpiaWsApi(ByVal hWnd As Long)

    Call LogApiSock("LimpiaWsApi")

    If WSAStartedUp Then
        Call EndWinsock

    End If

    If OldWProc <> 0 Then
        SetWindowLong hWndMsg, GWL_WNDPROC, OldWProc
        OldWProc = 0

    End If

    #If WSAPI_CREAR_LABEL Then

        If hWndMsg <> 0 Then
            DestroyWindow hWndMsg

        End If

    #End If

End Sub

Public Function BuscaSlotSock(ByVal s As Long, Optional ByVal CacheInd As Boolean = False) As Long

    On Error GoTo hayerror

    If WSAPISock2Usr.Count <> 0 Then    ' GSZAO
        BuscaSlotSock = WSAPISock2Usr.Item(CStr(s))
    Else
        BuscaSlotSock = -1

    End If

    Exit Function

hayerror:
    BuscaSlotSock = -1

End Function

Public Sub AgregaSlotSock(ByVal Sock As Long, ByVal Slot As Long)

    Debug.Print "AgregaSockSlot"

    'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("AgregaSlotSock:: sock=" & Sock & " slot=" & Slot)

    If WSAPISock2Usr.Count > MaxUsers Then
        'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("Imposible agregarSlotSock (wsapi2usr.count>maxusers)")
        Call CloseSocket(Slot)
        Exit Sub

    End If

    WSAPISock2Usr.Add CStr(Slot), CStr(Sock)

    'Dim Pri As Long, Ult As Long, Med As Long
    'Dim LoopC As Long
    '
    'If WSAPISockChacheCant > 0 Then
    '    Pri = 1
    '    Ult = WSAPISockChacheCant
    '    Med = Int((Pri + Ult) / 2)
    '
    '    Do While (Pri <= Ult) And (Ult > 1)
    '        If Sock < WSAPISockChache(Med).Sock Then
    '            Ult = Med - 1
    '        Else
    '            Pri = Med + 1
    '        End If
    '        Med = Int((Pri + Ult) / 2)
    '    Loop
    '
    '    Pri = IIf(Sock < WSAPISockChache(Med).Sock, Med, Med + 1)
    '    Ult = WSAPISockChacheCant
    '    For LoopC = Ult To Pri Step -1
    '        WSAPISockChache(LoopC + 1) = WSAPISockChache(LoopC)
    '    Next LoopC
    '    Med = Pri
    'Else
    '    Med = 1
    'End If
    'WSAPISockChache(Med).Slot = Slot
    'WSAPISockChache(Med).Sock = Sock
    'WSAPISockChacheCant = WSAPISockChacheCant + 1

End Sub

Public Sub BorraSlotSock(ByVal Sock As Long, Optional ByVal CacheIndice As Long)

    Dim Cant As Long

    Cant = WSAPISock2Usr.Count

    On Error Resume Next

    WSAPISock2Usr.Remove CStr(Sock)

    Debug.Print "BorraSockSlot " & Cant & " -> " & WSAPISock2Usr.Count

End Sub

Public Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error Resume Next

    Dim Ret As Long
    Dim Tmp As String

    Dim s As Long, e As Long
    Dim n As Integer

    Dim Dale As Boolean
    Dim UltError As Long

    WndProc = 0

    If CamaraLenta = 1 Then
        Sleep 1

    End If

    Select Case Msg

    Case 1025

        s = wParam
        e = WSAGetSelectEvent(lParam)
        'Debug.Print "Msg: " & msg & " W: " & wParam & " L: " & lParam
        Call LogApiSock("Msg: " & Msg & " W: " & wParam & " L: " & lParam)

        Select Case e

        Case FD_ACCEPT

            'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("FD_ACCEPT")
            If s = SockListen Then
                'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("sockLIsten = " & s & ". Llamo a Eventosocketaccept")
                Call EventoSockAccept(s)

            End If

            '    Case FD_WRITE
            '        N = BuscaSlotSock(s)
            '        If N < 0 And s <> SockListen Then
            '            'Call apiclosesocket(s)
            '            call WSApiCloseSocket(s)
            '            Exit Function
            '        End If
            '
            '        UserList(N).SockPuedoEnviar = True

            '        Call IntentarEnviarDatosEncolados(N)
            '
            ''        Dale = UserList(N).ColaSalida.Count > 0
            ''        Do While Dale
            ''            Ret = WsApiEnviar(N, UserList(N).ColaSalida.Item(1), False)
            ''            If Ret <> 0 Then
            ''                If Ret = WSAEWOULDBLOCK Then
            ''                    Dale = False
            ''                Else
            ''                    'y aca que hacemo' ?? help! i need somebody, help!
            ''                    Dale = False
            ''                    Debug.Print "ERROR AL ENVIAR EL DATO DESDE LA COLA " & Ret & ": " & GetWSAErrorString(Ret)
            ''                End If
            ''            Else
            ''            '    Debug.Print "Dato de la cola enviado"
            ''                UserList(N).ColaSalida.Remove 1
            ''                Dale = (UserList(N).ColaSalida.Count > 0)
            ''            End If
            ''        Loop

        Case FD_READ

            n = BuscaSlotSock(s)

            If n < 0 And s <> SockListen Then
                'Call apiclosesocket(s)
                Call WSApiCloseSocket(s)
                Exit Function

            End If

            'Call WSAAsyncSelect(s, hWndMsg, ByVal 1025, ByVal (0))

            '4k de buffer
            'buffer externo
            Tmp = Space$(SIZE_RCVBUF)   'si cambias este valor, tambien hacelo mas abajo
            'donde dice ret = 8192 :)

            Ret = recv(s, Tmp, Len(Tmp), 0)

            ' Comparo por = 0 ya que esto es cuando se cierra
            ' "gracefully". (mas abajo)
            If Ret < 0 Then
                UltError = Err.LastDllError

                If UltError = WSAEMSGSIZE Then
                    Debug.Print "WSAEMSGSIZE"
                    Ret = SIZE_RCVBUF
                Else
                    Debug.Print "Error en Recv: " & GetWSAErrorString(UltError)
                    Call LogApiSock("Error en Recv: N=" & n & " S=" & s & " Str=" & GetWSAErrorString(UltError))

                    'no hay q llamar a CloseSocket() directamente,
                    'ya q pueden abusar de algun error para
                    'desconectarse sin los 10segs. CREEME.
                    '    Call C l o s e Socket(N)

                    If UserList(n).flags.Privilegios = User Then
                        Call CloseSocketSL(n)
                        Call Cerrar_Usuario(n)

                    End If

                    Exit Function

                End If

            ElseIf Ret = 0 Then

                If UserList(n).flags.Privilegios = User Then
                    Call CloseSocketSL(n)
                    Call Cerrar_Usuario(n)

                End If

            End If

            'Call WSAAsyncSelect(s, hWndMsg, ByVal 1025, ByVal (FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT))

            Tmp = Left(Tmp, Ret)

            'Call LogApiSock("WndProc:FD_READ:N=" & N & ":TMP=" & Tmp)

            Call EventoSockRead(n, Tmp)

        Case FD_CLOSE
            n = BuscaSlotSock(s)

            If s <> SockListen Then Call apiclosesocket(s)

            Call LogApiSock("WndProc:FD_CLOSE:N=" & n & ":Err=" & WSAGetAsyncError(lParam))

            If n > 0 Then
                Call BorraSlotSock(UserList(n).ConnID)
                UserList(n).ConnID = -1
                UserList(n).ConnIDValida = False
                Call EventoSockClose(n)

            End If

        End Select

    Case Else
        WndProc = CallWindowProc(OldWProc, hWnd, Msg, wParam, lParam)

    End Select

End Function

'Retorna 0 cuando se envi� o se metio en la cola,
'retorna <> 0 cuando no se pudo enviar o no se pudo meter en la cola
Public Function WsApiEnviar(ByVal Slot As Integer, ByVal str As String) As Long

'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("WsApiEnviar:: slot=" & Slot & " str=" & str & " len(str)=" & Len(str) & " encolar=" & Encolar)

    Dim Ret As String
    Dim UltError As Long
    Dim Retorno As Long

    Retorno = 0

    'Debug.Print ">>>> " & str

    If UserList(Slot).ConnID <> -1 And UserList(Slot).ConnIDValida Then

        Ret = send(ByVal UserList(Slot).ConnID, ByVal str, ByVal Len(str), ByVal 0)

        If Ret < 0 Then
            UltError = Err.LastDllError

            'If UltError = WSAEWOULDBLOCK Then

            'End If

            Retorno = UltError

        End If

    ElseIf UserList(Slot).ConnID <> -1 And Not UserList(Slot).ConnIDValida Then

        If Not UserList(Slot).Counters.Saliendo Then
            Retorno = -1

        End If

    End If

    WsApiEnviar = Retorno

End Function

Public Sub LogCustom(ByVal str As String)

    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\custom.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & "(" & Timer & ") " & str
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogApiSock(ByVal str As String)

    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\wsapi.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub EventoSockAccept(ByVal SockID As Long)

'==========================================================
'USO DE LA API DE WINSOCK
'========================

    Dim NewIndex As Integer
    Dim Ret As Long
    Dim Tam As Long, sa As sockaddr
    Dim NuevoSock As Long
    Dim i As Long
    Dim tStr As String

    Tam = sockaddr_size

    '=============================================
    'SockID es en este caso es el socket de escucha,
    'a diferencia de socketwrench que es el nuevo
    'socket de la nueva conn

    'Modificado por Maraxus
    'Ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 0)
    Ret = accept(SockID, sa, Tam)

    If Ret = INVALID_SOCKET Then
        i = Err.LastDllError
        Call LogCriticEvent("Error en Accept() API " & i & ": " & GetWSAErrorString(i))
        Exit Sub

    End If

    If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
        Call WSApiCloseSocket(NuevoSock)
        Exit Sub

    End If

    'If Ret = INVALID_SOCKET Then
    '    If Err.LastDllError = 11002 Then
    '        ' We couldn't decide if to accept or reject the connection
    '        'Force reject so we can get it out of the queue
    '        LogCustom ("Pre WSAAccept CallbackData=1")
    '        Ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 1)
    '        LogCustom ("WSAccept Callbackdata 1, devuelve " & Ret)
    '        Call LogCriticEvent("Error en WSAAccept() API 11002: No se pudo decidir si aceptar o rechazar la conexi�n.")
    '    Else
    '        i = Err.LastDllError
    '        LogCustom ("Error en WSAAccept() API " & i & ": " & GetWSAErrorString(i))
    '        Call LogCriticEvent("Error en WSAAccept() API " & i & ": " & GetWSAErrorString(i))
    '        Exit Sub
    '    End If
    'End If

    NuevoSock = Ret

    'Seteamos el tama�o del buffer de entrada a 512 bytes
    If setsockopt(NuevoSock, SOL_SOCKET, SO_RCVBUFFER, SIZE_RCVBUF, 4) <> 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear el tama�o del buffer de entrada " & i & ": " & GetWSAErrorString(i))

    End If

    'Seteamos el tama�o del buffer de salida a 1 Kb
    If setsockopt(NuevoSock, SOL_SOCKET, SO_SNDBUFFER, SIZE_SNDBUF, 4) <> 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear el tama�o del buffer de salida " & i & ": " & GetWSAErrorString(i))

    End If

    If False Then
        'If SecurityIp.IPSecuritySuperaLimiteConexiones(sa.sin_addr) Then
        tStr = "ERRLimite de conexiones para su IP alcanzado." & ENDC
        Call send(ByVal NuevoSock, ByVal tStr, ByVal Len(tStr), ByVal 0)
        Call WSApiCloseSocket(NuevoSock)
        Exit Sub

    End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   BIENVENIDO AL SERVIDOR!!!!!!!!
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Mariano: Baje la busqueda de slot abajo de CondicionSocket y limite x ip
    NewIndex = NextOpenUser    ' Nuevo indice

    If NewIndex <= MaxUsers Then

        UserList(NewIndex).ip = GetAscIP(sa.sin_addr)

        'Busca si esta banneada la ip
        For i = 1 To BanIps.Count

            If BanIps.Item(i) = UserList(NewIndex).ip Then
                'Call apiclosesocket(NuevoSock)
                tStr = "ERRSu IP se encuentra bloqueada en este servidor." & ENDC
                Call send(ByVal NuevoSock, ByVal tStr, ByVal Len(tStr), ByVal 0)
                'Call SecurityIp.IpRestarConexion(sa.sin_addr)
                Call WSApiCloseSocket(NuevoSock)
                Exit Sub

            End If

        Next i

        ' anti bombardeo de ip
        Dim k As Integer, X As Integer
        X = 1

        For k = 1 To LastUser

            If (UserList(k).ip = UserList(NewIndex).ip) Then X = X + 1
        Next k

        If (X > 10) Then
            Call WSApiCloseSocket(NuevoSock)
            UserList(NewIndex).ConnID = -1
            Exit Sub

        End If

        ' terminamos el bombardeo de ip

        If NewIndex > LastUser Then LastUser = NewIndex

        'UserList(NewIndex).SockPuedoEnviar = True
        UserList(NewIndex).ConnID = NuevoSock
        UserList(NewIndex).ConnIDValida = True
        'Set UserList(NewIndex).CommandsBuffer = New CColaArray
        'Set UserList(NewIndex).ColaSalida = New Collection

        Call AgregaSlotSock(NuevoSock, NewIndex)
    Else
        tStr = "ERRServer lleno." & ENDC
        Dim AAA As Long
        AAA = send(ByVal NuevoSock, ByVal tStr, ByVal Len(tStr), ByVal 0)
        'Call SecurityIp.IpRestarConexion(sa.sin_addr)
        Call WSApiCloseSocket(NuevoSock)

    End If

End Sub

Public Sub EventoSockRead(ByVal Slot As Integer, ByRef Datos As String)

    Dim T() As String
    Dim LoopC As Long

    UserList(Slot).RDBuffer = UserList(Slot).RDBuffer & Datos

    T = Split(UserList(Slot).RDBuffer, ENDC)

    If UBound(T) > 0 Then
        UserList(Slot).RDBuffer = T(UBound(T))

        For LoopC = 0 To UBound(T) - 1

            '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            '%%% SI ESTA OPCION SE ACTIVA SOLUCIONA %%%
            '%%% EL PROBLEMA DEL SPEEDHACK          %%%
            '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

            If UserList(Slot).ConnID <> -1 Then
                Call HandleData(Slot, T(LoopC))
            Else
                Exit Sub

            End If

        Next LoopC

    End If

End Sub

Public Sub EventoSockClose(ByVal Slot As Integer)

'Es el mismo user al que est� revisando el centinela??
'Si estamos ac� es porque se cerr� la conexi�n, no es un /salir, y no queremos banearlo....
    If Centinela.RevisandoUserIndex = Slot Then Call modCentinela.CentinelaUserLogout

    If UserList(Slot).flags.UserLogged Then
        Call CloseSocketSL(Slot)
        Call Cerrar_Usuario(Slot)
    Else
        Call CloseSocket(Slot)

    End If

End Sub

Public Sub WSApiReiniciarSockets()

    Dim i As Long

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)

    'Cierra todas las conexiones
    For i = 1 To MaxUsers

        If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
            Call CloseSocket(i)

        End If

        'Call ResetUserSlot(i)
    Next i

    ' No 'ta el PRESERVE :p
    ReDim UserList(1 To MaxUsers)

    For i = 1 To MaxUsers
        UserList(i).ConnID = -1
        UserList(i).ConnIDValida = False
    Next i

    LastUser = 1
    NumUsers = 0

    Call Sleep(100)

    Call LimpiaWsApi(frmMain.hWnd)
    Call Sleep(100)
    Call IniciaWsApi(frmMain.hWnd)
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

End Sub

Public Sub WSApiCloseSocket(ByVal Socket As Long)

    Call WSAAsyncSelect(Socket, hWndMsg, ByVal 1025, ByVal (FD_CLOSE))
    Call ShutDown(Socket, SD_BOTH)

End Sub

Public Function CondicionSocket(ByRef lpCallerId As WSABUF, _
                                ByRef lpCallerData As WSABUF, _
                                ByRef lpSQOS As FLOWSPEC, _
                                ByVal Reserved As Long, _
                                ByRef lpCalleeId As WSABUF, _
                                ByRef lpCalleeData As WSABUF, _
                                ByRef Group As Long, _
                                ByVal dwCallbackData As Long) As Long

    Dim sa As sockaddr

    'Check if we were requested to force reject

    If dwCallbackData = 1 Then
        CondicionSocket = CF_REJECT
        Exit Function

    End If

    'Get the address

    CopyMemory sa, ByVal lpCallerId.lpBuffer, lpCallerId.dwBufferLen

    If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
        CondicionSocket = CF_REJECT
        Exit Function

    End If

    CondicionSocket = CF_ACCEPT    'En realdiad es al pedo, porque CondicionSocket se inicializa a 0, pero as� es m�s claro....

End Function
