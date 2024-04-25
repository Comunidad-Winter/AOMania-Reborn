VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEstadisticasIPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegisterWindowMessage _
                Lib "user32" _
                Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function SendNotifyMessage _
                Lib "user32" _
                Alias "SendNotifyMessageA" (ByVal hWnd As Long, _
                                            ByVal msg As Long, _
                                            ByVal wParam As Long, _
                                            ByVal lParam As Long) As Long
Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long
Private Declare Function SendMessageLong _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long

Private Const WM_USER = &H400

Private sMensaje As Long, hVentana As Long, hVentanaMia As Long

Private Declare Function GetWindowText _
                Lib "user32" _
                Alias "GetWindowTextA" (ByVal hWnd As Long, _
                                        ByVal lpString As String, _
                                        ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength _
                Lib "user32" _
                Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal wCmd As Long) As Long

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2

'*************************************************
Public Enum EstaNotificaciones

    CANTIDAD_ONLINE = 1
    RECORD_USUARIOS = 2
    UPTIME_SERVER = 3
    CANTIDAD_MAPAS = 4
    EVENTO_NUEVO_CLAN = 5
    
    HANDLE_WND_SERVER = 100

End Enum

'*************************************************

Private Function BuscaVentana(Wnd As Long, str As String) As Long

    Dim w As Long
    Dim L As Long
    Dim t As String

    w = GetWindow(Wnd, GW_HWNDFIRST)

    While w <> 0

        L = GetWindowTextLength(w)

        If L > 0 Then
            t = Space(L + 1)
            L = GetWindowText(w, t, L + 1)

            If Left(t, Len(str)) = str Then
                BuscaVentana = w
                Exit Function

            End If

        End If

        w = GetWindow(w, GW_HWNDNEXT)
    Wend

    BuscaVentana = 0

End Function

Public Function Informar(ByVal QueCosa As EstaNotificaciones, _
                         ByVal Parametro As Long) As Long

    Call BuscaWndEstadisticas

    If hVentana <> 0 Then
        Informar = SendMessageLong(hVentana, sMensaje, QueCosa, Parametro)

    End If

End Function

Public Function EstadisticasAndando() As Boolean

    Dim ret As Long

    Call BuscaWndEstadisticas
    'Ret = SendNotifyMessage(hVentana, sMensaje, 0, 0)
    EstadisticasAndando = (hVentana <> 0)

End Function

Public Sub Inicializa(ByVal hWnd As Long)

    hVentanaMia = hWnd
    sMensaje = RegisterWindowMessage("EstadisticasAO")

End Sub

Private Sub BuscaWndEstadisticas()

    hVentana = BuscaVentana(hVentanaMia, "Servidor de estadisticas AO")

End Sub
