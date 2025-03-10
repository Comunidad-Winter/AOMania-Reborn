VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundMapInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type p_tSoundMapInfo

    Cantidad As Integer
    SoundIndex() As Integer
    flags() As Long
    Probabilidad() As Single

End Type

Private Enum p_eSoundFlags

    ninguna = 0
    Lluvia = 1

End Enum

Private p_Mapas() As p_tSoundMapInfo

'sonidos conocidos, pasados a enum para intelisense
Public Enum e_SoundIndex

    MUERTE_HOMBRE = 11
    MUERTE_MUJER = 74
    FLECHA_IMPACTO = 65
    CONVERSION_BARCO = 55
    MORFAR_MANZANA = 82
    SOUND_COMIDA = 7

End Enum

Private Sub Class_Initialize()

    'armar el array
    ReDim p_Mapas(1 To NumMaps) As p_tSoundMapInfo
    Call LoadSoundMapInfo

End Sub

Public Sub LoadSoundMapInfo()

    Dim i       As Integer
    Dim j       As Integer
    Dim Temps   As String
    Dim MAPFILE As String
    MAPFILE = App.Path & MapPath & "MAPA"
    
    For i = 1 To UBound(p_Mapas)
        Temps = GetVar(MAPFILE & i & ".dat", "SONIDOS", "Cantidad")

        If IsNumeric(Temps) Then
            p_Mapas(i).Cantidad = CInt(GetVar(MAPFILE & i & ".dat", "SONIDOS", "Cantidad"))
            ReDim p_Mapas(i).flags(1 To p_Mapas(i).Cantidad) As Long
            ReDim p_Mapas(i).Probabilidad(1 To p_Mapas(i).Cantidad) As Single
            ReDim p_Mapas(i).SoundIndex(1 To p_Mapas(i).Cantidad) As Integer

            For j = 1 To p_Mapas(i).Cantidad
                p_Mapas(i).flags(j) = CInt(GetVar(MAPFILE & i & ".dat", "SONIDO" & j, "Flags"))
                p_Mapas(i).Probabilidad(j) = CSng(GetVar(MAPFILE & i & ".dat", "SONIDO" & j, "Probabilidad"))
                p_Mapas(i).SoundIndex(j) = CInt(GetVar(MAPFILE & i & ".dat", "SONIDO" & j, "Sonido"))
            Next j

        Else
            p_Mapas(i).Cantidad = 0

        End If

    Next i

End Sub

Public Sub ReproducirSonidosDeMapas()

    Dim i           As Integer
    Dim SonidosMapa As Byte
    Dim SonidoMapa  As Byte

    For i = 1 To UBound(p_Mapas)

        If p_Mapas(i).Cantidad > 0 Then
            SonidoMapa = RandomNumber(1, p_Mapas(i).Cantidad)

            If RandomNumber(1, 100) <= p_Mapas(i).Probabilidad(SonidoMapa) Then

                'tocarlo
                If SecondaryWeather Then
                    If p_Mapas(i).flags(SonidoMapa) Xor p_eSoundFlags.Lluvia Then
                        Call SendData(SendTarget.tomap, 0, i, "TW" & p_Mapas(i).SoundIndex(SonidoMapa))

                    End If

                Else

                    If p_Mapas(i).flags(SonidoMapa) Xor p_eSoundFlags.ninguna Then
                        Call SendData(SendTarget.tomap, 0, i, "TW" & p_Mapas(i).SoundIndex(SonidoMapa))

                    End If

                End If

            End If

        End If

    Next i
    
End Sub

Public Sub ReproducirSonido(ByVal Destino As Byte, ByVal UserIndex As Integer, ByVal Mapa As Integer, ByVal SoundIndex As Integer)
                            
    Call SendData(Destino, UserIndex, Mapa, "TW" & SoundIndex)

End Sub
