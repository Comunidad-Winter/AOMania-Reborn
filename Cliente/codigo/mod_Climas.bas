Attribute VB_Name = "mod_Climas"
'******************************************************************************
'Modulo Climas
'******************************************************************************
Option Explicit

'******************************************************************************
'Declaraciones del Tiempo
'******************************************************************************
Private Const MIN_MAINSTATUS      As Byte = 1
Private Const MAX_MAINSTATUS      As Byte = 4

Private Const MIN_SECONDARYSTATUS As Byte = 0
Private Const MAX_SECONDARYSTATUS As Byte = 1

' X% De que se de un cambio en el ambiente.
Private Const SECONDARYCHANCE     As Byte = 8

' Cada X min hacemos el chequeo para ver si llueve con el RND de arriba.
Private Const SECONDARYCHECK      As Byte = 30

Private Enum MainStatus

    Mañana = 1
    Dia = 2
    Tarde = 3
    Noche = 4

End Enum

Private Enum SecondaryStatus

    Nada = 0
    Lluvia = 1
    
End Enum

Private Type MeteoMain

    Current As MainStatus
    NextCurrent As MainStatus

End Type

Private Type MeteoSecondary

    Current As SecondaryStatus
    'NextCurrent As SecondaryStatus

End Type

Private Type Meteo

    MainStatus As MeteoMain
    SecondaryStatus As MeteoSecondary

End Type

'******************************************************************************
'Constantes de Tiempos
'******************************************************************************
Private Enum Life ' En Minutos
    
    Time_Mañana = 1     ' 100
    Time_Dia = 2        ' 100
    Time_Tarde = 3      ' 250
    Time_Noche = 4      ' 380
    Time_Lluvia = 5     ' 5

End Enum

Private Enum Particle

    Nada = 0
    Lluvia = 1

End Enum

Private ParticleAmbient() As Byte
Private LifeTime()        As Integer

'******************************************************************************
'Declaraciones del tiempo usado
'******************************************************************************
Private Weather           As Meteo
Private WeatherTime       As Long
Private WeatherSecondTime As Long
Private LastRnd           As Byte
Public SecondaryWeather   As Boolean

'******************************************************************************
'Funciones Del Tiempo:
'******************************************************************************

Public Sub InitTimeLife()

    ReDim ParticleAmbient(MIN_SECONDARYSTATUS To MAX_SECONDARYSTATUS) As Byte
    ReDim LifeTime(MIN_MAINSTATUS To MAX_MAINSTATUS + MAX_SECONDARYSTATUS) As Integer
    
    ' Minutos
    LifeTime(Life.Time_Mañana) = 100
    LifeTime(Life.Time_Dia) = 100
    LifeTime(Life.Time_Tarde) = 250
    LifeTime(Life.Time_Noche) = 380
    LifeTime(Life.Time_Lluvia) = 5
 
    ' Particle.ini
    ParticleAmbient(Particle.Nada) = 0
    ParticleAmbient(Particle.Lluvia) = 8
    
    WeatherTime = 0
    WeatherSecondTime = 0
    LastRnd = 0
    SecondaryWeather = False
    
    Weather.MainStatus.Current = MainStatus.Mañana
    Weather.MainStatus.NextCurrent = MainStatus.Dia
    
    Weather.SecondaryStatus.Current = SecondaryStatus.Nada
    
End Sub

Public Sub SpendTime()

    WeatherTime = WeatherTime + 1

    If WeatherTime >= LifeTime(Weather.MainStatus.Current) Then
       
        Weather.MainStatus.Current = Weather.MainStatus.NextCurrent
    
        Dim NextStatus As Byte
        NextStatus = Weather.MainStatus.Current + 1

        If NextStatus > MAX_MAINSTATUS Then NextStatus = MIN_MAINSTATUS

        Weather.MainStatus.NextCurrent = NextStatus
        Call SendData(SendTarget.toall, 0, 0, "CLM" & Weather.MainStatus.Current)
        
        WeatherTime = 0

    End If
    
    If Weather.SecondaryStatus.Current = SecondaryStatus.Nada Then
        
        ' con esto hacemos que haga este rnd tan seguido para evitar el spam de ambiente CON MUCHA SUERTE.
        LastRnd = LastRnd + 1

        If LastRnd > SECONDARYCHECK Then
            If RandomNumber(1, 100) <= SECONDARYCHANCE Then
        
                Dim NewSecondaryStatus As Byte
                NewSecondaryStatus = CByte(RandomNumber(MAX_SECONDARYSTATUS, MAX_SECONDARYSTATUS))

                If NewSecondaryStatus > SecondaryStatus.Nada Then
                    Weather.SecondaryStatus.Current = NewSecondaryStatus
                    SecondaryWeather = True
                    Call SendData(SendTarget.toall, 0, 0, "CLA" & ParticleAmbient(NewSecondaryStatus))
                
                End If

            End If

            LastRnd = 0

        End If

    Else
    
        WeatherSecondTime = WeatherSecondTime + 1

        If WeatherSecondTime >= LifeTime(Weather.SecondaryStatus.Current + MAX_MAINSTATUS) Then
            Call SendData(SendTarget.toall, 0, 0, "CLO")
            Weather.SecondaryStatus.Current = SecondaryStatus.Nada
            WeatherSecondTime = 0
            SecondaryWeather = False

        End If

    End If

End Sub

Public Sub SecondaryAmbient()

    SecondaryWeather = Not SecondaryWeather
    
    If SecondaryWeather Then
    
        If Weather.SecondaryStatus.Current = SecondaryStatus.Nada Then
        
            Dim NewSecondaryStatus As Byte
            NewSecondaryStatus = CByte(RandomNumber(MAX_SECONDARYSTATUS, MAX_SECONDARYSTATUS))

            Weather.SecondaryStatus.Current = NewSecondaryStatus
            WeatherSecondTime = 0
            
            SecondaryWeather = True
            Call SendData(SendTarget.toall, 0, 0, "CLA" & ParticleAmbient(NewSecondaryStatus))

        End If

    Else
       
        Weather.SecondaryStatus.Current = SecondaryStatus.Nada
        WeatherSecondTime = 0
        SecondaryWeather = False
        Call SendData(SendTarget.toall, 0, 0, "CLO")

    End If

End Sub

Public Sub SendMainAmbient(ByVal UserIndex As Integer)
 
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "CLM" & Weather.MainStatus.Current)

End Sub

Public Sub SendSecondaryAmbient(ByVal UserIndex As Integer)

    If Weather.SecondaryStatus.Current = SecondaryStatus.Nada Then Exit Sub

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "CLA" & ParticleAmbient(Weather.SecondaryStatus.Current))

End Sub

