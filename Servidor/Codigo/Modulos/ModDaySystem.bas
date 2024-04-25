Attribute VB_Name = "ModDaySystem"
Option Explicit

Public CountTC As Byte
Public NameDay As String
Public StatusTimeChange As Boolean
Public NocheLicantropo As Boolean

Sub TimeChange()

    CountTC = CountTC + 1

    If CountTC = 25 Then CountTC = 1

    Call SendData(SendTarget.ToAll, 0, 0, "HUCT" & CountTC)

    Call DayChange(CountTC)
    
    Call DayNameChange(CountTC)
    
    Call CambiaTimeHora(CountTC)
    
    If StatusTimeChange = True Then
        Call SystemTimeChange(CountTC)
    End If
    
    If StatusTimeChange = False Then StatusTimeChange = True
    
End Sub

Sub DayNameChange(ByVal Hora As Byte)

        If Hora >= 1 And Hora <= 7 Then NameDay = "Noche"
        If Hora >= 8 And Hora <= 12 Then NameDay = "Amanecer"
        If Hora >= 13 And Hora <= 18 Then NameDay = "Día"
        If Hora >= 19 And Hora <= 21 Then NameDay = "Tarde"
        If Hora >= 22 And Hora <= 24 Then NameDay = "Noche"
       
End Sub

Sub SystemTimeChange(ByVal Hora As Byte)
      
      If Hora = 8 Then
          Call CambiaClima
      End If
      
      If Hora = 13 Then
          Call CambiaClima
      End If
      
      If Hora = 19 Then
           Call CambiaClima
      End If
       
       If Hora = 22 Then
            Call CambiaClima
       End If
       
End Sub

Sub AdminCambiaDia(ByVal Hora As Byte)

      Call DayNameChange(Hora)
      Call LoadClima(Hora)
      Call SendMainAmbientAll
      Call CambiaTimeHora(Hora)
      Call DayChange(Hora)
      
End Sub

Sub CambiaTimeHora(ByVal Hora As Byte)

    Dim n As Integer
    
    CountTC = Hora
    
    'Noche
    If Hora >= 1 And Hora <= 7 Then
        Call SendData(SendTarget.ToAll, 0, 0, "TW53")
    End If

    'Amanecer
    If Hora >= 8 And Hora <= 12 Then
        Call SendData(SendTarget.ToAll, 0, 0, "TW55")
    End If

    'Día
    If Hora >= 13 And Hora <= 18 Then
        Call SendData(SendTarget.ToAll, 0, 0, "TW55")
    End If

    'Tarde
    If Hora >= 19 And Hora <= 21 Then
        Call SendData(SendTarget.ToAll, 0, 0, "TW55")
    End If

    'Noche
    If Hora >= 22 And Hora <= 24 Then
        Call SendData(SendTarget.ToAll, 0, 0, "TW53")
    End If
    
    Call SendData(SendTarget.ToAll, 0, 0, "HUCT" & Hora)

End Sub


Sub DayChange(ByVal Hora As Byte)

    Dim n As Integer
    Dim UserIndex As Integer

    'Noche
    If Hora >= 1 And Hora <= 7 Then

        If NocheLicantropo = False Then
            NocheLicantropo = True
            For n = 1 To LastUser

                UserIndex = n

                If UserList(UserIndex).flags.Privilegios = PlayerType.User And _
                   UCase$(UserList(UserIndex).Raza) = "LICANTROPO" Then
                    Call DarPoderLicantropo(UserIndex)
                End If

            Next n
        End If

    End If

    'Amanecer
    If Hora >= 8 And Hora <= 12 Then
        If NocheLicantropo = True Then
            NocheLicantropo = False
            For n = 1 To LastUser

                UserIndex = n

                If UserList(UserIndex).flags.Privilegios = PlayerType.User And _
                   UCase$(UserList(UserIndex).Raza) = "LICANTROPO" Then
                    Call QuitarPoderLicantropo(UserIndex)
                End If

            Next n
        End If
    End If

    'Día
    If Hora >= 13 And Hora <= 18 Then
        If NocheLicantropo = True Then
            NocheLicantropo = False
            For n = 1 To LastUser

                UserIndex = n

                If UserList(UserIndex).flags.Privilegios = PlayerType.User And _
                   UCase$(UserList(UserIndex).Raza) = "LICANTROPO" Then
                    Call QuitarPoderLicantropo(UserIndex)
                End If

            Next n
        End If
    End If

    'Tarde
    If Hora >= 19 And Hora <= 21 Then
        If NocheLicantropo = True Then
            NocheLicantropo = False
            For n = 1 To LastUser

                UserIndex = n

                If UserList(UserIndex).flags.Privilegios = PlayerType.User And _
                   UCase$(UserList(UserIndex).Raza) = "LICANTROPO" Then
                    Call QuitarPoderLicantropo(UserIndex)
                End If

            Next n
        End If
    End If

    'Noche
    If Hora >= 22 And Hora <= 24 Then
        If NocheLicantropo = False Then
            NocheLicantropo = True
            For n = 1 To LastUser

                UserIndex = n

                If UserList(UserIndex).flags.Privilegios = PlayerType.User And _
                   UCase$(UserList(UserIndex).Raza) = "LICANTROPO" Then
                    Call DarPoderLicantropo(UserIndex)
                End If

            Next n
        End If
    End If

End Sub

Sub DarPoderLicantropo(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).flags.Licantropo = 0 Then
        UserList(UserIndex).flags.Licantropo = "1"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + "3"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + "3"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) + "3"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) + "3"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) = UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) + "3"
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Es de noche, tu cuerpo empieza a transformarse y te sientes más poderoso." & FONTTYPE_INFO)
        Call DoLicantropo(UserIndex)
   End If
   
End Sub

Sub QuitarPoderLicantropo(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).flags.Licantropo = 1 Then
        UserList(UserIndex).flags.Licantropo = "0"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) - "3"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) - "3"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) - "3"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) - "3"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) = UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) - "3"
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya es de día, vuelves a tu apariencia normal" & FONTTYPE_INFO)
        Call DoLicantropo(UserIndex)
    End If
    
End Sub

Sub DesconectaPoderLicantropo(ByVal UserIndex As Integer)

  If UserList(UserIndex).flags.Licantropo = 1 Then
        UserList(UserIndex).flags.Licantropo = "0"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) - "3"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) - "3"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) - "3"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) - "3"
        UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) = UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) - "3"
        Call DoLicantropo(UserIndex)
    End If
      
End Sub
