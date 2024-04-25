Attribute VB_Name = "Mod_SOS"
Option Explicit

Sub CargarArchivosSos(ByVal UserIndex As Integer)

    Dim ArchSOS As String
    Dim NumSos  As Long
 
    NumSos = "0"
 
    ArchSOS = Dir(App.Path & "\logs\show\sos\*.ini")
 
    Do While ArchSOS > ""
        Call CargaSOS(ArchSOS, UserIndex)
        ArchSOS = Dir
        NumSos = NumSos + "1"
    Loop
    
    Dim ArchDEN As String
    Dim NumDEN  As Long
 
    NumDEN = "0"
 
    ArchDEN = Dir(App.Path & "\logs\show\denuncia\*.ini")
 
    Do While ArchDEN > ""
        Call CargaDEN(ArchDEN, UserIndex)
        ArchDEN = Dir
        NumDEN = NumDEN + "1"
    Loop

    Dim ArchBUG As String
    Dim NumBUG  As Long
 
    NumBUG = "0"
 
    ArchBUG = Dir(App.Path & "\logs\show\BUG\*.ini")
 
    Do While ArchBUG > ""
        Call CargaBUG(ArchBUG, UserIndex)
        ArchBUG = Dir
        NumBUG = NumBUG + "1"
    Loop
       
    Dim ArchSUG As String
    Dim NumSUG  As Long
 
    NumSUG = "0"
 
    ArchSUG = Dir(App.Path & "\logs\show\sugerencia\*.ini")
 
    Do While ArchSUG > ""
        Call CargaSUG(ArchSUG, UserIndex)
        ArchSUG = Dir
        NumSUG = NumSUG + "1"
    Loop

End Sub

Sub CargaSOS(ByVal ArchSOS As String, ByVal UserIndex As Integer)

    Dim Count As Long
    Dim X     As Long
    Dim Msj   As String
    Dim FH    As String
    Dim mm    As String

    ArchSOS = Left$(ArchSOS, Len(ArchSOS) - 4)
 
    Count = val(GetVar(App.Path & "\Logs\Show\SOS\" & ArchSOS & ".ini", "Config", "NumMsg"))
 
    If Count = "0" Then
        Kill (App.Path & "\Logs\Show\SOS\" & ArchSOS & ".ini")

    End If

    For X = 1 To Count

        Msj = GetVar(App.Path & "\Logs\Show\SOS\" & ArchSOS & ".ini", "Mensaje" & X, "Mensaje")
        FH = GetVar(App.Path & "\Logs\Show\SOS\" & ArchSOS & ".ini", "Mensaje" & X, "HoraFecha")
    
        mm = ArchSOS & "@" & "SOS@" & Msj & "@(" & FH & ")"

        Call SendData(SendTarget.toIndex, UserIndex, 0, "NSOS" & mm)
    Next X

End Sub

Sub CargaDEN(ByVal ArchDENU As String, ByVal UserIndex As Integer)

    Dim Count As Long
    Dim X     As Long
    Dim Msj   As String
    Dim FH    As String
    Dim mm    As String
 
    ArchDENU = Left$(ArchDENU, Len(ArchDENU) - 4)
 
    Count = val(GetVar(App.Path & "\Logs\Show\DENUNCIA\" & ArchDENU & ".ini", "Config", "NumMsg"))
 
    If Count = "0" Then
        Kill (App.Path & "\Logs\Show\DENUNCIA\" & ArchDENU & ".ini")

    End If
 
    For X = 1 To Count
    
        Msj = GetVar(App.Path & "\Logs\Show\DENUNCIA\" & ArchDENU & ".ini", "Mensaje" & X, "Mensaje")
        FH = GetVar(App.Path & "\Logs\Show\DENUNCIA\" & ArchDENU & ".ini", "Mensaje" & X, "HoraFecha")
    
        mm = ArchDENU & "@" & "DENUNCIA@" & Msj & "@(" & FH & ")"
    
        Call SendData(SendTarget.toIndex, UserIndex, 0, "NSOS" & mm)
    Next X

End Sub

Sub CargaBUG(ByVal ArchBUG As String, ByVal UserIndex As Integer)
   
    Dim Count As Long
    Dim X     As Long
    Dim Msj   As String
    Dim FH    As String
    Dim mm    As String
 
    ArchBUG = Left$(ArchBUG, Len(ArchBUG) - 4)
 
    Count = val(GetVar(App.Path & "\Logs\Show\BUG\" & ArchBUG & ".ini", "Config", "NumMsg"))
 
    If Count = "0" Then
        Kill (App.Path & "\Logs\Show\BUG\" & ArchBUG & ".ini")

    End If
 
    For X = 1 To Count
    
        Msj = GetVar(App.Path & "\Logs\Show\BUG\" & ArchBUG & ".ini", "Mensaje" & X, "Mensaje")
        FH = GetVar(App.Path & "\Logs\Show\BUG\" & ArchBUG & ".ini", "Mensaje" & X, "HoraFecha")
    
        mm = ArchBUG & "@" & "BUG@" & Msj & "@(" & FH & ")"
    
        Call SendData(SendTarget.toIndex, UserIndex, 0, "NSOS" & mm)
    Next X

End Sub

Sub CargaSUG(ByVal ArchSUG As String, ByVal UserIndex As Integer)
   
    Dim Count As Long
    Dim X     As Long
    Dim Msj   As String
    Dim FH    As String
    Dim mm    As String
 
    ArchSUG = Left$(ArchSUG, Len(ArchSUG) - 4)
 
    Count = val(GetVar(App.Path & "\Logs\Show\SUGERENCIA\" & ArchSUG & ".ini", "Config", "NumMsg"))
 
    If Count = "0" Then
        Kill (App.Path & "\Logs\Show\SUGERENCIA\" & ArchSUG & ".ini")

    End If
 
    For X = 1 To Count
    
        Msj = GetVar(App.Path & "\Logs\Show\SUGERENCIA\" & ArchSUG & ".ini", "Mensaje" & X, "Mensaje")
        FH = GetVar(App.Path & "\Logs\Show\SUGERENCIA\" & ArchSUG & ".ini", "Mensaje" & X, "HoraFecha")
    
        mm = ArchSUG & "@" & "SUG@" & Msj & "@(" & FH & ")"
    
        Call SendData(SendTarget.toIndex, UserIndex, 0, "NSOS" & mm)
    Next X

End Sub

Sub DropSOS(ByVal rData As String, ByVal UserIndex As Integer)

    Dim SOS() As String
      
    SOS() = Split(rData, Chr(64))
      
    Dim Name    As String
    Dim Section As String
    Dim Msg     As String
    Dim FH      As String
      
    Name = SOS(0)
    Section = SOS(1)
    Msg = SOS(2)
    FH = SOS(3)
      
    Dim X        As Long
    Dim NumAdmin As Long
       
    If UCase$(Section) = "SOS" Then
          
        Call BorrarSOS(rData, UserIndex)
             
    End If
      
    If UCase$(Section) = "DENUNCIA" Then
                  
        NumAdmin = val(GetVar(App.Path & "\Dat\Ini\Config.ini", "SOS", "NumAdmin"))
           
        For X = 1 To NumAdmin
                 
            Dim NickAdmin As String
            Dim RevNick   As String
                 
            RevNick = UserList(UserIndex).Name
            NickAdmin = GetVar(App.Path & "\Dat\ini\Config.ini", "SOS", "Admin" & X)
                 
            If (RevNick = NickAdmin) Then
                Call BorrarDenuncia(rData, UserIndex)
                Exit Sub
              
            End If
        
        Next X
           
        Call SendData(SendTarget.toIndex, UserIndex, 0, _
            "||Actualiza lista! Tu mensaje no fue borrado de DENUNCIA porque no tienes permiso para borrar estos mensajes." & FONTTYPE_INFO)

    End If
      
    If UCase$(Section) = "BUG" Then
          
        Call BorrarBUG(rData, UserIndex)
          
    End If
      
    If UCase$(Section) = "SUG" Then
           
        NumAdmin = val(GetVar(App.Path & "\Dat\Ini\Config.ini", "SOS", "NumAdmin"))
           
        For X = 1 To NumAdmin
                 
            RevNick = UserList(UserIndex).Name
            NickAdmin = GetVar(App.Path & "\Dat\ini\Config.ini", "SOS", "Admin" & X)
                 
            If (RevNick = NickAdmin) Then
                     
                Call BorrarSUG(rData, UserIndex)
                     
                Exit Sub
              
            End If
        
        Next X
           
        Call SendData(SendTarget.toIndex, UserIndex, 0, _
            "||Actualiza lista! Tu mensaje no fue borrado de SUGERENCIA porque no tienes permiso para borrar estos mensajes." & FONTTYPE_INFO)

    End If

End Sub

'Borrar seccion
Sub DeleteSection(sSection As String, sIniFile As String)

    'this call will remove the entire section
    'corresponding to sSection in the file.
    'This is accomplished by passing
    'vbNullString as both the sKeyName and
    'sValue parameters. For example, assuming
    'that an ini file had:
    ' [Colours]
    '  Colour1=Red
    '  Colour2=Blue
    '  Colour3=Green
    '
    'and this sub was called passing "Colours"
    'as sSection, the entire Colours
    'section and all keys and values in
    'the section would be deleted.
   
    Call writeprivateprofilestring(sSection, vbNullString, vbNullString, sIniFile)

End Sub

'Changesection
Sub ChangeSection(ByVal filename As String, ByVal edit1 As Long, ByVal edit2 As Long)
    Dim iniText As String
   
    'change section name
    Open filename For Input As #1
    iniText = Input(LOF(1), #1)
    Close #1
   
    iniText = Replace(iniText, "Mensaje" & edit2, "Mensaje" & edit1)
    
    Open filename For Output As #1
    Print #1, iniText
    Close #1
    
End Sub

Sub BorrarSOS(ByVal rData As String, ByVal UserIndex As Integer)
    
    Dim SOS() As String
      
    SOS() = Split(rData, Chr(64))
      
    Dim Name    As String
    Dim Section As String
    Dim Msg     As String
    Dim FH      As String
      
    Name = SOS(0)
    Section = SOS(1)
    Msg = SOS(2)
    FH = SOS(3)
      
    Dim X       As Long
    Dim NumMsg  As Long
    Dim Borrado As Long
    Dim VerMsg  As String
    Dim Count   As String
      
    NumMsg = val(GetVar(App.Path & "\logs\show\SOS\" & Name & ".ini", "Config", "NumMsg"))
      
    For X = 1 To NumMsg
         
        If Borrado = 1 Then
         
            Call ChangeSection(App.Path & "\logs\show\SOS\" & Name & ".ini", Count, X)
             
            Count = X
         
        ElseIf Borrado = 0 Then
            
            VerMsg = GetVar(App.Path & "\logs\show\SOS\" & Name & ".ini", "Mensaje" & X, "Mensaje")
             
            If Msg = VerMsg Then
               
                Call DeleteSection("Mensaje" & X, App.Path & "\logs\show\SOS\" & Name & ".ini")
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has borrado el SOS de: " & Name & FONTTYPE_INFO)
              
                Count = X
                Borrado = 1

            End If
            
        End If
         
    Next X
      
    Dim Resta As Long
      
    Resta = NumMsg - "1"
      
    Call WriteVar(App.Path & "\logs\show\SOS\" & Name & ".ini", "Config", "NumMsg", Resta)
      
End Sub

Sub BorrarDenuncia(ByVal rData As String, ByVal UserIndex As Integer)
    
    Dim SOS() As String
      
    SOS() = Split(rData, Chr(64))
      
    Dim Name    As String
    Dim Section As String
    Dim Msg     As String
    Dim FH      As String
      
    Name = SOS(0)
    Section = SOS(1)
    Msg = SOS(2)
    FH = SOS(3)
      
    Dim X       As Long
    Dim NumMsg  As Long
    Dim Borrado As Long
    Dim VerMsg  As String
    Dim Count   As String
      
    NumMsg = val(GetVar(App.Path & "\logs\show\DENUNCIA\" & Name & ".ini", "Config", "NumMsg"))
      
    For X = 1 To NumMsg
         
        If Borrado = 1 Then
         
            Call ChangeSection(App.Path & "\logs\show\DENUNCIA\" & Name & ".ini", Count, X)
             
            Count = X
         
        ElseIf Borrado = 0 Then
            
            VerMsg = GetVar(App.Path & "\logs\show\DENUNCIA\" & Name & ".ini", "Mensaje" & X, "Mensaje")
             
            If Msg = VerMsg Then
               
                Call DeleteSection("Mensaje" & X, App.Path & "\logs\show\DENUNCIA\" & Name & ".ini")
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has borrado la DENUNCIA de: " & Name & FONTTYPE_INFO)
              
                Count = X
                Borrado = 1

            End If
            
        End If
         
    Next X
      
    Dim Resta As Long
      
    Resta = NumMsg - "1"
      
    Call WriteVar(App.Path & "\logs\show\DENUNCIA\" & Name & ".ini", "Config", "NumMsg", Resta)
      
End Sub

Sub BorrarBUG(ByVal rData As String, ByVal UserIndex As Integer)
    
    Dim SOS() As String
      
    SOS() = Split(rData, Chr(64))
      
    Dim Name    As String
    Dim Section As String
    Dim Msg     As String
    Dim FH      As String
      
    Name = SOS(0)
    Section = SOS(1)
    Msg = SOS(2)
    FH = SOS(3)
      
    Dim X       As Long
    Dim NumMsg  As Long
    Dim Borrado As Long
    Dim VerMsg  As String
    Dim Count   As String
      
    NumMsg = val(GetVar(App.Path & "\logs\show\BUG\" & Name & ".ini", "Config", "NumMsg"))
      
    For X = 1 To NumMsg
         
        If Borrado = 1 Then
         
            Call ChangeSection(App.Path & "\logs\show\BUG\" & Name & ".ini", Count, X)
             
            Count = X
         
        ElseIf Borrado = 0 Then
            
            VerMsg = GetVar(App.Path & "\logs\show\BUG\" & Name & ".ini", "Mensaje" & X, "Mensaje")
             
            If Msg = VerMsg Then
               
                Call DeleteSection("Mensaje" & X, App.Path & "\logs\show\BUG\" & Name & ".ini")
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has borrado el BUG de: " & Name & FONTTYPE_INFO)
              
                Count = X
                Borrado = 1

            End If
            
        End If
         
    Next X
      
    Dim Resta As Long
      
    Resta = NumMsg - "1"
      
    Call WriteVar(App.Path & "\logs\show\BUG\" & Name & ".ini", "Config", "NumMsg", Resta)
      
End Sub

Sub BorrarSUG(ByVal rData As String, ByVal UserIndex As Integer)
    
    Dim SOS() As String
      
    SOS() = Split(rData, Chr(64))
      
    Dim Name    As String
    Dim Section As String
    Dim Msg     As String
    Dim FH      As String
      
    Name = SOS(0)
    Section = SOS(1)
    Msg = SOS(2)
    FH = SOS(3)
      
    Dim X       As Long
    Dim NumMsg  As Long
    Dim Borrado As Long
    Dim VerMsg  As String
    Dim Count   As String
      
    NumMsg = val(GetVar(App.Path & "\logs\show\SUGERENCIA\" & Name & ".ini", "Config", "NumMsg"))
      
    For X = 1 To NumMsg
         
        If Borrado = 1 Then
         
            Call ChangeSection(App.Path & "\logs\show\SUGERENCIA\" & Name & ".ini", Count, X)
             
            Count = X
         
        ElseIf Borrado = 0 Then
            
            VerMsg = GetVar(App.Path & "\logs\show\SUGERENCIA\" & Name & ".ini", "Mensaje" & X, "Mensaje")
             
            If Msg = VerMsg Then
               
                Call DeleteSection("Mensaje" & X, App.Path & "\logs\show\SUGERENCIA\" & Name & ".ini")
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has borrado la SUGERENCIA de: " & Name & FONTTYPE_INFO)
              
                Count = X
                Borrado = 1

            End If
            
        End If
         
    Next X
      
    Dim Resta As Long
      
    Resta = NumMsg - "1"
      
    Call WriteVar(App.Path & "\logs\show\SUGERENCIA\" & Name & ".ini", "Config", "NumMsg", Resta)
      
End Sub

'Panel GM

Sub CargarArchivosGM(ByVal UserIndex As Integer)

    Dim ArchSOS As String
    Dim NumSos  As Long
 
    NumSos = "0"
 
    ArchSOS = Dir(App.Path & "\logs\consultas\*.ini")
 
    Do While ArchSOS > ""
        Call CargaGM(ArchSOS, UserIndex)
        ArchSOS = Dir
        NumSos = NumSos + "1"
    Loop

End Sub

Sub CargaGM(ByVal ArchSOS As String, ByVal UserIndex As Integer)

    Dim Count As Long
    Dim X     As Long
    Dim Msj   As String
    Dim FH    As String
    Dim mm    As String

    ArchSOS = Left$(ArchSOS, Len(ArchSOS) - 4)
 
    Count = val(GetVar(App.Path & "\Logs\Consultas\" & ArchSOS & ".ini", "Config", "NumMsg"))
 
    If Count = "0" Then
        Kill (App.Path & "\Logs\Consultas\" & ArchSOS & ".ini")

    End If

    For X = 1 To Count

        Msj = GetVar(App.Path & "\Logs\Consultas\" & ArchSOS & ".ini", "Mensaje" & X, "Mensaje")
        FH = GetVar(App.Path & "\Logs\Consultas\" & ArchSOS & ".ini", "Mensaje" & X, "HoraFecha")
    
        mm = ArchSOS & "@" & "SOS@" & Msj & "@(" & FH & ")"

        Call SendData(SendTarget.toIndex, UserIndex, 0, "PSGM" & mm)
    Next X

End Sub

Sub BorrarGM(ByVal rData As String, ByVal UserIndex As Integer)
    
    Dim SOS() As String
      
    SOS() = Split(rData, Chr(64))
      
    Dim Name    As String
    Dim Section As String
    Dim Msg     As String
    Dim FH      As String
      
    Name = SOS(0)
    Section = SOS(1)
    Msg = SOS(2)
    FH = SOS(3)
      
    Dim X       As Long
    Dim NumMsg  As Long
    Dim Borrado As Long
    Dim VerMsg  As String
    Dim Count   As String
      
    NumMsg = val(GetVar(App.Path & "\logs\Consultas\" & Name & ".ini", "Config", "NumMsg"))
      
    For X = 1 To NumMsg
         
        If Borrado = 1 Then
         
            Call ChangeSection(App.Path & "\logs\Consultas\" & Name & ".ini", Count, X)
             
            Count = X
         
        ElseIf Borrado = 0 Then
            
            VerMsg = GetVar(App.Path & "\logs\Consultas\" & Name & ".ini", "Mensaje" & X, "Mensaje")
             
            If Msg = VerMsg Then
               
                Call DeleteSection("Mensaje" & X, App.Path & "\logs\Consultas\" & Name & ".ini")
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has borrado el SOS de: " & Name & FONTTYPE_INFO)
              
                Count = X
                Borrado = 1

            End If
            
        End If
         
    Next X
      
    Dim Resta As Long
      
    Resta = NumMsg - "1"
      
    Call WriteVar(App.Path & "\logs\Consultas\" & Name & ".ini", "Config", "NumMsg", Resta)
      
End Sub
