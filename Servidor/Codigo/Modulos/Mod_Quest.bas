Attribute VB_Name = "Mod_Quest"
Option Explicit

Public NumQuests As Long

Public Type tRecompensaObjeto
     ObjIndex As Integer
     Amount As Integer
End Type

Public Type tMataNpc
     NpcIndex As Integer
     Cantidad As Integer
End Type

Public Type tMataUser
    MinNivel As Byte
    MaxNivel As Byte
    NUMCLASES As Byte
    Clases(1 To NUMCLASES) As String
    NUMRAZAS As Byte
    Razas(1 To NUMRAZAS) As String
    Alineacion As Byte
    Faccion As Byte
    RangoFaccion As Byte
End Type

Public Type tBuscaObj
     ObjIndex As Integer
     Amount As Integer
End Type

Public Type tObjsNpc
      NpcIndex As Integer
      ObjIndex As Integer
      Amount As Integer
End Type

Public Type tDescubrePalabra
     NpcIndex As Integer
     Frase As String
End Type

Public Type tQuestList
    nombre As String
    Descripcion As String
    Rehacer As Byte
    MinNivel As Byte
    MaxNivel As Byte
    RecompensaOro As Long
    RecompensaExp As Long
    RecompensaItem As Byte
    RecompensaObjeto() As tRecompensaObjeto
    HablarNpc As Byte
    HablaNpc(1 To 10) As Integer
    NUMCLASES As Byte
    Clases(1 To NUMCLASES) As String
    NUMRAZAS As Byte
    Razas(1 To NUMRAZAS) As String
    Alineacion As Byte
    Faccion As Byte
    RangoFaccion As Byte
    NumNpc As Byte
    MataNpc(1 To 10) As tMataNpc
    NumUser As Integer
    MataUser As tMataUser
    NumObjs As Byte
    BuscaObj As tBuscaObj
    NumObjsNpc As Byte
    ObjsNpc As tObjsNpc
    NumNpcDD As Byte
    NpcDD(1 To 10) As Integer
    NumMapas As Integer
    Mapas(1 To 10) As Integer
    NumDescubre As Integer
    DescubrePalabra(1 To 10) As tDescubrePalabra
End Type

Public QuestList() As tQuestList

Public Sub Load_Quest()

    Dim Quest As Integer

    Dim LooPC As Integer

    Dim Datos As String
    
    Dim Leer  As New clsIniManager
    
    Call Leer.Initialize(DatPath & "Quest.dat")
    
    NumQuests = Leer.GetValue("INIT", "NumQuests")
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumQuests
    frmCargando.cargar.value = 0

    ReDim Preserve QuestList(1 To NumQuests) As tQuestList
    
    For Quest = 1 To NumQuests
       
        QuestList(Quest).nombre = Leer.GetValue("Quest" & Quest, "Nombre")
        QuestList(Quest).Descripcion = Leer.GetValue("Quest" & Quest, "Descripcion")
        QuestList(Quest).Rehacer = val(Leer.GetValue("Quest" & Quest, "Rehacer"))
        QuestList(Quest).MinNivel = val(Leer.GetValue("Quest" & Quest, "MinNivel"))
        QuestList(Quest).MaxNivel = val(Leer.GetValue("Quest" & Quest, "MaxNivel"))
        QuestList(Quest).RecompensaOro = val(Leer.GetValue("Quest" & Quest, "RecompensaOro"))
        QuestList(Quest).RecompensaExp = val(Leer.GetValue("Quest" & Quest, "RecompensaExp"))
        QuestList(Quest).RecompensaItem = val(Leer.GetValue("Quest" & Quest, "RecompensaItem"))
       
        If QuestList(Quest).RecompensaItem > 0 Then

            For LooPC = 1 To MAX_INVENTORY_SLOTS
             
                Datos = Leer.GetValue("Quest" & Quest, "RecompensaItem" & LooPC)
             
                QuestList(Quest).RecompensaObjeto(LooPC).ObjIndex = val(ReadField(1, Datos, 45))
                QuestList(Quest).RecompensaObjeto(LooPC).Amount = val(ReadField(2, Datos, 45))
             
            Next LooPC

        End If
        
        QuestList(Quest).HablarNpc = val(Leer.GetValue("Quest" & Quest, "HablarNPC"))
        
        If QuestList(Quest).HablarNpc > 0 Then
            
            For LooPC = 1 To MAX_INVENTORY_SLOTS
                   QuestList(Quest).HablaNpc(LooPC) = val(Leer.GetValue("Quest" & Quest, "HablarNPC" & LooPC))
            Next LooPC
            
        End If
        
        QuestList(Quest).NUMCLASES = val(Leer.GetValue("Quest" & Quest, "Clases"))
        
        If QuestList(Quest).NUMCLASES > 0 Then
            
            For LooPC = 1 To NUMCLASES
            
               QuestList(Quest).Clases(LooPC) = CStr(Leer.GetValue("Quest" & Quest, "Clases" & LooPC))
               
            Next LooPC
            
        End If
        
        QuestList(Quest).NUMRAZAS = val(Leer.GetValue("Quest" & Quest, "Razas"))
        
        If QuestList(Quest).NUMRAZAS > 0 Then
            
            For LooPC = 1 To NUMRAZAS
                 QuestList(Quest).Razas(LooPC) = CStr(Leer.GetValue("Quest" & Quest, "Razas" & LooPC))
            Next LooPC
            
        End If
        
        QuestList(Quest).Alineacion = val(Leer.GetValue("Quest" & Quest, "Alineacion"))
        QuestList(Quest).Faccion = val(Leer.GetValue("Quest" & Quest, "Faccion"))
        QuestList(Quest).RangoFaccion = val(Leer.GetValue("Quest" & Quest, "RangoFaccion"))
        QuestList(Quest).NumNpc = val(Leer.GetValue("Quest" & Quest, "MataNPC"))
        
        If QuestList(Quest).NumNpc > 0 Then
            For LooPC = 1 To MAX_INVENTORY_SLOTS
                
                Datos = Leer.GetValue("Quest" & Quest, "MataNPC" & LooPC)
                
                QuestList(Quest).MataNpc(LooPC).NpcIndex = val(ReadField(1, Datos, 45))
                QuestList(Quest).MataNpc(LooPC).Cantidad = val(ReadField(2, Datos, 45))
                
            Next LooPC
        End If
        
        QuestList(Quest).NumUser = val(Leer.GetValue("Quest" & Quest, "MataUSER"))
        
        If QuestList(Quest).NumUser > 0 Then
            QuestList(Quest).MataUser.MinNivel = val(Leer.GetValue("Quest" & Quest, "MUMinNivel"))
            QuestList(Quest).MataUser.MaxNivel = val(Leer.GetValue("Quest" & Quest, "MUMaxNivel"))
            QuestList(Quest).MataUser.NUMCLASES = val(Leer.GetValue("Quest" & Quest, "MUClases"))
            
            If QuestList(Quest).MataUser.NUMCLASES > 0 Then
                   
                   For LooPC = 1 To NUMCLASES
                         QuestList(Quest).MataUser.Clases(LooPC) = CStr(Leer.GetValue("Quest" & Quest, "MUClases" & LooPC))
                   Next LooPC
                   
            End If
            
            QuestList(Quest).MataUser.NUMRAZAS = val(Leer.GetValue("Quest" & Quest, "MURazas"))
            
            If QuestList(Quest).MataUser.NUMRAZAS > 0 Then
                  
                  For LooPC = 1 To NUMRAZAS
                       QuestList(Quest).MataUser.Razas(LooPC) = CStr(Leer.GetValue("Quest" & Quest, "MURazas" & LooPC))
                  Next LooPC
                  
            End If
            
            QuestList(Quest).MataUser.Alineacion = val(Leer.GetValue("Quest" & Quest, "MUAlineacion"))
            QuestList(Quest).MataUser.Faccion = val(Leer.GetValue("Quest" & Quest, "MUFaccion"))
            QuestList(Quest).MataUser.RangoFaccion = val(Leer.GetValue("Quest" & Quest, "MURangoFaccion"))
            
        End If
        
        QuestList(Quest).NumObjs = val(Leer.GetValue("Quest" & Quest, "BuscaObjetos"))
        
        If QuestList(Quest).NumObjs > 0 Then
             
             For LooPC = 1 To MAX_INVENTORY_SLOTS
                    
                    Datos = Leer.GetValue("Quest" & Quest, "BuscaObjetos" & LooPC)
                    
                    QuestList(Quest).BuscaObj.ObjIndex = val(ReadField(1, Datos, 45))
                    QuestList(Quest).BuscaObj.Amount = val(ReadField(2, Datos, 45))
                    
             Next LooPC
             
        End If
        
        QuestList(Quest).NumObjsNpc = val(Leer.GetValue("Quest" & Quest, "ObjetoNpc"))
        
        If QuestList(Quest).NumObjsNpc > 0 Then
            
            For LooPC = 1 To MAX_INVENTORY_SLOTS
                   
                   Datos = Leer.GetValue("Quest" & Quest, "ObjetoNpc" & LooPC)
                   
                   QuestList(Quest).ObjsNpc.NpcIndex = val(ReadField(1, Datos, 45))
                   QuestList(Quest).ObjsNpc.ObjIndex = val(ReadField(2, Datos, 45))
                   QuestList(Quest).ObjsNpc.Amount = val(ReadField(3, Datos, 45))
                   
            Next LooPC
            
        End If
        
        QuestList(Quest).NumNpcDD = val(Leer.GetValue("Quest" & Quest, "NpcDD"))
        
        If QuestList(Quest).NumNpcDD > 0 Then
            
            For LooPC = 1 To MAX_INVENTORY_SLOTS
                   QuestList(Quest).NpcDD(LooPC) = val(Leer.GetValue("Quest" & Quest, "NpcDD" & LooPC))
            Next LooPC
            
        End If
        
        QuestList(Quest).NumMapas = val(Leer.GetValue("Quest" & Quest, "EncontrarMapa"))
        
        If QuestList(Quest).NumMapas > 0 Then
            
            For LooPC = 1 To MAX_INVENTORY_SLOTS
                QuestList(Quest).Mapas(LooPC) = val(Leer.GetValue("Quest" & Quest, "EncontrarMapa" & LooPC))
            Next LooPC
            
        End If
        
        QuestList(Quest).NumDescubre = val(Leer.GetValue("Quest" & Quest, "DescubrePalabra"))
        
        If QuestList(Quest).NumDescubre > 0 Then
              
              For LooPC = 1 To MAX_INVENTORY_SLOTS
                     
                     Datos = Leer.GetValue("Quest" & Quest, "DescubrePalabra" & LooPC)
                     
                     QuestList(Quest).DescubrePalabra(LooPC).NpcIndex = val(ReadField(1, Datos, 45))
                     QuestList(Quest).DescubrePalabra(LooPC).Frase = val(ReadField(2, Datos, 45))
                     
              Next LooPC
              
        End If
        
        frmCargando.cargar.value = frmCargando.cargar.value + 1
       
    Next Quest
  
End Sub

Public Sub IniciarVentanaQuest(ByVal UserIndex As Integer)
     
    Dim LooPC As Integer

    Dim Datos As String
     
    With UserList(UserIndex)
     
        For LooPC = 1 To NumQuests
            Datos = Datos & .Quest.UserQuest(LooPC) & ", "
        Next LooPC
      
        Datos = Left$(Datos, Len(Datos) - 2)
     
        Call SendData(ToIndex, UserIndex, 0, "XU" & Datos)

    End With
        
End Sub

Public Sub ConnectQuest(ByVal UserIndex As Integer)
       
       With UserList(UserIndex)
       
            If .Quest.Start = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "0")
            ElseIf .Quest.Start = 1 Then
               Call SendData(ToIndex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "1")
            ElseIf .Quest.Start = 2 Then
               Call SendData(ToIndex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "2")
            End If
       
       End With
       
End Sub

Public Sub IniciarMisionQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
       
       Dim LooPC As Integer
       Dim N As Integer
       Dim c As Integer
       Dim Datos As String
        
        With UserList(UserIndex)
        
              If .Quest.Start > 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Ya tienes una misión iniciada!! Acabala antes de volver a empezar otra." & FONTTYPE_INFO)
                  Exit Sub
              End If
              
              If CompruebaIniciarQuest(UserIndex, Quest) = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Tienes otras misiones que realizar, antes que esta!! Revise la lista de misiones!!" & FONTTYPE_INFO)
                  Exit Sub
              ElseIf CompruebaIniciarQuest(UserIndex, Quest) = 2 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Debes completar todas las misiones, para poder repetir esta mision!!" & FONTTYPE_INFO)
                  Exit Sub
               ElseIf CompruebaIniciarQuest(UserIndex, Quest) = 3 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Esta mision no se puede rehacer, intente con otra!!" & FONTTYPE_INFO)
                  Exit Sub
              End If
              
              'AQUI DEBES PONER IDENTIFICADOR SI TIENE HECHO, REHAGA O NO.
              
              If QuestList(Quest).MinNivel > 0 Then
                  If .Stats.ELV < QuestList(Quest).MinNivel Then
                      Call SendData(ToIndex, UserIndex, 0, "||Para hacer esta quest, necesitas tener como minimo nivel " & QuestList(Quest).MinNivel & "." & FONTTYPE_INFO)
                      Exit Sub
                  End If
              End If
              
              If QuestList(Quest).MaxNivel > 0 Then
                  If .Stats.ELV > QuestList(Quest).MaxNivel Then
                       Call SendData(ToIndex, UserIndex, 0, "||Has alcansado el nivel maximo para realizar esta misión, puedes pasar a la siguiente!" & FONTTYPE_INFO)
                       .Quest.UserQuest(Quest) = 1
                       .Quest.Quest = Quest
                      Exit Sub
                  End If
              End If
              
              If QuestList(Quest).NUMCLASES > 0 Then
                  
                  N = QuestList(Quest).NUMCLASES
                  
                  For LooPC = 1 To N
                         Debug.Print UCase$(QuestList(Quest).Clases(LooPC))
                         If UCase$(QuestList(Quest).Clases(LooPC)) = UCase$(.Clase) Then
                             c = c + 1
                         End If
                         
                  Next LooPC
                  
                  If c = 0 Then
                     Call SendData(ToIndex, UserIndex, 0, "||Tu clase no esta permitida para realizar esta misión, puedes pasar a la siguiente!" & FONTTYPE_INFO)
                     .Quest.UserQuest(Quest) = 1
                     .Quest.Quest = Quest
                     Exit Sub
                  End If
                  
                  c = 0
              End If
              
              If QuestList(Quest).NUMRAZAS > 0 Then
                  
                  N = QuestList(Quest).NUMRAZAS
                  
                  For LooPC = 1 To N
                     If UCase$(QuestList(Quest).Razas(LooPC)) = UCase$(.Raza) Then
                         c = c + 1
                     End If
                  Next LooPC
                  
                  If c = 0 Then
                     Call SendData(ToIndex, UserIndex, 0, "||Tu raza no esta permitida para realizar esta misión, puedes pasar a la siguiente!" & FONTTYPE_INFO)
                     .Quest.UserQuest(Quest) = 1
                     .Quest.Quest = Quest
                     Exit Sub
                  End If
                  
                  c = 0
             End If
              
              If QuestList(Quest).Alineacion > 0 Then
                  
                  If QuestList(Quest).Alineacion = 1 Then
                      
                      If Criminal(UserIndex) Then
                         c = 0
                      Else
                         c = c + 1
                      End If
                      
                      If c = 0 Then
                          Call SendData(ToIndex, UserIndex, 0, "||No se permiten criminales en esta misión, puedes pasar a la siguiente!" & FONTTYPE_INFO)
                          .Quest.UserQuest(Quest) = 1
                          .Quest.Quest = Quest
                          Exit Sub
                      End If
                       c = 0
                  End If
                  
                  If QuestList(Quest).Alineacion = 2 Then
                      
                      If Criminal(UserIndex) Then
                         c = c + 1
                         Else
                         c = 0
                      End If
                      
                      If c = 0 Then
                          Call SendData(ToIndex, UserIndex, 0, "||No se permiten ciudadanos en esta misión, puedes pasar a la siguiente!" & FONTTYPE_INFO)
                          .Quest.UserQuest(Quest) = 1
                          .Quest.Quest = Quest
                          Exit Sub
                      End If
                      c = 0
                  End If
              
              End If
              
              If QuestList(Quest).Faccion > 0 Then
                  If QuestList(Quest).Faccion = 1 Then
                      If Not TieneFaccion(UserIndex) Then
                          Call SendData(ToIndex, UserIndex, 0, "||Solo se permiten usuarios con faccion en esta misión, puedes pasar a la siguiente!" & FONTTYPE_INFO)
                           .Quest.UserQuest(Quest) = 1
                           .Quest.Quest = Quest
                           Exit Sub
                      End If
                  End If
              End If
              
              If QuestList(Quest).RangoFaccion > 0 Then
                   If TieneFaccion(UserIndex) Then
                        If RangoFaccion(UserIndex) < QuestList(Quest).RangoFaccion Then
                            Call SendData(ToIndex, UserIndex, 0, "||Necesitas la " & QuestList(Quest).RangoFaccion & " rango de tu faccion para esta misión!" & FONTTYPE_INFO)
                          
                        End If
                   Else
                        Call SendData(ToIndex, UserIndex, 0, "||Solo se permiten usuarios con faccion en esta misión, puedes pasar a la siguiente!" & FONTTYPE_INFO)
                        .Quest.UserQuest(Quest) = 1
                        .Quest.Quest = Quest
                        Exit Sub
                   End If
              End If
              
              Call SendData(ToIndex, UserIndex, 0, "||Has iniciado nueva misión: " & QuestList(Quest).nombre & FONTTYPE_QUEST)
              .Quest.Start = 1
              .Quest.Quest = Quest
              
              Datos = "Objetivo: "
              
              If QuestList(Quest).NumNpc > 0 Then
                  
                  For LooPC = 1 To QuestList(Quest).NumNpc
                         Datos = Datos & "Mata " & QuestList(Quest).MataNpc(LooPC).Cantidad & " " & Npclist(BuscoNpcQuest(QuestList(Quest).MataNpc(LooPC).NpcIndex)).Name & "||"
                  Next LooPC
                  
                  .Quest.NumNpc = QuestList(Quest).NumNpc
                  
              End If
              
              Datos = Left$(Datos, Len(Datos) - 2)
              
              Call SendData(ToIndex, UserIndex, 0, "||" & Datos & FONTTYPE_GUILD)
              
              Call SendData(ToIndex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "1")
              
        End With
        
End Sub

Public Sub ActualizaQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
      
      Dim LooPC As Integer
       
       With UserList(UserIndex)
            
            If QuestList(Quest).NumNpc > 0 Then
                
                For LooPC = 1 To QuestList(Quest).NumNpc
                     
                     If QuestList(Quest).MataNpc(LooPC).Cantidad <> .Quest.MataNpc(LooPC) Then
                         Exit Sub
                     End If
                     
                Next LooPC
                
            End If
            
            Call SendData(ToIndex, UserIndex, 0, "||Tu quest ha finalizado, puedes ir a entregarla para recibir tu recompensa." & FONTTYPE_QUEST)
            .Quest.Start = 2
            Call SendData(ToIndex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "2")
            
       End With
       
End Sub

Public Sub MuereNpcQuest(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Quest As Integer)
      
      Dim LooPC As Integer
      Dim c As Integer
      
      With UserList(UserIndex)
          
          For LooPC = 1 To QuestList(Quest).NumNpc
                 
                 If QuestList(Quest).MataNpc(LooPC).NpcIndex = Npclist(NpcIndex).Numero Then
                      .Quest.MataNpc(LooPC) = .Quest.MataNpc(LooPC) + 1
                      c = c + 1
                 End If
                 
          Next LooPC
      
      End With
      
      If c > 0 Then
          Call ActualizaQuest(UserIndex, Quest)
      End If
      
End Sub

Function BuscoNpcQuest(ByVal IDNpc As Integer) As Integer
      
      Dim LooPC As Integer
      
      For LooPC = 1 To MAXNPCS
          
          If IDNpc = Npclist(LooPC).Numero Then
              BuscoNpcQuest = LooPC
              Exit Function
          End If
          
     Next LooPC
End Function

Function CompruebaIniciarQuest(ByVal UserIndex As Integer, _
                               ByVal Quest As Integer) As Integer
        
    Dim Update As Boolean
    Dim LooPC  As Integer
    Dim N As Integer
        
    With UserList(UserIndex)

             For LooPC = 1 To NumQuests
                    
                    If .Quest.UserQuest(LooPC) = 1 Then
                        N = N + 1
                    End If
                    
              Next
        
             If NumQuests = N Then
                 Update = True
             End If
        
             If Update = True Then
                 
                 For LooPC = 1 To NumQuests
                        
                        If Quest = LooPC Then
                            
                            If QuestList(Quest).Rehacer = 0 Then
                                CompruebaIniciarQuest = 0
                                Exit Function
                            ElseIf QuestList(Quest).Rehacer = 1 Then
                                CompruebaIniciarQuest = 3
                                Exit Function
                            End If
                            
                        End If
                        
                 Next LooPC
             
             ElseIf Update = False Then
                  
                  For LooPC = 1 To NumQuests
                          
                          If .Quest.UserQuest(LooPC) = 0 Then
                              
                              If LooPC = Quest Then
                                  CompruebaIniciarQuest = 0
                                  Exit Function
                              ElseIf Quest > LooPC Then
                                   CompruebaIniciarQuest = 1
                                   Exit Function
                              ElseIf Quest < LooPC Then
                                   CompruebaIniciarQuest = 2
                                   Exit Function
                              End If
                              
                          End If
                          
                  Next LooPC
                  
             End If

    End With
        
End Function
