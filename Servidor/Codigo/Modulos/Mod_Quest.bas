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
    RecompensaObjeto(1 To 10) As tRecompensaObjeto
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
    BuscaObj(1 To 10) As tBuscaObj
    NumObjsNpc As Byte
    ObjsNpc(1 To 10) As tObjsNpc
    NumNpcDD As Byte
    NpcDD As Integer
    NumMapas As Integer
    Mapas(1 To 10) As Integer
    NumDescubre As Integer
    DescubrePalabra(1 To 10) As tDescubrePalabra
End Type

Public QuestList() As tQuestList

Public Sub Load_Quest()

    Dim Quest As Integer

    Dim Loopc As Integer

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

            For Loopc = 1 To QuestList(Quest).RecompensaItem
             
                Datos = Leer.GetValue("Quest" & Quest, "RecompensaItem" & Loopc)
             
                QuestList(Quest).RecompensaObjeto(Loopc).ObjIndex = val(ReadField(1, Datos, 45))
                QuestList(Quest).RecompensaObjeto(Loopc).Amount = val(ReadField(2, Datos, 45))
             
            Next Loopc

        End If
        
        QuestList(Quest).HablarNpc = val(Leer.GetValue("Quest" & Quest, "HablarNPC"))
        
        If QuestList(Quest).HablarNpc > 0 Then
            
            For Loopc = 1 To QuestList(Quest).HablarNpc
                   QuestList(Quest).HablaNpc(Loopc) = val(Leer.GetValue("Quest" & Quest, "HablarNPC" & Loopc))
            Next Loopc
            
        End If
        
        QuestList(Quest).NUMCLASES = val(Leer.GetValue("Quest" & Quest, "Clases"))
        
        If QuestList(Quest).NUMCLASES > 0 Then
            
            For Loopc = 1 To NUMCLASES
            
               QuestList(Quest).Clases(Loopc) = CStr(Leer.GetValue("Quest" & Quest, "Clases" & Loopc))
               
            Next Loopc
            
        End If
        
        QuestList(Quest).NUMRAZAS = val(Leer.GetValue("Quest" & Quest, "Razas"))
        
        If QuestList(Quest).NUMRAZAS > 0 Then
            
            For Loopc = 1 To NUMRAZAS
                 QuestList(Quest).Razas(Loopc) = CStr(Leer.GetValue("Quest" & Quest, "Razas" & Loopc))
            Next Loopc
            
        End If
        
        QuestList(Quest).Alineacion = val(Leer.GetValue("Quest" & Quest, "Alineacion"))
        QuestList(Quest).Faccion = val(Leer.GetValue("Quest" & Quest, "Faccion"))
        QuestList(Quest).RangoFaccion = val(Leer.GetValue("Quest" & Quest, "RangoFaccion"))
        QuestList(Quest).NumNpc = val(Leer.GetValue("Quest" & Quest, "MataNPC"))
        
        If QuestList(Quest).NumNpc > 0 Then
            For Loopc = 1 To QuestList(Quest).NumNpc
                
                Datos = Leer.GetValue("Quest" & Quest, "MataNPC" & Loopc)
                
                QuestList(Quest).MataNpc(Loopc).NpcIndex = val(ReadField(1, Datos, 45))
                QuestList(Quest).MataNpc(Loopc).Cantidad = val(ReadField(2, Datos, 45))
                
            Next Loopc
        End If
        
        QuestList(Quest).NumUser = val(Leer.GetValue("Quest" & Quest, "MataUSER"))
        
        If QuestList(Quest).NumUser > 0 Then
            QuestList(Quest).MataUser.MinNivel = val(Leer.GetValue("Quest" & Quest, "MUMinNivel"))
            QuestList(Quest).MataUser.MaxNivel = val(Leer.GetValue("Quest" & Quest, "MUMaxNivel"))
            QuestList(Quest).MataUser.NUMCLASES = val(Leer.GetValue("Quest" & Quest, "MUClases"))
            
            If QuestList(Quest).MataUser.NUMCLASES > 0 Then
                   
                   For Loopc = 1 To NUMCLASES
                         QuestList(Quest).MataUser.Clases(Loopc) = CStr(Leer.GetValue("Quest" & Quest, "MUClases" & Loopc))
                   Next Loopc
                   
            End If
            
            QuestList(Quest).MataUser.NUMRAZAS = val(Leer.GetValue("Quest" & Quest, "MURazas"))
            
            If QuestList(Quest).MataUser.NUMRAZAS > 0 Then
                  
                  For Loopc = 1 To NUMRAZAS
                       QuestList(Quest).MataUser.Razas(Loopc) = CStr(Leer.GetValue("Quest" & Quest, "MURazas" & Loopc))
                  Next Loopc
                  
            End If
            
            QuestList(Quest).MataUser.Alineacion = val(Leer.GetValue("Quest" & Quest, "MUAlineacion"))
            QuestList(Quest).MataUser.Faccion = val(Leer.GetValue("Quest" & Quest, "MUFaccion"))
            QuestList(Quest).MataUser.RangoFaccion = val(Leer.GetValue("Quest" & Quest, "MURangoFaccion"))
            
        End If
        
        QuestList(Quest).NumObjs = val(Leer.GetValue("Quest" & Quest, "BuscaObjetos"))
        
        If QuestList(Quest).NumObjs > 0 Then
             
             For Loopc = 1 To QuestList(Quest).NumObjs
                    
                    Datos = Leer.GetValue("Quest" & Quest, "BuscaObjetos" & Loopc)
                    
                    QuestList(Quest).BuscaObj(Loopc).ObjIndex = val(ReadField(1, Datos, 45))
                    QuestList(Quest).BuscaObj(Loopc).Amount = val(ReadField(2, Datos, 45))
                    
             Next Loopc
             
        End If
        
        QuestList(Quest).NumObjsNpc = val(Leer.GetValue("Quest" & Quest, "ObjetoNpc"))
        
        If QuestList(Quest).NumObjsNpc > 0 Then
            
            For Loopc = 1 To QuestList(Quest).NumObjsNpc
                   
                   Datos = Leer.GetValue("Quest" & Quest, "ObjetoNpc" & Loopc)
                   
                   QuestList(Quest).ObjsNpc(Loopc).NpcIndex = val(ReadField(1, Datos, 45))
                   QuestList(Quest).ObjsNpc(Loopc).ObjIndex = val(ReadField(2, Datos, 45))
                   QuestList(Quest).ObjsNpc(Loopc).Amount = val(ReadField(3, Datos, 45))
                   
            Next Loopc
            
        End If
        
        QuestList(Quest).NumNpcDD = val(Leer.GetValue("Quest" & Quest, "NpcDD"))
        
        If QuestList(Quest).NumNpcDD > 0 Then
                   
                   Loopc = QuestList(Quest).NumNpcDD
        
                   QuestList(Quest).NpcDD = val(Leer.GetValue("Quest" & Quest, "NpcDD" & Loopc))
            
        End If
        
        QuestList(Quest).NumMapas = val(Leer.GetValue("Quest" & Quest, "EncontrarMapa"))
        
        If QuestList(Quest).NumMapas > 0 Then
            
            For Loopc = 1 To QuestList(Quest).NumMapas
                QuestList(Quest).Mapas(Loopc) = val(Leer.GetValue("Quest" & Quest, "EncontrarMapa" & Loopc))
            Next Loopc
            
        End If
        
        QuestList(Quest).NumDescubre = val(Leer.GetValue("Quest" & Quest, "DescubrePalabra"))
        
        If QuestList(Quest).NumDescubre > 0 Then
              
              For Loopc = 1 To QuestList(Quest).NumDescubre
                     
                     Datos = Leer.GetValue("Quest" & Quest, "DescubrePalabra" & Loopc)
                     
                     QuestList(Quest).DescubrePalabra(Loopc).NpcIndex = val(ReadField(1, Datos, 45))
                     QuestList(Quest).DescubrePalabra(Loopc).Frase = val(ReadField(2, Datos, 45))
                     
              Next Loopc
              
        End If
        
        frmCargando.cargar.value = frmCargando.cargar.value + 1
       
    Next Quest
  
End Sub

Public Sub IniciarVentanaQuest(ByVal UserIndex As Integer)
     
    Dim Loopc As Integer

    Dim Datos As String
     
    With UserList(UserIndex)
     
        For Loopc = 1 To NumQuests
            Datos = Datos & .Quest.UserQuest(Loopc) & ", "
        Next Loopc
      
        Datos = Left$(Datos, Len(Datos) - 2)
     
        Call SendData(Toindex, UserIndex, 0, "XU" & Datos)

    End With
        
End Sub

Public Sub ConnectQuest(ByVal UserIndex As Integer)
      
      Dim Quest As Integer
       
       With UserList(UserIndex)
          
            If .Quest.Start = 0 Then
                Call SendData(Toindex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "0")
            ElseIf .Quest.Start = 1 Then
               
               Quest = .Quest.Quest
               
               If QuestList(Quest).NumNpcDD > 0 Then
                    Call IconoNpcQuest(UserIndex, Quest)
               End If
               
               Call SendData(Toindex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "1")
            ElseIf .Quest.Start = 2 Then
               Call SendData(Toindex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "2")
            End If
       
       End With
       
End Sub

Public Sub IniciarMisionQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
       
       Dim Loopc As Integer
       Dim n As Integer
       Dim c As Integer
       Dim Datos As String
        
        With UserList(UserIndex)
        
              If .Quest.Start > 0 Then
                  Call SendData(Toindex, UserIndex, 0, "||Ya tienes una misión iniciada!! Acabala antes de volver a empezar otra." & FONTTYPE_INFO)
                  Exit Sub
              End If
              
              If CompruebaIniciarQuest(UserIndex, Quest) = 1 Then
                  Call SendData(Toindex, UserIndex, 0, "||Tienes otras misiones que realizar, antes que esta!! Revise la lista de misiones!!" & FONTTYPE_INFO)
                  Exit Sub
              ElseIf CompruebaIniciarQuest(UserIndex, Quest) = 2 Then
                  Call SendData(Toindex, UserIndex, 0, "||Debes completar todas las misiones, para poder repetir esta mision!!" & FONTTYPE_INFO)
                  Exit Sub
               ElseIf CompruebaIniciarQuest(UserIndex, Quest) = 3 Then
                  Call SendData(Toindex, UserIndex, 0, "||Esta mision no se puede rehacer, intente con otra!!" & FONTTYPE_INFO)
                  Exit Sub
              End If
                            
              If QuestList(Quest).MinNivel > 0 Then
                  If .Stats.ELV < QuestList(Quest).MinNivel Then
                      Call SendData(Toindex, UserIndex, 0, "||Para hacer esta quest, necesitas tener como minimo nivel " & QuestList(Quest).MinNivel & "." & FONTTYPE_INFO)
                      Exit Sub
                  End If
              End If
              
              If QuestList(Quest).MaxNivel > 0 Then
                  If .Stats.ELV > QuestList(Quest).MaxNivel Then
                       Call SendData(Toindex, UserIndex, 0, "||Has alcansado el nivel maximo para realizar esta misión, puedes pasar a la siguiente!" & FONTTYPE_INFO)
                       .Quest.UserQuest(Quest) = 1
                       .Quest.Quest = Quest
                      Exit Sub
                  End If
              End If
              
              If QuestList(Quest).NUMCLASES > 0 Then
                  
                  n = QuestList(Quest).NUMCLASES
                  
                  For Loopc = 1 To n
                         Debug.Print UCase$(QuestList(Quest).Clases(Loopc))
                         If UCase$(QuestList(Quest).Clases(Loopc)) = UCase$(.Clase) Then
                             c = c + 1
                         End If
                         
                  Next Loopc
                  
                  If c = 0 Then
                     Call SendData(Toindex, UserIndex, 0, "||Tu clase no esta permitida para realizar esta misión, puedes pasar a la siguiente!" & FONTTYPE_INFO)
                     .Quest.UserQuest(Quest) = 1
                     .Quest.Quest = Quest
                     Exit Sub
                  End If
                  
                  c = 0
              End If
              
              If QuestList(Quest).NUMRAZAS > 0 Then
                  
                  n = QuestList(Quest).NUMRAZAS
                  
                  For Loopc = 1 To n
                     If UCase$(QuestList(Quest).Razas(Loopc)) = UCase$(.Raza) Then
                         c = c + 1
                     End If
                  Next Loopc
                  
                  If c = 0 Then
                     Call SendData(Toindex, UserIndex, 0, "||Tu raza no esta permitida para realizar esta misión, puedes pasar a la siguiente!" & FONTTYPE_INFO)
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
                          Call SendData(Toindex, UserIndex, 0, "||No se permiten criminales en esta misión, puedes pasar a la siguiente!" & FONTTYPE_INFO)
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
                          Call SendData(Toindex, UserIndex, 0, "||No se permiten ciudadanos en esta misión, puedes pasar a la siguiente!" & FONTTYPE_INFO)
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
                          Call SendData(Toindex, UserIndex, 0, "||Solo se permiten usuarios con faccion en esta misión, puedes pasar a la siguiente!" & FONTTYPE_INFO)
                           .Quest.UserQuest(Quest) = 1
                           .Quest.Quest = Quest
                           Exit Sub
                      End If
                  End If
              End If
              
              If QuestList(Quest).RangoFaccion > 0 Then
                   If TieneFaccion(UserIndex) Then
                        If RangoFaccion(UserIndex) < QuestList(Quest).RangoFaccion Then
                            Call SendData(Toindex, UserIndex, 0, "||Necesitas la " & QuestList(Quest).RangoFaccion & " rango de tu faccion para esta misión!" & FONTTYPE_INFO)
                          
                        End If
                   Else
                        Call SendData(Toindex, UserIndex, 0, "||Solo se permiten usuarios con faccion en esta misión, puedes pasar a la siguiente!" & FONTTYPE_INFO)
                        .Quest.UserQuest(Quest) = 1
                        .Quest.Quest = Quest
                        Exit Sub
                   End If
              End If
              
              Call SendData(Toindex, UserIndex, 0, "||Has iniciado nueva misión: " & QuestList(Quest).nombre & FONTTYPE_QUEST)
              .Quest.Start = 1
              .Quest.Quest = Quest
              
              Datos = "Objetivo: "
              
              If QuestList(Quest).NumNpc > 0 Then
                  
                  For Loopc = 1 To QuestList(Quest).NumNpc
                         Datos = Datos & "Mata " & QuestList(Quest).MataNpc(Loopc).Cantidad & " " & Npclist(BuscoNpcQuest(QuestList(Quest).MataNpc(Loopc).NpcIndex)).Name & " || "
                  Next Loopc
                  
                  .Quest.NumNpc = QuestList(Quest).NumNpc
                  
              End If
              
              If QuestList(Quest).NumObjs > 0 Then
                  
                  For Loopc = 1 To QuestList(Quest).NumObjs
                         Datos = Datos & "Traeme " & QuestList(Quest).BuscaObj(Loopc).Amount & " " & ObjData(QuestList(Quest).BuscaObj(Loopc).ObjIndex).Name & " || "
                  Next Loopc
                  
                  .Quest.NumObj = QuestList(Quest).NumObjs
                  
              End If
              
              If QuestList(Quest).NumMapas > 0 Then
                
                For Loopc = 1 To QuestList(Quest).NumMapas
                       Datos = Datos & "Encuentra el mapa " & QuestList(Quest).Mapas(Loopc) & " || "
                Next Loopc
                
                .Quest.NumMap = QuestList(Quest).NumMapas
                
              End If
              
              If QuestList(Quest).NumNpcDD > 0 Then
                     Datos = Datos & "Busca/encuentra al npc y dale doble click. || "
                     .Quest.ValidNpcDD = QuestList(Quest).NumNpcDD
                     .Quest.Icono = 1
                     Call IconoNpcQuest(UserIndex, Quest)
              End If
              
              Datos = Left$(Datos, Len(Datos) - 4)
              
              Call SendData(Toindex, UserIndex, 0, "||" & Datos & FONTTYPE_GUILD)
              
              Call SendData(Toindex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "1")
              
        End With
        
End Sub

Public Sub EntregarMisionQuest(ByVal UserIndex As Integer)
        
        Dim Loopc As Integer
        Dim Quest As Integer
        
        With UserList(UserIndex)
             
             Quest = .Quest.Quest
        
             If .Quest.Start < 2 Then
                If .Quest.Start = 0 Then
                    Call SendData(Toindex, UserIndex, 0, "||Para entregar una misión, antes debes comenzar una!!" & FONTTYPE_INFO)
                    Exit Sub
                ElseIf .Quest.Start = 1 Then
                    Call SendData(Toindex, UserIndex, 0, "||Para entregar la misión, primero debes finalizarla!!" & FONTTYPE_INFO)
                    Exit Sub
                End If
             End If
             
             If QuestList(Quest).NumNpc > 0 Then
                 For Loopc = 1 To QuestList(Quest).NumNpc
                        If .Quest.MataNpc(Loopc) < QuestList(Quest).MataNpc(Loopc).Cantidad Then
                             Call SendData(Toindex, UserIndex, 0, "||Te faltan NPC's que matar antes de entregar la misión!!" & FONTTYPE_INFO)
                             Exit Sub
                        End If
                 Next Loopc
             End If
             
             If QuestList(Quest).NumObjs > 0 Then
                 For Loopc = 1 To QuestList(Quest).NumObjs
                       If .Quest.BuscaObj(Loopc) < QuestList(Quest).BuscaObj(Loopc).Amount Then
                           Call SendData(Toindex, UserIndex, 0, "||Te faltan Objetos que traerme" & FONTTYPE_INFO)
                           Exit Sub
                       End If
                       
                      If Not TieneObjetos(QuestList(Quest).BuscaObj(Loopc).ObjIndex, QuestList(Quest).BuscaObj(Loopc).Amount, UserIndex) Then
                          Call SendData(Toindex, UserIndex, 0, "||No tienes los objetos de la mision en el inventario!!" & FONTTYPE_INFO)
                          Exit Sub
                      End If
               Next Loopc
                 
             End If
             
             If QuestList(Quest).NumMapas > 0 Then
                 For Loopc = 1 To QuestList(Quest).NumMapas
                        If .Quest.Mapa(Loopc) = 0 Then
                            Call SendData(Toindex, UserIndex, 0, "||Te faltan mapas por encontrar!!" & FONTTYPE_INFO)
                            Exit Sub
                        End If
                 Next Loopc
             End If
             
             If QuestList(Quest).NumNpcDD > 0 Then
                 If .Quest.MapaNpcDD = 0 Then
                      Call SendData(Toindex, UserIndex, 0, "||Aun no le diste doble click al npc!!" & FONTTYPE_INFO)
                      Exit Sub
                 End If
             End If
             
             Call SendData(Toindex, UserIndex, 0, "||Has entregado la misión: " & QuestList(Quest).nombre & FONTTYPE_QUEST)
              
              Call RecompensaQuest(UserIndex, Quest)
              Call ResetQuest(UserIndex, Quest)
              
             .Quest.UserQuest(Quest) = 1
             .Quest.Start = 0
             Call SendData(Toindex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "0")
        End With
        
End Sub

Public Sub ActualizaQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
      
      Dim Loopc As Integer
       
       With UserList(UserIndex)
            
            If QuestList(Quest).NumNpc > 0 Then
                
                For Loopc = 1 To QuestList(Quest).NumNpc
                     
                     If .Quest.MataNpc(Loopc) < QuestList(Quest).MataNpc(Loopc).Cantidad Then
                         Exit Sub
                     End If
                     
                Next Loopc
                
            End If
            
            If QuestList(Quest).NumObjs > 0 Then
                 
                 For Loopc = 1 To QuestList(Quest).NumObjs
                       
                       If .Quest.BuscaObj(Loopc) < QuestList(Quest).BuscaObj(Loopc).Amount Then
                           Exit Sub
                       End If
                       
                 Next Loopc
                 
            End If
            
            If QuestList(Quest).NumMapas > 0 Then
            
                For Loopc = 1 To QuestList(Quest).NumMapas
                       
                       If .Quest.Mapa(Loopc) = 0 Then
                           Exit Sub
                       End If
                       
                Next Loopc
            
            End If
            
            If QuestList(Quest).NpcDD > 0 Then
                    
                    If .Quest.MapaNpcDD = 0 Then
                       Exit Sub
                    End If
                    
            End If
            
            Call SendData(Toindex, UserIndex, 0, "||Tu quest ha finalizado, puedes ir a entregarla para recibir tu recompensa." & FONTTYPE_QUEST)
            .Quest.Start = 2
            Call SendData(Toindex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "2")
            
       End With
       
End Sub

Public Sub MuereNpcQuest(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Quest As Integer)
      
      Dim Loopc As Integer
      Dim c As Integer
      
      With UserList(UserIndex)
          
          For Loopc = 1 To QuestList(Quest).NumNpc
                 
                 If QuestList(Quest).MataNpc(Loopc).NpcIndex = Npclist(NpcIndex).Numero Then
                      .Quest.MataNpc(Loopc) = .Quest.MataNpc(Loopc) + 1
                      
                      'If QuestList(Quest).MataNpc(LoopC).Cantidad <= .Quest.MataNpc(LoopC) Then
                      '    Call SendData(ToPCArea, UserIndex, .pos.Map, "||" & vbCyan & "°Mata a " & Npclist(NpcIndex).Name & " (" & .Quest.MataNpc(LoopC) & "/" & QuestList(Quest).MataNpc(LoopC).Cantidad & ")°" & CStr(.char.CharIndex))
                      'End If
                      
                      c = c + 1
                 End If
                 
          Next Loopc
        
      End With
      
      If c > 0 Then
          Call ActualizaQuest(UserIndex, Quest)
      End If
      
End Sub

Public Sub BuscaObjQuest(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Amount As Integer, ByVal Quest As Integer)
     
     Dim Loopc As Integer
     Dim c As Integer
     
     With UserList(UserIndex)
         
         If QuestList(Quest).NumObjs > 0 Then
             
             For Loopc = 1 To QuestList(Quest).NumObjs
                     If QuestList(Quest).BuscaObj(Loopc).ObjIndex = ObjIndex Then
                          .Quest.BuscaObj(Loopc) = .Quest.BuscaObj(Loopc) + Amount
                          c = c + 1
                     End If
             Next Loopc
             
         End If
           
         If c > 0 Then
            Call ActualizaQuest(UserIndex, Quest)
         End If
         
     End With
     
End Sub

Public Sub EncuentraMapaQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
     
     Dim Loopc As Integer
     Dim Map As Integer
     Dim c As Integer
     
     With UserList(UserIndex)
          
          Map = .pos.Map
     
          For Loopc = 1 To QuestList(Quest).NumMapas
                  If QuestList(Quest).Mapas(Loopc) = Map Then
                       .Quest.Mapa(Loopc) = 1
                       c = c + 1
                  End If
          Next Loopc
          
          If c > 0 Then
            Call ActualizaQuest(UserIndex, Quest)
         End If
          
     End With
     
End Sub

Public Sub ClickMisionesQuest(ByVal UserIndex As Integer)
      Dim Quest As Integer
      
      With UserList(UserIndex)
            
            Quest = .Quest.Quest
            
            If .Quest.Start <> 1 Then Exit Sub
            
            If QuestList(Quest).NumNpcDD > 0 Then
                 Call DobleClickNpcQuest(UserIndex, Quest)
            End If
            
      End With
End Sub

Public Sub DobleClickNpcQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
      
       Dim Map As Integer
       Dim c As Byte
       
       With UserList(UserIndex)
         
         Map = .pos.Map
         
         If QuestList(Quest).NumNpcDD > 0 Then
             If QuestList(Quest).NpcDD = Map Then
                Call SendData(ToPCArea, UserIndex, .pos.Map, "||" & vbCyan & "°¡Le has dado Doble Click!°" & CStr(.char.CharIndex))
                 .Quest.MapaNpcDD = 1
                 c = c + 1
             End If
         End If
         
         If c > 0 Then
            Call ActualizaQuest(UserIndex, Quest)
         End If
         
       End With
       
End Sub

Public Sub RecompensaQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
      Dim Loopc As Integer
      Dim Obj As Obj
      
      With UserList(UserIndex)
      
             If QuestList(Quest).RecompensaOro > 0 Then
                 .Stats.GLD = .Stats.GLD + QuestList(Quest).RecompensaOro
                 Call EnviarOro(UserIndex)
             End If
             
             If QuestList(Quest).RecompensaExp > 0 Then
                 If .Stats.ELV < STAT_MAXELV Then
                     .Stats.Exp = .Stats.Exp + QuestList(Quest).RecompensaExp
                     Call EnviarExp(UserIndex)
                 End If
             End If
             
             If QuestList(Quest).RecompensaItem > 0 Then
                 
                 For Loopc = 1 To QuestList(Quest).RecompensaItem
                         
                         Obj.ObjIndex = QuestList(Quest).RecompensaObjeto(Loopc).ObjIndex
                         Obj.Amount = QuestList(Quest).RecompensaObjeto(Loopc).Amount
                         
                         Call MeterItemEnInventario(UserIndex, Obj)
                         
                 Next Loopc
             
             End If
             
      End With
      
End Sub

Public Sub IconoNpcQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
        
        Dim Map As Integer
        Dim Loopc As Integer
         
         With UserList(UserIndex)
               
               If QuestList(Quest).NumNpcDD > 0 Then
                   
                   Map = QuestList(Quest).NpcDD
                   
                   For Loopc = 1 To NumNPCs
                   
                   If Npclist(Loopc).NPCtype = eNPCType.Misiones Then
                       
                       If Npclist(Loopc).pos.Map = Map Then
                           
                           If .Quest.Icono = 0 Then
                                Call SendData(Toindex, UserIndex, 0, "XI" & Npclist(Loopc).char.CharIndex & "," & 0)
                           ElseIf .Quest.Icono = 1 Then
                                Call SendData(Toindex, UserIndex, 0, "XI" & Npclist(Loopc).char.CharIndex & "," & 1)
                           End If
                       
                       End If
                       
                   End If
                   
                   Next Loopc
               End If
               
         End With
         
End Sub

Public Sub ResetQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
       Dim Loopc As Integer
       
       With UserList(UserIndex)
            
            If QuestList(Quest).NumNpc > 0 Then
                For Loopc = 1 To QuestList(Quest).NumNpc
                
                       .Quest.MataNpc(Loopc) = 0
                       
                Next Loopc
                
                .Quest.NumNpc = 0
                
            End If
            
            If QuestList(Quest).NumObjs > 0 Then
                
                For Loopc = 1 To QuestList(Quest).NumObjs
                
                      .Quest.BuscaObj(Loopc) = 0
                      Call QuitarObjetos(QuestList(Quest).BuscaObj(Loopc).ObjIndex, QuestList(Quest).BuscaObj(Loopc).Amount, UserIndex)
                
                Next Loopc
                
                .Quest.NumObj = 0
                
            End If
            
            If QuestList(Quest).NumMapas > 0 Then
                
                For Loopc = 1 To QuestList(Quest).NumMapas
                       .Quest.Mapa(Loopc) = 0
                Next Loopc
                
                .Quest.NumMap = 0
                
            End If
            
            If QuestList(Quest).NumNpcDD > 0 Then
                 .Quest.ValidNpcDD = 0
                 .Quest.MapaNpcDD = 0
                 .Quest.Icono = 0
                 Call IconoNpcQuest(UserIndex, Quest)
            End If
            
       End With
       
End Sub

Function BuscoNpcQuest(ByVal IDNpc As Integer) As Integer
      
      Dim Loopc As Integer
      
      For Loopc = 1 To MAXNPCS
          
          If IDNpc = Npclist(Loopc).Numero Then
              BuscoNpcQuest = Loopc
              Exit Function
          End If
          
     Next Loopc
End Function

Function CompruebaIniciarQuest(ByVal UserIndex As Integer, _
                               ByVal Quest As Integer) As Integer
        
    Dim Update As Boolean
    Dim Loopc  As Integer
    Dim n As Integer
        
    With UserList(UserIndex)

             For Loopc = 1 To NumQuests
                    
                    If .Quest.UserQuest(Loopc) = 1 Then
                        n = n + 1
                    End If
                    
              Next
        
             If NumQuests = n Then
                 Update = True
             End If
        
             If Update = True Then
                 
                 For Loopc = 1 To NumQuests
                        
                        If Quest = Loopc Then
                            
                            If QuestList(Quest).Rehacer = 0 Then
                                CompruebaIniciarQuest = 0
                                Exit Function
                            ElseIf QuestList(Quest).Rehacer = 1 Then
                                CompruebaIniciarQuest = 3
                                Exit Function
                            End If
                            
                        End If
                        
                 Next Loopc
             
             ElseIf Update = False Then
                  
                  For Loopc = 1 To NumQuests
                          
                          If .Quest.UserQuest(Loopc) = 0 Then
                              
                              If Loopc = Quest Then
                                  CompruebaIniciarQuest = 0
                                  Exit Function
                              ElseIf Quest > Loopc Then
                                   CompruebaIniciarQuest = 1
                                   Exit Function
                              ElseIf Quest < Loopc Then
                                   CompruebaIniciarQuest = 2
                                   Exit Function
                              End If
                              
                          End If
                          
                  Next Loopc
                  
             End If

    End With
        
End Function
