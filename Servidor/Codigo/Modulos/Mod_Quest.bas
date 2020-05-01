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
      ObjIndex As Integer
      Amount As Integer
End Type

Public Type tDescubrePalabra
     Mapa As Integer
     Pregunta As String
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
    NumHablarNpc As Byte
    MapaHablaNpc As Integer
    NumMsjHablar As Integer
    MsjHablar(1 To 10) As String
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
    CantidadMataUser As Integer
    NumObjs As Byte
    BuscaObj(1 To 10) As tBuscaObj
    NumObjsNpc As Byte
    ObjsNpc(1 To 10) As tObjsNpc
    MapaObjsNpc As Integer
    NumNpcDD As Byte
    NpcDD As Integer
    NumMapas As Integer
    Mapas(1 To 10) As Integer
    NumDescubre As Integer
    DescubrePalabra As tDescubrePalabra
End Type

Public Type tQuestDesc
       
       DobleClick As String
       Descubridor As String
       Hablador As String
       DarObjNpc As String
       
End Type

Public QuestList() As tQuestList
Public QuestDesc As tQuestDesc

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

            For LooPC = 1 To QuestList(Quest).RecompensaItem
             
                Datos = Leer.GetValue("Quest" & Quest, "RecompensaItem" & LooPC)
             
                QuestList(Quest).RecompensaObjeto(LooPC).ObjIndex = val(ReadField(1, Datos, 45))
                QuestList(Quest).RecompensaObjeto(LooPC).Amount = val(ReadField(2, Datos, 45))
             
            Next LooPC

        End If
        
        QuestList(Quest).NumHablarNpc = val(Leer.GetValue("Quest" & Quest, "HablarNPC"))
        
        If QuestList(Quest).NumHablarNpc > 0 Then
                                   
               QuestList(Quest).MapaHablaNpc = val(Leer.GetValue("Quest" & Quest, "MapaHablarNPC"))
               QuestList(Quest).NumMsjHablar = val(Leer.GetValue("Quest" & Quest, "NumMsjHablar"))
               
               For LooPC = 1 To QuestList(Quest).NumMsjHablar
                      
                      QuestList(Quest).MsjHablar(LooPC) = Leer.GetValue("Quest" & Quest, "MsjHablar" & LooPC)
                      
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
            For LooPC = 1 To QuestList(Quest).NumNpc
                
                Datos = Leer.GetValue("Quest" & Quest, "MataNPC" & LooPC)
                
                QuestList(Quest).MataNpc(LooPC).NpcIndex = val(ReadField(1, Datos, 45))
                QuestList(Quest).MataNpc(LooPC).Cantidad = val(ReadField(2, Datos, 45))
                
            Next LooPC
        End If
        
        QuestList(Quest).NumUser = val(Leer.GetValue("Quest" & Quest, "NumMataUser"))
        
        If QuestList(Quest).NumUser > 0 Then
            
            QuestList(Quest).CantidadMataUser = val(Leer.GetValue("Quest" & Quest, "CantidadMataUser"))
            
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
             
             For LooPC = 1 To QuestList(Quest).NumObjs
                    
                    Datos = Leer.GetValue("Quest" & Quest, "BuscaObjetos" & LooPC)
                    
                    QuestList(Quest).BuscaObj(LooPC).ObjIndex = val(ReadField(1, Datos, 45))
                    QuestList(Quest).BuscaObj(LooPC).Amount = val(ReadField(2, Datos, 45))
                    
             Next LooPC
             
        End If
        
        QuestList(Quest).NumObjsNpc = val(Leer.GetValue("Quest" & Quest, "ObjetoNpc"))
        
        If QuestList(Quest).NumObjsNpc > 0 Then
            
            For LooPC = 1 To QuestList(Quest).NumObjsNpc
                   
                   Datos = Leer.GetValue("Quest" & Quest, "ObjetoNpc" & LooPC)
                   
                   QuestList(Quest).MapaObjsNpc = val(ReadField(1, Datos, 45))
                   QuestList(Quest).ObjsNpc(LooPC).ObjIndex = val(ReadField(2, Datos, 45))
                   QuestList(Quest).ObjsNpc(LooPC).Amount = val(ReadField(3, Datos, 45))
                   
            Next LooPC
            
        End If
        
        QuestList(Quest).NumNpcDD = val(Leer.GetValue("Quest" & Quest, "NpcDD"))
        
        If QuestList(Quest).NumNpcDD > 0 Then
                   
                   LooPC = QuestList(Quest).NumNpcDD
        
                   QuestList(Quest).NpcDD = val(Leer.GetValue("Quest" & Quest, "NpcDD" & LooPC))
            
        End If
        
        QuestList(Quest).NumMapas = val(Leer.GetValue("Quest" & Quest, "EncontrarMapa"))
        
        If QuestList(Quest).NumMapas > 0 Then
            
            For LooPC = 1 To QuestList(Quest).NumMapas
                QuestList(Quest).Mapas(LooPC) = val(Leer.GetValue("Quest" & Quest, "EncontrarMapa" & LooPC))
            Next LooPC
            
        End If
        
        QuestList(Quest).NumDescubre = val(Leer.GetValue("Quest" & Quest, "DescubrePalabra"))
        
        If QuestList(Quest).NumDescubre > 0 Then
                                   
                      LooPC = QuestList(Quest).NumDescubre
                      
                      Datos = Leer.GetValue("Quest" & Quest, "DescubrePalabra" & LooPC)
                                   
                     QuestList(Quest).DescubrePalabra.Mapa = val(ReadField(1, Datos, 45))
                     QuestList(Quest).DescubrePalabra.Pregunta = ReadField(2, Datos, 45)
                     QuestList(Quest).DescubrePalabra.Frase = ReadField(3, Datos, 45)
              
        End If
        
        frmCargando.cargar.value = frmCargando.cargar.value + 1
       
    Next Quest
    
    QuestDesc.DobleClick = "¡Me has encontrado! ¡Hazme doble click para realizar tu mision"
    QuestDesc.Descubridor = "¡Me has encontrado! ¡Clickeame dos veces para saber que pregunta es!"
    QuestDesc.DarObjNpc = "¡Me has encontrado! Si ya tienes mis objetos escribe /QUESTENTREGA"
    QuestDesc.Hablador = "¡Me has encontrado! ¡Clickeame dos veces y escucha mi historia"
    
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
      
      Dim Quest As Integer
       
       With UserList(UserIndex)
          
            If .Quest.Start = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "0")
            ElseIf .Quest.Start = 1 Then
               
               Quest = .Quest.Quest
               
               If QuestList(Quest).NumNpcDD > 0 Then
                    Call IconoNpcQuest(UserIndex, Quest)
               End If
               
               Call SendData(ToIndex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "1")
            ElseIf .Quest.Start = 2 Then
               Call SendData(ToIndex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "2")
            End If
       
       End With
       
End Sub

Public Sub UserMataQuest(ByVal UserIndex As Integer, ByVal Victima As Integer, ByVal Quest As Integer)
        
        Dim LooPC As Integer
        Dim c As Integer
        Dim n As Integer
        
        With UserList(UserIndex)
        
           If QuestList(Quest).MataUser.MinNivel > UserList(Victima).Stats.ELV Then
                       Exit Sub
              End If
              
           If QuestList(Quest).MataUser.MaxNivel > UserList(Victima).Stats.ELV Then
                      Exit Sub
           End If
              
           If QuestList(Quest).MataUser.NUMCLASES > 0 Then
                  
                  n = QuestList(Quest).MataUser.NUMCLASES
                  
                  For LooPC = 1 To n
                         If UCase$(QuestList(Quest).MataUser.Clases(LooPC)) = UCase$(UserList(Victima).Clase) Then
                             c = c + 1
                         End If
                  Next LooPC
                  
                  If c = 0 Then
                     Exit Sub
                  End If
                  
                  c = 0
           End If
           
           If QuestList(Quest).MataUser.NUMRAZAS > 0 Then
                  
                  n = QuestList(Quest).MataUser.NUMRAZAS
                  
                  For LooPC = 1 To n
                     If UCase$(QuestList(Quest).MataUser.Razas(LooPC)) = UCase$(UserList(Victima).Raza) Then
                         c = c + 1
                     End If
                  Next LooPC
                  
                  If c = 0 Then
                     Exit Sub
                  End If
                  
                  c = 0
             End If
             
             If QuestList(Quest).MataUser.Alineacion > 0 Then
                  
                  If QuestList(Quest).MataUser.Alineacion = 1 Then
                      
                      If Criminal(Victima) Then
                         c = 0
                      Else
                         c = c + 1
                      End If
                      
                      If c = 0 Then
                          Exit Sub
                      End If
                       c = 0
                  End If
                  
                  If QuestList(Quest).Alineacion = 2 Then
                      
                      If Criminal(Victima) Then
                         c = c + 1
                         Else
                         c = 0
                      End If
                      
                      If c = 0 Then
                          Exit Sub
                      End If
                      c = 0
                  End If
              
              End If
              
              If QuestList(Quest).MataUser.Faccion > 0 Then
                  If QuestList(Quest).MataUser.Faccion = 1 Then
                      If Not TieneFaccion(Victima) Then
                           Exit Sub
                      End If
                  End If
              End If
              
              If QuestList(Quest).MataUser.RangoFaccion > 0 Then
                   If TieneFaccion(Victima) Then
                        If RangoFaccion(Victima) < QuestList(Quest).MataUser.RangoFaccion Then
                            Exit Sub
                        End If
                   Else
                        Exit Sub
                   End If
              End If
              
               If QuestList(Quest).NumUser > 0 Then
                   c = 0
                   
                   If .Quest.UserMatados <= QuestList(Quest).CantidadMataUser Then
                       .Quest.UserMatados = .QuestMatados + 1
                       Call SendData(ToIndex, UserIndex, 0, "||Has matado a un usuario! (" & .Quest.UserMatados & "/" & QuestList(Quest).CantidadMataUser & ")" & FONTTYPE_GUILD)
                       c = c + 1
                   End If
                   
               End If
               
               If c > 0 Then
                  Call ActualizaQuest(UserIndex, Quest)
               End If
               
        End With
        
End Sub

Public Sub IniciarMisionQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
       
       Dim LooPC As Integer
       Dim n As Integer
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
                  
                  n = QuestList(Quest).NUMCLASES
                  
                  For LooPC = 1 To n
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
                  
                  n = QuestList(Quest).NUMRAZAS
                  
                  For LooPC = 1 To n
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
                             Exit Sub
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
                         Datos = Datos & "Mata " & QuestList(Quest).MataNpc(LooPC).Cantidad & " " & Npclist(BuscoNpcQuest(QuestList(Quest).MataNpc(LooPC).NpcIndex)).Name & " || "
                  Next LooPC
                  
                  .Quest.NumNpc = QuestList(Quest).NumNpc
                  
              End If
              
              If QuestList(Quest).NumObjs > 0 Then
                  
                  For LooPC = 1 To QuestList(Quest).NumObjs
                         Datos = Datos & "Traeme " & QuestList(Quest).BuscaObj(LooPC).Amount & " " & ObjData(QuestList(Quest).BuscaObj(LooPC).ObjIndex).Name & " || "
                  Next LooPC
                  
                  .Quest.NumObj = QuestList(Quest).NumObjs
                  
              End If
              
              If QuestList(Quest).NumMapas > 0 Then
                
                For LooPC = 1 To QuestList(Quest).NumMapas
                       Datos = Datos & "Encuentra el mapa " & QuestList(Quest).Mapas(LooPC) & " || "
                Next LooPC
                
                .Quest.NumMap = QuestList(Quest).NumMapas
                
              End If
              
              If QuestList(Quest).NumNpcDD > 0 Then
                     Datos = Datos & "Busca/encuentra al npc y dale doble click. || "
                     .Quest.ValidNpcDD = QuestList(Quest).NumNpcDD
                     .Quest.Icono = 1
                     Call IconoNpcQuest(UserIndex, Quest)
              End If
              
              If QuestList(Quest).NumDescubre > 0 Then
                   Datos = Datos & "Busca/encuentra al npc y responde su pregunta. || "
                    .Quest.ValidNpcDescubre = QuestList(Quest).NumDescubre
                    .Quest.Icono = 1
                    Call IconoNpcQuest(UserIndex, Quest)
              End If
              
              If QuestList(Quest).NumObjsNpc > 0 Then
                 For LooPC = 1 To QuestList(Quest).NumObjsNpc
                        Datos = Datos & "Llevale " & QuestList(Quest).ObjsNpc(LooPC).Amount & " " & ObjData(QuestList(Quest).ObjsNpc(LooPC).ObjIndex).Name & " al npc de mision escondido. || "
                 Next LooPC
                  .Quest.NumObjNpc = QuestList(Quest).NumObjsNpc
                  .Quest.Icono = 1
                  Call IconoNpcQuest(UserIndex, Quest)
              End If
              
              If QuestList(Quest).NumHablarNpc > 0 Then
                  Datos = Datos & "Busca y habla con el npc de misiones del mapa " & QuestList(Quest).MapaHablaNpc & " || "
                  .Quest.ValidHablarNpc = QuestList(Quest).NumHablarNpc
                  .Quest.Icono = 1
                  Call IconoNpcQuest(UserIndex, Quest)
              End If
              
              If QuestList(Quest).NumUser > 0 Then
                  Datos = Datos & "Debes matar a " & QuestList(Quest).CantidadMataUser & " usuarios. || "
                  .Quest.ValidMatarUser = QuestList(Quest).NumUser
              End If
              
              Datos = Left$(Datos, Len(Datos) - 4)
              
              Call SendData(ToIndex, UserIndex, 0, "||" & Datos & FONTTYPE_GUILD)
              
              Call SendData(ToIndex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "1")
              
        End With
        
End Sub

Public Sub EntregarMisionQuest(ByVal UserIndex As Integer)
        
        Dim LooPC As Integer
        Dim Quest As Integer
        
        With UserList(UserIndex)
             
             Quest = .Quest.Quest
        
             If .Quest.Start < 2 Then
                If .Quest.Start = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Para entregar una misión, antes debes comenzar una!!" & FONTTYPE_INFO)
                    Exit Sub
                ElseIf .Quest.Start = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Para entregar la misión, primero debes finalizarla!!" & FONTTYPE_INFO)
                    Exit Sub
                End If
             End If
             
             If QuestList(Quest).NumNpc > 0 Then
                 For LooPC = 1 To QuestList(Quest).NumNpc
                        If .Quest.MataNpc(LooPC) < QuestList(Quest).MataNpc(LooPC).Cantidad Then
                             Call SendData(ToIndex, UserIndex, 0, "||Te faltan NPC's que matar antes de entregar la misión!!" & FONTTYPE_INFO)
                             Exit Sub
                        End If
                 Next LooPC
             End If
             
             If QuestList(Quest).NumObjs > 0 Then
                 For LooPC = 1 To QuestList(Quest).NumObjs
                       If .Quest.BuscaObj(LooPC) < QuestList(Quest).BuscaObj(LooPC).Amount Then
                           Call SendData(ToIndex, UserIndex, 0, "||Te faltan Objetos que traerme" & FONTTYPE_INFO)
                           Exit Sub
                       End If
                       
                      If Not TieneObjetos(QuestList(Quest).BuscaObj(LooPC).ObjIndex, QuestList(Quest).BuscaObj(LooPC).Amount, UserIndex) Then
                          Call SendData(ToIndex, UserIndex, 0, "||No tienes los objetos de la mision en el inventario!!" & FONTTYPE_INFO)
                          Exit Sub
                      End If
               Next LooPC
                 
             End If
             
             If QuestList(Quest).NumMapas > 0 Then
                 For LooPC = 1 To QuestList(Quest).NumMapas
                        If .Quest.Mapa(LooPC) = 0 Then
                            Call SendData(ToIndex, UserIndex, 0, "||Te faltan mapas por encontrar!!" & FONTTYPE_INFO)
                            Exit Sub
                        End If
                 Next LooPC
             End If
             
             If QuestList(Quest).NumNpcDD > 0 Then
                 If .Quest.MapaNpcDD = 0 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Aun no le diste doble click al npc!!" & FONTTYPE_INFO)
                      Exit Sub
                 End If
             End If
             
             If QuestList(Quest).NumDescubre > 0 Then
                 If .Quest.PreguntaDescubre = 0 Then
                       Call SendData(ToIndex, UserIndex, 0, "||Te falta responde la pregunta al npc!!" & FONTTYPE_INFO)
                     Exit Sub
                 End If
             End If
             
             If QuestList(Quest).NumObjsNpc > 0 Then
                  If .Quest.DarObjNpcEntrega = 0 Then
                       Call SendData(ToIndex, UserIndex, 0, "||No has entregado los objetos al npc de misiones!!" & FONTTYPE_INFO)
                       Exit Sub
                  End If
             End If
             
             If QuestList(Quest).NumHablarNpc > 0 Then
                  If .Quest.UserHablaNpc = 0 Then
                      Call SendData(ToIndex, UserIndex, 0, "||No has hablado con el npc de mision!!" & FONTTYPE_INFO)
                      Exit Sub
                  End If
             End If
             
             If QuestList(Quest).NumUser > 0 Then
                 If .Quest.UserMatados < QuestList(Quest).CantidadMataUser Then
                     Call SendData(ToIndex, UserIndex, 0, "||Te faltan usuarios por matar!!" & FONTTYPE_INFO)
                     Exit Sub
                 End If
             End If
             
             Call SendData(ToIndex, UserIndex, 0, "||Has entregado la misión: " & QuestList(Quest).nombre & FONTTYPE_QUEST)
              
              Call RecompensaQuest(UserIndex, Quest)
              Call ResetQuest(UserIndex, Quest)
              
             .Quest.UserQuest(Quest) = 1
             .Quest.Start = 0
             Call SendData(ToIndex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "0")
        End With
        
End Sub

Public Sub ActualizaQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
      
      Dim LooPC As Integer
       
       With UserList(UserIndex)
            
            If QuestList(Quest).NumNpc > 0 Then
                
                For LooPC = 1 To QuestList(Quest).NumNpc
                     
                     If .Quest.MataNpc(LooPC) < QuestList(Quest).MataNpc(LooPC).Cantidad Then
                         Exit Sub
                     End If
                     
                Next LooPC
                
            End If
            
            If QuestList(Quest).NumObjs > 0 Then
                 
                 For LooPC = 1 To QuestList(Quest).NumObjs
                       
                       If .Quest.BuscaObj(LooPC) < QuestList(Quest).BuscaObj(LooPC).Amount Then
                           Exit Sub
                       End If
                       
                 Next LooPC
                 
            End If
            
            If QuestList(Quest).NumMapas > 0 Then
            
                For LooPC = 1 To QuestList(Quest).NumMapas
                       
                       If .Quest.Mapa(LooPC) = 0 Then
                           Exit Sub
                       End If
                       
                Next LooPC
            
            End If
            
            If QuestList(Quest).NpcDD > 0 Then
                    
                    If .Quest.MapaNpcDD = 0 Then
                       Exit Sub
                    End If
                    
            End If
            
            If QuestList(Quest).NumDescubre > 0 Then
                    
                    If .Quest.PreguntaDescubre = 0 Then
                        Exit Sub
                    End If
                    
            End If
            
            If QuestList(Quest).NumObjsNpc > 0 Then
                  
                  If .Quest.DarObjNpcEntrega = 0 Then
                       Exit Sub
                  End If
                  
            End If
            
            If QuestList(Quest).NumHablarNpc > 0 Then
                  
                  If .Quest.UserHablaNpc = 0 Then
                      Exit Sub
                  End If
                  
            End If
            
            If QuestList(Quest).NumUser > 0 Then
                 
                 If .Quest.UserMatados < QuestList(Quest).CantidadMataUser Then
                     Exit Sub
                 End If
                 
            End If
            
            Call SendData(ToIndex, UserIndex, 0, "||Tu quest ha finalizado, puedes ir a entregarla para recibir tu recompensa." & FONTTYPE_QUEST)
            .Quest.Start = 2
            Call SendData(ToIndex, UserIndex, 0, "XP" & UserList(UserIndex).char.CharIndex & "," & "2")
            
       End With
       
End Sub

Public Sub ActualizaObjNpc(ByVal UserIndex As Integer, ByVal Quest As Integer)
       
       Dim LooPC As Integer
       
        With UserList(UserIndex)
        
            If .Quest.NumObjNpc > 0 Then
               For LooPC = 1 To .Quest.NumObjNpc
                If .Quest.DarObjNpc(LooPC) < QuestList(Quest).ObjsNpc(LooPC).Amount Then
                    Exit Sub
                End If
                Next LooPC
            End If
          
          Call SendData(ToIndex, UserIndex, 0, "||Has conseguido todos los items que debes entregar al npc de misiones, buscalo y entregaselo!" & FONTTYPE_GUILD)
          
        End With
End Sub

Public Sub MuereNpcQuest(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Quest As Integer)
      
      Dim LooPC As Integer
      Dim c As Integer
      
      With UserList(UserIndex)
          
          For LooPC = 1 To QuestList(Quest).NumNpc
                 
                 If QuestList(Quest).MataNpc(LooPC).NpcIndex = Npclist(NpcIndex).Numero Then
                      .Quest.MataNpc(LooPC) = .Quest.MataNpc(LooPC) + 1
                      
                      If QuestList(Quest).MataNpc(LooPC).Cantidad >= .Quest.MataNpc(LooPC) Then
                          Call SendData(ToIndex, UserIndex, 0, "||Mata a " & Npclist(NpcIndex).Name & " (" & .Quest.MataNpc(LooPC) & "/" & QuestList(Quest).MataNpc(LooPC).Cantidad & ")" & FONTTYPE_GUILD)
                      End If
                      
                      c = c + 1
                 End If
                 
          Next LooPC
        
      End With
      
      If c > 0 Then
          Call ActualizaQuest(UserIndex, Quest)
      End If
      
End Sub

Public Sub BuscaObjNpcQuest(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Amount As Integer, ByVal Quest As Integer)
     
     Dim LooPC As Integer
     Dim c As Integer
     
     With UserList(UserIndex)
          
          If QuestList(Quest).NumObjsNpc > 0 Then
              
              For LooPC = 1 To QuestList(Quest).NumObjsNpc
                     
                     If QuestList(Quest).ObjsNpc(LooPC).ObjIndex = ObjIndex Then
                     
                         .Quest.DarObjNpc(LooPC) = .Quest.DarObjNpc(LooPC) + Amount
                         
                         If QuestList(Quest).ObjsNpc(LooPC).Amount >= .Quest.DarObjNpc(LooPC) Then
                            Call SendData(ToIndex, UserIndex, 0, "||" & vbCyan & "°¡Has conseguido " & Amount & " " & ObjData(ObjIndex).Name & " (" & .Quest.DarObjNpc(LooPC) & "/" & QuestList(Quest).ObjsNpc(LooPC).Amount & ")!°" & CStr(.char.CharIndex))
                         End If
                         
                         c = c + 1
                     End If
                     
              Next LooPC
              
          End If
          
          If c > 0 Then
             Call ActualizaObjNpc(UserIndex, Quest)
          End If
          
     End With
     
End Sub

Public Sub BuscaObjQuest(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Amount As Integer, ByVal Quest As Integer)
     
     Dim LooPC As Integer
     Dim c As Integer
     
     With UserList(UserIndex)
         
         If QuestList(Quest).NumObjs > 0 Then
             
             For LooPC = 1 To QuestList(Quest).NumObjs
                     If QuestList(Quest).BuscaObj(LooPC).ObjIndex = ObjIndex Then
                          .Quest.BuscaObj(LooPC) = .Quest.BuscaObj(LooPC) + Amount
                          
                          If QuestList(Quest).BuscaObj(LooPC).Amount >= .Quest.BuscaObj(LooPC) Then
                              Call SendData(ToIndex, UserIndex, 0, "||" & vbCyan & "°¡Has conseguido " & Amount & " " & ObjData(ObjIndex).Name & " (" & .Quest.BuscaObj(LooPC) & "/" & QuestList(Quest).BuscaObj(LooPC).Amount & ")!°" & CStr(.char.CharIndex))
                          End If
                          
                          c = c + 1
                     End If
             Next LooPC
             
         End If
           
         If c > 0 Then
            Call ActualizaQuest(UserIndex, Quest)
         End If
         
     End With
     
End Sub

Public Sub EncuentraMapaQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
     
     Dim LooPC As Integer
     Dim Map As Integer
     Dim c As Integer
     
     With UserList(UserIndex)
          
          Map = .pos.Map
     
          For LooPC = 1 To QuestList(Quest).NumMapas
                If .Quest.Mapa(LooPC) = 0 Then
                  If QuestList(Quest).Mapas(LooPC) = Map Then
                       .Quest.Mapa(LooPC) = 1
                      Call SendData(ToPCArea, UserIndex, .pos.Map, "||" & vbCyan & "°¡Has encontrado un mapa!°" & CStr(.char.CharIndex))
                       c = c + 1
                  End If
                End If
          Next LooPC
          
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
            
            If QuestList(Quest).NumDescubre > 0 Then
                 Call DescubreNpcQuest(UserIndex, Quest)
            End If
            
            If QuestList(Quest).NumHablarNpc > 0 Then
                  Call EnviaVentanaHablarQuest(UserIndex, Quest)
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
                Call SendData(ToIndex, UserIndex, 0, "||" & vbCyan & "°¡Le has dado Doble Click!°" & CStr(.char.CharIndex))
                 .Quest.MapaNpcDD = 1
                 c = c + 1
             End If
         End If
         
         If c > 0 Then
            Call ActualizaQuest(UserIndex, Quest)
         End If
         
       End With
       
End Sub

Public Sub DescubreNpcQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
      
      Dim Map As Integer
      Dim c As Byte
      Dim LooPC As Integer
      Dim NpcIndex As Integer
      
      With UserList(UserIndex)
           
           Map = .pos.Map
           
           If QuestList(Quest).NumDescubre > 0 Then
                If QuestList(Quest).DescubrePalabra.Mapa = Map Then
                    If .Quest.PreguntaDescubre = 0 Then
                        
                        For LooPC = 1 To NumNPCs
                               
                               If Npclist(LooPC).NPCtype = eNPCType.Misiones Then
                                    
                                   If Npclist(LooPC).pos.Map = Map Then
                                      
                                       Call SendData(ToIndex, UserIndex, 0, "||" & vbCyan & "°" & QuestList(Quest).DescubrePalabra.Pregunta & "°" & CStr(Npclist(LooPC).char.CharIndex))
                                      
                                   End If
                                    
                               End If
                               
                        Next LooPC
                        
                    End If
                End If
           End If
                      
      End With
      
End Sub

Public Sub RespuestaNpcQuest(ByVal UserIndex As Integer, ByVal Quest As Integer, ByVal Mensaje As String)
       
      Dim LooPC As Integer
      Dim Map As Integer
      Dim c As Integer
           
      With UserList(UserIndex)
          
          If .Quest.Start <> 1 Then Exit Sub
          
          If QuestList(Quest).NumDescubre > 0 Then
              
              Map = .pos.Map
              
              If QuestList(Quest).DescubrePalabra.Mapa = Map Then
              
                  If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = eNPCType.Misiones Then
                      If .Quest.PreguntaDescubre = 0 Then
                          
                          If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                              Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                              Exit Sub
                           End If
                           
                           If UCase$(QuestList(Quest).DescubrePalabra.Frase) = UCase$(Mensaje) Then
                               Call SendData(ToIndex, UserIndex, 0, "||¡Respuesta correcta!" & FONTTYPE_GUILD)
                               .Quest.PreguntaDescubre = 1
                               c = c + 1
                           End If
                           
                      End If
                      
                  End If
              
              End If
              
          End If
          
          If c > 0 Then
            Call ActualizaQuest(UserIndex, Quest)
         End If
      
      End With
           
End Sub

Public Sub RecompensaQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
      Dim LooPC As Integer
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
                 
                 For LooPC = 1 To QuestList(Quest).RecompensaItem
                         
                         Obj.ObjIndex = QuestList(Quest).RecompensaObjeto(LooPC).ObjIndex
                         Obj.Amount = QuestList(Quest).RecompensaObjeto(LooPC).Amount
                         
                         Call MeterItemEnInventario(UserIndex, Obj)
                         
                 Next LooPC
             
             End If
             
      End With
      
End Sub

Public Sub IconoNpcQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
        
    Dim Map   As Integer

    Dim LooPC As Integer
         
    With UserList(UserIndex)
               
        If QuestList(Quest).NumNpcDD > 0 Then
                   
            Map = QuestList(Quest).NpcDD
                   
            For LooPC = 1 To NumNPCs
                   
                If Npclist(LooPC).NPCtype = eNPCType.Misiones Then
                       
                    If Npclist(LooPC).pos.Map = Map Then
                           
                        If .Quest.Icono = 0 Then
                            Call SendData(ToIndex, UserIndex, 0, "XI" & Npclist(LooPC).char.CharIndex & "," & 0)
                        ElseIf .Quest.Icono = 1 Then
                            Call SendData(ToIndex, UserIndex, 0, "XI" & Npclist(LooPC).char.CharIndex & "," & 1)

                        End If
                       
                    End If
                       
                End If
                    
            Next LooPC
                       
        End If
                   
        If QuestList(Quest).NumDescubre > 0 Then
                       
            Map = QuestList(Quest).DescubrePalabra.Mapa
                       
            For LooPC = 1 To NumNPCs
                       
                If Npclist(LooPC).NPCtype = eNPCType.Misiones Then
                              
                    If Npclist(LooPC).pos.Map = Map Then
                              
                        If .Quest.Icono = 0 Then
                            Call SendData(ToIndex, UserIndex, 0, "XI" & Npclist(LooPC).char.CharIndex & "," & 0)
                        ElseIf .Quest.Icono = 1 Then
                            Call SendData(ToIndex, UserIndex, 0, "XI" & Npclist(LooPC).char.CharIndex & "," & 1)

                        End If
                              
                    End If
                              
                End If
                       
            Next LooPC
                   
        End If
                   
        If QuestList(Quest).NumObjsNpc > 0 Then
                       
            Map = QuestList(Quest).MapaObjsNpc
                       
            For LooPC = 1 To NumNPCs
                           
                If Npclist(LooPC).NPCtype = eNPCType.Misiones Then
                               
                    If Npclist(LooPC).pos.Map = Map Then
                                   
                        If .Quest.Icono = 0 Then
                            Call SendData(ToIndex, UserIndex, 0, "XI" & Npclist(LooPC).char.CharIndex & "," & 0)
                        ElseIf .Quest.Icono = 1 Then
                            Call SendData(ToIndex, UserIndex, 0, "XI" & Npclist(LooPC).char.CharIndex & "," & 1)

                        End If
                                   
                    End If
                               
                End If
                       
            Next LooPC
                       
        End If
                   
        If QuestList(Quest).NumHablarNpc > 0 Then
                      
            Map = QuestList(Quest).MapaHablaNpc
                       
            For LooPC = 1 To NumNPCs
                              
                If Npclist(LooPC).NPCtype = eNPCType.Misiones Then
                                  
                    If Npclist(LooPC).pos.Map = Map Then
                                      
                        If .Quest.Icono = 0 Then
                            Call SendData(ToIndex, UserIndex, 0, "XI" & Npclist(LooPC).char.CharIndex & "," & 0)
                        ElseIf .Quest.Icono = 1 Then
                            Call SendData(ToIndex, UserIndex, 0, "XI" & Npclist(LooPC).char.CharIndex & "," & 1)

                        End If
                                      
                    End If
                                  
                End If
                              
            Next LooPC

        End If
               
    End With
         
End Sub

Public Sub CambiaDescQuest(ByVal UserIndex As Integer, ByVal Quest As Integer, ByVal TempCharIndex As Integer)
        
        Dim LooPC As Integer
        
        With UserList(UserIndex)
            
            If .Quest.Start <> 1 Then Exit Sub
            
            If .Quest.ValidNpcDD = 1 Then
               If Npclist(TempCharIndex).NPCtype <> eNPCType.Misiones Then
                   Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).char.CharIndex _
                                                                  & FONTTYPE_INFO)
                    Exit Sub
               End If
               
               If Npclist(TempCharIndex).NPCtype = eNPCType.Misiones Then
                  If Npclist(TempCharIndex).pos.Map = QuestList(Quest).NpcDD Then
                     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & QuestDesc.DobleClick & "°" & Npclist(TempCharIndex).char.CharIndex _
                                                                  & FONTTYPE_INFO)
                  Else
                     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).char.CharIndex _
                                                                  & FONTTYPE_INFO)
                  End If
              End If
            
            ElseIf .Quest.ValidNpcDescubre = 1 Then
               If Npclist(TempCharIndex).NPCtype <> eNPCType.Misiones Then
                   Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).char.CharIndex _
                                                                  & FONTTYPE_INFO)
                    Exit Sub
              End If
              
              If Npclist(TempCharIndex).NPCtype = eNPCType.Misiones Then
                  If Npclist(TempCharIndex).pos.Map = QuestList(Quest).DescubrePalabra.Mapa Then
                     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & QuestDesc.Descubridor & "°" & Npclist(TempCharIndex).char.CharIndex _
                                                                  & FONTTYPE_INFO)
                  Else
                     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).char.CharIndex _
                                                                  & FONTTYPE_INFO)
                  End If
              End If
              
            ElseIf .Quest.NumObjNpc > 0 Then
              If Npclist(TempCharIndex).NPCtype <> eNPCType.Misiones Then
                   Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).char.CharIndex _
                                                                  & FONTTYPE_INFO)
                    Exit Sub
              End If
                 
              If Npclist(TempCharIndex).NPCtype = eNPCType.Misiones Then
                  If Npclist(TempCharIndex).pos.Map = QuestList(Quest).MapaObjsNpc Then
                     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & QuestDesc.DarObjNpc & "°" & Npclist(TempCharIndex).char.CharIndex _
                                                                  & FONTTYPE_INFO)
                  Else
                     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).char.CharIndex _
                                                                  & FONTTYPE_INFO)
                  End If
              End If
              
             ElseIf .Quest.ValidHablarNpc > 0 Then
                If Npclist(TempCharIndex).NPCtype <> eNPCType.Misiones Then
                   Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).char.CharIndex _
                                                                  & FONTTYPE_INFO)
                    Exit Sub
              End If
                 
              If Npclist(TempCharIndex).NPCtype = eNPCType.Misiones Then
                  If Npclist(TempCharIndex).pos.Map = QuestList(Quest).MapaHablaNpc Then
                     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & QuestDesc.Hablador & "°" & Npclist(TempCharIndex).char.CharIndex _
                                                                  & FONTTYPE_INFO)
                  Else
                     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).char.CharIndex _
                                                                  & FONTTYPE_INFO)
                  End If
              End If
                 
            Else
               
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).char.CharIndex _
                                                                  & FONTTYPE_INFO)
            
            End If
        
        End With
        
End Sub

Public Sub EntregaObjNpcQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
       
       Dim LooPC As Integer
       Dim Map As Integer
       
       With UserList(UserIndex)
             
             If .Quest.Start <> 1 Then Exit Sub
             
             If QuestList(Quest).NumObjsNpc > 0 Then
                 
                 Map = .pos.Map
                 
                If QuestList(Quest).MapaObjsNpc = Map Then
                
                    If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                              Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                              Exit Sub
                    End If
                    
                    For LooPC = 1 To QuestList(Quest).NumObjsNpc
                           
                           If Not TieneObjetos(QuestList(Quest).ObjsNpc(LooPC).ObjIndex, QuestList(Quest).ObjsNpc(LooPC).Amount, UserIndex) Then
                                Call SendData(ToIndex, UserIndex, 0, "||No tienes los objetos de la mision en el inventario!!" & FONTTYPE_INFO)
                               Exit Sub
                           End If
                           
                    Next LooPC
                    
                    For LooPC = 1 To QuestList(Quest).NumObjsNpc
                         Call QuitarObjetos(QuestList(Quest).ObjsNpc(LooPC).ObjIndex, QuestList(Quest).ObjsNpc(LooPC).Amount, UserIndex)
                    Next LooPC
                    
                    .Quest.DarObjNpcEntrega = 1
                    Call ActualizaQuest(UserIndex, Quest)
                    
                End If
                 
             End If
             
             
       End With
       
End Sub

Public Sub FinalizaHablarQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
     
     Dim c As Byte
      
      With UserList(UserIndex)
           
           If QuestList(Quest).NumHablarNpc > 0 Then
                .Quest.UserHablaNpc = 1
                c = c + 1
           End If
           
           If c > 0 Then
            Call ActualizaQuest(UserIndex, Quest)
         End If
           
      End With
      
End Sub

Public Sub EnviaVentanaHablarQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
    
    Dim LooPC As Integer
    Dim Datos As String
                 
       With UserList(UserIndex)
       
        If QuestList(Quest).NumHablarNpc > 0 Then
            
            Datos = QuestList(Quest).NumMsjHablar & ", "
            
            For LooPC = 1 To QuestList(Quest).NumMsjHablar
                   Datos = Datos & QuestList(Quest).MsjHablar(LooPC) & ", "
            Next LooPC
            
            Datos = Left$(Datos, Len(Datos) - 2)
            
            Call SendData(ToIndex, UserIndex, 0, "XV" & Datos)
            
        End If
       
       End With
                 
End Sub

Public Sub ResetQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)

    Dim LooPC As Integer
       
    With UserList(UserIndex)
            
        If QuestList(Quest).NumNpc > 0 Then

            For LooPC = 1 To QuestList(Quest).NumNpc
                
                .Quest.MataNpc(LooPC) = 0
                       
            Next LooPC
                
            .Quest.NumNpc = 0
                
        End If
            
        If QuestList(Quest).NumObjs > 0 Then
                
            For LooPC = 1 To QuestList(Quest).NumObjs
                
                .Quest.BuscaObj(LooPC) = 0
                Call QuitarObjetos(QuestList(Quest).BuscaObj(LooPC).ObjIndex, QuestList(Quest).BuscaObj(LooPC).Amount, UserIndex)
                
            Next LooPC
                
            .Quest.NumObj = 0
                
        End If
            
        If QuestList(Quest).NumMapas > 0 Then
                
            For LooPC = 1 To QuestList(Quest).NumMapas
                .Quest.Mapa(LooPC) = 0
            Next LooPC
                
            .Quest.NumMap = 0
                
        End If
            
        If QuestList(Quest).NumNpcDD > 0 Then
            .Quest.ValidNpcDD = 0
            .Quest.MapaNpcDD = 0
            .Quest.Icono = 0
            Call IconoNpcQuest(UserIndex, Quest)
        End If
        
        If QuestList(Quest).NumDescubre > 0 Then
            .Quest.ValidNpcDescubre = 0
            .Quest.PreguntaDescubre = 0
            .Quest.Icono = 0
            Call IconoNpcQuest(UserIndex, Quest)
        End If
        
        If QuestList(Quest).NumObjsNpc > 0 Then
            .Quest.NumObjNpc = 0
            For LooPC = 1 To QuestList(Quest).NumObjsNpc
                 .Quest.DarObjNpc(LooPC) = 0
            Next LooPC
            .Quest.DarObjNpcEntrega = 0
            .Quest.Icono = 0
            Call IconoNpcQuest(UserIndex, Quest)
        End If
        
        If QuestList(Quest).NumHablarNpc Then
            .Quest.ValidHablarNpc = 0
            .Quest.UserHablaNpc = 0
            .Quest.Icono = 0
            Call IconoNpcQuest(UserIndex, Quest)
        End If
        
        If QuestList(Quest).NumUser > 0 Then
           .Quest.ValidMatarUser = 0
           .Quest.UserMatados = 0
        End If
            
    End With
       
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
    Dim n As Integer
        
    With UserList(UserIndex)

             For LooPC = 1 To NumQuests
                    
                    If .Quest.UserQuest(LooPC) = 1 Then
                        n = n + 1
                    End If
                    
              Next
        
             If NumQuests = n Then
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
