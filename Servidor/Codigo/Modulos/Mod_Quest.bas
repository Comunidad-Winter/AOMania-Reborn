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
    NumClases As Byte
    Clases(1 To NumClases) As String
    NumRazas As Byte
    Razas(1 To NumRazas) As String
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
    Nombre As String
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
    NumClases As Byte
    Clases(1 To NumClases) As String
    NumRazas As Byte
    Razas(1 To NumRazas) As String
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
       
        QuestList(Quest).Nombre = Leer.GetValue("Quest" & Quest, "Nombre")
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
        
        QuestList(Quest).NumClases = val(Leer.GetValue("Quest" & Quest, "Clases"))
        
        If QuestList(Quest).NumClases > 0 Then
            
            For LooPC = 1 To NumClases
            
               QuestList(Quest).Clases(LooPC) = CStr(Leer.GetValue("Quest" & Quest, "Clases" & LooPC))
               
            Next LooPC
            
        End If
        
        QuestList(Quest).NumRazas = val(Leer.GetValue("Quest" & Quest, "Razas"))
        
        If QuestList(Quest).NumRazas > 0 Then
            
            For LooPC = 1 To NumRazas
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
            QuestList(Quest).MataUser.NumClases = val(Leer.GetValue("Quest" & Quest, "MUClases"))
            
            If QuestList(Quest).MataUser.NumClases > 0 Then
                   
                   For LooPC = 1 To NumClases
                         QuestList(Quest).MataUser.Clases(LooPC) = CStr(Leer.GetValue("Quest" & Quest, "MUClases" & LooPC))
                   Next LooPC
                   
            End If
            
            QuestList(Quest).MataUser.NumRazas = val(Leer.GetValue("Quest" & Quest, "MURazas"))
            
            If QuestList(Quest).MataUser.NumRazas > 0 Then
                  
                  For LooPC = 1 To NumRazas
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

Public Sub IniciarMisionQuest(ByVal UserIndex As Integer, ByVal Quest As Integer)
       
       Dim LooPC As Integer
       Dim N As Integer
       Dim C As Integer
        
        With UserList(UserIndex)
        
              If .Quest.Start = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Ya tienes una misión iniciada!! Acabala antes de volver a empezar otra." & FONTTYPE_INFO)
                  Exit Sub
              End If
              
              'If .Quest.Quest < Quest Then
              '    Call SendData(ToIndex, UserIndex, 0, "||No te adelantes!! Debes hacer una misión antes que esta!!" & FONTTYPE_INFO)
              '    Exit Sub
              'End If
              
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
              
              If QuestList(Quest).NumClases > 0 Then
                  
                  N = QuestList(Quest).NumClases
                  
                  For LooPC = 1 To N
                         Debug.Print UCase$(QuestList(Quest).Clases(LooPC))
                         If UCase$(QuestList(Quest).Clases(LooPC)) = UCase$(.Clase) Then
                             C = C + 1
                         End If
                         
                  Next LooPC
                  
                  If C = 0 Then
                     Call SendData(ToIndex, UserIndex, 0, "||Tu clase no esta permitida para realizar esta misión, puedes pasar a la siguiente!" & FONTTYPE_INFO)
                     .Quest.UserQuest(Quest) = 1
                     .Quest.Quest = Quest
                     Exit Sub
                  End If
                  
                  C = 0
              End If
              
              If QuestList(Quest).NumRazas > 0 Then
                  
                  N = QuestList(Quest).NumRazas
                  
                  For LooPC = 1 To N
                     If UCase$(QuestList(Quest).Razas(LooPC)) = UCase$(.Raza) Then
                         C = C + 1
                     End If
                  Next LooPC
                  
                  If C = 0 Then
                     Call SendData(ToIndex, UserIndex, 0, "||Tu raza no esta permitida para realizar esta misión, puedes pasar a la siguiente!" & FONTTYPE_INFO)
                     .Quest.UserQuest(Quest) = 1
                     .Quest.Quest = Quest
                  End If
                  
                  C = 0
              End If
              
             
        End With
        
End Sub
