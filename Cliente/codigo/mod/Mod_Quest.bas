Attribute VB_Name = "Mod_Quest"
Option Explicit

Public NumQuests As Long

Public Type tRecompensaObjeto
     ObjIndex As String
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

Public Type tInfoUser
     UserQuest(1 To 1000) As Integer
End Type

Public Type tQuest
     NumQuests As Integer
     InfoUser As tInfoUser
End Type

Public Quest As tQuest

Public Sub Load_Quest()
    
    Dim Quest As Integer

    Dim LooPC As Integer

    Dim Datos As String
    
    Dim Data() As Byte
    
    Dim handle As Integer
    
    Dim arch As String
    
   If Not Get_File_Data(DirRecursos, "QUEST.DAT", Data, INIT_RESOURCE_FILE) Then Exit Sub
    
    arch = DirRecursos & "Quest.dat"
    
    handle = FreeFile
    Open arch For Binary Access Write As handle
    Put handle, , Data
    Close handle
    
    Dim Leer As New clsIniManager
    
    Call Leer.Initialize(arch)
    
    NumQuests = Leer.GetValue("INIT", "NumQuests")

    ReDim Preserve QuestList(1 To NumQuests) As tQuestList
    
    For Quest = 1 To NumQuests
       
        QuestList(Quest).Nombre = Leer.GetValue("Quest" & Quest, "Nombre")
        QuestList(Quest).Descripcion = Leer.GetValue("Quest" & Quest, "Descripcion")
        QuestList(Quest).Rehacer = Val(Leer.GetValue("Quest" & Quest, "Rehacer"))
        QuestList(Quest).MinNivel = Val(Leer.GetValue("Quest" & Quest, "MinNivel"))
        QuestList(Quest).MaxNivel = Val(Leer.GetValue("Quest" & Quest, "MaxNivel"))
        QuestList(Quest).RecompensaOro = Val(Leer.GetValue("Quest" & Quest, "RecompensaOro"))
        QuestList(Quest).RecompensaExp = Val(Leer.GetValue("Quest" & Quest, "RecompensaExp"))
        QuestList(Quest).RecompensaItem = Val(Leer.GetValue("Quest" & Quest, "RecompensaItem"))
       
        If QuestList(Quest).RecompensaItem > 0 Then

            For LooPC = 1 To MAX_INVENTORY_SLOTS
             
                Datos = Leer.GetValue("Quest" & Quest, "RecompensaItem" & LooPC)
             
                QuestList(Quest).RecompensaObjeto(LooPC).ObjIndex = Val(ReadField(1, Datos, 45))
                QuestList(Quest).RecompensaObjeto(LooPC).Amount = Val(ReadField(2, Datos, 45))
             
            Next LooPC

        End If
        
        QuestList(Quest).HablarNpc = Val(Leer.GetValue("Quest" & Quest, "HablarNPC"))
        
        If QuestList(Quest).HablarNpc > 0 Then
            
            For LooPC = 1 To MAX_INVENTORY_SLOTS
                   QuestList(Quest).HablaNpc(LooPC) = Val(Leer.GetValue("Quest" & Quest, "HablarNPC" & LooPC))
            Next LooPC
            
        End If
        
        QuestList(Quest).NUMCLASES = Val(Leer.GetValue("Quest" & Quest, "Clases"))
        
        If QuestList(Quest).NUMCLASES > 0 Then
            
            For LooPC = 1 To NUMCLASES
            
               QuestList(Quest).Clases(LooPC) = CStr(Leer.GetValue("Quest" & Quest, "Clases" & LooPC))
               
            Next LooPC
            
        End If
        
        QuestList(Quest).NUMRAZAS = Val(Leer.GetValue("Quest" & Quest, "Razas"))
        
        If QuestList(Quest).NUMRAZAS > 0 Then
            
            For LooPC = 1 To NUMRAZAS
                 QuestList(Quest).Razas(LooPC) = CStr(Leer.GetValue("Quest" & Quest, "Razas" & LooPC))
            Next LooPC
            
        End If
        
        QuestList(Quest).Alineacion = Val(Leer.GetValue("Quest" & Quest, "Alineacion"))
        QuestList(Quest).Faccion = Val(Leer.GetValue("Quest" & Quest, "Faccion"))
        QuestList(Quest).RangoFaccion = Val(Leer.GetValue("Quest" & Quest, "RangoFaccion"))
        QuestList(Quest).NumNpc = Val(Leer.GetValue("Quest" & Quest, "MataNPC"))
        
        If QuestList(Quest).NumNpc > 0 Then
            For LooPC = 1 To MAX_INVENTORY_SLOTS
                
                Datos = Leer.GetValue("Quest" & Quest, "MataNPC" & LooPC)
                
                QuestList(Quest).MataNpc(LooPC).NpcIndex = Val(ReadField(1, Datos, 45))
                QuestList(Quest).MataNpc(LooPC).Cantidad = Val(ReadField(2, Datos, 45))
                
            Next LooPC
        End If
        
        QuestList(Quest).NumUser = Val(Leer.GetValue("Quest" & Quest, "MataUSER"))
        
        If QuestList(Quest).NumUser > 0 Then
            QuestList(Quest).MataUser.MinNivel = Val(Leer.GetValue("Quest" & Quest, "MUMinNivel"))
            QuestList(Quest).MataUser.MaxNivel = Val(Leer.GetValue("Quest" & Quest, "MUMaxNivel"))
            QuestList(Quest).MataUser.NUMCLASES = Val(Leer.GetValue("Quest" & Quest, "MUClases"))
            
            If QuestList(Quest).MataUser.NUMCLASES > 0 Then
                   
                   For LooPC = 1 To NUMCLASES
                         QuestList(Quest).MataUser.Clases(LooPC) = CStr(Leer.GetValue("Quest" & Quest, "MUClases" & LooPC))
                   Next LooPC
                   
            End If
            
            QuestList(Quest).MataUser.NUMRAZAS = Val(Leer.GetValue("Quest" & Quest, "MURazas"))
            
            If QuestList(Quest).MataUser.NUMRAZAS > 0 Then
                  
                  For LooPC = 1 To NUMRAZAS
                       QuestList(Quest).MataUser.Razas(LooPC) = CStr(Leer.GetValue("Quest" & Quest, "MURazas" & LooPC))
                  Next LooPC
                  
            End If
            
            QuestList(Quest).MataUser.Alineacion = Val(Leer.GetValue("Quest" & Quest, "MUAlineacion"))
            QuestList(Quest).MataUser.Faccion = Val(Leer.GetValue("Quest" & Quest, "MUFaccion"))
            QuestList(Quest).MataUser.RangoFaccion = Val(Leer.GetValue("Quest" & Quest, "MURangoFaccion"))
            
        End If
        
        QuestList(Quest).NumObjs = Val(Leer.GetValue("Quest" & Quest, "BuscaObjetos"))
        
        If QuestList(Quest).NumObjs > 0 Then
             
             For LooPC = 1 To MAX_INVENTORY_SLOTS
                    
                    Datos = Leer.GetValue("Quest" & Quest, "BuscaObjetos" & LooPC)
                    
                    QuestList(Quest).BuscaObj.ObjIndex = Val(ReadField(1, Datos, 45))
                    QuestList(Quest).BuscaObj.Amount = Val(ReadField(2, Datos, 45))
                    
             Next LooPC
             
        End If
        
        QuestList(Quest).NumObjsNpc = Val(Leer.GetValue("Quest" & Quest, "ObjetoNpc"))
        
        If QuestList(Quest).NumObjsNpc > 0 Then
            
            For LooPC = 1 To MAX_INVENTORY_SLOTS
                   
                   Datos = Leer.GetValue("Quest" & Quest, "ObjetoNpc" & LooPC)
                   
                   QuestList(Quest).ObjsNpc.NpcIndex = Val(ReadField(1, Datos, 45))
                   QuestList(Quest).ObjsNpc.ObjIndex = Val(ReadField(2, Datos, 45))
                   QuestList(Quest).ObjsNpc.Amount = Val(ReadField(3, Datos, 45))
                   
            Next LooPC
            
        End If
        
        QuestList(Quest).NumNpcDD = Val(Leer.GetValue("Quest" & Quest, "NpcDD"))
        
        If QuestList(Quest).NumNpcDD > 0 Then
            
            For LooPC = 1 To MAX_INVENTORY_SLOTS
                   QuestList(Quest).NpcDD(LooPC) = Val(Leer.GetValue("Quest" & Quest, "NpcDD" & LooPC))
            Next LooPC
            
        End If
        
        QuestList(Quest).NumMapas = Val(Leer.GetValue("Quest" & Quest, "EncontrarMapa"))
        
        If QuestList(Quest).NumMapas > 0 Then
            
            For LooPC = 1 To MAX_INVENTORY_SLOTS
                QuestList(Quest).Mapas(LooPC) = Val(Leer.GetValue("Quest" & Quest, "EncontrarMapa" & LooPC))
            Next LooPC
            
        End If
        
        QuestList(Quest).NumDescubre = Val(Leer.GetValue("Quest" & Quest, "DescubrePalabra"))
        
        If QuestList(Quest).NumDescubre > 0 Then
              
              For LooPC = 1 To MAX_INVENTORY_SLOTS
                     
                     Datos = Leer.GetValue("Quest" & Quest, "DescubrePalabra" & LooPC)
                     
                     QuestList(Quest).DescubrePalabra(LooPC).NpcIndex = Val(ReadField(1, Datos, 45))
                     QuestList(Quest).DescubrePalabra(LooPC).Frase = Val(ReadField(2, Datos, 45))
                     
              Next LooPC
              
        End If
       
    Next Quest
  
    
End Sub
