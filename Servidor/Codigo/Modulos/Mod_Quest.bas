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
    Clases() As String
    NUMRAZAS As Byte
    Razas() As String
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
    HablaNpc() As Integer
    NUMCLASES As Byte
    Clases() As String
    NUMRAZAS As Byte
    Razas() As String
    Alineacion As Byte
    Faccion As Byte
    RangoFaccion As Byte
    NumNpc As Byte
    MataNpc() As tMataNpc
    NumUser As Integer
    MataUser As tMataUser
    NumObjs As Byte
    BuscaObj As tBuscaObj
    NumObjsNpc As Byte
    ObjsNpc As tObjsNpc
    NumNpcDD As Byte
    NpcDD() As Integer
    NumMapas As Integer
    Mapas() As Integer
    NumDescubre As Integer
    DescubrePalabra() As tDescubrePalabra
End Type

Public QuestList() As tQuestList

Public Sub Load_Quest()

    Dim Quest As Integer

    Dim LoopC As Integer

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

            For LoopC = 1 To MAX_INVENTORY_SLOTS
             
                Datos = Leer.GetValue("Quest" & Quest, "RecompensaItem" & LoopC)
             
                QuestList(Quest).RecompensaObjeto(LoopC).ObjIndex = val(ReadField(1, Datos, 45))
                QuestList(Quest).RecompensaObjeto(LoopC).Amount = val(ReadField(2, Datos, 45))
             
            Next LoopC

        End If
        
        QuestList(Quest).HablarNpc = val(Leer.GetValue("Quest" & Quest, "HablarNPC"))
        
        If QuestList(Quest).HablarNpc > 0 Then
            
            For LoopC = 1 To MAX_INVENTORY_SLOTS
                   QuestList(Quest).HablaNpc(LoopC) = val(Leer.GetValue("Quest" & Quest, "HablarNPC" & LoopC))
            Next LoopC
            
        End If
        
        QuestList(Quest).NUMCLASES = val(Leer.GetValue("Quest" & Quest, "Clases"))
        
        If QuestList(Quest).NUMCLASES > 0 Then
            
            For LoopC = 1 To NUMCLASES
            
               QuestList(Quest).Clases(LoopC) = CStr(Leer.GetValue("Quest" & Quest, "Clases" & LoopC))
               
            Next LoopC
            
        End If
        
        QuestList(Quest).NUMRAZAS = val(Leer.GetValue("Quest" & Quest, "Razas"))
        
        If QuestList(Quest).NUMRAZAS > 0 Then
            
            For LoopC = 1 To NUMRAZAS
                 QuestList(Quest).Razas(LoopC) = CStr(Leer.GetValue("Quest" & Quest, "Razas" & LoopC))
            Next LoopC
            
        End If
        
        QuestList(Quest).Alineacion = val(Leer.GetValue("Quest" & Quest, "Alineacion"))
        QuestList(Quest).Faccion = val(Leer.GetValue("Quest" & Quest, "Faccion"))
        QuestList(Quest).RangoFaccion = val(Leer.GetValue("Quest" & Quest, "RangoFaccion"))
        QuestList(Quest).NumNpc = val(Leer.GetValue("Quest" & Quest, "MataNPC"))
        
        If QuestList(Quest).NumNpc > 0 Then
            For LoopC = 1 To MAX_INVENTORY_SLOTS
                
                Datos = Leer.GetValue("Quest" & Quest, "MataNPC" & LoopC)
                
                QuestList(Quest).MataNpc(LoopC).NpcIndex = val(ReadField(1, Datos, 45))
                QuestList(Quest).MataNpc(LoopC).Cantidad = val(ReadField(2, Datos, 45))
                
            Next LoopC
        End If
        
        QuestList(Quest).NumUser = val(Leer.GetValue("Quest" & Quest, "MataUSER"))
        
        If QuestList(Quest).NumUser > 0 Then
            QuestList(Quest).MataUser.MinNivel = val(Leer.GetValue("Quest" & Quest, "MUMinNivel"))
            QuestList(Quest).MataUser.MaxNivel = val(Leer.GetValue("Quest" & Quest, "MUMaxNivel"))
            QuestList(Quest).MataUser.NUMCLASES = val(Leer.GetValue("Quest" & Quest, "MUClases"))
            
            If QuestList(Quest).MataUser.NUMCLASES > 0 Then
                   
                   For LoopC = 1 To NUMCLASES
                         QuestList(Quest).MataUser.Clases(LoopC) = CStr(Leer.GetValue("Quest" & Quest, "MUClases" & LoopC))
                   Next LoopC
                   
            End If
            
            QuestList(Quest).MataUser.NUMRAZAS = val(Leer.GetValue("Quest" & Quest, "MURazas"))
            
            If QuestList(Quest).MataUser.NUMRAZAS > 0 Then
                  
                  For LoopC = 1 To NUMRAZAS
                       QuestList(Quest).MataUser.Razas(LoopC) = CStr(Leer.GetValue("Quest" & Quest, "MURazas" & LoopC))
                  Next LoopC
                  
            End If
            
            QuestList(Quest).MataUser.Alineacion = val(Leer.GetValue("Quest" & Quest, "MUAlineacion"))
            QuestList(Quest).MataUser.Faccion = val(Leer.GetValue("Quest" & Quest, "MUFaccion"))
            QuestList(Quest).MataUser.RangoFaccion = val(Leer.GetValue("Quest" & Quest, "MURangoFaccion"))
            
        End If
        
        QuestList(Quest).NumObjs = val(Leer.GetValue("Quest" & Quest, "BuscaObjetos"))
        
        If QuestList(Quest).NumObjs > 0 Then
             
             For LoopC = 1 To MAX_INVENTORY_SLOTS
                    
                    Datos = Leer.GetValue("Quest" & Quest, "BuscaObjetos" & LoopC)
                    
                    QuestList(Quest).BuscaObj.ObjIndex = val(ReadField(1, Datos, 45))
                    QuestList(Quest).BuscaObj.Amount = val(ReadField(2, Datos, 45))
                    
             Next LoopC
             
        End If
        
        QuestList(Quest).NumObjsNpc = val(Leer.GetValue("Quest" & Quest, "ObjetoNpc"))
        
        If QuestList(Quest).NumObjsNpc > 0 Then
            
            For LoopC = 1 To MAX_INVENTORY_SLOTS
                   
                   Datos = Leer.GetValue("Quest" & Quest, "ObjetoNpc" & LoopC)
                   
                   QuestList(Quest).ObjsNpc.NpcIndex = val(ReadField(1, Datos, 45))
                   QuestList(Quest).ObjsNpc.ObjIndex = val(ReadField(2, Datos, 45))
                   QuestList(Quest).ObjsNpc.Amount = val(ReadField(3, Datos, 45))
                   
            Next LoopC
            
        End If
        
        QuestList(Quest).NumNpcDD = val(Leer.GetValue("Quest" & Quest, "NpcDD"))
        
        If QuestList(Quest).NumNpcDD > 0 Then
            
            For LoopC = 1 To MAX_INVENTORY_SLOTS
                   QuestList(Quest).NpcDD(LoopC) = val(Leer.GetValue("Quest" & Quest, "NpcDD" & LoopC))
            Next LoopC
            
        End If
        
        QuestList(Quest).NumMapas = val(Leer.GetValue("Quest" & Quest, "EncontrarMapa"))
        
        If QuestList(Quest).NumMapas > 0 Then
            
            For LoopC = 1 To MAX_INVENTORY_SLOTS
                QuestList(Quest).Mapas(LoopC) = val(Leer.GetValue("Quest" & Quest, "EncontrarMapa" & LoopC))
            Next LoopC
            
        End If
        
        QuestList(Quest).NumDescubre = val(Leer.GetValue("Quest" & Quest, "DescubrePalabra"))
        
        If QuestList(Quest).NumDescubre > 0 Then
              
              For LoopC = 1 To MAX_INVENTORY_SLOTS
                     
                     Datos = Leer.GetValue("Quest" & Quest, "DescubrePalabra" & LoopC)
                     
                     QuestList(Quest).DescubrePalabra(LoopC).NpcIndex = val(ReadField(1, Datos, 45))
                     QuestList(Quest).DescubrePalabra(LoopC).Frase = val(ReadField(2, Datos, 45))
                     
              Next LoopC
              
        End If
        
        frmCargando.cargar.value = frmCargando.cargar.value + 1
       
    Next Quest
    
End Sub

Public Sub IniciarVentanaQuest(ByVal UserIndex As Integer)
     
     
     
     
End Sub
