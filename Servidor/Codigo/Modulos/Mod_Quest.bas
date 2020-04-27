Attribute VB_Name = "Mod_Quest"
Option Explicit

Private NumQuests As Integer

Public Type tRecompensaObjeto
     ObjIndex As Integer
     Amount As Integer
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
    HablaNpc() As Integer
    NumClases As Byte
    Clases() As String
    NumRazas As Byte
    Razas() As String
    Alineacion As Byte
    Faccion As Byte
    RangoFaccion As Byte
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
       
        QuestList(Quest).Nombre = Leer.GetValue("Quest" & Quest, "Nombre")
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
        
        QuestList(Quest).NumClases = val(Leer.GetValue("Quest" & Quest, "Clases"))
        
        If QuestList(Quest).NumClases > 0 Then
            
            For LoopC = 1 To NumClases
            
               QuestList(Quest).Clases(LoopC) = CStr(Leer.GetValue("Quest" & Quest, "Clases" & LoopC))
               
            Next LoopC
            
        End If
        
        QuestList(Quest).NumRazas = val(Leer.GetValue("Quest" & Quest, "Razas"))
        
        If QuestList(Quest).NumRazas > 0 Then
            
            For LoopC = 1 To NumRazas
                 QuestList(Quest).Razas(LoopC) = CStr(Leer.GetValue("Quest" & Quest, "Razas" & LoopC))
            Next LoopC
            
        End If
        
        QuestList(Quest).Alineacion = val(Leer.GetValue("Quest" & Quest, "Alineacion"))
        QuestList(Quest).Faccion = val(Leer.GetValue("Quest" & Quest, "Faccion"))
        QuestList(Quest).RangoFaccion = val(Leer.GetValue("Quest" & Quest, "RangoFaccion"))
       
        frmCargando.cargar.value = frmCargando.cargar.value + 1
       
    Next Quest
    
End Sub
