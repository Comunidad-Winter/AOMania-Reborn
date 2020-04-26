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
       
        frmCargando.cargar.value = frmCargando.cargar.value + 1
       
    Next Quest
    
End Sub
