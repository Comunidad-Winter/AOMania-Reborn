Attribute VB_Name = "Mod_Quest"
Option Explicit

Public NumQuests As Long

Public Type tRecompensaObjeto
     ObjIndex As String
     Amount As Integer
End Type

Public Type tQuestList
    Nombre As String
    Descripcion As String
    RecompensaOro As Long
    RecompensaExp As Long
    RecompensaItem As Byte
    RecompensaObjeto(1 To 10) As tRecompensaObjeto
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

Public Type tHablarQuest
     NumMsj As Byte
     Mensaje(1 To 10) As String
     Proceso As Byte
End Type

Public HablarQuest As tHablarQuest

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
        QuestList(Quest).RecompensaOro = Val(Leer.GetValue("Quest" & Quest, "RecompensaOro"))
        QuestList(Quest).RecompensaExp = Val(Leer.GetValue("Quest" & Quest, "RecompensaExp"))
        QuestList(Quest).RecompensaItem = Val(Leer.GetValue("Quest" & Quest, "RecompensaItem"))
       
        If QuestList(Quest).RecompensaItem > 0 Then

            For LooPC = 1 To QuestList(Quest).RecompensaItem
             
                Datos = Leer.GetValue("Quest" & Quest, "RecompensaItem" & LooPC)
             
                QuestList(Quest).RecompensaObjeto(LooPC).ObjIndex = Val(ReadField(1, Datos, 45))
                QuestList(Quest).RecompensaObjeto(LooPC).Amount = Val(ReadField(2, Datos, 45))
             
            Next LooPC

        End If
       
    Next Quest
  
    
End Sub
