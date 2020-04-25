Attribute VB_Name = "Mod_Quest"
Option Explicit

Private NumQuests As Integer

Public Type tQuestList
    Nombre As String
    Descripcion As String
    MinNivel As Byte
    MaxNivel As Byte
End Type

Public QuestList() As tQuestList

Public Sub Load_Quest()
    Dim Quest As Integer
    
    Dim Leer As New clsIniManager
    
    Call Leer.Initialize(DatPath & "Quest.dat")
    
    NumQuests = Leer.GetValue("INIT", "NumQuests")
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumQuests
    frmCargando.cargar.value = 0

    ReDim Preserve QuestList(1 To NumQuests) As tQuestList
    
    For Quest = 1 To NumQuests
       
       QuestList(Quest).Nombre = Leer.GetValue("Quest" & Quest, "Nombre")
       QuestList(Quest).Descripcion = Leer.GetValue("Quest" & Quest, "Descripcion")
       
    Next Quest
    
End Sub
