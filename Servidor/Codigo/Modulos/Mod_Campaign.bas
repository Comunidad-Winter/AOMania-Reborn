Attribute VB_Name = "Mod_Campaign"

Option Explicit

'///////////////////////////////////////////////////////////////////////////////////////////////////////////
'Esto va antes del Type user (Modulo Declaraciones)

Public Type UserCampaign
    Current As Integer  'Campaña actual
    CurrentStage As Integer       'Stage actual
    WaitingReward As Boolean        'Esperando la recompensa
    NpcIndexGiver As Integer        'Index del NpcStage
    Timing As Long    'Tiempo de la campaña
    StageTiming As Long         'Tiempo de la etapa
    AmountTargetNpc As Integer          'Cantidad de Npc matados
    AmountTargetUser As Integer           'Cantidad de usuarios matados
    AmountTargetObj As Integer          'Cantidad de objetos obtenidos
    HaveToFoundPlace As Boolean           'Si tiene que encontrar un lugar
    CampaignsHasDone As String            'Las campañas que hizo
    CanDoThisOne As Integer       'Número de campaña que puede hacer
    OffererNpcIndex As Integer          'Index del Npc ofrecedor
    CanViewCampaignList As Boolean              'Si le ofrecieron ver la lista de misiones
End Type

'///////////////////////////////////////////////////////////////////////////////////////////////////////////
'De acá hasta el Enum CampaignFailReason va debajo del Type MapInfo (Módulo Declaraciones)

Public Type CampaignLocation
    Trigger As Byte     'Número de trigger de la ubicación
    PosX As Byte  'Valor X de la ubicación
    PosY As Byte  'Valor Y de la ubicación
    Map As Byte    'Mapa de la ubicación
End Type

Public Type CampaignUser
    MinLvl As Integer    'Nivel mínimo
    MaxLvl As Integer    'Nivel máximo
    Amount As Integer    'Cantidad de Usuarios
    Alignment As Byte       'Alineamiento
    Faction As Byte     'Facción del usuario
    Rank As Byte  'Rango del usuario
End Type

Public Type CampaignNpc
    Index As Integer    'Indice del Npc
    Amount As Integer    'Cantidad
    Action As Integer    '1)Matar 2)Hablar 3)Entregar 4)Encontrar 5)Decir keyword
End Type

Public Type CampaignDialogs
    FirstTalk As String     'Conversación previa
    SecondTalk As String      'Conversación inicial
    ThirdTalk As String     'Conversación de paso
    FourthTalk As String      'Conversación final
    FifthTalk As String     'Conversación de recompensa
End Type

Public Type StageProperties
    Description As String       'Descripción de la etapa
    Dialog As CampaignDialogs    'Conversaciones con el StageNpc
    Timing As Integer    'Tiempo requerido para finalizar
End Type

Public Type StageObjetives
    TargetNpc As CampaignNpc    'Información del npc
    TargetObj As Obj        'Información del objeto
    TargetUser As CampaignUser    'Información del TargetUser
    TargetLocation As CampaignLocation    'Información del TargetLocation
    KeyWord As String   'Palabra clave
End Type

Public Type CampaignProperties
    Name As String    'Nombre principal
    Description As String       'Descripción principal
    Dialog As CampaignDialogs    'Conversaciones con el CampaignNpc
    NumStages As Integer    'Cantidad de etapas
    Redoable As Boolean   'Se puede repetir
    Timing As Integer    'Tiempo requerido para finalizar
    FinishWhenDie As Boolean        'Se termina si el usuario muere
End Type

Public Type CampaignRequires
    MinLvl As Integer    'Nivel mínimo requerido
    MaxLvl As Integer    'Nivel máximo permitido
    PreviousCampaign As Integer           'Número de campaña requerida
    Class As Integer    'Clase requerida
    Race As Integer    'Raza requerida
    Genre As Integer    'Género requerido
    Alignment As Byte       'Alineamiento
    Faction As Byte     'Facción del usuario
    Rank As Byte  'Rango del usuario
End Type

Public Type CampaignRewards
    RewardExp As Long       'Experiencia entregada
    RewardGold As Long        'Oro entregado
    NumRewardObj As Integer       'Cantidad de objetos entregados
    RewardObj() As Obj          'Información del objeto entregado
End Type

Public Type CampaignStage
    Properties As StageProperties    'Propiedades de la etapa
    Objetives As StageObjetives    'Objetivos a cumplir
    NumGivenObj As Integer      'Cantidad de objetos entregados al iniciar
    GivenObj() As Obj         'Información del objeto entregado al iniciar
End Type

Public Type Campaign
    Properties As CampaignProperties    'Propiedades de la campaña
    Requirements As CampaignRequires    'Requerimientos
    Rewards As CampaignRewards    'Recompensas
    Stages() As CampaignStage    'Declaración de uso
End Type

Public Enum CampaignFailReason
    Leave = 1
    Reject = 2
    death = 3
    Timing = 4
End Enum

'///////////////////////////////////////////////////////////////////////////////////////////////////////////
'Esto va debajo de Public distanceToCities() As HomeDistance (Módulo Declaraciones)

Public Campaign() As Campaign       'Vector principal
Public NumCampaigns As Integer          'Número de campañas cargadas

