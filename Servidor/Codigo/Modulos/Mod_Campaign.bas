Attribute VB_Name = "Mod_Campaign"

Option Explicit

'///////////////////////////////////////////////////////////////////////////////////////////////////////////
'Esto va antes del Type user (Modulo Declaraciones)

Public Type UserCampaign
    Current As Integer  'Campa�a actual
    CurrentStage As Integer       'Stage actual
    WaitingReward As Boolean        'Esperando la recompensa
    NpcIndexGiver As Integer        'Index del NpcStage
    Timing As Long    'Tiempo de la campa�a
    StageTiming As Long         'Tiempo de la etapa
    AmountTargetNpc As Integer          'Cantidad de Npc matados
    AmountTargetUser As Integer           'Cantidad de usuarios matados
    AmountTargetObj As Integer          'Cantidad de objetos obtenidos
    HaveToFoundPlace As Boolean           'Si tiene que encontrar un lugar
    CampaignsHasDone As String            'Las campa�as que hizo
    CanDoThisOne As Integer       'N�mero de campa�a que puede hacer
    OffererNpcIndex As Integer          'Index del Npc ofrecedor
    CanViewCampaignList As Boolean              'Si le ofrecieron ver la lista de misiones
End Type

'///////////////////////////////////////////////////////////////////////////////////////////////////////////
'De ac� hasta el Enum CampaignFailReason va debajo del Type MapInfo (M�dulo Declaraciones)

Public Type CampaignLocation
    Trigger As Byte     'N�mero de trigger de la ubicaci�n
    PosX As Byte  'Valor X de la ubicaci�n
    PosY As Byte  'Valor Y de la ubicaci�n
    Map As Byte    'Mapa de la ubicaci�n
End Type

Public Type CampaignUser
    MinLvl As Integer    'Nivel m�nimo
    MaxLvl As Integer    'Nivel m�ximo
    Amount As Integer    'Cantidad de Usuarios
    Alignment As Byte       'Alineamiento
    Faction As Byte     'Facci�n del usuario
    Rank As Byte  'Rango del usuario
End Type

Public Type CampaignNpc
    Index As Integer    'Indice del Npc
    Amount As Integer    'Cantidad
    Action As Integer    '1)Matar 2)Hablar 3)Entregar 4)Encontrar 5)Decir keyword
End Type

Public Type CampaignDialogs
    FirstTalk As String     'Conversaci�n previa
    SecondTalk As String      'Conversaci�n inicial
    ThirdTalk As String     'Conversaci�n de paso
    FourthTalk As String      'Conversaci�n final
    FifthTalk As String     'Conversaci�n de recompensa
End Type

Public Type StageProperties
    Description As String       'Descripci�n de la etapa
    Dialog As CampaignDialogs    'Conversaciones con el StageNpc
    Timing As Integer    'Tiempo requerido para finalizar
End Type

Public Type StageObjetives
    TargetNpc As CampaignNpc    'Informaci�n del npc
    TargetObj As Obj        'Informaci�n del objeto
    TargetUser As CampaignUser    'Informaci�n del TargetUser
    TargetLocation As CampaignLocation    'Informaci�n del TargetLocation
    KeyWord As String   'Palabra clave
End Type

Public Type CampaignProperties
    Name As String    'Nombre principal
    Description As String       'Descripci�n principal
    Dialog As CampaignDialogs    'Conversaciones con el CampaignNpc
    NumStages As Integer    'Cantidad de etapas
    Redoable As Boolean   'Se puede repetir
    Timing As Integer    'Tiempo requerido para finalizar
    FinishWhenDie As Boolean        'Se termina si el usuario muere
End Type

Public Type CampaignRequires
    MinLvl As Integer    'Nivel m�nimo requerido
    MaxLvl As Integer    'Nivel m�ximo permitido
    PreviousCampaign As Integer           'N�mero de campa�a requerida
    Class As Integer    'Clase requerida
    Race As Integer    'Raza requerida
    Genre As Integer    'G�nero requerido
    Alignment As Byte       'Alineamiento
    Faction As Byte     'Facci�n del usuario
    Rank As Byte  'Rango del usuario
End Type

Public Type CampaignRewards
    RewardExp As Long       'Experiencia entregada
    RewardGold As Long        'Oro entregado
    NumRewardObj As Integer       'Cantidad de objetos entregados
    RewardObj() As Obj          'Informaci�n del objeto entregado
End Type

Public Type CampaignStage
    Properties As StageProperties    'Propiedades de la etapa
    Objetives As StageObjetives    'Objetivos a cumplir
    NumGivenObj As Integer      'Cantidad de objetos entregados al iniciar
    GivenObj() As Obj         'Informaci�n del objeto entregado al iniciar
End Type

Public Type Campaign
    Properties As CampaignProperties    'Propiedades de la campa�a
    Requirements As CampaignRequires    'Requerimientos
    Rewards As CampaignRewards    'Recompensas
    Stages() As CampaignStage    'Declaraci�n de uso
End Type

Public Enum CampaignFailReason
    Leave = 1
    Reject = 2
    death = 3
    Timing = 4
End Enum

'///////////////////////////////////////////////////////////////////////////////////////////////////////////
'Esto va debajo de Public distanceToCities() As HomeDistance (M�dulo Declaraciones)

Public Campaign() As Campaign       'Vector principal
Public NumCampaigns As Integer          'N�mero de campa�as cargadas

