Attribute VB_Name = "Declaraciones"
Option Explicit

Public IndexNPC As Integer
Public CvcFunciona As Boolean
Public Const SkillPointInicial As Byte = 10

Public CountNpc As Long
Public CountNpcH As Long

Public ValidMap() As Byte

Public Type LevelSkill

    LevelValue As Integer

End Type

Public LevelSkill() As LevelSkill

Public Const MapaNoValidoNemesis As Integer = 34
Public Const MapaNoValidoTemplario As Integer = 34

Public Type MBarcos

    TiempoRest As Single
    Zona As Integer
    Pasajeros As Integer

End Type

Public Type MZonas

    nombre As String
    Map As Integer
    x As Integer
    Y As Integer

End Type

Public Barcos As MBarcos
Public Const MAX_ZONAS = 100
Public Const MAX_PASAJEROS = 30
Public Const TIEMPO_LLEGADA = -6
Public Zonas(MAX_ZONAS) As MZonas
Public NumZonas As Integer

Public entrarReto As Long
Public entrarPlante As Long
Public entrarReto2v2 As Long

Public lvlGuerra As Long
Public lvlMedusa As Long
Public lvlTorneo As Long
Public lvlDeath As Long

Public Lac_Camina As Long
Public Lac_Pociones As Long
Public Lac_Pegar As Long
Public Lac_Lanzar As Long
Public Lac_Usar As Long
Public Lac_Tirar As Long

Public Type TLac

    LCaminar As New Cls_InterGTC
    LPociones As New Cls_InterGTC
    LPegar As New Cls_InterGTC
    LUsar As New Cls_InterGTC
    LTirar As New Cls_InterGTC
    LLanzar As New Cls_InterGTC

End Type

Public Retos1 As String
Public Retos2 As String
Public Plante1 As String
Public Plante2 As String
Public PrecioQl As Byte
Public ComerciarAc As Boolean
Public KATA As Boolean
Public terminodeat As Boolean
Public bandasqls As Integer
'invocaciones
'Mapa
Public tukiql As Integer
Public Const mapainvo = 96
' posi 1
Public Const mapainvoX1 = 27
Public Const mapainvoY1 = 21
' posi 2
Public Const mapainvoX2 = 21
Public Const mapainvoY2 = 27
'posi 3
Public Const mapainvoX3 = 33
Public Const mapainvoY3 = 27
'posi 4
Public Const mapainvoX4 = 27
Public Const mapainvoY4 = 33

Public YaHayPlante As Boolean
Public denuncias As Boolean
Public keyA As String
Public keyB As String
''
' Modulo de declaraciones. Aca hay de todo.
Public duelosespera As Integer
Public duelosreta As Integer
Public numduelos As Integer
Public Const MAPADUELO As Integer = 154

Public CuentaRegresiva As Long
'2vs2
Public HayPareja As Boolean

Type tEstadisticasDiarias

    Segundos As Double
    MaxUsuarios As Integer
    Promedio As Integer


End Type

Public DayStats As tEstadisticasDiarias

Public aClon As New clsAntiMassClon
Public TrashCollector As New Collection

Public Const MAXSPAWNATTEMPS = 60
Public Const MAXUSERMATADOS = 9000000
Public Const LoopAdEternum = 999
Public Const FXSANGRE = 14

Public Const iFragataFantasmal = 87

Public Enum iMinerales

    HierroCrudo = 192
    PlataCruda = 193
    OroCrudo = 194
    MercurioCrudo = 1583

    LingoteDeHierro = 386
    LingoteDePlata = 387
    LingoteDeOro = 388
    LingoteDeMercurio = 1584

End Enum

Public Type tLlamadaGM

    Usuario As String * 255
    Desc As String * 255

End Type

Public Enum PlayerType

    User = 0
    Consejero = 1
    SemiDios = 2
    Dios = 3
    Admin = 4

End Enum

Public Const LimiteNewbie As Byte = 13

Public Type tCabecera    'Cabecera de los con

    Desc As String * 255
    crc As Long
    MagicWord As Long

End Type

Public MiCabecera As tCabecera

'Barrin 3/10/03
Public Const TIEMPO_INICIOMEDITAR As Byte = 3

Public Const NingunEscudo As Integer = 2
Public Const NingunCasco As Integer = 2
Public Const NingunArma As Integer = 2
Public Const NingunAlas As Integer = 0

Public Const EspadaMataDragonesIndex As Integer = 402
Public Const LAUDMAGICO As Integer = 696

Public Const MAXMASCOTASENTRENADOR As Byte = 7

Public Enum FXIDs

    FXWARP = 42
    FXMEDITARNW = 4
    FXMEDITARAZULNW = 5
    FXMEDITARFUEGUITO = 6
    FXMEDITARFUEGO = 27
    FXMEDITARMEDIANO = 54
    FXMEDITARAZULCITO = 55
    FXMEDITARGRIS = 53
    FXMEDITARFULL = 52

End Enum

Public Const TIEMPO_CARCEL_PIQUETE As Long = 3

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum eTrigger

    Nada = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
    resu = 7

End Enum

''
' constantes para el trigger 6
'
' @see eTrigger
' @param TRIGGER6_PERMITE TRIGGER6_PERMITE
' @param TRIGGER6_PROHIBE TRIGGER6_PROHIBE
' @param TRIGGER6_AUSENTE El trigger no aparece
'
Public Enum eTrigger6

    TRIGGER6_PERMITE = 1
    TRIGGER6_PROHIBE = 2
    TRIGGER6_AUSENTE = 3

End Enum

'TODO : Reemplazar por un enum
Public Const Bosque = "BOSQUE"
Public Const Nieve = "NIEVE"
Public Const Desierto = "DESIERTO"
Public Const Ciudad = "CIUDAD"
Public Const Campo = "CAMPO"
Public Const Dungeon = "DUNGEON"

' <<<<<< Targets >>>>>>
Public Enum TargetType

    uUsuarios = 1
    uNPC = 2
    uUsuariosYnpc = 3
    uTerreno = 4

End Enum

' <<<<<< Acciona sobre >>>>>>
Public Enum TipoHechizo

    uPropiedades = 1
    uEstado = 2
    uMaterializa = 3    'Nose usa
    uInvocacion = 4
    uArea = 5

End Enum

Public Const MAX_MENSAJES_FORO As Byte = 35

Public Const MAXUSERHECHIZOS As Byte = 35

' TODO: Y ESTO ? LO CONOCE GD ?
Public Const EsfuerzoTalarGeneral As Byte = 4
Public Const EsfuerzoTalarLeñador As Byte = 2

Public Const EsfuerzoPescarPescador As Byte = 1
Public Const EsfuerzoPescarGeneral As Byte = 3

Public Const EsfuerzoExcavarMinero As Byte = 2
Public Const EsfuerzoExcavarGeneral As Byte = 5

Public Const FX_TELEPORT_INDEX As Integer = 1

' La utilidad de esto es casi nula, sólo se revisa si fue a la cabeza...
Public Enum PartesCuerpo

    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6

End Enum

Public Const Guardias As Integer = 6

Public Const MAXREP As Long = 6000000
Public Const MaxOro As Long = 999999999
Public Const MAXEXP As Long = 999999999

Public Const MAXATRIBUTOS As Byte = 35
Public Const MINATRIBUTOS As Byte = 6

Public Const LingoteHierro As Integer = 386
Public Const LingotePlata As Integer = 387
Public Const LingoteOro As Integer = 388
Public Const LingoteMercurio As Integer = 1584
Public Const Diamante As Integer = 1274
Public Const GemaMagica As Integer = 1316

Public Const Leña As Integer = 58

Public Const MAXNPCS As Integer = 10000
Public Const MAXCHARS As Integer = 10000

Public Const HACHA_LEÑADOR As Integer = 127
Public Const PIQUETE_MINERO As Integer = 187

Public Const DAGA As Integer = 15

Public Const FOGATA_APAG As Integer = 136
Public Const FOGATA As Integer = 63

Public Const ORO_MINA As Integer = 194
Public Const PLATA_MINA As Integer = 193
Public Const HIERRO_MINA As Integer = 192

Public Const MARTILLO_HERRERO As Integer = 389
Public Const MARTILLO_HERRERO_MAGICO As Integer = 1272
Public Const SERRUCHO_CARPINTERO As Integer = 198

Public Const ObjArboles As Integer = 4

Public Const RED_PESCA As Integer = 543
Public Const CAÑA_PESCA As Integer = 138

Public Const Lana As Integer = 880
Public Const TIJERA As Integer = 881
Public Const AGUJA As Integer = 882
Public Const PielLobo As Integer = 414
Public Const PielLoboPolar As Integer = 139
Public Const PielOsos As Integer = 415
Public Const PielTigre As Integer = 545
Public Const PielOsosPolar As Integer = 416
Public Const PielVaca As Integer = 544
Public Const PielJabali As Integer = 1166

Public Const HOZ_DE_MANO As Integer = 878
Public Const MORTERO As Integer = 879
Public Const Hierba As Integer = 884

Public ContReSpawnNpc As Integer
Public mariano As Integer
Public xao As Integer

Public Enum eNPCType

    Comun = 0
    Revividor = 1
    GuardiaReal = 2
    Entrenador = 3
    Banquero = 4
    armada = 5
    dragon = 6
    Cirujia = 7
    Guardia = 9
    Duelos = 10
    CambiaCabeza = 11
    Teleport = 12
    Guardiascaos = 17
    Casamiento = 18
    Clero = 20
    Abbadon = 21
    Timbero = 23
    Templario = 26
    Tiniebla = 29
    Banda = 33
    Medusa = 38
    OlvidarHechizo = 54
    Misiones = 96
    nQuest = 97
    Canjes = 98
    Creditos = 99

End Enum

Public Const MIN_APUÑALAR As Byte = 10

'********** CONSTANTANTES ***********

''
' Cantidad de skills
Public Const NUMSKILLS As Byte = 26

''
' Cantidad de Atributos
Public Const NUMATRIBUTOS As Byte = 5

''
' Cantidad de Clases
Public Const NUMCLASES As Byte = 14

''
' Cantidad de Razas
Public Const NUMRAZAS As Byte = 10

''
' Valor maximo de cada skill
Public Const MAXSKILLPOINTS As Byte = 100

''
'Direccion
'
' @param NORTH Norte
' @param EAST Este
' @param SOUTH Sur
' @param WEST Oeste
'
Public Enum eHeading

    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4

End Enum

''
' Cantidad maxima de mascotas
Public Const MAXMASCOTAS As Byte = 3

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const vlASALTO As Integer = 100
Public Const vlASESINO As Integer = 1000
Public Const vlCAZADOR As Integer = 5
Public Const vlNoble As Integer = 5
Public Const vlLadron As Integer = 25
Public Const vlProleta As Integer = 2

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const iCuerpoMuerto As Integer = 8
Public Const iCabezaMuerto As Integer = 500
Public Const iCuerpoMuertoCrimi As Integer = 145
Public Const iCabezaMuertoCrimi As Integer = 501

Public Const iORO As Byte = 12
Public Const Pescado As Integer = 1161
Public Const PescadoCofre As Integer = 11

Public Enum PECES_POSIBLES

    PESCADO1 = 1162
    PESCADO2 = 1163
    PESCADO3 = 1164
    PESCADO4 = 1165

End Enum

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Enum eSkill

    Suerte = 1
    Magia = 2
    Robar = 3
    Tacticas = 4
    Armas = 5
    Meditar = 6
    Apuñalar = 7
    Ocultarse = 8
    Supervivencia = 9
    talar = 10
    comerciar = 11
    Defensa = 12
    Resistencia = 13
    Pesca = 14
    Mineria = 15
    Carpinteria = 16
    Herreria = 17
    Liderazgo = 18
    Domar = 19
    Proyectiles = 20
    Wresterling = 21
    Navegacion = 22
    Sastreria = 23
    Recolectar = 24
    Hechiceria = 25
    Herrero = 26

End Enum

Public Const FundirMetal = 88

Public Enum eAtributos

    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    constitucion = 5

End Enum

Public Const AdicionalHPGuerrero As Byte = 2    'HP adicionales cuando sube de nivel
Public Const AdicionalHPCazador As Byte = 1    'HP adicionales cuando sube de nivel

Public Const AumentoSTDef As Byte = 15
Public Const AumentoSTLadron As Byte = AumentoSTDef + 3
Public Const AumentoSTMago As Byte = AumentoSTDef - 1
Public Const AumentoSTLeñador As Byte = AumentoSTDef + 23
Public Const AumentoSTPescador As Byte = AumentoSTDef + 20
Public Const AumentoSTMinero As Byte = AumentoSTDef + 25

'Sonidos
Public Const SND_SWING As Byte = 2
Public Const SND_TALAR As Byte = 13
Public Const SND_PESCAR As Byte = 14
Public Const SND_MINERO As Byte = 15
Public Const SND_WARP As Byte = 3
Public Const SND_PUERTA As Byte = 5
Public Const SND_NIVEL As Byte = 6

Public Const SND_USERMUERTE As Byte = 11
Public Const SND_IMPACTO As Byte = 10
Public Const SND_IMPACTO2 As Byte = 12
Public Const SND_LEÑADOR As Byte = 13
Public Const SND_FOGATA As Byte = 14
Public Const SND_AVE As Byte = 21
Public Const SND_AVE2 As Byte = 22
Public Const SND_AVE3 As Byte = 34
Public Const SND_GRILLO As Byte = 28
Public Const SND_GRILLO2 As Byte = 29
Public Const SND_SACARARMA As Byte = 25
Public Const SND_ESCUDO As Byte = 37
Public Const MARTILLOHERRERO As Byte = 41
Public Const LABUROCARPINTERO As Byte = 42
Public Const SND_POTEAR As Byte = 46
Public Const SND_CHIRP As Byte = 47
Public Const SND_BEBER As Byte = 175

''
' Cantidad maxima de objetos por slot de inventario
Public Const MAX_INVENTORY_OBJS As Integer = 10000

''
' Cantidad de "slots" en el inventario
Public Const MAX_INVENTORY_SLOTS As Byte = 25

''
' Constante para indicar que se esta usando ORO
Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1

' CATEGORIAS PRINCIPALES
Public Enum eOBJType

    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otCONTENEDORES = 7
    otCARTELES = 8
    otLlaves = 9
    otFOROS = 10
    otPociones = 11
    otBebidas = 13
    otLeña = 14
    otFogata = 15
    otESCUDO = 16
    otCASCO = 17
    otherramientas = 18
    otTELEPORT = 19
    otRegalos = 20
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otCheques = 25
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otMANCHAS = 35
    otPARAA = 36
    otAlas = 37
    otHierba = 38
    otOveja = 40
    otPasaje = 41
    otAmuletoDefensa = 45
    otAmuleto = 47
    otVales = 50
    otPack = 51
    otMontura = 60
    otPLATA = 77
    otLibromagico = 999
    otCualquiera = 1000

End Enum

Public Enum eAmuleto
    otMagia = 1
    otFisico = 3
End Enum

'Texto
Public Const FONTTYPE_TALK As String = "~255~255~255~0~0"
Public Const FONTTYPE_TURQ As String = "~5~205~216~0~0"
Public Const fonttype_HABLAR As String = "~6~159~60~0~0"
Public Const FONTTYPE_FIGHT As String = "~206~4~4~0~0"
Public Const FONTTYPE_WARNING As String = "~54~69~245~1~0"
Public Const FONTTYPE_WARNIN As String = "~4~249~90~0~0"
Public Const FONTTYPE_INFO As String = "~65~190~156~0~0"
Public Const FONTTYPE_APU As String = "~77~238~238~1~0"
Public Const FONTTYPE_INFON As String = "~65~190~156~1~0"
Public Const FONTTYPE_INFOBOLD As String = "~65~190~156~1~0"
Public Const FONTTYPE_EJECUCION As String = "~199~32~7~1~0"
Public Const FONTTYPE_PARTY As String = "~67~134~201~1~0"
Public Const FONTTYPE_VENENO As String = "~0~255~0~1~0"
Public Const FONTTYPE_CHEAT As String = "~77~198~36~1~0"
Public Const FONTTYPE_GUILD As String = "~255~255~255~1~0"
Public Const FONTTYPE_TALKMSG As String = "~244~244~244~1~0"
Public Const FONTTYPE_DENUNCIAR As String = "~200~168~147~0~0"
Public Const FONTTYPE_SERVER As String = "~199~200~209~0~0"
Public Const FONTTYPE_GUILDMSG As String = "~0~255~0~0~0"
Public Const FONTTYPE_CONSEJO As String = "~106~181~255~1~0"
Public Const FONTTYPE_RETOS As String = "~77~198~36~0~"
Public Const FONTTYPE_RETOS2V2 As String = "~5~205~216~0~0"
Public Const FONTTYPE_GUERRA As String = "~77~198~36~1~0"
Public Const FONTTYPE_DEATH As String = "~106~181~255~1~0"
Public Const FONTTYPE_PLANTE As String = "~255~128~64~0~"
Public Const FONTTYPE_CONSEJOCAOS As String = "~255~128~54~1~0"
Public Const FONTTYPE_CONSEJOVesA As String = "~106~181~255~1~0"
Public Const FONTTYPE_CONSEJOCAOSVesA As String = "~255~150~0~1~0"
Public Const FONTTYPE_WETAS As String = "~128~128~128~1~0"
Public Const FONTTYPE_CENTINELA As String = "~0~255~0~1~0"
Public Const FONTTYPE_ORO As String = "~255~255~0~1~0"
Public Const FONTTYPE_CYAN As String = "~67~188~188~0~0"
Public Const FONTTYPE_WorldCarga As String = "~0~128~255~1~0"
Public Const FONTTYPE_WorldSave As String = "~255~128~64~1~0"
Public Const FONTTYPE_Motd1 As String = "~235~0~8~0~0"
Public Const FONTTYPE_Motd2 As String = "~255~50~0~0~0"
Public Const FONTTYPE_Motd3 As String = "~0~255~0~1~0"
Public Const FONTTYPE_Motd4 As String = "~255~0~0~1~0"
Public Const FONTTYPE_Motd5 As String = "~255~255~0~1~0"
Public Const FONTTYPE_AMARILLON As String = "~255~255~0~1~0"
Public Const FONTTYPE_BLANCO As String = "~255~255~255~0~0"
Public Const FONTTYPE_BORDO As String = "~128~0~0~0~0"
Public Const FONTTYPE_VERDE As String = "~0~255~0~0~0"
Public Const FONTTYPE_ROJO As String = "~255~0~0~0~0"
Public Const FONTTYPE_AZUL As String = "~0~0~255~0~0"
Public Const FONTTYPE_VIOLETA As String = "~128~0~128~0~0"
Public Const FONTTYPE_AMARILLO As String = "~255~255~0~0~0"
Public Const FONTTYPE_CELESTE As String = "~128~255~255~0~0"
Public Const FONTTYPE_GRIS As String = "~130~130~130~0~0"
Public Const FONTTYPE_BLANCON As String = "~255~255~255~1~0"
Public Const FONTTYPE_BORDON As String = "~128~0~0~1~0"
Public Const FONTTYPE_VERDEN As String = "~0~255~0~1~0"
Public Const FONTTYPE_ROJON As String = "~255~0~0~1~0"
Public Const FONTTYPE_AZULN As String = "~0~0~255~1~0"
Public Const FONTTYPE_VIOLETAN As String = "~128~0~128~1~0"
Public Const FONTTYPE_CELESTEN As String = "~128~255~255~1~0"
Public Const FONTTYPE_GRISN As String = "~130~130~130~1~0"
Public Const FONTTYPE_QUEST As String = "~128~255~0~1~0"


'Estadisticas
Public Const STAT_MAXELV As Byte = 55
Public Const STAT_MAXHP As Integer = 999
Public Const STAT_MAXSTA As Integer = 999
Public Const STAT_MAXMAN As Integer = 3000
Public Const STAT_MAXHIT_UNDER36 As Byte = 99
Public Const STAT_MAXHIT_OVER36 As Integer = 999
Public Const STAT_MAXDEF As Byte = 99

' **************************************************************
' **************************************************************
' ************************ TIPOS *******************************
' **************************************************************
' **************************************************************

Public Type NpcDrop

    NumDrop As Integer
    DropIndex(1 To 20) As Integer
    Porcentaje(1 To 20) As Integer
    Amount(1 To 20) As Integer

End Type

Public Type hMetamorfosis
    Status As Byte
    Body As Integer
    Fuerza As String
    Agilidad As String
    Inteligencia As String
End Type

Public Type tHechizo

    nombre As String
    Desc As String
    PalabrasMagicas As String

    ExclusivoClase As String

    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String

    Resis As Byte

    Tipo As TipoHechizo

    WAV As Integer
    FXgrh As Integer
    loops As Byte

    SubeHP As Integer
    MinHP As Integer
    MaxHP As Integer

    SubeMana As Byte
    MinMana As Long
    ManMana As Long

    SubeSta As Byte
    MinSta As Integer
    MaxSta As Integer

    SubeHam As Byte
    MinHam As Integer
    MaxHam As Integer

    SubeSed As Byte
    MinSed As Integer
    MaxSed As Integer

    SubeAgilidad As Byte
    MinAgilidad As Integer
    MaxAgilidad As Integer

    SubeFuerza As Byte
    MinFuerza As Integer
    MaxFuerza As Integer

    SubeCarisma As Byte
    MinCarisma As Integer
    MaxCarisma As Integer

    Invisibilidad As Byte
    Paraliza As Byte
    Inmoviliza As Byte
    RemoverParalisis As Byte
    RemoverEstupidez As Byte
    CuraVeneno As Byte
    Envenena As Byte
    Maldicion As Byte
    RemoverMaldicion As Byte
    Bendicion As Byte
    Estupidez As Byte
    Ceguera As Byte
    Revivir As Byte
    Morph As Byte
    Mimetiza As Byte
    RemueveInvisibilidadParcial As Byte
    Lluvia As Byte

    invoca As Byte
    NumNpc As Integer
    Cant As Integer

    Materializa As Byte
    ItemIndex As Byte

    MinSkill As Integer
    ManaRequerido As Integer

    'Barrin 29/9/03
    StaRequerido As Integer

    Target As TargetType

    NeedStaff As Integer
    StaffAffected As Boolean

    ClaseProhibida(1 To 20) As String

    Real As Integer
    Caos As Integer
    Nemes As Integer
    Templ As Integer

    ParalisisArea As Byte

    Metamorfosis As hMetamorfosis

End Type

Public Type UserOBJ

    ObjIndex As Integer
    Amount As Integer
    Equipped As Byte
    ProbTirar As Byte

End Type

Public Type Inventario

    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte
    EscudoEqpObjIndex As Integer
    EscudoEqpSlot As Byte
    CascoEqpObjIndex As Integer
    CascoEqpSlot As Byte
    AlaEqpObjIndex As Integer
    AlaEqpSlot As Byte
    MunicionEqpObjIndex As Integer
    MunicionEqpSlot As Byte
    HerramientaEqpObjIndex As Integer
    HerramientaEqpSlot As Integer
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    NroItems As Integer
    AmuletoEqpObjIndex As Integer
    AmuletoEqpSlot As Byte

End Type

Public Type tPartyData

    PIndex As Integer
    RemXP As Double    'La exp. en el server se cuenta con Doubles
    TargetUser As Integer    'Para las invitaciones

End Type

Public Type Position

    x As Integer
    Y As Integer

End Type

Public Type WorldPos

    Map As Integer
    x As Integer
    Y As Integer

End Type

'Datos de user o npc
Public Type char

    CharIndex As Integer
    Head As Integer
    Body As Integer

    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer

    '[MaTeO 9]
    Alas As Integer
    '[/MaTeO 9]

    FX As Integer
    loops As Integer

    heading As eHeading

    delay As Integer    'CRAW; 18/09/2019

    Fuerza As Integer
    Agilidad As Integer
    Inteligencia As Integer

End Type

Public Type AmuletoDefensa
    TipoBonifica As Byte
    Bonifica As Integer
End Type

Public Type tObjPack
    Objeto As Integer
    Cantidad As Integer
End Type

Public Type tPack
    NumObjs As Byte
    Obj(1 To 10) As tObjPack
End Type

'Tipos de objetos
Public Type ObjData

    Name As String    'Nombre del obj

    ObjType As eOBJType    'Tipo enum que determina cuales son las caract del obj

    GrhIndex As Long    ' Indice del grafico que representa el obj
    GrhSecundario As Long

    'Solo contenedores
    MaxItems As Integer
    Conte As Inventario
    Apuñala As Byte
    Pegadoble As Byte
    DosManos As Byte

    HechizoIndex As Integer

    ForoID As String

    MinHP As Integer    ' Minimo puntos de vida
    MaxHP As Integer    ' Maximo puntos de vida

    MineralIndex As Integer
    LingoteInex As Integer

    '[MaTeO 9]
    Alas As Integer
    '[/MaTeO 9]

    Proyectil As Integer
    Municion As Integer

    Nivel As Byte    'nivel minimo usar item

    Crucial As Byte
    Newbie As Integer

    'Puntos de Stamina que da
    MinSta As Integer    ' Minimo puntos de stamina

    'Pociones
    TipoPocion As Byte
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    MinSkill As Integer
    LingoteIndex As Integer

    MinHit As Integer    'Minimo golpe
    MaxHit As Integer    'Maximo golpe

    MinHam As Integer
    MinSed As Integer

    def As Integer
    MinDef As Integer    ' Armaduras
    MaxDef As Integer    ' Armaduras

    Ropaje As Integer    'Indice del grafico del ropaje

    WeaponAnim As Integer    ' Apunta a una anim de armas
    ShieldAnim As Integer    ' Apunta a una anim de escudo
    CascoAnim As Integer

    Valor As Long     ' Precio

    Cerrada As Integer
    Llave As Byte
    clave As Long    'si clave=llave la puerta se abre o cierra

    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer

    RazaEnana As Byte
    RazaHobbit As Byte
    RazaVampiro As Byte
    RazaOrco As Byte

    Mujer As Byte
    Hombre As Byte

    Envenena As Byte
    Paraliza As Byte
    Agarrable As Byte

    Zona As Integer

    LingH As Integer
    LingO As Integer
    LingP As Integer
    LingM As Integer
    Gemas As Integer
    Diamantes As Integer

    Madera As Integer

    Lana As Integer
    Lobo As Integer
    Osos As Integer
    Tigre As Integer
    Jabali As Integer
    LoboPolar As Integer
    OsoPolar As Integer
    Vaca As Integer

    Hierba As Long

    SkHerreria As Integer
    SkCarpinteria As Integer
    SkSastreria As Integer
    SkHechiceria As Integer

    texto As String

    'Clases que no tienen permitido usar este obj
    ClaseProhibida(1 To 20) As String

    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer

    Real As Integer
    Caos As Integer
    Nemes As Integer
    Templ As Integer

    Cae As Integer

    ObjetoEspecial As Long

    StaffPower As Integer
    VaraDragon As Byte

    StaffDamageBonus As Integer
    DefensaMagicaMax As Integer
    DefensaMagicaMin As Integer
    Refuerzo As Byte

    Expe As Long

    Regalos As Integer

    Gm As Byte
    sagrado As Byte
    Limpiar As Byte

    AmuletoDefensa As AmuletoDefensa

    TipoRegalo As Byte
    Pack As tPack

    NoRobable As Byte
End Type

Public Type Regalos
    ObjIndex As Long
End Type

Public Type Obj

    ObjIndex As Integer
    Amount As Integer

End Type

Public Type EncZeus

    ACT As Byte
    Tiempo As Integer
    EncSI As Integer
    EncNO As Integer

End Type

Public Encuesta As EncZeus
'[KEVIN]
'Banco Objs
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
'[/KEVIN]

'[KEVIN]
Public Type BancoInventario

    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserOBJ
    NroItems As Integer

End Type

'[/KEVIN]

'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'******* T I P O S   D E    U S U A R I O S **************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Type tReputacion    'Fama del usuario

    NobleRep As Double
    BurguesRep As Double
    PlebeRep As Double
    LadronesRep As Double
    BandidoRep As Double
    AsesinoRep As Double
    Promedio As Double

End Type

'Estadisticas de los usuarios
Public Type UserStats

    TrofOro As Byte
    TrofBronce As Byte
    TrofPlata As Byte
    TrofMadera As Byte
    GLD As Long    'Dinero
    Banco As Long
    MET As Integer

    ' puntos
    PuntosTorneo As Integer
    PuntosRetos As Integer
    PuntosDuelos As Integer

    MaxHP As Integer
    MinHP As Integer

    FIT As Integer
    MaxSta As Integer
    MinSta As Integer
    MaxMAN As Long
    MinMAN As Long
    MaxHit As Integer
    MinHit As Integer

    MaxHam As Integer
    MinHam As Integer

    MaxAGU As Integer
    MinAGU As Integer

    def As Integer
    Exp As Double
    ELV As Long
    ELU As Long
    UserSkills(1 To NUMSKILLS) As Integer
    UserAtributos(1 To NUMATRIBUTOS) As Integer
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Integer
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    UsuariosMatados As Integer
    CriminalesMatados As Integer
    NPCsMuertos As Integer

    SkillPts As Integer

    CleroMatados As Long
    AbbadonMatados As Long
    TemplarioMatados As Long
    TinieblaMatados As Long

End Type
'2vs2
Public Pareja As Pareja

Public Type Pareja
    Jugador1 As Integer
    Jugador2 As Integer
    Jugador3 As Integer
    Jugador4 As Integer
End Type

'Flags
Public Type UserFlags

    SeleccioneA As String
    EstoySelec As Long
    
    SuPareja As Integer
    EsperaPareja As Boolean
    EnPareja As Boolean

    'Casarse?
    Casado As Byte
    Casandose As Boolean    '¿Esta con otro wacho la muy perra? O con otra wacha el chamuyerosonononon?
    Quien As Integer    'El otro
    QuienName As String
    SolicitudC As String

    EnCvc As Boolean
    CvcBlue As Byte
    CvcRed As Byte    ' by ZaikO Dieguito; Tu Papá !
    SeguroCVC As Boolean

    'Jua
    Embarcado As Byte
    PuedeSumon As Boolean
    envioSol As Boolean
    Potea As Boolean
    Soporteo As Boolean
    RecibioSol As Boolean
    compa As Integer
    EnDosVDos As Boolean
    ParejaMuerta As Boolean
    Montado As Boolean
    NumeroMont As Integer
    VotEnc As Boolean
    EsperandoDuelo As Boolean
    EsperandoDuelo1 As Boolean
    EstaDueleando As Boolean
    EstaDueleando1 As Boolean
    automatico As Boolean

    bandas As Boolean
    Demonio As Boolean
    Angel As Boolean

    medusas As Boolean
    Corsarios As Boolean
    Piratas As Boolean

    death As Boolean
    Oponente As Integer
    Oponente1 As Integer
    EstaEmpo As Byte    'Empollando (by yb)
    Muerto As Byte    '¿Esta muerto?
    Escondido As Byte    '¿Esta escondido?
    Comerciando As Boolean    '¿Esta comerciando?
    UserLogged As Boolean    '¿Esta online?
    Meditando As Boolean
    Descuento As String
    Hambre As Byte
    Sed As Byte
    PuedeMoverse As Byte
    TimerLanzarSpell As Long
    PuedeTrabajar As Byte
    Envenenado As Byte
    ClienteOK As Boolean    'CHOTS | Comprobacion de cliente
    YaDenuncio As Byte
    Paralizado As Byte
    Estupidez As Byte
    Ceguera As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    Oculto As Byte
    Desnudo As Byte
    Descansar As Boolean
    Hechizo As Integer
    Navegando As Byte

    Seguro As Boolean
    SeguroClan As Boolean
    SeguroCombate As Boolean
    SeguroObjetos As Boolean
    SeguroHechizos As Boolean

    TomoPocionAmarilla As Boolean
    TomoPocionVerde As Boolean
    DuracionEfectoVerdes As Integer
    DuracionEfectoAmarillas As Integer

    TargetNpc As Integer    ' Npc señalado por el usuario
    TargetNpcTipo As eNPCType    ' Tipo del npc señalado
    NpcInv As Integer

    Ban As Byte
    AdministrativeBan As Byte

    TargetUser As Integer    ' Usuario señalado

    TargetObj As Integer    ' Obj señalado
    TargetObjMap As Integer
    TargetObjX As Integer
    TargetObjY As Integer

    TargetMap As Integer
    TargetX As Integer
    TargetY As Integer

    TargetObjInvIndex As Integer
    TargetObjInvSlot As Integer

    AtacadoPorNpc As Integer
    AtacadoPorUser As Integer

    StatsChanged As Byte
    Privilegios As PlayerType
    EsRolesMaster As Boolean

    LastCrimMatado As String
    LastCiudMatado As String

    OldBody As Integer
    OldHead As Integer
    AdminInvisible As Byte

    '[Barrin 30-11-03]
    TimesWalk As Long
    StartWalk As Long
    CountSH As Long
    '[/Barrin 30-11-03]

    '[CDT 17-02-04]
    UltimoMensaje As Byte
    '[/CDT]

    PertAlCons As Byte
    PertAlConsCaos As Byte

    Silenciado As Byte
    Mimetizado As Byte
    CentinelaOK As Boolean    'Centinela
    Quest As Byte

    EspecialFuerza As Integer
    EspecialAgilidad As Integer
    EspecialArco As Integer
    EspecialObjArco As Integer

    HechizoVeneno As Integer

    pendingUpdate As Boolean    'CRAW; 18/09/2019

    RPasswd As String

    ValidBank As Byte

    Metamorfosis As Byte

    Licantropo As Byte

    UsoLibroHP As Byte
    
    HablanMute As Byte

End Type

Public Type UserCounters
    IdleCount As Long
    AttackCounter As Integer
    HPCounter As Integer
    STACounter As Integer
    Frio As Integer
    COMCounter As Integer
    AGUACounter As Integer
    Veneno As Integer
    Paralisis As Integer
    Ceguera As Integer
    Estupidez As Integer
    Invisibilidad As Integer
    Mimetismo As Integer
    PiqueteC As Long
    Pena As Long
    SendMapCounter As WorldPos
    Pasos As Integer
    AntiSH As Integer
    '[Gonzalo]
    Saliendo As Boolean
    Salir As Integer
    '[/Gonzalo]

    'Barrin 3/10/03
    tInicioMeditar As Long
    bPuedeMeditar As Boolean
    'Barrin

    TimerLanzarSpell As Long
    TimerPuedeAtacar As Long
    TimerPuedeTrabajar As Long
    TimerUsar As Long

    Trabajando As Long  ' Para el centinela
    Ocultando As Long   ' Unico trabajo no revisado por el centinela

    validInputs As Long    'CRAW; 18/09/2019

    Cerdo As Integer

    Metamorfosis As Integer

    TimerAttack As Integer

End Type

Public MuereSpell As Integer
Public LoopSpell As Integer

Public Type tFacciones

    ArmadaReal As Long
    FuerzasCaos As Long

    Templario As Byte
    Nemesis As Byte

    CriminalesMatados As Double
    CiudadanosMatados As Double

    RecompensasReal As Byte
    RecompensasCaos As Byte

    RecompensasTemplaria As Byte
    RecompensasNemesis As Byte

    RecibioExpInicialReal As Byte
    RecibioExpInicialCaos As Byte

    RecibioExpInicialTemplaria As Byte
    RecibioExpInicialNemesis As Byte

    RecibioArmaduraReal As Byte
    RecibioArmaduraCaos As Byte

    RecibioArmaduraNemesis As Byte
    RecibioArmaduraTemplaria As Byte

    Reenlistadas As Byte
    ArmaduraFaccionaria As Integer
    NextRecompensas As Integer

    Reenlistado As Integer

    FEnlistado As String

End Type

Public Type UserSagrada

    Enabled As Long
    MinHit As Integer
    MaxHit As Integer

End Type

Public Type tClan

    FundoClan As Integer
    PuntosClan As Long
    UMuerte As String
    Timer As Integer
    ParticipoClan As Integer

End Type

Public Type tCastillos
    Norte As Byte
    Oeste As Byte
    Este As Byte
    Sur As Byte
    Fortaleza As Byte

    tNorte As Byte
    tOeste As Byte
    tEste As Byte
    tSur As Byte
    tFortaleza As Byte
End Type

Public Type tMetamorfosis
    Angel As Byte
    Demonio As Byte
End Type

Public Type tGm
    Command(1 To 18) As Byte
End Type

Public Type UserQuest
      UserQuest(1 To 1000) As Integer
      Quest As Integer
      Start As Byte
      NumNpc As Byte
      MataNpc(1 To 10) As Integer
      NumObj As Byte
      BuscaObj(1 To 10) As Integer
      NumMap As Byte
      Mapa(1 To 10) As Integer
      ValidNpcDD As Byte
      MapaNpcDD As Integer
      Icono As Integer
      ValidNpcDescubre As Byte
      PreguntaDescubre As Integer
      NumObjNpc As Byte
      DarObjNpc(1 To 10) As Integer
      DarObjNpcEntrega As Byte
      ValidHablarNpc As Byte
      UserHablaNpc As Byte
      ValidMatarUser As Byte
      UserMatados As Integer
End Type

Public Type tIgnore
     NumIgnores As Integer
     Usuario(1 To 100) As String
     MaximoIgnores As Integer
End Type

Public Type User

    Pareja As String
    Zona As Integer
    Lac As TLac

    ' soporte
    Pregunta As String
    Respuesta As String

    Name As String
    Id As Long
    EnCvc As Boolean
    ViejaPos As WorldPos

    showName As Boolean    'Permite que los GMs oculten su nick con el comando /SHOWNAME

    modName As String
    Password As String
    PalabraSecreta As String
    RecuperarPassword As String
    hd_String As String

    char As char    'Define la apariencia
    CharMimetizado As char
    OrigChar As char

    Desc As String    ' Descripcion
    DescRM As String

    Clase As String
    Raza As String
    Genero As String
    Email As String
    Hogar As String

    Invent As Inventario

    pos As WorldPos

    ConnIDValida As Boolean
    ConnID As Long    'ID
    RDBuffer As String    'Buffer roto

    BancoInvent As BancoInventario

    Counters As UserCounters

    MascotasIndex(1 To MAXMASCOTAS) As Integer
    MascotasType(1 To MAXMASCOTAS) As Integer
    NroMacotas As Integer

    Stats As UserStats
    flags As UserFlags


    BytesTransmitidosUser As Long
    BytesTransmitidosSvr As Long

    Reputacion As tReputacion
    Faccion As tFacciones

    ip As String

    ComUsu As tCOmercioUsuario

    Asedio As flagsAsedio

    EmpoCont As Byte

    GuildIndex As Integer   'puntero al array global de guilds
    EscucheClan As Integer

    PartyIndex As Integer   'index a la party q es miembro
    PartySolicitud As Integer   'index a la party q solicito

    AreasInfo As AreaInfo

    Telepatia As Byte

    Sagrada As UserSagrada

    AoMCreditos As Long
    AoMCanjes As Long

    Clan As tClan

    GranPoder As Byte

    DañoVeneno As Integer
    TipoVeneno As Integer
    AumentoVeneno As Integer

    Castillos As tCastillos

    Metamorfosis As tMetamorfosis

    SnapShot As Boolean
    SnapShotAdmin As Integer

    Gm As tGm
    
    Quest As UserQuest
    
    clave2 As Long
    clave As String
    
    Ignore As tIgnore

End Type

'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'**  T I P O S   D E    N P C S **************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Type NPCStats

    Alineacion As Integer
    MaxHP As Long
    MinHP As Long
    MaxHit As Integer
    MinHit As Integer
    def As Integer
    UsuariosMatados As Integer

End Type

Public Type NpcCounters

    Paralisis As Integer
    TiempoExistencia As Long

End Type

Public Type NPCFlags

    AfectaParalisis As Byte
    GolpeExacto As Byte
    Domable As Integer
    Respawn As Byte
    NPCActive As Boolean    '¿Esta vivo?
    Follow As Boolean
    Faccion As Byte
    LanzaSpells As Byte

    '[KEVIN]
    'DeQuest As Byte

    'ExpDada As Long
    ExpCount As Long    '[ALEJO]
    '[/KEVIN]

    OldMovement As TipoAI
    OldHostil As Byte

    AguaValida As Byte
    TierraInvalida As Byte

    Sound As Integer
    Attacking As Integer
    AttackedBy As String
    Category1 As String
    Category2 As String
    Category3 As String
    Category4 As String
    Category5 As String
    BackUp As Byte
    RespawnOrigPos As Byte

    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte

    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer

    AtacaAPJ As Integer
    AtacaANPC As Integer
    AIAlineacion As e_Alineacion
    AIPersonalidad As e_Personalidad

End Type

Public Type tCriaturasEntrenador

    NpcIndex As Integer
    NpcName As String
    TmpIndex As Integer

End Type

' New type for holding the pathfinding info
Public Type NpcPathFindingInfo

    Path() As tVertice      ' This array holds the path
    Target As Position      ' The location where the NPC has to go
    PathLenght As Integer   ' Number of steps *
    CurPos As Integer       ' Current location of the npc
    TargetUser As Integer   ' UserIndex chased
    NoPath As Boolean       ' If it is true there is no path to the target location

    '* By setting PathLenght to 0 we force the recalculation
    '  of the path, this is very useful. For example,
    '  if a NPC or a User moves over the npc's path, blocking
    '  its way, the function NpcLegalPos set PathLenght to 0
    '  forcing the seek of a new path.

End Type

' New type for holding the pathfinding info

Public Type npc

    Name As String
    char As char    'Define como se vera
    Desc As String
    DescExtra As String

    NPCtype As eNPCType
    Numero As Integer

    level As Integer

    InvReSpawn As Byte

    Comercia As Integer
    Target As Long
    TargetNpc As Long
    TipoItems As Integer

    Veneno As Byte

    pos As WorldPos    'Posicion
    Orig As WorldPos
    SkillDomar As Integer

    Movement As TipoAI
    Attackable As Byte
    Hostile As Byte
    PoderAtaque As Long
    PoderEvasion As Long

    Inflacion As Long
    GiveEXP As Long
    GiveGLD As Long
    Drops As NpcDrop
    DefensaMagica As Byte

    Stats As NPCStats
    flags As NPCFlags
    Contadores As NpcCounters

    MurallaEquipo As Byte
    MurallaIndex As Byte

    Invent As Inventario
    CanAttack As Byte

    NroExpresiones As Byte
    Expresiones() As String    ' le da vida ;)

    NroSpells As Byte
    Spells() As Integer  ' le da vida ;)

    '<<<<Entrenadores>>>>>
    NroCriaturas As Integer
    Criaturas() As tCriaturasEntrenador
    MaestroUser As Integer
    MaestroNpc As Integer
    Mascotas As Integer

    ' New!! Needed for pathfindig
    PFINFO As NpcPathFindingInfo
    AreasInfo As AreaInfo

End Type

'**********************************************************
'**********************************************************
'******************** Tipos del mapa **********************
'**********************************************************
'**********************************************************
'Tile
Public Type MapBlock

    Blocked As Byte
    Graphic(1 To 4) As Long
    UserIndex As Integer
    NpcIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Trigger As eTrigger

End Type

'Info del mapa
Type MapInfo

    criatinv As Integer
    NumUsers As Integer
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    Pk As Boolean
    MagiaSinEfecto As Byte
    OcultarSinEfecto As Byte

    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte

End Type

'********** V A R I A B L E S     P U B L I C A S ***********

Public SERVERONLINE As Boolean
Public ULTIMAVERSION As String
Public BackUp As Boolean            ' TODO: Se usa esta variable ?

Public ListaRazas(1 To NUMRAZAS) As String
Public Torneo_Clases_Validas(1 To 8) As String
Public Torneo_Alineacion_Validas(1 To 8) As String
Public Torneo_Clases_Validas2(1 To 8) As Integer
Public Torneo_Alineacion_Validas2(1 To 4) As Integer
Public SkillsNames(1 To NUMSKILLS) As String
Public ListaClases(1 To NUMCLASES) As String

Public Const ENDL As String * 2 = vbCrLf
Public Const ENDC As String * 1 = vbNullChar
Public Multexp As Byte
Public MultOro As Byte
Public MultMsg As String
Public recordusuarios As Long

'
'Directorios
'

''
'Ruta base del server, en donde esta el "server.ini"
Public IniPath As String

''
'Ruta base para guardar los chars
Public CharPath As String

''
'Ruta base para los archivos de mapas
Public MapPath As String

''
'Ruta base para los DATs
Public DatPath As String

''
'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

Public ResPos As WorldPos             ' TODO: Se usa esta variable ?

''
'Posicion de comienzo
Public StartPos As WorldPos             ' TODO: Se usa esta variable ?

''
'Numero de usuarios actual
Public NumUsers As Integer
Public LastUser As Integer
Public LastChar As Integer
Public NumChars As Integer
Public LastNPC As Integer
Public NumNPCs As Integer
Public NumFX As Integer
Public NumMaps As Integer
Public NumObjDatas As Integer
Public NumeroHechizos As Integer
Public NumRegalos As Integer
Public AllowMultiLogins As Byte
Public IdleLimit As Integer
Public MaxUsers As Integer
Public HideMe As Byte
Public LastBackup As String
Public Minutos As String
Public haciendoBK As Boolean
Public Torneo_SumAuto As Integer
Public Torneo_Map As Integer
Public Torneo_X As Integer
Public Torneo_Y As Integer
Public Hay_Torneo As Boolean
'Public Torneo_Clases_Validas() As String
'Public Torneo_Clases_Validas2() As Integer
'Public Torneo_Alineacion_Validas() As String
'Public Torneo_Alineacion_Validas2() As Integer
Public Torneo_Nivel_Minimo As Long
Public Torneo_Nivel_Maximo As Long
Public Torneo_Cantidad As Long
Public Torneo_Inscriptos As Long
Public Oscuridad As Integer
Public NocheDia As Integer
Public PuedeCrearPersonajes As Integer
Public CamaraLenta As Integer
Public ServerSoloGMs As Integer

''
'Esta activada la verificacion MD5 ?
Public MD5ClientesActivado As Byte

Public EnPausa As Boolean
Public EnTesting As Boolean
Public EncriptarProtocolosCriticos As Boolean
Public Nombre1 As String
Public Nombre2 As String

'*****************ARRAYS PUBLICOS*************************
Public UserList() As User         'USUARIOS
Public Npclist() As npc        'NPCS
Public MapData() As MapBlock
Public MapInfo() As MapInfo
Public Hechizos() As tHechizo
Public CharList() As Integer
Public ObjData() As ObjData
Public Regalos() As Regalos
Public SpawnList() As tCriaturasEntrenador
Public ForbidenNames() As String
Public ArmasHerrero() As Integer
Public ArmadurasHerrero() As Integer
Public ObjCarpintero() As Integer
Public ObjSastre() As Integer
Public ObjHechizeria() As Integer
Public ObjArmaHerreroMagico() As Integer
Public ObjArmaduraHerreroMagico() As Integer
Public MD5s() As String
Public BanIps As New Collection
Public Parties() As clsParty
'*********************************************************

Public Nix As WorldPos
Public Ullathorpe As WorldPos
Public Banderbill As WorldPos
Public Lindos As WorldPos

Public Prision As WorldPos
Public Libertad As WorldPos

Public Ayuda As New cCola
Public Quest As New cCola
Public Torneo As New cCola
Public SonidosMapas As New SoundMapInfo

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                                                              ByVal lpOperation As String, _
                                                                              ByVal lpFile As String, _
                                                                              ByVal lpParameters As String, _
                                                                              ByVal lpDirectory As String, _
                                                                              ByVal nShowCmd As Long) As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                                                                     ByVal lpKeyname As Any, _
                                                                                                     ByVal lpString As String, _
                                                                                                     ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                                                                 ByVal lpKeyname As Any, _
                                                                                                 ByVal lpdefault As String, _
                                                                                                 ByVal lpreturnedstring As String, _
                                                                                                 ByVal nsize As Long, _
                                                                                                 ByVal lpfilename As String) As Long

Public Enum e_ObjetosCriticos

    Manzana = 1
    Manzana2 = 2
    ManzanaNewbie = 467

End Enum

Public MaxLevel As Long
Public UserMaxLevel As String

'Max/Min Guerrero
Public GCONST21MAXVIDA As Byte
Public GCONST21MINVIDA As Byte
Public GCONST20MAXVIDA As Byte
Public GCONST20MINVIDA As Byte
Public GCONST19MAXVIDA As Byte
Public GCONST19MINVIDA As Byte
Public GCONST18MAXVIDA As Byte
Public GCONST18MINVIDA As Byte
Public GCONST17MAXVIDA As Byte
Public GCONST17MINVIDA As Byte
Public GCONSTOTROMAXVIDA As Byte
Public GCONSTOTROMINVIDA As Byte
'Max/Min Cazador
Public CCONST21MAXVIDA As Byte
Public CCONST21MINVIDA As Byte
Public CCONST20MAXVIDA As Byte
Public CCONST20MINVIDA As Byte
Public CCONST19MAXVIDA As Byte
Public CCONST19MINVIDA As Byte
Public CCONST18MAXVIDA As Byte
Public CCONST18MINVIDA As Byte
Public CCONST17MAXVIDA As Byte
Public CCONST17MINVIDA As Byte
Public CCONSTOTROMAXVIDA As Byte
Public CCONSTOTROMINVIDA As Byte
'Max/Min Paladin
Public PCONST21MAXVIDA As Byte
Public PCONST21MINVIDA As Byte
Public PCONST20MAXVIDA As Byte
Public PCONST20MINVIDA As Byte
Public PCONST19MAXVIDA As Byte
Public PCONST19MINVIDA As Byte
Public PCONST18MAXVIDA As Byte
Public PCONST18MINVIDA As Byte
Public PCONST17MAXVIDA As Byte
Public PCONST17MINVIDA As Byte
Public PCONSTOTROMAXVIDA As Byte
Public PCONSTOTROMINVIDA As Byte
'Max/Min Mago
Public MCONST21MAXVIDA As Byte
Public MCONST21MINVIDA As Byte
Public MCONST20MAXVIDA As Byte
Public MCONST20MINVIDA As Byte
Public MCONST19MAXVIDA As Byte
Public MCONST19MINVIDA As Byte
Public MCONST18MAXVIDA As Byte
Public MCONST18MINVIDA As Byte
Public MCONST17MAXVIDA As Byte
Public MCONST17MINVIDA As Byte
Public MCONSTOTROMAXVIDA As Byte
Public MCONSTOTROMINVIDA As Byte
'Max/Min Clerigo
Public CLCONST21MAXVIDA As Byte
Public CLCONST21MINVIDA As Byte
Public CLCONST20MAXVIDA As Byte
Public CLCONST20MINVIDA As Byte
Public CLCONST19MAXVIDA As Byte
Public CLCONST19MINVIDA As Byte
Public CLCONST18MAXVIDA As Byte
Public CLCONST18MINVIDA As Byte
Public CLCONST17MAXVIDA As Byte
Public CLCONST17MINVIDA As Byte
Public CLCONSTOTROMAXVIDA As Byte
Public CLCONSTOTROMINVIDA As Byte
'Max/Min Asesino
Public ACONST21MAXVIDA As Byte
Public ACONST21MINVIDA As Byte
Public ACONST20MAXVIDA As Byte
Public ACONST20MINVIDA As Byte
Public ACONST19MAXVIDA As Byte
Public ACONST19MINVIDA As Byte
Public ACONST18MAXVIDA As Byte
Public ACONST18MINVIDA As Byte
Public ACONST17MAXVIDA As Byte
Public ACONST17MINVIDA As Byte
Public ACONSTOTROMAXVIDA As Byte
Public ACONSTOTROMINVIDA As Byte
'Max/Min Bardo
Public BACONST21MAXVIDA As Byte
Public BACONST21MINVIDA As Byte
Public BACONST20MAXVIDA As Byte
Public BACONST20MINVIDA As Byte
Public BACONST19MAXVIDA As Byte
Public BACONST19MINVIDA As Byte
Public BACONST18MAXVIDA As Byte
Public BACONST18MINVIDA As Byte
Public BACONST17MAXVIDA As Byte
Public BACONST17MINVIDA As Byte
Public BACONSTOTROMAXVIDA As Byte
Public BACONSTOTROMINVIDA As Byte
'Max/Min Ladron
Public LCONST21MAXVIDA As Byte
Public LCONST21MINVIDA As Byte
Public LCONST20MAXVIDA As Byte
Public LCONST20MINVIDA As Byte
Public LCONST19MAXVIDA As Byte
Public LCONST19MINVIDA As Byte
Public LCONST18MAXVIDA As Byte
Public LCONST18MINVIDA As Byte
Public LCONST17MAXVIDA As Byte
Public LCONST17MINVIDA As Byte
Public LCONSTOTROMAXVIDA As Byte
Public LCONSTOTROMINVIDA As Byte
'Max/Min Druida
Public DCONST21MAXVIDA As Byte
Public DCONST21MINVIDA As Byte
Public DCONST20MAXVIDA As Byte
Public DCONST20MINVIDA As Byte
Public DCONST19MAXVIDA As Byte
Public DCONST19MINVIDA As Byte
Public DCONST18MAXVIDA As Byte
Public DCONST18MINVIDA As Byte
Public DCONST17MAXVIDA As Byte
Public DCONST17MINVIDA As Byte
Public DCONSTOTROMAXVIDA As Byte
Public DCONSTOTROMINVIDA As Byte
'Max/Min Trabajador
Public TCONST21MAXVIDA As Byte
Public TCONST21MINVIDA As Byte
Public TCONST20MAXVIDA As Byte
Public TCONST20MINVIDA As Byte
Public TCONST19MAXVIDA As Byte
Public TCONST19MINVIDA As Byte
Public TCONST18MAXVIDA As Byte
Public TCONST18MINVIDA As Byte
Public TCONST17MAXVIDA As Byte
Public TCONST17MINVIDA As Byte
Public TCONSTOTROMAXVIDA As Byte
Public TCONSTOTROMINVIDA As Byte
'Max/Min Brujo
Public BCONST21MAXVIDA As Byte
Public BCONST21MINVIDA As Byte
Public BCONST20MAXVIDA As Byte
Public BCONST20MINVIDA As Byte
Public BCONST19MAXVIDA As Byte
Public BCONST19MINVIDA As Byte
Public BCONST18MAXVIDA As Byte
Public BCONST18MINVIDA As Byte
Public BCONST17MAXVIDA As Byte
Public BCONST17MINVIDA As Byte
Public BCONSTOTROMAXVIDA As Byte
Public BCONSTOTROMINVIDA As Byte
'Max/Min Arquero
Public ARCONST21MAXVIDA As Byte
Public ARCONST21MINVIDA As Byte
Public ARCONST20MAXVIDA As Byte
Public ARCONST20MINVIDA As Byte
Public ARCONST19MAXVIDA As Byte
Public ARCONST19MINVIDA As Byte
Public ARCONST18MAXVIDA As Byte
Public ARCONST18MINVIDA As Byte
Public ARCONST17MAXVIDA As Byte
Public ARCONST17MINVIDA As Byte
Public ARCONSTOTROMAXVIDA As Byte
Public ARCONSTOTROMINVIDA As Byte

'Dateador de experencia cada niveles
Public levelELU(1 To STAT_MAXELV) As Long

Public StatusInvo As Boolean
Public ConfInvo As Long
Public OnMin As Long
Public OnHor As Long
Public OnDay As Long
Public NumGm As Long
Public NumClan As Long

Public UserArmada As String
Public UserRecompensas As Long

Public Const OroDivorciarse As Long = 500000
Public Const OroHechizo As Long = 100000
Public Const OroCirujia As Long = 10000

Public Const H_Demonio As Integer = 63
Public Const H_DemonioII As Integer = 64
Public Const H_Angel As Integer = 65
Public Const H_AngelII As Integer = 66

Public Enum CabezaDragon
    Roja = 899
    negra = 893
    Verde = 896
    lila = 897
    Blanca = 898
    naranja = 895
    Azul = 894
End Enum

Public Enum Plumas
    Ampere = 1688
    Bassinger = 1689
    Seth = 1690
End Enum

Public Enum AlasCaos
    One = 1691
    Second = 1692
    Thir = 1693
    Four = 1694
End Enum

Public Enum AlasReal
    One = 1704
    Second = 1736
    Thir = 1703
    Four = 1697
End Enum

Public Enum AlasNemesis
    One = 1751
    Second = 1752
    Thir = 1770
    Four = 1698
End Enum

Public Enum AlasTemplario
    One = 1737
    Second = 1738
    Thir = 1749
    Four = 1699
End Enum

Public Const IntervaloAttack As Integer = 12

Public Const ObjCreacionAlas As Integer = 1695

Public AodefConv As AoDefenderConverter
Public SuperClave As String
