Attribute VB_Name = "Mod_Declaraciones"
Option Explicit

'CRAW; 18/09/2019 --> VARIABLES AUTO L PUBLICAS
Public TickCountClient As Long
Public TickCountServer As Long
Public delayCl(0 To 3) As Long
Public delaySv(0 To 1) As Long
Public requestPing As Byte
Public SeguroCvc As Boolean

Public Const CASPER_HEAD       As Integer = 500
Public Const FRAGATA_FANTASMAL As Integer = 87

Public IsSeguro                As Boolean
Public IsSeguroClan            As Boolean
Public IsSeguroCombate         As Boolean
Public IsSeguroHechizos        As Boolean
Public IsSeguroObjetos         As Boolean

Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
                             
Public Const VK_SNAPSHOT As Byte = 44 ' PrintScreen virtual keycode
Public Const PS_TheForm = 0
Public Const PS_TheScreen As Byte = 1

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Public DragPantalla        As Boolean
Public ChangeFont          As Boolean
Public CustomKeys          As New clsCustomKeys
'"51.38.175.174"
'"213.239.214.69"
'"127.0.0.1"
Public Const CurServerIp   As String = "127.0.0.1"
Public Const CurServerPort As Integer = 9879

Public Centrada            As Boolean
Public CartelInvisibilidad As Integer

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Estadisticas As Boolean, UserClick As String, ClickMatados As String, ClickClase As String
Public TiempoEst As Byte


'CRAW; 24/03/2020
Public UserClicado As Integer

Public Msn                    As Boolean
Public NameMap                As String


Public TiempoAsedio As Long

'Objetos públicos
Public DialogosClanes         As New clsGuildDlg
Public Dialogos               As New cDialogos
Public Audio                  As New clsAudio
Public Inventario             As New clsGrapchicalInventory

'Sonidos
Public Const SND_CLICK        As String = "click.Wav"
Public Const SND_PASOS1       As String = "23.Wav"
Public Const SND_PASOS2       As String = "24.Wav"
Public Const SND_PASOS3       As String = "176.Wav"
Public Const SND_PASOS4       As String = "177.Wav"
Public Const SND_NAVEGANDO    As String = "50.wav"
Public Const SND_OVER         As String = "click2.Wav"
Public Const SND_DICE         As String = "cupdice.Wav"
Public Const SND_LLUVIAINEND  As String = "lluviainend.wav"
Public Const SND_LLUVIAOUTEND As String = "lluviaoutend.wav"

Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

'Musica
Public Const MIdi_Inicio  As Byte = 6

Public ColoresPJ(0 To 50) As Long

Public currentMidi        As Long

Public ArmaMin            As Integer
Public ArmaMax            As Integer
Public ArmorMin           As Integer
Public ArmorMax           As Integer
Public EscuMin            As Integer
Public EscuMax            As Integer
Public CascMin            As Integer
Public MagMin             As Integer
Public MagMax             As Integer
Public CascMax            As Integer
Public Verde              As Integer
Public Amarilla           As Integer
Public VidaVerde          As Long
Public VidaAmarilla       As Long
Public HDD                As Long
Public UserClan           As String
Public CreandoClan        As Boolean
Public ClanName           As String
Public Site               As String

Public UserCiego          As Boolean
Public UserEstupido       As Boolean

Public NoRes              As Boolean 'no cambiar la resolucion

Public RainBufferIndex    As Long
Public FogataBufferIndex  As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 0
Public Const tUs = 130

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87

Public NumEscudosAnims            As Integer

Public ArmasHerrero(0 To 100)     As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100)    As Integer
Public ObjSastre(0 To 100) As Integer
Public ObjHechizeria(0 To 100) As Integer
Public ObjHerreroMagico(0 To 100) As Integer

Public NumSastre As Integer
Public NumHechizeria As Integer
Public NumHerrero As Integer

Public UsaMacro                   As Boolean
Public CnTd                       As Byte
Public SecuenciaMacroHechizos     As Byte

Public Const Mensaje1             As String = "Estás muy cansado para lanzar este hechizo."
Public Const Mensaje2             As String = "No tienes suficientes puntos de magia para lanzar este hechizo."
Public Const Mensaje3             As String = "No tienes suficiente mana."
Public Const Mensaje4             As String = "No puedes lanzar hechizos porque estas muerto."
Public Const Mensaje5             As String = "En zona segura no puedes invocar criaturas."
Public Const Mensaje6             As String = "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos"
Public Const Mensaje7             As String = "No puedes atacar a ese npc."
Public Const Mensaje8             As String = "Debes quitarte el seguro para de poder atacar guardias"
Public Const Mensaje9             As String = "El npc es inmune a este hechizo."
Public Const Mensaje10            As String = "Este hechizo solo afecta NPCs que tengan amo."
Public Const Mensaje11            As String = "¡Has vuelto a ser visible!"
Public Const Mensaje12            As String = "¡¡Estas muerto!!"
Public Const Mensaje13            As String = "Estas muy lejos del usuario."
Public Const Mensaje14            As String = "Usuario inexistente."
Public Const Mensaje15            As String = "/salir cancelado."
Public Const Mensaje16            As String = "Has Terminado de meditar."
Public Const Mensaje17            As String = "No puedes moverte porque estas paralizado."
Public Const Mensaje18            As String = "¡¡No puedes atacar a nadie porque estas muerto!!."
Public Const Mensaje19            As String = "No puedes usar asi esta arma."
Public Const Mensaje20            As String = "¡¡Estas muerto!! Los muertos no pueden tomar objetos."
Public Const Mensaje21            As String = "Escribe /SEG para quitar el seguro"
Public Const Mensaje22            As String = "No puedes atacarte a ti mismo."
Public Const Mensaje23            As String = "Comienzas a Meditar"
Public Const Mensaje24            As String = "No puedo Cargar mas Objetos"
Public Const Mensaje25            As String = "¡Has Ganado 100 puntos de Experiencia!"
Public Const Mensaje26            As String = "Objetivo inválido."
Public Const Mensaje27            As String = "Estas demasiado lejos."
Public Const Mensaje28            As String = "Ya Estas Oculto."
Public Const Mensaje29            As String = "¡Primero selecciona el hechizo que quieres lanzar!"
Public Const Mensaje30            As String = "¡Primero tienes que seleccionar un personaje, hace click izquierdo sobre el."
Public Const Mensaje31            As String = "Primero hace click izquierdo sobre el personaje."
Public Const Mensaje32            As String = "El sacerdote no puede curarte debido a que estas demasiado lejos."
Public Const Mensaje33            As String = "Puedes Utilizar el Comando /HOGAR para ir a tu ciudad (Nix)."
Public Const Mensaje34            As String = "Puedes Utilizar el Comando /HOGAR para ir a tu ciudad (Ullathorpe)."
Public Const Mensaje35            As String = "Estas envenenado, si no te curas moriras."
Public Const Mensaje36            As String = "Has Sanado."
Public Const Mensaje37            As String = "Te estas concentrando, en 3 segundos comenzarás a meditar."
Public Const Mensaje38            As String = "AntiCheat> Tu Cliente es Valido, Gracias por jugar AOMania!!"
Public Const Mensaje39            As String = "Estás obstruyendo la via publica, muévete o seras encarcelado."
Public Const Mensaje40            As String = "Has sido resucitado!!"
Public Const Mensaje41            As String = "Has sido curado!!"
Public Const Mensaje42            As String = "Tu Clase, Genero o Raza, no puede usar este Objeto."

'[KEVIN]
Public Const MAX_BANCOINVENTORY_SLOTS = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
'[/KEVIN]

Public Const LoopAdEternum = 999

'Direcciones
Public Enum E_Heading

    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4

End Enum

'Objetos
Public Const MAX_INVENTORY_OBJS      As Integer = 10000
Public Const MAX_INVENTORY_SLOTS     As Byte = 25
Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 50
Public Const MAXHECHI                As Byte = 35

Public Const MAXSKILLPOINTS          As Byte = 100
Public Const FLAGORO                 As Byte = MAX_INVENTORY_SLOTS + 1
Public Const Fogata                  As Integer = 1521

Public Enum Skills

    Suerte = 1
    Magia = 2
    Robar = 3
    Tacticas = 4
    Armas = 5
    Meditar = 6
    Apuñalar = 7
    Ocultarse = 8
    Supervivencia = 9
    Talar = 10
    Comerciar = 11
    Defensa = 12
    Resistencia = 13
    Pesca = 14
    Mineria = 15
    Carpinteria = 16
    Herreria = 17
    Liderazgo = 18 ' NOTA: Solia decir "Curacion"
    Domar = 19
    Proyectiles = 20
    Wresterling = 21
    Navegacion = 22
    Sastreria = 23
    Recolectar = 24
    Hechiceria = 25
    Herrero = 26
End Enum


Public Enum eNPCType

    Comun = 0
    Revividor = 1
    GuardiaReal = 2
    Entrenador = 3
    Banquero = 4
    Armada = 5
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
    nQuest = 97
    Canjes = 98
    Creditos = 99

End Enum

' CATEGORIAS PRINCIPALES
Public Enum eObjType

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
    otHerramientas = 18
    otTELEPORT = 19
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otMontura = 66
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otMANCHAS = 35          'No se usa
    otPARAA = 36
    '[MaTeO 9]
    otAlas = 37
    '[/MaTeO 9]
    otPasaje = 41
    otVales = 50
    otPLATA = 77
    otCualquiera = 1000

End Enum

Public Const FundirMetal                           As Integer = 88

'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

Public Const MENSAJE_CRIATURA_FALLA_GOLPE          As String = "La criatura fallo el golpe!!!"
Public Const MENSAJE_CRIATURA_MATADO               As String = "La criatura te ha matado!!!"
Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO         As String = "Has rechazado el ataque con el escudo!!!"
Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO As String = "El usuario rechazo el ataque con su escudo!!!"
Public Const MENSAJE_FALLADO_GOLPE                 As String = "Has fallado el golpe!!!"
Public Const MENSAJE_SEGURO_ACTIVADO               As String = ">>SEGURO ACTIVADO<<"
Public Const MENSAJE_SEGURO_DESACTIVADO            As String = ">>SEGURO DESACTIVADO<<"
Public Const MENSAJE_PIERDE_NOBLEZA                As String = _
    "¡¡Has perdido puntaje de nobleza y ganado puntaje de criminalidad!! Si sigues ayudando a criminales te convertirás en uno de ellos y serás perseguido por las tropas de las ciudades."
Public Const MENSAJE_USAR_MEDITANDO                As String = "¡Estás meditando! Debes dejar de meditar para usar objetos."
Public Const MENSAJE_GOLPE_CABEZA                  As String = "¡¡La criatura te ha pegado en la cabeza por "
Public Const MENSAJE_GOLPE_BRAZO_IZQ               As String = "¡¡La criatura te ha pegado el brazo izquierdo por "
Public Const MENSAJE_GOLPE_BRAZO_DER               As String = "¡¡La criatura te ha pegado el brazo derecho por "
Public Const MENSAJE_GOLPE_PIERNA_IZQ              As String = "¡¡La criatura te ha pegado la pierna izquierda por "
Public Const MENSAJE_GOLPE_PIERNA_DER              As String = "¡¡La criatura te ha pegado la pierna derecha por "
Public Const MENSAJE_GOLPE_TORSO                   As String = "¡¡La criatura te ha pegado en el torso por "

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1                             As String = "¡¡"
Public Const MENSAJE_2                             As String = "!!"

Public Const MENSAJE_GOLPE_CRIATURA_1              As String = "¡¡Le has pegado a la criatura por "

Public Const MENSAJE_ATAQUE_FALLO                  As String = " te ataco y fallo!!"

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA         As String = " te ha pegado en la cabeza por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ      As String = " te ha pegado el brazo izquierdo por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER      As String = " te ha pegado el brazo derecho por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ     As String = " te ha pegado la pierna izquierda por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER     As String = " te ha pegado la pierna derecha por "
Public Const MENSAJE_RECIVE_IMPACTO_TORSO          As String = " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1             As String = "¡¡Le has pegado a "
Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA        As String = " en la cabeza por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ     As String = " en el brazo izquierdo por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER     As String = " en el brazo derecho por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ    As String = " en la pierna izquierda por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER    As String = " en la pierna derecha por "
Public Const MENSAJE_PRODUCE_IMPACTO_TORSO         As String = " en el torso por "

Public Const MENSAJE_TRABAJO_MAGIA                 As String = "Haz click sobre el objetivo..."
Public Const MENSAJE_TRABAJO_PESCA                 As String = "Haz click sobre el sitio donde quieres pescar..."
Public Const MENSAJE_TRABAJO_ROBAR                 As String = "Haz click sobre la victima..."
Public Const MENSAJE_TRABAJO_TALAR                 As String = "Haz click sobre el árbol..."
Public Const MENSAJE_TRABAJO_MINERIA               As String = "Haz click sobre el yacimiento..."
Public Const MENSAJE_TRABAJO_FUNDIRMETAL           As String = "Haz click sobre la fragua..."
Public Const MENSAJE_TRABAJO_PROYECTILES           As String = "Haz click sobre la victima..."

Public Const MENSAJE_ENTRAR_PARTY_1                As String = "Si deseas entrar en una party con "
Public Const MENSAJE_ENTRAR_PARTY_2                As String = ", escribe /entrarparty"

Public Const MENSAJE_NENE                          As String = "Cantidad de NPCs: "

'Inventario
Type Inventory

    ObjIndex As Integer
    Name As String
    GrhIndex As Long
    '[Alejo]: tipo de datos ahora es Long
    Amount As Long
    '[/Alejo]
    Equipped As Byte
    Valor As Long
    ObjType As Integer
    MinDef As Integer
    MaxDef As Integer
    MaxHit As Integer
    MinHit As Integer
    
End Type

Type NpCinV

    ObjIndex As Integer
    Name As String
    GrhIndex As Integer
    Amount As Integer
    Valor As Long
    ObjType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
    
End Type

Type CredInv
     
    Name As String
    GrhIndex As Long
    ObjIndex As Long
    Monedas As Long
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
    ObjType As Integer
     
End Type

Type CanjInv
      
    Name As String
    GrhIndex As Long
    ObjIndex As Long
    Monedas As Long
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
    ObjType As Integer
    Cantidad As Integer
    
End Type

Type tReputacion 'Fama del usuario

    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    
    Promedio As Long

End Type

Type tStats
    Nivel As Long
    MinExp As Long
    MaxExp As Long
    MinHP As Long
    MaxHP As Long
    MinMan As Long
    MaxMan As Long
    MinSta As Long
    MaxSta As Long
    Oro As Long
    Banco As Long
    SkillPoins As Long
End Type

Type tPos
    Map As String
    PosX As Integer
    PosY As Integer
End Type

Type tFaccion
   
    Armada As String
    Reenlistado As Byte
    Recompensas As Byte
    CiudadanosMatados As Long
    CriminalesMatados As Long
    FEnlistado As String
   
End Type

Type tEstadisticasUsu

    CiudadanosMatados As Long
    CriminalesMatados As Long
    AbbadonMatados As Long
    CleroMatados As Long
    TemplarioMatados As Long
    TinieblaMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
    Raza As String
    PuntosClan As Long
    Name As String
    Genero As String
    PuntosRetos As Long
    PuntosTorneos As Long
    PuntosDuelos As Long
    Stats As tStats
    pos As tPos
    ParticipoClan As Integer
    Faccion As tFaccion
    
End Type

Public Nombres                                    As Boolean

Public MixedKey                                   As Long

'User status vars
Global OtroInventario(1 To MAX_INVENTORY_SLOTS)   As Inventory

Public UserInventory(1 To MAX_INVENTORY_SLOTS)    As Inventory
Public UserHechizos(1 To MAXHECHI)                As Integer

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV
Public NPCInvDim                                  As Integer

Public CREDInventory(1 To MAX_NPC_INVENTORY_SLOTS) As CredInv
Public CREDInvDim As Integer

Public CANJInventory(1 To MAX_NPC_INVENTORY_SLOTS) As CanjInv
Public CANJInvDim As Integer


Public UserMeditar                                As Boolean
Public UserName                                   As String
Public UserPassword                               As String
Public UserMaxHP                                  As Long
Public UserMinHP                                  As Long
Public UserMaxMAN                                 As Integer
Public UserMinMAN                                 As Integer
Public UserMaxSTA                                 As Integer
Public UserMinSTA                                 As Integer
Public UserGLD                                    As Long
Public UserCreditos                             As Long
Public UserCanjes                                As Long
Public UserLvl                                    As Integer
Public UserPort                                   As Integer
Public UserServerIP                               As String
Public UserCanAttack                              As Integer
Public UserEstado                                 As Byte '0 = Vivo & 1 = Muerto
Public UserPasarNivel                             As Long
Public UserExp                                    As Long
Public UserReputacion                             As tReputacion
Public UserEstadisticas                           As tEstadisticasUsu
Public UserDescansar                              As Boolean
Public UserParalizado                             As Boolean
Public UserNavegando                              As Boolean
Public UserHogar                                  As String
Public UserClase                                  As String
Public UserSexo                                   As String
Public UserRaza                                   As String
Public UserEmail                                  As String
Public UserBanco                                  As String
Public UserPersonaje                           As String
Public UserFuerza As Byte
Public UserAgilidad As Byte
Public UserInteligencia As Byte
Public UserCarisma As Byte
Public UserConstitucion As Byte


Public NumUsers                                   As String
Public FPSFLAG                                    As Boolean
Public pausa                                      As Boolean

'<-------------------------NUEVO-------------------------->
Public Comerciando                                As Boolean
'<-------------------------NUEVO-------------------------->

Public Const NUMCIUDADES                          As Byte = 3
Public Const NUMSKILLS                            As Byte = 26
Public Const NUMATRIBUTOS                         As Byte = 5
Public Const NUMCLASES                            As Byte = 13
Public Const NUMRAZAS                             As Byte = 10

Public UserSkills(1 To NUMSKILLS)                 As Integer
Public SkillsNames(1 To NUMSKILLS)                As String

Public UserAtributos(1 To NUMATRIBUTOS)           As Integer
Public AtributosNames(1 To NUMATRIBUTOS)          As String

Public Ciudades(1 To NUMCIUDADES)                 As String
Public CityDesc(1 To NUMCIUDADES)                 As String

Public ListaRazas(1 To NUMRAZAS)                  As String
Public ListaClases(1 To NUMCLASES)                As String

Public SkillPoints                                As Integer
Public Alocados                                   As Integer
Public Flags()                                    As Integer

Public NoPuedeUsar                                As Boolean

'Barrin 30/9/03
Public UserPuedeRefrescar                         As Boolean
Public UsingSkill                                 As Integer

Public Enum E_MODO

    Normal = 1
    BorrarPj = 2
    CrearNuevoPj = 3
    Dados = 4
    RecuperarPass = 5

End Enum

Public EstadoLogin As E_MODO
   
Public Enum FxMeditar

    '    FXMEDITARCHICO = 4
    '    FXMEDITARMEDIANO = 5
    '    FXMEDITARGRANDE = 6
    '    FXMEDITARXGRANDE = 16
    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16

End Enum

'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer      As String 'Holds temp raw data from server
Public stxtbuffercmsg  As String 'Holds temp raw data from server
Public SendNewChar     As Boolean 'Used during login
Public Connected       As Boolean 'True when connected to server
Public DownloadingMap  As Boolean 'Currently downloading a map from server
Public UserMap         As Integer

'String contants
Public Const ENDC      As String * 1 = vbNullChar    'Endline character for talking with server
Public Const ENDL      As String * 2 = vbCrLf        'Holds the Endline character for textboxes

'Control
Public prgRun          As Boolean 'When true the program ends

'
'********** FUNCIONES API ***********
'

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
    ByVal lpKeyname As Any, _
    ByVal lpString As String, _
    ByVal lpFilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
    ByVal lpKeyname As Any, _
    ByVal lpdefault As String, _
    ByVal lpreturnedstring As String, _
    ByVal nSize As Long, _
    ByVal lpFilename As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Lista de cabezas
Public Type tIndiceCabeza

    Head(1 To 4) As Long

End Type

Public Type tIndiceCuerpo

    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer

End Type

Public Type tIndiceFx

    Animacion As Long
    OffsetX As Integer
    OffsetY As Integer

End Type

Public TimerPing(1 To 2) As Long

Public Type tSetupMods

    bTransparencia    As Byte
    bMusica   As Byte
    bSonido    As Byte
    bResolucion    As Byte
    bEjecutar As Byte
    bMover As Byte
    bPajaritos As Byte

End Type

Public AoSetup        As tSetupMods

Public VOLUMEN_FX     As Integer
Public VOLUMEN_MUSICA As Integer

Public SoundPajaritos As Boolean

Public StatusVerde    As Boolean
Public StatusAmarilla As Boolean

Public Type tMayor
      
    CiudadanoMaxNivel As String
    CriminalMaxNivel As String
    MaxCiudadano As String
    MaxCriminal As String
    OnlineCiudadano As Long
    OnlineCriminal As Long
    MaxOroOnline As String
    MaxOro As String

End Type

Public Mayores As tMayor

Public FloodStats As Integer

Public Const PARTYMAXMEMBER As Integer = 10

Public Type tParty
      
      Name As String
      MinHP As Integer
      MaxHP As Integer
      
End Type

Public PartyData(1 To PARTYMAXMEMBER) As tParty

Public MaxVerParty As Integer

Public Heads() As Integer

Public TimeChange As Byte
Public NameDay As String

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public ProcesoQuest As Byte

Public AodefConv As AoDefenderConverter
Public SuperClave As String

Public Type tClanPos
       X As Byte
       Y As Byte
End Type

Public ClanPos(1 To 10) As tClanPos
