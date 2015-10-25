Attribute VB_Name = "modDeclaraciones"
Option Explicit

Public Declare Sub mouse_event Lib "user32" _
(ByVal dwFlags As Long, ByVal dX As Long, _
ByVal dy As Long, ByVal cButtons As Long, _
ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Enum eChatType
    Norm = 1
    Guil = 2
    Komp = 3
    Priv = 4
    Glob = 5
    Yell = 6
    GM = 7
    GM_Yell = 8
End Enum

'Public MiniMapa1 As Integer
'Public MiniMapa2 As Integer
'Public MiniMapa3 As Integer
'Public MiniMapa4 As Integer

Public UserLogged As Boolean

'Objetos públicos
Public Dialogos As New clsDialogs
Public Audio As New clsAudio
Public Inventario As New clsGrapchicalInventory
Public Cinturon As New clsGrapchicalInventory

'Inventario de hechizos
Public Hechizos As New clsGrapchicalInventory

'Inventario con los Items que ofrece el npc
Public InvNpc As New clsGrapchicalInventory

'Inventario de companieros :D
Public Companieros As New clsGrapchicalInventory

'Inventarios de comercio con usuario
Public InvComUsu As New clsGrapchicalInventory 'Inventario del usuario visible en el comercio
Public InvOroComUsu(2) As New clsGrapchicalInventory 'Inventarios de oro (ambos usuarios)
Public InvOfferComUsu(1) As New clsGrapchicalInventory 'Inventarios de ofertas (ambos usuarios)

'Inventarios de herreria
Public Const MAX_LIST_Items As Byte = 4
Public InvLingosHerreria(1 To MAX_LIST_Items) As New clsGrapchicalInventory
Public InvMaderasCarpinteria(1 To MAX_LIST_Items) As New clsGrapchicalInventory

Public SurfaceDB As clsSurfaceManager   'No va new porque es una interfaz, el new se pone al decidir que clase de objeto es
Public CustomKeys As New clsCustomKeys
Public CustomMessages As New clsCustomMessages

Public incomingData As New clsByteQueue
Public outgoingData As New clsByteQueue

'The main timer of the game.
Public MainTimer As New clsTimer

'Sonidos
Public Const SND_PORTAL As String = "223.wav"
Public Const SND_FOGATA As String = "224.wav"

Public Const SND_CLICK As String = "click.wav"
Public Const SND_PASOS1 As String = "23.wav"
Public Const SND_PASOS2 As String = "24.wav"
Public Const SND_NAVEGANDO As String = "50.wav"
Public Const SND_OVER As String = "click2.wav"
Public Const SND_DICE As String = "cupdice.wav"
Public Const SND_LLUVIAINEND As String = "lluviainend.wav"
Public Const SND_LLUVIAOUTEND As String = "lluviaoutend.wav"

Public Const SND_HEARTBEAT As String = "217.wav" 'POCA VIDA

Public Const SND_PICKUP As String = "231.wav" 'AGARRAR

Public Const SND_PICKUP_GOLD As String = "227.wav" 'AGARRAR ORO

Public Const SND_DROP As String = "214.wav" 'TIRAR

'Public Const SND_DROP_GOLD As String = "228.wav" 'TIRAR ORO

Public Const SND_SELL As String = "225.wav" 'VENDER

Public Const SND_DRINK As String = "46.wav" 'BEBER

Public Const SND_AMANECER As String = "215.wav"
Public Const SND_NOCHE As String = "216.wav"

Public Const SND_ESCUDO As String = "37.wav" 'RECHAZA GOLPE CON ESCUDO

Public Const SND_SWING As String = "212.wav"        'USER FALLA
Public Const SND_SWING2 As String = "213.wav"       'USER FALLA 2
Public Const SND_NPCSWING As String = "2.wav"       'NPC FALLA

Public Const SND_APU As String = "221.wav"          'STAB

'Musica
Public Const SND_MUSIC_INTRO As Byte = 135 'MÚSICA CONECTAR/CREARPJ

Public Const SND_INTRO As String = "218.wav" 'SONIDO CONECTAR/CREARPJ?

'Constantes de intervalo
Public Const INT_MACRO_HECHIS As Integer = 2500
Public Const INT_MACRO_TRABAJO As Integer = 1500

Public Const INT_ATTACK As Integer = 1150
Public Const INT_ARROWS As Integer = 1150
Public Const INT_CAST_SPELL As Integer = 1150
Public Const INT_CAST_ATTACK As Integer = 1150
Public Const INT_WORK As Integer = 1000
Public Const INT_USEItemU As Integer = 400
Public Const INT_USEItemDCK As Integer = 150
Public Const INT_SENTRPU As Integer = 3000
Public Const INT_DROP As Integer = 1000
Public Const INT_BUY_SELL As Integer = 400
Public Const INT_PUB_MSG As Integer = 5000
Public Const INT_TALK As Integer = 1000
Public Const INT_MEDIT As Integer = 1000
Public Const INT_RAND_NAME As Integer = 1000

Public MacroBltIndex As Integer

Public Const CASPER_HEAD As Integer = 500
Public Const FRAGATA_FANTASMAL As Integer = 87

Public Type tColor
    r As Byte
    g As Byte
    B As Byte
End Type

Public ColoresPJ(0 To 50) As tColor

Public Const FX_MEDITARCHICO = 4
Public Const FX_MEDITARMEDIANO = 5
Public Const FX_MEDITARGRANDE = 6
Public Const FX_MEDITARXGRANDE = 16
Public Const FX_MEDITARXXGRANDE = 34

Public Type tServerInfo
    Ip As String
    Puerto As Integer
    Desc As String
    PassRecPort As Integer
End Type

Public CreandoGuilda As Boolean
Public GuildName As String

Public UserCiego As Boolean
Public UserEstupido As Boolean

Public Const GrhFile As String = "Grh.ind"

Public Const GrhComprimidos As Boolean = False

Public RainBufferIndex As Long
Public FogataBufferIndex As Long
Public PortalBufferIndex As Long
Public ApuBufferIndex As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 2000
Public Const tUs = 600

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87

Public NumEscudosAnims As Integer

Public ArmasHerrero() As tItemsConstruibles
Public ArmadurasHerrero() As tItemsConstruibles
Public ObjCarpintero() As tItemsConstruibles
Public CarpinteroMejorar() As tItemsConstruibles
Public HerreroMejorar() As tItemsConstruibles

Public UsaMacro As Boolean
Public CnTd As Byte

'Public DañoEnConsola As Boolean
'Public TransparenciaEnForms As Boolean

Public TradingUserName As String

Public Const LoopAdEternum As Integer = 999

'Direcciones
Public Enum eHeading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum
    
'Cantidad Maxima de objetos por slot del inventario
Public Const MaxInvObjs As Integer = 1000

'Cantidad Maxima de objetos por slot del cinturón
Public Const MaxBeltObjs As Integer = 50

'Cantidad de slots en el inventario
Public Const MaxInvSlots As Byte = 20
Public Const MaxBeltSlots As Byte = 4
Public Const MaxNpcInvSlots As Byte = 50
Public Const MaxSpellSlots As Byte = 35
Public Const MaxBankSlots As Byte = 30
Public Const MaxCompaSlots As Byte = 50
Public Const MaxPlataformSlots As Byte = 10

'Plataformas
Public Plataforma(1 To MaxPlataformSlots) As Integer

Public Banco(1 To MaxBankSlots) As tInv

Public Const MAXSKILLPOINTS As Byte = 100

Public Const MAXATRIBUTOS As Byte = 40

Public Const FLAGORO As Integer = MaxInvSlots

Public Const FOgata As Integer = 1521

Public Enum eClass
    Mage = 1    'Mago
    Cleric      'Clérigo
    Warrior     'Guerrero
    Assasin     'Asesino
    Thief       'Ladrón
    Bard        'Bardo
    Druid       'Druida
    Bandit      'Bandido
    Paladin     'Paladín
    Hunter      'Cazador
    Pirat       'Pirata
End Enum

Public Enum eCiudad
    cUllathorpe = 1
    cNix
    cBanderbill
    cLindos
    cArghal
End Enum

Enum eRaza
    Humano = 1
    Elfo
    ElfoOscuro
    Gnomo
    Enano
End Enum

Public Enum eSkill
    Magia = 1
    Robar = 2
    Tacticas = 3
    Armas = 4
    Meditar = 5
    Apuñalar = 6
    Ocultarse = 7
    Supervivencia = 8
    Talar = 9
    Defensa = 10
    Pescar = 11
    Minar = 12
    Carpinteria = 13
    Herreria = 14
    Liderazgo = 15
    Domar = 16
    Proyectiles = 17
    Wrestling = 18
    Navegacion = 19
End Enum

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Enum eGenero
    Hombre = 1
    Mujer
End Enum

Public Enum PlayerType
    User = 1
    Consejero = 2
    SemiDios = 4
    Dios = 8
    Admin = 10
    RoleMaster = 20
End Enum

Public Enum eObjType
    otUseOnce = 1
    otArma = 2
    otArmadura = 3
    otArbol = 4
    otGuita = 5
    otPuerta = 6
    otContenedor = 7
    otCartel = 8
    otLlave = 9
    otForo = 10
    otPocion = 11
    otBebida = 13
    otLeña = 14
    otFogata = 15
    otEscudo = 16
    otCasco = 17
    otAnillo = 18
    otTeleport = 19
    otYacimiento = 22
    otMineral = 23
    otPergamino = 24
    otInstrumento = 26
    otYunque = 27
    otFragua = 28
    otBarco = 31
    otFlecha = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otMancha = 35          'No se usa
    otArbolElfico = 36
    otPasaje = 37
    otCuerpoMuerto = 38
    otCinturon = 39
    otAlijo = 40
    otCualquiera = 1000
End Enum

'Tamaño del mapa
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

'Mensajes
Public Const MENSAJE_CRIATURA_FALLA_GOLPE As String = "¡La criatura falló el golpe!"
Public Const MENSAJE_CRIATURA_MATADO As String = "¡La criatura te mató!"
Public Const MENSAJE_rechazó_ATAQUE_ESCUDO As String = "¡Rechazaste el ataque con el escudo!"
Public Const MENSAJE_USUARIO_rechazó_ATAQUE_ESCUDO  As String = "¡El usuario rechazó el ataque con su escudo!"
Public Const MENSAJE_FALLADO_GOLPE As String = "¡Has fallado el golpe!"
Public Const MENSAJE_SEGURO_ACTIVADO As String = "Seguro activado."
Public Const MENSAJE_SEGURO_DESACTIVADO As String = "Seguro desactivado."

Public Const MENSAJE_SEGURO_RESU_ON As String = "Seguro de resurreccion activado"
Public Const MENSAJE_SEGURO_RESU_OFF As String = "Seguro de resurreccion desactivado"

Public Const MENSAJE_GOLPE_CABEZA As String = "¡La criatura te pegó en la cabeza por "
Public Const MENSAJE_GOLPE_BRAZO_IZQ As String = "¡La criatura te pegó en el brazo izquierdo por "
Public Const MENSAJE_GOLPE_BRAZO_DER As String = "¡La criatura te pegó en el brazo derecho por "
Public Const MENSAJE_GOLPE_PIERNA_IZQ As String = "¡La criatura te pegó en la pierna izquierda por "
Public Const MENSAJE_GOLPE_PIERNA_DER As String = "¡La criatura te pegó en la pierna derecha por "
Public Const MENSAJE_GOLPE_TORSO  As String = "¡La criatura te pegó en el torso por "

'MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1 As String = "¡"
Public Const MENSAJE_2 As String = "!"

Public Const MENSAJE_GOLPE_CRIATURA_1 As String = "¡Le pegaste a la criatura por "

Public Const MENSAJE_ATAQUE_falló As String = " te atacó y falló!"

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA As String = " te pegó en la cabeza por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ As String = " te pegó el brazo izquierdo por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER As String = " te pegó el brazo derecho por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ As String = " te pegó la pierna izquierda por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER As String = " te pegó la pierna derecha por "
Public Const MENSAJE_RECIVE_IMPACTO_TORSO As String = " te pegó en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1 As String = "¡Le pegaste a "
Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA As String = " en la cabeza por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ As String = " en el brazo izquierdo por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER As String = " en el brazo derecho por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ As String = " en la pierna izquierda por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER As String = " en la pierna derecha por "
Public Const MENSAJE_PRODUCE_IMPACTO_TORSO As String = " en el torso por "

Public Const MENSAJE_ENTRAR_PARTY_1 As String = "Si deseas entrar en una party con "
Public Const MENSAJE_ENTRAR_PARTY_2 As String = ", escribe /entrarparty"

Public Const MENSAJE_NENE As String = "Cantidad de NPCs: "

Public Const MENSAJE_FRAGSHOOTER_TE_HA_MATADO As String = "te mató!"
Public Const MENSAJE_FRAGSHOOTER_HAS_MATADO As String = "Mataste a"
Public Const MENSAJE_FRAGSHOOTER_HAS_GANADO As String = "Ganaste "
Public Const MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA As String = "puntos de experiencia."

'Inventario
Type tInv
    ObjIndex As Integer
    Name As String
    GrhIndex As Integer
    Amount As Long
    Valor As Long
    ObjType As eObjType
    MaxDef As Byte
    MinDef As Byte
    MaxHit As Integer
    MinHit As Integer
    PuedeUsar As Boolean
    Proyectil As Boolean
End Type

Public Inv(1 To MaxInvSlots) As tInv

Type tBelt
    ObjIndex As Integer
    Name As String
    GrhIndex As Integer
    Amount As Integer
    Valor As Long
    PuedeUsar As Boolean
End Type

Public Belt(1 To MaxBeltSlots) As tBelt

Type tCompa
    Nombre As String
    Online As Boolean
    Body As Integer
    Head As Integer
    CascoAnim As Integer
    ShieldAnim As Integer
    WeaponAnim As Integer
End Type

Public Compa(1 To MaxCompaSlots) As tCompa

'Hechizos Locura
Type tSpell
    Grh As Integer
    Nombre As String
    MinSkill As Boolean
    ManaRequerido As Integer
    StaRequerido As Integer
    NeedStaff As Integer
    PuedeLanzar As Boolean
End Type

Public Spell(1 To MaxSpellSlots) As tSpell

Type tNpcInv
    ObjIndex As Integer
    Name As String
    GrhIndex As Integer
    Amount As Integer
    Valor As Long
    ObjType As eObjType
    MaxDef As Byte
    MinDef As Byte
    MaxHit As Integer
    MinHit As Integer
    PuedeUsar As Boolean
End Type

Public NpcInv(1 To MaxNpcInvSlots) As tNpcInv

Type tUserStats
    Matados As Long
    Muertes As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
    Silencio As Long
End Type

Type tItemsConstruibles
    Name As String
    ObjIndex As Integer
    GrhIndex As Integer
    LinH As Integer
    LinP As Integer
    LinO As Integer
    Madera As Integer
    MaderaElfica As Integer
    Upgrade As Integer
    UpgradeName As String
    UpgradeGrhIndex As Integer
End Type

Public Nombres As Boolean

Public Meditando As Boolean
Public UserName As String
Public UserPassword As String
Public UserMaxHP As Integer
Public UserMinHP As Integer
Public UserMaxMan As Integer
Public UserMinMan As Integer
Public UserMaxSTA As Integer
Public UserMinSTA As Integer
Public UserMinSed As Byte
Public UserMinHam As Byte
Public UserGld As Long
Public UserLvl As Integer
Public UserMuerto As Boolean
Public UserPasarNivel As Long
Public UserExp As Long
Public NroItems As Byte
Public NroBeltItems As Byte
Public NroSpells As Byte
Public NroCompas As Byte
Public NroPlataformas As Byte
Public UserEstadisticas As tUserStats
Public Descansando As Boolean
Public FPSFLAG As Boolean
Public Pausa As Boolean
Public UserParalizado As Boolean
Public UserNavegando As Boolean

Public UserFuerza As Byte
Public UserAgilidad As Byte

Public HeadEqp As tInv
Public BodyEqp As tInv
Public LeftHandEqp As tInv
Public RightHandEqp As tInv
Public BeltEqp As tInv
Public RingEqp As tInv
Public Ship As tInv

Public Comerciando As Boolean
Public MirandoForo As Boolean
Public MirandoAsignarSkills As Boolean
Public MirandoEstadisticas As Boolean
Public MirandoParty As Boolean

Public UserClase As eClass
Public UserSexo As eGenero
Public UserRaza As eRaza
Public UserEmail As String

Public Const NUMCIUDADES As Byte = 5
Public Const NUMSKILLS As Byte = 19
Public Const NUMATRIBUTOS As Byte = 5
Public Const NUMCLASES As Byte = 12
Public Const NUMRAZAS As Byte = 5

Public UserSkills(1 To NUMSKILLS) As Byte
Public PorcentajeSkills(1 To NUMSKILLS) As Byte
Public SkillName(1 To NUMSKILLS) As String

Public UserAtributos(1 To NUMATRIBUTOS) As Byte
Public AtributosNames(1 To NUMATRIBUTOS) As String

Public Ciudades(1 To NUMCIUDADES) As String

Public ListaRazas(1 To NUMRAZAS) As String
Public ListaClases(1 To NUMCLASES) As String

Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer

Public UsingSkill As Byte

Public pingTime As Long

Public EsPartyLeader As Boolean
   
Public Enum EstadoLog
    Normal = 1
    Creado = 2 'en cliente
    Creando = 3
    BuscandoNombre = 4
    Recuperando = 5
End Enum

Public EstadoLogin As EstadoLog

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Sex
    eo_Raza
    eo_addGold
End Enum

'
'TRIGGERS
'
'@param NADA nada
'@param BAJOTECHO bajo techo
'@param EnPlataforma
'@param POSINVALIDA los npcs no pueden pisar tiles con este trigger
'@param ZONASEGURA no se puede robar o pelear desde este trigger
'@param ANTIPIQUETE
'@param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum eTrigger
    NADA = 0
    BAJOTECHO = 1
    EnPlataforma = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

Public LastParsedString As String

Public stxtbuffercmsg As String 'Holds temp raw data from server

Public prgRun As Boolean 'When true the program ends

Public ServerIP As String
Public ServerPort As Integer

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type

Public EsperandoLevel As Boolean

Public GuildNames() As String
Public GuildMembers() As String

'Hardcoded grhs and Items
Public Const Gld_Index As Integer = 12
Public Const Gld_Grh As Integer = 511

Public Const LH_Grh As Integer = 724
Public Const LP_Grh As Integer = 725
Public Const LO_Grh As Integer = 723

Public Const Wood_Grh As Integer = 550
Public Const Elf_Wood_Grh As Integer = 1999

Public picMouseIcon As Picture

'MINIMAPA
Public SupBMiniMap As DirectDrawSurface7
Public SupMiniMap As DirectDrawSurface7

Public Enum EstadoTiempo
    Amanecer = 1
    Mañana = 2
    Mediodía = 3
    Tarde = 4
    Anochecer = 5
    Noche = 6
End Enum

Public Tiempo As EstadoTiempo

Public DataPath As String
Public GrhPath As String
Public MapPath As String
Public MusicPath As String
Public SfxPath As String

Public UserBankGold As Long

Public SoundEffectsActivated As Boolean
Public MusicActivated As Boolean
Public ChangeResolution As Boolean
Public AlphaBActivated As Boolean

'Consola en render
Public Type ConsolaLoqui
    MensajeConsola(1 To 10) As String
    Color_Red(1 To 10) As Integer
    Color_Green(1 To 10) As Integer
    Color_Blue(1 To 10) As Integer
End Type
 
Public Consola As ConsolaLoqui

'Drag & Drop
'Public Enum BeginDrag
'None = 0
'MiInventario = 1
'InventarioNpc = 2
'End Enum

'Public DragType As BeginDrag

Public PicInvDragging As Boolean

Public InvSelSlot As Byte
Public BeltSelSlot As Byte
Public SpellSelSlot As Byte
Public CompaSelSlot As Byte
Public NpcInvSelSlot As Byte

Public TempSlot As Byte
Public BeltTempSlot As Byte
Public CompaTempSlot As Byte
Public NpcTempSlot As Byte

Public Tomando As Boolean

Public PuedeMacrear As Boolean

Public last_i As Long

Public BeltSlots() As Integer
