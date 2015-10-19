Attribute VB_Name = "Protocol"
Option Explicit

'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 255

'Auxiliar ByteQueue used as buffer to generate Messages not intended to be sent right away.
'Specially usefull to create a Message once and send it over to several clients.
Private auxiliarBuffer As New clsByteQueue

Public PacketName(0 To 255) As String

Private Enum ConnectPacketID
    LoginChar
    LoginNewChar
    RecoverChar
    KillChar
    RequestRandomName
End Enum

Public Enum GMPacketID
'GM commands
    SearchObj               '/B OBJETO
    GMMessage               '/GMSG
    ShowName                '/NAME
    GoNearby                '/IRCERCA
    comment                 '/REM
    ServerTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpMeToTarget          '/TELEPLOC
    WarpChar                '/T
    Silence                 '/SILENCIAR
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/I
    Invisible               '/INVI
    GMPanel                 '/P
    RequestUserList         'LISTUSU
    Working                 '/TRABAJANDO
    Hiding                  '/OCULTANDO
    Jail                    '/CARCEL
    KillNpc                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    RequestCharInfo         '/INFO
    RequestCharStats        '/STAT
    RequestCharGold         '/BAL
    RequestCharInventory    '/INV
    RequestCharBank         '/BOV
    RequestCharSkills       '/SKILLS
    ReviveChar              '/R
    OnlineGM                '/ONGM
    OnlineMap               '/ONMAP
    Kick                    '/E
    Execute                 '/EJECUTAR
    BanChar                 '/BAN
    UnbanChar               '/UNBAN
    NpcFollow               '/FOLLOW
    SummonChar              '/S
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNpcInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONGUILDA
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    RainToggle              '/LLUVIA
    Weather
    SetCharDescription      '/SETDESC
    ForceMP3ToMap           '/FORCEMP3MAP
    ForceWAVEToMap          '/FORCEWAVMAP
    TalkAsNpc               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    MakeDumb                '/ESTUPIDO
    MakeDumbNoMore          '/NOESTUPIDO
    dumpIPTables            '/DUMPSECURITY
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSGUILDA
    GuildBan                '/BANGUILDA
    BanIP                   '/BANIP
    UnbanIP                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ForceMP3All             '/FORCEMP3
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    TileBlockedToggle       '/BLOQ
    KillNpcNoRespawn        '/M
    KillAllNearbyNpcs       '/MASSKILL
    LastIP                  '/LASTIP
    ChangeMOTD              '/MOTDCAMBIA
    SetMOTD                 'ZMOTD
    SystemMessage           '/SMSG
    CreateNpc               '/ACC
    CreateNpcWithRespawn    '/RACC
    ServerOpenToUsersToggle '/HABILITAR
    TurnOffServer           '/APAGAR
    RemoveCharFromGuild     '/RAJARCLAN
    RequestCharMail         '/LASTEMAIL
    AlterPassword           '/APASS
    AlterMail               '/AEMAIL
    AlterName               '/ANAME
    ToggleCentinelActivated '/CENTINELAACTIVADO
    DoBackUp                '/DOBACKUP
    ShowGuildMessages       '/SHOWCMSG
    SaveMap                 '/GUARDARMAPA
    ChangeMapInfoPK         '/MODMAPINFO PK
    ChangeMapInfoBackup     '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand       '/MODMAPINFO TERRENO
    ChangeMapInfoZone       '/MODMAPINFO ZONA
    ChangeMapInfoStealNpc   '/MODMAPINFO ROBONPC
    ChangeMapInfoNoOcultar  '/MODMAPINFO OCULTARSINEFECTO
    ChangeMapInfoNoInvocar  '/MODMAPINFO INVOCARSINEFECTO
    SaveChars               '/GUARDAR /GRABAR /G
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    KickAllChars            '/ETODOSPJS
    ReloadNpcs              '/RELOADNpcS
    ReloadServidorIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjs           '/RELOADOBJ
    Restart                 '/REINICIAR
    ChatColor               '/ChatCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/Slot
    SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
End Enum

Public Enum ServerPacketID
'GM commands
    SpawnList               'SPL
    ShowSOSForm             'MSOS
    ShowMOTDEditionForm     'ZMOTD
    ShowGMPanelForm         'ABPANEL
    UserNameList            'LISTUSU

    Logged                  'LOGGED
    RandomName
    NavigateToggle          'NAVEG
    UserCommerceInit        'INITCOMUSU
    UserCommerceEnd         'FINCOMUSUOK
    UserOfferConfirm
    CommerceChat
    CharSwing
    NpcKillUser             '6
    BlockedWithShield
    BlockedWithShieldOther
    UserSwing
    ResuscitationSafeOn
    ResuscitationSafeOff
    UpdateSta
    UpdateMana
    UpdateHP
    UpdateGold
    UpdateExp
    ChangeMap
    PosUpdate
    Damage
    UserDamaged
    ChatOverHead
    DeleteChatOverHead
    ChatNormal
    'ChatGM
    ChatGuild
    ChatCompa
    ChatPrivate
    ChatGlobal
    ConsoleMsg
    ShowMessageBox
    CharCreate
    NpcCharCreate
    CharRemove
    CharChangeNick
    CharMove
    ForceCharMove
    CharChange
    ChangeCharHeading
    ObjCreate
    ObjDelete
    BlockPosition
    PlayMP3
    PlayWave
    GuildList
    AreaChanged
    PauseToggle
    RainToggle
    Weather
    CreateFX
    CreateCharFX
    UpdateUserStats
    SlotMenosUno
    Inventory
    InventorySlot
    BeltInv
    BeltSlot
    Bank
    BankSlot
    NpcInventory
    NpcInventorySlot
    Spells
    SpellSlot
    Compas
    AddCompa
    QuitarCompa
    CompaConnected
    CompaDisconnected
    Attributes
    UserPlatforms
    BlacksmithWeapons
    BlacksmithArmors
    CarpenterObjs
    RestOK
    ErrorMsg
    Blind
    Dumb
    ShowSignal
    UpdateHungerAndThirst
    MiniStats
    SkillUp
    LevelUp
    SetInvisible
    SetParalized
    BlindNoMore
    DumbNoMore
    Skills
    FreeSkills
    TrainerCreatureList
    GuildNews
    OfferDetails
    AlianceProposalsList
    PeaceProposalsList
    CharInfo
    GuildLeaderInfo
    GuildMemberInfo
    GuildDetails
    ShowGuildFundationForm
    ShowUserRequest
    ChangeUserTradeSlot
    Pong
    UpdateTagAndStatus
    Population
    AnimAttack
    CharMeditate
    ShowPartyForm
    UpdateStrenghtAndDexterity
    UpdateStrenght
    UpdateDexterity
    StopWorking
    CancelOfferItem
End Enum

Private Enum ClientPacketID
    connect
    Online
    Talk                    ';
    CompaMessage
    PrivateMessage
    DeleteChat
    WalkNorth
    WalkEast
    WalkSouth
    WalkWest
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    ResuscitationSafeToggle
    RequestGuildLeaderInfo  'GLINFO
    RequestAttributes        'ATR
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    UserCommerceConfirm
    CommerceChat
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    MoveInvSlot
    MoveBeltSlot
    MoveBankSlot
    MoveSpellSlot
    Drop                    'TI
    DropGold
    LeftClick               'LC
    RightClick
    DoubleClick             'RC
    Work                    'UK
    UseSpellMacro           'UMH
    UseItem
    UseBeltItem
    CraftBlacksmith         'CNS
    CraftCarpenter          'CNC
    WorkLeftClick           'WLC
    CreateNewGuild          'CIG
    CastSpell
    SpellInfo               'INFS
    EquipItem
    UnEquipItem
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    BankDepositItem
    GuildDescUpdate
    UserCommerceOffer       'OFRECER
    GuildAcceptPeace        'ACEPPEAT
    GuildRejectAlliance     'RECPALIA
    GuildRejectPeace        'RECPPEAT
    GuildAcceptAlliance     'ACEPALIA
    GuildOfferPeace         'PEACEOFF
    GuildOfferAlliance      'ALLIEOFF
    GuildAllianceDetails    'ALLIEDET
    GuildPeaceDetails       'PEACEDET
    GuildRequestJoinerInfo  'ENVCOMEN
    GuildAlliancePropList   'ENVALPRO
    GuildPeacePropList      'ENVPROPP
    GuildDeclareWar         'DECGUERR
    GuildAcceptNewMember    'ACEPTARI
    GuildRejectNewMember    'RECHAZAR
    GuildKickMember         'ECHARCLA
    GuildUpdateNews         'ACTGNEWS
    GuildMemberInfo         '1HRINFO<
    GuildOpenElections      'ABREELEC
    GuildRequestMembership  'SOLICITUD
    GuildRequestDetails     'CLANDETAILS
    Quit                    '/SALIR
    GuildLeave              '/DEJARGUILDA
    RequestAccountState     '/BALANCE
    PetStand                '/QUIETO
    PetFollow               '/SEGUIR
    ReleasePet              '/LIBERAR
    Rest                    '/DESCANSAR
    Meditate                '/MEDITAR
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    RequestMOTD             '/MOTD
    UpTime                  '/UPTIME
    PartyLeave              '/SALIRPARTY
    PartyCreate             '/CREARPARTY
    PartyJoin               '/PARTY
    Inquiry                 '/ENCUESTA (with no params)
    GuildMessage            '/CMSG
    PartyMessage            '/PMSG
    CentinelReport          '/CENTINELA
    GuildOnline             '/ONGUILDA
    PartyOnline             '/ONLINEPARTY
    RoleMasterRequest       '/ROL
    GMRequest               '/GM
    BugReport               '/BUG
    ChangeDescription       '/DESC
    GuildVote               '/VOTO
    Punishments             '/PENAS
    ChangePassword          '/clave
    Gamble                  '/APOSTAR
    InquiryVote             '/ENCUESTA (with parameters)
    BankExtractGold         '/RETIRAR (with arguments)
    BankDepositGold         '/DEPOSITAR
    Denounce                '/DENUNCIAR
    GuildFundate            '/FUNDARGUILDA
    PartyKick               '/EPARTY
    PartySetLeader          '/PARTYLIDER
    PartyAcceptMember       '/ACCEPTPARTY
    Ping                    '/PING
    
    RequestPartyForm
    ItemUpgrade
    GMCommand
    InitCrafting
    ShowGuildNews
    Consultation
    PublicMessage
    AuctionCreate
    AuctionBid
    AuctionView
    AniadirCompaniero
    EliminarCompaniero
    Home
    PlatformTeleport
End Enum

Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_TALKGM
    FONTTYPE_YELL
    FONTTYPE_YELLGM
    FONTTYPE_PUBLICMESSAGE
    FONTTYPE_COMPAMESSAGE
    FONTTYPE_PRIVATEMESSAGE
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CONSE
    FONTTYPE_DIOS
    FONTTYPE_NUMERO
    FONTTYPE_HABILIDAD
    FONTTYPE_HECHIZO
End Enum

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

Public Sub HandleIncomingData(ByVal UserIndex As Integer)

On Error Resume Next
    Dim PacketID As Byte
    
    PacketID = UserList(UserIndex).incomingData.PeekByte()
        
    'Does the packet requires a logged user??
    If PacketID <> ClientPacketID.connect Then
        
        'Is the user actually logged?
        If Not UserList(UserIndex).flags.Logged Then
            Call CloseSocket(UserIndex)
            Exit Sub
        
        'He is logged. Reset idle counter if id is valid.
        ElseIf PacketID <= LAST_CLIENT_PACKET_ID Then
            UserList(UserIndex).Counters.IdleCount = 0
        End If
        
    ElseIf PacketID <= LAST_CLIENT_PACKET_ID Then
        UserList(UserIndex).Counters.IdleCount = 0
        
        'Is the user logged?
        If UserList(UserIndex).flags.Logged Then
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If
    
    'Ante cualquier paquete, pierde la proteccion de ser atacado.
    UserList(UserIndex).flags.NoPuedeSerAtacado = False
    
    Select Case PacketID
            
        Case ClientPacketID.connect
            Call HandleConnect(UserIndex)
        
        Case ClientPacketID.Online
            Call HandleOnline(UserIndex)
    
        Case ClientPacketID.Talk                    ';
            Call HandleTalk(UserIndex)
                
        Case ClientPacketID.CompaMessage
            Call HandleCompaMessage(UserIndex)
    
        Case ClientPacketID.PrivateMessage
            Call HandlePrivateMessage(UserIndex)
        
        Case ClientPacketID.DeleteChat
            Call HandleDeleteChat(UserIndex)
        
        Case ClientPacketID.WalkNorth
            Call HandleWalk(UserIndex, eHeading.NORTH)
        
        Case ClientPacketID.WalkEast
            Call HandleWalk(UserIndex, eHeading.EAST)
        
        Case ClientPacketID.WalkSouth
            Call HandleWalk(UserIndex, eHeading.SOUTH)
        
        Case ClientPacketID.WalkWest
            Call HandleWalk(UserIndex, eHeading.WEST)
        
        Case ClientPacketID.RequestPositionUpdate   'RPU
            Call HandleRequestPositionUpdate(UserIndex)
        
        Case ClientPacketID.Attack                  'AT
            Call HandleAttack(UserIndex)
        
        Case ClientPacketID.PickUp                  'AG
            Call HandlePickUp(UserIndex)
                
        Case ClientPacketID.ResuscitationSafeToggle
            Call HandleResuscitationToggle(UserIndex)
        
        Case ClientPacketID.RequestGuildLeaderInfo  'GLINFO
            Call HandleRequestGuildLeaderInfo(UserIndex)
        
        Case ClientPacketID.RequestAttributes        'ATR
            Call HandleRequestAttributes(UserIndex)
        
        Case ClientPacketID.RequestSkills           'ESKI
            Call HandleRequestSkills(UserIndex)
        
        Case ClientPacketID.RequestMiniStats        'FEST
            Call HandleRequestMiniStats(UserIndex)
        
        Case ClientPacketID.CommerceEnd             'FINCOM
            Call HandleCommerceEnd(UserIndex)
        
        Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
            Call HandleUserCommerceEnd(UserIndex)
        
        Case ClientPacketID.BankEnd                 'FINBAN
            Call HandleBankEnd(UserIndex)
        
        Case ClientPacketID.UserCommerceOk          'COMUSUOK
            Call HandleUserCommerceOk(UserIndex)
        
        Case ClientPacketID.UserCommerceReject      'COMUSUNO
            Call HandleUserCommerceReject(UserIndex)
        
        Case ClientPacketID.Drop                    'TI
            Call HandleDrop(UserIndex)
        
        Case ClientPacketID.DropGold
            Call HandleDropGold(UserIndex)
        
        Case ClientPacketID.LeftClick               'LC
            Call HandleLeftClick(UserIndex)
        
        Case ClientPacketID.RightClick
            Call HandleRightClick(UserIndex)
        
        Case ClientPacketID.DoubleClick             'RC
            Call HandleDoubleClick(UserIndex)
        
        Case ClientPacketID.Work                    'UK
            Call HandleWork(UserIndex)
        
        Case ClientPacketID.UseSpellMacro           'UMH
            Call HandleUseSpellMacro(UserIndex)
        
        Case ClientPacketID.UseItem                 'USA
            Call HandleUseItem(UserIndex)
        
        Case ClientPacketID.UseBeltItem
            Call HandleUseBeltItem(UserIndex)
        
        Case ClientPacketID.CraftBlacksmith         'CNS
            Call HandleCraftBlacksmith(UserIndex)
        
        Case ClientPacketID.CraftCarpenter          'CNC
            Call HandleCraftCarpenter(UserIndex)
        
        Case ClientPacketID.WorkLeftClick           'WLC
            Call HandleWorkLeftClick(UserIndex)
        
        Case ClientPacketID.CreateNewGuild          'CIG
            Call HandleCreateNewGuild(UserIndex)
        
        Case ClientPacketID.CastSpell
            Call HandleCastSpell(UserIndex)
            
        Case ClientPacketID.SpellInfo               'INFS
            Call HandleSpellInfo(UserIndex)
        
        Case ClientPacketID.EquipItem
            Call HandleEquipItem(UserIndex)
        
        Case ClientPacketID.UnEquipItem
            Call HandleUnEquipItem(UserIndex)
        
        Case ClientPacketID.ChangeHeading           'CHEA
            Call HandleChangeHeading(UserIndex)
        
        Case ClientPacketID.ModifySkills            'SKSE
            Call HandleModifySkills(UserIndex)
        
        Case ClientPacketID.Train                   'ENTR
            Call HandleTrain(UserIndex)
        
        Case ClientPacketID.CommerceBuy             'COMP
            Call HandleCommerceBuy(UserIndex)
        
        Case ClientPacketID.BankExtractItem         'RETI
            Call HandleBankExtractItem(UserIndex)
        
        Case ClientPacketID.CommerceSell            'VEND
            Call HandleCommerceSell(UserIndex)
        
        Case ClientPacketID.BankDepositItem
            Call HandleBankDepositItem(UserIndex)
                
        Case ClientPacketID.MoveInvSlot
            Call HandleMoveInvSlot(UserIndex)
            
        Case ClientPacketID.MoveBeltSlot
            Call HandleMoveBeltSlot(UserIndex)
            
        Case ClientPacketID.MoveSpellSlot
            Call HandleMoveSpellSlot(UserIndex)
            
        Case ClientPacketID.MoveBankSlot
            Call HandleMoveBankSlot(UserIndex)
        
        Case ClientPacketID.GuildDescUpdate
            Call HandleGuildDescUpdate(UserIndex)
        
        Case ClientPacketID.UserCommerceOffer       'OFRECER
            Call HandleUserCommerceOffer(UserIndex)
        
        Case ClientPacketID.GuildAcceptPeace        'ACEPPEAT
            Call HandleGuildAcceptPeace(UserIndex)
        
        Case ClientPacketID.GuildRejectAlliance     'RECPALIA
            Call HandleGuildRejectAlliance(UserIndex)
        
        Case ClientPacketID.GuildRejectPeace        'RECPPEAT
            Call HandleGuildRejectPeace(UserIndex)
        
        Case ClientPacketID.GuildAcceptAlliance     'ACEPALIA
            Call HandleGuildAcceptAlliance(UserIndex)
        
        Case ClientPacketID.GuildOfferPeace         'PEACEOFF
            Call HandleGuildOfferPeace(UserIndex)
        
        Case ClientPacketID.GuildOfferAlliance      'ALLIEOFF
            Call HandleGuildOfferAlliance(UserIndex)
        
        Case ClientPacketID.GuildAllianceDetails    'ALLIEDET
            Call HandleGuildAllianceDetails(UserIndex)
        
        Case ClientPacketID.GuildPeaceDetails       'PEACEDET
            Call HandleGuildPeaceDetails(UserIndex)
        
        Case ClientPacketID.GuildRequestJoinerInfo  'ENVCOMEN
            Call HandleGuildRequestJoinerInfo(UserIndex)
        
        Case ClientPacketID.GuildAlliancePropList   'ENVALPRO
            Call HandleGuildAlliancePropList(UserIndex)
        
        Case ClientPacketID.GuildPeacePropList      'ENVPROPP
            Call HandleGuildPeacePropList(UserIndex)
        
        Case ClientPacketID.GuildDeclareWar         'DECGUERR
            Call HandleGuildDeclareWar(UserIndex)
                
        Case ClientPacketID.GuildAcceptNewMember    'ACEPTARI
            Call HandleGuildAcceptNewMember(UserIndex)
        
        Case ClientPacketID.GuildRejectNewMember    'RECHAZAR
            Call HandleGuildRejectNewMember(UserIndex)
        
        Case ClientPacketID.GuildKickMember         'ECHARCLA
            Call HandleGuildKickMember(UserIndex)
        
        Case ClientPacketID.GuildUpdateNews         'ACTGNEWS
            Call HandleGuildUpdateNews(UserIndex)
        
        Case ClientPacketID.GuildMemberInfo         '1HRINFO<
            Call HandleGuildMemberInfo(UserIndex)
        
        Case ClientPacketID.GuildOpenElections      'ABREELEC
            Call HandleGuildOpenElections(UserIndex)
        
        Case ClientPacketID.GuildRequestMembership  'SOLICITUD
            Call HandleGuildRequestMembership(UserIndex)
        
        Case ClientPacketID.GuildRequestDetails     'CLANDETAILS
            Call HandleGuildRequestDetails(UserIndex)
                          
        Case ClientPacketID.Quit                    '/SALIR
            Call HandleQuit(UserIndex)
        
        Case ClientPacketID.GuildLeave              '/DEJARGUILDA
            Call HandleGuildLeave(UserIndex)
        
        Case ClientPacketID.RequestAccountState     '/BALANCE
            Call HandleRequestAccountState(UserIndex)
        
        Case ClientPacketID.PetStand                '/QUIETO
            Call HandlePetStand(UserIndex)
        
        Case ClientPacketID.PetFollow               '/SEGUIR
            Call HandlePetFollow(UserIndex)
        
        Case ClientPacketID.ReleasePet              '/LIBERAR
            Call HandleReleasePet(UserIndex)
        
        Case ClientPacketID.Rest                    '/DESCANSAR
            Call HandleRest(UserIndex)
        
        Case ClientPacketID.Meditate                '/MEDITAR
            Call HandleMeditate(UserIndex)
               
        Case ClientPacketID.Help                    '/AYUDA
            Call HandleHelp(UserIndex)
        
        Case ClientPacketID.RequestStats            '/EST
            Call HandleRequestStats(UserIndex)
        
        Case ClientPacketID.CommerceStart           '/COMERCIAR
            Call HandleCommerceStart(UserIndex)
        
        Case ClientPacketID.BankStart               '/BOVEDA
            Call HandleBankStart(UserIndex)
                
        Case ClientPacketID.RequestMOTD             '/MOTD
            Call HandleRequestMOTD(UserIndex)
        
        Case ClientPacketID.UpTime                  '/UPTIME
            Call HandleUpTime(UserIndex)
        
        Case ClientPacketID.PartyLeave              '/SALIRPARTY
            Call HandlePartyLeave(UserIndex)
        
        Case ClientPacketID.PartyCreate             '/CREARPARTY
            Call HandlePartyCreate(UserIndex)
        
        Case ClientPacketID.PartyJoin               '/PARTY
            Call HandlePartyJoin(UserIndex)
        
        Case ClientPacketID.Inquiry                 '/ENCUESTA ( with no params )
            Call HandleInquiry(UserIndex)
        
        Case ClientPacketID.GuildMessage            '/CMSG
            Call HandleGuildMessage(UserIndex)
        
        Case ClientPacketID.PartyMessage            '/PMSG
            Call HandlePartyMessage(UserIndex)
        
        Case ClientPacketID.CentinelReport          '/CENTINELA
            Call HandleCentinelReport(UserIndex)
        
        Case ClientPacketID.GuildOnline             '/ONLINECLAN
            Call HandleGuildOnline(UserIndex)
        
        Case ClientPacketID.PartyOnline             '/ONLINEPARTY
            Call HandlePartyOnline(UserIndex)
                
        Case ClientPacketID.RoleMasterRequest       '/ROL
            Call HandleRoleMasterRequest(UserIndex)
        
        Case ClientPacketID.GMRequest               '/GM
            Call HandleGMRequest(UserIndex)
        
        Case ClientPacketID.BugReport               '/BUG
            Call HandleBugReport(UserIndex)
        
        Case ClientPacketID.ChangeDescription       '/DESC
            Call HandleChangeDescription(UserIndex)
        
        Case ClientPacketID.GuildVote               '/VOTO
            Call HandleGuildVote(UserIndex)
        
        Case ClientPacketID.Punishments             '/PENAS
            Call HandlePunishments(UserIndex)
        
        Case ClientPacketID.ChangePassword          '/CONTRASEÑA
            Call HandleChangePassword(UserIndex)
        
        Case ClientPacketID.Gamble                  '/APOSTAR
            Call HandleGamble(UserIndex)
        
        Case ClientPacketID.InquiryVote             '/ENCUESTA ( with parameters )
            Call HandleInquiryVote(UserIndex)
                
        Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
            Call HandleBankExtractGold(UserIndex)
        
        Case ClientPacketID.BankDepositGold         '/DEPOSITAR
            Call HandleBankDepositGold(UserIndex)
        
        Case ClientPacketID.Denounce                '/DENUNCIAR
            Call HandleDenounce(UserIndex)
        
        Case ClientPacketID.GuildFundate            '/FUNDARGUILDA
            Call HandleGuildFundate(UserIndex)
                        
        Case ClientPacketID.PartyKick               '/EPARTY
            Call HandlePartyKick(UserIndex)
        
        Case ClientPacketID.PartySetLeader          '/PARTYLIDER
            Call HandlePartySetLeader(UserIndex)
        
        Case ClientPacketID.PartyAcceptMember       '/ACCEPTPARTY
            Call HandlePartyAcceptMember(UserIndex)
        
        Case ClientPacketID.Ping                    '/PING
            Call HandlePing(UserIndex)
        
        Case ClientPacketID.RequestPartyForm
            Call HandlePartyForm(UserIndex)
            
        Case ClientPacketID.ItemUpgrade
            Call HandleItemUpgrade(UserIndex)
            
        Case ClientPacketID.GMCommand
            Call HandleGMCommand(UserIndex)
            
        Case ClientPacketID.InitCrafting
            Call HandleInitCrafting(UserIndex)
            
        Case ClientPacketID.ShowGuildNews
            Call HandleShowGuildNews(UserIndex)
            
        Case ClientPacketID.Consultation
            Call HandleConsultation(UserIndex)
            
        Case ClientPacketID.PublicMessage
            Call HandlePublicMessage(UserIndex)
            
        Case ClientPacketID.AuctionCreate
            Call HandleAuctionCreate(UserIndex)
            
        Case ClientPacketID.AuctionBid
            Call HandleAuctionBid(UserIndex)
        
        Case ClientPacketID.AuctionView
            Call HandleAuctionView(UserIndex)

        Case ClientPacketID.AniadirCompaniero
            Call HandleAniadirCompaniero(UserIndex)
            
        Case ClientPacketID.EliminarCompaniero
            Call HandleEliminarCompaniero(UserIndex)
            
        Case ClientPacketID.Home
            Call HandleHome(UserIndex)
            
        Case ClientPacketID.PlatformTeleport
            Call HandlePlatformTeleport(UserIndex)

        Case Else
            'ERROR: Abort!
            Call CloseSocket(UserIndex)
    End Select
    
    'Done with this packet, move on to next one or send everything if no more packets found
    If UserList(UserIndex).incomingData.length > 0 And Err.Number = 0 Then
        Err.Clear
        Call HandleIncomingData(UserIndex)
    
    ElseIf Err.Number > 0 And Not Err.Number = UserList(UserIndex).incomingData.NotEnoughDataErrCode Then
        'An error ocurred, log it and kick player.
        Call LogError("Error: " & Err.Number & " [" & Err.description & "] " & " Source: " & Err.source & _
                        vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & _
                        vbTab & " LastDllError: " & Err.LastDllError & vbTab & _
                        " - UserIndex: " & UserIndex & " - producido al manejar el paquete: " & CStr(PacketID))
        Call CloseSocket(UserIndex)
    
    Else
        'Flush buffer - send everything that has been written
        Call FlushBuffer(UserIndex)
    End If
End Sub

Private Sub HandleConnect(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    
    With UserList(UserIndex)
        Call .incomingData.ReadByte

        Select Case .incomingData.PeekByte
            Case ConnectPacketID.LoginChar
                Call HandleLoginChar(UserIndex)
            
            Case ConnectPacketID.LoginNewChar
                Call HandleLoginNewChar(UserIndex)
            
            Case ConnectPacketID.RecoverChar
                Call HandleRecoverChar(UserIndex)
            
            Case ConnectPacketID.KillChar
                Call HandleKillChar(UserIndex)
                
            Case ConnectPacketID.RequestRandomName
                Call HandleRequestRandomName(UserIndex)
        End Select
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en HandleConnect. Error: " & Err.Number & " - " & Err.description & _
                  ". Paquete: " & Command)
                  
End Sub

Private Sub HandleGMCommand(ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    Dim Command As Byte
    
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        Command = .incomingData.PeekByte

        Select Case Command
            'GM Messages
            Case GMPacketID.SearchObj               '/B [OBJETO]
                Call HandleSearchObj(UserIndex)
    
            Case GMPacketID.GMMessage               '/GMSG
                Call HandleGMMessage(UserIndex)
            
            Case GMPacketID.ShowName                '/NAME
                Call HandleShowName(UserIndex)
            
            Case GMPacketID.GoNearby                '/IRCERCA
                Call HandleGoNearby(UserIndex)
            
            Case GMPacketID.comment                 '/REM
                Call HandleComment(UserIndex)
            
            Case GMPacketID.ServerTime              '/HORA
                Call HandleServerTime(UserIndex)
            
            Case GMPacketID.Where                   '/DONDE
                Call HandleWhere(UserIndex)
            
            Case GMPacketID.CreaturesInMap          '/NENE
                Call HandleCreaturesInMap(UserIndex)
            
            Case GMPacketID.WarpMeToTarget          '/TELEPLOC
                Call HandleWarpMeToTarget(UserIndex)
            
            Case GMPacketID.WarpChar                '/TELEP
                Call HandleWarpChar(UserIndex)
            
            Case GMPacketID.Silence                 '/SILENCIAR
                Call HandleSilence(UserIndex)
            
            Case GMPacketID.SOSShowList             '/SHOW SOS
                Call HandleSOSShowList(UserIndex)
            
            Case GMPacketID.SOSRemove               'SOSDONE
                Call HandleSOSRemove(UserIndex)
            
            Case GMPacketID.GoToChar                '/I
                Call HandleGoToChar(UserIndex)
            
            Case GMPacketID.Invisible               '/INVI
                Call HandleInvisible(UserIndex)
            
            Case GMPacketID.GMPanel                 '/P
                Call HandleGMPanel(UserIndex)
            
            Case GMPacketID.RequestUserList         'LISTUSU
                Call HandleRequestUserList(UserIndex)
            
            Case GMPacketID.Working                 '/TRABAJANDO
                Call HandleWorking(UserIndex)
            
            Case GMPacketID.Hiding                  '/OCULTANDO
                Call HandleHiding(UserIndex)
            
            Case GMPacketID.Jail                    '/CARCEL
                Call HandleJail(UserIndex)
            
            Case GMPacketID.KillNpc                 '/RMATA
                Call HandleKillNpc(UserIndex)
            
            Case GMPacketID.WarnUser                '/ADVERTENCIA
                Call HandleWarnUser(UserIndex)
            
            Case GMPacketID.EditChar                '/MOD
                Call HandleEditChar(UserIndex)
                
            Case GMPacketID.RequestCharInfo         '/INFO
                Call HandleRequestCharInfo(UserIndex)
            
            Case GMPacketID.RequestCharStats        '/STAT
                Call HandleRequestCharStats(UserIndex)
                
            Case GMPacketID.RequestCharGold         '/BAL
                Call HandleRequestCharGold(UserIndex)
                
            Case GMPacketID.RequestCharInventory    '/INV
                Call HandleRequestCharInventory(UserIndex)
                
            Case GMPacketID.RequestCharBank         '/BOV
                Call HandleRequestCharBank(UserIndex)
            
            Case GMPacketID.RequestCharSkills       '/SKILLS
                Call HandleRequestCharSkills(UserIndex)
            
            Case GMPacketID.ReviveChar              '/R
                Call HandleReviveChar(UserIndex)
            
            Case GMPacketID.OnlineGM                '/ONGM
                Call HandleOnlineGM(UserIndex)
            
            Case GMPacketID.OnlineMap               '/ONMAP
                Call HandleOnlineMap(UserIndex)
                
            Case GMPacketID.Kick                    '/E
                Call HandleKick(UserIndex)
                
            Case GMPacketID.Execute                 '/EJECUTAR
                Call HandleExecute(UserIndex)
                
            Case GMPacketID.BanChar                 '/BAN
                Call HandleBanChar(UserIndex)
                
            Case GMPacketID.UnbanChar               '/UNBAN
                Call HandleUnbanChar(UserIndex)
                
            Case GMPacketID.NpcFollow               '/FOLLOW
                Call HandleNpcFollow(UserIndex)
                
            Case GMPacketID.SummonChar              '/SUM
                Call HandleSummonChar(UserIndex)
                
            Case GMPacketID.SpawnListRequest        '/CC
                Call HandleSpawnListRequest(UserIndex)
                
            Case GMPacketID.SpawnCreature           'SPA
                Call HandleSpawnCreature(UserIndex)
                
            Case GMPacketID.ResetNpcInventory       '/RESETINV
                Call HandleResetNpcInventory(UserIndex)
                
            Case GMPacketID.CleanWorld              '/LIMPIAR
                Call HandleCleanWorld(UserIndex)
            
            Case GMPacketID.ServerMessage           '/RMSG
                Call HandleServerMessage(UserIndex)
                
            Case GMPacketID.NickToIP                '/NICK2IP
                Call HandleNickToIP(UserIndex)
            
            Case GMPacketID.IPToNick                '/IP2NICK
                Call HandleIPToNick(UserIndex)
                
            Case GMPacketID.GuildOnlineMembers      '/ONCLAN
                Call HandleGuildOnlineMembers(UserIndex)
                    
            Case GMPacketID.TeleportCreate          '/CT
                Call HandleTeleportCreate(UserIndex)
                
            Case GMPacketID.TeleportDestroy         '/DT
                Call HandleTeleportDestroy(UserIndex)
                
            Case GMPacketID.RainToggle              '/LLUVIA
                Call HandleRainToggle(UserIndex)
                
            Case GMPacketID.Weather
                Call HandleWeather(UserIndex)
            
            Case GMPacketID.SetCharDescription      '/SETDESC
                Call HandleSetCharDescription(UserIndex)
            
            Case GMPacketID.ForceMP3ToMap          '/FORCEMP3MAP
                Call HandleForceMP3ToMap(UserIndex)
                
            Case GMPacketID.ForceWAVEToMap          '/FORCEWAVMAP
                Call HandleForceWAVEToMap(UserIndex)
                                
            Case GMPacketID.TalkAsNpc               '/TALKAS
                Call HandleTalkAsNpc(UserIndex)
            
            Case GMPacketID.DestroyAllItemsInArea   '/MASSDEST
                Call HandleDestroyAllItemsInArea(UserIndex)
                
            Case GMPacketID.MakeDumb                '/ESTUPIDO
                Call HandleMakeDumb(UserIndex)
                
            Case GMPacketID.MakeDumbNoMore          '/NOESTUPIDO
                Call HandleMakeDumbNoMore(UserIndex)
                
            Case GMPacketID.dumpIPTables            '/DUMPSECURITY
                Call HandleDumpIPTables(UserIndex)
                            
            Case GMPacketID.SetTrigger              '/TRIGGER
                Call HandleSetTrigger(UserIndex)
            
            Case GMPacketID.AskTrigger              '/TRIGGER
                Call HandleAskTrigger(UserIndex)
                
            Case GMPacketID.BannedIPList            '/BANIPLIST
                Call HandleBannedIPList(UserIndex)
            
            Case GMPacketID.BannedIPReload          '/BANIPRELOAD
                Call HandleBannedIPReload(UserIndex)
                
            Case GMPacketID.GuildMemberList         '/MIEMBROSCLAN
                Call HandleGuildMemberList(UserIndex)
            
            Case GMPacketID.GuildBan                '/BANCLAN
                Call HandleGuildBan(UserIndex)
            
            Case GMPacketID.BanIP                   '/BANIP
                Call HandleBanIP(UserIndex)
            
            Case GMPacketID.UnbanIP                 '/UNBANIP
                Call HandleUnbanIP(UserIndex)

            Case GMPacketID.CreateItem              '/CI
                Call HandleCreateItem(UserIndex)
            
            Case GMPacketID.DestroyItems            '/DEST
                Call HandleDestroyItems(UserIndex)
            
            Case GMPacketID.ForceMP3All             '/FORCEMP3
                Call HandleForceMP3All(UserIndex)
            
            Case GMPacketID.ForceWAVEAll            '/FORCEWAV
                Call HandleForceWAVEAll(UserIndex)
            
            Case GMPacketID.RemovePunishment        '/BORRARPENA
                Call HandleRemovePunishment(UserIndex)
            
            Case GMPacketID.TileBlockedToggle       '/BLOQ
                Call HandleTileBlockedToggle(UserIndex)
            
            Case GMPacketID.KillNpcNoRespawn        '/M
                Call HandleKillNpcNoRespawn(UserIndex)
            
            Case GMPacketID.KillAllNearbyNpcs       '/MASSKILL
                Call HandleKillAllNearbyNpcs(UserIndex)
            
            Case GMPacketID.LastIP                  '/LASTIP
                Call HandleLastIP(UserIndex)
            
            Case GMPacketID.ChangeMOTD              '/MOTDCAMBIA
                Call HandleChangeMOTD(UserIndex)
            
            Case GMPacketID.SetMOTD                 'ZMOTD
                Call HandleSetMOTD(UserIndex)
            
            Case GMPacketID.SystemMessage           '/SMSG
                Call HandleSystemMessage(UserIndex)
            
            Case GMPacketID.CreateNpc               '/ACC
                Call HandleCreateNpc(UserIndex)
            
            Case GMPacketID.CreateNpcWithRespawn    '/RACC
                Call HandleCreateNpcWithRespawn(UserIndex)
                        
            Case GMPacketID.ServerOpenToUsersToggle '/HABILITAR
                Call HandleServerOpenToUsersToggle(UserIndex)
            
            Case GMPacketID.TurnOffServer           '/APAGAR
                Call HandleTurnOffServer(UserIndex)
            
            Case GMPacketID.RemoveCharFromGuild     '/RAJARCLAN
                Call HandleRemoveCharFromGuild(UserIndex)
            
            Case GMPacketID.RequestCharMail         '/LASTEmail
                Call HandleRequestCharMail(UserIndex)
            
            Case GMPacketID.AlterPassword           '/APASS
                Call HandleAlterPassword(UserIndex)
            
            Case GMPacketID.AlterMail               '/AEmail
                Call HandleAlterMail(UserIndex)
            
            Case GMPacketID.AlterName               '/ANAME
                Call HandleAlterName(UserIndex)
            
            Case GMPacketID.ToggleCentinelActivated '/CENTINELAACTIVADO
                Call HandleToggleCentinelActivated(UserIndex)
            
            Case GMPacketID.DoBackUp                '/DOBACKUP
                Call HandleDoBackUp(UserIndex)
            
            Case GMPacketID.ShowGuildMessages       '/SHOWCMSG
                Call HandleShowGuildMessages(UserIndex)
            
            Case GMPacketID.SaveMap                 '/GUARDARMAPA
                Call HandleSaveMap(UserIndex)
            
            Case GMPacketID.ChangeMapInfoPK         '/MODMAPINFO PK
                Call HandleChangeMapInfoPK(UserIndex)
                
            Case GMPacketID.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
                Call HandleChangeMapInfoBackup(UserIndex)
            
            Case GMPacketID.ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
                Call HandleChangeMapInfoRestricted(UserIndex)
            
            Case GMPacketID.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
                Call HandleChangeMapInfoNoMagic(UserIndex)
            
            Case GMPacketID.ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
                Call HandleChangeMapInfoNoInvi(UserIndex)
            
            Case GMPacketID.ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
                Call HandleChangeMapInfoNoResu(UserIndex)
            
            Case GMPacketID.ChangeMapInfoLand       '/MODMAPINFO TERRENO
                Call HandleChangeMapInfoLand(UserIndex)
            
            Case GMPacketID.ChangeMapInfoZone       '/MODMAPINFO ZONA
                Call HandleChangeMapInfoZone(UserIndex)
            
            Case GMPacketID.ChangeMapInfoStealNpc   '/MODMAPINFO ROBONPC
                Call HandleChangeMapInfoStealNpc(UserIndex)
                
            Case GMPacketID.ChangeMapInfoNoOcultar  '/MODMAPINFO OCULTARSINEFECTO
                Call HandleChangeMapInfoNoOcultar(UserIndex)
                
            Case GMPacketID.ChangeMapInfoNoInvocar  '/MODMAPINFO INVOCARSINEFECTO
                Call HandleChangeMapInfoNoInvocar(UserIndex)
            
            Case GMPacketID.SaveChars               '/GUARDAR /GRABAR /G
                Call HandleSaveChars(UserIndex)
            
            Case GMPacketID.CleanSOS                '/BORRAR SOS
                Call HandleCleanSOS(UserIndex)
            
            Case GMPacketID.ShowServerForm          '/SHOW INT
                Call HandleShowServerForm(UserIndex)
                                    
            Case GMPacketID.KickAllChars            '/ETODOSPJS
                Call HandleKickAllChars(UserIndex)
            
            Case GMPacketID.ReloadNpcs              '/RELOADNpcS
                Call HandleReloadNpcs(UserIndex)
            
            Case GMPacketID.ReloadServidorIni         '/RELOADSINI
                Call HandleReloadServidorIni(UserIndex)
            
            Case GMPacketID.ReloadSpells            '/RELOADHECHIZOS
                Call HandleReloadSpells(UserIndex)
            
            Case GMPacketID.ReloadObjs           '/RELOADOBJ
                Call HandleReloadObjs(UserIndex)
            
            Case GMPacketID.Restart                 '/REINICIAR
                Call HandleRestart(UserIndex)
                    
            Case GMPacketID.ChatColor               '/CHATCOLOR
                Call HandleChatColor(UserIndex)
            
            Case GMPacketID.Ignored                 '/IGNORADO
                Call HandleIgnored(UserIndex)
            
            Case GMPacketID.CheckSlot               '/SLOT
                Call HandleCheckSlot(UserIndex)
                
            Case GMPacketID.SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
                Call HandleSetIniVar(UserIndex)

        End Select
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en GmCommand. Error: " & Err.Number & " - " & Err.description & _
                  ". Paquete: " & Command)

End Sub

Private Sub HandleOnline(ByVal UserIndex As Integer)
    Dim i As Long
    Dim Count As Long
    Dim Usuarios As String
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
       
        For i = 1 To LastUser
            If LenB(UserList(i).name) > 0 Then
                If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                    If LenB(Usuarios) > 0 Then
                        Usuarios = Usuarios & ", " & UserList(i).name
                    Else
                        Usuarios = UserList(i).name
                    End If
                End If
            End If
        Next i
        
        Call WriteConsoleMsg(UserIndex, "Población actual: " & Usuarios & ".", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "La población máxima fue de " & RecordPoblacion & " habitantes.", FONTTYPE_INFO)
     End With
End Sub

Private Sub HandleLoginChar(ByVal UserIndex As Integer)

'If UserList(UserIndex).incomingData.length < 13 Then
'Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
'EXIT SUB
'End If

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    Call buffer.ReadByte

    Dim UserName As String
    Dim Password As String
    
    Dim Id As Integer
    
    Dim SoyYo As Boolean
    
    UserName = buffer.ReadASCIIString
    Password = buffer.ReadASCIIString
    
    If Not SoyYo Then
        'Existe el personaje?
        If Not User_Exist(UserName) Then
            Call WriteErrorMsg(UserIndex, "NmNo")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    
        If Ban_Check(UserName) Then
            Call WriteErrorMsg(UserIndex, "Fuiste desterrado del mundo de Abraxas.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        '¿Es el password válido?
        If Not Check_Password(UserName, Password) Then
            Call WriteErrorMsg(UserIndex, "PsNo")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If
    
    '¿Ya esta conectado el personaje?
    If CheckForSameName(UserName) Then
        Dim index As Integer
        index = NameIndex(UserName)
        If UserList(index).Counters.Saliendo Then
            If UserList(index).Counters.Salir > 0 Then
                Call WriteErrorMsg(UserIndex, UserName & " está saliendo. Volvé a intentar en " & UserList(index).Counters.Salir + 1 & " segundos.")
            End If
        Else
            Call WriteErrorMsg(UserIndex, UserName & " ya está conectado.")
        End If
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    '¿Este Ip ya esta conectado?
    If AllowMultiLogins = 0 Then
        If CheckForSameIP(UserIndex, UserList(UserIndex).Ip) Then
            Call WriteErrorMsg(UserIndex, "No es posible jugar con más de un personaje al mismo tiempo.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If

    Call ConnectUser(UserIndex, UserName, Password, SoyYo)
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
End Sub

Private Sub HandleLoginNewChar(ByVal UserIndex As Integer)

'If UserList(UserIndex).incomingData.length < 5 Then
'Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
'EXIT SUB
'End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)

    Call buffer.ReadByte
    
    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(UserIndex, "La creación de personajes en este servidor se ha deshabilitado.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If

    If ServerSoloGMs > 0 Then
        Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Consulte la página oficial o el foro oficial para más información.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    Dim name As String
    Dim Pass As String
    Dim Race As eRaza
    Dim Gender As eGenero
    Dim Class As eClass
    Dim Atributos(NUMATRIBUTOS - 1) As Byte
    Dim Mail As String
    Dim Head As Integer
    
    name = buffer.ReadASCIIString
    Pass = buffer.ReadASCIIString

    Race = buffer.ReadByte
    Gender = buffer.ReadByte
    Class = buffer.ReadByte
    Head = buffer.ReadInteger
    Call buffer.ReadBlock(Atributos, NUMATRIBUTOS)
    Mail = buffer.ReadASCIIString
        
    Call ConnectNewUser(UserIndex, name, Pass, Race, Gender, Class, Atributos, Mail, Head)

    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleRecoverChar(ByVal UserIndex As Integer)
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    Call buffer.ReadByte
    
    Dim name As String
    Dim Email As String
    
    name = buffer.ReadASCIIString
    Email = buffer.ReadASCIIString
    
    If User_Exist(name) Then
    
        If UCase$(Email) = UCase$(GetVar(CharPath & name & ".chr", "CONTACTO", "Email")) Then
            If NameIndex(name) > 0 Then
                Call CloseSocket(NameIndex(name))
            End If
            
            Dim RandomPassword As String
            
            Dim ChrStr As String
            
            ChrStr = "abcdefghijklmnopqrstuvwxyz"
            ChrStr = ChrStr & UCase(ChrStr) & "0123456789"
        
            Dim i As Byte
            
            For i = 1 To 7
                RandomPassword = RandomPassword & mid$(ChrStr, Int(Rnd() * Len(ChrStr) + 1), 1)
            Next
            
            Call WriteVar(CharPath & name & ".chr", "INIT", "Password", RandomPassword)

            'frmMain.Inet1.OpenURL ("http://abraxas-online.com/cuentas?name=" & Name & "&confirmcode=" & RandomPassword)

            'Call WriteErrorMsg(UserIndex, "Se te envió enviado una nueva clave a " & Email & ".")
        
            Call CloseSocket(UserIndex)
            
        Else
            Call WriteErrorMsg(UserIndex, "La direccion de correo electrónico es incorrecta.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
    Else
        Call WriteErrorMsg(UserIndex, "No existe nadie llamado " & name & ".")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
End Sub

Private Sub HandleKillChar(ByVal UserIndex As Integer)

    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Call buffer.ReadByte
    
        Dim name As String
        Dim Pass As String
        Dim Email As String
                   
        name = UserList(UserIndex).name
        Pass = buffer.ReadASCIIString
        Email = buffer.ReadASCIIString

        If Pass <> GetVar(CharPath & name & ".chr", "INIT", "Password") Then
            Call WriteConsoleMsg(UserIndex, "Datos incorrectos.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Email <> GetVar(CharPath & name & ".chr", "CONTACTO", "Email") Then
            Call WriteConsoleMsg(UserIndex, "Datos incorrectos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            Call CloseSocket(UserIndex)
            Call KillCharInfo(name)
            Call BorrarUsuario(name)
        End If
            
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
End Sub

Private Sub HandleRequestRandomName(ByVal UserIndex As Integer)
    Dim Nombre As String
    
    Call UserList(UserIndex).incomingData.ReadByte

    Call DB_RS_Open("SELECT * from nombres ORDER BY RAND() LIMIT 1")
    
    Nombre = DB_RS!Nombre
    
    DB_RS.Close
    
    Call WriteRandomName(UserIndex, Nombre)
    
    Call FlushBuffer(UserIndex)
End Sub

Private Sub HandleTalk(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString
        
        If .Counters.Silencio > 0 Then
            'If we got here then packet is complete, copy data back to original queue
            Call .incomingData.CopyBuffer(buffer)
            Exit Sub
        End If
        
        If LenB(Chat) > 0 Then
                
            '[Consejeros & GMs]d
            If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                Call LogGM(.name, "Dijo: " & Chat)
            End If
            
            'I see you....
            If .flags.Oculto > 0 Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                
                If .flags.Navegando Then
                    If .Clase = eClass.Pirat Then
                        'Pierde la apariencia de fragata fantasmal
                        Call ToogleBoatBody(UserIndex)
                        Call WriteConsoleMsg(UserIndex, "¡Recuperaste tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, _
                                            NingunEscudo, NingunCasco)
                    End If
                Else
                    If .flags.Invisible < 1 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
                    End If
                End If
            End If
                        
            If .flags.AdminInvisible < 1 Then
                Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageChat(Chat, eChatType.Norm, .Char.CharIndex))
            Else
                Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageConsoleMsg(.name & ": " & Chat, FontTypeNames.FONTTYPE_GM))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleCompaMessage(ByVal UserIndex As Integer)

    'If UserList(UserIndex).incomingData.length < 5 Then
    'Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
    'EXIT SUB
    'End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Call buffer.ReadByte
        
        Dim Slot As Byte
        Dim Chat As String
        Dim TargetUserIndex As Integer
        Dim targetPriv As PlayerType
        
        Slot = buffer.ReadByte
        
        TargetUserIndex = NameIndex(.Compas.Compa(Slot))
        
        Chat = buffer.ReadASCIIString
        
        If LenB(Chat) < 1 Then
            Exit Sub
        End If

        Dim CompaSlot As Byte
        CompaSlot = EsCompaniero(TargetUserIndex, UserList(UserIndex).name)
        
        If CompaSlot < 1 Then
            Exit Sub
        End If

        If Not .Stats.Muerto Then
            If TargetUserIndex = 0 Then
                Exit Sub
            Else
                If Not .flags.Privilegios And PlayerType.User Then
                    If UserList(TargetUserIndex).flags.Privilegios And PlayerType.User Then
                        Call LogGM(.name, "Le dijo a su compañero " & UserList(TargetUserIndex).name & ": " & Chat)
                    End If
                End If
                    
                If .flags.AdminInvisible < 1 Then
                    Call WriteChat(TargetUserIndex, Chat, eChatType.Komp, , , CompaSlot)
                Else
                    Call WriteConsoleMsg(TargetUserIndex, .name & ": " & Chat, FontTypeNames.FONTTYPE_GM)
                End If
            End If
        End If
       
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandlePrivateMessage(ByVal UserIndex As Integer)

    'If UserList(UserIndex).incomingData.length < 5 Then
    'Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
    'EXIT SUB
    'End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Nombre As String
        Dim Chat As String
        Dim TargetUserIndex As Integer
        Dim targetPriv As PlayerType
                
        Nombre = buffer.ReadASCIIString
        Chat = buffer.ReadASCIIString
        
        TargetUserIndex = NameIndex(Nombre)

        If Not .Stats.Muerto Then
            
            If TargetUserIndex > 0 Then
            
                If UserList(TargetUserIndex).flags.Privilegios And PlayerType.User Then
                    If Not .flags.Privilegios And PlayerType.User Then
                        Call LogGM(.name, "Le dijo a " & UserList(TargetUserIndex).name & ": " & Chat)
                    End If
                    
                ElseIf .flags.Privilegios And PlayerType.User Then
                    Exit Sub
                End If
                                
                If LenB(Chat) > 0 Then
                    If .flags.AdminInvisible < 1 Then
                        Call WriteChat(UserIndex, Chat, eChatType.Priv, , .name)
                        Call WriteChat(TargetUserIndex, Chat, eChatType.Priv, , .name)
                    Else
                        Call WriteConsoleMsg(UserIndex, .name & ": " & Chat, FontTypeNames.FONTTYPE_GM)
                        Call WriteConsoleMsg(TargetUserIndex, .name & ": " & Chat, FontTypeNames.FONTTYPE_GM)
                    End If
                End If
                
            ElseIf User_Exist(Nombre) Then
                Call WriteConsoleMsg(UserIndex, Nombre & " no está.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "No existe nadie llamado " & Nombre & ".", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
       
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleDeleteChat(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageDeleteChatOverHead(UserList(UserIndex).Char.CharIndex))
    End With
End Sub

Private Sub HandleWalk(ByVal UserIndex As Integer, Heading As eHeading)
        
    Dim dummy As Long
    Dim TempTick As Long
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
                
        'Prevent SpeedHack
        
        If .Stats.Muerto Then
            Exit Sub
        End If
        
        'Prevent SpeedHack
        If .flags.TimesWalk >= 30 Then
            TempTick = GetTickCount And &H7FFFFFFF
            dummy = (TempTick - .flags.StartWalk)
            
            '5800 is actually less than what would be needed in perfect conditions to take 30 steps
            '(it's about 193 ms per step against the over 200 needed in perfect conditions)
            If dummy < 5800 Then
                If TempTick - .flags.CountSH > 90000 Then
                    .flags.CountSH = 0
                End If
                
                'If Not .flags.CountSH = 0 Then
                'If dummy > 0 Then
                'dummy = 126000 \ dummy
                'End If
                    
                'Call LogHackAttemp("Tramposo SH: " & .Name & " , " & dummy)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " fue echado por el servidor por posible uso de SH.", FontTypeNames.FONTTYPE_SERVER))
                Call CloseSocket(UserIndex)
                    
                Exit Sub
                'Else
                '.flags.CountSH = TempTick
                'End If
            End If
            .flags.StartWalk = TempTick
            .flags.TimesWalk = 0
        End If
        
        .flags.TimesWalk = .flags.TimesWalk + 1
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        If .flags.Paralizado < 1 Then
            If .flags.Meditando Then
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                .Char.FX = 0
            End If
            
            If .flags.Descansando Then
                .flags.Descansando = False
            End If

            'Move user
            Call MoveUserChar(UserIndex, Heading)
        
        Else    'paralized
            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                Call WriteConsoleMsg(UserIndex, "No podés moverte porque estás paralizado.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.CountSH = 0
        End If
        
        'Can't move while hidden except he is a thief
        If .flags.Oculto > 0 And .flags.AdminInvisible < 1 Then
            If .Clase <> eClass.Thief And .Clase <> eClass.Bandit Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                
                If .flags.Navegando Then
                    If .Clase = eClass.Pirat Then
                        'Pierde la apariencia de fragata fantasmal
                        Call ToogleBoatBody(UserIndex)
                        Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                    End If
                Else
                    'If not under a spell effect, show char
                    If .flags.Invisible < 1 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                    End If
                End If
            End If
        End If
        
    End With
    
End Sub

Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)
    UserList(UserIndex).incomingData.ReadByte
    
    Call WritePosUpdate(UserIndex)
End Sub

Private Sub HandleAttack(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        'If dead, can't attack
        If .Stats.Muerto Or .flags.Meditando Or .flags.Descansando Then
            Exit Sub
        End If
        
        'If equipped weapon is ranged, can't attack this way
        If UsaArco(UserIndex) > 0 Then
            Exit Sub
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Attack!
        Call UserAtaca(UserIndex)
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
        'I see you...
        If .flags.Oculto > 0 And .flags.AdminInvisible < 1 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando Then
                If .Clase = eClass.Pirat Then
                    'Pierde la apariencia de fragata fantasmal
                    Call ToogleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                                        
                ElseIf .flags.Invisible < 1 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                End If
            End If
        End If
    End With
End Sub

Private Sub HandlePickUp(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        'If dead, it can't pick up objs
        If .Stats.Muerto Then
            Exit Sub
        End If
        
        'If user is trading Items and attempts to pickup an Item, he's cheating, so we kick him.
        If .flags.Comerciando Then
            Exit Sub
        End If
        
        'Lower rank administrators can't pick up Items
        If .flags.Privilegios And PlayerType.Consejero Then
            If Not .flags.Privilegios And PlayerType.RoleMaster Then
                Call WriteConsoleMsg(UserIndex, "No podés tomar ningún objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        Call GetObj(UserIndex)
    End With
End Sub

Private Sub HandleResuscitationToggle(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        .flags.SeguroResu = Not .flags.SeguroResu
        
        If .flags.SeguroResu Then
            Call WriteResuscitationSafeOn(UserIndex)
        Else
            Call WriteResuscitationSafeOff(UserIndex)
        End If
    End With
End Sub

Private Sub HandleRequestGuildLeaderInfo(ByVal UserIndex As Integer)
    UserList(UserIndex).incomingData.ReadByte
    Call modGuilds.SendGuildLeaderInfo(UserIndex)
End Sub

Private Sub HandleRequestAttributes(ByVal UserIndex As Integer)
    Call UserList(UserIndex).incomingData.ReadByte
    Call WriteAttributes(UserIndex)
End Sub

Private Sub HandleRequestSkills(ByVal UserIndex As Integer)
    Call UserList(UserIndex).incomingData.ReadByte
    Call WriteSkills(UserIndex)
End Sub

Private Sub HandleRequestMiniStats(ByVal UserIndex As Integer)
    Call UserList(UserIndex).incomingData.ReadByte
    Call WriteMiniStats(UserIndex)
End Sub

Private Sub HandleBankEnd(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        'User exits banking mode
        .flags.Comerciando = False
        Call WriteObjCreate(UserIndex, ObjData(1054).GrhIndex, ObjData(1054).Type, .flags.TargetObjX, .flags.TargetObjY, ObjData(1054).name, 1)
    End With
End Sub

Private Sub HandleUserCommerceOk(ByVal UserIndex As Integer)
    Call UserList(UserIndex).incomingData.ReadByte
    'Trade accepted
    Call AceptarComercioUsu(UserIndex)
End Sub

Private Sub HandleUserCommerceReject(ByVal UserIndex As Integer)
    Dim otherUser As Integer
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        otherUser = .ComUsu.DestUsu
        
        'Offer rejected
        If otherUser > 0 Then
            If UserList(otherUser).flags.Logged Then
                Call WriteConsoleMsg(otherUser, .name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(otherUser)
                
                'Send data in the outgoing buffer of the other user
                Call FlushBuffer(otherUser)
            End If
        End If
        
        Call WriteConsoleMsg(UserIndex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
        Call FinComerciarUsu(UserIndex)
    End With
End Sub

Private Sub HandleCommerceEnd(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        'User quits commerce mode
        .flags.Comerciando = False
    End With
End Sub

Private Sub HandleUserCommerceEnd(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        'Quits commerce mode with user
        If .ComUsu.DestUsu > 0 Then
            If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
            Call WriteConsoleMsg(.ComUsu.DestUsu, .name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(.ComUsu.DestUsu)
            
            'Send data in the outgoing buffer of the other user
            Call FlushBuffer(.ComUsu.DestUsu)
        End If
        End If
        
        Call FinComerciarUsu(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has dejado de comerciar.", FontTypeNames.FONTTYPE_TALK)
    End With
End Sub

Private Sub HandleUserCommerceConfirm(ByVal UserIndex As Integer)
    
    Call UserList(UserIndex).incomingData.ReadByte

    'Validate the commerce
    If PuedeSeguirComerciando(UserIndex) Then
        'Tell the other user the confirmation of the offer
        Call WriteUserOfferConfirm(UserList(UserIndex).ComUsu.DestUsu)
        UserList(UserIndex).ComUsu.Confirmo = True
    End If
End Sub

Private Sub HandleCommerceChat(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
    
        Call buffer.ReadByte
    
        Dim Chat As String

        Chat = buffer.ReadASCIIString

        If LenB(Chat) > 0 Then
            If PuedeSeguirComerciando(UserIndex) Then
                Chat = UserList(UserIndex).name & "> " & Chat
                Call WriteCommerceChat(UserIndex, Chat, FontTypeNames.FONTTYPE_PARTY)
                Call WriteCommerceChat(UserList(UserIndex).ComUsu.DestUsu, Chat, FontTypeNames.FONTTYPE_PARTY)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleDrop(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim Slot As Byte
    Dim Amount As Integer
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte

        Slot = .incomingData.ReadByte
        Amount = .incomingData.ReadInteger
        
        'low rank admins can't drop Item. Neither can the dead nor those sailing.
        If .Stats.Muerto Or _
           ((.flags.Privilegios And PlayerType.Consejero) > 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) > 0) Then
            Exit Sub
        End If

        'If the user is trading, he can't drop Items => He's cheating, we kick him.
        If .flags.Comerciando Then
            Exit Sub
        End If
        
        If Slot < 1 Then
            Exit Sub
        End If
        
        If Slot > 200 Then
            Slot = Slot - 200
            
            If Slot > MaxBeltObjs Then
                Exit Sub
            End If
            
            If Amount < 1 Or Amount > MaxBeltObjs Then
                Exit Sub
            End If
                                    
            If .Belt.Obj(Slot).index = 0 Then
                Exit Sub
            End If
            
            Call DropBeltObj(UserIndex, Slot, Amount)
            
        Else
        
            If Slot > MaxInvSlots Then
                Exit Sub
            End If
            
            If Amount < 1 Or Amount > MaxInvObjs Then
                Exit Sub
            End If
                                    
            If .Inv.Obj(Slot).index = 0 Then
                Exit Sub
            End If
            
            Call DropObj(UserIndex, Slot, Amount)
        End If
        
    End With
End Sub

Private Sub HandleDropGold(ByVal UserIndex As Integer)
    
    Dim Amount As Long
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte

        Amount = .incomingData.ReadLong
        
        'low rank admins can't drop Item. Neither can the dead nor those sailing.
        If .Stats.Muerto Or _
           ((.flags.Privilegios And PlayerType.Consejero) > 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) > 0) Then
            Exit Sub
        End If

        'If the user is trading, he can't drop Items => He's cheating, we kick him.
        If .flags.Comerciando Then
            Exit Sub
        End If

        If Amount > 100000 Then
            Exit Sub 'Don't drop too much gold
        End If

        Call TirarOro(Amount, UserIndex)
        
        Call WriteUpdateGold(UserIndex)
        
    End With
    
End Sub

Private Sub HandleLeftClick(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        
        Call .ReadByte
        
        Dim x As Byte
        Dim y As Byte
        
        x = .ReadByte
        y = .ReadByte
        
        Call LookatTile(UserIndex, UserList(UserIndex).Pos.map, x, y)
    End With
End Sub

Private Sub HandleRightClick(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
    
        
        Call .ReadByte
        
        Dim x As Byte
        Dim y As Byte
        
        x = .ReadByte
        y = .ReadByte
        
        Call LookatTile(UserIndex, UserList(UserIndex).Pos.map, x, y, False, True)
    End With
End Sub

Private Sub HandleDoubleClick(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        
        Call .ReadByte
        
        Dim x As Byte
        Dim y As Byte
        
        x = .ReadByte
        y = .ReadByte
        
        'Call Accion(UserIndex, UserList(UserIndex).Pos.map, X, Y)
    End With
End Sub

Private Sub HandleWork(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim Skill As eSkill
        
        Skill = .incomingData.ReadByte
        
        If UserList(UserIndex).Stats.Muerto Then
            Exit Sub
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
            Case Ocultarse
            
                If .flags.EnConsulta Then
                    Call WriteConsoleMsg(UserIndex, "No podés ocultarte si estás en consulta.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            
                If .flags.Navegando Then
                    If .Clase <> eClass.Pirat Then
                        If Not .flags.UltimoMensaje = 3 Then
                                Call WriteConsoleMsg(UserIndex, "No podés ocultarte si estás navegando.", FontTypeNames.FONTTYPE_INFO)
                            .flags.UltimoMensaje = 3
                        End If
                        Exit Sub
                    End If
                End If
                
                If .flags.Oculto > 0 Then
                    Exit Sub
                End If
                
                Call DoOcultarse(UserIndex)
        End Select
    End With
End Sub

Private Sub HandleInitCrafting(ByVal UserIndex As Integer)

    Dim TotalItems As Integer
    Dim ItemsPorCiclo As Integer
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        TotalItems = .incomingData.ReadInteger
        ItemsPorCiclo = .incomingData.ReadInteger
        
        If TotalItems > 0 Then
            .Construir.Cantidad = TotalItems
            .Construir.PorCiclo = TotalItems 'MinimoInt(MaxItemsConstruibles(UserIndex), ItemsPorCiclo)
        End If
    End With
End Sub

Private Sub HandleUseSpellMacro(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(.name & " fue expulsado por Anti-macro de hechizos.", FontTypeNames.FONTTYPE_VENENO))
        Call WriteErrorMsg(UserIndex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
    End With
End Sub

Private Sub HandleUseItem(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        
        Slot = .incomingData.ReadByte
        
        If Slot > 0 And Slot <= MaxInvSlots Then
            If .Inv.Obj(Slot).index = 0 Then
                Exit Sub
            End If
        End If
        
        If .Stats.Muerto Then
            Exit Sub
        End If

        Call UseInvItem(UserIndex, Slot)
    End With
End Sub

Private Sub HandleUseBeltItem(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        
        Slot = .incomingData.ReadByte
        
        If Slot > 0 And Slot <= MaxBeltSlots Then
            If .Belt.Obj(Slot).index = 0 Then
                Exit Sub
            End If
        End If
        
        If .Stats.Muerto Then
            Exit Sub
        End If

        Call UseBeltInvItem(UserIndex, Slot)
    End With
End Sub

Private Sub HandleCraftBlacksmith(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        
        Call .ReadByte
        
        Dim Item As Integer
        
        Item = .ReadInteger
        
        If Item < 1 Then
            Exit Sub
        End If
        
        Call HerreroConstruirItem(UserIndex, Item)
    End With
End Sub

Private Sub HandleCraftCarpenter(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        
        Call .ReadByte
        
        Dim Item As Integer
        
        Item = .ReadInteger
        
        If Item < 1 Then
           Exit Sub
        End If
        
        Call CarpinteroConstruirItem(UserIndex, Item)
    End With
End Sub

Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim x As Byte
        Dim y As Byte
        Dim Skill As eSkill
        Dim DummyInt As Integer
        Dim tU As Integer   'Target user
        Dim tN As Integer   'Target Npc
        
        x = .incomingData.ReadByte
        y = .incomingData.ReadByte
        
        Skill = .incomingData.ReadByte

        If .Stats.Muerto Or .flags.Descansando Or .flags.Meditando Or _
            Not InMapBounds(.Pos.map, x, y) Then
            Exit Sub
        End If
        
        If Not InRangoVision(UserIndex, x, y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
                         
        Select Case Skill
            Case eSkill.Proyectiles
                'Check attack interval
                If Not IntervaloPermiteAtacar(UserIndex) Then
                    Exit Sub
                End If
                
                'Check Magic interval
                If Not IntervaloPermiteLanzarSpell(UserIndex, False) Then
                    Exit Sub
                End If
                
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex) Then
                    Exit Sub
                End If
                
                Dim Atacked As Boolean
                Atacked = True
                
                'Make sure the Item is valid and there is ammo equipped.
                With .Inv
                    'Tiene arma equipada?
                    If .LeftHand = 0 Then
                        DummyInt = 1
                        
                    'Usa munición? (Si no la usa, puede ser un arma arrojadiza)
                    ElseIf ObjData(.LeftHand).Municion Then
                        'La municion esta equipada en un slot válido?
                        If .AmmoAmount < 1 Then
                            DummyInt = 1
                        End If
                    End If
                    
                    If DummyInt > 0 Then
                        Call WriteConsoleMsg(UserIndex, "No tenés flechas.", FontTypeNames.FONTTYPE_INFO)
                        Call Desequipar(UserIndex, otArma)
                        Exit Sub
                    End If
                End With
                
                'Quitamos stamina
                If .Stats.MinSta >= 10 Then
                    Call QuitarSta(UserIndex, RandomNumber(1, 10))
                Else
                    Exit Sub
                End If
                
                Call LookatTile(UserIndex, .Pos.map, x, y, True)
                
                tU = .flags.TargetUser
                tN = .flags.TargetNpc
                
                'Validate target
                If tU > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(UserList(tU).Pos.y - .Pos.y) > RANGO_VISION_Y Then
                        Exit Sub
                    End If
                    
                    'Prevent from hitting self
                    If tU = UserIndex Then
                        Exit Sub
                    End If
                    
                    'Attack!
                    Atacked = UserAtacaUser(UserIndex, tU)
                    
                ElseIf tN > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(NpcList(tN).Pos.y - .Pos.y) > RANGO_VISION_Y And Abs(NpcList(tN).Pos.x - .Pos.x) > RANGO_VISION_X Then
                        Exit Sub
                    End If

                    If NpcList(tN).Attackable > 0 Then
                        Atacked = UserAtacaNpc(UserIndex, tN)
                    End If
                End If
                
                'Solo pierde la munición si pudo atacar al target, o tiro al aire
                If Atacked Then
                    'Call QuitarMunicion(UserIndex)
                End If

            Case eSkill.Pesca
                DummyInt = .Inv.RightHand
                
                If DummyInt = 0 Then
                    Exit Sub
                End If
                
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then
                    Exit Sub
                End If
                
                If maps(.Pos.map).mapData(.Pos.x, .Pos.y).Trigger = eTrigger.BAJOTECHO Or maps(.Pos.map).mapData(.Pos.x, .Pos.y).Trigger = eTrigger.EnPlataforma Then
                    Call WriteConsoleMsg(UserIndex, "No podés pescar desde ahí.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If HayAgua(.Pos.map, x, y) Then
                    Select Case DummyInt
                        Case CAÑA_PESCA
                            Call DoPescar(UserIndex)
                        
                        Case RED_PESCA
                            If Abs(.Pos.x - x) + Abs(.Pos.y - y) > 2 Then
                                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            Call DoPescarRed(UserIndex)
                        
                        Case Else
                            Exit Sub    'Invalid Item!
                    End Select
                    
                    'Play sound!
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.x, .Pos.y))
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, río o mar.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Robar
                'Does the map allow us to steal here?
                If MapInfo(.Pos.map).PK Then
                    
                    'Check interval
                    If Not IntervaloPermiteTrabajar(UserIndex) Then
                        Exit Sub
                    End If
                    
                    'Target whatever is in that tile
                    Call LookatTile(UserIndex, UserList(UserIndex).Pos.map, x, y, True)
                    
                    tU = .flags.TargetUser
                    
                    If tU > 0 And tU <> UserIndex Then
                        'Can't steal administrative players
                        If UserList(tU).flags.Privilegios And PlayerType.User Then
                            If Not UserList(tU).Stats.Muerto Then
                                If Abs(.Pos.x - x) + Abs(.Pos.y - y) > 2 Or _
                                .Pos.map <> UserList(tU).Pos.map Then
                                    Exit Sub
                                End If
                                 
                                'Check the trigger
                                If maps(UserList(tU).Pos.map).mapData(x, y).Trigger = eTrigger.ZONASEGURA Then
                                    Exit Sub
                                End If
                                
                                If maps(.Pos.map).mapData(.Pos.x, .Pos.y).Trigger = eTrigger.ZONASEGURA Then
                                    Exit Sub
                                End If
                                
                                Call DoRobar(UserIndex, tU)
                            End If
                        End If
                    End If
                End If
            
            Case eSkill.Talar
            
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then
                    Exit Sub
                End If
                
                If .Inv.RightHand = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Deberías equiparte el hacha.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteStopWorking(UserIndex)
                    Exit Sub
                End If
                
                If .Inv.RightHand <> HACHA_LEÑADOR Then
                    If .Inv.RightHand <> HACHA_LEÑA_ELFICA Then
                        Exit Sub
                    End If
                End If
                
                DummyInt = maps(.Pos.map).mapData(x, y).ObjInfo.index
                
                If DummyInt > 0 Then
                    If Abs(.Pos.x - x) + Abs(.Pos.y - y) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Call WriteStopWorking(UserIndex)
                        Exit Sub
                    End If
                    
                    If .Pos.x = x And .Pos.y = y Then
                        Call WriteStopWorking(UserIndex)
                        Exit Sub
                    End If
        
                    If maps(.Pos.map).mapData(.Pos.x, .Pos.y).Trigger = eTrigger.BAJOTECHO Or maps(.Pos.map).mapData(.Pos.x, .Pos.y).Trigger = eTrigger.EnPlataforma Then
                        Call WriteConsoleMsg(UserIndex, "No podés pescar desde ahí.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
        
                    '¿Hay un arbol donde clickeo?
                    If ObjData(DummyInt).Type = otArbol And .Inv.RightHand = HACHA_LEÑADOR Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.x, .Pos.y))
                        Call DoTalar(UserIndex)
                    ElseIf ObjData(DummyInt).Type = otArbolElfico And .Inv.RightHand = HACHA_LEÑA_ELFICA Then
                        If .Inv.RightHand = HACHA_LEÑA_ELFICA Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.x, .Pos.y))
                            Call DoTalar(UserIndex, True)
                        Else
                            Call WriteConsoleMsg(UserIndex, "El hacha utilizado no es suficientemente poderosa.", FontTypeNames.FONTTYPE_INFO)
                            Call WriteStopWorking(UserIndex)
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)
                        Call WriteStopWorking(UserIndex)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteStopWorking(UserIndex)
                End If
            
            Case eSkill.Mineria
                If Not IntervaloPermiteTrabajar(UserIndex) Then
                    Exit Sub
                End If
                
                If .Inv.RightHand <> PIQUETE_MINERO Then
                    Exit Sub
                End If
                
                'Target whatever is in the tile
                Call LookatTile(UserIndex, .Pos.map, x, y, True)
                
                DummyInt = maps(.Pos.map).mapData(x, y).ObjInfo.index
                
                If DummyInt > 0 Then
                    'Check distance
                    If Abs(.Pos.x - x) + Abs(.Pos.y - y) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Call WriteStopWorking(UserIndex)
                        Exit Sub
                    End If
                    
                    DummyInt = maps(.Pos.map).mapData(x, y).ObjInfo.index  'CHECK
                    '¿Hay un yacimiento donde clickeo?
                    If ObjData(DummyInt).Type = otYacimiento Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_MINERO, .Pos.x, .Pos.y))
                        Call DoMinar(UserIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
                        Call WriteStopWorking(UserIndex)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteStopWorking(UserIndex)
                End If
            
            Case eSkill.Domar
                'Target whatever is that tile
                Call LookatTile(UserIndex, .Pos.map, x, y, True)
                tN = .flags.TargetNpc
                
                If tN > 0 Then
                    'If NpcList(tN).flags.Domable > 0 Then
                        If Abs(.Pos.x - x) + Abs(.Pos.y - y) > 2 Then
                            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If NpcList(tN).TargetUser > 0 Then
                            If LenB(NpcList(tN).TargetUser) = LenB(UserList(UserIndex).name) Then
                                If UCase$(NpcList(tN).TargetUser) <> UCase$(UserList(UserIndex).name) Then
                                    Call WriteConsoleMsg(UserIndex, "No podés domar una criatura que está luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Sub
                                End If
                            End If
                        End If
                        
                        Call DoDomar(UserIndex, tN)
                    'Else
                    '    Call WriteConsoleMsg(UserIndex, "No podés domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)
                    'End If
                End If
                
            Case eSkill.Herreria
                'Target wehatever is in that tile
                Call LookatTile(UserIndex, .Pos.map, x, y, True)
                
                If .flags.TargetObjIndex > 0 Then
                    If ObjData(.flags.TargetObjIndex).Type = otYunque Then
                        Call WriteBlacksmithWeapons(UserIndex)
                        Call WriteBlacksmithArmors(UserIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yunque.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yunque.", FontTypeNames.FONTTYPE_INFO)
                End If
                
            Case Is > 200
                .flags.TargetObjInvSlot = Skill - 200
                
                .flags.TargetObjInvIndex = .Inv.Obj(.flags.TargetObjInvSlot).index
                
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then
                    Exit Sub
                End If
                
                'Target whatever is in the tile
                Call LookatTile(UserIndex, .Pos.map, x, y, True)
                
                'Check there is a proper Item there
                If .flags.TargetObjIndex > 0 Then
                    If ObjData(.flags.TargetObjIndex).Type = otFragua Then
                    
                        'Validate other Items
                        If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > MaxInvSlots Then
                            Exit Sub
                        End If
                        
                        'chequeamos que no se zarpe duplicando oro
                        'If .inv.Obj(.flags.TargetObjInvSlot).index <> .flags.TargetObjInvIndex Then
                        'If .inv.Obj(.flags.TargetObjInvSlot).index = 0 Or .inv.Obj(.flags.TargetObjInvSlot).Amount = 0 Then
                        'Call WriteConsoleMsg(UserIndex, "No tenés más minerales.", FontTypeNames.FONTTYPE_INFO)
                        'EXIT SUB
                        'End If
                        '
                        'FUISTE
                        'Call WriteErrorMsg(UserIndex, "Has sido expulsado por el sistema anti cheats.")
                        'Call FlushBuffer(UserIndex)
                        'Call CloseSocket(UserIndex)
                        'EXIT SUB
                        'End If

                        If ObjData(.flags.TargetObjInvIndex).Type = otMineral Then
                            Call FundirMineral(UserIndex)
                        ElseIf ObjData(.flags.TargetObjInvIndex).Type = otArma Then
                            Call FundirArmas(UserIndex)
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                        Call WriteStopWorking(UserIndex)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteStopWorking(UserIndex)
                End If
        End Select
    End With
End Sub

Public Sub HandleCastSpell(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte

        Dim Spell As Byte
        Dim x As Byte
        Dim y As Byte
        Dim map As Integer
        Dim tU As Integer   'Target user
        Dim tN As Integer   'Target Npc
                       
        Spell = .incomingData.ReadByte
        x = .incomingData.ReadByte
        y = .incomingData.ReadByte
        map = UserList(UserIndex).Pos.map
  
        If .Stats.Muerto Or .flags.Descansando Or .flags.Meditando Or Not InMapBounds(map, x, y) Then
            Exit Sub
        End If
                
        If Not InRangoVision(UserIndex, x, y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub
        End If

        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        .flags.Hechizo = Spell
        
        If .flags.Hechizo < 1 Then
            .flags.Hechizo = 0
        ElseIf .flags.Hechizo > MaxSpellSlots Then
            .flags.Hechizo = 0
        End If

        'Check the map allows spells to be casted.
        If MapInfo(map).MagiaSinEfecto > 0 Then
            Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energía", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        'Target whatever is in that tile
        Call LookatTile(UserIndex, map, x, y, True)
        
        'If it's outside range log it and exit
        If Abs(.Pos.x - x) > RANGO_VISION_X Or Abs(.Pos.y - y) > RANGO_VISION_Y Then
            Call LogCheating("Ataque fuera de rango de " & .name & "(" & map & "/" & .Pos.x & "/" & .Pos.y & ") ip: " & .Ip & " a la posición (" & .Pos.map & "/" & x & "/" & y & ")")
            Exit Sub
        End If
        
        'Check bow's interval
        If Not IntervaloPermiteUsarArcos(UserIndex, False) Then
            Exit Sub
        End If
        
        'Check Spell-Hit interval
        If Not IntervaloPermiteGolpeMagia(UserIndex) Then
            'Check Magic interval
            If Not IntervaloPermiteLanzarSpell(UserIndex) Then
                Exit Sub
            End If
        End If
    
        'Check intervals and cast
        If .flags.Hechizo > 0 Then
            Call LanzarHechizo(.flags.Hechizo, UserIndex)
        End If
        
        .flags.Hechizo = 0
    End With
End Sub

Private Sub HandleCreateNewGuild(ByVal UserIndex As Integer)
    'If UserList(UserIndex).incomingData.Length < 9 Then
    'Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
    'EXIT SUB
    'End If

On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Desc As String
        Dim GuildName As String
        Dim ErrorStr As String
        
        Desc = buffer.ReadASCIIString
        GuildName = Trim$(buffer.ReadASCIIString)
        
        If modGuilds.CrearNuevaGuilda(UserIndex, Desc, GuildName, ErrorStr) Then
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg(.name & " fundó la guilda " & GuildName & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44))

            'Update tag
             Call RefreshCharStatus(UserIndex)
        Else
            Call WriteConsoleMsg(UserIndex, ErrorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleSpellInfo(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim SpellSlot As Byte
        Dim Spell As Integer
        
        SpellSlot = .incomingData.ReadByte
        
        'Validate slot
        If SpellSlot < 1 Or SpellSlot > UserList(UserIndex).Spells.Nro Then
            Exit Sub
        End If
        
        'Validate spell in the slot
        Spell = .Spells.Spell(SpellSlot)
        If Spell > 0 And Spell < NumeroHechizos + 1 Then
            'Send information
            Call WriteConsoleMsg(UserIndex, Hechizos(Spell).Nombre & ": " & Hechizos(Spell).Desc, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleEquipItem(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim ItemSlot As Byte
        
        ItemSlot = .incomingData.ReadByte
        
        'Dead users can't equip Items
        If .Stats.Muerto Then
            Exit Sub
        End If
        
        If .flags.Navegando Then
            Exit Sub
        End If
        
        'Validate Item slot
        If ItemSlot > MaxInvSlots Or ItemSlot < 1 Then
            Exit Sub
        End If
        
        If .Inv.Obj(ItemSlot).index = 0 Then
            Exit Sub
        End If
        
        Call Equipar(UserIndex, ItemSlot)
    End With
End Sub

Private Sub HandleUnEquipItem(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim ObjType As eObjType
        
        ObjType = .incomingData.ReadByte
        
        'Dead users can't equip Items
        If .Stats.Muerto Then
            Exit Sub
        End If
        
        If .flags.Navegando Then
            Exit Sub
        End If
        
        Call Desequipar(UserIndex, ObjType)
    End With
End Sub

Private Sub HandleChangeHeading(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim Heading As eHeading
        Dim posX As Integer
        Dim posY As Integer
                
        Heading = .incomingData.ReadByte
        
        If .flags.Paralizado > 0 And .flags.Inmovilizado < 1 Then
            Select Case Heading
                Case eHeading.NORTH
                    posY = -1
                Case eHeading.EAST
                    posX = 1
                Case eHeading.SOUTH
                    posY = 1
                Case eHeading.WEST
                    posX = -1
            End Select
            
            If LegalPos(.Pos.map, .Pos.x + posX, .Pos.y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then
                Exit Sub
            End If
        End If
        
        'Validate Heading (VB won't say invalid cast if not a valid Index like .Net languages would do... *sigh*)
        If Heading > 0 And Heading < 5 Then
            .Char.Heading = Heading
            If .Char.CharIndex > 0 Then
                Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageChangeCharHeading(.Char.CharIndex, .Char.Heading))
            End If
        End If
    End With
End Sub

Private Sub HandleModifySkills(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 1 + NumSkills Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim i As Long
        Dim Count As Integer
        Dim points(1 To NumSkills) As Byte
        
        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NumSkills
            points(i) = .incomingData.ReadByte
            
            If points(i) < 0 Then
                Call LogHackAttemp(.name & " IP:" & .Ip & " trató de hackear los skills.")
                .Skills.NroFree = 0
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
            Count = Count + points(i)
        Next i
        
        If Count > .Skills.NroFree Then
            Call LogHackAttemp(.name & " IP:" & .Ip & " trató de hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        For i = 1 To NumSkills
            If points(i) > 0 Then
                .Skills.NroFree = .Skills.NroFree - points(i)
                .Skills.Skill(i).Elv = .Skills.Skill(i).Elv + points(i)

                If .Skills.Skill(i).Elv > MaxSkillPoints Then
                    .Skills.NroFree = .Skills.NroFree + .Skills.Skill(i).Elv - MaxSkillPoints
                End If
                
                Call CheckEluSkill(UserIndex, i, True)
            End If
        Next i
    End With
End Sub

Private Sub HandleTrain(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim SpawnedNpc As Integer
        Dim PetIndex As Byte
        
        PetIndex = .incomingData.ReadByte
        
        If .flags.TargetNpc < 1 Then
            Exit Sub
        End If
        
        If NpcList(.flags.TargetNpc).Type <> eNpcType.Entrenador Then
            Exit Sub
        End If
        
        If NpcList(.flags.TargetNpc).Nro < MaxPetsENTRENADOR Then
            If PetIndex > 0 And PetIndex < NpcList(.flags.TargetNpc).NroCriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(NpcList(.flags.TargetNpc).Criaturas(PetIndex).NpcIndex, NpcList(.flags.TargetNpc).Pos, True, False)
                
                If SpawnedNpc > 0 Then
                    NpcList(SpawnedNpc).MaestroNpc = .flags.TargetNpc
                    NpcList(.flags.TargetNpc).Nro = NpcList(.flags.TargetNpc).Nro + 1
                End If
            End If
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes.", NpcList(.flags.TargetNpc).Char.CharIndex, vbWhite))
        End If
    End With
End Sub

Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte
        Amount = .incomingData.ReadInteger
        
        'Dead people can't commerce...
        If .Stats.Muerto Then
            Exit Sub
        End If
        
        '¿El target es un Npc válido?
        If .flags.TargetNpc < 1 Then
            Exit Sub
        End If
        
        '¿El Npc puede comerciar?
        If NpcList(.flags.TargetNpc).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", NpcList(.flags.TargetNpc).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "No estás comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'User compra el Item
        Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNpc, Slot, Amount)
    End With
End Sub

Private Sub HandleBankExtractItem(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte
        Amount = .incomingData.ReadInteger
        
        If Slot < 1 Then
            Exit Sub
        End If
            
        If Amount < 1 Then
            Exit Sub
        End If

        If Not .flags.Comerciando Then
            Exit Sub
        End If
        
        '¿El target es un objeto válido?
        If .flags.TargetObjIndex < 1 Then
            Exit Sub
        End If
        
        '¿Es el alijo?
        If ObjData(.flags.TargetObjIndex).Type <> otAlijo Then
            Exit Sub
        End If
        
        Dim Pos As WorldPos
        
        Pos.map = .flags.TargetObjMap
        Pos.x = .flags.TargetObjX
        Pos.y = .flags.TargetObjY
            
        If Distancia(Pos, .Pos) > 5 Then
            Exit Sub
        End If
        
        If UserList(UserIndex).Bank.Obj(Slot).Amount < Amount Then
            Exit Sub
        End If
            
        'Agregamos el obj que compro al inventario y actualiza el slot del inv
        Call UserReciveObj(UserIndex, Slot, Amount)
        
    End With
End Sub

Private Sub HandleCommerceSell(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte
        Amount = .incomingData.ReadInteger
        
        If Slot < 1 Then
            Exit Sub
        End If
        
        If Amount < 1 Then
            Exit Sub
        End If
        
        'Dead people can't commerce...
        If .Stats.Muerto Then
            Exit Sub
        End If
        
        '¿El target es un Npc válido?
        If .flags.TargetNpc < 1 Then
            Exit Sub
        End If
        
        '¿El Npc puede comerciar?
        If NpcList(.flags.TargetNpc).Comercia = 0 Then
            Exit Sub
        End If
        
        'User compra el Item del slot
        Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNpc, Slot, Amount)
    End With
End Sub

Private Sub HandleBankDepositItem(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte
        Amount = .incomingData.ReadInteger
        
        If Slot < 1 Then
            Exit Sub
        End If
            
        If Amount < 1 Then
            Exit Sub
        End If
        
        If .Stats.Muerto Then
            Exit Sub
        End If
        
        If Not .flags.Comerciando Then
            Exit Sub
        End If
        
        '¿El target es un objeto válido?
        If .flags.TargetObjIndex < 1 Then
            Exit Sub
        End If
        
        '¿Es el alijo?
        If ObjData(.flags.TargetObjIndex).Type <> otAlijo Then
            Exit Sub
        End If
        
        Dim Pos As WorldPos
        
        Pos.map = .flags.TargetObjMap
        Pos.x = .flags.TargetObjX
        Pos.y = .flags.TargetObjY
            
        If Distancia(Pos, .Pos) > 5 Then
            Exit Sub
        End If
        
        If Slot > 200 Then
            If UserList(UserIndex).Inv.Obj(Slot - 200).index < 1 Then
                Exit Sub
            End If
        
            If UserList(UserIndex).Inv.Obj(Slot - 200).Amount < Amount Then
                Exit Sub
            End If
        Else
            If UserList(UserIndex).Inv.Obj(Slot).index < 1 Then
                Exit Sub
            End If
        
            If UserList(UserIndex).Inv.Obj(Slot).Amount < Amount Then
                Exit Sub
            End If
        End If
        
        'User deposita el Item del slot rdata
        Call UserDejaObj(UserIndex, Slot, Amount)
    End With
End Sub

Private Sub HandleMoveInvSlot(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        
        Call .ReadByte
        
        Dim Slot As Byte 'DRAGGING SLOT
        Dim Slot2 As Byte 'DEST SLOT

        Slot = .ReadByte
        Slot2 = .ReadByte
    End With
    
    With UserList(UserIndex).Inv
    
        'SI HAY DRAGItem ES Item
        If .Obj(Slot).index > 0 Then
        
            'SI DESTSLOT ES Item
            If .Obj(Slot2).index > 0 Then
                Dim ObjIndex As Integer
                Dim Amount As Long
                
                ObjIndex = .Obj(Slot).index
                Amount = .Obj(Slot).Amount
                
                .Obj(Slot).index = .Obj(Slot2).index
                .Obj(Slot).Amount = .Obj(Slot2).Amount
                
                .Obj(Slot2).index = ObjIndex
                .Obj(Slot2).Amount = Amount
                
            'SI DEST SLOT NO ES Item (SLOT VACIO)
            Else
                .Obj(Slot2).index = .Obj(Slot).index
                .Obj(Slot2).Amount = .Obj(Slot).Amount
                
                .Obj(Slot).index = 0
                .Obj(Slot).Amount = 0
            End If
            
        End If
    End With
End Sub

Private Sub HandleMoveBeltSlot(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        
        Call .ReadByte
        
        Dim Slot As Byte 'DRAGGING SLOT
        Dim Slot2 As Byte 'DEST SLOT

        Slot = .ReadByte
        Slot2 = .ReadByte
    End With
    
    With UserList(UserIndex).Belt
    
        'SI HAY DRAGItem ES Item
        If .Obj(Slot).index > 0 Then
        
            'SI DESTSLOT ES Item
            If .Obj(Slot2).index > 0 Then
                Dim ObjIndex As Integer
                Dim Amount As Long
                
                ObjIndex = .Obj(Slot).index
                Amount = .Obj(Slot).Amount
                
                .Obj(Slot).index = .Obj(Slot2).index
                .Obj(Slot).Amount = .Obj(Slot2).Amount
                
                .Obj(Slot2).index = ObjIndex
                .Obj(Slot2).Amount = Amount
                
            'SI DEST SLOT NO ES Item (SLOT VACIO)
            Else
                .Obj(Slot2).index = .Obj(Slot).index
                .Obj(Slot2).Amount = .Obj(Slot).Amount
                
                .Obj(Slot).index = 0
                .Obj(Slot).Amount = 0
            End If
        End If
    End With
End Sub

Private Sub HandleMoveSpellSlot(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim Slot As Byte 'DRAGING SLOT
        Dim Slot2 As Byte 'DEST SLOT
        Dim TempSlot As Byte
        
        Slot = .incomingData.ReadByte
        Slot2 = .incomingData.ReadByte
    
        TempSlot = .Spells.Spell(Slot)
        .Spells.Spell(Slot) = .Spells.Spell(Slot2)
        .Spells.Spell(Slot2) = TempSlot
    End With
End Sub

Private Sub HandleMoveBankSlot(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        
        Call .ReadByte
        
        Dim dir As Integer
        Dim Slot As Byte
        Dim TempItem As Obj
        
        If .ReadBoolean Then
            dir = 1
        Else
            dir = -1
        End If
        
        Slot = .ReadByte
    End With
        
    With UserList(UserIndex)
        TempItem.index = .Bank.Obj(Slot).index
        TempItem.Amount = .Bank.Obj(Slot).Amount
        
        If dir = 1 Then 'Mover arriba
            .Bank.Obj(Slot) = .Bank.Obj(Slot - 1)
            .Bank.Obj(Slot - 1).index = TempItem.index
            .Bank.Obj(Slot - 1).Amount = TempItem.Amount
        Else 'mover abajo
            .Bank.Obj(Slot) = .Bank.Obj(Slot + 1)
            .Bank.Obj(Slot + 1).index = TempItem.index
            .Bank.Obj(Slot + 1).Amount = TempItem.Amount
        End If
    End With
    
End Sub

Private Sub HandleGuildDescUpdate(ByVal UserIndex As Integer)

    'If UserList(UserIndex).incomingData.Length < 5 Then
    'Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
    'EXIT SUB
    'End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Desc As String
        
        Desc = buffer.ReadASCIIString
        
        Call modGuilds.ChangeDesc(Desc, .Guild_Id)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleUserCommerceOffer(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 7 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        Dim Slot As Byte
        Dim tUser As Integer
        Dim OfferSlot As Byte
        Dim ObjIndex As Integer
        
        Slot = .incomingData.ReadByte
        Amount = .incomingData.ReadLong
        OfferSlot = .incomingData.ReadByte
        
        'Get the other player
        tUser = .ComUsu.DestUsu
        
        'If he's already confirmed his offer, but now tries to change it, then he's cheating
        If UserList(UserIndex).ComUsu.Confirmo Then
            
            'Finish the trade
            Call FinComerciarUsu(UserIndex)
        
            If tUser < 1 Or tUser > MaxPoblacion Then
                Call FinComerciarUsu(tUser)
                Call Protocol.FlushBuffer(tUser)
            End If
        
            Exit Sub
        End If
        
        'If slot is invalid and it's not gold or it's not 0 (Substracting), then ignore it.
        If ((Slot < 0 Or Slot > MaxInvSlots) And Slot <> FLAGORO) Then
            Exit Sub
        End If
        
        'If OfferSlot is invalid, then ignore it.
        If OfferSlot < 1 Or OfferSlot > Max_OFFER_SLOTS + 1 Then
            Exit Sub
        End If
        
        'Can be negative if substracted from the offer, but never 0.
        If Amount = 0 Then
            Exit Sub
        End If
        
        'Has he got enough??
        If Slot = FLAGORO Then
            'Can't offer more than he has
            If Amount > .Stats.Gld - .ComUsu.GoldAmount Then
                Call WriteCommerceChat(UserIndex, "No tenés esa cantidad de oro para agregar a la oferta.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
        Else
            'If modifing a filled offerSlot, we already got the objIndex, then we don't need to know it
            If Slot > 0 Then ObjIndex = .Inv.Obj(Slot).index
            'Can't offer more than he has
            If Not TieneObjetos(ObjIndex, _
                TotalOfferItems(ObjIndex, UserIndex) + Amount, UserIndex) Then
                
                Call WriteCommerceChat(UserIndex, "No tenés esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
            
        End If
                
        Call AgregarOferta(UserIndex, OfferSlot, ObjIndex, Amount, Slot = FLAGORO)
        
        Call EnviarOferta(tUser, OfferSlot)
    End With
End Sub

Private Sub HandleGuildAcceptPeace(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Guild As String
        Dim ErrorStr As String
        Dim otherClanIndex As String
        
        Guild = buffer.ReadASCIIString
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(UserIndex, Guild, ErrorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, ErrorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .Guild_Id, PrepareMessageConsoleMsg("Tu guilda ha firmado la paz con " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu guilda ha firmado la paz con " & modGuilds.GuildName(.Guild_Id) & ".", FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildRejectAlliance(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Guild As String
        Dim ErrorStr As String
        Dim otherClanIndex As String
        
        Guild = buffer.ReadASCIIString
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(UserIndex, Guild, ErrorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, ErrorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .Guild_Id, PrepareMessageConsoleMsg("Tu guilda rechazado la propuesta de alianza de " & Guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.Guild_Id) & " ha rechazado nuestra propuesta de alianza con su guilda.", FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildRejectPeace(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Guild As String
        Dim ErrorStr As String
        Dim otherClanIndex As String
        
        Guild = buffer.ReadASCIIString
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(UserIndex, Guild, ErrorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, ErrorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .Guild_Id, PrepareMessageConsoleMsg("Tu guilda rechazado la propuesta de paz de " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.Guild_Id) & " ha rechazado nuestra propuesta de paz con su guilda.", FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildAcceptAlliance(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Guild As String
        Dim ErrorStr As String
        Dim otherClanIndex As String
        
        Guild = buffer.ReadASCIIString
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(UserIndex, Guild, ErrorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, ErrorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .Guild_Id, PrepareMessageConsoleMsg("Tu guilda ha firmado la alianza con " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu guilda ha firmado la paz con " & modGuilds.GuildName(.Guild_Id) & ".", FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildOfferPeace(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Guild As String
        Dim proposal As String
        Dim ErrorStr As String
        
        Guild = buffer.ReadASCIIString
        proposal = buffer.ReadASCIIString
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, Guild, RELACIONES_GUILD.PAZ, proposal, ErrorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de paz enviada.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, ErrorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildOfferAlliance(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Guild As String
        Dim proposal As String
        Dim ErrorStr As String
        
        Guild = buffer.ReadASCIIString
        proposal = buffer.ReadASCIIString
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, Guild, RELACIONES_GUILD.ALIADOS, proposal, ErrorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de alianza enviada.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, ErrorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildAllianceDetails(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Guild As String
        Dim ErrorStr As String
        Dim details As String
        
        Guild = buffer.ReadASCIIString
        
        details = modGuilds.r_VerPropuesta(UserIndex, Guild, RELACIONES_GUILD.ALIADOS, ErrorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, ErrorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildPeaceDetails(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Guild As String
        Dim ErrorStr As String
        Dim details As String
        
        Guild = buffer.ReadASCIIString
        
        details = modGuilds.r_VerPropuesta(UserIndex, Guild, RELACIONES_GUILD.PAZ, ErrorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, ErrorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildRequestJoinerInfo(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim User As String
        Dim details As String
        
        User = buffer.ReadASCIIString
        
        details = modGuilds.a_DetallesAspirante(UserIndex, User)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, "El personaje no ha mandado solicitud, o no estás habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteShowUserRequest(UserIndex, details)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildAlliancePropList(ByVal UserIndex As Integer)
    
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteAlianceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.ALIADOS))
End Sub

'
'Handles the "GuildPeacePropList" Message.
'
'UserIndex The Index of the user sending the Message.

Private Sub HandleGuildPeacePropList(ByVal UserIndex As Integer)


'Last Modification: 05\17\06
'

    
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WritePeaceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.PAZ))
End Sub

'
'Handles the "GuildDeclareWar" Message.
'
'UserIndex The Index of the user sending the Message.

Private Sub HandleGuildDeclareWar(ByVal UserIndex As Integer)


'Last Modification: 05\17\06
'

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Guild As String
        Dim ErrorStr As String
        Dim otherGuild_Id As Integer
        
        Guild = buffer.ReadASCIIString
        
        otherGuild_Id = modGuilds.r_DeclararGuerra(UserIndex, Guild, ErrorStr)
        
        If otherGuild_Id = 0 Then
            Call WriteConsoleMsg(UserIndex, ErrorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            'WAR shall be!
            Call SendData(SendTarget.ToGuildMembers, .Guild_Id, PrepareMessageConsoleMsg("TU CLAN HA ENTRADO EN GUERRA CON " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherGuild_Id, PrepareMessageConsoleMsg(modGuilds.GuildName(.Guild_Id) & " LE DECLARA LA GUERRA A TU CLAN.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .Guild_Id, PrepareMessagePlayWave(45))
            Call SendData(SendTarget.ToGuildMembers, otherGuild_Id, PrepareMessagePlayWave(45))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildAcceptNewMember(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim ErrorStr As String
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString
        
        If Not modGuilds.a_AceptarAspirante(UserIndex, UserName, ErrorStr) Then
            Call WriteConsoleMsg(UserIndex, ErrorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                Call RefreshCharStatus(tUser)
            End If
            
            Call SendData(SendTarget.ToGuildMembers, .Guild_Id, PrepareMessageConsoleMsg(UserName & " fue aceptado como miembro dla guilda.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .Guild_Id, PrepareMessagePlayWave(43))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildRejectNewMember(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim ErrorStr As String
        Dim UserName As String
        Dim Reason As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString
        Reason = buffer.ReadASCIIString
        
        If Not modGuilds.a_RechazarAspirante(UserIndex, UserName, ErrorStr) Then
            Call WriteConsoleMsg(UserIndex, ErrorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call WriteConsoleMsg(tUser, ErrorStr & ": " & Reason, FontTypeNames.FONTTYPE_GUILD)
            Else
                'hay que grabar en el char su rechazo
                Call modGuilds.a_RechazarAspiranteChar(UserName, .Guild_Id, Reason)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildKickMember(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim Guild_Id As Integer
        
        UserName = buffer.ReadASCIIString
        
        Guild_Id = modGuilds.m_EcharMiembroDeGuilda(UserIndex, UserName)
        
        If Guild_Id > 0 Then
            Call SendData(SendTarget.ToGuildMembers, Guild_Id, PrepareMessageConsoleMsg(UserName & " fue expulsado dla guilda.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, Guild_Id, PrepareMessagePlayWave(45))
        Else
            Call WriteConsoleMsg(UserIndex, "No podés expulsar ese personaje dla guilda.", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildUpdateNews(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Call modGuilds.ActualizarNoticias(UserIndex, buffer.ReadASCIIString)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildMemberInfo(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Call modGuilds.SendDetallesPersonaje(UserIndex, buffer.ReadASCIIString)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildOpenElections(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim Error As String
        
        If Not modGuilds.v_AbrirElecciones(UserIndex, Error) Then
            Call WriteConsoleMsg(UserIndex, Error, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .Guild_Id, PrepareMessageConsoleMsg("¡Han comenzado las elecciones dla guilda! Puedes votar escribiendo /vOTO seguido del nombre del personaje, por ejemplo: /vOTO " & .name, FontTypeNames.FONTTYPE_GUILD))
        End If
    End With
End Sub

'
'Handles the "GuildRequestMembership" Message.
'
'UserIndex The Index of the user sending the Message.

Private Sub HandleGuildRequestMembership(ByVal UserIndex As Integer)


'Last Modification: 05\17\06
'

    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Guild As String
        Dim application As String
        Dim ErrorStr As String
        
        Guild = buffer.ReadASCIIString
        application = buffer.ReadASCIIString
        
        If Not modGuilds.a_NuevoAspirante(UserIndex, Guild, application, ErrorStr) Then
           Call WriteConsoleMsg(UserIndex, ErrorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
           Call WriteConsoleMsg(UserIndex, "Tu solicitud fue enviada. Espera prontas noticias del líder de " & Guild & ".", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildRequestDetails(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Call modGuilds.SendGuildDetails(UserIndex, buffer.ReadASCIIString)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleQuit(ByVal UserIndex As Integer)

    Dim tUser As Integer
    Dim isNotVisible As Boolean
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Paralizado > 0 Then
            Call WriteConsoleMsg(UserIndex, "No podés salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        'Subastas
        If UserIndex = Subasta.UserIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "No podés salir mientras estás subastando.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        ElseIf UserIndex = Subasta.OfertaIndex Then
            Call WriteConsoleMsg(UserIndex, "No podés salir mientras estás ganando una subasta.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        'exit secure commerce
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu
            
            If UserList(tUser).flags.Logged Then
                If UserList(tUser).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(tUser)
                End If
            End If
            
            Call WriteConsoleMsg(UserIndex, "Comercio cancelado. ", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(UserIndex)
        End If
        
        Call CerrarUsuario(UserIndex)
    End With
End Sub

Private Sub HandleGuildLeave(ByVal UserIndex As Integer)
    
    Dim Guild_Id As Integer
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        'obtengo el Guild_Id
        Guild_Id = m_EcharMiembroDeGuilda(UserIndex, .name)
        
            If Guild_Id > 0 Then
                Call WriteConsoleMsg(UserIndex, "Dejas la guilda.", FontTypeNames.FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, Guild_Id, PrepareMessageConsoleMsg(.name & " deja la guilda.", FontTypeNames.FONTTYPE_GUILD))
            Else
            Call WriteConsoleMsg(UserIndex, "Tú no podés salir de esta guilda.", FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub

Private Sub HandleRequestAccountState(ByVal UserIndex As Integer)

    Dim earnings As Integer
    Dim percentage As Integer
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .Stats.Muerto Then
            Exit Sub
        End If
        
        'Validate target Npc
        If .flags.TargetNpc < 1 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(NpcList(.flags.TargetNpc).Pos, .Pos) > 5 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del Timbero.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If NpcList(.flags.TargetNpc).Type = eNpcType.Timbero Then
            If Not .flags.Privilegios And PlayerType.User Then
                earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    
                If earnings > 1 And Apuestas.Ganancias > 0 Then
                    percentage = Int(earnings * 100 \ Apuestas.Ganancias)
                End If
                    
                If earnings < 0 And Apuestas.Perdidas > 0 Then
                    percentage = Int(earnings * 100 \ Apuestas.Perdidas)
                End If
                    
                Call WriteConsoleMsg(UserIndex, "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & earnings & " (" & percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
End Sub

Private Sub HandlePetStand(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .Stats.Muerto Then
            Exit Sub
        End If
        
        Dim i As Byte
        
        i = .incomingData.ReadByte
        
        If i > .Pets.Nro Then
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(NpcList(.Pets.Pet(i).index).Pos, .Pos) > 10 Then
            Exit Sub
        End If
        
        'Do it!
        NpcList(.Pets.Pet(i).index).Movement = TipoAI.Estatico
        
        Call Expresar(.Pets.Pet(i).index, UserIndex)
    End With
End Sub

Private Sub HandlePetFollow(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .Stats.Muerto Then
            Exit Sub
        End If

        Dim i As Byte
        
        i = .incomingData.ReadByte
        
        If i > .Pets.Nro Then
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(NpcList(.Pets.Pet(i).index).Pos, .Pos) > 10 Then
            Exit Sub
        End If
        
        'Do it
        Call FollowAmo(.Pets.Pet(i).index)
        
        Call Expresar(.Pets.Pet(i).index, UserIndex)
    End With
End Sub

Private Sub HandleReleasePet(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .Stats.Muerto Then
            Exit Sub
        End If
        
        'Validate target Npc
        If .flags.TargetNpc < 1 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(NpcList(.flags.TargetNpc).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make usre it's the user's pet
        If NpcList(.flags.TargetNpc).MaestroUser <> UserIndex Then
            Exit Sub
        End If
        
        Call QuitarMascota(UserIndex, .flags.TargetNpc)
    End With
End Sub

Private Sub HandleRest(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .Stats.Muerto Then
            Exit Sub
        End If
        
        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(UserIndex)
            
            If Not .flags.Descansando Then
                Call WriteConsoleMsg(UserIndex, "Te acomodás junto a la fogata y comienzas a descansar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Te levantás.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.Descansando = Not .flags.Descansando
        Else
            If .flags.Descansando Then
                Call WriteRestOK(UserIndex)
                .flags.Descansando = False
                Exit Sub
            End If
            
            Call WriteConsoleMsg(UserIndex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleMeditate(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte

        If .Stats.Muerto Or .Stats.MinMan = .Stats.MaxMan Then
            If .flags.Privilegios And PlayerType.User Then
                Exit Sub
            End If
        End If
                         
        If Not .flags.Meditando Then
            .flags.Meditando = True
            
            .Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
            
            Select Case .Stats.Elv
                'Show proper FX according to level
                Case Is < 15
                    .Char.FX = FXIDs.FX_MEDITARCHICO
                Case Is < 25
                    .Char.FX = FXIDs.FX_MEDITARMEDIANO
                Case Is < 35
                    .Char.FX = FXIDs.FX_MEDITARGRANDE
                Case Is < 40
                    .Char.FX = FXIDs.FX_MEDITARXGRANDE
                Case Else
                    .Char.FX = FXIDs.FX_MEDITARXXGRANDE
            End Select
                        
            .Char.Loops = INFInitE_Loops
            
            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharMeditate(.Char.CharIndex))

        Else
            .flags.Meditando = False
            .Counters.bPuedeMeditar = False
            
            .Char.FX = 0
            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCreateCharFX(.Char.CharIndex))
        End If
    End With
End Sub

Private Sub HandleConsultation(ByVal UserIndex As String)
    
    Dim UserConsulta As Integer
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        'Comando exclusivo para gms
        If Not EsGM(UserIndex) Then
            Exit Sub
        End If
        
        UserConsulta = .flags.TargetUser
        
        'Se asegura que el target es un usuario
        If UserConsulta = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un usuario, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'No podes ponerte a vos mismo en modo consulta.
        If UserConsulta = UserIndex Then
            Exit Sub
        End If
        
        'No podes estra en consulta con otro gm
        If EsGM(UserConsulta) Then
            Call WriteConsoleMsg(UserIndex, "No podés iniciar el modo consulta con otro administrador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim UserName As String
        UserName = UserList(UserConsulta).name
        
        'Si ya estaba en consulta, termina la consulta
        If UserList(UserConsulta).flags.EnConsulta Then
            Call WriteConsoleMsg(UserIndex, "Has terminado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has terminado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.name, "Termino consulta con " & UserName)
            
            UserList(UserConsulta).flags.EnConsulta = False
        
        'Sino la inicia
        Else
            Call WriteConsoleMsg(UserIndex, "Has iniciado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has iniciado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.name, "Inicio consulta con " & UserName)
            
            With UserList(UserConsulta)
            
                .flags.EnConsulta = True
                
                'Pierde invi u ocu
                If .flags.Invisible > 0 Or .flags.Oculto > 0 Then
                    .flags.Oculto = 0
                    .flags.Invisible = 0
                    
                    .Counters.TiempoOculto = 0
                    .Counters.Invisibilidad = 0
                    
                    Call SendData(SendTarget.ToPCArea, UserConsulta, PrepareMessageSetInvisible(UserList(UserConsulta).Char.CharIndex, False))
                End If
                
            End With

        End If
        
    End With

End Sub

Private Sub HandleRequestStats(ByVal UserIndex As Integer)
    
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendUserStatsTxt(UserIndex, UserIndex)
End Sub

Private Sub HandleHelp(ByVal UserIndex As Integer)
    
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendHelp(UserIndex)
End Sub

Private Sub HandleCommerceStart(ByVal UserIndex As Integer)

    Dim i As Integer
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .Stats.Muerto Or .flags.Comerciando Then
            Exit Sub
        End If
        
        'Validate target Npc
        If .flags.TargetNpc > 0 Then
            'Does the Npc want to trade??
            If NpcList(.flags.TargetNpc).Comercia = 0 Then
                If LenB(NpcList(.flags.TargetNpc).Desc) > 0 Then
                    Call WriteChatOverHead(UserIndex, "No tengo ningún interés en comerciar.", NpcList(.flags.TargetNpc).Char.CharIndex, vbWhite)
                End If
                
                Exit Sub
            End If
            
            If Distancia(NpcList(.flags.TargetNpc).Pos, .Pos) > 5 Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
                    
            Call WriteNpcInventory(UserIndex)

            UserList(UserIndex).flags.Comerciando = True
            
        ElseIf .flags.TargetUser > 0 Then
            'User commerce...
            'Can he commerce??
            If .flags.Privilegios And PlayerType.Consejero Then
                Call WriteConsoleMsg(UserIndex, "No podés vender ítems.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            
            'Is the other one dead??
            If UserList(.flags.TargetUser).Stats.Muerto Then
                Call WriteConsoleMsg(UserIndex, "¡¡No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is it me??
            If .flags.TargetUser = UserIndex Then
                Exit Sub
            End If
            
            'Check distance
            If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is he already trading?? is it with me or someone else??
            If UserList(.flags.TargetUser).flags.Comerciando And _
                UserList(.flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No podés comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Initialize some variables...
            .ComUsu.DestUsu = .flags.TargetUser
            .ComUsu.DestNick = UserList(.flags.TargetUser).name
            For i = 1 To Max_OFFER_SLOTS
                .ComUsu.Cant(i) = 0
                .ComUsu.Objeto(i) = 0
            Next i
            .ComUsu.GoldAmount = 0
            
            .ComUsu.Acepto = False
            .ComUsu.Confirmo = False
            
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(UserIndex, .flags.TargetUser)
        Else
            Call WriteConsoleMsg(UserIndex, "Primero hacé click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleBankStart(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .Stats.Muerto Then
            Exit Sub
        End If

        If .flags.Comerciando Then
            Exit Sub
        End If
        
        '¿El target es un objeto válido?
        If .flags.TargetObjIndex < 1 Then
            Call WriteConsoleMsg(UserIndex, "Primero hacé click izquierdo sobre el alijo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿Es el alijo?
        If ObjData(.flags.TargetObjIndex).Type <> otAlijo Then
            Exit Sub
        End If
        
        Dim Pos As WorldPos
        
        Pos.map = .flags.TargetObjMap
        Pos.x = .flags.TargetObjX
        Pos.y = .flags.TargetObjY
            
        If Distancia(Pos, .Pos) > 5 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del alijo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
                                    
        Call WriteBank(UserIndex)
        UserList(UserIndex).flags.Comerciando = True
        
        Call WriteObjCreate(UserIndex, ObjData(1055).GrhIndex, ObjData(1055).Type, Pos.x, Pos.y, ObjData(1055).name, 1)
    End With
End Sub

Private Sub HandleRequestMOTD(ByVal UserIndex As Integer)
    Call UserList(UserIndex).incomingData.ReadByte
    Call SendMOTD(UserIndex)
End Sub

Private Sub HandleUpTime(ByVal UserIndex As Integer)

    Call UserList(UserIndex).incomingData.ReadByte
    
    Dim Time As Long
    Dim UpTimeStr As String
    
    'Get total time in seconds
    Time = ((GetTickCount() And &H7FFFFFFF) - tInicioServer) \ 1000
    
    'Get times in dd:hh:mm:ss format
    UpTimeStr = (Time Mod 60) & " segundos."
    Time = Time \ 60
    
    UpTimeStr = (Time Mod 60) & " minutos, " & UpTimeStr
    Time = Time \ 60
    
    UpTimeStr = (Time Mod 24) & " horas, " & UpTimeStr
    Time = Time \ 24
    
    If Time = 1 Then
        UpTimeStr = Time & " día, " & UpTimeStr
    Else
        UpTimeStr = Time & " días, " & UpTimeStr
    End If
    
    Call WriteConsoleMsg(UserIndex, "Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
End Sub

Private Sub HandlePartyLeave(ByVal UserIndex As Integer)

    
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call mdParty.SalirDeParty(UserIndex)
End Sub

Private Sub HandlePartyCreate(ByVal UserIndex As Integer)
    
    
    Call UserList(UserIndex).incomingData.ReadByte
    
    If Not mdParty.PuedeCrearParty(UserIndex) Then
        Exit Sub
    End If
    
    Call mdParty.CrearParty(UserIndex)
End Sub

Private Sub HandlePartyJoin(ByVal UserIndex As Integer)
    
    
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call mdParty.SolicitarIngresoAParty(UserIndex)
End Sub

Private Sub HandleInquiry(ByVal UserIndex As Integer)

    
    Call UserList(UserIndex).incomingData.ReadByte
    
    ConsultaPopular.SendInfoEncuesta (UserIndex)
End Sub

Private Sub HandleGuildMessage(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString
        
        If LenB(Chat) > 0 Then
            If .Guild_Id > 0 Then
                Call SendData(SendTarget.ToDiosesYGuilda, .Guild_Id, PrepareMessageChat(Chat, eChatType.Guil, , .name))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandlePartyMessage(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString
        
        If LenB(Chat) > 0 Then
            Call mdParty.BroadCastParty(UserIndex, Chat)
'TODO: Con la 0.12.1 se debe definir si esto vuelve o se borra (/cMSG overhead)
            'Call SendData(SendTarget.ToPartyArea, UserIndex, UserList(UserIndex).Pos.map, "||" & vbYellow & "°< " & mid$(rData, 7) & " >°" & CStr(UserList(UserIndex).Char.CharIndex))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleCentinelReport(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Call CentinelaCheckClave(UserIndex, .incomingData.ReadInteger)
    End With
End Sub

Private Sub HandleGuildOnline(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim onlineList As String
        
        onlineList = modGuilds.m_ListaDeMiembrosOnline(UserIndex, .Guild_Id)
        
        If .Guild_Id > 0 Then
            Call WriteConsoleMsg(UserIndex, "Mascotas de tu guilda conectados: " & onlineList, FontTypeNames.FONTTYPE_GUILDMSG)
        Else
            Call WriteConsoleMsg(UserIndex, "No pertences a ninguna guilda.", FontTypeNames.FONTTYPE_GUILDMSG)
        End If
    End With
End Sub

Private Sub HandlePartyOnline(ByVal UserIndex As Integer)
    
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call mdParty.OnlineParty(UserIndex)
End Sub

Private Sub HandleRoleMasterRequest(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim request As String
        
        request = buffer.ReadASCIIString
        
        If LenB(request) > 0 Then
            Call WriteConsoleMsg(UserIndex, "Su solicitud fue enviada.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGMRequest(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If Not Ayuda.Existe(.name) Then
            Call WriteConsoleMsg(UserIndex, "El mensaje fue entregado, ahora sólo debes esperar que se desocupe algún GM.", FontTypeNames.FONTTYPE_INFO)
            Call Ayuda.Push(.name)
        Else
            Call Ayuda.Quitar(.name)
            Call Ayuda.Push(.name)
            Call WriteConsoleMsg(UserIndex, "Ya habías mandado un mensaje, tu mensaje fue movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleBugReport(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Dim N As Integer
        
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim BugReport As String
        
        BugReport = buffer.ReadASCIIString
        
        N = FreeFile
        Open App.Path & "/LOGS/bUGs.log" For Append Shared As N
        Print #N, "Usuario:" & .name & "  Fecha:" & Date & "    Hora:" & Time
        Print #N, "BUG:"
        Print #N, BugReport
        Print #N, "########################################################################"
        Close #N
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleChangeDescription(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim description As String
            
        description = buffer.ReadASCIIString
        
        If Not .Stats.Muerto Then
            If Len(description) > 0 And Len(description) < 51 Then
                If Not AsciiValidos(description) Then
                    Call WriteConsoleMsg(UserIndex, "La descripción tiene caracteres inválidos.", FontTypeNames.FONTTYPE_INFO)
                Else
                    .Desc = Trim$(description)
                    Call WriteConsoleMsg(UserIndex, "La descripción ha cambiado.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildVote(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim vote As String
        Dim ErrorStr As String
        
        vote = buffer.ReadASCIIString
        
        If Not modGuilds.v_UsuarioVota(UserIndex, vote, ErrorStr) Then
            Call WriteConsoleMsg(UserIndex, "Voto NO contabilizado: " & ErrorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Voto contabilizado.", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleShowGuildNews(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Call modGuilds.SendGuildNews(UserIndex)
    End With
End Sub

Private Sub HandlePunishments(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim name As String
        Dim Count As Integer
        
        name = buffer.ReadASCIIString
        
        If LenB(name) > 0 Then
            If (InStrB(name, "/") > 0) Then
                name = Replace(name, "/", vbNullString)
            End If
            If (InStrB(name, "/") > 0) Then
                name = Replace(name, "/", vbNullString)
            End If
            If (InStrB(name, ":") > 0) Then
                name = Replace(name, ":", vbNullString)
            End If
            If (InStrB(name, "|") > 0) Then
                name = Replace(name, "|", vbNullString)
            End If
            
            If (EsAdmin(name) Or EsDios(name) Or EsSemiDios(name) Or EsConsejero(name) Or EsRolesMaster(name)) And (UserList(UserIndex).flags.Privilegios And PlayerType.User) Then
                Call WriteConsoleMsg(UserIndex, "No podés ver las penas de los administradores.", FontTypeNames.FONTTYPE_INFO)
            Else
                If User_Exist(name) Then
                    Count = Val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))
                    If Count = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                    Else
                        While Count > 0
                            Call WriteConsoleMsg(UserIndex, Count & " - " & GetVar(CharPath & name & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
                            Count = Count - 1
                        Wend
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Personaje " & name & " inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleChangePassword(ByVal UserIndex As Integer)

If UserList(UserIndex).incomingData.length < 5 Then
    Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
    Exit Sub
End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Dim oldPass As String
        Dim newPass As String
        Dim oldPass2 As String
        
        
        Call buffer.ReadByte

        oldPass = buffer.ReadASCIIString
        newPass = buffer.ReadASCIIString
        
        If LenB(newPass) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Debes especificar una contraseña nueva, inténtalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
        Else
            oldPass2 = GetVar(CharPath & UserList(UserIndex).name & ".chr", "INIT", "Password")
            
            If oldPass2 <> oldPass Then
                Call WriteConsoleMsg(UserIndex, "La contraseña actual proporcionada no es correcta. La contraseña no fue cambiada, inténtalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteVar(CharPath & UserList(UserIndex).name & ".chr", "INIT", "Password", newPass)
                Call WriteConsoleMsg(UserIndex, "La contraseña fue cambiada con éxito.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGamble(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim Amount As Integer
        
        Amount = .incomingData.ReadInteger
        
        If .Stats.Muerto Then
            Exit Sub
        End If
        
        If .flags.TargetNpc < 1 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Distancia(NpcList(.flags.TargetNpc).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        ElseIf NpcList(.flags.TargetNpc).Type <> eNpcType.Timbero Then
            Call WriteChatOverHead(UserIndex, "No tengo ningún interés en apostar.", NpcList(.flags.TargetNpc).Char.CharIndex, vbWhite)
        ElseIf Amount < 1 Then
            Call WriteChatOverHead(UserIndex, "El mínimo de apuesta es 1 moneda.", NpcList(.flags.TargetNpc).Char.CharIndex, vbWhite)
        ElseIf Amount > 5000 Then
            Call WriteChatOverHead(UserIndex, "El máximo de apuesta es 5000 monedas.", NpcList(.flags.TargetNpc).Char.CharIndex, vbWhite)
        ElseIf .Stats.Gld < Amount Then
            Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", NpcList(.flags.TargetNpc).Char.CharIndex, vbWhite)
        Else
            If RandomNumber(1, 100) <= 47 Then
                .Stats.Gld = .Stats.Gld + Amount
                Call WriteChatOverHead(UserIndex, "¡Felicidades! Has ganado " & CStr(Amount) & " monedas de oro.", NpcList(.flags.TargetNpc).Char.CharIndex, vbWhite)
                
                Apuestas.Perdidas = Apuestas.Perdidas + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.Gld = .Stats.Gld - Amount
                Call WriteChatOverHead(UserIndex, "Lo siento, has perdido " & CStr(Amount) & " monedas de oro.", NpcList(.flags.TargetNpc).Char.CharIndex, vbWhite)
                
                Apuestas.Ganancias = Apuestas.Ganancias + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
            End If
            
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
            Call WriteUpdateGold(UserIndex)
        End If
    End With
End Sub

Private Sub HandleInquiryVote(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim opt As Byte
        
        opt = .incomingData.ReadByte
        
        Call WriteConsoleMsg(UserIndex, ConsultaPopular.doVotar(UserIndex, opt), FontTypeNames.FONTTYPE_GUILD)
    End With
End Sub

Private Sub HandleBankExtractGold(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
        
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        
        Amount = .incomingData.ReadLong
        
        'Dead people can't leave a faction.. they can't talk...
        If .Stats.Muerto Then
            Exit Sub
        End If
        
        If Not .flags.Comerciando Then
            Exit Sub
        End If
        
        '¿El target es un objeto válido?
        If .flags.TargetObjIndex < 1 Then
            Exit Sub
        End If
        
        '¿Es el alijo?
        If ObjData(.flags.TargetObjIndex).Type <> otAlijo Then
            Exit Sub
        End If
        
        Dim Pos As WorldPos
        
        Pos.map = .flags.TargetObjMap
        Pos.x = .flags.TargetObjX
        Pos.y = .flags.TargetObjY
            
        If Distancia(Pos, .Pos) > 5 Then
            Exit Sub
        End If
        
        If Amount > 0 And Amount <= .Stats.BankGld Then
             .Stats.BankGld = .Stats.BankGld - Amount
             .Stats.Gld = .Stats.Gld + Amount
             'Call WriteUpdateGold(UserIndex)
        End If
    End With
End Sub

Private Sub HandleBankDepositGold(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        
        Amount = .incomingData.ReadLong
        
        If .Stats.Muerto Then
            Exit Sub
        End If
        
        If Not .flags.Comerciando Then
            Exit Sub
        End If
        
        '¿El target es un objeto válido?
        If .flags.TargetObjIndex < 1 Then
            Exit Sub
        End If
        
        '¿Es el alijo?
        If ObjData(.flags.TargetObjIndex).Type <> otAlijo Then
            Exit Sub
        End If
        
        Dim Pos As WorldPos
        
        Pos.map = .flags.TargetObjMap
        Pos.x = .flags.TargetObjX
        Pos.y = .flags.TargetObjY
            
        If Distancia(Pos, .Pos) > 5 Then
            Exit Sub
        End If
        
        If Amount > 0 And Amount <= .Stats.Gld Then
            .Stats.BankGld = .Stats.BankGld + Amount
            .Stats.Gld = .Stats.Gld - Amount
            'Call WriteUpdateGold(UserIndex)
        End If
    End With
End Sub

Private Sub HandleDenounce(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Text As String
        
        Text = buffer.ReadASCIIString
        
        If .Counters.Silencio < 1 Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(LCase$(.name) & " DENUNCIA: " & Text, FontTypeNames.FONTTYPE_GUILDMSG))
            Call WriteConsoleMsg(UserIndex, "Denuncia enviada, espere..", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildFundate(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 1 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim Error As String

    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        If HasFound(.name) Then
            Call WriteConsoleMsg(UserIndex, "¡Ya has fundado una guilda, no podés fundar otra!", FontTypeNames.FONTTYPE_INFOBOLD)
            Exit Sub
        End If
        
        If modGuilds.PuedeFundarUnaGuilda(UserIndex, Error) Then
            Call WriteShowGuildFundationForm(UserIndex)
        Else
            Call WriteConsoleMsg(UserIndex, Error, FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub
        
Private Sub HandlePartyKick(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString
        
        If UserPuedeEjecutarComandos(UserIndex) Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call mdParty.ExpulsarDeParty(UserIndex, tUser)
            Else
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                End If
                
                Call WriteConsoleMsg(UserIndex, LCase(UserName) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandlePartySetLeader(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
'On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim rank As Integer
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        UserName = buffer.ReadASCIIString
        If UserPuedeEjecutarComandos(UserIndex) Then
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                'Don't allow users to spoof online GMs
                If (UserDarPrivilegioLevel(UserName) And rank) <= (.flags.Privilegios And rank) Then
                    Call mdParty.TransformarEnLider(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, LCase(UserList(tUser).name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
                End If
                
            Else
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                End If
                Call WriteConsoleMsg(UserIndex, LCase(UserName) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandlePartyAcceptMember(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim rank As Integer
        Dim bUserVivo As Boolean
        
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        UserName = buffer.ReadASCIIString
        If UserList(UserIndex).Stats.Muerto Then
            Call WriteConsoleMsg(UserIndex, "Estás muerto.", FontTypeNames.FONTTYPE_PARTY)
        Else
            bUserVivo = True
        End If
        
        If mdParty.UserPuedeEjecutarComandos(UserIndex) And bUserVivo Then
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                'Validate administrative ranks - don't allow users to spoof online GMs
                If (UserList(tUser).flags.Privilegios And rank) <= (.flags.Privilegios And rank) Then
                    Call mdParty.AprobarIngresoAParty(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No podés incorporar a tu party a personajes de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                End If
                
                'Don't allow users to spoof online GMs
                If (UserDarPrivilegioLevel(UserName) And rank) <= (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(UserIndex, LCase(UserName) & " no ha solicitado ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
                Else
                    Call WriteConsoleMsg(UserIndex, "No podés incorporar a tu party a personajes de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGuildMemberList(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Guild As String
        Dim MemberCount As Integer
        Dim i As Long
        Dim UserName As String
        
        Guild = buffer.ReadASCIIString
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If (InStrB(Guild, "/") > 0) Then
                Guild = Replace(Guild, "/", vbNullString)
            End If
            If (InStrB(Guild, "/") > 0) Then
                Guild = Replace(Guild, "/", vbNullString)
            End If
            
            If Not FileExist(App.Path & "/guilds/" & Guild & "-members.mem") Then
                Call WriteConsoleMsg(UserIndex, "No existe la guilda: " & Guild, FontTypeNames.FONTTYPE_INFO)
            Else
                MemberCount = Val(GetVar(App.Path & "/Guilds/" & Guild & "-Members" & ".mem", "INIT", "NroMembers"))
                
                For i = 1 To MemberCount
                    UserName = GetVar(App.Path & "/Guilds/" & Guild & "-Members" & ".mem", "Members", "Member" & i)
                    
                    Call WriteConsoleMsg(UserIndex, UserName & "<" & Guild & ">", FontTypeNames.FONTTYPE_INFO)
                Next i
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Public Sub HandleSearchObj(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
       
        
        Call buffer.ReadByte
       
        Dim UserObj As String
        Dim tUser As Integer
        Dim rank As Integer
        Dim N As Integer
        Dim i As Integer
       
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
       
        UserObj = buffer.ReadASCIIString
       
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
           
            For i = 1 To UBound(ObjData)
                If InStr(1, Tilde(ObjData(i).name), Tilde(UserObj)) Then
                    Call WriteConsoleMsg(UserIndex, i & " " & ObjData(i).name & ".", FontTypeNames.FONTTYPE_CENTINELA)
                    N = N + 1
                End If
            Next
            If N = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hubo resultados de la busqueda: " & UserObj, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Hubo " & N & " resultados de la busqueda: " & UserObj, FontTypeNames.FONTTYPE_INFO)
            End If
           
        End If
       
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
 
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
   
    'Destroy auxiliar buffer
    Set buffer = Nothing
   
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGMMessage(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Call buffer.ReadByte
        
        Dim Message As String
        
        Message = buffer.ReadASCIIString
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.name, "Mensaje a Gms:" & Message)
        
            If LenB(Message) > 0 Then
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & ": " & Message, FontTypeNames.FONTTYPE_GMMSG))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleShowName(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            .ShowName = Not .ShowName 'Show \ Hide the name
            
            Call RefreshCharStatus(UserIndex)
        End If
    End With
End Sub

Private Sub HandleGoNearby(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        
        UserName = buffer.ReadASCIIString
        
        Dim tIndex As Integer
        Dim x As Long
        Dim y As Long
        Dim i As Long
        Dim Found As Boolean
        
        tIndex = NameIndex(UserName)
        
        'Check the user has enough powers
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
                If tIndex < 1 Then 'existe el usuario destino?
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
                        For x = UserList(tIndex).Pos.x - i To UserList(tIndex).Pos.x + i
                            For y = UserList(tIndex).Pos.y - i To UserList(tIndex).Pos.y + i
                                If maps(UserList(tIndex).Pos.map).mapData(x, y).UserIndex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.map, x, y, True, True) Then
                                        Call WarpUserChar(UserIndex, UserList(tIndex).Pos.map, x, y, True)
                                        Call LogGM(.name, "/IRCERCA " & UserName & " Mapa:" & UserList(tIndex).Pos.map & " X:" & UserList(tIndex).Pos.x & " Y:" & UserList(tIndex).Pos.y)
                                        Found = True
                                        Exit For
                                    End If
                                End If
                            Next y
                            
                            If Found Then
                                Exit For  'Feo, pero hay que abortar 3 fors sin usar GoTo
                            End If
                        Next x
                        
                        If Found Then
                            Exit For  'Feo, pero hay que abortar 3 fors sin usar GoTo
                        End If
                    Next i
                    
                    'No space found??
                    If Not Found Then
                        Call WriteConsoleMsg(UserIndex, "Todos los lugares están ocupados.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleComment(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim comment As String
        comment = buffer.ReadASCIIString
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.name, "Comentario: " & comment)
            Call WriteConsoleMsg(UserIndex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleServerTime(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
    
        If .flags.Privilegios And PlayerType.User Then
            Exit Sub
        End If
        
        Call LogGM(.name, "Hora.")
    End With
    
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & Time & " " & Date, FontTypeNames.FONTTYPE_INFO))
End Sub

Private Sub HandleWhere(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
            If tUser < 1 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (UserList(tUser).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) > 0 Or ((UserList(tUser).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) > 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) > 0) Then
                    Call WriteConsoleMsg(UserIndex, "Ubicación  " & UserName & ": " & UserList(tUser).Pos.map & ", " & UserList(tUser).Pos.x & ", " & UserList(tUser).Pos.y & ".", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, "/Donde " & UserName)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleCreaturesInMap(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim map As Integer
        Dim i, j As Long
        Dim Npccount1, Npccount2 As Integer
        Dim Npccant1() As Integer
        Dim Npccant2() As Integer
        Dim List1() As String
        Dim List2() As String
        
        map = .incomingData.ReadInteger
        
        If .flags.Privilegios And PlayerType.User Then
            Exit Sub
        End If
        
        If MapaValido(map) Then
            For i = 1 To LastNpc
                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If NpcList(i).Pos.map = map Then
                    '¿esta vivo?
                    If NpcList(i).flags.NpcActive And NpcList(i).Hostile = 1 Then
                        If Npccount1 = 0 Then
                            ReDim List1(0) As String
                            ReDim Npccant1(0) As Integer
                            Npccount1 = 1
                            List1(0) = NpcList(i).name & ": (" & NpcList(i).Pos.x & "," & NpcList(i).Pos.y & ")"
                            Npccant1(0) = 1
                        Else
                            For j = 0 To Npccount1 - 1
                                If Left$(List1(j), Len(NpcList(i).name)) = NpcList(i).name Then
                                    List1(j) = List1(j) & ", (" & NpcList(i).Pos.x & "," & NpcList(i).Pos.y & ")"
                                    Npccant1(j) = Npccant1(j) + 1
                                    Exit For
                                End If
                            Next j
                            If j = Npccount1 Then
                                ReDim Preserve List1(0 To Npccount1) As String
                                ReDim Preserve Npccant1(0 To Npccount1) As Integer
                                Npccount1 = Npccount1 + 1
                                List1(j) = NpcList(i).name & ": (" & NpcList(i).Pos.x & "," & NpcList(i).Pos.y & ")"
                                Npccant1(j) = 1
                            End If
                        End If
                    Else
                        If Npccount2 = 0 Then
                            ReDim List2(0) As String
                            ReDim Npccant2(0) As Integer
                            Npccount2 = 1
                            List2(0) = NpcList(i).name & ": (" & NpcList(i).Pos.x & "," & NpcList(i).Pos.y & ")"
                            Npccant2(0) = 1
                        Else
                            For j = 0 To Npccount2 - 1
                                If Left$(List2(j), Len(NpcList(i).name)) = NpcList(i).name Then
                                    List2(j) = List2(j) & ", (" & NpcList(i).Pos.x & "," & NpcList(i).Pos.y & ")"
                                    Npccant2(j) = Npccant2(j) + 1
                                    Exit For
                                End If
                            Next j
                            If j = Npccount2 Then
                                ReDim Preserve List2(0 To Npccount2) As String
                                ReDim Preserve Npccant2(0 To Npccount2) As Integer
                                Npccount2 = Npccount2 + 1
                                List2(j) = NpcList(i).name & ": (" & NpcList(i).Pos.x & "," & NpcList(i).Pos.y & ")"
                                Npccant2(j) = 1
                            End If
                        End If
                    End If
                End If
            Next i
            
            Call WriteConsoleMsg(UserIndex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            If Npccount1 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay NpcS Hostiles.", FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To Npccount1 - 1
                    Call WriteConsoleMsg(UserIndex, Npccant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            Call WriteConsoleMsg(UserIndex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            If Npccount2 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay más NpcS.", FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To Npccount2 - 1
                    Call WriteConsoleMsg(UserIndex, Npccant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            Call LogGM(.name, "Numero enemigos en mapa " & map)
        End If
    End With
End Sub

Private Sub HandleWarpMeToTarget(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim x As Integer
        Dim y As Integer
        
        If .flags.Privilegios And PlayerType.User Then
            Exit Sub
        End If
        
        x = .flags.TargetX
        y = .flags.TargetY
        
        Call FindLegalPos(UserIndex, .flags.TargetMap, x, y)
        Call WarpUserChar(UserIndex, .flags.TargetMap, x, y, True)
        Call LogGM(.name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.map)
    End With
End Sub

Private Sub HandleWarpChar(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 7 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim map As Integer
        Dim x As Integer
        Dim y As Integer
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString
        map = buffer.ReadInteger
        x = buffer.ReadByte
        y = buffer.ReadByte
        
        If Not .flags.Privilegios And PlayerType.User Then
            If MapaValido(map) And LenB(UserName) > 0 Then
                If UCase$(UserName) <> "YO" Then
                    If Not .flags.Privilegios And PlayerType.Consejero Then
                        tUser = NameIndex(UserName)
                    End If
                Else
                    tUser = UserIndex
                End If
            
                If tUser < 1 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                ElseIf InMapBounds(map, x, y) Then
                    Call FindLegalPos(tUser, map, x, y)
                    Call WarpUserChar(tUser, map, x, y, True)
                    If UserIndex <> tUser Then
                        Call WriteConsoleMsg(UserIndex, UserList(tUser).name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.name, "Transportó a " & UserList(tUser).name & " hacia " & "Mapa" & map & " X:" & x & " Y:" & y)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleSilence(ByVal UserIndex As Integer)

    'If UserList(UserIndex).incomingData.length < 3 Then
    'Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
    'EXIT SUB
    'End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim TiempoSilencio As Long
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString
        
        TiempoSilencio = buffer.ReadLong
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
        
            If tUser < 1 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            ElseIf TiempoSilencio > 0 Then
                UserList(tUser).Counters.Silencio = TiempoSilencio
                Call WriteConsoleMsg(UserIndex, "Usuario silenciado por " & TiempoSilencio & " minutos.", FontTypeNames.FONTTYPE_INFO)
                Call WriteShowMessageBox(tUser, "Has sido castigado con el silencio por " & TiempoSilencio & " minutos.")
                Call LogGM(.name, "/silenciar " & UserList(tUser).name & " " & TiempoSilencio)
            
                'Flush the other user's buffer
                Call FlushBuffer(tUser)
            Else
                UserList(tUser).Counters.Silencio = 0
                Call WriteConsoleMsg(UserIndex, "Usuario des-silenciado.", FontTypeNames.FONTTYPE_INFO)
                Call WriteShowMessageBox(tUser, "El efecto del silencio ha desaparecido.")
                Call LogGM(.name, "/DESsilenciar " & UserList(tUser).name)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleSOSShowList(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then
            Exit Sub
        End If
        
        Call WriteShowSOSForm(UserIndex)
    End With
End Sub

Private Sub HandlePartyForm(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        If .PartyIndex > 0 Then
            Call WriteShowPartyForm(UserIndex)
            
        Else
            Call WriteConsoleMsg(UserIndex, "No perteneces a ningún grupo!", FontTypeNames.FONTTYPE_INFOBOLD)
        End If
    End With
End Sub

Private Sub HandleItemUpgrade(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Dim ItemIndex As Integer
        
        
        Call .incomingData.ReadByte
        
        ItemIndex = .incomingData.ReadInteger
        
        If ItemIndex < 1 Then
            Exit Sub
        End If
        
        If Not TieneObjetos(ItemIndex, 1, UserIndex) Then
            Exit Sub
        End If
        
        Call DoUpgrade(UserIndex, ItemIndex)
    End With
End Sub

Private Sub HandleSOSRemove(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        UserName = buffer.ReadASCIIString
        
        If Not .flags.Privilegios And PlayerType.User Then _
            Call Ayuda.Quitar(UserName)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleGoToChar(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim x As Byte
        Dim y As Byte
        
        UserName = buffer.ReadASCIIString
        tUser = NameIndex(UserName)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) > 0 Then
                If tUser < 1 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    x = UserList(tUser).Pos.x
                    y = UserList(tUser).Pos.y + 1
                    Call FindLegalPos(UserIndex, UserList(tUser).Pos.map, x, y)
                    
                    Call WarpUserChar(UserIndex, UserList(tUser).Pos.map, x, y, True)
                    
                    If .flags.AdminInvisible < 1 Then
                        Call FlushBuffer(tUser)
                    End If
                    
                    Call LogGM(.name, "/I " & UserName & " Mapa:" & UserList(tUser).Pos.map & " X:" & UserList(tUser).Pos.x & " Y:" & UserList(tUser).Pos.y)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleInvisible(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then
            Exit Sub
        End If
        
        Call DoAdminInvisible(UserIndex)
        
        Call LogGM(.name, "/INVI")
    End With
End Sub

Private Sub HandleGMPanel(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then
            Exit Sub
        End If
        
        Call WriteShowGMPanelForm(UserIndex)
    End With
End Sub

Private Sub HandleRequestUserList(ByVal UserIndex As Integer)
    
    Dim i As Long
    Dim names() As String
    Dim Count As Long
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        ReDim names(1 To LastUser) As String
        Count = 1
        
        For i = 1 To LastUser
            If (LenB(UserList(i).name) > 0) Then
                If UserList(i).flags.Privilegios And PlayerType.User Then
                    names(Count) = UserList(i).name
                    Count = Count + 1
                End If
            End If
        Next i
        
        If Count > 1 Then Call WriteUserNameList(UserIndex, names(), Count - 1)
    End With
End Sub

Private Sub HandleWorking(ByVal UserIndex As Integer)

    Dim i As Long
    Dim users As String
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        For i = 1 To LastUser
            If UserList(i).flags.Logged And UserList(i).Counters.Trabajando > 0 Then
                users = users & ", " & UserList(i).name
                
                ' Display the user being checked by the centinel
                If UserList(i).flags.CentinelaIndex <> 0 Then _
                    users = users & " (*)"
            End If
        Next i
        
        If LenB(users) > 0 Then
            users = Right$(users, Len(users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios trabajando: " & users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios trabajando.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleHiding(ByVal UserIndex As Integer)

    Dim i As Long
    Dim users As String
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        For i = 1 To LastUser
            If (LenB(UserList(i).name) > 0) And UserList(i).Counters.Ocultando > 0 Then
                users = users & UserList(i).name & ", "
            End If
        Next i
        
        If LenB(users) > 0 Then
            users = Left$(users, Len(users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios ocultandose: " & users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios ocultándose.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

'
'Handles the "Jail" Message.
'
'UserIndex The Index of the user sending the Message.

Private Sub HandleJail(ByVal UserIndex As Integer)


'Last Modification: 05\17\06
'

    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim Reason As String
        Dim jailTime As Byte
        Dim Count As Byte
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString
        Reason = buffer.ReadASCIIString
        jailTime = buffer.ReadByte
        
        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")
        End If
        
        '/carcel nick@motivo@<tiempo>
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (Not .flags.Privilegios And PlayerType.User) > 0 Then
            If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(UserName)
                
                If tUser < 1 Then
                    Call WriteConsoleMsg(UserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                        Call WriteConsoleMsg(UserIndex, "No podés encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf jailTime > 60 Then
                        Call WriteConsoleMsg(UserIndex, "No puedés encarcelar por más de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        If (InStrB(UserName, "/") > 0) Then
                            UserName = Replace(UserName, "/", vbNullString)
                        End If
                        If (InStrB(UserName, "/") > 0) Then
                            UserName = Replace(UserName, "/", vbNullString)
                        End If
                        
                        If User_Exist(UserName) Then
                            Count = Val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(Reason) & " " & Date & " " & Time)
                        End If
                        
                        Call Encarcelar(tUser, jailTime, .name)
                        Call LogGM(.name, " encarceló a " & UserName)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleKillNpc(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then
            Exit Sub
        End If
        
        Dim tNpc As Integer
        Dim auxNpc As Npc
        
        tNpc = .flags.TargetNpc
        
        If tNpc > 0 Then
            Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & NpcList(tNpc).name, FontTypeNames.FONTTYPE_INFO)
            
            auxNpc = NpcList(tNpc)
            Call QuitarNpc(tNpc)
            
            If auxNpc.flags.Respawn > 0 Then
                Call CrearNpc(auxNpc.Numero, auxNpc.Pos.map, auxNpc.Orig)
            End If
            
            .flags.TargetNpc = 0
        Else
            Call WriteConsoleMsg(UserIndex, "Antes debes hacer click sobre el Npc.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleWarnUser(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim Reason As String
        Dim Privs As PlayerType
        Dim Count As Byte
        
        UserName = buffer.ReadASCIIString
        Reason = buffer.ReadASCIIString
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (Not .flags.Privilegios And PlayerType.User) > 0 Then
            If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                Privs = UserDarPrivilegioLevel(UserName)
                
                If Not Privs And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "No podés advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If InStrB(UserName, "/") > 0 Then
                            UserName = Replace(UserName, "/", vbNullString)
                    End If
                    If InStrB(UserName, "/") > 0 Then
                            UserName = Replace(UserName, "/", vbNullString)
                    End If
                    
                    If User_Exist(UserName) Then
                        Count = Val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.name) & ": ADVERTENCIA por: " & LCase$(Reason) & " " & Date & " " & Time)
                        
                        Call WriteConsoleMsg(UserIndex, "Has advertido a " & UserName & ".", FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.name, " advirtio a " & UserName)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleEditChar(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 8 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim opcion As Byte
        Dim Arg1 As String
        Dim Arg2 As String
        Dim válido As Boolean
        Dim LoopC As Byte
        Dim CommandString As String
        Dim N As Byte
        Dim UserCharPath As String
        Dim Var As Long
        
        UserName = Replace(buffer.ReadASCIIString, "+", " ")
        
        If UCase$(UserName) = "YO" Then
            tUser = UserIndex
        Else
            tUser = NameIndex(UserName)
        End If
        
        opcion = buffer.ReadByte
        Arg1 = buffer.ReadASCIIString
        Arg2 = buffer.ReadASCIIString
        
        If .flags.Privilegios And PlayerType.RoleMaster Then
            Select Case .flags.Privilegios And (PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)
                Case PlayerType.Consejero
                    'Los RMs consejeros sólo se pueden editar su head, body y level
                    válido = tUser = UserIndex And _
                            (opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head Or opcion = eEditOptions.eo_Level)
                
                Case PlayerType.SemiDios
                    'Los RMs sólo se pueden editar su level y el head y body de cualquiera
                    válido = (opcion = eEditOptions.eo_Level And tUser = UserIndex) _
                            Or opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head
                
                Case PlayerType.Dios
                    'Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                    'pero si quiere modificar el level sólo lo puede hacer sobre sí mismo
                    válido = (opcion = eEditOptions.eo_Level And tUser = UserIndex) Or _
                            opcion = eEditOptions.eo_Body Or _
                            opcion = eEditOptions.eo_Head Or _
                            opcion = eEditOptions.eo_Class Or _
                            opcion = eEditOptions.eo_Skills Or _
                            opcion = eEditOptions.eo_addGold
            End Select
            
        ElseIf .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then   'Si no es RM debe ser dios para poder usar este comando
            válido = True
        End If

        If válido Then
            UserCharPath = CharPath & UserName & ".chr"
            If tUser < 1 And Not FileExist(UserCharPath) Then
                Call WriteConsoleMsg(UserIndex, "Estás intentando editar un personaje inexistente.", FontTypeNames.FONTTYPE_INFO)
                Call LogGM(.name, "Intentó editar un personaje inexistente.")
            Else
                'For making the Log
                CommandString = "/MOD "
                
                Select Case opcion
                    Case eEditOptions.eo_Gold
                    
                        If Val(Arg1) > MaxOro Then
                            Arg1 = MaxOro
                        End If
                        
                        If tUser < 1 Then 'Esta offline?
                            Call WriteVar(UserCharPath, "STATS", "GLD", Val(Arg1))
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else 'Online
                            UserList(tUser).Stats.Gld = Val(Arg1)
                            Call WriteUpdateGold(tUser)
                        End If
                    
                        'Log it
                        CommandString = CommandString & "ORO "
                
                    Case eEditOptions.eo_Experience
                        If Val(Arg1) > MaxExp Then
                            Arg1 = MaxExp
                        End If
                        
                        If tUser < 1 Then 'Offline
                            Var = GetVar(UserCharPath, "STATS", "EXP")
                            Call WriteVar(UserCharPath, "STATS", "EXP", Var + Val(Arg1))
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else 'Online
                            UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + Val(Arg1)
                            Call WriteUpdateExp(tUser)
                        End If
                        
                        'Log it
                        CommandString = CommandString & "EXP "
                    
                    Case eEditOptions.eo_Body
                        If tUser < 1 Then
                            Call WriteVar(UserCharPath, "INIT", "Body", Arg1)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ChangeUserChar(tUser, Val(Arg1), UserList(tUser).Char.Head, UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.HeadAnim)
                        End If
                        
                        'Log it
                        CommandString = CommandString & "BODY "
                    
                    Case eEditOptions.eo_Head
                        If tUser < 1 Then
                            Call WriteVar(UserCharPath, "INIT", "Head", Arg1)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ChangeUserChar(tUser, UserList(tUser).Char.Body, Val(Arg1), UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.HeadAnim)
                        End If
                        
                        'Log it
                        CommandString = CommandString & "HEAD "
                    
                    Case eEditOptions.eo_Level
                        If Val(Arg1) > STAT_MaxELV Then
                            Arg1 = CStr(STAT_MaxELV)
                            Call WriteConsoleMsg(UserIndex, "No podés tener un nivel superior a " & STAT_MaxELV & ".", FONTTYPE_INFO)
                        End If
                        
                        If tUser < 1 Then 'Offline
                            Call WriteVar(UserCharPath, "STATS", "ELV", Val(Arg1))
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else 'Online
                            UserList(tUser).Stats.Elv = Val(Arg1)
                            Call WriteUpdateUserStats(tUser)
                        End If
                    
                        'Log it
                        CommandString = CommandString & "LEVEL "
                    
                    Case eEditOptions.eo_Class
                        For LoopC = 1 To NUMCLASES
                            If UCase$(ListaClases(LoopC)) = UCase$(Arg1) Then
                                Exit For
                            End If
                        Next LoopC
                            
                        If LoopC > NUMCLASES Then
                            Call WriteConsoleMsg(UserIndex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser < 1 Then 'Offline
                                Call WriteVar(UserCharPath, "INIT", "Clase", LoopC)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else 'Online
                                UserList(tUser).Clase = LoopC
                            End If
                        End If
                    
                        'Log it
                        CommandString = CommandString & "CLASE "
                        
                    Case eEditOptions.eo_Skills
                        For LoopC = 1 To NumSkills
                            If UCase$(Replace$(SkillName(LoopC), " ", "+")) = UCase$(Arg1) Then
                                Exit For
                            End If
                        Next LoopC
                        
                        If LoopC > NumSkills Then
                            Call WriteConsoleMsg(UserIndex, "Skill Inexistente!", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser < 1 Then 'Offline
                                Call WriteVar(UserCharPath, "Skills", "SK" & LoopC, Arg2)
                                Call WriteVar(UserCharPath, "Skills", "EXPSK" & LoopC, 0)
                                
                                If Arg2 < MaxSkillPoints Then
                                    Call WriteVar(UserCharPath, "Skills", "ELUSK" & LoopC, 0)
                                Else
                                    Call WriteVar(UserCharPath, "Skills", "ELUSK" & LoopC, ELU_SKILL_INICIAL * 1.03 ^ Arg2)
                                End If
    
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else 'Online
                                UserList(tUser).Skills.Skill(LoopC).Elv = Val(Arg2)
                                Call CheckEluSkill(tUser, LoopC, True)
                            End If
                        End If
                        
                        'Log it
                        CommandString = CommandString & "SKILLS "
                    
                    Case eEditOptions.eo_SkillPointsLeft
                        If tUser < 1 Then 'Offline
                            Call WriteVar(UserCharPath, "STATS", "FreeLibres", Arg1)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else 'Online
                            UserList(tUser).Skills.NroFree = Val(Arg1)
                            If UserList(tUser).Skills.NroFree > 0 Then
                                Call WriteFreeSkills(tUser)
                            End If
                        End If
                        
                        'Log it
                        CommandString = CommandString & "SKILLSLIBRES "
                    
                    Case eEditOptions.eo_Sex
                        Dim Sex As Byte
                        Sex = IIf(UCase(Arg1) = "MUJER", eGenero.Mujer, 0) 'Mujer?
                        Sex = IIf(UCase(Arg1) = "HOMBRE", eGenero.Hombre, Sex) 'Hombre?
                        
                        If Sex > 0 Then 'Es Hombre o mujer?
                            If tUser < 1 Then 'OffLine
                                Call WriteVar(UserCharPath, "INIT", "Genero", Sex)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else 'Online
                                UserList(tUser).Genero = Sex
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "Genero desconocido. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        
                        'Log it
                        CommandString = CommandString & "SEX "
                    
                    Case eEditOptions.eo_Raza
                        Dim Raza As Byte
                        
                        Arg1 = UCase$(Arg1)
                        Select Case Arg1
                            Case "HUMANO"
                                Raza = eRaza.Humano
                            Case "ELFO"
                                Raza = eRaza.Elfo
                            Case "DROW"
                                Raza = eRaza.Drow
                            Case "ENANO"
                                Raza = eRaza.Enano
                            Case "GNOMO"
                                Raza = eRaza.Gnomo
                            Case Else
                                Raza = 0
                        End Select
                        
                            
                        If Raza = 0 Then
                            Call WriteConsoleMsg(UserIndex, "Raza desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser < 1 Then
                                Call WriteVar(UserCharPath, "INIT", "Raza", Raza)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else
                                UserList(tUser).Raza = Raza
                            End If
                        End If
                            
                        'Log it
                        CommandString = CommandString & "RAZA "
                        
                    Case eEditOptions.eo_addGold
                    
                        Dim BankGold As Long
            
                        If tUser < 1 Then
                            BankGold = GetVar(CharPath & UserName & ".chr", "STATS", "BANCO")
                            Call WriteVar(UserCharPath, "STATS", "BANCO", IIf(BankGold + Val(Arg1) < 1, 0, BankGold + Val(Arg1)))
                            Call WriteConsoleMsg(UserIndex, "Se le ha agregado " & Arg1 & " monedas de oro a " & UserName & ".", FONTTYPE_TALK)
                        Else
                            UserList(tUser).Stats.BankGld = IIf(UserList(tUser).Stats.BankGld + Val(Arg1) < 1, 0, UserList(tUser).Stats.BankGld + Val(Arg1))
                            Call WriteConsoleMsg(tUser, STANDARD_BOUNTY_HUNTER_Message, FONTTYPE_TALK)
                        End If
                        
                        'Log it
                        CommandString = CommandString & "AGREGAR "
                        
                    Case Else
                        Call WriteConsoleMsg(UserIndex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)
                        CommandString = CommandString & "UNKOWN "
                        
                End Select
                
                CommandString = CommandString & Arg1 & " " & Arg2
                Call LogGM(.name, CommandString & " " & UserName)
                
            End If
        End If
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleRequestCharInfo(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
                
        Dim targetName As String
        Dim TargetIndex As Integer
        
        targetName = Replace$(buffer.ReadASCIIString, "+", " ")
        TargetIndex = NameIndex(targetName)
        
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            'is the player offline?
            If TargetIndex < 1 Then
                'don't allow to retrieve administrator's info
                If Not (EsDios(targetName) Or EsAdmin(targetName)) Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, Buscando en Charfile.", FontTypeNames.FONTTYPE_INFO)
                    Call SendUserStatsTxtOFF(UserIndex, targetName)
                End If
            Else
                'don't allow to retrieve administrator's info
                If UserList(TargetIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(UserIndex, TargetIndex)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleRequestCharStats(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        UserName = buffer.ReadASCIIString
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) > 0 Then
            Call LogGM(.name, "/STAT " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser < 1 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo Charfile... ", FontTypeNames.FONTTYPE_INFO)
                
                Call SendUserMiniStatsTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserMiniStatsTxt(UserIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleRequestCharGold(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString
        tUser = NameIndex(UserName)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/BAL " & UserName)
            
            If tUser < 1 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserOROTxtFromChar(UserIndex, UserName)
            Else
                Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " tiene " & UserList(tUser).Stats.BankGld & " en el banco.", FontTypeNames.FONTTYPE_TALK)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleRequestCharInventory(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/INV " & UserName)
            
            If tUser < 1 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo del charfile...", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserInvTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserInvTxt(UserIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleRequestCharBank(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/BOV " & UserName)
            
            If tUser < 1 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserBovedaTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserBovedaTxt(UserIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleRequestCharSkills(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Long
        Dim Message As String
        
        UserName = buffer.ReadASCIIString
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/STATS " & UserName)
            
            If tUser < 1 Then
                If (InStrB(UserName, "/") > 0) Then
                        UserName = Replace(UserName, "/", vbNullString)
                End If
                If (InStrB(UserName, "/") > 0) Then
                        UserName = Replace(UserName, "/", vbNullString)
                End If
                
                For LoopC = 1 To NumSkills
                    Message = Message & "CHAR>" & SkillName(LoopC) & " = " & GetVar(CharPath & UserName & ".chr", "SKILLS", "SK" & LoopC) & vbNewLine
                Next LoopC
                
                Call WriteConsoleMsg(UserIndex, Message & "CHAR> Libres:" & GetVar(CharPath & UserName & ".chr", "STATS", "FreeLIBRES"), FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendSkillsTxt(UserIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleReviveChar(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte
        
        UserName = buffer.ReadASCIIString
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
            Else
                tUser = UserIndex
            End If
            
            If tUser < 1 Then
                If Not User_Exist(UserName) Then
                    Call WriteConsoleMsg(UserIndex, "No existe nadie llamado " & UserName & ".", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call DB_RS_Open("SELECT * FROM people WHERE `name`='" & UserName & "'")
                                        
                    DB_RS!MinHP = DB_RS!MaxHP
                    DB_RS.Update
                    DB_RS.Close
                      
                    Call WriteConsoleMsg(UserIndex, UserName & " revivido.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Else
                With UserList(tUser)
                    'If dead, show him alive (naked).
                    If .Stats.Muerto Then
                        Call RevivirUsuario(tUser)
                        
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).name & " te resucitó.", FontTypeNames.FONTTYPE_INFO)
                        Call WriteConsoleMsg(UserIndex, UserList(tUser).name & " fue resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).name & " te curó.", FontTypeNames.FONTTYPE_INFO)
                        Call WriteConsoleMsg(UserIndex, UserList(tUser).name & " fue curado.", FontTypeNames.FONTTYPE_INFO)
                        .Stats.MinHP = .Stats.MaxHP
                        Call WriteUpdateHP(tUser)
                        Call FlushBuffer(tUser)
                    End If
                End With
                Call LogGM(.name, "Resucito a " & UserName)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleOnlineGM(ByVal UserIndex As Integer)

    Dim i As Long
    Dim list As String
    Dim Priv As PlayerType
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
            Exit Sub
        End If
        
        Priv = PlayerType.Consejero Or PlayerType.SemiDios
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            Priv = Priv Or PlayerType.Dios Or PlayerType.Admin
        End If
        
        For i = 1 To LastUser
            If UserList(i).flags.Logged Then
                If UserList(i).flags.Privilegios And Priv Then _
                    list = list & UserList(i).name & ", "
            End If
        Next i
        
        If LenB(list) > 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(UserIndex, list & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleOnlineMap(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim map As Integer
        map = .incomingData.ReadInteger
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
            Exit Sub
        End If
        
        Dim LoopC As Long
        Dim list As String
        Dim Priv As PlayerType
        
        Priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            Priv = Priv + (PlayerType.Dios Or PlayerType.Admin)
        End If
        
        For LoopC = 1 To LastUser
            If LenB(UserList(LoopC).name) > 0 And UserList(LoopC).Pos.map = map Then
                If UserList(LoopC).flags.Privilegios And Priv Then
                    list = list & UserList(LoopC).name & ", "
                End If
            End If
        Next LoopC
        
        If Len(list) > 2 Then
            list = Left$(list, Len(list) - 2)
        End If
        
        Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)
    End With
    
End Sub

Private Sub HandleKick(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim rank As Integer
        
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        UserName = buffer.ReadASCIIString
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser < 1 Then
                Call WriteConsoleMsg(UserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(UserIndex, "No podés echar a alguien con jerarquía mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " echó a " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
                    Call CloseSocket(tUser)
                    Call LogGM(.name, "Echó a " & UserName)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleExecute(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) > 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "¿¿Estás loco?? ¿¿Cómo vas a piñatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call UserDie(tUser)
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " ha ejecutado a " & UserName & ".", FontTypeNames.FONTTYPE_EJECUCION))
                    Call LogGM(.name, " ejecuto a " & UserName)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "No está online.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleBanChar(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim Reason As String
        
        UserName = buffer.ReadASCIIString
        Reason = buffer.ReadASCIIString
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) > 0 Then
            Call BanCharacter(UserIndex, UserName, Reason)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleUnbanChar(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim cantPenas As Byte
        
        UserName = buffer.ReadASCIIString
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) > 0 Then
            If (InStrB(UserName, "/") > 0) Then
                UserName = Replace(UserName, "/", vbNullString)
            End If
            If (InStrB(UserName, "/") > 0) Then
                UserName = Replace(UserName, "/", vbNullString)
            End If
            
            If Not User_Exist(UserName) Then
                Call WriteConsoleMsg(UserIndex, "Charfile inexistente (no use +).", FontTypeNames.FONTTYPE_INFO)
            Else
                If (Val(GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban")) = 1) Then
                    Call UnBan(UserName)
                
                    'penas
                    cantPenas = Val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.name) & ": UNBAN. " & Date & " " & Time)
                
                    Call LogGM(.name, "/UNBAN a " & UserName)
                    Call WriteConsoleMsg(UserIndex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & " no está baneado. Imposible unbanear.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleNpcFollow(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
            Exit Sub
        End If
        
        If .flags.TargetNpc > 0 Then
            Call DoFollow(.flags.TargetNpc, UserIndex)
            NpcList(.flags.TargetNpc).flags.Inmovilizado = 0
            NpcList(.flags.TargetNpc).flags.Paralizado = 0
            NpcList(.flags.TargetNpc).Contadores.Paralisis = 0
        End If
    End With
End Sub

Private Sub HandleSummonChar(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim x As Byte
        Dim y As Byte
        
        UserName = buffer.ReadASCIIString
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser < 1 Then
                Call WriteConsoleMsg(UserIndex, "El jugador no está online.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) > 0 Or _
                  (UserList(tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.User)) > 0 Then
                    Call WriteConsoleMsg(tUser, .name & " transportó.", FontTypeNames.FONTTYPE_INFO)
                    x = .Pos.x
                    y = .Pos.y + 1
                    Call FindLegalPos(tUser, .Pos.map, x, y)
                    Call WarpUserChar(tUser, .Pos.map, x, y, True)
                    Call LogGM(.name, "/SUM " & UserName & " Map:" & .Pos.map & " X:" & .Pos.x & " Y:" & .Pos.y)
                Else
                    Call WriteConsoleMsg(UserIndex, "No podés invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleSpawnListRequest(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
            Exit Sub
        End If
        
        Call EnviarSpawnList(UserIndex)
    End With
End Sub

Private Sub HandleSpawnCreature(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim Npc As Integer
        Npc = .incomingData.ReadInteger
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If Npc > 0 And Npc <= UBound(Declaraciones.Spawn_List()) Then _
              Call SpawnNpc(Declaraciones.Spawn_List(Npc).NpcIndex, .Pos, True, False)
            
            Call LogGM(.name, "Sumoneo " & Declaraciones.Spawn_List(Npc).NpcName)
        End If
    End With
End Sub

Private Sub HandleResetNpcInventory(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
            Exit Sub
            
        End If
        If .flags.TargetNpc < 1 Then
            Exit Sub
        End If
        
        Call ResetNpcInv(.flags.TargetNpc)
        Call LogGM(.name, "/RESETINV " & NpcList(.flags.TargetNpc).name)
    End With
End Sub

Private Sub HandleCleanWorld(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        Call LimpiarMundo
    End With
End Sub

Private Sub HandleServerMessage(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Message As String
        Message = buffer.ReadASCIIString
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(Message) > 0 Then
                Call LogGM(.name, "Mensaje Broadcast:" & Message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & "> " & Message, FontTypeNames.FONTTYPE_TALK))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleNickToIP(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim Priv As PlayerType
        
        UserName = buffer.ReadASCIIString
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) > 0 Then
            tUser = NameIndex(UserName)
            Call LogGM(.name, "NICK2IP Solicito la Ip de " & UserName)

            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                Priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
                Priv = PlayerType.User
            End If
            
            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And Priv Then
                    Call WriteConsoleMsg(UserIndex, "El Ip de " & UserName & " es " & UserList(tUser).Ip, FontTypeNames.FONTTYPE_INFO)
                    Dim Ip As String
                    Dim Lista As String
                    Dim LoopC As Long
                    Ip = UserList(tUser).Ip
                    For LoopC = 1 To LastUser
                        If UserList(LoopC).Ip = Ip Then
                            If LenB(UserList(LoopC).name) > 0 And UserList(LoopC).flags.Logged Then
                                If UserList(LoopC).flags.Privilegios And Priv Then
                                    Lista = Lista & UserList(LoopC).name & ", "
                                End If
                            End If
                        End If
                    Next LoopC
                    If LenB(Lista) > 0 Then Lista = Left$(Lista, Len(Lista) - 2)
                    Call WriteConsoleMsg(UserIndex, "Los personajes con Ip " & Ip & " son: " & Lista, FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "No hay ningún personaje con ese nick.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleIPToNick(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim Ip As String
        Dim LoopC As Long
        Dim Lista As String
        Dim Priv As PlayerType
        
        Ip = .incomingData.ReadByte & "."
        Ip = Ip & .incomingData.ReadByte & "."
        Ip = Ip & .incomingData.ReadByte & "."
        Ip = Ip & .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        Call LogGM(.name, "IP2NICK Solicito los Nicks de Ip " & Ip)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            Priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
        Else
            Priv = PlayerType.User
        End If

        For LoopC = 1 To LastUser
            If UserList(LoopC).Ip = Ip Then
                If LenB(UserList(LoopC).name) > 0 And UserList(LoopC).flags.Logged Then
                    If UserList(LoopC).flags.Privilegios And Priv Then
                        Lista = Lista & UserList(LoopC).name & ", "
                    End If
                End If
            End If
        Next LoopC
        
        If LenB(Lista) > 0 Then Lista = Left$(Lista, Len(Lista) - 2)
        Call WriteConsoleMsg(UserIndex, "Los personajes con Ip " & Ip & " son: " & Lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

'
'Handles the "GuildOnlineMembers" Message.
'
'UserIndex The Index of the user sending the Message.

Private Sub HandleGuildOnlineMembers(ByVal UserIndex As Integer)

'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12\29\06
'

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim GuildName As String
        Dim tGuild As Integer
        
        GuildName = buffer.ReadASCIIString
        
        If (InStrB(GuildName, "+") > 0) Then
            GuildName = Replace(GuildName, "+", " ")
        End If
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) > 0 Then
            tGuild = Guild_Id(GuildName)
            
            If tGuild > 0 Then
                Call WriteConsoleMsg(UserIndex, "Guilda " & UCase(GuildName) & ": " & _
                  modGuilds.m_ListaDeMiembrosOnline(UserIndex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleTeleportCreate(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim mapa As Integer
        Dim x As Byte
        Dim y As Byte
        Dim Radio As Byte
        
        mapa = .incomingData.ReadInteger
        x = .incomingData.ReadByte
        y = .incomingData.ReadByte
        Radio = .incomingData.ReadByte
        
        Radio = MinimoInt(Radio, 6)

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
            Exit Sub
        End If
        
        Call LogGM(.name, "/CT " & mapa & "," & x & "," & y & "," & Radio)
        
        If Not MapaValido(mapa) Or Not InMapBounds(mapa, x, y) Then
            Exit Sub
        End If
        
        If maps(.Pos.map).mapData(.Pos.x, .Pos.y - 1).ObjInfo.index > 0 Then
            Exit Sub
        End If
        
        If maps(.Pos.map).mapData(.Pos.x, .Pos.y - 1).TileExit.map > 0 Then
            Exit Sub
        End If
        
        If maps(mapa).mapData(x, y).ObjInfo.index > 0 Then
            Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If maps(mapa).mapData(x, y).TileExit.map > 0 Then
            Call WriteConsoleMsg(UserIndex, "No podés crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim ET As Obj
        ET.Amount = 1
        
        'Es el numero en el dat. El indice es el comienzo + el radio, todo harcodeado :(.
        ET.index = TELEP_OBJ_Index + Radio
        
        With maps(.Pos.map).mapData(.Pos.x, .Pos.y - 1)
            .TileExit.map = mapa
            .TileExit.x = x
            .TileExit.y = y
        End With
        
        Call MakeObj(ET, .Pos.map, .Pos.x, .Pos.y - 1)
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_WARP, .Pos.x, .Pos.y))
    End With
End Sub

Private Sub HandleTeleportDestroy(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        Dim mapa As Integer
        Dim x As Byte
        Dim y As Byte
    
        Call .incomingData.ReadByte
        
        '/dt
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
            Exit Sub
        End If
        
        mapa = .flags.TargetMap
        x = .flags.TargetX
        y = .flags.TargetY
        
        If Not InMapBounds(mapa, x, y) Then
            Exit Sub
        End If
        
        With maps(mapa).mapData(x, y)
            If .ObjInfo.index = 0 Then
                Exit Sub
            End If
            
            If ObjData(.ObjInfo.index).Type = otPortal And .TileExit.map > 0 Then
                Call LogGM(UserList(UserIndex).name, "/DT: " & mapa & "," & x & "," & y)
                
                Call EraseObj(mapa, x, y, -1)
                
                If maps(.TileExit.map).mapData(.TileExit.x, .TileExit.y).ObjInfo.index = 651 Then
                    Call EraseObj(.TileExit.map, .TileExit.x, .TileExit.y, -1)
                End If
                
                .TileExit.map = 0
                .TileExit.x = 0
                .TileExit.y = 0
            End If
        End With
    End With
End Sub

Private Sub HandleRainToggle(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Exit Sub
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
            Exit Sub
        End If
        
        Call LogGM(.name, "/LLUVIA")
        Lloviendo = Not Lloviendo
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
    End With
End Sub

Public Sub HandleWeather(ByVal UserIndex As Integer)
 
    Dim Weather As Byte
    
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        Weather = .incomingData.ReadByte
    
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Tiempo = Weather
            Call SendData(SendTarget.ToAll, 0, PrepareMessageWeather())
        End If
    End With
End Sub

'
'Handles the "SetCharDescription" Message.
'
'UserIndex The Index of the user sending the Message.

Private Sub HandleSetCharDescription(ByVal UserIndex As Integer)

'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12\29\06
'

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim tUser As Integer
        Dim Desc As String
        
        Desc = buffer.ReadASCIIString
        
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) > 0 Or (.flags.Privilegios And PlayerType.RoleMaster) > 0 Then
            tUser = .flags.TargetUser
            If tUser > 0 Then
                UserList(tUser).DescRM = Desc
            Else
                Call WriteConsoleMsg(UserIndex, "Haz click sobre un personaje antes.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleForceMP3ToMap(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim mp3ID As Byte
        Dim mapa As Integer
        
        mp3ID = .incomingData.ReadByte
        mapa = .incomingData.ReadInteger
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, 50, 50) Then
                mapa = .Pos.map
            End If
        
            If mp3ID = 0 Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayMP3(MapInfo(.Pos.map).Music))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayMP3(mp3ID))
            End If
        End If
    End With
End Sub

Private Sub HandleForceWAVEToMap(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim waveID As Byte
        Dim mapa As Integer
        Dim x As Byte
        Dim y As Byte
        
        waveID = .incomingData.ReadByte
        mapa = .incomingData.ReadInteger
        x = .incomingData.ReadByte
        y = .incomingData.ReadByte
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
        'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, x, y) Then
                mapa = .Pos.map
                x = .Pos.x
                y = .Pos.y
            End If
            
            'Ponemos el pedido por el GM
            Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayWave(waveID, x, y))
        End If
    End With
End Sub

Private Sub HandlePublicMessage(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
       
        
        Call buffer.ReadByte
       
        Dim Message As String
        Message = buffer.ReadASCIIString
       
        'If .Stats.Gld < 10000 Then
        '    Call WriteConsoleMsg(UserIndex, "El hablar en modo global cuesta 10.000 monedas de oro cada mensaje.", FontTypeNames.FONTTYPE_INFO)
        '    Call .incomingData.CopyBuffer(buffer)
        '    EXIT SUB
        'End If
        
        If .Counters.Silencio > 0 Then
            Call WriteConsoleMsg(UserIndex, "Estás silenciado.", FontTypeNames.FONTTYPE_INFO)
            Call .incomingData.CopyBuffer(buffer)
            Exit Sub
        End If
        
        Call SendData(SendTarget.ToAllButIndex, UserIndex, PrepareMessageChat(Message, eChatType.Glob, , .name))
        
        .Stats.Gld = .Stats.Gld - 10000
        Call WriteUpdateGold(UserIndex)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
 
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
   
    'Destroy auxiliar buffer
    Set buffer = Nothing
   
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Public Sub HandleAuctionCreate(ByVal UserIndex As Integer)
     
    With UserList(UserIndex)
    
        Call .incomingData.ReadByte
        
        If maps(.Pos.map).mapData(.Pos.x, .Pos.y).ObjInfo.index > 0 Then
            If Not maps(.Pos.map).mapData(.Pos.x, .Pos.y).Blocked Then
                If ObjData(maps(.Pos.map).mapData(.Pos.x, .Pos.y).ObjInfo.index).Agarrable Then
                    If ObjData(maps(.Pos.map).mapData(.Pos.x, .Pos.y).ObjInfo.index).Type <> otGuita Then
                        Call Iniciar_Subasta(UserIndex, maps(.Pos.map).mapData(.Pos.x, .Pos.y).ObjInfo.index, maps(.Pos.map).mapData(.Pos.x, .Pos.y).ObjInfo.Amount, 0)
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleAuctionBid(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        Call Ofertar_Subasta(UserIndex, .incomingData.ReadLong)
    End With
End Sub

Public Sub HandleAuctionView(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
               
        With Subasta
            Call WriteConsoleMsg(UserIndex, "[Subasta] " & UserList(.UserIndex).name & " está subastando " & .Objeto.Amount & " " & ObjData(.Objeto.index).name & ".", FontTypeNames.FONTTYPE_INFO)
        End With
    End With
End Sub

Private Sub HandleSendProcessList(ByVal UserIndex As Integer)
 
On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Procesos As String
        Procesos = buffer.ReadASCIIString
        
        Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg("Procesos de" & UserList(UserIndex).name & ": " & Procesos, FontTypeNames.FONTTYPE_INFO))
       
        Call .incomingData.CopyBuffer(buffer)
    End With
   
ErrHandler:
       
End Sub

Private Sub HandleAniadirCompaniero(ByVal UserIndex As Integer)
 
On Error GoTo ErrHandler

    With UserList(UserIndex)
    
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
       
        Dim CompaName As String
        
        CompaName = buffer.ReadASCIIString
       
        Call .incomingData.CopyBuffer(buffer)
        
        If LenB(CompaName) < 1 Then
            Exit Sub
        End If
        
        If CompaName = .name Then
            Exit Sub
        End If
        
        If Not User_Exist(CompaName) Then
            Call WriteConsoleMsg(UserIndex, "No existe nadie llamado " & CompaName & ".", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
             
        If EsCompaniero(UserIndex, CompaName) > 0 Then
            Exit Sub
        End If
        
        Call AgregarCompaniero(UserIndex, CompaName)

    End With
   
ErrHandler:
       
End Sub

Private Sub HandleEliminarCompaniero(ByVal UserIndex As Integer)
 
On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Slot As Byte
        
        Slot = buffer.ReadByte

        Call .incomingData.CopyBuffer(buffer)
        
        If LenB(.Compas.Compa(Slot)) < 0 Then
            Exit Sub
        End If
                        
        Call QuitarCompaniero(UserIndex, Slot)
    End With
   
ErrHandler:
       
End Sub

Private Sub HandleHome(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        If .Stats.Muerto Then
            Call RespawnearUsuario(UserIndex)
        End If
    End With
End Sub

Private Sub HandlePlatformTeleport(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        If maps(.Pos.map).mapData(.Pos.x, .Pos.y).Trigger = eTrigger.EnPlataforma Then
            If ExistePlataforma(.Pos.map) > 0 Then
                Dim mapa As Integer

                mapa = .incomingData.ReadInteger
                
                If mapa > 0 Then
                    Dim Nro As Byte

                    Nro = ExistePlataforma(mapa)
                    
                    If Nro > 0 Then
                        If TienePlataforma(UserIndex, mapa) Then
                            Call WarpUserChar(UserIndex, Plataforma(Nro).map, Plataforma(Nro).x, Plataforma(Nro).y, True)
                            UserList(UserIndex).Counters.EnPlataforma = 1
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub HandleTalkAsNpc(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Message As String
        Message = buffer.ReadASCIIString
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Asegurarse haya un Npc seleccionado
            If .flags.TargetNpc > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNpc, PrepareMessageChatOverHead(Message, NpcList(.flags.TargetNpc).Char.CharIndex, vbWhite))
            Else
                Call WriteConsoleMsg(UserIndex, "Debes seleccionar el Npc por el que quieres hablar antes de usar este comando.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleDestroyAllItemsInArea(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
            Exit Sub
        End If
        
        Dim x As Long
        Dim y As Long
                
        For y = .Pos.y - MinYBorder + 1 To .Pos.y + MinYBorder - 1
            For x = .Pos.x - MinXBorder + 1 To .Pos.x + MinXBorder - 1
                If x > 0 And y > 0 And x < 101 And y < 101 Then
                    If maps(.Pos.map).mapData(x, y).ObjInfo.index > 0 Then
                        If Not EsObjetoFijo(maps(.Pos.map).mapData(x, y).ObjInfo.index) Then
                            Call EraseObj(.Pos.map, x, y, -1)
                        End If
                    End If
                End If
            Next x
        Next y
        
        Call LogGM(UserList(UserIndex).name, "/MASSDEST")
    End With
End Sub

Private Sub HandleMakeDumb(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString
        
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) > 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser < 1 Then
                Call WriteConsoleMsg(UserIndex, "Offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleMakeDumbNoMore(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString
        
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) > 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser < 1 Then
                Call WriteConsoleMsg(UserIndex, "Offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)
                Call FlushBuffer(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleDumpIPTables(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
            Exit Sub
        End If
        
        Call SecurityIp.DumpTables
    End With
End Sub

Private Sub HandleSetTrigger(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim tTrigger As Byte
        Dim tLog As String
        
        tTrigger = .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        If tTrigger > 1 Then
            maps(.Pos.map).mapData(.Pos.x, .Pos.y).Trigger = tTrigger
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.map & " " & .Pos.x & "," & .Pos.y
            
            Call LogGM(.name, tLog)
            Call WriteConsoleMsg(UserIndex, tLog, FontTypeNames.FONTTYPE_INFO)
        End If
    End With

End Sub

Private Sub HandleAskTrigger(ByVal UserIndex As Integer)
    Dim tTrigger As Byte
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        tTrigger = maps(.Pos.map).mapData(.Pos.x, .Pos.y).Trigger
        
        Call LogGM(.name, "Miro el trigger en " & .Pos.map & "," & .Pos.x & "," & .Pos.y & ". Era " & tTrigger)
        
        Call WriteConsoleMsg(UserIndex, _
            "Trigger " & tTrigger & " en mapa " & .Pos.map & " " & .Pos.x & ", " & .Pos.y _
            , FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Private Sub HandleBannedIPList(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        Dim Lista As String
        Dim LoopC As Long
        
        Call LogGM(.name, "/BANIPLIST")
        
        For LoopC = 1 To BanIps.Count
            Lista = Lista & BanIps.Item(LoopC) & ", "
        Next LoopC
        
        If LenB(Lista) > 0 Then
            Lista = Left$(Lista, Len(Lista) - 2)
        End If
        
        Call WriteConsoleMsg(UserIndex, Lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Private Sub HandleBannedIPReload(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        Call BanIpGuardar
        Call BanIpCargar
    End With
End Sub

Private Sub HandleGuildBan(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim GuildName As String
        Dim cantMembers As Integer
        Dim LoopC As Long
        Dim member As String
        Dim Count As Byte
        Dim tIndex As Integer
        Dim tFile As String
        
        GuildName = buffer.ReadASCIIString
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tFile = App.Path & "/guilds/" & GuildName & "-members.mem"
            
            If Not FileExist(tFile) Then
                Call WriteConsoleMsg(UserIndex, "No existe la guilda: " & GuildName, FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " banned al guilda " & UCase$(GuildName), FontTypeNames.FONTTYPE_FIGHT))
                
                'baneamos a los miembros
                Call LogGM(.name, "BANCLAN a " & UCase$(GuildName))
                
                cantMembers = Val(GetVar(tFile, "INIT", "NroMembers"))
                
                For LoopC = 1 To cantMembers
                    member = GetVar(tFile, "Members", "Member" & LoopC)
                    'member es la victima
                    Call Ban(member, "Administracion del servidor", "Guilda Banned")
                    
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("   " & member & "<" & GuildName & "> fue expulsado del servidor.", FontTypeNames.FONTTYPE_FIGHT))
                    
                    tIndex = NameIndex(member)
                    If tIndex > 0 Then
                        'esta online
                        UserList(tIndex).flags.Ban = 1
                        Call CloseSocket(tIndex)
                    End If
                    
                    'ponemos el flag de ban a 1
                    Call WriteVar(CharPath & member & ".chr", "FLAGS", "Ban", "1")
                    'ponemos la pena
                    Count = Val(GetVar(CharPath & member & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & member & ".chr", "PENAS", "Cant", Count + 1)
                    Call WriteVar(CharPath & member & ".chr", "PENAS", "P" & Count + 1, LCase$(.name) & ": BAN AL CLAN: " & GuildName & " " & Date & " " & Time)
                Next LoopC
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleBanIP(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim bannedIp As String
        Dim tUser As Integer
        Dim Reason As String
        Dim i As Long
        
        'Is it by ip??
        If buffer.ReadBoolean Then
            bannedIp = buffer.ReadByte & "."
            bannedIp = bannedIp & buffer.ReadByte & "."
            bannedIp = bannedIp & buffer.ReadByte & "."
            bannedIp = bannedIp & buffer.ReadByte
        Else
            Dim name As String

            name = buffer.ReadASCIIString
        
            tUser = NameIndex(name)
            
            If tUser > 0 Then
                bannedIp = UserList(tUser).Ip
            End If
        End If
        
        Reason = buffer.ReadASCIIString
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If LenB(bannedIp) > 0 Then
                Call LogGM(.name, "/BanIP " & bannedIp & " por " & Reason)
                
                If BanIpBuscar(bannedIp) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "La Ip " & bannedIp & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call BanIpAgrega(bannedIp)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " baneó la Ip " & bannedIp & " por " & Reason, FontTypeNames.FONTTYPE_FIGHT))
                    
                    'Find every player with that Ip and ban him!
                    For i = 1 To LastUser
                        If UserList(i).ConnIDValida Then
                            If UserList(i).Ip = bannedIp Then
                                Call BanCharacter(UserIndex, UserList(i).name, "IP POR " & Reason)
                            End If
                        End If
                    Next i
                End If
            ElseIf tUser < 1 Then
                Call WriteConsoleMsg(UserIndex, name & " no está online.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleUnbanIP(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim bannedIp As String
        
        bannedIp = .incomingData.ReadByte & "."
        bannedIp = bannedIp & .incomingData.ReadByte & "."
        bannedIp = bannedIp & .incomingData.ReadByte & "."
        bannedIp = bannedIp & .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        If BanIpQuita(bannedIp) Then
            Call WriteConsoleMsg(UserIndex, "La Ip """ & bannedIp & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "La Ip """ & bannedIp & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleCreateItem(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte

        Dim tObj As Integer
        Dim ObjAmount As Long
        tObj = .incomingData.ReadInteger
        ObjAmount = .incomingData.ReadLong
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
            Exit Sub
        End If
        
        Call LogGM(.name, "/CI: " & tObj)
        
        If maps(.Pos.map).mapData(.Pos.x, .Pos.y - 1).TileExit.map > 0 Then
            Exit Sub
        End If
        
        If tObj < 1 Or tObj > NumObjDatas Then
            Exit Sub
        End If
        
        If ObjAmount < 1 Then
            Exit Sub
        End If
        
        'Is the obj not null?
        If LenB(ObjData(tObj).name) = 0 Then
            Exit Sub
        End If
                    
        Dim Objeto As Obj
        
        Objeto.Amount = ObjAmount
        Objeto.index = tObj
        
        Call TirarItemAlPiso(.Pos, Objeto)
    End With
End Sub

Private Sub HandleDestroyItems(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
            Exit Sub
        End If
        
        If maps(.Pos.map).mapData(.Pos.x, .Pos.y).ObjInfo.index = 0 Then
            Exit Sub
        End If
        
        Call LogGM(.name, "/DEST")
        
        If ObjData(maps(.Pos.map).mapData(.Pos.x, .Pos.y).ObjInfo.index).Type = otPortal Then
            Call WriteConsoleMsg(UserIndex, "No puede destruir teleports así. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call EraseObj(.Pos.map, .Pos.x, .Pos.y, -1)
    End With
End Sub

Private Sub HandleForceMP3All(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte

        Dim mp3ID As Byte
        mp3ID = .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
            Exit Sub
        End If
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " broadcast musica: " & mp3ID, FontTypeNames.FONTTYPE_SERVER))
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMP3(mp3ID))
    End With
End Sub

Private Sub HandleForceWAVEAll(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte

        Dim waveID As Byte
        waveID = .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
            Exit Sub
        End If
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID))
    End With
End Sub

Private Sub HandleRemovePunishment(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim punishment As Byte
        Dim NewText As String
        
        UserName = buffer.ReadASCIIString
        punishment = buffer.ReadByte
        NewText = buffer.ReadASCIIString
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
            Else
                If (InStrB(UserName, "/") > 0) Then
                        UserName = Replace(UserName, "/", vbNullString)
                End If
                If (InStrB(UserName, "/") > 0) Then
                        UserName = Replace(UserName, "/", vbNullString)
                End If
                
                If User_Exist(UserName) Then
                    Call LogGM(.name, " borro la pena: " & punishment & "-" & _
                      GetVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment) _
                      & " de " & UserName & " y la cambió por: " & NewText)
                    
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment, LCase$(.name) & ": <" & NewText & "> " & Date & " " & Time)
                    
                    Call WriteConsoleMsg(UserIndex, "Pena modificada.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Private Sub HandleTileBlockedToggle(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
            Exit Sub
        End If
        
        Call LogGM(.name, "/BLOQ")
        
        maps(.Pos.map).mapData(.Pos.x, .Pos.y).Blocked = Not maps(.Pos.map).mapData(.Pos.x, .Pos.y).Blocked
            
        Call Bloquear(True, .Pos.map, .Pos.x, .Pos.y, maps(.Pos.map).mapData(.Pos.x, .Pos.y).Blocked)
        
        If maps(.Pos.map).mapData(.Pos.x, .Pos.y).Blocked Then
            Call WriteConsoleMsg(UserIndex, "Tile X:" & .Pos.x & " Y:" & .Pos.y & " bloqueado.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "Tile X:" & .Pos.x & " Y:" & .Pos.y & " desbloqueado.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleKillNpcNoRespawn(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
            Exit Sub
        End If
        
        If .flags.TargetNpc < 1 Then
            Exit Sub
        End If
        
        Call QuitarNpc(.flags.TargetNpc)
        Call LogGM(.name, "/M " & NpcList(.flags.TargetNpc).name)
    End With
End Sub

Private Sub HandleKillAllNearbyNpcs(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
            Exit Sub
        End If
        
        Dim x As Long
        Dim y As Long
        
        For y = .Pos.y - MinYBorder + 1 To .Pos.y + MinYBorder - 1
            For x = .Pos.x - MinXBorder + 1 To .Pos.x + MinXBorder - 1
                If x > 0 And y > 0 And x < 101 And y < 101 Then
                    If maps(.Pos.map).mapData(x, y).NpcIndex > 0 Then
                        Call QuitarNpc(maps(.Pos.map).mapData(x, y).NpcIndex)
                    End If
                End If
            Next x
        Next y
        Call LogGM(.name, "/MASSKILL")
    End With
    
End Sub

Private Sub HandleLastIP(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim Lista As String
        Dim LoopC As Byte
        Dim Priv As Integer
        Dim validCheck As Boolean
        
        Priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        UserName = buffer.ReadASCIIString
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) > 0 Then
            'Handle special chars
            If (InStrB(UserName, "/") > 0) Then
                UserName = Replace(UserName, "/", vbNullString)
            End If
            If (InStrB(UserName, "/") > 0) Then
                UserName = Replace(UserName, "/", vbNullString)
            End If
            If (InStrB(UserName, "+") > 0) Then
                UserName = Replace(UserName, "+", " ")
            End If
            
            'Only Gods and Admins can see the ips of adminsitrative Chars. All others can be seen by every adminsitrative char.
            If NameIndex(UserName) > 0 Then
                validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And Priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) > 0
            Else
                validCheck = (UserDarPrivilegioLevel(UserName) And Priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) > 0
            End If
            
            If validCheck Then
                Call LogGM(.name, "/LASTIP " & UserName)
                
                If User_Exist(UserName) Then
                    Lista = "Las ultimas IPs con las que " & UserName & " se conectó son:"
                    For LoopC = 1 To 5
                        Lista = Lista & vbNewLine & LoopC & " - " & GetVar(CharPath & UserName & ".chr", "INIT", "LastIP" & LoopC)
                    Next LoopC
                    Call WriteConsoleMsg(UserIndex, Lista, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Charfile """ & UserName & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, UserName & " es de mayor jerarquía que vos.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Public Sub HandleChatColor(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim color As Long
        
        color = RGB(.incomingData.ReadByte, .incomingData.ReadByte, .incomingData.ReadByte)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            .flags.ChatColor = color
        End If
    End With
End Sub

Public Sub HandleIgnored(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible
        End If
    End With
End Sub

Public Sub HandleCheckSlot(ByVal UserIndex As Integer)
'Check one Users Slot in Particular from Inventory

    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim Slot As Byte
        Dim tIndex As Integer
        
        UserName = buffer.ReadASCIIString 'Que UserName?
        Slot = buffer.ReadByte 'Que Slot?
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Dios) Then
            tIndex = NameIndex(UserName)  'Que user Index?
            
            Call LogGM(.name, .name & " Checkeo el slot " & Slot & " de " & UserName)
               
            If tIndex > 0 Then
                If Slot > 0 And Slot <= MaxInvSlots Then
                    If UserList(tIndex).Inv.Obj(Slot).index > 0 Then
                        Call WriteConsoleMsg(UserIndex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Inv.Obj(Slot).index).name & " Cantidad:" & UserList(tIndex).Inv.Obj(Slot).Amount, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No hay Objeto en slot seleccionado", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Slot Inválido.", FontTypeNames.FONTTYPE_TALK)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Public Sub HandleRestart(ByVal UserIndex As Integer)
'Restart the game

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
    
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
            Exit Sub
        End If
            
        'time and Time BUG!
        Call LogGM(.name, .name & " reinicio el mundo")
        
        Call ReiniciarServidor(True)
    End With
End Sub

Public Sub HandleReloadObjs(ByVal UserIndex As Integer)
'Reload the objs
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        Call LogGM(.name, .name & " ha recargado a los objetos.")
        
        Call LoadOBJData
    End With
End Sub

Public Sub HandleReloadSpells(ByVal UserIndex As Integer)
'Reload the spells
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        Call LogGM(.name, .name & " ha recargado los hechizos.")
        
        Call CargarHechizos
    End With
End Sub

Public Sub HandleReloadServidorIni(ByVal UserIndex As Integer)
'Reload the Server`s INI

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        Call LogGM(.name, .name & " ha recargado los Inits.")
        
        Call LoadSini
    End With
End Sub

Public Sub HandleReloadNpcs(ByVal UserIndex As Integer)
'Reload the Server`s Npc

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
         
        Call LogGM(.name, .name & " ha recargado los Npcs.")
    
        Call CargaNpcsDat
    
        Call WriteConsoleMsg(UserIndex, "Npcs.dat recargado.", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Public Sub HandleKickAllChars(ByVal UserIndex As Integer)
'Kick all the chars that are online

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        Call LogGM(.name, .name & " ha echado a todos los personajes.")
        
        Call EcharPjsNoPrivilegiados
    End With
End Sub

Public Sub HandleShowServerForm(ByVal UserIndex As Integer)
'Show the server form

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        Call LogGM(.name, .name & " ha solicitado mostrar el formulario del servidor.")
        Call frmMain.mnuMostrar_Click
    End With
End Sub

Public Sub HandleCleanSOS(ByVal UserIndex As Integer)
'Clean the SOS
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        Call LogGM(.name, .name & " ha borrado los SOS")
        
        Call Ayuda.Reset
    End With
End Sub

Public Sub HandleSaveChars(ByVal UserIndex As Integer)
'Save the Chars
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        Call LogGM(.name, .name & " ha guardado todos los chars")
        
        Call mdParty.ActualizaExperiencias
        Call GuardarUsuarios
    End With
End Sub


''
' Handle the "ChangeMapInfoBackup" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the backup`s info of the map
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim doTheBackUp As Boolean
        
        doTheBackUp = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.name, .name & " ha cambiado la información sobre el BackUp.")
        
        'Change the boolean to byte in a fast way
        If doTheBackUp Then
            MapInfo(.Pos.map).BackUp = 1
        Else
            MapInfo(.Pos.map).BackUp = 0
        End If
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.map & ".dat", "Mapa" & .Pos.map, "backup", MapInfo(.Pos.map).BackUp)
        
        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " Backup: " & MapInfo(.Pos.map).BackUp, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ChangeMapInfoPK" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the pk`s info of the  map
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim isMapPk As Boolean
        
        isMapPk = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.name, .name & " ha cambiado la información sobre si es PK el mapa.")
        
        MapInfo(.Pos.map).PK = isMapPk
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.map & ".dat", "Mapa" & .Pos.map, "Pk", IIf(isMapPk, "1", "0"))

        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " PK: " & MapInfo(.Pos.map).PK, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Then
                Call LogGM(.name, .name & " ha cambiado la información sobre si es restringido el mapa.")
                
                MapInfo(UserList(UserIndex).Pos.map).restringir = RestrictStringToByte(tStr)
                
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "Restringir", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " Restringido: " & RestrictByteToString(MapInfo(.Pos.map).restringir), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handle the "ChangeMapInfoNoMagic" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'MagiaSinEfecto -> Options: "1" , "0".
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim nomagic As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        nomagic = .incomingData.ReadBoolean
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.name, .name & " ha cambiado la información sobre si está permitido usar la magia el mapa.")
            MapInfo(UserList(UserIndex).Pos.map).MagiaSinEfecto = nomagic
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "MagiaSinEfecto", nomagic)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " MagiaSinEfecto: " & MapInfo(.Pos.map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoNoInvi" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvi(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'InviSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim noinvi As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noinvi = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.name, .name & " ha cambiado la información sobre si está permitido usar la invisibilidad en el mapa.")
            MapInfo(UserList(UserIndex).Pos.map).InviSinEfecto = noinvi
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "InviSinEfecto", noinvi)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " InviSinEfecto: " & MapInfo(.Pos.map).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub
            
''
' Handle the "ChangeMapInfoNoResu" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'ResuSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim noresu As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noresu = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.name, .name & " ha cambiado la información sobre si está permitido usar el resucitar en el mapa.")
            MapInfo(UserList(UserIndex).Pos.map).ResuSinEfecto = noresu
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "ResuSinEfecto", noresu)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " ResuSinEfecto: " & MapInfo(.Pos.map).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoLand" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.name, .name & " ha cambiado la información del terreno del mapa.")
                
                MapInfo(UserList(UserIndex).Pos.map).terreno = TerrainStringToByte(tStr)
                
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "Terreno", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " Terreno: " & TerrainByteToString(MapInfo(.Pos.map).terreno), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el único útil es 'NIEVE' ya que al ingresarlo, la gente muere de frío en el mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handle the "ChangeMapInfoZone" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.name, .name & " ha cambiado la información de la zona del mapa.")
                MapInfo(UserList(UserIndex).Pos.map).Zona = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "Zona", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " Zona: " & MapInfo(.Pos.map).Zona, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el único útil es 'DUNGEON' ya que al ingresarlo, NO se sentirá el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub
            
''
' Handle the "ChangeMapInfoStealNp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoStealNpc(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 25/07/2010
'RoboNpcsPermitido -> Options: "1", "0"
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim RoboNpc As Byte
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        RoboNpc = Val(IIf(.incomingData.ReadBoolean(), 1, 0))
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.name, .name & " ha cambiado la información sobre si está permitido robar npcs en el mapa.")
            
            MapInfo(UserList(UserIndex).Pos.map).RoboNpcsPermitido = RoboNpc
            
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "RoboNpcsPermitido", RoboNpc)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " RoboNpcsPermitido: " & MapInfo(.Pos.map).RoboNpcsPermitido, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub
            
''
' Handle the "ChangeMapInfoNoOcultar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoOcultar(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 18/09/2010
'OcultarSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim NoOcultar As Byte
    Dim mapa As Integer
    
    With UserList(UserIndex)
    
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        NoOcultar = Val(IIf(.incomingData.ReadBoolean(), 1, 0))
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            
            mapa = .Pos.map
            
            Call LogGM(.name, .name & " ha cambiado la información sobre si está permitido ocultarse en el mapa " & mapa & ".")
            
            MapInfo(mapa).OcultarSinEfecto = NoOcultar
            
            Call WriteVar(App.Path & MapPath & "mapa" & mapa & ".dat", "Mapa" & mapa, "OcultarSinEfecto", NoOcultar)
            Call WriteConsoleMsg(UserIndex, "Mapa " & mapa & " OcultarSinEfecto: " & NoOcultar, FontTypeNames.FONTTYPE_INFO)
        End If
        
    End With
    
End Sub
           
''
' Handle the "ChangeMapInfoNoInvocar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvocar(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 18/09/2010
'InvocarSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim NoInvocar As Byte
    Dim mapa As Integer
    
    With UserList(UserIndex)
    
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        NoInvocar = Val(IIf(.incomingData.ReadBoolean(), 1, 0))
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            
            mapa = .Pos.map
            
            Call LogGM(.name, .name & " ha cambiado la información sobre si está permitido invocar en el mapa " & mapa & ".")
            
            MapInfo(mapa).InvocarSinEfecto = NoInvocar
            
            Call WriteVar(App.Path & MapPath & "mapa" & mapa & ".dat", "Mapa" & mapa, "InvocarSinEfecto", NoInvocar)
            Call WriteConsoleMsg(UserIndex, "Mapa " & mapa & " InvocarSinEfecto: " & NoInvocar, FontTypeNames.FONTTYPE_INFO)
        End If
        
    End With
    
End Sub

Public Sub HandleSaveMap(ByVal UserIndex As Integer)
'Saves the map

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        Call LogGM(.name, .name & " ha guardado el mapa " & CStr(.Pos.map))
        
        Call GrabarMapa(.Pos.map, App.Path & "/WorldBackUp/Mapa" & .Pos.map)
        
        Call WriteConsoleMsg(UserIndex, "Mapa Guardado", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Public Sub HandleShowGuildMessages(ByVal UserIndex As Integer)
'Allows admins to read guild Messages

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Guild As String
        
        Guild = buffer.ReadASCIIString
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call modGuilds.GMEscuchaGuilda(UserIndex, Guild)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Public Sub HandleDoBackUp(ByVal UserIndex As Integer)
'Show guilds Messages

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        Call LogGM(.name, .name & " ha hecho un backup")
        
        Call ES.DoBackUp 'Sino lo confunde con la id del paquete
    End With
End Sub

''
' Handle the "ToggleCentinelActivated" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleToggleCentinelActivated(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/26/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Activate or desactivate the Centinel
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        centinelaActivado = Not centinelaActivado
        
        Call ResetCentinelas
        
        If centinelaActivado Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido activado.", FontTypeNames.FONTTYPE_SERVER))
        Else
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido desactivado.", FontTypeNames.FONTTYPE_SERVER))
        End If
    End With
End Sub

Public Sub HandleAlterName(ByVal UserIndex As Integer)
'Change user name

    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        'Reads the userName and newUser Packets
        Dim UserName As String
        Dim newName As String
        Dim changeNameUI As Integer
        Dim Guild_Id As Integer
        
        UserName = buffer.ReadASCIIString
        newName = buffer.ReadASCIIString
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Usar: /aNAME origen@destino", FontTypeNames.FONTTYPE_INFO)
            Else
                changeNameUI = NameIndex(UserName)
                
                If changeNameUI > 0 Then
                    Call WriteConsoleMsg(UserIndex, UserName & " está online, debe desconectarse para el cambio.", FontTypeNames.FONTTYPE_WARNING)
                Else
                    If Not User_Exist(UserName) Then
                        Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " es inexistente ", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Guild_Id = Val(GetVar(CharPath & UserName & ".chr", "GUILD", "Guild_Id"))
                        
                        If Guild_Id > 0 Then
                            Call WriteConsoleMsg(UserIndex, UserName & " pertenece a una guilda, debe salir del mismo con /DEJARGUILDA para ser transferido.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If Not User_Exist(newName) Then
                                Call FileCopy(CharPath & UserName & ".chr", CharPath & UCase$(newName) & ".chr")
                                
                                Call WriteConsoleMsg(UserIndex, "Transferencia exitosa", FontTypeNames.FONTTYPE_INFO)
                                
                                Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                                
                                Dim cantPenas As Byte
                                
                                cantPenas = Val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                                
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", CStr(cantPenas + 1))
                                
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & CStr(cantPenas + 1), LCase$(.name) & ": BAN POR Cambio de nick a " & newName & " " & Date & " " & Time)
                                
                                Call LogGM(.name, "Ha cambiado de nombre al usuario " & UserName & ". Ahora se llama " & newName)
                            Else
                                Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe", FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Public Sub HandleAlterMail(ByVal UserIndex As Integer)
'Change user mail

    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim newMail As String
        
        UserName = buffer.ReadASCIIString
        newMail = buffer.ReadASCIIString
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newMail) = 0 Then
                Call WriteConsoleMsg(UserIndex, "usar /aEmail <pj>-<nuevomail>", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not User_Exist(UserName) Then
                    Call WriteConsoleMsg(UserIndex, "No existe el charfile " & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteVar(CharPath & UserName & ".chr", "CONTACTO", "Email", newMail)
                    Call WriteConsoleMsg(UserIndex, "Email de " & UserName & " cambiado a: " & newMail, FontTypeNames.FONTTYPE_INFO)
                End If
                
                Call LogGM(.name, "Le ha cambiado el mail a " & UserName)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Public Sub HandleAlterPassword(ByVal UserIndex As Integer)
'Change user password

    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim copyFrom As String
        Dim Password As String
        
        UserName = Replace(buffer.ReadASCIIString, "+", " ")
        copyFrom = Replace(buffer.ReadASCIIString, "+", " ")
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "Ha alterado la contraseña de " & UserName)
            
            If LenB(UserName) = 0 Or LenB(copyFrom) = 0 Then
                Call WriteConsoleMsg(UserIndex, "usar /aPASS <pjsinpass>@<pjconpass>", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not User_Exist(UserName) Or Not User_Exist(copyFrom) Then
                    Call WriteConsoleMsg(UserIndex, "Alguno de los PJs no existe " & UserName & "@" & copyFrom, FontTypeNames.FONTTYPE_INFO)
                Else
                    Password = GetVar(CharPath & copyFrom & ".chr", "INIT", "Password")
                    Call WriteVar(CharPath & UserName & ".chr", "INIT", "Password", Password)
                    
                    Call WriteConsoleMsg(UserIndex, "Password de " & UserName & " ha cambiado por la de " & copyFrom, FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Public Sub HandleCreateNpc(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
            Exit Sub
        End If
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)
        
        If NpcIndex > 0 Then
            Call LogGM(.name, "Sumoneo a " & NpcList(NpcIndex).name & " en mapa " & .Pos.map)
        End If
    End With
End Sub

Public Sub HandleCreateNpcWithRespawn(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger
            
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
            Exit Sub
        End If
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)
        
        If NpcIndex > 0 Then
            Call LogGM(.name, "Sumoneo con respawn " & NpcList(NpcIndex).name & " en mapa " & .Pos.map)
        End If
    End With
End Sub

Public Sub HandleServerOpenToUsersToggle(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        If ServerSoloGMs > 0 Then
            Call WriteConsoleMsg(UserIndex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 0
        Else
            Call WriteConsoleMsg(UserIndex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 1
        End If
    End With
End Sub

Public Sub HandleTurnOffServer(ByVal UserIndex As Integer)

    Dim handle As Integer
    
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Exit Sub
        End If
        
        Call LogGM(.name, "/APAGAR")
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " VA A APAGAR EL SERVIDOR!!!", FontTypeNames.FONTTYPE_FIGHT))
        
        'Log
        handle = FreeFile
        Open App.Path & "/logs/main.log" For Append Shared As #handle
        
        Print #handle, Date & " " & Time & " server apagado por " & .name & ". "
        
        Close #handle
        
        Unload frmMain
    End With
End Sub

Public Sub HandleRemoveCharFromGuild(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim Guild_Id As Integer
        
        UserName = buffer.ReadASCIIString
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "/RAJARCLAN " & UserName)
            
            Guild_Id = modGuilds.m_EcharMiembroDeGuilda(UserIndex, UserName)
            
            If Guild_Id = 0 Then
                Call WriteConsoleMsg(UserIndex, "No pertenece a ningúna guilda o es el lider.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, Guild_Id, PrepareMessageConsoleMsg(UserName & " fue expulsado del guilda por los administradores del servidor.", FontTypeNames.FONTTYPE_GUILD))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Public Sub HandleRequestCharMail(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim Mail As String
        
        UserName = buffer.ReadASCIIString
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If User_Exist(UserName) Then
                Mail = GetVar(CharPath & UserName & ".chr", "CONTACTO", "Email")
                
                Call WriteConsoleMsg(UserIndex, "Last Email de " & UserName & ":" & Mail, FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

'
'Handle the "SystemMessage" Message
'
'UserIndex The Index of the user sending the Message

Public Sub HandleSystemMessage(ByVal UserIndex As Integer)

'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12\29\06
'Send a Message to all the users

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim Message As String
        Message = buffer.ReadASCIIString
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "Mensaje de sistema:" & Message)
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(Message))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Public Sub HandleSetMOTD(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim newMOTD As String
        Dim auxiliaryString() As String
        Dim LoopC As Long
        
        newMOTD = buffer.ReadASCIIString
        auxiliaryString = Split(newMOTD, vbNewLine)
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) > 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "Ha fijado un nuevo MOTD")
            
            MaxLines = UBound(auxiliaryString()) + 1
            
            ReDim Motd(1 To MaxLines)
            
            Call WriteVar(DatPath & "motd.ini", "INIT", "NumLines", CStr(MaxLines))
            
            For LoopC = 1 To MaxLines
                Call WriteVar(DatPath & "motd.ini", "Motd", "Line" & CStr(LoopC), auxiliaryString(LoopC - 1))
                
                Motd(LoopC).texto = auxiliaryString(LoopC - 1)
            Next LoopC
            
            Call WriteConsoleMsg(UserIndex, "Se ha cambiado el MOTD con exito", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub

Public Sub HandleChangeMOTD(ByVal UserIndex As Integer)
'Change the MOTD

    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        If (.flags.Privilegios And (PlayerType.RoleMaster Or PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) Then
            Exit Sub
        End If
        
        Dim auxiliaryString As String
        Dim LoopC As Long
        
        For LoopC = LBound(Motd()) To UBound(Motd())
            auxiliaryString = auxiliaryString & Motd(LoopC).texto & vbNewLine
        Next LoopC
        
        If Len(auxiliaryString) >= 2 Then
            If Right$(auxiliaryString, 2) = vbNewLine Then
                auxiliaryString = Left$(auxiliaryString, Len(auxiliaryString) - 2)
            End If
        End If
        
        Call WriteShowMOTDEditionForm(UserIndex, auxiliaryString)
    End With
End Sub

Public Sub HandlePing(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Call WritePong(UserIndex)
    End With
End Sub

Public Sub HandleSetIniVar(ByVal UserIndex As Integer)
'Modify Servidor.ini

    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        
        Call buffer.ReadByte
        
        Dim sLlave As String
        Dim sClave As String
        Dim sValor As String
        
        'Obtengo los parámetros
        sLlave = buffer.ReadASCIIString
        sClave = buffer.ReadASCIIString
        sValor = buffer.ReadASCIIString
    
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            Dim sTmp As String
            'Obtengo el valor según llave y clave
            sTmp = GetVar(ServidorIni, sLlave, sClave)
            
            'Si obtengo un valor escribo en el Servidor.ini
            If LenB(sTmp) Then
                Call WriteVar(ServidorIni, sLlave, sClave, sValor)
                Call LogGM(.name, "Modificó en Servidor.ini (" & sLlave & " " & sClave & ") el valor " & sTmp & " por " & sValor)
                Call WriteConsoleMsg(UserIndex, "Modificó " & sLlave & " " & sClave & " a " & sValor & ". Valor anterior " & sTmp, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "No existe la llave y/o clave", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error > 0 Then
        Err.Raise Error
    End If
End Sub
       
Public Sub WriteLogged(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Logged)
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteRandomName(ByVal UserIndex As Integer, ByVal name As String)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RandomName)
    Call UserList(UserIndex).outgoingData.WriteASCIIString(name)
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUserOfferConfirm(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserOfferConfirm)
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceInit)
    Call UserList(UserIndex).outgoingData.WriteASCIIString(UserList(UserIndex).ComUsu.DestNick)
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceEnd)
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCharSwing(ByVal UserIndex As Integer, ByVal CharIndex As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CharSwing)
        Call .WriteInteger(CharIndex)
    End With
    
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteNpcKillUser(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NpcKillUser)
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBlockedWithShieldOther(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlockedWithShieldOther)
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUserSwing(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserSwing)
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteResuscitationSafeOn(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ResuscitationSafeOn)
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteResuscitationSafeOff(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ResuscitationSafeOff)
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    With UserList(UserIndex)
    
        If Not .flags.Privilegios And PlayerType.User Then
            .Stats.MinSta = .Stats.MaxSta
            Exit Sub
        ElseIf .Stats.MinSta < 1 Then
            .Stats.MinSta = 0
        ElseIf .Stats.MinSta > .Stats.MaxSta Then
            .Stats.MinSta = .Stats.MaxSta
        End If
    
    End With
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSta)
        
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    With UserList(UserIndex)
    
        If Not .flags.Privilegios And PlayerType.User Then
            .Stats.MinMan = .Stats.MaxMan
            Exit Sub
        ElseIf .Stats.MinMan < 1 Then
            .Stats.MinMan = 0
        ElseIf .Stats.MinMan > .Stats.MaxMan Then
            .Stats.MinMan = .Stats.MaxMan
        End If
        
    End With
    
    With UserList(UserIndex).outgoingData
    
        Call .WriteByte(ServerPacketID.UpdateMana)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMan)
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
        
    With UserList(UserIndex)
    
        If Not .flags.Privilegios And PlayerType.User Then
            .Stats.MinHP = .Stats.MaxHP
            Exit Sub
        ElseIf .Stats.MinHP > .Stats.MaxHP Then
            .Stats.MinHP = .Stats.MaxHP
        End If
        
    End With

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHP)
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUpdateGold(ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    If UserList(UserIndex).Stats.Gld < 1 Then
        UserList(UserIndex).Stats.Gld = 0
    End If
    
    If UserList(UserIndex).Stats.Gld > MaxOro Then
        UserList(UserIndex).Stats.Gld = MaxOro
    End If
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(UserIndex).Stats.Gld)
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    With UserList(UserIndex)
        If .Stats.Exp < 1 Then
            .Stats.Exp = 0
        End If
        
        If .Stats.Exp > MaxExp Then
            .Stats.Exp = MaxExp
        End If
    
        If UserList(UserIndex).Stats.Exp >= UserList(UserIndex).Stats.Elu Then
            Call CheckUserLevel(UserIndex)
        End If
    End With
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUpdateStrenghtAndDexterity(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenghtAndDexterity)
        Call .WriteByte(UserList(UserIndex).Stats.Atributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.Atributos(eAtributos.Agilidad))
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUpdateDexterity(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateDexterity)
        Call .WriteByte(UserList(UserIndex).Stats.Atributos(eAtributos.Agilidad))
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


Public Sub WriteUpdateStrenght(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenght)
        Call .WriteByte(UserList(UserIndex).Stats.Atributos(eAtributos.Fuerza))
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal map As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeMap)
        
        Call .WriteInteger(map)
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WritePosUpdate(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PosUpdate)
        Call .WriteByte(UserList(UserIndex).Pos.x)
        Call .WriteByte(UserList(UserIndex).Pos.y)
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteDamage(ByVal UserIndex As Integer, ByVal Char As Integer, ByVal Damage As Integer, ByVal MinHP As Long, ByVal MaxHP As Long, ByVal DamageNpcType As Byte)

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Damage)
    
        Call .WriteInteger(Char)
        
        If MinHP < 1 Then
            Call .WriteByte(0)
        Else
            Call .WriteByte(MinHP * 100 \ MaxHP)
        End If
        
        Call .WriteByte(DamageNpcType)
        Call .WriteInteger(Damage)
    End With

End Sub

Public Sub WriteUserDamaged(ByVal UserIndex As Integer, ByVal Char As Integer, ByVal Damage As Integer, ByVal DamagedType As Byte)

On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserDamaged)
    
        Call .WriteInteger(Char)
        Call .WriteInteger(Damage)
        Call .WriteByte(DamagedType)
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteChatOverHead(ByVal UserIndex As Integer, ByVal Chat As String, ByVal CharIndex As Integer, ByVal color As Long)

On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(Chat, CharIndex, color))
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, ByVal Chat As String, ByVal FontIndex As FontTypeNames)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(Chat, FontIndex))
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteChat(ByVal UserIndex As Integer, ByVal Chat As String, ByVal TipoChat As eChatType, Optional ByVal CharIndex As Integer = 0, Optional ByVal Nombre As String = vbNullString, Optional ByVal Compa As Byte = 0)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageChat(Chat, TipoChat, CharIndex, Nombre, Compa))
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCommerceChat(ByVal UserIndex As Integer, ByVal Chat As String, ByVal FontIndex As FontTypeNames)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareCommerceConsoleMsg(Chat, FontIndex))
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal Message As String)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(Message)
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteForceCharMove(ByVal UserIndex, ByVal Direccion As eHeading)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageForceCharMove(Direccion))
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteObjCreate(ByVal UserIndex As Integer, ByVal GrhIndex As Integer, ByVal ObjType As Byte, ByVal x As Byte, ByVal y As Byte, Optional ByVal ObjName As String = vbNullString, Optional ByVal ObjAmount As Long = 0)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjCreate(GrhIndex, ObjType, x, y, ObjName, ObjAmount))
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteObjDelete(ByVal UserIndex As Integer, ByVal x As Byte, ByVal y As Byte)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjDelete(x, y))
Exit Sub
            
ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBlockPosition(ByVal UserIndex As Integer, ByVal x As Byte, ByVal y As Byte, ByVal Blocked As Boolean)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(x)
        Call .WriteByte(y)
        Call .WriteBoolean(Blocked)
    End With
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WritePlayWave(ByVal UserIndex As Integer, ByVal wave As Byte, ByVal x As Byte, ByVal y As Byte)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, x, y))
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteGuildList(ByVal UserIndex As Integer, ByRef GuildList() As String)

On Error GoTo ErrHandler
    Dim Tmp As String
    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildList)
        
        'Prepare guild name's list
        For i = LBound(GuildList()) To UBound(GuildList())
            Tmp = Tmp & GuildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteAreaChanged(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AreaChanged)
        Call .WriteByte(UserList(UserIndex).Pos.x)
        Call .WriteByte(UserList(UserIndex).Pos.y)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WritePauseToggle(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteRainToggle(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRainToggle())
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCreateFX(ByVal UserIndex As Integer, ByVal x As Byte, ByVal y As Byte, ByVal FX As Integer, ByVal FXLoops As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(x, y, FX, FXLoops))
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserStats)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxMan)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMan)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
        Call .WriteLong(UserList(UserIndex).Stats.Gld)
        Call .WriteByte(UserList(UserIndex).Stats.Elv)
        Call .WriteLong(UserList(UserIndex).Stats.Elu)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
        
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
        Call .WriteByte(UserList(UserIndex).Pos.x)
        Call .WriteByte(UserList(UserIndex).Pos.y)
        
        Call .WriteInteger(UserList(UserIndex).Char.Head)
        Call .WriteInteger(UserList(UserIndex).Char.Body)
        Call .WriteByte(UserList(UserIndex).Char.Heading)
        
        Call .WriteByte(UserList(UserIndex).Char.WeaponAnim)
        Call .WriteByte(UserList(UserIndex).Char.ShieldAnim)
        Call .WriteByte(UserList(UserIndex).Char.HeadAnim)
        
        Call .WriteASCIIString(modGuilds.GuildName(UserList(UserIndex).Guild_Id))
        
        Call .WriteByte(UserList(UserIndex).flags.Privilegios)

    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteSlotMenosUno(ByVal UserIndex As Integer, ByVal Slot As Byte)
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SlotMenosUno)
        Call .WriteByte(Slot)
    End With
End Sub

Private Sub WriteObjEqp(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional ByVal Amount As Integer = 0)

    With UserList(UserIndex).outgoingData
    
        Call .WriteInteger(ObjIndex)
        
        If ObjIndex > 0 Then
            Call .WriteInteger(ObjData(ObjIndex).GrhIndex)
            Call .WriteASCIIString(ObjData(ObjIndex).name)
            Call .WriteLong(ObjData(ObjIndex).Valor)
            Call .WriteByte(ObjData(ObjIndex).Type)
            
            Select Case ObjData(ObjIndex).Type
            
                Case otArma
                    Call .WriteByte(ObjData(ObjIndex).MinHit)
                    Call .WriteByte(ObjData(ObjIndex).MaxHit)
                    
                Case otFlecha
                    Call .WriteByte(ObjData(ObjIndex).MinHit)
                    Call .WriteByte(ObjData(ObjIndex).MaxHit)
                    
                    Call .WriteInteger(Amount)
                    
                Case otArmadura, otCasco, otEscudo
                    Call .WriteByte(ObjData(ObjIndex).MinDef)
                    Call .WriteByte(ObjData(ObjIndex).MaxDef)
                    
                Case otAnillo
                    Call .WriteByte(ObjData(ObjIndex).MinDefM)
                    Call .WriteByte(ObjData(ObjIndex).MaxDefM)
                    
            End Select
            
        End If
    End With
End Sub

Public Sub WriteInventory(ByVal UserIndex As Integer)

    Dim i As Byte
    Dim ObjIndex As Integer
    Dim ObjInfo As ObjData
    
    With UserList(UserIndex).outgoingData
    
        Call .WriteByte(ServerPacketID.Inventory)
        
        Call WriteObjEqp(UserIndex, UserList(UserIndex).Inv.Head)
        Call WriteObjEqp(UserIndex, UserList(UserIndex).Inv.Body)
        Call WriteObjEqp(UserIndex, UserList(UserIndex).Inv.LeftHand)
        Call WriteObjEqp(UserIndex, UserList(UserIndex).Inv.RightHand, UserList(UserIndex).Inv.AmmoAmount)
        Call WriteObjEqp(UserIndex, UserList(UserIndex).Inv.Belt)
        Call WriteObjEqp(UserIndex, UserList(UserIndex).Inv.Ring)
        'Call WriteObjEqp(UserIndex, UserList(UserIndex).Inv.Ship)
        
        If UserList(UserIndex).Inv.NroItems > 0 Then
        
            'Actualiza todos los slots
            For i = 1 To MaxInvSlots
            
                ObjIndex = UserList(UserIndex).Inv.Obj(i).index
            
                If ObjIndex >= LBound(ObjData()) And ObjIndex <= UBound(ObjData()) Then
                                            
                    ObjInfo = ObjData(ObjIndex)
    
                    Call .WriteByte(i)
                    
                    Call .WriteInteger(ObjIndex)
                    Call .WriteInteger(ObjInfo.GrhIndex)
                    Call .WriteASCIIString(ObjInfo.name)
                    Call .WriteInteger(UserList(UserIndex).Inv.Obj(i).Amount)
                    Call .WriteLong(ObjInfo.Valor)
                    
                    Select Case ObjInfo.Type
                        
                        Case otArma
                        
                            If ObjInfo.MinHit > 255 Or ObjInfo.MaxHit > 255 Then
                                Call .WriteByte(ObjInfo.Type + 100)
                                
                                Call .WriteInteger(ObjInfo.MinHit)
                                Call .WriteInteger(ObjInfo.MaxHit)
                            Else
                                Call .WriteByte(ObjInfo.Type)
                                
                                Call .WriteByte(ObjInfo.MinHit)
                                Call .WriteByte(ObjInfo.MaxHit)
                            End If
                            
                            Call .WriteBoolean(ObjInfo.Proyectil)
                                                    
                            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                            SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                            GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                                Call .WriteBoolean(True)
                            Else
                                Call .WriteBoolean(False)
                            End If
                            
                        Case otFlecha
                        
                            If ObjInfo.MinHit > 255 Or ObjInfo.MaxHit > 255 Then
                                Call .WriteByte(ObjInfo.Type + 100)
                                
                                Call .WriteInteger(ObjInfo.MinHit)
                                Call .WriteInteger(ObjInfo.MaxHit)
                            Else
                                Call .WriteByte(ObjInfo.Type)
                                
                                Call .WriteByte(ObjInfo.MinHit)
                                Call .WriteByte(ObjInfo.MaxHit)
                            End If
                        
                            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                            SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                            GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                                Call .WriteBoolean(True)
                            Else
                                Call .WriteBoolean(False)
                            End If
                                                                      
                        Case otBarco
                                                  
                            If ObjInfo.MinHit > 255 Or ObjInfo.MaxHit > 255 Then
                                Call .WriteByte(ObjInfo.Type + 100)
                                
                                Call .WriteInteger(ObjInfo.MinHit)
                                Call .WriteInteger(ObjInfo.MaxHit)
                            Else
                                Call .WriteByte(ObjInfo.Type)
                                
                                Call .WriteByte(ObjInfo.MinHit)
                                Call .WriteByte(ObjInfo.MaxHit)
                            End If
                                                                                     
                        Case otArmadura
                        
                            Call .WriteByte(ObjInfo.Type)
                            
                            Call .WriteByte(ObjInfo.MinDef)
                            Call .WriteByte(ObjInfo.MaxDef)
                        
                            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                            SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                            GuildaPuedeUsarItem(UserIndex, ObjIndex) And _
                            CheckRazaUsaRopa(UserIndex, ObjIndex) Then
                                Call .WriteBoolean(True)
                            Else
                                Call .WriteBoolean(False)
                            End If
                            
                        Case otCasco, otEscudo
                                
                            Call .WriteByte(ObjInfo.Type)
                            
                            Call .WriteByte(ObjInfo.MinDef)
                            Call .WriteByte(ObjInfo.MaxDef)
                
                            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                            SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                            GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                                Call .WriteBoolean(True)
                            Else
                                Call .WriteBoolean(False)
                            End If
                    
                        Case otAnillo
                        
                            Call .WriteByte(ObjInfo.Type)
            
                            Call .WriteByte(ObjInfo.MinDefM)
                            Call .WriteByte(ObjInfo.MaxDefM)
                                                    
                            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                            SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                            GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                                Call .WriteBoolean(True)
                            Else
                                Call .WriteBoolean(False)
                            End If
                            
                        Case Else
                        
                            Call .WriteByte(ObjInfo.Type)
            
                    End Select
                End If
            Next i
        End If
        
        Call .WriteByte(0)
        
    End With

End Sub

Public Sub WriteInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

On Error GoTo ErrHandler

    Dim ObjIndex As Integer

    ObjIndex = UserList(UserIndex).Inv.Obj(Slot).index
    
    If ObjIndex < LBound(ObjData()) Or ObjIndex > UBound(ObjData()) Then
        Exit Sub
    End If

    Dim ObjInfo As ObjData
        
    ObjInfo = ObjData(ObjIndex)

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.InventorySlot)
        Call .WriteByte(Slot)
        
        Call .WriteInteger(ObjIndex)
        Call .WriteInteger(ObjInfo.GrhIndex)
        Call .WriteASCIIString(ObjInfo.name)
        Call .WriteInteger(UserList(UserIndex).Inv.Obj(Slot).Amount)
        Call .WriteLong(ObjInfo.Valor)
        
        Select Case ObjInfo.Type
            
            Case otArma
            
                If ObjInfo.MinHit > 255 Or ObjInfo.MaxHit > 255 Then
                    Call .WriteByte(ObjInfo.Type + 100)
                    
                    Call .WriteInteger(ObjInfo.MinHit)
                    Call .WriteInteger(ObjInfo.MaxHit)
                Else
                    Call .WriteByte(ObjInfo.Type)
                    
                    Call .WriteByte(ObjInfo.MinHit)
                    Call .WriteByte(ObjInfo.MaxHit)
                End If
                
                Call .WriteBoolean(ObjInfo.Proyectil)
                
                If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                    Call .WriteBoolean(True)
                Else
                    Call .WriteBoolean(False)
                End If
                
            Case otFlecha
            
                If ObjInfo.MinHit > 255 Or ObjInfo.MaxHit > 255 Then
                    Call .WriteByte(ObjInfo.Type + 100)
                    
                    Call .WriteInteger(ObjInfo.MinHit)
                    Call .WriteInteger(ObjInfo.MaxHit)
                Else
                    Call .WriteByte(ObjInfo.Type)
                    
                    Call .WriteByte(ObjInfo.MinHit)
                    Call .WriteByte(ObjInfo.MaxHit)
                End If
            
                If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                    Call .WriteBoolean(True)
                Else
                    Call .WriteBoolean(False)
                End If
                
            Case otBarco
                                      
                If ObjInfo.MinHit > 255 Or ObjInfo.MaxHit > 255 Then
                    Call .WriteByte(ObjInfo.Type + 100)
                    
                    Call .WriteInteger(ObjInfo.MinHit)
                    Call .WriteInteger(ObjInfo.MaxHit)
                Else
                    Call .WriteByte(ObjInfo.Type)
                    
                    Call .WriteByte(ObjInfo.MinHit)
                    Call .WriteByte(ObjInfo.MaxHit)
                End If
                                                 
            Case otArmadura
            
                Call .WriteByte(ObjInfo.Type)
                
                Call .WriteByte(ObjInfo.MinDef)
                Call .WriteByte(ObjInfo.MaxDef)
            
                If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                GuildaPuedeUsarItem(UserIndex, ObjIndex) And _
                CheckRazaUsaRopa(UserIndex, ObjIndex) Then
                    Call .WriteBoolean(True)
                Else
                    Call .WriteBoolean(False)
                End If
                                           
            Case otCasco, otEscudo
                    
                Call .WriteByte(ObjInfo.Type)
                
                Call .WriteByte(ObjInfo.MinDef)
                Call .WriteByte(ObjInfo.MaxDef)
                    
                If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                    Call .WriteBoolean(True)
                Else
                    Call .WriteBoolean(False)
                End If
            
            Case otAnillo
            
                Call .WriteByte(ObjInfo.Type)

                Call .WriteByte(ObjInfo.MinDefM)
                Call .WriteByte(ObjInfo.MaxDefM)
                    
                If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                    Call .WriteBoolean(True)
                Else
                    Call .WriteBoolean(False)
                End If
            
            Case Else

                Call .WriteByte(ObjInfo.Type)

        End Select

    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBeltInv(ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    Dim i As Byte
    Dim ObjIndex As Integer
    Dim ObjInfo As ObjData
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BeltInv)
    
        'Actualiza todos los slots
        For i = 1 To MaxBeltSlots
            ObjIndex = UserList(UserIndex).Belt.Obj(i).index
            
            If ObjIndex >= LBound(ObjData()) And ObjIndex <= UBound(ObjData()) Then
                ObjInfo = ObjData(ObjIndex)
            
                With UserList(UserIndex).outgoingData
                    Call .WriteByte(i)
                    
                    Call .WriteInteger(ObjIndex)
                    Call .WriteInteger(ObjInfo.GrhIndex)
                    Call .WriteASCIIString(ObjInfo.name)
                    Call .WriteInteger(UserList(UserIndex).Belt.Obj(i).Amount)
                    Call .WriteLong(ObjInfo.Valor)
                End With
            End If
        Next i
        
        Call .WriteByte(0)
    End With
    
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBeltSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

On Error GoTo ErrHandler

    Dim ObjIndex As Integer

    ObjIndex = UserList(UserIndex).Belt.Obj(Slot).index
    
    'If ObjIndex < LBound(ObjData()) Or ObjIndex > UBound(ObjData()) Then
    '    EXIT SUB
    'End If

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BeltSlot)
        Call .WriteByte(Slot)
        
        Call .WriteInteger(ObjIndex)
        
        If ObjIndex > 0 Then
            Dim ObjInfo As ObjData
                
            ObjInfo = ObjData(ObjIndex)
            
            Call .WriteInteger(ObjInfo.GrhIndex)
            Call .WriteASCIIString(ObjInfo.name)
            Call .WriteInteger(UserList(UserIndex).Belt.Obj(Slot).Amount)
            Call .WriteLong(ObjInfo.Valor)
        End If
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBank(ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    Dim i As Byte
    Dim ObjIndex As Integer
    Dim ObjInfo As ObjData
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Bank)
        
        If UserList(UserIndex).Bank.NroItems > 0 Then
            'Actualiza todos los slots
            For i = 1 To MaxBankSlots
            
                ObjIndex = UserList(UserIndex).Bank.Obj(i).index
            
                If ObjIndex >= LBound(ObjData()) And ObjIndex <= UBound(ObjData()) Then
    
                    ObjInfo = ObjData(ObjIndex)
    
                    Call .WriteByte(i)
                    
                    Call .WriteInteger(ObjIndex)
                    Call .WriteInteger(ObjInfo.GrhIndex)
                    Call .WriteASCIIString(ObjInfo.name)
                    
                    Call .WriteInteger(UserList(UserIndex).Bank.Obj(i).Amount)
                    Call .WriteLong(ObjInfo.Valor)
                    
                    Select Case ObjInfo.Type
                        
                        Case otArma, otFlecha
                        
                            If ObjInfo.MinHit > 255 Or ObjInfo.MaxHit > 255 Then
                                Call .WriteByte(ObjInfo.Type + 100)
                                
                                Call .WriteInteger(ObjInfo.MinHit)
                                Call .WriteInteger(ObjInfo.MaxHit)
                            Else
                                Call .WriteByte(ObjInfo.Type)
                                
                                Call .WriteByte(ObjInfo.MinHit)
                                Call .WriteByte(ObjInfo.MaxHit)
                            End If
                            
                            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                            SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                            GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                                Call .WriteBoolean(True)
                            Else
                                Call .WriteBoolean(False)
                            End If
                                                                          
                        Case otBarco
                                                  
                            If ObjInfo.MinHit > 255 Or ObjInfo.MaxHit > 255 Then
                                Call .WriteByte(ObjInfo.Type + 100)
                                
                                Call .WriteInteger(ObjInfo.MinHit)
                                Call .WriteInteger(ObjInfo.MaxHit)
                            Else
                                Call .WriteByte(ObjInfo.Type)
                                
                                Call .WriteByte(ObjInfo.MinHit)
                                Call .WriteByte(ObjInfo.MaxHit)
                            End If
                            
                        Case otArmadura
                        
                            Call .WriteByte(ObjInfo.Type)
                            
                            Call .WriteByte(ObjInfo.MinDef)
                            Call .WriteByte(ObjInfo.MaxDef)
                                                       
                            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                            SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                            GuildaPuedeUsarItem(UserIndex, ObjIndex) And _
                            CheckRazaUsaRopa(UserIndex, ObjIndex) Then
                                Call .WriteBoolean(True)
                            Else
                                Call .WriteBoolean(False)
                            End If
                                                       
                        Case otCasco, otEscudo
                                
                            Call .WriteByte(ObjInfo.Type)
                            
                            Call .WriteByte(ObjInfo.MinDef)
                            Call .WriteByte(ObjInfo.MaxDef)
                        
                            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                            SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                            GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                                Call .WriteBoolean(True)
                            Else
                                Call .WriteBoolean(False)
                            End If
                        
                        Case otAnillo
                        
                            Call .WriteByte(ObjInfo.Type)
            
                            Call .WriteByte(ObjInfo.MinDefM)
                            Call .WriteByte(ObjInfo.MaxDefM)
    
                            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                            SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                            GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                                Call .WriteBoolean(True)
                            Else
                                Call .WriteBoolean(False)
                            End If
               
                        Case Else
                            Call .WriteByte(ObjInfo.Type)
            
                    End Select
                End If
            Next i
        End If
        
        Call .WriteByte(0)
                        
        Call .WriteLong(UserList(UserIndex).Stats.BankGld)
                        
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If

End Sub

Public Sub WriteBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

On Error GoTo ErrHandler
     
    Dim ObjIndex As Integer
    Dim ObjInfo As ObjData
    
    ObjIndex = UserList(UserIndex).Bank.Obj(Slot).index
    
    If ObjIndex >= LBound(ObjData()) And ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(ObjIndex)
    End If
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BankSlot)
        Call .WriteByte(Slot)
        
        Call .WriteInteger(ObjIndex)
        Call .WriteInteger(ObjInfo.GrhIndex)
        Call .WriteASCIIString(ObjInfo.name)
        Call .WriteInteger(UserList(UserIndex).Bank.Obj(Slot).Amount)
        Call .WriteLong(ObjInfo.Valor)
                
        Select Case ObjInfo.Type
            
            Case otArma, otFlecha
            
                If ObjInfo.MinHit > 255 Or ObjInfo.MaxHit > 255 Then
                    Call .WriteByte(ObjInfo.Type + 100)
                    
                    Call .WriteInteger(ObjInfo.MinHit)
                    Call .WriteInteger(ObjInfo.MaxHit)
                Else
                    Call .WriteByte(ObjInfo.Type)
                    
                    Call .WriteByte(ObjInfo.MinHit)
                    Call .WriteByte(ObjInfo.MaxHit)
                End If
                
                If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                    Call .WriteBoolean(True)
                Else
                    Call .WriteBoolean(False)
                End If
                                                              
            Case otBarco
                                      
                If ObjInfo.MinHit > 255 Or ObjInfo.MaxHit > 255 Then
                    Call .WriteByte(ObjInfo.Type + 100)
                    
                    Call .WriteInteger(ObjInfo.MinHit)
                    Call .WriteInteger(ObjInfo.MaxHit)
                Else
                    Call .WriteByte(ObjInfo.Type)
                    
                    Call .WriteByte(ObjInfo.MinHit)
                    Call .WriteByte(ObjInfo.MaxHit)
                End If
                
            Case otArmadura
            
                Call .WriteByte(ObjInfo.Type)
                
                Call .WriteByte(ObjInfo.MinDef)
                Call .WriteByte(ObjInfo.MaxDef)
                                           
                If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                GuildaPuedeUsarItem(UserIndex, ObjIndex) And _
                CheckRazaUsaRopa(UserIndex, ObjIndex) Then
                    Call .WriteBoolean(True)
                Else
                    Call .WriteBoolean(False)
                End If
                                           
            Case otCasco, otEscudo
                    
                Call .WriteByte(ObjInfo.Type)
                
                Call .WriteByte(ObjInfo.MinDef)
                Call .WriteByte(ObjInfo.MaxDef)
            
                If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                    Call .WriteBoolean(True)
                Else
                    Call .WriteBoolean(False)
                End If
            
            Case otAnillo
            
                Call .WriteByte(ObjInfo.Type)

                Call .WriteByte(ObjInfo.MinDefM)
                Call .WriteByte(ObjInfo.MaxDefM)

                If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                    Call .WriteBoolean(True)
                Else
                    Call .WriteBoolean(False)
                End If
           
                
            Case Else
            
                Call .WriteByte(ObjInfo.Type)

        End Select

    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteNpcInventory(ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    Dim i As Byte
    Dim ObjIndex As Integer
    Dim Amount As Integer

    Dim ObjInfo As ObjData
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.NpcInventory)
        
        If NpcList(UserList(UserIndex).flags.TargetNpc).Inv.NroItems > 0 Then
            
            'Actualiza todos los slots
            For i = 1 To MaxNpcInvSlots
            
                ObjIndex = NpcList(UserList(UserIndex).flags.TargetNpc).Inv.Obj(i).index
            
                If ObjIndex >= LBound(ObjData()) And ObjIndex <= UBound(ObjData()) Then
                    
                    Amount = NpcList(UserList(UserIndex).flags.TargetNpc).Inv.Obj(i).Amount
    
                    ObjInfo = ObjData(ObjIndex)
    
                    Call .WriteByte(i)
                    
                    Call .WriteInteger(ObjIndex)
                    Call .WriteInteger(ObjInfo.GrhIndex)
                    Call .WriteASCIIString(ObjInfo.name)
                    Call .WriteInteger(Amount)
                    Call .WriteLong(ObjInfo.Valor)
                    
                    Select Case ObjInfo.Type
                        
                        Case otArma, otFlecha
                        
                            If ObjInfo.MinHit > 255 Or ObjInfo.MaxHit > 255 Then
                                Call .WriteByte(ObjInfo.Type + 100)
                                
                                Call .WriteInteger(ObjInfo.MinHit)
                                Call .WriteInteger(ObjInfo.MaxHit)
                            Else
                                Call .WriteByte(ObjInfo.Type)
                                
                                Call .WriteByte(ObjInfo.MinHit)
                                Call .WriteByte(ObjInfo.MaxHit)
                            End If
                            
                            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                            SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                            GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                                Call .WriteBoolean(True)
                            Else
                                Call .WriteBoolean(False)
                            End If
                                                                          
                        Case otBarco
                                                  
                            If ObjInfo.MinHit > 255 Or ObjInfo.MaxHit > 255 Then
                                Call .WriteByte(ObjInfo.Type + 100)
                                
                                Call .WriteInteger(ObjInfo.MinHit)
                                Call .WriteInteger(ObjInfo.MaxHit)
                            Else
                                Call .WriteByte(ObjInfo.Type)
                                
                                Call .WriteByte(ObjInfo.MinHit)
                                Call .WriteByte(ObjInfo.MaxHit)
                            End If
                            
                        Case otArmadura
                        
                            Call .WriteByte(ObjInfo.Type)
                            
                            Call .WriteByte(ObjInfo.MinDef)
                            Call .WriteByte(ObjInfo.MaxDef)
                                                       
                            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                            SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                            GuildaPuedeUsarItem(UserIndex, ObjIndex) And _
                            CheckRazaUsaRopa(UserIndex, ObjIndex) Then
                                Call .WriteBoolean(True)
                            Else
                                Call .WriteBoolean(False)
                            End If
                                                       
                        Case otCasco, otEscudo
                                
                            Call .WriteByte(ObjInfo.Type)
                            
                            Call .WriteByte(ObjInfo.MinDef)
                            Call .WriteByte(ObjInfo.MaxDef)
                        
                            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                            SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                            GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                                Call .WriteBoolean(True)
                            Else
                                Call .WriteBoolean(False)
                            End If
                        
                        Case otAnillo
                        
                            Call .WriteByte(ObjInfo.Type)
            
                            Call .WriteByte(ObjInfo.MinDefM)
                            Call .WriteByte(ObjInfo.MaxDefM)
    
                            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                            SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                            GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                                Call .WriteBoolean(True)
                            Else
                                Call .WriteBoolean(False)
                            End If
               
                        Case Else
                            Call .WriteByte(ObjInfo.Type)
            
                    End Select
                End If
            Next i
        End If
        
        Call .WriteByte(0)
        
        Call .WriteASCIIString(NpcList(UserList(UserIndex).flags.TargetNpc).name)
                        
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If

End Sub

Public Sub WriteNpcInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef NpcIndex As Integer)

On Error GoTo ErrHandler
     
    Dim Obj As Obj
    Dim ObjInfo As ObjData

    Obj.index = NpcList(NpcIndex).Inv.Obj(Slot).index
    Obj.Amount = NpcList(NpcIndex).Inv.Obj(Slot).Amount
    
    If Obj.index >= LBound(ObjData()) And Obj.index <= UBound(ObjData()) Then
        ObjInfo = ObjData(Obj.index)
    End If
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.NpcInventorySlot)
        Call .WriteByte(Slot)
        
        Call .WriteInteger(Obj.index)
        Call .WriteInteger(ObjInfo.GrhIndex)
        Call .WriteASCIIString(ObjInfo.name)
        Call .WriteInteger(Obj.Amount)
        Call .WriteLong(ObjInfo.Valor)
                
        Select Case ObjInfo.Type
            
            Case otArma, otFlecha
            
                If ObjInfo.MinHit > 255 Or ObjInfo.MaxHit > 255 Then
                    Call .WriteByte(ObjInfo.Type + 100)
                    
                    Call .WriteInteger(ObjInfo.MinHit)
                    Call .WriteInteger(ObjInfo.MaxHit)
                Else
                    Call .WriteByte(ObjInfo.Type)
                    
                    Call .WriteByte(ObjInfo.MinHit)
                    Call .WriteByte(ObjInfo.MaxHit)
                End If
                
                If ClasePuedeUsarItem(UserIndex, Obj.index) And _
                SexoPuedeUsarItem(UserIndex, Obj.index) And _
                GuildaPuedeUsarItem(UserIndex, Obj.index) Then
                    Call .WriteBoolean(True)
                Else
                    Call .WriteBoolean(False)
                End If
                                                              
            Case otBarco
                                      
                If ObjInfo.MinHit > 255 Or ObjInfo.MaxHit > 255 Then
                    Call .WriteByte(ObjInfo.Type + 100)
                    
                    Call .WriteInteger(ObjInfo.MinHit)
                    Call .WriteInteger(ObjInfo.MaxHit)
                Else
                    Call .WriteByte(ObjInfo.Type)
                    
                    Call .WriteByte(ObjInfo.MinHit)
                    Call .WriteByte(ObjInfo.MaxHit)
                End If
                
            Case otArmadura
            
                Call .WriteByte(ObjInfo.Type)
                
                Call .WriteByte(ObjInfo.MinDef)
                Call .WriteByte(ObjInfo.MaxDef)
                                           
                If ClasePuedeUsarItem(UserIndex, Obj.index) And _
                SexoPuedeUsarItem(UserIndex, Obj.index) And _
                GuildaPuedeUsarItem(UserIndex, Obj.index) And _
                CheckRazaUsaRopa(UserIndex, Obj.index) Then
                    Call .WriteBoolean(True)
                Else
                    Call .WriteBoolean(False)
                End If
                                           
            Case otCasco, otEscudo
                    
                Call .WriteByte(ObjInfo.Type)
                
                Call .WriteByte(ObjInfo.MinDef)
                Call .WriteByte(ObjInfo.MaxDef)
            
                If ClasePuedeUsarItem(UserIndex, Obj.index) And _
                SexoPuedeUsarItem(UserIndex, Obj.index) And _
                GuildaPuedeUsarItem(UserIndex, Obj.index) Then
                    Call .WriteBoolean(True)
                Else
                    Call .WriteBoolean(False)
                End If
            
            Case otAnillo
            
                Call .WriteByte(ObjInfo.Type)

                Call .WriteByte(ObjInfo.MinDefM)
                Call .WriteByte(ObjInfo.MaxDefM)

                If ClasePuedeUsarItem(UserIndex, Obj.index) And _
                SexoPuedeUsarItem(UserIndex, Obj.index) And _
                GuildaPuedeUsarItem(UserIndex, Obj.index) Then
                    Call .WriteBoolean(True)
                Else
                    Call .WriteBoolean(False)
                End If
           
                
            Case Else
            
                Call .WriteByte(ObjInfo.Type)

        End Select

    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteSpells(ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    Dim i As Byte
    Dim Spell As Byte
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Spells)

        'Actualiza todos los slots
        For i = 1 To MaxSpellSlots
            Call .WriteByte(i)
            
            Spell = UserList(UserIndex).Spells.Spell(i)
            
            Call .WriteByte(Spell)
            
            If Spell > 0 Then
                Call .WriteASCIIString(Hechizos(Spell).Nombre)
                Call .WriteBoolean(Hechizos(Spell).MinSkill < UserList(UserIndex).Skills.Skill(eSkill.Magia).Elv)
                Call .WriteInteger(Hechizos(Spell).ManaRequerido)
                Call .WriteInteger(Hechizos(Spell).StaRequerido)
                Call .WriteInteger(Hechizos(Spell).NeedStaff)
            End If
        Next i
        
    End With
    
    Exit Sub
    
ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SpellSlot)
        
        Call .WriteByte(Slot)
        
        Call .WriteASCIIString(Hechizos(UserList(UserIndex).Spells.Spell(Slot)).Nombre)
        Call .WriteByte(UserList(UserIndex).Spells.Spell(Slot))
        Call .WriteBoolean(Hechizos(UserList(UserIndex).Spells.Spell(Slot)).MinSkill < UserList(UserIndex).Skills.Skill(eSkill.Magia).Elv)
        Call .WriteInteger(Hechizos(UserList(UserIndex).Spells.Spell(Slot)).ManaRequerido)
        Call .WriteInteger(Hechizos(UserList(UserIndex).Spells.Spell(Slot)).StaRequerido)
        Call .WriteInteger(Hechizos(UserList(UserIndex).Spells.Spell(Slot)).NeedStaff)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCompas(ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    Dim i As Byte
    Dim CompaIndex As Integer
    Dim CompaName As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Compas)
        
        For i = 1 To MaxCompaSlots
            
            CompaName = UserList(UserIndex).Compas.Compa(i)
        
            If LenB(CompaName) > 0 Then

                Debug.Print "Compa " & i & " " & CompaName
                
                Call .WriteByte(i)

                Call .WriteASCIIString(CompaName)
                
                CompaIndex = NameIndex(CompaName)
                
                If CompaIndex > 0 Then
                    Call .WriteBoolean(True)

                    Call WriteCompaConnected(CompaIndex, EsCompaniero(CompaIndex, UserList(UserIndex).name))
                Else
                    Call .WriteBoolean(False)
                End If
                
                'Call .WriteInteger(UserList(CompaIndex).Char.Body)
                'Call .WriteInteger(UserList(CompaIndex).Char.Head)
                'Call .WriteByte(UserList(CompaIndex).Char.HeadAnim)
                'Call .WriteByte(UserList(CompaIndex).Char.ShieldAnim)
                'Call .WriteByte(UserList(CompaIndex).Char.WeaponAnim)
            End If
            
        Next i
        
        Debug.Print "Compa O"
        
        Call .WriteByte(0)
        
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteAddCompa(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Online As Boolean, Optional ByVal Added As Boolean = False)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AddCompa)
                
        Call .WriteByte(Slot)
        
        Call .WriteASCIIString(UserList(UserIndex).Compas.Compa(Slot))
            
        Call .WriteBoolean(Online)
        
        Call .WriteBoolean(Added)
        
        'Call .WriteInteger(UserList(CompaIndex).Char.Body)
        'Call .WriteInteger(UserList(CompaIndex).Char.Head)
        'Call .WriteByte(UserList(CompaIndex).Char.HeadAnim)
        'Call .WriteByte(UserList(CompaIndex).Char.ShieldAnim)
        'Call .WriteByte(UserList(CompaIndex).Char.WeaponAnim)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteQuitarCompa(ByVal UserIndex As Integer, ByVal Slot As Byte)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.QuitarCompa)
                
        Call .WriteByte(Slot)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCompaConnected(ByVal UserIndex As Integer, ByVal Slot As Byte)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CompaConnected)
                
        Call .WriteByte(Slot)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCompaDisconnected(ByVal UserIndex As Integer, ByVal Slot As Byte)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CompaDisconnected)
                
        Call .WriteByte(Slot)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteAttributes(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Attributes)
        Call .WriteByte(UserList(UserIndex).Stats.Atributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.Atributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.Atributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.Atributos(eAtributos.Carisma))
        Call .WriteByte(UserList(UserIndex).Stats.Atributos(eAtributos.Constitucion))
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUserPlatforms(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Dim i As Byte
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserPlatforms)
        
        Dim mapa As Integer
                
        For i = 1 To UserList(UserIndex).Plataformas.Nro
            mapa = UserList(UserIndex).Plataformas.Plataforma(i).map
            
            If mapa > 0 Then
                If mapa <> UserList(UserIndex).Pos.map Then
                    Call .WriteInteger(mapa)
                End If
            End If
        Next i
        
        Call .WriteInteger(0)
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBlacksmithWeapons(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ArmasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithWeapons)
        
        For i = 1 To UBound(ArmasHerrero())
            'Can the user create this obj? If so add it to the list....
            If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(UserIndex).Skills.Skill(eSkill.Herreria).Elv Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        'Write the number of objs in the list
        Call .WriteInteger(Count)
        
        'Write the needed data of each obj
        For i = 1 To Count
            Obj = ObjData(ArmasHerrero(validIndexes(i)))
            Call .WriteASCIIString(Obj.name)
            Call .WriteInteger(Obj.GrhIndex)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(ArmasHerrero(validIndexes(i)))
            Call .WriteInteger(Obj.Upgrade)
        Next i
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBlacksmithArmors(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithArmors)
        
        For i = 1 To UBound(ArmadurasHerrero())
            'Can the user create this obj? If so add it to the list....
            If ObjData(ArmadurasHerrero(i)).SkHerreria <= UserList(UserIndex).Skills.Skill(eSkill.Herreria).Elv Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        'Write the number of objs in the list
        Call .WriteInteger(Count)
        
        'Write the needed data of each obj
        For i = 1 To Count
            Obj = ObjData(ArmadurasHerrero(validIndexes(i)))
            Call .WriteASCIIString(Obj.name)
            Call .WriteInteger(Obj.GrhIndex)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(ArmadurasHerrero(validIndexes(i)))
            Call .WriteInteger(Obj.Upgrade)
        Next i
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCarpenterObjs(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ObjCarpintero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CarpenterObjs)
        
        For i = 1 To UBound(ObjCarpintero())
            'Can the user create this obj? If so add it to the list....
            If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Skills.Skill(eSkill.Carpinteria).Elv Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        'Write the number of objs in the list
        Call .WriteInteger(Count)
        
        'Write the needed data of each obj
        For i = 1 To Count
            Obj = ObjData(ObjCarpintero(validIndexes(i)))
            Call .WriteASCIIString(Obj.name)
            Call .WriteInteger(Obj.GrhIndex)
            Call .WriteInteger(Obj.Madera)
            Call .WriteInteger(Obj.MaderaElfica)
            Call .WriteInteger(ObjCarpintero(validIndexes(i)))
            Call .WriteInteger(Obj.Upgrade)
        Next i
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteRestOK(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RestOK)
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal Message As String)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(Message))
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBlind(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Blind)
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteDumb(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Dumb)
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSignal)
        Call .WriteASCIIString(ObjData(ObjIndex).texto)
        Call .WriteInteger(ObjData(ObjIndex).GrhSecundario)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
        Call .WriteByte(UserList(UserIndex).Stats.MinSed)
        Call .WriteByte(UserList(UserIndex).Stats.MinHam)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteMiniStats(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MiniStats)
        
        Call .WriteLong(UserList(UserIndex).Stats.Matados)
        Call .WriteLong(UserList(UserIndex).Stats.Muertes)
        
        Call .WriteLong(UserList(UserIndex).Stats.NpcMatados)
        
        Call .WriteByte(UserList(UserIndex).Clase)
        Call .WriteLong(UserList(UserIndex).Counters.Pena)
        Call .WriteLong(UserList(UserIndex).Counters.Silencio)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteSkillUp(ByVal UserIndex As Integer, ByVal Skill As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SkillUp)
        Call .WriteByte(Skill)
        Call .WriteByte(UserList(UserIndex).Skills.Skill(Skill).Elv)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal Pts As Byte, ByVal AumentoHP As Byte, _
    ByVal AumentoSTA As Byte, ByVal AumentoMANA As Byte, ByVal AumentoHIT As Byte)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.LevelUp)
        Call .WriteByte(Pts)
        Call .WriteByte(AumentoHP)
        Call .WriteByte(AumentoSTA)
        Call .WriteByte(AumentoMANA)
        Call .WriteByte(AumentoHIT)
        Call .WriteLong(UserList(UserIndex).Stats.Elu)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal Invisible As Boolean)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, Invisible))
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteSetParalized(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal Paralizado As Boolean)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetParalized(CharIndex, Paralizado))
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteSkills(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Dim i As Byte
    
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.Skills)
        
        For i = 1 To NumSkills
            Call .outgoingData.WriteByte(.Skills.Skill(i).Elv)
            
            If .Skills.Skill(i).Elv < MaxSkillPoints Then
                If .Skills.Skill(i).Elu > 0 And .Skills.Skill(i).Exp < .Skills.Skill(i).Elu Then
                    Call .outgoingData.WriteByte(Int(.Skills.Skill(i).Exp * 100 \ .Skills.Skill(i).Elu))
                Else
                    Call .outgoingData.WriteByte(0)
                End If
            Else
                Call .outgoingData.WriteByte(0)
            End If
        Next i
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteFreeSkills(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.FreeSkills)
        Call .outgoingData.WriteInteger(.Skills.NroFree)
    End With
End Sub

Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

On Error GoTo ErrHandler
    Dim i As Long
    Dim str As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.TrainerCreatureList)
        
        For i = 1 To NpcList(NpcIndex).NroCriaturas
            str = str & NpcList(NpcIndex).Criaturas(i).NpcName & SEPARATOR
        Next i
        
        If LenB(str) > 0 Then _
            str = Left$(str, Len(str) - 1)
        
        Call .WriteASCIIString(str)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

'UserIndex User to which the Message is intended.
'GuildNews The guild's news.
'enemies The list of the guild's enemies.
'allies The list of the guild's allies.
Public Sub WriteGuildNews(ByVal UserIndex As Integer, ByVal GuildNews As String, ByRef enemies() As String, ByRef allies() As String)

On Error GoTo ErrHandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildNews)
        
        Call .WriteASCIIString(GuildNews)
        
        'Prepare enemies'list
        For i = LBound(enemies()) To UBound(enemies())
            Tmp = Tmp & enemies(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        Tmp = vbNullString
        'Prepare allies'list
        For i = LBound(allies()) To UBound(allies())
            Tmp = Tmp & allies(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteOfferDetails(ByVal UserIndex As Integer, ByVal details As String)

On Error GoTo ErrHandler
    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.OfferDetails)
        
        Call .WriteASCIIString(details)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteAlianceProposalsList(ByVal UserIndex As Integer, ByRef Guilds() As String)

On Error GoTo ErrHandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AlianceProposalsList)
        
        'Prepare guild's list
        For i = LBound(Guilds()) To UBound(Guilds())
            Tmp = Tmp & Guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WritePeaceProposalsList(ByVal UserIndex As Integer, ByRef Guilds() As String)

On Error GoTo ErrHandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PeaceProposalsList)
                
        'Prepare guilds'list
        For i = LBound(Guilds()) To UBound(Guilds())
            Tmp = Tmp & Guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCharInfo(ByVal UserIndex As Integer, ByVal CharName As String, ByVal Race As eRaza, ByVal Class As eClass, _
                            ByVal Gender As eGenero, ByVal level As Byte, ByVal Gold As Long, ByVal Bank As Long, _
                            ByVal PreviousPetitions As String, ByVal CurrentGuild As String, ByVal PreviousGuilds As String, _
                            ByVal PeopleKilled As Long, ByVal Deaths As Long)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CharInfo)
        
        Call .WriteASCIIString(CharName)
        Call .WriteByte(Race)
        Call .WriteByte(Class)
        Call .WriteByte(Gender)
        
        Call .WriteByte(level)
        Call .WriteLong(Gold)
        Call .WriteLong(Bank)
        
        Call .WriteASCIIString(PreviousPetitions)
        Call .WriteASCIIString(CurrentGuild)
        Call .WriteASCIIString(PreviousGuilds)
        
        Call .WriteLong(PeopleKilled)
        Call .WriteLong(Deaths)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteGuildLeaderInfo(ByVal UserIndex As Integer, ByRef GuildList() As String, ByRef MemberList() As String, _
                            ByVal GuildNews As String, ByRef joinRequests() As String)

On Error GoTo ErrHandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildLeaderInfo)
        
        'Prepare guild name's list
        For i = LBound(GuildList()) To UBound(GuildList())
            Tmp = Tmp & GuildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        'Prepare guild member's list
        Tmp = vbNullString
        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        'Store guild news
        Call .WriteASCIIString(GuildNews)
        
        'Prepare the join request's list
        Tmp = vbNullString
        For i = LBound(joinRequests()) To UBound(joinRequests())
            Tmp = Tmp & joinRequests(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteGuildMemberInfo(ByVal UserIndex As Integer, ByRef GuildList() As String, ByRef MemberList() As String)

On Error GoTo ErrHandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildMemberInfo)
        
        'Prepare guild name's list
        For i = LBound(GuildList()) To UBound(GuildList())
            Tmp = Tmp & GuildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        'Prepare guild member's list
        Tmp = vbNullString
        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

'UserIndex User to which the Message is intended.
'GuildName The requested guild's name.
'Founder The requested guild's Founder.
'FoundationDate The requested guild's foundation date.
'Leader The requested guild's current Leader.
'MemberCount The requested guild's member count.
'ElectionsOpen True if the guilda is electing it's new Leader.
'EnemiesCount The requested guild's enemy count.
'AlliesCount The requested guild's ally count.
'GuildDesc The requested guild's description.

Public Sub WriteGuildDetails(ByVal UserIndex As Integer, ByVal GuildName As String, ByVal Founder As String, ByVal FoundationDate As String, _
                            ByVal Leader As String, ByVal MemberCount As Integer, ByVal ElectionsOpen As Boolean, _
                            ByVal EnemiesCount As Integer, ByVal AlliesCount As Integer, ByVal GuildDesc As String)

On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildDetails)
        
        Call .WriteASCIIString(GuildName)
        Call .WriteASCIIString(Founder)
        Call .WriteASCIIString(FoundationDate)
        Call .WriteASCIIString(Leader)
        
        Call .WriteInteger(MemberCount)
        Call .WriteBoolean(ElectionsOpen)
                
        Call .WriteInteger(EnemiesCount)
        Call .WriteInteger(AlliesCount)
        
        Call .WriteASCIIString(GuildDesc)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteShowGuildFundationForm(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGuildFundationForm)
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowUserRequest)
        
        Call .WriteASCIIString(details)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, ByVal OfferSlot As Byte, ByVal ObjIndex As Integer, ByVal Amount As Long)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeUserTradeSlot)
            Call .WriteByte(OfferSlot)
        If ObjIndex > 0 Then
            Call .WriteInteger(ObjIndex)
            Call .WriteASCIIString(ObjData(ObjIndex).name)
            Call .WriteLong(Amount)
            Call .WriteInteger(ObjData(ObjIndex).GrhIndex)
            Call .WriteByte(ObjData(ObjIndex).Type)
            Call .WriteInteger(ObjData(ObjIndex).MaxHit)
            Call .WriteInteger(ObjData(ObjIndex).MinHit)
            Call .WriteInteger(ObjData(ObjIndex).MaxDef)
            Call .WriteInteger(ObjData(ObjIndex).MinDef)
            Call .WriteLong(ObjData(ObjIndex).Valor)
        Else
            Call .WriteInteger(0)
            Call .WriteASCIIString(vbNullString)
            Call .WriteLong(0)
            Call .WriteInteger(0)
            Call .WriteByte(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteLong(0)
        End If
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByRef npcNames() As String)

On Error GoTo ErrHandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SpawnList)
        
        For i = LBound(npcNames()) To UBound(npcNames())
            Tmp = Tmp & npcNames(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

'
'Writes the "ShowSOSForm" Message to the given user's outgoing data buffer.
'
'UserIndex User to which the Message is intended.


Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)


'Last Modification: 05\17\06
'Writes the "ShowSOSForm" Message to the given user's outgoing data buffer

On Error GoTo ErrHandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSOSForm)
        
        For i = 1 To Ayuda.Longitud
            Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i
        
        If LenB(Tmp) > 0 Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


'
'Writes the "ShowSOSForm" Message to the given user's outgoing data buffer.
'
'UserIndex User to which the Message is intended.


Public Sub WriteShowPartyForm(ByVal UserIndex As Integer)

'Author: Budi
'Last Modification: 11\26\09
'Writes the "ShowPartyForm" Message to the given user's outgoing data buffer

On Error GoTo ErrHandler
    Dim i As Long
    Dim Tmp As String
    Dim PI As Integer
    Dim members(PARTY_MaxMEMBERS) As Integer
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowPartyForm)
        
        PI = UserList(UserIndex).PartyIndex
        Call .WriteByte(CByte(Parties(PI).EsPartyLeader(UserIndex)))
        
        If PI > 0 Then
            Call Parties(PI).ObtenerMiembrosOnline(members())
            For i = 1 To PARTY_MaxMEMBERS
                If members(i) > 0 Then
                    Tmp = Tmp & UserList(members(i)).name & " (" & Fix(Parties(PI).MiExperiencia(members(i))) & ")" & SEPARATOR
                End If
            Next i
        End If
        
        If LenB(Tmp) > 0 Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
            
        Call .WriteASCIIString(Tmp)
        Call .WriteLong(Parties(PI).ObtenerExperienciaTotal)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

'
'Writes the "ShowMOTDEditionForm" Message to the given user's outgoing data buffer.
'
'UserIndex User to which the Message is intended.
'currentMOTD The current Message Of The Day.


Public Sub WriteShowMOTDEditionForm(ByVal UserIndex As Integer, ByVal currentMOTD As String)


'Last Modification: 05\17\06
'Writes the "ShowMOTDEditionForm" Message to the given user's outgoing data buffer

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMOTDEditionForm)
        
        Call .WriteASCIIString(currentMOTD)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

'
'Writes the "ShowGMPanelForm" Message to the given user's outgoing data buffer.
'
'UserIndex User to which the Message is intended.


Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)


'Last Modification: 05\17\06
'Writes the "ShowGMPanelForm" Message to the given user's outgoing data buffer

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGMPanelForm)
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUserNameList(ByVal UserIndex As Integer, ByRef userNamesList() As String, ByVal Cant As Integer)

On Error GoTo ErrHandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserNameList)
        
        'Prepare user's names list
        For i = 1 To Cant
            Tmp = Tmp & userNamesList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WritePong(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Pong)
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub FlushBuffer(ByVal UserIndex As Integer)

    Dim sndData As String
    
    With UserList(UserIndex).outgoingData
        If .length = 0 Then
            Exit Sub
        End If
        
        sndData = .ReadASCIIStringFixed(.length)
        
        Call EnviarDatosASlot(UserIndex, sndData)
    End With
End Sub

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, ByVal Invisible As Boolean) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetInvisible)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(Invisible)
        
        PrepareMessageSetInvisible = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageSetParalized(ByVal CharIndex As Integer, ByVal Paralizado As Boolean) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetParalized)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(Paralizado)
        
        PrepareMessageSetParalized = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageCharChangeNick(ByVal CharIndex As Integer, ByVal newNick As String) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharChangeNick)
        
        Call .WriteInteger(CharIndex)
        Call .WriteASCIIString(newNick)
        
        PrepareMessageCharChangeNick = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageChatOverHead(ByVal Chat As String, ByVal CharIndex As Integer, ByVal color As Long) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChatOverHead)
        Call .WriteASCIIString(Chat)
        Call .WriteInteger(CharIndex)
        
        'Write rgb channels and save one byte from long :D
        Call .WriteByte(color And &HFF)
        Call .WriteByte((color And &HFF00&) \ &H100&)
        Call .WriteByte((color And &HFF0000) \ &H10000)
        
        PrepareMessageChatOverHead = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageDeleteChatOverHead(ByVal CharIndex As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.DeleteChatOverHead)
        
        Call .WriteInteger(CharIndex)
        
        PrepareMessageDeleteChatOverHead = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageConsoleMsg(ByVal Chat As String, ByVal FontIndex As FontTypeNames) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ConsoleMsg)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)
        
        PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareCommerceConsoleMsg(ByRef Chat As String, ByVal FontIndex As FontTypeNames) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CommerceChat)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)
        
        PrepareCommerceConsoleMsg = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageCreateFX(ByVal x As Byte, ByVal y As Byte, Optional ByVal FX As Integer = 0, Optional ByVal FXLoops As Integer = 0) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateFX)
        Call .WriteByte(x)
        Call .WriteByte(y)
        Call .WriteInteger(FX)
        
        If FX > 0 Then
            Call .WriteByte(FXLoops)
        End If
        
        PrepareMessageCreateFX = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageCreateCharFX(ByVal CharIndex As Integer, Optional ByVal FX As Integer = 0, Optional ByVal FXLoops As Integer = 0) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateCharFX)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(FX)
        
        If FX > 0 Then
            Call .WriteByte(FXLoops)
        End If

        PrepareMessageCreateCharFX = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessagePlayWave(ByVal wave As Byte, Optional ByVal x As Byte = 0, Optional ByVal y As Byte = 0) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayWave)
        Call .WriteByte(wave)
        Call .WriteByte(x)
        
        If x > 0 Then
            Call .WriteByte(y)
        End If
        
        PrepareMessagePlayWave = .ReadASCIIStringFixed(.length)
    End With
    
End Function

Public Function PrepareMessageChat(ByVal Chat As String, ByVal TipoChat As eChatType, Optional ByVal CharIndex As Integer = 0, Optional ByVal Nombre As String = vbNullString, Optional ByVal Compa As Byte = 0) As String

    With auxiliarBuffer
        
        Select Case TipoChat
            
            Case eChatType.Norm
                Call .WriteByte(ServerPacketID.ChatNormal)
                
                Call .WriteASCIIString(Chat)
                Call .WriteInteger(CharIndex)
            
            Case eChatType.Guil
                Call .WriteByte(ServerPacketID.ChatGuild)

                Call .WriteASCIIString(Chat)
                Call .WriteASCIIString(Nombre)
            
            Case eChatType.Komp
                Call .WriteByte(ServerPacketID.ChatCompa)
                
                Call .WriteASCIIString(Chat)
                Call .WriteByte(Compa)
            
            Case eChatType.Priv
                Call .WriteByte(ServerPacketID.ChatPrivate)
                
                Call .WriteASCIIString(Chat)
                Call .WriteASCIIString(Nombre)
                
            Case eChatType.Glob
                Call .WriteByte(ServerPacketID.ChatGlobal)
                
                Call .WriteASCIIString(Chat)
                Call .WriteASCIIString(Nombre)
                
        End Select
        
        PrepareMessageChat = .ReadASCIIStringFixed(.length)
    End With

End Function

Public Function PrepareMessageShowMessageBox(ByVal Chat As String) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(Chat)
        
        PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.length)
    End With
    
End Function

Public Function PrepareMessagePlayMP3(ByVal FileNumber As Byte) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayMP3)
        Call .WriteByte(FileNumber)
        
        PrepareMessagePlayMP3 = .ReadASCIIStringFixed(.length)
    End With
    
End Function

Public Function PrepareMessagePauseToggle() As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PauseToggle)
        PrepareMessagePauseToggle = .ReadASCIIStringFixed(.length)
    End With
    
End Function

Public Function PrepareMessageRainToggle() As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RainToggle)
        
        PrepareMessageRainToggle = .ReadASCIIStringFixed(.length)
    End With
    
End Function

Public Function PrepareMessageWeather() As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.Weather)
        Call .WriteByte(Tiempo)
        PrepareMessageWeather = .ReadASCIIStringFixed(.length)
    End With
    
End Function

Public Function PrepareMessagePopulation() As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.Population)
        Call .WriteInteger(Poblacion)
        PrepareMessagePopulation = .ReadASCIIStringFixed(.length)
    End With
    
End Function

Public Function PrepareMessageAnimAttack(ByVal CharIndex As Integer) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.AnimAttack)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageAnimAttack = .ReadASCIIStringFixed(.length)
    End With
    
End Function

Public Function PrepareMessageCharMeditate(ByVal CharIndex As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharMeditate)
        Call .WriteInteger(CharIndex)

        PrepareMessageCharMeditate = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageBlockedWithShield(ByVal CharIndex As Integer) As String
    
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.BlockedWithShield)
        Call .WriteInteger(CharIndex)
        PrepareMessageBlockedWithShield = .ReadASCIIStringFixed(.length)
    End With
    
End Function

Public Function PrepareMessageObjDelete(ByVal x As Byte, ByVal y As Byte) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjDelete)
        Call .WriteByte(x)
        Call .WriteByte(y)
        
        PrepareMessageObjDelete = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageBlockPosition(ByVal x As Byte, ByVal y As Byte, ByVal Blocked As Boolean) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(x)
        Call .WriteByte(y)
        Call .WriteBoolean(Blocked)
        
        PrepareMessageBlockPosition = .ReadASCIIStringFixed(.length)
    End With
    
End Function

Public Function PrepareMessageObjCreate(ByVal GrhIndex As Integer, ByVal ObjType As Byte, ByVal x As Byte, ByVal y As Byte, Optional ByVal ObjName As String = vbNullString, Optional ByVal ObjAmount As Long = 0) As String
    With auxiliarBuffer
    
        Call .WriteByte(ServerPacketID.ObjCreate)
        Call .WriteByte(x)
        Call .WriteByte(y)
        
        Call .WriteByte(ObjType)

        If ObjType = otGuita Then
            Call .WriteLong(ObjAmount)
            
        ElseIf ObjType = otCuerpoMuerto Then
            Call .WriteASCIIString(ObjName)
            
        Else
            Call .WriteLong(ObjAmount)
            
            Call .WriteInteger(GrhIndex)
                
            If ObjAmount > 0 Then
                Call .WriteASCIIString(ObjName)
            End If
        End If
        
        PrepareMessageObjCreate = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageCharRemove(ByVal CharIndex As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharRemove)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageCharRemove = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageCharCreate(ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal x As Byte, ByVal y As Byte, ByVal Weapon As Integer, ByVal Shield As Integer, _
                                ByVal Helmet As Integer, ByVal name As String, ByVal Guild As String, ByVal Privileges As Byte, ByVal Lvl As Byte) As String
    
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharCreate)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(Body)
        Call .WriteInteger(Head)
        Call .WriteByte(Heading)
        Call .WriteByte(x)
        Call .WriteByte(y)
        Call .WriteByte(Weapon)
        Call .WriteByte(Shield)
        Call .WriteByte(Helmet)
        Call .WriteASCIIString(name)
        Call .WriteASCIIString(Guild)
        Call .WriteByte(Privileges)
        Call .WriteByte(Lvl)
        
        PrepareMessageCharCreate = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageNpcCharCreate(ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal x As Byte, ByVal y As Byte, ByVal name As String, ByVal Lvl As Byte, ByVal MascoIndex As Byte) As String
    
    If Head > 0 Then
        Body = Body + 4000
    End If
    
    If Lvl > 1 Then
        Body = Body + 2000
    End If
    
    If MascoIndex > 0 Then
        Body = Body + 1000
    End If

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.NpcCharCreate)
        
        Call .WriteInteger(CharIndex)
        
        Call .WriteInteger(Body)
        
        If Head > 0 Then
            Call .WriteInteger(Head)
        End If
        
        Call .WriteByte(Heading)
        Call .WriteByte(x)
        Call .WriteByte(y)
        Call .WriteASCIIString(name)

        If Lvl > 1 Then
            Call .WriteByte(Lvl)
        End If
        
        If MascoIndex > 0 Then
            Call .WriteByte(MascoIndex)
        End If
        
        PrepareMessageNpcCharCreate = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageCharChange(ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal Weapon As Integer, ByVal Shield As Integer, _
                                ByVal Helmet As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharChange)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(Body)
        Call .WriteInteger(Head)
        Call .WriteByte(Heading)
        Call .WriteByte(Weapon)
        Call .WriteByte(Shield)
        Call .WriteByte(Helmet)
        
        PrepareMessageCharChange = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageChangeCharHeading(ByVal CharIndex As Integer, ByVal Heading As eHeading) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChangeCharHeading)
        
        Call .WriteInteger(CharIndex)
        Call .WriteByte(Heading)
        
        PrepareMessageChangeCharHeading = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageCharMove(ByVal CharIndex As Integer, ByVal x As Byte, ByVal y As Byte) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharMove)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(x)
        Call .WriteByte(y)
        
        PrepareMessageCharMove = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ForceCharMove)
        Call .WriteByte(Direccion)
        
        PrepareMessageForceCharMove = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, name As String, Guild As String, RelacionGuilda As Byte) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.UpdateTagAndStatus)
        
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
        Call .WriteASCIIString(name)
        Call .WriteASCIIString(Guild)
        Call .WriteByte(RelacionGuilda)
        
        PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageErrorMsg(ByVal Message As String) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ErrorMsg)
        Call .WriteASCIIString(Message)
        
        PrepareMessageErrorMsg = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Sub WriteStopWorking(ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.StopWorking)

Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCancelOfferItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

On Error GoTo ErrHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CancelOfferItem)
        Call .WriteByte(Slot)
    End With
    
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub
