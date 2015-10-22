Attribute VB_Name = "modProtocol"
Option Explicit

'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

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
    Comment                 '/REM
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
    KillNPC                 '/RMATA
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
    banChar                 '/BAN
    UnbanChar               '/UNBAN
    NPCFollow               '/FOLLOW
    SummonChar              '/S
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNpcInv       '/RESETINV
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
    TalkAsNPC               '/TALKAS
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
    KillNPCNoRespawn        '/M
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    ChangeMOTD              '/MOTDCAMBIA
    SetMOTD                 'ZMOTD
    SystemMessage           '/SMSG
    CreateNPC               '/ACC
    CreateNPCWithRespawn    '/RACC
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
    SaveMap                 '/GUARDARMAPA, /GRABARMAPA
    ChangeMapInfoPK         '/MODMAPINFO PK
    ChangeMapInfoBackup     '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand       '/MODMAPINFO TERRENO
    ChangeMapInfoZone       '/MODMAPINFO ZONA
    SaveChars               '/GUARDAR /GRABAR /G
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    KickAllChars            '/ETODOSPJS
    ReloadNPCs              '/RELOADNPCS
    ReloadServidorIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
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
    NPCKillUser             '6
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
    ObjectDelete
    BlockPosition
    PlayMP3
    SoundFX
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
    CarpenterObjects
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
    FreeSkillPts
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
    Connect
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
    RequestAttributes
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
    GuildRequestDetails     'GuildaDETAILS
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

Public Sub HandleIncomingData()

    Dim PacketID As Byte
    
    PacketID = incomingData.PeekByte()

    Debug.Print PacketName(PacketID)

    Select Case PacketID
    
        Case ServerPacketID.Logged                  'LOGGED
            Call HandleLogged
        
        Case ServerPacketID.RandomName
            Call HandleRandomName
            
        Case ServerPacketID.NavigateToggle          'NAVEG
            Call HandleNavigateToggle
                        
        Case ServerPacketID.CommerceChat
            Call HandleCommerceChat
                
        Case ServerPacketID.UserCommerceInit        'INITCOMUSU
            Call HandleUserCommerceInit
        
        Case ServerPacketID.UserCommerceEnd         'FINCOMUSUOK
            Call HandleUserCommerceEnd
        
        Case ServerPacketID.UserOfferConfirm
            Call HandleUserOfferConfirm
                
        Case ServerPacketID.CharSwing
            Call HandleCharSwing
        
        Case ServerPacketID.NPCKillUser             '6
            Call HandleNPCKillUser
        
        Case ServerPacketID.BlockedWithShield
            Call HandleBlockedWithShield
        
        Case ServerPacketID.BlockedWithShieldOther  '8
            Call HandleBlockedWithShieldOther
        
        Case ServerPacketID.UserSwing               'U1
            Call HandleUserSwing
        
        Case ServerPacketID.ResuscitationSafeOff
            Call HandleResuscitationSafeOff
        
        Case ServerPacketID.ResuscitationSafeOn
            Call HandleResuscitationSafeOn
        
        Case ServerPacketID.UpdateSta               'ASS
            Call HandleUpdateSta
        
        Case ServerPacketID.UpdateMana              'ASM
            Call HandleUpdateMana
        
        Case ServerPacketID.UpdateHP                'ASH
            Call HandleUpdateHP
        
        Case ServerPacketID.UpdateGold              'ASG
            Call HandleUpdateGold
        
        Case ServerPacketID.UpdateExp               'ASE
            Call HandleUpdateExp
        
        Case ServerPacketID.ChangeMap               'CM
            Call HandleChangeMap
        
        Case ServerPacketID.PosUpdate               'PU
            Call HandlePosUpdate
            
        Case ServerPacketID.Damage
            Call HandleDamage
        
        Case ServerPacketID.UserDamaged
            Call HandleUserDamaged
                
        Case ServerPacketID.ChatOverHead
            Call HandleChatOverHead
        
        Case ServerPacketID.DeleteChatOverHead
            Call HandleDeleteChatOverHead

        Case ServerPacketID.ConsoleMsg
            Call HandleConsoleMessage
        
        Case ServerPacketID.ChatNormal
            Call HandleChatNormal
            
        'Case ServerPacketID.ChatGM
            'Call HandleChatGM
                 
        Case ServerPacketID.ChatGuild
            Call HandleChatGuild
            
        Case ServerPacketID.ChatCompa
            Call HandleChatCompa
            
        Case ServerPacketID.ChatPrivate
            Call HandleChatPrivate
                        
        Case ServerPacketID.ChatGlobal
            Call HandleChatGlobal
                        
        Case ServerPacketID.ShowMessageBox
            Call HandleShowMessageBox
                
        Case ServerPacketID.CharCreate
            Call HandleCharCreate
            
        Case ServerPacketID.NpcCharCreate
            Call HandleNpcCharCreate
        
        Case ServerPacketID.CharRemove
            Call HandleCharRemove
        
        Case ServerPacketID.CharChangeNick
            Call HandleCharChangeNick
            
        Case ServerPacketID.CharMove
            Call HandleCharMove
            
        Case ServerPacketID.ForceCharMove
            Call HandleForceCharMove
        
        Case ServerPacketID.CharChange
            Call HandleCharChange
        
        Case ServerPacketID.ChangeCharHeading
            Call HandleChangeCharHeading
        
        Case ServerPacketID.ObjCreate
            Call HandleObjCreate
        
        Case ServerPacketID.ObjectDelete
            Call HandleObjectDelete
        
        Case ServerPacketID.BlockPosition
            Call HandleBlockPosition
        
        Case ServerPacketID.PlayMP3
            Call HandlePlayMP3
        
        Case ServerPacketID.SoundFX
            Call HandleSoundFX
        
        Case ServerPacketID.GuildList
            Call HandleGuildList
        
        Case ServerPacketID.AreaChanged
            Call HandleAreaChanged
        
        Case ServerPacketID.PauseToggle
            Call HandlePauseToggle
        
        Case ServerPacketID.RainToggle
            Call HandleRainToggle
            
        Case ServerPacketID.Weather
            Call HandleWeather
        
        Case ServerPacketID.CreateFX
            Call HandleCreateFX
        
        Case ServerPacketID.CreateCharFX
            Call HandleCreateCharFX
        
        Case ServerPacketID.UpdateUserStats
            Call HandleUpdateUserStats

        Case ServerPacketID.SlotMenosUno
            Call HandleSlotMenosUno
        
        Case ServerPacketID.Inventory
            Call HandleInventory
        
        Case ServerPacketID.InventorySlot
            Call HandleInventorySlot
            
        Case ServerPacketID.BeltInv
            Call HandleBeltInv
        
        Case ServerPacketID.BeltSlot
            Call HandleBeltSlot
            
        Case ServerPacketID.Bank
            Call HandleBank
        
        Case ServerPacketID.BankSlot
            Call HandleBankSlot
        
        Case ServerPacketID.NpcInventory
            Call HandleNpcInventory
        
        Case ServerPacketID.NpcInventorySlot
            Call HandleNpcInventorySlot
        
        Case ServerPacketID.Spells
            Call HandleSpells
        
        Case ServerPacketID.SpellSlot
            Call HandleSpellSlot
        
        Case ServerPacketID.Compas
            Call HandleCompas

        Case ServerPacketID.AddCompa
            Call HandleAddCompa
        
        Case ServerPacketID.QuitarCompa
            Call HandleQuitarCompa
            
        Case ServerPacketID.CompaConnected
            Call HandleCompaConnected
                        
        Case ServerPacketID.CompaDisconnected
            Call HandleCompaDisconnected
                
        Case ServerPacketID.Attributes
            Call HandleAttributes
            
        Case ServerPacketID.UserPlatforms
            Call HandleUserPlatforms
        
        Case ServerPacketID.BlacksmithWeapons       'LAH
            Call HandleBlacksmithWeapons
        
        Case ServerPacketID.BlacksmithArmors        'LAR
            Call HandleBlacksmithArmors

        Case ServerPacketID.CarpenterObjects        'OBR
            Call HandleCarpenterObjects
        
        Case ServerPacketID.RestOK                  'DOK
            Call HandleRestOK
        
        Case ServerPacketID.ErrorMsg                'ERR
            Call HandleErrorMessage
        
        Case ServerPacketID.Blind
            Call HandleBlind
        
        Case ServerPacketID.Dumb
            Call HandleDumb
        
        Case ServerPacketID.ShowSignal
            Call HandleShowSignal

        Case ServerPacketID.UpdateHungerAndThirst
            Call HandleUpdateHungerAndThirst
        
        Case ServerPacketID.MiniStats
            Call HandleMiniStats
            
        Case ServerPacketID.SkillUp
            Call HandleSkillUp
        
        Case ServerPacketID.LevelUp
            Call HandleLevelUp
        
        Case ServerPacketID.SetInvisible
            Call HandleSetInvisible
            
        Case ServerPacketID.SetParalized
            Call HandleSetParalized
                    
        Case ServerPacketID.BlindNoMore
            Call HandleBlindNoMore
        
        Case ServerPacketID.DumbNoMore
            Call HandleDumbNoMore
        
        Case ServerPacketID.Skills
            Call HandleSkills
            
        Case ServerPacketID.FreeSkillPts
            Call HandleFreeSkillPts
        
        Case ServerPacketID.TrainerCreatureList
            Call HandleTrainerCreatureList
        
        Case ServerPacketID.GuildNews
            Call HandleGuildNews

        Case ServerPacketID.OfferDetails
            Call HandleOfferDetails
        
        Case ServerPacketID.AlianceProposalsList
            Call HandleAlianceProposalsList
        
        Case ServerPacketID.PeaceProposalsList
            Call HandlePeaceProposalsList
        
        Case ServerPacketID.CharInfo
            Call HandleCharInfo
        
        Case ServerPacketID.GuildLeaderInfo
            Call HandleGuildLeaderInfo
        
        Case ServerPacketID.GuildDetails
            Call HandleGuildDetails
        
        Case ServerPacketID.ShowGuildFundationForm
            Call HandleShowGuildFundationForm
        
        Case ServerPacketID.ShowUserRequest
            Call HandleShowUserRequest
                
        Case ServerPacketID.ChangeUserTradeSlot
            Call HandleChangeUserTradeSlot
        
        Case ServerPacketID.Pong
            Call HandlePong
        
        Case ServerPacketID.UpdateTagAndStatus
            Call HandleUpdateTagAndStatus
            
        Case ServerPacketID.Population
            Call HandlePopulation
        
        Case ServerPacketID.AnimAttack
            Call HandleAnimAttack
            
        Case ServerPacketID.CharMeditate
            Call HandleCharMeditate
        
        Case ServerPacketID.GuildMemberInfo
            Call HandleGuildMemberInfo

        Case ServerPacketID.ShowPartyForm
            Call HandleShowPartyForm
        
        Case ServerPacketID.UpdateStrenghtAndDexterity
            Call HandleUpdateStrenghtAndDexterity
            
        Case ServerPacketID.UpdateStrenght
            Call HandleUpdateStrenght
            
        Case ServerPacketID.UpdateDexterity
            Call HandleUpdateDexterity
        
        Case ServerPacketID.StopWorking
            Call HandleStopWorking
            
        Case ServerPacketID.CancelOfferItem
            Call HandleCancelOfferItem

        'GM messages
        Case ServerPacketID.SpawnList               'SPL
            Call HandleSpawnList
        
        Case ServerPacketID.ShowSOSForm             'RSOS and MSOS
            Call HandleShowSOSForm
        
        Case ServerPacketID.ShowMOTDEditionForm     'ZMOTD
            Call HandleShowMOTDEditionForm
        
        Case ServerPacketID.ShowGMPanelForm         'ABPANEL
            Call HandleShowGMPanelForm
        
        Case ServerPacketID.UserNameList            'LISTUSU
            Call HandleUserNameList
                
        Case Else
            'Error : Abort!
            Exit Sub
    End Select
    
    'Done with this packet, move on to next one
    If incomingData.length > 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then
        Err.Clear
        Call HandleIncomingData
    End If

End Sub

Private Sub HandleLogged()
    Call incomingData.ReadByte
    Call SetConnected
End Sub

Private Sub HandleRandomName()
    Call incomingData.ReadByte
    
    frmCrearPersonaje.NameTxt.Text = incomingData.ReadASCIIString
    frmCrearPersonaje.NameTxt.SetFocus
End Sub

Private Sub HandleNavigateToggle()
    Call incomingData.ReadByte
    UserNavegando = Not UserNavegando
End Sub

Private Sub HandleUserCommerceInit()

End Sub

Private Sub HandleUserCommerceEnd()

End Sub

Private Sub HandleUserOfferConfirm()

End Sub

Private Sub HandleCharSwing()
    Call incomingData.ReadByte
    
    AttackerCharIndex = incomingData.ReadInteger
    
    CharDamage = "Falla"
    
    CharDamageType = 1
    
    If Meditando Then
        Meditando = False
        Charlist(UserCharIndex).FxIndex = 0
    ElseIf Descansando Then
        Descansando = False
    End If
End Sub

Private Sub HandleNPCKillUser()
    Call incomingData.ReadByte
    
    Call ShowConsoleMsg(MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)
End Sub

Private Sub HandleBlockedWithShield()
    
    Dim CharBlocked As Integer
    
    Call incomingData.ReadByte
    
    CharBlocked = incomingData.ReadInteger
    
    Charlist(CharBlocked).Escudo.ShieldWalk(Charlist(CharBlocked).Heading).Started = 1

    If CharBlocked = UserCharIndex Then
        Call Audio.Play(SND_ESCUDO)
    Else
        Call Audio.Play(SND_ESCUDO, Charlist(CharBlocked).Pos.x, Charlist(CharBlocked).Pos.y)
    End If
End Sub

Private Sub HandleBlockedWithShieldOther()
    Call incomingData.ReadByte
    
    Call ShowConsoleMsg(MENSAJE_USUARIO_rechazó_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
End Sub

Private Sub HandleUserSwing()
    Call incomingData.ReadByte
    
    Call InitDamage("Fallás")
    
    DamageType = 1
    
    If RandomNumber(1, 2) = 1 Then
        Call Audio.Play(SND_SWING)
    Else
        Call Audio.Play(SND_SWING2)
    End If
End Sub

Private Sub HandleResuscitationSafeOff()
    Call incomingData.ReadByte
    
    Call ShowConsoleMsg(MENSAJE_SEGURO_RESU_OFF, 255, 0, 0, True, False, False)
End Sub

Private Sub HandleResuscitationSafeOn()
    Call incomingData.ReadByte
    
    Call ShowConsoleMsg(MENSAJE_SEGURO_RESU_ON, 0, 255, 0, True, False, False)
End Sub

Private Sub HandleUpdateSta()

    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinSTA = incomingData.ReadInteger
    
    Dim i As Byte
    
    For i = 1 To MaxSpellSlots
        If Spell(i).Grh > 0 Then
            If Spell(i).PuedeLanzar = False Then
                If Spell(i).ManaRequerido <= UserMinMan And Spell(i).StaRequerido <= UserMinSTA Then
                    Spell(i).PuedeLanzar = True
                    Call Hechizos.DrawSpellSlot(i)
                End If
                
            ElseIf Spell(i).ManaRequerido > UserMinMan Or Spell(i).StaRequerido > UserMinSTA Then
                Spell(i).PuedeLanzar = False
                Call Hechizos.DrawSpellSlot(i)
            End If
        End If
    Next i
End Sub

Private Sub HandleUpdateMana()
    
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinMan = incomingData.ReadInteger
    
    If Meditando Then
        If UserMinMan = UserMaxMan Then
            Meditando = False
            Charlist(UserCharIndex).FxIndex = 0
        End If
    End If
    
    Dim i As Byte
    
    For i = 1 To MaxSpellSlots
        If Spell(i).Grh > 0 Then
            If Spell(i).PuedeLanzar = False Then
                If Spell(i).ManaRequerido <= UserMinMan Then
                    Spell(i).PuedeLanzar = True
                    Call Hechizos.DrawSpellSlot(i)
                End If
                
            ElseIf Spell(i).ManaRequerido > UserMinMan Or Spell(i).StaRequerido > UserMinSTA Then
                Spell(i).PuedeLanzar = False
                Call Hechizos.DrawSpellSlot(i)
            End If
        End If
    Next i
End Sub

Private Sub HandleUpdateHP()

    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinHP = incomingData.ReadInteger
    
    If UserMuerto Then
        If UserMinHP > 0 Then
            UserMuerto = False
        End If
    Else
        If UserMinHP < 1 Then
            Call Morir
        End If
    End If
    
    If UserMinHP < UserMaxHP * 0.2 Then
        If Not Audio.PlayingSound Then
            Call Audio.Play(SND_HEARTBEAT)
        End If
    End If
    
End Sub

Private Sub HandleUpdateGold()

On Error GoTo Error
        
    Dim Gld As Long
    
    Call incomingData.ReadByte
    
    Gld = UserGld
    
    'Get data and update form
    UserGld = incomingData.ReadLong
    
    Call InitGld(UserGld - Gld)
    
    frmMain.GldLbl.Caption = PonerPuntos(UserGld)
    
Error:
End Sub

Private Sub HandleUpdateExp()
    Dim Exp As Long
    
    Call incomingData.ReadByte
    
    'Get data and update form
    Exp = UserExp
    UserExp = incomingData.ReadLong
    
    Call InitExp(CStr(UserExp - Exp))
    
    frmMain.ExpLbl.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
End Sub

Private Sub HandleUpdateStrenghtAndDexterity()
    Call incomingData.ReadByte
    
    'Get data and update form
    UserFuerza = incomingData.ReadByte
    UserAgilidad = incomingData.ReadByte
    frmMain.lblStrg.Caption = UserFuerza
    frmMain.lblDext.Caption = UserAgilidad
    frmMain.lblStrg.ForeColor = getStrenghtColor
    frmMain.lblDext.ForeColor = getDexterityColor
End Sub

Private Sub HandleUpdateStrenght()
    Call incomingData.ReadByte
    
    'Get data and update form
    UserFuerza = incomingData.ReadByte
    frmMain.lblStrg.Caption = UserFuerza
    frmMain.lblStrg.ForeColor = getStrenghtColor
End Sub

Private Sub HandleUpdateDexterity()
    Call incomingData.ReadByte
    
    'Get data and update form
    UserAgilidad = incomingData.ReadByte
    frmMain.lblDext.Caption = UserAgilidad
    frmMain.lblDext.ForeColor = getDexterityColor
End Sub

Private Sub HandlePopulation()
    Call incomingData.ReadByte
    
    'Get data and update form
    frmMain.LblPoblacion.Caption = incomingData.ReadInteger
End Sub

Private Sub HandleAnimAttack()
    Dim CharAttacking As Integer
    
    Call incomingData.ReadByte
    
    CharAttacking = incomingData.ReadInteger
        
    'Get data and update form
    Charlist(CharAttacking).Arma.WeaponWalk(Charlist(CharAttacking).Heading).Started = 1
    Charlist(CharAttacking).Escudo.ShieldWalk(Charlist(CharAttacking).Heading).Started = 1
    
End Sub

Private Sub HandleCharMeditate()
    Dim CharMeditating As Integer
    
    Call incomingData.ReadByte
    
    CharMeditating = incomingData.ReadInteger
        
    With Charlist(CharMeditating)
        
        Select Case .Lvl
            'Show proper FX according to level
            Case Is < 15
                .FxIndex = FX_MEDITARCHICO
            Case Is < 25
                .FxIndex = FX_MEDITARMEDIANO
            Case Is < 35
                .FxIndex = FX_MEDITARGRANDE
            Case Is < 40
                .FxIndex = FX_MEDITARXGRANDE
            Case Else
                .FxIndex = FX_MEDITARXXGRANDE
        End Select
        
        .fX.Loops = -1
            
        Call InitGrh(.fX, FxData(.FxIndex).Animacion)
            
        If PortalBufferIndex = 0 Then
            PortalBufferIndex = Audio.Play(SND_PORTAL, .Pos.x, .Pos.y, LoopStyle.Enabled)
        End If
            
    End With
    
End Sub

Private Sub HandleChangeMap()
    
On Error Resume Next

    Call incomingData.ReadByte
    
    UserMap = incomingData.ReadInteger

    MapInfo(UserMap).Name = vbNullString
    MapInfo(UserMap).Music = 0
    
    If bLluvia(UserMap) = 0 Then
        If bRain Then
            Call Audio.StopWave(RainBufferIndex)
            RainBufferIndex = 0
            frmMain.IsPlaying = PlayLoop.plNone
        End If
    End If
    
    MapInfo(UserMap).Name = GetVar(MapPath & "Maps.dat", "Mapa" & UserMap, "Name")
    MapInfo(UserMap).Name = Replace(MapInfo(UserMap).Name, ".", vbNullString)

    If LenB(MapInfo(UserMap).Name) > 0 Then
        frmMain.lblMapName.Caption = Replace(MapInfo(UserMap).Name, ".", vbNullString)
    End If

    MapInfo(UserMap).Zone = GetVar(MapPath & "Maps.dat", "Mapa" & UserMap, "Zone")
    MapInfo(UserMap).Music = GetVar(MapPath & "Maps.dat", "Mapa" & UserMap, "Music")
    MapInfo(UserMap).Top = GetVar(MapPath & "Maps.dat", "Mapa" & UserMap, "Top")
    MapInfo(UserMap).Left = GetVar(MapPath & "Maps.dat", "Mapa" & UserMap, "Left")
    
    If MusicActivated Then
        If MapInfo(UserMap).Music > 134 Then
            Call Audio.MusicMP3Play(MapInfo(UserMap).Music)
        Else
            Call Audio.PlayMIDI(MapInfo(UserMap).Music)
        End If
    End If

    'Call GenerarMiniMapa
        
    'CargarMapa UserMap

End Sub

Private Sub HandlePosUpdate()
    
    Dim x As Integer
    Dim y As Integer

    Call incomingData.ReadByte
    
    x = incomingData.ReadInteger
    y = incomingData.ReadInteger
    
    'Set new pos
    If UserPos.x <> x Or UserPos.y <> y Then
        'Set new pos
        UserPos.x = x
        UserPos.y = y
        
        'Remove char from old position
        If MapData(UserPos.x, UserPos.y).CharIndex = UserCharIndex Then
            MapData(UserPos.x, UserPos.y).CharIndex = 0
        End If
        
        'Set char
        MapData(UserPos.x, UserPos.y).CharIndex = UserCharIndex
        Charlist(UserCharIndex).Pos = UserPos
        
        'Are we under a roof?
        bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
            MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
            MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
            
        Call DibujarMiniMapa
    End If
End Sub

Private Sub HandleDamage()
    
    Dim Damg As Integer
    
    Call incomingData.ReadByte
    
    AttackedCharIndex = incomingData.ReadInteger
                    
    CharMinHP = incomingData.ReadByte
    
    DamageType = incomingData.ReadByte
    
    Damg = incomingData.ReadInteger
    
    Call InitDamage(CStr(Damg))
    
    If AttackedCharIndex = UserCharIndex Then
        'Update UserMinHP from CharMinHP
        UserMinHP = UserMinHP + Damg
        
        If UserMinHP > UserMaxHP Then
            UserMinHP = UserMaxHP
        End If
    End If
    
End Sub

Private Sub HandleUserDamaged()

    Call incomingData.ReadByte
            
    AttackerCharIndex = incomingData.ReadInteger
    
    CharDamage = CStr(incomingData.ReadInteger)
    
    CharDamageType = incomingData.ReadByte
    
    If Meditando Then
        Meditando = False
        Charlist(UserCharIndex).FxIndex = 0
    ElseIf Descansando Then
        Descansando = False
    End If
    
    If Charlist(UserCharIndex).Priv > 1 Then
        Exit Sub
    End If
    
    Select Case CharDamageType
        Case 2
            'Update UserMinHP from CharDamage
            UserMinHP = UserMinHP + CharDamage
            
            If UserMinHP > UserMaxHP Then
                UserMinHP = UserMaxHP
            End If
            
        Case Else
            'Update UserMinHP from CharDamage
            UserMinHP = UserMinHP - CharDamage
            
            If UserMinHP < 1 Then
                Call Morir
            End If
            
    End Select
    
    Call InitGrh(Charlist(UserCharIndex).fX, FxData(14).Animacion)

    Charlist(UserCharIndex).fX.Loops = 0
End Sub

Private Sub HandleChatOverHead()
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim Chat As String
    Dim CharIndex As Integer
    Dim r As Byte
    Dim g As Byte
    Dim B As Byte
    
    Chat = Buffer.ReadASCIIString
    CharIndex = Buffer.ReadInteger
    
    r = Buffer.ReadByte
    g = Buffer.ReadByte
    B = Buffer.ReadByte
    
    'Only add the Chat if the Char exists (a CharRemove may have been sent to the PC / NPC area before the buffer was flushed)
    'If Charlist(CharIndex).Active Then
        Call Dialogos.CreateDialog(Chat, CharIndex, RGB(r, g, B))
    'End If
        
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleDeleteChatOverHead()
    Call incomingData.ReadByte
    Call Dialogos.CreateDialog(vbNullString, incomingData.ReadInteger, vbWhite)
End Sub
Private Sub HandleConsoleMessage()

On Error Resume Next
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim Chat As String
    Dim FontIndex As Integer
    Dim str As String
    Dim r As Byte
    Dim g As Byte
    Dim B As Byte
    
    Chat = Buffer.ReadASCIIString
    FontIndex = Buffer.ReadByte
    
    If InStr(1, Chat, "~") Then
        str = ReadField(2, Chat, 126)
            If Val(str) > 255 Then
                r = 255
            Else
                r = Val(str)
            End If
            
            str = ReadField(3, Chat, 126)
            If Val(str) > 255 Then
                g = 255
            Else
                g = Val(str)
            End If
            
            str = ReadField(4, Chat, 126)
            If Val(str) > 255 Then
                B = 255
            Else
                B = Val(str)
            End If
            
        Call ShowConsoleMsg(Left$(Chat, InStr(1, Chat, "~") - 1), r, g, B, Val(ReadField(5, Chat, 126)) > 0, Val(ReadField(6, Chat, 126)) > 0)
    Else
        If Left$(LastParsedString, 1) = ":" Then
            If LenB(RTrim$(LastParsedString)) = LenB(Left$(Chat, Len(Chat) - 9)) + 2 Then
                If RTrim$(LastParsedString) = ":" & Left$(Chat, Len(Chat) - 9) Then
                    LastParsedString = vbNullString
                End If
            ElseIf LenB(Chat) = LenB(Right$(RTrim$(LastParsedString), Len(LastParsedString) - 2)) + 50 Then
                If Chat = "No existe nadie llamado " & Right$(RTrim$(LastParsedString), Len(LastParsedString) - 2) & "." Then
                    LastParsedString = vbNullString
                End If
            End If
        End If
        
        With FontTypes(FontIndex)
            Call ShowConsoleMsg(Chat, .Red, .Green, .Blue, .Bold, .Italic)
        End With
    End If
    
    'Call checkText(Chat)
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleChatNormal()

On Error Resume Next

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim Chat As String
    Dim CharIndex As Integer
            
    Chat = Buffer.ReadASCIIString
    CharIndex = Buffer.ReadInteger
    
    If Charlist(CharIndex).Priv < 2 Then
        With FontTypes(FontTypeNames.FONTTYPE_TALK)
            Call ShowConsoleMsg(Charlist(CharIndex).Nombre & ": ", .Red, .Green, .Blue, .Bold, .Italic, True, True)
            Call Dialogos.CreateDialog(Chat, CharIndex, RGB(.Red, .Green, .Blue))
        End With
             
    Else
        With FontTypes(FontTypeNames.FONTTYPE_GM)
            Call ShowConsoleMsg(Charlist(CharIndex).Nombre & ": ", .Red, .Green, .Blue, .Bold, .Italic, True, True)
            Call Dialogos.CreateDialog(Chat, CharIndex, RGB(.Red, .Green, .Blue))
        End With
    End If
    
    If Right$(Chat, 1) = "!" Then
        If Charlist(CharIndex).Priv < 2 Then
            With FontTypes(FontTypeNames.FONTTYPE_YELL)
                Call ShowConsoleMsg(Chat, .Red, .Green, .Blue, .Bold, .Italic, , True)
                Call Dialogos.CreateDialog(Chat, CharIndex, RGB(.Red, .Green, .Blue))
            End With
                 
        Else
            With FontTypes(FontTypeNames.FONTTYPE_YELLGM)
                Call ShowConsoleMsg(Chat, .Red, .Green, .Blue, .Bold, .Italic, , True)
                Call Dialogos.CreateDialog(Chat, CharIndex, RGB(.Red, .Green, .Blue))
            End With
        End If
        
    Else
        If Charlist(CharIndex).Priv < 2 Then
            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                Call ShowConsoleMsg(Chat, .Red, .Green, .Blue, .Bold, .Italic, , True)
            End With
                 
        Else
            With FontTypes(FontTypeNames.FONTTYPE_TALKGM)
                Call ShowConsoleMsg(Chat, .Red, .Green, .Blue, .Bold, .Italic, , True)
            End With
        End If
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

End Sub

'PRIVATE SUB HandleChatGM()

'On Error Resume Next

'    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
'    Dim Buffer As New clsByteQueue
'    Call Buffer.CopyBuffer(incomingData)
    
'    Call Buffer.ReadByte
    
'    Dim Chat As String
'    Dim CharIndex As Integer
            
'    Chat = Buffer.ReadASCIIString
'    CharIndex = Buffer.ReadInteger
    
'    With FontTypes(FontTypeNames.FONTTYPE_TALK)
'        Call ShowConsoleMsg(Charlist(CharIndex).Nombre & ": " & Chat, .Red, .Green, .Blue, .Bold, .Italic, , True)
'        Call Dialogos.CreateDialog(Chat, CharIndex, RGB(.Red, .Green, .Blue))
'    End With
    
    'If we got here then packet is complete, copy data back to original queue
'    Call incomingData.CopyBuffer(Buffer)

'END SUB

Private Sub HandleChatGuild()

On Error Resume Next

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim Chat As String
    Dim Nombre As String
    Dim i As Integer
            
    Chat = Buffer.ReadASCIIString
    Nombre = Buffer.ReadASCIIString
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        Call ShowConsoleMsg(Nombre & ": ", .Red, .Green, .Blue, .Bold, .Italic, True, True)
        
        For i = 1 To LastChar
            If Charlist(i).EsUser Then
                If LenB(Charlist(i).Nombre) = LenB(Nombre) Then
                    If Charlist(i).Nombre = Nombre Then
                        Call Dialogos.CreateDialog(Chat, i, RGB(.Red, .Green, .Blue))
                        Exit For
                    End If
                End If
            End If
        Next i
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        Call ShowConsoleMsg(Chat, .Red, .Green, .Blue, .Bold, .Italic, , True)
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

End Sub

Private Sub HandleChatCompa()

On Error Resume Next

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim Chat As String
    Dim Nro As Byte
    Dim i As Integer
    
    Chat = Buffer.ReadASCIIString
    Nro = Buffer.ReadByte
    
    With FontTypes(FontTypeNames.FONTTYPE_COMPAMESSAGE)
        Call ShowConsoleMsg(Compa(Nro).Nombre & ": ", .Red, .Green, .Blue, .Bold, .Italic, True, True)
        
        For i = 1 To LastChar
            If Charlist(i).EsUser Then
                If LenB(Charlist(i).Nombre) = LenB(Compa(Nro).Nombre) Then
                    If Charlist(i).Nombre = Compa(Nro).Nombre Then
                        Call Dialogos.CreateDialog(Chat, i, RGB(.Red, .Green, .Blue))
                        Exit For
                    End If
                End If
            End If
        Next i
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        Call ShowConsoleMsg(Chat, .Red, .Green, .Blue, .Bold, .Italic, , True)
    End With

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

End Sub

Private Sub HandleChatPrivate()

On Error Resume Next

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim Chat As String
    Dim Nombre As String
    Dim i As Integer
    
    Chat = Buffer.ReadASCIIString
    Nombre = Buffer.ReadASCIIString

    With FontTypes(FontTypeNames.FONTTYPE_PRIVATEMESSAGE)
    
        Call ShowConsoleMsg(Nombre & ": ", .Red, .Green, .Blue, .Bold, .Italic, True, True)
        
        For i = 1 To LastChar
            If Charlist(i).EsUser Then
                If LenB(Charlist(i).Nombre) = LenB(Nombre) Then
                    If Charlist(i).Nombre = Nombre Then
                        Call Dialogos.CreateDialog(Chat, i, RGB(.Red, .Green, .Blue))
                        Exit For
                    End If
                End If
            End If
        Next i
        
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        Call ShowConsoleMsg(Chat, .Red, .Green, .Blue, .Bold, .Italic, , True)
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

End Sub

Private Sub HandleChatGlobal()

On Error Resume Next

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim Chat As String
    Dim Nombre As String
    
    Chat = Buffer.ReadASCIIString
    Nombre = Buffer.ReadASCIIString

    With FontTypes(FontTypeNames.FONTTYPE_PUBLICMESSAGE)
        Call ShowConsoleMsg(Nombre & ": ", .Red, .Green, .Blue, .Bold, .Italic, True, True)
    End With

    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        Call ShowConsoleMsg(Chat, 170, 170, 170, .Bold, True, , True)
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

End Sub

Private Sub HandleCommerceChat()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim Chat As String
    Dim FontIndex As Integer
    Dim str As String
    Dim r As Byte
    Dim g As Byte
    Dim B As Byte
    
    Chat = Buffer.ReadASCIIString
    FontIndex = Buffer.ReadByte
    
    If InStr(1, Chat, "~") Then
        str = ReadField(2, Chat, 126)
            If Val(str) > 255 Then
                r = 255
            Else
                r = Val(str)
            End If
            
            str = ReadField(3, Chat, 126)
            If Val(str) > 255 Then
                g = 255
            Else
                g = Val(str)
            End If
            
            str = ReadField(4, Chat, 126)
            If Val(str) > 255 Then
                B = 255
            Else
                B = Val(str)
            End If
            
        Call AddtoCommerceRecTxt(Left$(Chat, InStr(1, Chat, "~") - 1), r, g, B, Val(ReadField(5, Chat, 126)) > 0, Val(ReadField(6, Chat, 126)) > 0)
    Else
        With FontTypes(FontIndex)
            Call AddtoCommerceRecTxt(Chat, .Red, .Green, .Blue, .Bold, .Italic)
        End With
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleShowMessageBox()
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    frmMensaje.msg.Caption = Buffer.ReadASCIIString
    frmMensaje.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleCharCreate()
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As eHeading
    Dim x As Integer
    Dim y As Integer
    Dim Weapon As Integer
    Dim Shield As Integer
    Dim Helmet As Integer
    Dim Privs As Integer
    
    CharIndex = Buffer.ReadInteger
    Body = Buffer.ReadInteger
    Head = Buffer.ReadInteger
    Heading = Buffer.ReadByte
    x = Buffer.ReadInteger
    y = Buffer.ReadInteger
    Weapon = Buffer.ReadByte
    Shield = Buffer.ReadByte
    Helmet = Buffer.ReadByte
    
    With Charlist(CharIndex)
        .EsUser = True
        .Nombre = Buffer.ReadASCIIString
        .Guilda = Buffer.ReadASCIIString
        Privs = Buffer.ReadByte
        .Lvl = Buffer.ReadByte
        
        If Privs > 0 Then
            'If the player is a RM, ignore other flags
            If Privs And PlayerType.RoleMaster Then
                Privs = PlayerType.RoleMaster
            End If
            
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .Priv = Log(Privs) / Log(2)
        Else
            .Priv = 0
        End If
        
        .CompaIndex = EsCompaniero(.Nombre)
        
    End With
    
    Call MakeChar(CharIndex, Body, Head, Heading, x, y, Weapon, Shield, Helmet)
        
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleNpcCharCreate()
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As eHeading
    Dim x As Integer
    Dim y As Integer
    
    CharIndex = Buffer.ReadInteger
    
    Body = Buffer.ReadInteger

    If Body > 4000 Then
        Head = Buffer.ReadInteger
        Body = Body - 4000
    End If
    
    Heading = Buffer.ReadByte
    
    x = Buffer.ReadInteger
    y = Buffer.ReadInteger
    
    Charlist(CharIndex).Nombre = Buffer.ReadASCIIString

    If Body > 2000 Then
        Charlist(CharIndex).Lvl = Buffer.ReadByte
        Body = Body - 2000
    Else
        Charlist(CharIndex).Lvl = 1
    End If
         
    If Body > 1000 Then
        Charlist(CharIndex).MascoIndex = Buffer.ReadByte
        Body = Body - 1000
    End If
    
    Call MakeChar(CharIndex, Body, Head, Heading, x, y)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleCharChangeNick()
    Call incomingData.ReadByte
    Dim CharIndex As Integer
    CharIndex = incomingData.ReadInteger
    Charlist(CharIndex).Nombre = incomingData.ReadASCIIString
End Sub

Private Sub HandleCharRemove()
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger

    Call EraseChar(CharIndex)
    
    Call DibujarMiniMapa
End Sub

Private Sub HandleCharMove()
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim x As Integer
    Dim y As Integer
    
    CharIndex = incomingData.ReadInteger
    x = incomingData.ReadInteger
    y = incomingData.ReadInteger
    
    If CharIndex < 1 Then
        Exit Sub
    End If
    
    If Charlist(CharIndex).FxIndex = FX_MEDITARCHICO Or _
    Charlist(CharIndex).FxIndex = FX_MEDITARMEDIANO Or _
    Charlist(CharIndex).FxIndex = FX_MEDITARGRANDE Or _
    Charlist(CharIndex).FxIndex = FX_MEDITARXGRANDE Or _
    Charlist(CharIndex).FxIndex = FX_MEDITARXXGRANDE Then
        Charlist(CharIndex).FxIndex = 0
    End If
    
    Call DoPasosFx(CharIndex)
    
    Call MoveCharbyPos(CharIndex, x, y)
End Sub

Private Sub HandleForceCharMove()
    Call incomingData.ReadByte
    
    Dim Direccion As Byte
    
    Direccion = incomingData.ReadByte

    Call MoveCharbyHead(Direccion)
End Sub

Private Sub HandleCharChange()
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim TempInt As Integer
    Dim headIndex As Integer
    
    CharIndex = incomingData.ReadInteger

    With Charlist(CharIndex)
        TempInt = incomingData.ReadInteger
        
        If TempInt < LBound(BodyData()) Or TempInt > UBound(BodyData()) Then
            .Body = BodyData(0)
            .iBody = 0
        Else
            .Body = BodyData(TempInt)
            .iBody = TempInt
        End If
        
        headIndex = incomingData.ReadInteger
        
        If headIndex < LBound(HeadData()) Or headIndex > UBound(HeadData()) Then
            .Head = HeadData(0)
            .iHead = 0
        Else
            .Head = HeadData(headIndex)
            .iHead = headIndex
        End If
        
        .Heading = incomingData.ReadByte
        
        TempInt = incomingData.ReadByte
        
        If TempInt > 0 Then
            .Arma = WeaponAnimData(TempInt)
        End If
        
        TempInt = incomingData.ReadByte
        
        If TempInt > 0 Then
            .Escudo = ShieldAnimData(TempInt)
        End If
        
        TempInt = incomingData.ReadByte
        
        If TempInt > 0 Then
            .Casco = CascoAnimData(TempInt)
        End If
        
    End With
End Sub

Private Sub HandleChangeCharHeading()

On Error Resume Next

    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger

    Charlist(CharIndex).Heading = incomingData.ReadByte
End Sub

Private Sub HandleObjCreate()
    
    Call incomingData.ReadByte
    
    Dim x As Integer
    Dim y As Integer
    
    x = incomingData.ReadInteger
    y = incomingData.ReadInteger
    
    With MapData(x, y)
    
        .Obj.ObjType = incomingData.ReadByte
        
        If .Obj.ObjType = otGuita Then
        
            .Obj.ObjType = .Obj.ObjType - 100
            
            .Obj.Amount = incomingData.ReadLong
            
            .Obj.Name = "Monedas de oro (" & PonerPuntos(.Obj.Amount) & ")"
                    
            Select Case .Obj.Amount
                
                Case Is > 999999
                    .Obj.Grh.GrhIndex = 23957
                Case Is > 99999
                    .Obj.Grh.GrhIndex = 23956
                Case Is > 49999
                    .Obj.Grh.GrhIndex = 23955
                Case Is > 19999
                    .Obj.Grh.GrhIndex = 23954
                Case Is > 9999
                    .Obj.Grh.GrhIndex = 23953
                Case Is > 4999
                    .Obj.Grh.GrhIndex = 23952
                Case Is > 2449
                    .Obj.Grh.GrhIndex = 23951
                Case Is > 999
                    .Obj.Grh.GrhIndex = 23950
                Case Is > 499
                    .Obj.Grh.GrhIndex = 23949
                Case Is > 249
                    .Obj.Grh.GrhIndex = 23948
                Case Is > 99
                    .Obj.Grh.GrhIndex = 23947
                Case Is > 49
                    .Obj.Grh.GrhIndex = 23946
                Case Is > 24
                    .Obj.Grh.GrhIndex = 23945
                Case Is > 9
                    .Obj.Grh.GrhIndex = 23944
                Case Else
                    .Obj.Grh.GrhIndex = 23943
                
            End Select
                                
            If .Obj.Amount > 1 Then
                .Obj.Name = "Monedas de oro (" & PonerPuntos(.Obj.Amount) & ")"
            Else
                .Obj.Name = "Moneda de oro"
            End If
            
        ElseIf .Obj.ObjType = otCuerpoMuerto Then
            Dim Name As String
            
            Name = incomingData.ReadASCIIString
            
            .Obj.Name = "Cuerpo de " & Name
            
            .Obj.Grh.GrhIndex = 23958
            
            If Name = UserName Then
                If Not UserMuerto Then
                    Exit Sub
                End If
                
                'Call Morir
                If UserPos.x <> x Or UserPos.y <> y Then
                    UserPos.x = x
                    UserPos.y = y
                End If
            End If
        
        Else
            .Obj.Amount = incomingData.ReadLong
        
            .Obj.Grh.GrhIndex = incomingData.ReadInteger

            If .Obj.Amount > 0 Then
            
                .Obj.Name = incomingData.ReadASCIIString
                               
                If .Obj.Amount > 1 Then
                    .Obj.Name = .Obj.Name & " (" & PonerPuntos(.Obj.Amount) & ")"
                End If
                        
            ElseIf .Obj.ObjType = otTeleport Then
            
                .Obj.Amount = .Obj.Amount - 10000
                
                If .Obj.Amount > 0 Then
                    MapInfo(UserMap).Name = GetVar(MapPath & "Mapa" & .Obj.Amount & ".dat", "Mapa" & .Obj.Amount, "Name")
                    
                    If LenB(MapInfo(UserMap).Name) > 0 Then
                        If Right$(MapInfo(UserMap).Name, 1) = "." Then
                            MapInfo(UserMap).Name = Replace(MapInfo(UserMap).Name, ".", vbNullString)
                        End If
    
                        .Obj.Name = "Portal a " & MapInfo(UserMap).Name
                        .Obj.Amount = 0
                    Else
                        .Obj.Name = "Portal"
                        .Obj.Amount = 0
                    End If
                
                Else
                    .Obj.Name = "Portal"
                    .Obj.Amount = 0
                End If
            End If
        End If
            
        Call InitGrh(.Obj.Grh, .Obj.Grh.GrhIndex)
    End With
End Sub

Private Sub HandleObjectDelete()
    Call incomingData.ReadByte
    
    Dim x As Integer
    Dim y As Integer
    
    x = incomingData.ReadInteger
    y = incomingData.ReadInteger
    
    With MapData(x, y)
        If x = UserPos.x And y = UserPos.y Then
            If .Obj.Amount > 0 Then
                If LenB(.Obj.Name) > 0 Then
                    If .Obj.ObjType = otGuita Then
                        Call Audio.Play(SND_PICKUP_GOLD)
                    Else
                        Call Audio.Play(SND_PICKUP)
                    End If
                End If
            End If
        End If
        
        .Obj.Grh.GrhIndex = 0
        .Obj.Name = vbNullString
        .Obj.Amount = 0
    End With
End Sub

Private Sub HandleBlockPosition()
    Call incomingData.ReadByte
    
    Dim x As Integer
    Dim y As Integer
    
    x = incomingData.ReadInteger
    y = incomingData.ReadInteger
    
    MapData(x, y).Blocked = incomingData.ReadBoolean
End Sub

Private Sub HandlePlayMP3()
    Call incomingData.ReadByte
    
    Dim Number As Byte
    
    Number = incomingData.ReadByte()
    
    If Number > 134 Then
        Call Audio.MusicMP3Play(Number)
    Else
        Call Audio.PlayMIDI(Number)
    End If
End Sub

Private Sub HandleSoundFX()
    Call incomingData.ReadByte
        
    Dim wave As Byte
    Dim srcX As Integer
    Dim srcY As Integer
    
    wave = incomingData.ReadByte
    srcX = incomingData.ReadInteger
    
    If srcX > 0 Then
        srcY = incomingData.ReadInteger
    End If

    Call Audio.Play(CStr(wave) & ".wav", srcX, srcY)
End Sub

Private Sub HandleGuildList()
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    'Clear guild's list
    frmGuildAdm.GuildsList.Clear
    
    Dim guilds() As String
    guilds = Split(Buffer.ReadASCIIString, SEPARATOR)
    
    Dim i As Long
    For i = 0 To UBound(guilds())
        Call frmGuildAdm.GuildsList.AddItem(guilds(i))
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
    frmGuildAdm.Show vbModeless, frmMain
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleAreaChanged()
    Call incomingData.ReadByte
    
    Dim x As Integer
    Dim y As Integer
    
    x = incomingData.ReadInteger
    y = incomingData.ReadInteger
        
    Call CambioDeArea(x, y)
End Sub

Private Sub HandlePauseToggle()
    
    Call incomingData.ReadByte
    
    Pausa = Not Pausa
End Sub

Private Sub HandleRainToggle()
    
    Call incomingData.ReadByte
    
    If Not InMapBounds(UserPos.x, UserPos.y) Then
        Exit Sub
    End If
    
    bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
        MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
        MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
            
    If bRain Then
        If bLluvia(UserMap) Then
            'Stop playing the rain sound
            Call Audio.StopWave(RainBufferIndex)
            RainBufferIndex = 0
            If bTecho Then
                Call Audio.Play("lluviainend.wav", 0, 0, LoopStyle.Disabled)
            Else
                Call Audio.Play("lluviaoutend.wav", 0, 0, LoopStyle.Disabled)
            End If
            frmMain.IsPlaying = PlayLoop.plNone
        End If
    End If
    
    bRain = Not bRain
End Sub

Private Sub HandleWeather()
    Call incomingData.ReadByte
    
    Dim PrevTiempo As Byte
    
    PrevTiempo = Tiempo
    
    Tiempo = incomingData.ReadByte
    
    If PrevTiempo <> 0 Then
        If Tiempo = Amanecer Then
            Call Audio.Play(SND_AMANECER)
        ElseIf Tiempo = Noche Then
            Call Audio.Play(SND_NOCHE)
        End If
    End If
    
    If Not InMapBounds(UserPos.x, UserPos.y) Then
        Exit Sub
    End If

    'bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
        MapData(Usermap,UserPos.x, UserPos.y).Trigger = 2 Or _
        MapData(Usermap,UserPos.x, UserPos.y).Trigger = 4, True, False)
End Sub

Private Sub HandleCreateFX()
    Call incomingData.ReadByte
    
    Dim x As Integer
    Dim y As Integer
    Dim fX As Integer
    Dim Loops As Byte
    
    x = incomingData.ReadInteger
    y = incomingData.ReadInteger
    fX = incomingData.ReadInteger
    
    If fX > 0 Then
        Loops = incomingData.ReadByte
    End If
    
    With MapData(x, y)
        .FxIndex = fX
            
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
          
            If Loops = 255 Then
                .fX.Loops = -1
            Else
                .fX.Loops = Loops
            End If
        End If
    End With
End Sub

Private Sub HandleCreateCharFX()
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim Loops As Byte
    
    CharIndex = incomingData.ReadInteger
    
    With Charlist(CharIndex)
        .FxIndex = incomingData.ReadInteger
        
        If .FxIndex > 0 Then
            Loops = incomingData.ReadByte
    
            .fX.Loops = Loops
            
            Call InitGrh(.fX, FxData(.FxIndex).Animacion)
          
            If Loops = 255 Then
                .fX.Loops = -1
            Else
                .fX.Loops = Loops
            End If
        End If
    End With
End Sub

Private Sub HandleUpdateUserStats()

    Dim Body As Integer
    Dim Head As Integer

    Dim Weapon As Integer
    Dim Shield As Integer
    Dim Helmet As Integer
    
    Dim Privs As Integer

    Call incomingData.ReadByte
    
    UserMaxHP = incomingData.ReadInteger
    UserMinHP = incomingData.ReadInteger
    UserMaxMan = incomingData.ReadInteger
    UserMinMan = incomingData.ReadInteger
    UserMaxSTA = incomingData.ReadInteger
    UserMinSTA = incomingData.ReadInteger
    UserGld = incomingData.ReadLong
    UserLvl = incomingData.ReadByte
    UserPasarNivel = incomingData.ReadLong
    UserExp = incomingData.ReadLong
    
    If UserCharIndex > 0 Then
        Call EraseChar(UserCharIndex)
    End If
    
    UserCharIndex = incomingData.ReadInteger
    UserPos.x = incomingData.ReadInteger
    UserPos.y = incomingData.ReadInteger
    
    With Charlist(UserCharIndex)
        If .Pos.x > 0 And .Pos.y > 0 Then
            MapData(.Pos.x, .Pos.y).CharIndex = 0
        End If
        
        .Pos = UserPos
        
        Head = incomingData.ReadInteger
        Body = incomingData.ReadInteger
        .Heading = incomingData.ReadByte
        
        Weapon = incomingData.ReadByte
        Shield = incomingData.ReadByte
        Helmet = incomingData.ReadByte
        
        .Guilda = incomingData.ReadASCIIString
        
        Privs = incomingData.ReadByte
    
        .EsUser = True
        .Nombre = UserName

        .Lvl = UserLvl
        
        If Privs > 0 Then
            'If the player is a RM, ignore other flags
            If Privs And PlayerType.RoleMaster Then
                Privs = PlayerType.RoleMaster
            End If
            
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .Priv = Log(Privs) / Log(2)
        Else
            .Priv = 0
        End If
           
        Call MakeChar(UserCharIndex, Body, Head, .Heading, .Pos.x, .Pos.y, Weapon, Shield, Helmet)
        
    End With

    If UserMinHP < 1 Then
        Call Morir
    Else
        UserMuerto = False
    End If

    frmMain.GldLbl.Caption = PonerPuntos(UserGld)
    frmMain.LvlLbl.Caption = UserLvl
    
    If UserPasarNivel > 0 Then
        frmMain.ExpLbl.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
    End If

    'Are we under a roof?
    bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
        MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
        MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
        
    Call DibujarMiniMapa
End Sub

Private Sub HandleSlotMenosUno()
    
    Dim Slot As Byte
    
    Call incomingData.ReadByte
    
    Slot = incomingData.ReadByte
    
    If Slot > 200 Then
        Slot = Slot - 200
        Call Inventario.UnSetSlot(Slot)
    Else
        Call Inventario.SetSlotAmount(Slot, Inv(Slot).Amount - 1)
    End If
    
    If Tomando Then
        If Inv(Slot).ObjType = otBebida Or Inv(Slot).ObjType = otPocion Or Inv(Slot).ObjType = otBotellaLlena Then
            Call Audio.Play(SND_DRINK)
            Tomando = False
        End If
    End If
    
End Sub

Private Sub HandleInventory()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Dim i As Byte
        
    Dim Slot As Byte
    Dim ObjIndex As Integer
    Dim GrhIndex As Integer
    Dim Name As String
    Dim Amount As Integer
    Dim Value As Long
    Dim ObjType As Byte
    
    Dim MinHit As Integer
    Dim MaxHit As Integer
    Dim MinDef As Byte
    Dim MaxDef As Byte
    Dim PuedeUsar As Boolean
    Dim Proyectil As Boolean
    
    Dim SR As RECT
    Dim DR As RECT

    SR.Right = 32
    SR.Bottom = 32
 
    DR.Right = 32
    DR.Bottom = 32
    
    Call Buffer.ReadByte
    
    NroItems = 0
    
    With HeadEqp
        .ObjIndex = Buffer.ReadInteger
     
        If .ObjIndex > 0 Then
            .GrhIndex = Buffer.ReadInteger
            .Name = Buffer.ReadASCIIString
            .Valor = Buffer.ReadLong
            .ObjType = Buffer.ReadByte
            
            .MinDef = Buffer.ReadByte
            .MaxDef = Buffer.ReadByte
        
            .Amount = 1
            .PuedeUsar = True
            
            frmMain.lblHeadEqp.Caption = .MinDef & "/" & .MaxDef
            frmMain.picHeadEqp.Picture = frmMain.picHeadEqp.Picture
            
            Call DrawTransparentGrhtoHdc(frmMain.picHeadEqp.hdc, 0, 0, .GrhIndex, SR)
    
            frmMain.picHeadEqp.Refresh
        End If

    End With
    
    With BodyEqp
        .ObjIndex = Buffer.ReadInteger
        
        If .ObjIndex > 0 Then
            .GrhIndex = Buffer.ReadInteger
            .Name = Buffer.ReadASCIIString
            .Valor = Buffer.ReadLong
            .ObjType = Buffer.ReadByte
            
            .MinDef = Buffer.ReadByte
            .MaxDef = Buffer.ReadByte
            
            .Amount = 1
            .PuedeUsar = True
            
            frmMain.lblBodyEqp.Caption = .MinDef & "/" & .MaxDef
            frmMain.picBodyEqp.Picture = frmMain.picBodyEqp.Picture
            
            Call DrawTransparentGrhtoHdc(frmMain.picBodyEqp.hdc, 0, 0, .GrhIndex, SR)
    
            frmMain.picBodyEqp.Refresh
        End If
    End With
    
    With LeftHandEqp
        .ObjIndex = Buffer.ReadInteger
        
        If .ObjIndex > 0 Then
            .GrhIndex = Buffer.ReadInteger
            .Name = Buffer.ReadASCIIString
            .Valor = Buffer.ReadLong
            .ObjType = Buffer.ReadByte
            
            If .ObjType = otArma Then
                .MinHit = Buffer.ReadByte
                .MaxHit = Buffer.ReadByte
                
                .Proyectil = True
                    
                frmMain.lblLeftHandEqp.Caption = .MinHit & "/" & .MaxHit
                frmMain.picLeftHandEqp.Picture = frmMain.picLeftHandEqp.Picture
            Else
                .MinDef = Buffer.ReadByte
                .MaxDef = Buffer.ReadByte
                
                frmMain.lblLeftHandEqp.Caption = .MinDef & "/" & .MaxDef
                frmMain.picLeftHandEqp.Picture = frmMain.picLeftHandEqp.Picture
            End If
            
            .Amount = 1
            .PuedeUsar = True
    
            Call DrawTransparentGrhtoHdc(frmMain.picLeftHandEqp.hdc, 0, 0, .GrhIndex, SR)
    
            frmMain.picLeftHandEqp.Refresh
        End If
    End With
    
    With RightHandEqp
        .ObjIndex = Buffer.ReadInteger
        
        If .ObjIndex > 0 Then
            .GrhIndex = Buffer.ReadInteger
            .Name = Buffer.ReadASCIIString
            .Valor = Buffer.ReadLong
            .ObjType = Buffer.ReadByte
            
            .MinHit = Buffer.ReadByte
            .MaxHit = Buffer.ReadByte
            
            If .ObjIndex > 0 Then
                If .ObjType = otFlecha Then
                    .Amount = Buffer.ReadInteger
                End If
            End If
            
            .PuedeUsar = True
            
            If .Amount > 0 Then
                frmMain.lblRightHandEqp.Caption = .Amount
            Else
                frmMain.lblRightHandEqp.Caption = .MinHit & "/" & .MaxHit
            End If
            
            frmMain.picRightHandEqp.Picture = frmMain.picRightHandEqp.Picture
            
            Call DrawTransparentGrhtoHdc(frmMain.picRightHandEqp.hdc, 0, 0, .GrhIndex, SR)
    
            frmMain.picRightHandEqp.Refresh
        End If
    End With
    
    With BeltEqp
        .ObjIndex = Buffer.ReadInteger
            
        If .ObjIndex > 0 Then
            .GrhIndex = Buffer.ReadInteger
            .Name = Buffer.ReadASCIIString
            .Valor = Buffer.ReadLong
            .ObjType = Buffer.ReadByte
            
            .Amount = 1
            .PuedeUsar = True
            
            frmMain.lblBeltEqp.Caption = .MinHit & "/" & .MaxHit
            frmMain.picBeltEqp.Picture = frmMain.picBeltEqp.Picture
            
            Call DrawTransparentGrhtoHdc(frmMain.picBeltEqp.hdc, 0, 0, .GrhIndex, SR)
    
            frmMain.picBeltEqp.Refresh
        End If
    End With
    
    With RingEqp
        .ObjIndex = Buffer.ReadInteger
        
        If .ObjIndex > 0 Then
            .GrhIndex = Buffer.ReadInteger
            .Name = Buffer.ReadASCIIString
            .Valor = Buffer.ReadLong
            .ObjType = Buffer.ReadByte
            
            .MinDef = Buffer.ReadByte
            .MaxDef = Buffer.ReadByte
            
            .Amount = 1
            .PuedeUsar = True
            
            frmMain.lblRingEqp.Caption = .MinHit & "/" & .MaxHit
            frmMain.picRingEqp.Picture = frmMain.picRingEqp.Picture
            
            Call DrawTransparentGrhtoHdc(frmMain.picRingEqp.hdc, 0, 0, .GrhIndex, SR)
    
            frmMain.picRingEqp.Refresh
        End If
    End With
    
    'With Ship
        '.ObjIndex = Buffer.ReadInteger
        
        'If .ObjIndex > 0 Then
            '.GrhIndex = Buffer.ReadInteger
            '.Name = Buffer.ReadASCIIString
            '.Valor = Buffer.ReadLong
            '.ObjType = Buffer.ReadByte
            '.MinHit = Buffer.ReadByte
            '.MaxHit = Buffer.ReadByte
            '.MinDef = Buffer.ReadByte
            '.MaxDef = Buffer.ReadByte
            '.PuedeUsar = True
            
            'frmMain.lblShip.Caption = .MinHit & "/" & .MaxHit
            'frmMain.picShip.Picture = frmMain.picShip.Picture
            
            'Call DrawTransparentGrhtoHdc(frmMain.picShip.hdc, 0, 0, .GrhIndex, SR)
    
            'frmMain.picShip.Refresh
        'End If
    'End With
    
    For i = 1 To MaxInvSlots + 1
        
        Slot = Buffer.ReadByte
        
        If Slot < 1 Then
            Exit For
        End If
        
        ObjIndex = Buffer.ReadInteger
        GrhIndex = Buffer.ReadInteger
        Name = Buffer.ReadASCIIString
        Amount = Buffer.ReadInteger
        Value = Buffer.ReadLong
        ObjType = Buffer.ReadByte
    
        MinHit = 0
        MaxHit = 0
        MinDef = 0
        MaxDef = 0
        PuedeUsar = True
        
        Select Case ObjType
                    
            Case Is > 100
                ObjType = ObjType - 100
                
                MinHit = Buffer.ReadInteger
                MaxHit = Buffer.ReadInteger
                
                If ObjType = otArma Then
                    Proyectil = Buffer.ReadBoolean
                End If
                
                PuedeUsar = Buffer.ReadBoolean
                
            Case otArma
                MinHit = Buffer.ReadByte
                MaxHit = Buffer.ReadByte
                
                Proyectil = Buffer.ReadBoolean
                
                PuedeUsar = Buffer.ReadBoolean
                
            Case otFlecha
                MinHit = Buffer.ReadByte
                MaxHit = Buffer.ReadByte
                
                PuedeUsar = Buffer.ReadBoolean
                
            Case otBarco
                MinHit = Buffer.ReadByte
                MaxHit = Buffer.ReadByte
                
            Case otArmadura, otCasco, otEscudo, otAnillo
                MinDef = Buffer.ReadByte
                MaxDef = Buffer.ReadByte
                
                PuedeUsar = Buffer.ReadBoolean
                
        End Select
        
        NroItems = NroItems + 1
        
        Call Inventario.SetSlot(Slot, ObjIndex, Amount, GrhIndex, ObjType, MinHit, MaxHit, MinDef, MaxDef, Value, Name, PuedeUsar, Proyectil)
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleInventorySlot()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
        
    Dim Slot As Byte
    Dim ObjIndex As Integer
    Dim GrhIndex As Integer
    Dim Name As String
    Dim Amount As Integer
    Dim Value As Long
    Dim ObjType As Byte
    
    Dim MinHit As Integer
    Dim MaxHit As Integer
    Dim MinDef As Byte
    Dim MaxDef As Byte
    Dim PuedeUsar As Boolean
    Dim Proyectil As Boolean
    
    Slot = Buffer.ReadByte
    
    ObjIndex = Buffer.ReadInteger
    GrhIndex = Buffer.ReadInteger
    Name = Buffer.ReadASCIIString
    Amount = Buffer.ReadInteger
    Value = Buffer.ReadLong
    ObjType = Buffer.ReadByte

    Select Case ObjType
                
        Case Is > 100
            ObjType = ObjType - 100
            
            MinHit = Buffer.ReadInteger
            MaxHit = Buffer.ReadInteger
            
            If ObjType = otArma Then
                Proyectil = Buffer.ReadBoolean
            End If
            
            PuedeUsar = Buffer.ReadBoolean
            
        Case otArma
            MinHit = Buffer.ReadByte
            MaxHit = Buffer.ReadByte
            
            Proyectil = Buffer.ReadBoolean
            
            PuedeUsar = Buffer.ReadBoolean
            
        Case otFlecha
            MinHit = Buffer.ReadByte
            MaxHit = Buffer.ReadByte
            
            PuedeUsar = Buffer.ReadBoolean
            
        Case otBarco
            MinHit = Buffer.ReadByte
            MaxHit = Buffer.ReadByte
            
            PuedeUsar = True
            
        Case otArmadura, otCasco, otEscudo, otAnillo
            MinDef = Buffer.ReadByte
            MaxDef = Buffer.ReadByte
            
            PuedeUsar = Buffer.ReadBoolean
            
        Case Else
            PuedeUsar = True
            
    End Select
    
    If Inv(Slot).ObjIndex = ObjIndex Then
        Call Inventario.SetSlotAmount(Slot, Amount)
    Else
        Call Inventario.SetSlot(Slot, ObjIndex, Amount, GrhIndex, ObjType, MinHit, MaxHit, MinDef, MaxDef, Value, Name, PuedeUsar, Proyectil)
    
        If Inv(Slot).ObjIndex < 1 Then
            NroItems = NroItems + 1
        End If
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleBeltInv()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Dim i As Byte
        
    Dim Slot As Byte
    Dim ObjIndex As Integer
    Dim GrhIndex As Integer
    Dim Name As String
    Dim Amount As Integer
    Dim Value As Long
    
    Call Buffer.ReadByte
    
    NroBeltItems = 0
    
    For i = 1 To MaxBeltSlots + 1
        
        Slot = Buffer.ReadByte

        If Slot < 1 Then
            Exit For
        End If
        
        ObjIndex = Buffer.ReadInteger
        GrhIndex = Buffer.ReadInteger
        Name = Buffer.ReadASCIIString
        Amount = Buffer.ReadInteger
        Value = Buffer.ReadLong
        
        If ObjIndex > 0 Then
            Call SetBeltSlot(Slot, ObjIndex, Amount, GrhIndex, Value, Name, True)
            NroBeltItems = NroBeltItems + 1
        End If
                
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleBeltSlot()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
        
    Dim Slot As Byte
    Dim ObjIndex As Integer
    Dim GrhIndex As Integer
    Dim Name As String
    Dim Amount As Integer
    Dim Value As Long
    
    Slot = Buffer.ReadByte
    
    ObjIndex = Buffer.ReadInteger
    GrhIndex = Buffer.ReadInteger
    Name = Buffer.ReadASCIIString
    Amount = Buffer.ReadInteger
    Value = Buffer.ReadLong
    
    If Belt(Slot).ObjIndex < 1 Then
        NroBeltItems = NroBeltItems + 1
        Call SetBeltSlot(Slot, ObjIndex, Amount, GrhIndex, Value, Name, True)
        
    ElseIf Belt(Slot).ObjIndex = ObjIndex Then
        Call Cinturon.SetSlotAmount(Slot, Amount)
        
    ElseIf ObjIndex < 1 Then
        NroBeltItems = NroBeltItems - 1
        Call Cinturon.UnSetSlot(Slot)
    End If
        
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleBank()

On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Dim i As Byte
        
    Dim Slot As Byte
            
    Dim ObjIndex As Integer
    Dim Name As String
    Dim Amount As Integer
    Dim GrhIndex As Integer
    Dim ObjType As eObjType
    Dim MaxHit As Integer
    Dim MinHit As Integer
    Dim MaxDef As Byte
    Dim MinDef As Byte
    Dim Valor As Long
    Dim PuedeUsar As Boolean
        
    Call InvNpc.Initialize(DirectDraw, frmBanco.PicBancoInv, MaxBankSlots)
    
    Call Buffer.ReadByte
    
    For i = 1 To MaxNpcInvSlots + 1
        
        Slot = Buffer.ReadByte

        If Slot < 1 Then
            Exit For
        End If
            
        ObjIndex = Buffer.ReadInteger
        GrhIndex = Buffer.ReadInteger
        Name = Buffer.ReadASCIIString
        Amount = Buffer.ReadInteger
        Valor = Buffer.ReadLong
        ObjType = Buffer.ReadByte

        MinHit = 0
        MaxHit = 0
        MinDef = 0
        MaxDef = 0
        PuedeUsar = True
        
        Select Case ObjType
                    
            Case Is > 100
                ObjType = ObjType - 100
                
                MinHit = Buffer.ReadInteger
                MaxHit = Buffer.ReadInteger
                
                If ObjType <> otBarco Then
                    PuedeUsar = Buffer.ReadBoolean
                End If
                
            Case otArma, otFlecha
                MinHit = Buffer.ReadByte
                MaxHit = Buffer.ReadByte
                
                PuedeUsar = Buffer.ReadBoolean
                
            Case otBarco
                MinHit = Buffer.ReadByte
                MaxHit = Buffer.ReadByte
                
            Case otArmadura, otCasco, otEscudo, otAnillo
                MinDef = Buffer.ReadByte
                MaxDef = Buffer.ReadByte
                
                PuedeUsar = Buffer.ReadBoolean
            
        End Select
    
        Call InvNpc.SetSlot(Slot, ObjIndex, Amount, GrhIndex, ObjType, _
        MinHit, MaxHit, MinDef, MaxDef, Valor, Name, PuedeUsar, False)
        
    Next i
    
    UserBankGold = Buffer.ReadLong
    frmBanco.lblUserBankGold = PonerPuntos(UserBankGold)

    If Not frmMain.PicInv.Visible Then
        Call Audio.Play(SND_CLICK)
    
        frmMain.imgInv.Visible = False
        frmMain.PicSpellInv.Visible = False
        
        frmMain.PicInv.Visible = True
    End If

    'Set state and show form
    Comerciando = True
    
    frmBanco.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleBankSlot()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim Slot As Byte
    
    Dim ObjIndex As Integer
    Dim Name As String
    Dim Amount As Integer
    Dim GrhIndex As Integer
    Dim ObjType As eObjType
    Dim MaxHit As Integer
    Dim MinHit As Integer
    Dim MaxDef As Byte
    Dim MinDef As Byte
    Dim Valor As Long
    Dim PuedeUsar As Boolean
    
    Slot = Buffer.ReadByte
    
    ObjIndex = Buffer.ReadInteger
    GrhIndex = Buffer.ReadInteger
    Name = Buffer.ReadASCIIString
    Amount = Buffer.ReadInteger
    Valor = Buffer.ReadLong
    ObjType = Buffer.ReadByte

    MinHit = 0
    MaxHit = 0
    MinDef = 0
    MaxDef = 0
    PuedeUsar = True
    
    Select Case ObjType
                
        Case Is > 100
            ObjType = ObjType - 100
            
            MinHit = Buffer.ReadInteger
            MaxHit = Buffer.ReadInteger
            
            If ObjType <> otBarco Then
                PuedeUsar = Buffer.ReadBoolean
            End If
            
        Case otArma, otFlecha
            MinHit = Buffer.ReadByte
            MaxHit = Buffer.ReadByte
            
            PuedeUsar = Buffer.ReadBoolean
            
        Case otBarco
            MinHit = Buffer.ReadByte
            MaxHit = Buffer.ReadByte
            
        Case otArmadura, otCasco, otEscudo, otAnillo
            MinDef = Buffer.ReadByte
            MaxDef = Buffer.ReadByte
            
            PuedeUsar = Buffer.ReadBoolean
        
    End Select

    Call InvNpc.SetSlot(Slot, ObjIndex, Amount, GrhIndex, ObjType, _
        MinHit, MaxHit, MinDef, MaxDef, Valor, Name, PuedeUsar, False)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleNpcInventory()

On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Dim i As Byte
        
    Dim Slot As Byte

    Dim ObjIndex As Integer
    Dim Name As String
    Dim Amount As Integer
    Dim GrhIndex As Integer
    Dim ObjType As eObjType
    Dim MaxHit As Integer
    Dim MinHit As Integer
    Dim MaxDef As Byte
    Dim MinDef As Byte
    Dim Valor As Long
    Dim PuedeUsar As Boolean
        
    Call InvNpc.Initialize(DirectDraw, frmComerciar.PicComercianteInv, MaxNpcInvSlots)
    
    Call Buffer.ReadByte
    
    For i = 1 To MaxNpcInvSlots + 1
        
        Slot = Buffer.ReadByte

        If Slot < 1 Then
            Exit For
        End If
            
        ObjIndex = Buffer.ReadInteger
        GrhIndex = Buffer.ReadInteger
        Name = Buffer.ReadASCIIString
        Amount = Buffer.ReadInteger
        Valor = Buffer.ReadLong
        ObjType = Buffer.ReadByte

        MinHit = 0
        MaxHit = 0
        MinDef = 0
        MaxDef = 0
        PuedeUsar = True
        
        Select Case ObjType
                    
            Case Is > 100
                ObjType = ObjType - 100
                
                MinHit = Buffer.ReadInteger
                MaxHit = Buffer.ReadInteger
                
                If ObjType <> otBarco Then
                    PuedeUsar = Buffer.ReadBoolean
                End If
                
            Case otArma, otFlecha
                MinHit = Buffer.ReadByte
                MaxHit = Buffer.ReadByte
                
                PuedeUsar = Buffer.ReadBoolean
                
            Case otBarco
                MinHit = Buffer.ReadByte
                MaxHit = Buffer.ReadByte
                
            Case otArmadura, otCasco, otEscudo, otAnillo
                MinDef = Buffer.ReadByte
                MaxDef = Buffer.ReadByte
                
                PuedeUsar = Buffer.ReadBoolean
            
        End Select
    
        Call InvNpc.SetSlot(Slot, ObjIndex, Amount, GrhIndex, ObjType, _
        MinHit, MaxHit, MinDef, MaxDef, Valor, Name, PuedeUsar, False)
        
    Next i

    frmComerciar.lblNpcName.Caption = Buffer.ReadASCIIString

    If frmMain.imgInv.Visible Then
        Call Audio.Play(SND_CLICK)
    
        frmMain.imgInv.Visible = False
        frmMain.PicSpellInv.Visible = False
        
        frmMain.PicInv.Visible = True
    End If

    'Set state and show form
    Comerciando = True
    
    frmComerciar.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleNpcInventorySlot()
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim Slot As Byte

    Dim ObjIndex As Integer
    Dim Name As String
    Dim Amount As Integer
    Dim GrhIndex As Integer
    Dim ObjType As eObjType
    Dim MaxHit As Integer
    Dim MinHit As Integer
    Dim MaxDef As Byte
    Dim MinDef As Byte
    Dim Valor As Long
    Dim PuedeUsar As Boolean
            
    Slot = Buffer.ReadByte
    
    ObjIndex = Buffer.ReadInteger
    GrhIndex = Buffer.ReadInteger
    Name = Buffer.ReadASCIIString
    Amount = Buffer.ReadInteger
    Valor = Buffer.ReadLong
    ObjType = Buffer.ReadByte

    MinHit = 0
    MaxHit = 0
    MinDef = 0
    MaxDef = 0
    PuedeUsar = True
    
    Select Case ObjType
                
        Case Is > 100
            ObjType = ObjType - 100
            
            MinHit = Buffer.ReadInteger
            MaxHit = Buffer.ReadInteger
            
            If ObjType <> otBarco Then
                PuedeUsar = Buffer.ReadBoolean
            End If
            
        Case otArma, otFlecha
            MinHit = Buffer.ReadByte
            MaxHit = Buffer.ReadByte
            
            PuedeUsar = Buffer.ReadBoolean
            
        Case otBarco
            MinHit = Buffer.ReadByte
            MaxHit = Buffer.ReadByte
            
        Case otArmadura, otCasco, otEscudo, otAnillo
            MinDef = Buffer.ReadByte
            MaxDef = Buffer.ReadByte
            
            PuedeUsar = Buffer.ReadBoolean
        
    End Select

    Call InvNpc.SetSlot(Slot, ObjIndex, Amount, GrhIndex, ObjType, _
        MinHit, MaxHit, MinDef, MaxDef, Valor, Name, PuedeUsar, False)
        
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleSpells()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim i As Byte
    Dim Slot As Byte
    Dim Spell As Byte
    
    NroSpells = 0
    
    For i = 1 To MaxSpellSlots
        
        Slot = Buffer.ReadByte
        
        Spell = Buffer.ReadByte
    
        If Spell > 0 Then
            Call SetSpellSlot(i, Spell + 24000, Buffer.ReadASCIIString, Buffer.ReadBoolean, Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger)
        Else
            Call Hechizos.UnSetSlot(i)
        End If
        
        NroSpells = NroSpells + 1
        
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleSpellSlot()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim Slot As Byte
    
    Slot = Buffer.ReadByte
    
    With Spell(Slot)
        .Nombre = Buffer.ReadASCIIString
                
        .Grh = Buffer.ReadByte + 24000
        .MinSkill = Buffer.ReadBoolean
        .ManaRequerido = Buffer.ReadInteger
        .StaRequerido = Buffer.ReadInteger
        .NeedStaff = Buffer.ReadInteger
        
        If .Grh > 24032 Then
            .Grh = 24032
        End If
            
        If .Grh = 24031 Then
            .Grh = 24032
        End If
        
        If .Grh = 24003 Then
            .Grh = 24002
        End If

        If .ManaRequerido > UserMinMan Or .StaRequerido > UserMinSTA Then
           .PuedeLanzar = False
        Else
           .PuedeLanzar = True
        End If
        
        If .Grh > 0 Then
            Call Hechizos.DrawSpellSlot(Slot)
        End If
        
        'Call frmMain.lstSpells.AddItem(.Nombre)
    End With
    
    NroSpells = NroSpells + 1

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleCompas()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Dim i As Byte
    Dim Slot As Byte
    
    Call Buffer.ReadByte
    
    NroCompas = 0
    
    For i = 1 To MaxCompaSlots + 1
        Slot = Buffer.ReadByte
        
        If Slot < 1 Then
            Debug.Print "Compa O"
            Exit For
        End If

        Call SetCompaSlot(Slot, Buffer.ReadASCIIString, Buffer.ReadBoolean)
        
        NroCompas = NroCompas + 1
    Next i
            
    Call RenderCompas
            
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:

    Call incomingData.CopyBuffer(Buffer)

    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleAddCompa()

    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim Slot As Byte

    Dim Agregado As Boolean
        
    Slot = Buffer.ReadByte
    
    Call SetCompaSlot(Slot, Buffer.ReadASCIIString, Buffer.ReadBoolean)
    
    If Buffer.ReadBoolean Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(Compa(Slot).Nombre & " te agregó como compañero, para eliminarlo escribe '-" & Compa(Slot).Nombre & "'.", .Red, .Green, .Blue, .Bold, .Italic)
        End With
    End If
    
    NroCompas = NroCompas + 1

    Call RenderCompas
    
    Call incomingData.CopyBuffer(Buffer)
    
End Sub

Private Sub HandleQuitarCompa()

On Error Resume Next

    Call incomingData.ReadByte
    
    Dim Slot As Byte
    Dim Nombre As String
    
    Slot = incomingData.ReadByte
    
    Nombre = Compa(Slot).Nombre
    
    Call UnSetCompaSlot(Slot)
    
    'If Buffer.ReadBoolean Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("Ya no eres compañero de " & Nombre & ".", .Red, .Green, .Blue, .Bold, .Italic)
        End With
    'End If
            
    NroCompas = NroCompas - 1
    
    Call RenderCompas
    
End Sub

Private Sub HandleCompaConnected()

On Error Resume Next

    Call incomingData.ReadByte
    
    Dim Slot As Byte
        
    Slot = incomingData.ReadByte
    
    Compa(Slot).Online = True
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg("Entró " & Compa(Slot).Nombre & ".", .Red, .Green, .Blue, .Bold, .Italic)
    End With
    
    Call RenderCompas
    
End Sub

Private Sub HandleCompaDisconnected()

On Error Resume Next

    Call incomingData.ReadByte
    
    Dim Slot As Byte
        
    Slot = incomingData.ReadByte
    
    Compa(Slot).Online = False
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg(Compa(Slot).Nombre & " salió.", .Red, .Green, .Blue, .Bold, .Italic)
    End With
    
    Call RenderCompas
        
End Sub

Private Sub HandleStopWorking()
    Call incomingData.ReadByte
    
    If frmMain.MacroTrabajo.Enabled Then
        Call frmMain.DesactivarMacroTrabajo
    End If
End Sub

Private Sub HandleCancelOfferItem()

End Sub

Private Sub HandleDeleteFile()
'SEGURIDAD
On Error GoTo ER
    
    Call incomingData.ReadByte

    Kill incomingData.ReadASCIIString
    
ER:
End Sub

Private Sub HandleAttributes()

    Call incomingData.ReadByte
    
    Dim i As Long
    
    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = incomingData.ReadByte
    Next i
    
    LlegaronAtrib = True
End Sub

Private Sub HandleUserPlatforms()

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim i As Byte
    Dim NroMapa As Integer
    
    NroPlataformas = 0

    For i = 1 To MaxPlataformSlots
        NroMapa = Buffer.ReadInteger
        
        If NroMapa < 1 Then
            Exit For
        Else
            Plataforma(i) = NroMapa
            NroPlataformas = NroPlataformas + 1
        End If
        
        Call frmPlataforma.lstLugares.AddItem(Replace(GetVar(MapPath & "Mapa" & NroMapa & ".dat", "Mapa" & NroMapa, "Name"), ".", vbNullString))
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

    frmPlataforma.Show , frmMain

End Sub

Private Sub HandleBlacksmithWeapons()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Count = Buffer.ReadInteger
    
    ReDim ArmasHerrero(Count) As tItemsConstruibles
    ReDim HerreroMejorar(0) As tItemsConstruibles
    
    For i = 1 To Count
        With ArmasHerrero(i)
            .Name = Buffer.ReadASCIIString    'Get the object's name
            .GrhIndex = Buffer.ReadInteger
            .LinH = Buffer.ReadInteger        'The iron needed
            .LinP = Buffer.ReadInteger        'The silver needed
            .LinO = Buffer.ReadInteger        'The gold needed
            .ObjIndex = Buffer.ReadInteger
            .Upgrade = Buffer.ReadInteger
        End With
    Next i
    
    With frmHerrero
        'Inicializo los inventarios
        Call InvLingosHerreria(1).Initialize(DirectDraw, .picLingotes0, 3, , , , , , False)
        Call InvLingosHerreria(2).Initialize(DirectDraw, .picLingotes1, 3, , , , , , False)
        Call InvLingosHerreria(3).Initialize(DirectDraw, .picLingotes2, 3, , , , , , False)
        Call InvLingosHerreria(4).Initialize(DirectDraw, .picLingotes3, 3, , , , , , False)
        
        Call .HideExtraControls(Count)
        Call .RenderList(1, True)
    End With
    
    For i = 1 To Count
        With ArmasHerrero(i)
            If .Upgrade Then
                For k = 1 To Count
                    If .Upgrade = ArmasHerrero(k).ObjIndex Then
                        j = j + 1
                
                        ReDim Preserve HerreroMejorar(j) As tItemsConstruibles
                        
                        HerreroMejorar(j).Name = .Name
                        HerreroMejorar(j).GrhIndex = .GrhIndex
                        HerreroMejorar(j).ObjIndex = .ObjIndex
                        HerreroMejorar(j).UpgradeName = ArmasHerrero(k).Name
                        HerreroMejorar(j).UpgradeGrhIndex = ArmasHerrero(k).GrhIndex
                        HerreroMejorar(j).LinH = ArmasHerrero(k).LinH - .LinH * 0.85
                        HerreroMejorar(j).LinP = ArmasHerrero(k).LinP - .LinP * 0.85
                        HerreroMejorar(j).LinO = ArmasHerrero(k).LinO - .LinO * 0.85
                        
                        Exit For
                    End If
                Next k
            End If
        End With
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleBlacksmithArmors()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Count = Buffer.ReadInteger
    
    ReDim ArmadurasHerrero(Count) As tItemsConstruibles
    
    For i = 1 To Count
        With ArmadurasHerrero(i)
            .Name = Buffer.ReadASCIIString    'Get the object's name
            .GrhIndex = Buffer.ReadInteger
            .LinH = Buffer.ReadInteger        'The iron needed
            .LinP = Buffer.ReadInteger        'The silver needed
            .LinO = Buffer.ReadInteger        'The gold needed
            .ObjIndex = Buffer.ReadInteger
            .Upgrade = Buffer.ReadInteger
        End With
    Next i
    
    j = UBound(HerreroMejorar)
    
    For i = 1 To Count
        With ArmadurasHerrero(i)
            If .Upgrade Then
                For k = 1 To Count
                    If .Upgrade = ArmadurasHerrero(k).ObjIndex Then
                        j = j + 1
                
                        ReDim Preserve HerreroMejorar(j) As tItemsConstruibles
                        
                        HerreroMejorar(j).Name = .Name
                        HerreroMejorar(j).GrhIndex = .GrhIndex
                        HerreroMejorar(j).ObjIndex = .ObjIndex
                        HerreroMejorar(j).UpgradeName = ArmadurasHerrero(k).Name
                        HerreroMejorar(j).UpgradeGrhIndex = ArmadurasHerrero(k).GrhIndex
                        HerreroMejorar(j).LinH = ArmadurasHerrero(k).LinH - .LinH * 0.85
                        HerreroMejorar(j).LinP = ArmadurasHerrero(k).LinP - .LinP * 0.85
                        HerreroMejorar(j).LinO = ArmadurasHerrero(k).LinO - .LinO * 0.85
                        
                        Exit For
                    End If
                Next k
            End If
        End With
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
    frmHerrero.Show , frmMain
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleCarpenterObjects()
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Count = Buffer.ReadInteger
    
    ReDim ObjCarpintero(Count) As tItemsConstruibles
    ReDim CarpinteroMejorar(0) As tItemsConstruibles
    
    For i = 1 To Count
        With ObjCarpintero(i)
            .Name = Buffer.ReadASCIIString        'Get the object's name
            .GrhIndex = Buffer.ReadInteger
            .Madera = Buffer.ReadInteger          'The wood needed
            .MaderaElfica = Buffer.ReadInteger    'The elfic wood needed
            .ObjIndex = Buffer.ReadInteger
            .Upgrade = Buffer.ReadInteger
        End With
    Next i
    
    With frmCarp
        'Inicializo los inventarios
        Call InvMaderasCarpinteria(1).Initialize(DirectDraw, .picMaderas0, 2, , , , , , False)
        Call InvMaderasCarpinteria(2).Initialize(DirectDraw, .picMaderas1, 2, , , , , , False)
        Call InvMaderasCarpinteria(3).Initialize(DirectDraw, .picMaderas2, 2, , , , , , False)
        Call InvMaderasCarpinteria(4).Initialize(DirectDraw, .picMaderas3, 2, , , , , , False)
        
        Call .HideExtraControls(Count)
        Call .RenderList(1)
    End With
    
    For i = 1 To Count
        With ObjCarpintero(i)
            If .Upgrade Then
                For k = 1 To Count
                    If .Upgrade = ObjCarpintero(k).ObjIndex Then
                        j = j + 1
                
                        ReDim Preserve CarpinteroMejorar(j) As tItemsConstruibles
                        
                        CarpinteroMejorar(j).Name = .Name
                        CarpinteroMejorar(j).GrhIndex = .GrhIndex
                        CarpinteroMejorar(j).ObjIndex = .ObjIndex
                        CarpinteroMejorar(j).UpgradeName = ObjCarpintero(k).Name
                        CarpinteroMejorar(j).UpgradeGrhIndex = ObjCarpintero(k).GrhIndex
                        CarpinteroMejorar(j).Madera = ObjCarpintero(k).Madera - .Madera * 0.85
                        CarpinteroMejorar(j).MaderaElfica = ObjCarpintero(k).MaderaElfica - .MaderaElfica * 0.85
                        
                        Exit For
                    End If
                Next k
            End If
        End With
    Next i
    
    frmCarp.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleRestOK()
    
    Call incomingData.ReadByte
    
    Descansando = Not Descansando
End Sub

Private Sub HandleErrorMessage()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim box As String
    
    box = Buffer.ReadASCIIString
    
    If box = "PsNo" Then
        frmConnect.PasswordTxt.Text = vbNullString
    
        Call SendMessage(frmConnect.PasswordTxt.hWnd, &H7, ByVal 0&, ByVal 0)
    
        frmConnect.PasswordTxtBorder.BorderColor = &H80&
        'frmConnect.PsNo.Visible = True
    
    ElseIf box = "NmNo" Then
        frmConnect.NameTxtBorder.BorderColor = &H80&
    
        Call SendMessage(frmConnect.NameTxt.hWnd, &H7, ByVal 0&, ByVal 0)
    
        frmConnect.PasswordTxt.Text = vbNullString
    Else
        Call MsgBox(box)
    End If
    
    If frmCrearPersonaje.Visible Then
        frmCrearPersonaje.MousePointer = vbDefault
    End If
        
    frmConnect.MousePointer = vbDefault
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleBlind()

    
    Call incomingData.ReadByte
    
    UserCiego = True
End Sub

Private Sub HandleDumb()
    Call incomingData.ReadByte
    
    UserEstupido = True
End Sub

Private Sub HandleShowSignal()
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Call InitCartel(Buffer.ReadASCIIString, Buffer.ReadInteger)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleUpdateHungerAndThirst()
    Call incomingData.ReadByte
    
    UserMinSed = incomingData.ReadByte
    UserMinHam = incomingData.ReadByte
    
    frmMain.AGUALbl.Caption = UserMinSed & "%"
    frmMain.COMIDALbl.Caption = UserMinHam & "%"
End Sub

Private Sub HandleMiniStats()
    
On Error Resume Next

    Call incomingData.ReadByte
    
    With UserEstadisticas
        .Matados = incomingData.ReadLong
        .Muertes = incomingData.ReadLong
        .NpcsMatados = incomingData.ReadLong
        .Clase = ListaClases(incomingData.ReadByte)
        .PenaCarcel = incomingData.ReadLong
        .Silencio = incomingData.ReadLong
    End With
End Sub

Private Sub HandleSkillUp()
    
    Dim SkillId As Byte
    Dim SkillLvl As Byte
    
    Call incomingData.ReadByte

    SkillId = incomingData.ReadByte
    
    SkillLvl = incomingData.ReadByte
    
    With FontTypes(FontTypeNames.FONTTYPE_HABILIDAD)
        Call ShowConsoleMsg(SkillName(SkillId), .Red, .Green, .Blue, .Bold, .Italic, True)
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg(" subió a nivel ", .Red, .Green, .Blue, .Bold, .Italic, True)
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_NUMERO)
        Call ShowConsoleMsg(SkillLvl, .Red, .Green, .Blue, .Bold, .Italic, True)
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg(".", .Red, .Green, .Blue, .Bold, .Italic)
    End With
    
    UserSkills(SkillId) = SkillLvl
End Sub

Private Sub HandleLevelUp()
    Dim Pts As Byte
    Dim AumentoHP As Byte
    Dim AumentoSTA As Byte
    Dim AumentoMANA As Byte
    Dim AumentoHIT As Byte

    Call incomingData.ReadByte
    
    Pts = incomingData.ReadByte
    AumentoHP = incomingData.ReadByte
    AumentoSTA = incomingData.ReadByte
    AumentoMANA = incomingData.ReadByte
    AumentoHIT = incomingData.ReadByte
    UserPasarNivel = incomingData.ReadLong
    UserExp = incomingData.ReadLong
    
    UserLvl = UserLvl + 1
    
    frmMain.LvlLbl.Caption = UserLvl
    
    'Call RemoveDamage
    'Call Dialogos.CreateDialog("¡Nivel " & UserLvl & "!", UserCharIndex, RGB(100, 200, 100))
   
    If Pts > 0 Then
        SkillPoints = SkillPoints + Pts
        frmMain.imgAsignarSkill.Visible = True
    End If
    
    UserMaxHP = UserMaxHP + AumentoHP
    UserMinHP = UserMaxHP
    
    UserMaxSTA = UserMaxSTA + AumentoSTA
    UserMinSTA = UserMaxSTA
        
    If AumentoMANA > 0 Then
        UserMaxMan = UserMaxMan + AumentoMANA
        UserMinMan = UserMaxMan
        
        If Meditando Then
            Meditando = False
            Charlist(UserCharIndex).FxIndex = 0
        End If
    End If

    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg("Subiste " & AumentoHP & " puntos de salud, " & AumentoSTA & " puntos de estamina y " & AumentoMANA & " puntos de maná.", .Red, .Green, .Blue, .Bold, .Italic)
    
        Call ShowConsoleMsg("Tu vida aumentó en " & AumentoHP & " puntos.", .Red, .Green, .Blue, .Bold, .Italic)
        
        Call ShowConsoleMsg("Tu energía aumentó en " & AumentoSTA & " puntos.", .Red, .Green, .Blue, .Bold, .Italic)
        
        If AumentoMANA > 0 Then
            Call ShowConsoleMsg("Tu maná aumentó en " & AumentoMANA & " puntos.", .Red, .Green, .Blue, .Bold, .Italic)
        End If
        
        Call ShowConsoleMsg("Tu golpe mínimo y máximo aumentaron en " & AumentoHIT & " puntos.", .Red, .Green, .Blue, .Bold, .Italic)
    End With
End Sub

Private Sub HandleSetInvisible()
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger
    Charlist(CharIndex).Invisible = incomingData.ReadBoolean
    
    Exit Sub
    
    If CharIndex = UserCharIndex Then
        If Not Charlist(CharIndex).Invisible Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("Has vuelto a ser visible.", .Red, .Green, .Blue, .Bold, .Italic)
            End With
        End If
    End If
End Sub

Private Sub HandleSetParalized()
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger
    Charlist(CharIndex).Paralizado = incomingData.ReadBoolean
                
    If CharIndex = UserCharIndex Then
        UserParalizado = Charlist(CharIndex).Paralizado
        
        If Not UserParalizado Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("Has vuelto a tener movilidad.", .Red, .Green, .Blue, .Bold, .Italic)
            End With
        End If
    End If
End Sub

Private Sub HandleBlindNoMore()
    
    Call incomingData.ReadByte
    
    UserCiego = False
End Sub

Private Sub HandleDumbNoMore()
    
    Call incomingData.ReadByte
    
    UserEstupido = False
End Sub

Private Sub HandleSkills()
    Call incomingData.ReadByte
    
    Dim i As Long
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = incomingData.ReadByte
        PorcentajeSkills(i) = incomingData.ReadByte
    Next i
End Sub

Private Sub HandleFreeSkillPts()

On Error Resume Next

    Call incomingData.ReadByte
    
    SkillPoints = incomingData.ReadInteger
    
    frmSkills.puntos.Caption = SkillPoints
    frmMain.imgAsignarSkill.Visible = True
End Sub

Private Sub HandleTrainerCreatureList()
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim creatures() As String
    Dim i As Long
    
    creatures = Split(Buffer.ReadASCIIString, SEPARATOR)
    
    For i = 0 To UBound(creatures())
        Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
    Next i
    frmEntrenador.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleGuildNews()
'Handles the GuildNews message.
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim GuildList() As String
    Dim i As Long
    Dim sTemp As String
    
    'Get news'string
    frmGuildNews.news = Buffer.ReadASCIIString
    
    'Get Enemy guilds list
    GuildList = Split(Buffer.ReadASCIIString, SEPARATOR)
    
    For i = 0 To UBound(GuildList)
        sTemp = frmGuildNews.txtClanesGuerra.Text
        frmGuildNews.txtClanesGuerra.Text = sTemp & GuildList(i) & vbCrLf
    Next i
    
    'Get Allied guilds list
    GuildList = Split(Buffer.ReadASCIIString, SEPARATOR)
    
    For i = 0 To UBound(GuildList)
        sTemp = frmGuildNews.txtClanesAliados.Text
        frmGuildNews.txtClanesAliados.Text = sTemp & GuildList(i) & vbCrLf
    Next i
    
    'If ClientSetup.GuildNews Then
    'frmGuildNews.Show vbModeless, frmMain
    'End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleOfferDetails()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleAlianceProposalsList()
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim vsGuildList() As String
    Dim i As Long
    
    vsGuildList = Split(Buffer.ReadASCIIString, SEPARATOR)
    
    Call frmPeaceProp.lista.Clear
    For i = 0 To UBound(vsGuildList())
        Call frmPeaceProp.lista.AddItem(vsGuildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.ALIANZA
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandlePeaceProposalsList()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim GuildList() As String
    Dim i As Long
    
    GuildList = Split(Buffer.ReadASCIIString, SEPARATOR)
    
    Call frmPeaceProp.lista.Clear
    For i = 0 To UBound(GuildList())
        Call frmPeaceProp.lista.AddItem(GuildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.PAZ
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleCharInfo()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    With frmCharInfo
        If .frmType = CharInfoFrmType.frmMembers Then
            .imgRechazar.Visible = False
            .imgAceptar.Visible = False
            .imgEchar.Visible = True
            .imgPeticion.Visible = False
        Else
            .imgRechazar.Visible = True
            .imgAceptar.Visible = True
            .imgEchar.Visible = False
            .imgPeticion.Visible = True
        End If
        
        .Nombre.Caption = Buffer.ReadASCIIString
        .Raza.Caption = ListaRazas(Buffer.ReadByte)
        .Clase.Caption = ListaClases(Buffer.ReadByte)
        
        If Buffer.ReadByte = 1 Then
            .Genero.Caption = "Hombre"
        Else
            .Genero.Caption = "Mujer"
        End If
        
        .Nivel.Caption = Buffer.ReadByte
        .Oro.Caption = Buffer.ReadLong
        .Banco.Caption = Buffer.ReadLong
        
        .txtPeticiones.Text = Buffer.ReadASCIIString
        .guildactual.Caption = Buffer.ReadASCIIString
        .txtMiembro.Text = Buffer.ReadASCIIString
        
        .Matados.Caption = CStr(Buffer.ReadLong)
        
        .Muertes.Caption = CStr(Buffer.ReadLong)
        
        Call .Show(vbModeless, frmMain)
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleGuildLeaderInfo()
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim i As Long
    Dim List() As String
    
    With frmGuildLeader
        'Get list of existing guilds
        GuildNames = Split(Buffer.ReadASCIIString, SEPARATOR)
        
        'Empty the list
        Call .GuildsList.Clear
        
        For i = 0 To UBound(GuildNames())
            Call .GuildsList.AddItem(GuildNames(i))
        Next i
        
        'Get list of guild's members
        GuildMembers = Split(Buffer.ReadASCIIString, SEPARATOR)
        .Miembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
        'Empty the list
        Call .members.Clear
        
        For i = 0 To UBound(GuildMembers())
            Call .members.AddItem(GuildMembers(i))
        Next i
        
        .txtGuildNews = Buffer.ReadASCIIString
        
        'Get list of join requests
        List = Split(Buffer.ReadASCIIString, SEPARATOR)
        
        'Empty the list
        Call .Solicitudes.Clear
        
        For i = 0 To UBound(List())
            Call .Solicitudes.AddItem(List(i))
        Next i
        
        .Show , frmMain
    End With

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleGuildDetails()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    With frmGuildBrief
        .imgDeclararGuerra.Visible = .EsLeader
        .imgOfrecerAlianza.Visible = .EsLeader
        .imgOfrecerPaz.Visible = .EsLeader
        
        .Nombre.Caption = Buffer.ReadASCIIString
        .fundador.Caption = Buffer.ReadASCIIString
        .Creacion.Caption = Buffer.ReadASCIIString
        .lider.Caption = Buffer.ReadASCIIString
        .Miembros.Caption = Buffer.ReadInteger
        
        If Buffer.ReadBoolean Then
            .eleccion.Caption = "ABIERTA"
        Else
            .eleccion.Caption = "CERRADA"
        End If
        
        .Enemigos.Caption = Buffer.ReadInteger
        .Aliados.Caption = Buffer.ReadInteger
        
        .Desc.Text = Buffer.ReadASCIIString
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
    frmGuildBrief.Show vbModeless, frmMain
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleShowGuildFundationForm()
    Call incomingData.ReadByte
    
    CreandoGuilda = True
    frmGuildFoundation.Show , frmMain
End Sub

Private Sub HandleShowUserRequest()
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString)
    Call frmUserRequest.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleChangeUserTradeSlot()

End Sub

Private Sub HandleSpawnList()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim creatureList() As String
    Dim i As Long
    
    creatureList = Split(Buffer.ReadASCIIString, SEPARATOR)
    
    For i = 0 To UBound(creatureList())
        Call frmSpawnList.lstCriaturas.AddItem(creatureList(i))
    Next i
    frmSpawnList.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
End Sub

Private Sub HandleShowSOSForm()
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim sosList() As String
    Dim i As Long
    
    sosList = Split(Buffer.ReadASCIIString, SEPARATOR)
    
    For i = 0 To UBound(sosList())
        Call frmMSG.List1.AddItem(sosList(i))
    Next i
    
    frmMSG.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleShowPartyForm()
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim members() As String
    Dim i As Long
    
    EsPartyLeader = CBool(Buffer.ReadByte)
       
    members = Split(Buffer.ReadASCIIString, SEPARATOR)
    For i = 0 To UBound(members())
        Call frmParty.lstMembers.AddItem(members(i))
    Next i
    
    frmParty.lblTotalExp.Caption = Buffer.ReadLong
    frmParty.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleShowMOTDEditionForm()

On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    frmCambiaMotd.txtMotd.Text = Buffer.ReadASCIIString
    frmCambiaMotd.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleShowGMPanelForm()

    
    Call incomingData.ReadByte

    frmPanelGm.Show vbModeless, frmMain
End Sub

Private Sub HandleUserNameList()
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim userList() As String
    Dim i As Long
    
    userList = Split(Buffer.ReadASCIIString, SEPARATOR)
    
    If frmPanelGm.Visible Then
        frmPanelGm.cboListaUsus.Clear
        For i = 0 To UBound(userList())
            Call frmPanelGm.cboListaUsus.AddItem(userList(i))
        Next i
        If frmPanelGm.cboListaUsus.ListCount > 0 Then
            frmPanelGm.cboListaUsus.ListIndex = 0
        End If
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandlePong()

    Call incomingData.ReadByte
    
    Call ShowConsoleMsg("Ping: " & (GetTickCount - pingTime) & " ms.", 255, 0, 0, True, False, False)
    
    pingTime = 0
End Sub

Private Sub HandleGuildMemberInfo()
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    With frmGuildMember
        'Clear guild's list
        .lstClanes.Clear
        
        GuildNames = Split(Buffer.ReadASCIIString, SEPARATOR)
        
        Dim i As Long
        For i = 0 To UBound(GuildNames())
            Call .lstClanes.AddItem(GuildNames(i))
        Next i
        
        'Get list of guild's members
        GuildMembers = Split(Buffer.ReadASCIIString, SEPARATOR)
        .lblCantMiembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
        'Empty the list
        Call .lstMiembros.Clear
        
        For i = 0 To UBound(GuildMembers())
            Call .lstMiembros.AddItem(GuildMembers(i))
        Next i
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
        
        .Show vbModeless, frmMain
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Private Sub HandleUpdateTagAndStatus()
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
    
    Dim CharIndex As Integer

    CharIndex = Buffer.ReadInteger

    'Update char status adn tag!
    With Charlist(CharIndex)
        .Nombre = Buffer.ReadASCIIString
        .Guilda = Buffer.ReadASCIIString
        .AlineacionGuilda = Buffer.ReadByte
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error > 0 Then
        'Err.Raise Error
    End If
End Sub

Public Sub WriteLoginChar()
    With outgoingData
        Call .WriteByte(ClientPacketID.Connect)
        Call .WriteByte(ConnectPacketID.LoginChar)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(UserPassword)
        Call .WriteASCIIString(CStr(GetDriveSerialNumber))
    End With
End Sub

Public Sub WriteLoginNewChar()
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.Connect)
        Call .WriteByte(ConnectPacketID.LoginNewChar)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(UserPassword)
        
        Call .WriteByte(UserRaza)
        Call .WriteByte(UserSexo)
        Call .WriteByte(UserClase)
        Call .WriteInteger(UserHead)
        
        For i = 1 To NUMATRIBUTOS
            Call .WriteByte(UserAtributos(i))
        Next i
        
        Call .WriteASCIIString(UserEmail)
        
        Call .WriteASCIIString(CStr(GetDriveSerialNumber))
    End With
End Sub

Public Sub WriteRecoverChar(UserName, UserEmail)
    With outgoingData
        Call .WriteByte(ClientPacketID.Connect)
        Call .WriteByte(ConnectPacketID.RecoverChar)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(UserEmail)
    End With
End Sub

Public Sub WriteKillChar()
    With outgoingData
        Call .WriteByte(ClientPacketID.Connect)
        Call .WriteByte(ConnectPacketID.KillChar)
        
        Call .WriteASCIIString(UserPassword)
        Call .WriteASCIIString(UserEmail)
    End With
End Sub

Public Sub WriteRequestRandomName()
    With outgoingData
        Call .WriteByte(ClientPacketID.Connect)
        Call .WriteByte(ConnectPacketID.RequestRandomName)
    End With
End Sub

Public Sub WriteOnline()
    Call outgoingData.WriteByte(ClientPacketID.Online)
End Sub

Public Sub WriteTalk(ByVal Chat As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.Talk)
        Call .WriteASCIIString(Chat)
    End With
End Sub

Public Sub WriteCompaMessage(ByVal Slot As Byte, ByVal Chat As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.CompaMessage)
        
        Call .WriteByte(Slot)
        Call .WriteASCIIString(Chat)
    End With
End Sub

Public Sub WritePrivateMessage(ByVal CharName As String, ByVal Chat As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.PrivateMessage)
        
        Call .WriteASCIIString(CharName)
        Call .WriteASCIIString(Chat)
    End With
End Sub

Public Sub WritePublicMessage(ByVal Message As String)
    Call outgoingData.WriteByte(ClientPacketID.PublicMessage)
    Call outgoingData.WriteASCIIString(Message)
End Sub

Public Sub WriteWalk(ByVal Heading As eHeading)
    With outgoingData
        
        Select Case Heading
            
            Case eHeading.NORTH
                Call .WriteByte(ClientPacketID.WalkNorth)

            Case eHeading.EAST
                Call .WriteByte(ClientPacketID.WalkEast)

            Case eHeading.SOUTH
                Call .WriteByte(ClientPacketID.WalkSouth)

            Case eHeading.WEST
                Call .WriteByte(ClientPacketID.WalkWest)

        End Select
        
    End With
End Sub

Public Sub WriteDeleteChat()
    Call outgoingData.WriteByte(ClientPacketID.DeleteChat)
End Sub
Public Sub WriteRequestPositionUpdate()
    Call outgoingData.WriteByte(ClientPacketID.RequestPositionUpdate)
End Sub

Public Sub WriteAttack()
    Call outgoingData.WriteByte(ClientPacketID.Attack)
End Sub

Public Sub WritePickUp()
    Call outgoingData.WriteByte(ClientPacketID.PickUp)
End Sub

Public Sub WriteResuscitationToggle()
    Call outgoingData.WriteByte(ClientPacketID.ResuscitationSafeToggle)
End Sub

Public Sub WriteRequestGuildLeaderInfo()
    Call outgoingData.WriteByte(ClientPacketID.RequestGuildLeaderInfo)
End Sub

Public Sub WriteRequestPartyForm()
    Call outgoingData.WriteByte(ClientPacketID.RequestPartyForm)
End Sub

Public Sub WriteItemUpgrade(ByVal ItemIndex As Integer)
    Call outgoingData.WriteByte(ClientPacketID.ItemUpgrade)
    Call outgoingData.WriteInteger(ItemIndex)
End Sub

Public Sub WriteRequestAttributes()
    Call outgoingData.WriteByte(ClientPacketID.RequestAttributes)
End Sub

Public Sub WriteRequestSkills()
    Call outgoingData.WriteByte(ClientPacketID.RequestSkills)
End Sub

Public Sub WriteRequestMiniStats()
    Call outgoingData.WriteByte(ClientPacketID.RequestMiniStats)
End Sub

Public Sub WriteCommerceEnd()
    Call outgoingData.WriteByte(ClientPacketID.CommerceEnd)
End Sub

Public Sub WriteUserCommerceEnd()
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceEnd)
End Sub

Public Sub WriteUserCommerceConfirm()
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceConfirm)
End Sub

Public Sub WriteBankEnd()
    Call outgoingData.WriteByte(ClientPacketID.BankEnd)
End Sub
Public Sub WriteUserCommerceOk()
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceOk)
End Sub

Public Sub WriteUserCommerceReject()
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceReject)
End Sub

Public Sub WriteDrop(ByVal Slot As Byte, ByVal Amount As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.Drop)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteDropGold(ByVal Amount As Long)
    Call outgoingData.WriteByte(ClientPacketID.DropGold)
    Call outgoingData.WriteLong(Amount)
End Sub

Public Sub WriteLeftClick(ByVal x As Integer, ByVal y As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.LeftClick)
        
        Call .WriteInteger(x)
        Call .WriteInteger(y)
    End With
End Sub

Public Sub WriteRightClick(ByVal x As Integer, ByVal y As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.RightClick)
        
        Call .WriteInteger(x)
        Call .WriteInteger(y)
    End With
End Sub

Public Sub WriteDoubleClick(ByVal x As Integer, ByVal y As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.DoubleClick)
        
        Call .WriteInteger(x)
        Call .WriteInteger(y)
    End With
End Sub

Public Sub WriteWork(ByVal Skill As eSkill)
    With outgoingData
        Call .WriteByte(ClientPacketID.Work)
        
        Call .WriteByte(Skill)
    End With
End Sub

Public Sub WriteUseSpellMacro()
    Call outgoingData.WriteByte(ClientPacketID.UseSpellMacro)
End Sub

Public Sub WriteUseItem(ByVal Slot As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.UseItem)
        Call .WriteByte(Slot)
    End With
End Sub

Public Sub WriteUseBeltItem(ByVal Slot As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.UseBeltItem)
        Call .WriteByte(Slot)
    End With
End Sub

Public Sub WriteCraftBlacksmith(ByVal Item As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftBlacksmith)
        Call .WriteInteger(Item)
    End With
End Sub

Public Sub WriteCraftCarpenter(ByVal Item As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftCarpenter)
        
        Call .WriteInteger(Item)
    End With
End Sub

Public Sub WriteShowGuildNews()
     outgoingData.WriteByte (ClientPacketID.ShowGuildNews)
End Sub

Public Sub WriteWorkLeftClick(ByVal x As Integer, ByVal y As Integer, ByVal Skill As eSkill)
    With outgoingData
        Call .WriteByte(ClientPacketID.WorkLeftClick)
        
        Call .WriteInteger(x)
        Call .WriteInteger(y)
        
        Call .WriteByte(Skill)
    End With
End Sub

Public Sub WriteCreateNewGuild(ByVal Desc As String, ByVal Name As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNewGuild)
        
        Call .WriteASCIIString(Desc)
        Call .WriteASCIIString(Name)
    End With
End Sub

Public Sub WriteCastSpell(ByVal Spell As Byte, ByVal x As Integer, ByVal y As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.CastSpell)
        
        Call .WriteByte(Spell)
        Call .WriteInteger(x)
        Call .WriteInteger(y)
    End With
End Sub

Public Sub WriteSpellInfo(ByVal Slot As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.SpellInfo)
        
        Call .WriteByte(Slot)
    End With
End Sub

Public Sub WriteEquipItem(ByVal Slot As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.EquipItem)
        
        Call .WriteByte(Slot)
    End With
End Sub

Public Sub WriteUnEquipItem(ByVal ObjType As eObjType)
    With outgoingData
        Call .WriteByte(ClientPacketID.UnEquipItem)
        
        Call .WriteByte(ObjType)
    End With
End Sub

Public Sub WriteChangeHeading(ByVal Heading As eHeading)
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeHeading)
        
        Call .WriteByte(Heading)
    End With
End Sub

Public Sub WrItemodifySkills(ByRef skillEdt() As Byte)
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ModifySkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(skillEdt(i))
        Next i
    End With
End Sub

Public Sub WriteTrain(ByVal creature As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.Train)
        
        Call .WriteByte(creature)
    End With
End Sub

Public Sub WriteCommerceBuy(ByVal Slot As Byte, ByVal Amount As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceBuy)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteBankExtractItem(ByVal Slot As Byte, ByVal Amount As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractItem)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteCommerceSell(ByVal Slot As Byte, ByVal Amount As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceSell)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
    End With
    
    Call Audio.Play(SND_SELL)
End Sub

Public Sub WriteBankDepositItem(ByVal Slot As Byte, ByVal Amount As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDepositItem)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteMoveInvSlot(ByVal Slot As Byte, ByVal Slot2 As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveInvSlot)

        Call .WriteByte(Slot)
        Call .WriteByte(Slot2)
    End With
End Sub

Public Sub WriteMoveBeltSlot(ByVal Slot As Byte, ByVal Slot2 As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveBeltSlot)

        Call .WriteByte(Slot)
        Call .WriteByte(Slot2)
    End With
End Sub

Public Sub WriteMoveBankSlot(ByVal upwards As Boolean, ByVal Slot As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveBankSlot)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(Slot)
    End With
End Sub

Public Sub WriteMoveSpellSlot(ByVal Slot As Byte, ByVal Slot2 As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveSpellSlot)

        Call .WriteByte(Slot)
        Call .WriteByte(Slot2)
    End With
End Sub

Public Sub WriteGuildDescUpdate(ByVal Desc As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildDescUpdate)
        
        Call .WriteASCIIString(Desc)
    End With
End Sub

Public Sub WriteUserCommerceOffer(ByVal Slot As Byte, ByVal Amount As Long, ByVal OfferSlot As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.UserCommerceOffer)
        
        Call .WriteByte(Slot)
        Call .WriteLong(Amount)
        Call .WriteByte(OfferSlot)
    End With
End Sub

Public Sub WriteCommerceChat(ByVal Chat As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceChat)
        
        Call .WriteASCIIString(Chat)
    End With
End Sub

Public Sub WriteGuildAcceptPeace(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptPeace)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteGuildRejectAlliance(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectAlliance)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteGuildRejectPeace(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectPeace)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteGuildAcceptAlliance(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptAlliance)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteGuildOfferPeace(ByVal guild As String, ByVal proposal As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferPeace)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)
    End With
End Sub

Public Sub WriteGuildOfferAlliance(ByVal guild As String, ByVal proposal As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferAlliance)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)
    End With
End Sub

Public Sub WriteGuildAllianceDetails(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAllianceDetails)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteGuildPeaceDetails(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildPeaceDetails)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteGuildRequestJoinerInfo(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestJoinerInfo)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteGuildAlliancePropList()
    Call outgoingData.WriteByte(ClientPacketID.GuildAlliancePropList)
End Sub

Public Sub WriteGuildPeacePropList()
    Call outgoingData.WriteByte(ClientPacketID.GuildPeacePropList)
End Sub

Public Sub WriteGuildDeclareWar(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildDeclareWar)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteGuildAcceptNewMember(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptNewMember)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteGuildRejectNewMember(ByVal UserName As String, ByVal reason As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectNewMember)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
    End With
End Sub

Public Sub WriteGuildKickMember(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildKickMember)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteGuildUpdateNews(ByVal news As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildUpdateNews)
        Call .WriteASCIIString(news)
    End With
End Sub

Public Sub WriteGuildMemberInfo(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMemberInfo)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteGuildOpenElections()
    Call outgoingData.WriteByte(ClientPacketID.GuildOpenElections)
End Sub

Public Sub WriteGuildRequestMembership(ByVal guild As String, ByVal Application As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestMembership)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(Application)
    End With
End Sub

Public Sub WriteGuildRequestDetails(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestDetails)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteQuit()

    If UserParalizado Then 'Inmo
        If Charlist(UserCharIndex).Priv < 2 Then
            With FontTypes(FontTypeNames.FONTTYPE_WARNING)
                Call ShowConsoleMsg("No podés salir estando paralizado.", .Red, .Green, .Blue, .Bold, .Italic)
            End With
            Exit Sub
        End If
    End If
    
    If frmMain.MacroTrabajo.Enabled Then
        frmMain.DesactivarMacroTrabajo
    End If
    
    Call outgoingData.WriteByte(ClientPacketID.Quit)
End Sub

Public Sub WriteGuildLeave()
    Call outgoingData.WriteByte(ClientPacketID.GuildLeave)
End Sub

Public Sub WriteRequestAccountState()
    Call outgoingData.WriteByte(ClientPacketID.RequestAccountState)
End Sub

Public Sub WritePetStand(ByVal MascoIndex As Byte)
    Call outgoingData.WriteByte(ClientPacketID.PetStand)
    Call outgoingData.WriteByte(MascoIndex)
End Sub

Public Sub WritePetFollow(ByVal MascoIndex As Byte)
    Call outgoingData.WriteByte(ClientPacketID.PetFollow)
    Call outgoingData.WriteByte(MascoIndex)
End Sub

Public Sub WriteReleasePet()
    Call outgoingData.WriteByte(ClientPacketID.ReleasePet)
End Sub

Public Sub WriteRest()
    Call outgoingData.WriteByte(ClientPacketID.Rest)
End Sub

Public Sub WriteMeditate()
    Call outgoingData.WriteByte(ClientPacketID.Meditate)
End Sub

Public Sub WriteHome()
    Call outgoingData.WriteByte(ClientPacketID.Home)
End Sub

Public Sub WritePlatformTeleport(ByVal Index As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.PlatformTeleport)
        Call .WriteInteger(Plataforma(Index))
    End With
End Sub

Public Sub WriteConsulta()
    Call outgoingData.WriteByte(ClientPacketID.Consultation)
End Sub

Public Sub WriteHelp()
    Call outgoingData.WriteByte(ClientPacketID.Help)
End Sub

Public Sub WriteRequestStats()
    Call outgoingData.WriteByte(ClientPacketID.RequestStats)
End Sub

Public Sub WriteCommerceStart()
    Call outgoingData.WriteByte(ClientPacketID.CommerceStart)
End Sub

Public Sub WriteBankStart()
    Call outgoingData.WriteByte(ClientPacketID.BankStart)
End Sub

Public Sub WriteRequestMOTD()
    Call outgoingData.WriteByte(ClientPacketID.RequestMOTD)
End Sub

Public Sub WriteUpTime()
    Call outgoingData.WriteByte(ClientPacketID.UpTime)
End Sub

Public Sub WritePartyLeave()
    Call outgoingData.WriteByte(ClientPacketID.PartyLeave)
End Sub

Public Sub WritePartyCreate()
    Call outgoingData.WriteByte(ClientPacketID.PartyCreate)
End Sub

Public Sub WritePartyJoin()
    Call outgoingData.WriteByte(ClientPacketID.PartyJoin)
End Sub

Public Sub WriteInquiry()
    Call outgoingData.WriteByte(ClientPacketID.Inquiry)
End Sub

Public Sub WriteGuildMessage(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WritePartyMessage(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteCentinelReport(ByVal Number As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.CentinelReport)
        
        Call .WriteInteger(Number)
    End With
End Sub

Public Sub WriteGuildOnline()
    Call outgoingData.WriteByte(ClientPacketID.GuildOnline)
End Sub

Public Sub WritePartyOnline()
    Call outgoingData.WriteByte(ClientPacketID.PartyOnline)
End Sub

Public Sub WriteRoleMasterRequest(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.RoleMasterRequest)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteGMRequest()
    Call outgoingData.WriteByte(ClientPacketID.GMRequest)
End Sub

Public Sub WriteBugReport(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.BugReport)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteChangeDescription(ByVal Desc As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeDescription)
        
        Call .WriteASCIIString(Desc)
    End With
End Sub

Public Sub WriteGuildVote(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildVote)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WritePunishments(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.Punishments)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteChangePassword(ByRef oldPass As String, ByRef newPass As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangePassword)
        Call .WriteASCIIString(oldPass)
        Call .WriteASCIIString(newPass)
    End With
End Sub

Public Sub WriteGamble(ByVal Amount As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.Gamble)
        
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteInquiryVote(ByVal opt As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.InquiryVote)
        
        Call .WriteByte(opt)
    End With
End Sub

Public Sub WriteBankExtractGold(ByVal Amount As Long)
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractGold)
        
        Call .WriteLong(Amount)
    End With
End Sub

Public Sub WriteBankDepositGold(ByVal Amount As Long)
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDepositGold)
        
        Call .WriteLong(Amount)
    End With
End Sub

Public Sub WriteDenounce(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.Denounce)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteGuildFundate()
    Call outgoingData.WriteByte(ClientPacketID.GuildFundate)
End Sub

Public Sub WritePartyKick(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyKick)
            
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WritePartySetLeader(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.PartySetLeader)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WritePartyAcceptMember(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyAcceptMember)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteGuildMemberList(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.GuildMemberList)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteInitCrafting(ByVal Cantidad As Integer, ByVal NroPorCiclo As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.InitCrafting)
        
        Call .WriteInteger(Cantidad)
        Call .WriteInteger(NroPorCiclo)
    End With
End Sub

Public Sub WriteSearchObj(ByVal restrict As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.SearchObj)
        
        Call .WriteASCIIString(restrict)
    End With
End Sub

Public Sub WriteGMMessage(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.GMMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteShowName()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.ShowName)
End Sub

Public Sub WriteGoNearby(ByVal UserName As String)
    With outgoingData
        Call outgoingData.WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.GoNearby)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteComment(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.Comment)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteServerTime()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.ServerTime)
End Sub

Public Sub WriteWhere(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.Where)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteCreaturesInMap(ByVal map As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.CreaturesInMap)
        
        Call .WriteInteger(map)
    End With
End Sub

Public Sub WriteWarpMeToTarget()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.WarpMeToTarget)
End Sub

Public Sub WriteWarpChar(ByVal UserName As String, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.WarpChar)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
        
        Call .WriteInteger(map)
        
        Call .WriteInteger(x)
        Call .WriteInteger(y)
    End With
End Sub

Public Sub WriteSilence(ByVal UserName As String, ByVal TiempoSilencio As Long)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.Silence)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
        Call .WriteLong(TiempoSilencio)
    End With
End Sub

Public Sub WriteSOSShowList()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.SOSShowList)
End Sub

Public Sub WriteSOSRemove(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.SOSRemove)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteGoToChar(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.GoToChar)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteInvisible()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.Invisible)
End Sub

Public Sub WriteGMPanel()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.GMPanel)
End Sub

Public Sub WriteRequestUserList()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.RequestUserList)
End Sub

Public Sub WriteWorking()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.Working)
End Sub

Public Sub WriteHiding()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.Hiding)
End Sub

Public Sub WriteJail(ByVal UserName As String, ByVal reason As String, ByVal time As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.Jail)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
        Call .WriteASCIIString(reason)
        
        Call .WriteByte(time)
    End With
End Sub

Public Sub WriteKillNPC()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.KillNPC)
End Sub

Public Sub WriteWarnUser(ByVal UserName As String, ByVal reason As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.WarnUser)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
        Call .WriteASCIIString(reason)
    End With
End Sub

Public Sub WriteEditChar(ByVal UserName As String, ByVal EditOption As eEditOptions, ByVal arg1 As String, ByVal arg2 As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.EditChar)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
        
        Call .WriteByte(EditOption)
        
        Call .WriteASCIIString(arg1)
        Call .WriteASCIIString(arg2)
    End With
End Sub

Public Sub WriteRequestCharInfo(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.RequestCharInfo)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteRequestCharStats(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.RequestCharStats)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteRequestCharGold(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.RequestCharGold)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteRequestCharInv(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.RequestCharInventory)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteRequestCharBank(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.RequestCharBank)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteRequestCharSkills(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.RequestCharSkills)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteReviveChar(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.ReviveChar)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteOnlineGM()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.OnlineGM)
End Sub

Public Sub WriteOnlineMap(ByVal map As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.OnlineMap)
        
        Call .WriteInteger(map)
    End With
End Sub

Public Sub WriteKick(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.Kick)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteExecute(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.Execute)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteBanChar(ByVal UserName As String, ByVal reason As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.banChar)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
        
        Call .WriteASCIIString(reason)
    End With
End Sub

Public Sub WriteUnbanChar(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.UnbanChar)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteNPCFollow()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.NPCFollow)
End Sub

Public Sub WriteSummonChar(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.SummonChar)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteSpawnListRequest()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.SpawnListRequest)
End Sub

Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.SpawnCreature)
        
        Call .WriteInteger(creatureIndex)
    End With
End Sub

Public Sub WriteResetNpcInv()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.ResetNpcInv)
End Sub

Public Sub WriteCleanWorld()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.CleanWorld)
End Sub

Public Sub WriteServerMessage(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.ServerMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteNickToIP(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.NickToIP)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteIPToNick(ByRef Ip() As Byte)
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then
        Exit Sub   'Invalid IP
    End If
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.IPToNick)
        
        For i = LBound(Ip()) To UBound(Ip())
            Call .WriteByte(Ip(i))
        Next i
    End With
End Sub

Public Sub WriteGuildOnlineMembers(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.GuildOnlineMembers)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteTeleportCreate(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal Radio As Byte = 0)
    With outgoingData
            Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.TeleportCreate)
        
        Call .WriteInteger(map)
        
        Call .WriteInteger(x)
        Call .WriteInteger(y)
        
        Call .WriteByte(Radio)
    End With
End Sub

Public Sub WriteTeleportDestroy()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.TeleportDestroy)
End Sub

Public Sub WriteRainToggle()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.RainToggle)
End Sub

Public Sub WriteWeather(ByVal Estado As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.Weather)
        Call .WriteByte(Estado)
    End With
End Sub

Public Sub WriteSetCharDescription(ByVal Desc As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.SetCharDescription)
        
        Call .WriteASCIIString(Desc)
    End With
End Sub

Public Sub WriteForceMP3ToMap(ByVal MP3ID As Byte, ByVal map As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.ForceMP3ToMap)
        
        Call .WriteByte(MP3ID)
        
        Call .WriteInteger(map)
    End With
End Sub

Public Sub WriteForceWAVEToMap(ByVal waveID As Byte, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.ForceWAVEToMap)
        
        Call .WriteByte(waveID)
        
        Call .WriteInteger(map)
        
        Call .WriteInteger(x)
        Call .WriteInteger(y)
    End With
End Sub

Public Sub WriteTalkAsNPC(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.TalkAsNPC)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteDestroyAllItemsInArea()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.DestroyAllItemsInArea)
End Sub

Public Sub WrItemakeDumb(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.MakeDumb)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WrItemakeDumbNoMore(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.MakeDumbNoMore)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteDumpIPTables()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.dumpIPTables)
End Sub

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.SetTrigger)
        
        Call .WriteByte(Trigger)
    End With
End Sub

Public Sub WriteAskTrigger()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.AskTrigger)
End Sub

Public Sub WriteBannedIPList()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.BannedIPList)
End Sub

Public Sub WriteBannedIPReload()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.BannedIPReload)
End Sub

Public Sub WriteGuildBan(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.GuildBan)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteBanIP(ByVal byIp As Boolean, ByRef Ip() As Byte, ByVal Nick As String, ByVal reason As String)
    If byIp And UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then
        Exit Sub   'Invalid IP
    End If
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.BanIP)
        
        Call .WriteBoolean(byIp)
        
        If byIp Then
            For i = LBound(Ip()) To UBound(Ip())
                Call .WriteByte(Ip(i))
            Next i
        Else
            Call .WriteASCIIString(Nick)
        End If
        
        Call .WriteASCIIString(reason)
    End With
End Sub

Public Sub WriteUnbanIP(ByRef Ip() As Byte)
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then
        Exit Sub   'Invalid IP
    End If
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.UnbanIP)
        
        For i = LBound(Ip()) To UBound(Ip())
            Call .WriteByte(Ip(i))
        Next i
    End With
End Sub

Public Sub WriteCreateItem(ByVal ItemIndex As Integer, ByVal ItemAmount As Long)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.CreateItem)
        Call .WriteInteger(ItemIndex)
        Call .WriteLong(ItemAmount)
    End With
End Sub

Public Sub WriteDestroyItems()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.DestroyItems)
End Sub

Public Sub WriteForceMP3All(ByVal FileNumber As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.ForceMP3All)
        
        Call .WriteByte(FileNumber)
    End With
End Sub

Public Sub WriteForceWAVEAll(ByVal waveID As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.ForceWAVEAll)
        
        Call .WriteByte(waveID)
    End With
End Sub

Public Sub WriteRemovePunishment(ByVal UserName As String, ByVal punishment As Byte, ByVal NewText As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.RemovePunishment)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
        Call .WriteByte(punishment)
        Call .WriteASCIIString(NewText)
    End With
End Sub

Public Sub WriteTileBlockedToggle()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.TileBlockedToggle)
End Sub

Public Sub WriteKillNPCNoRespawn()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.KillNPCNoRespawn)
End Sub

Public Sub WriteKillAllNearbyNPCs()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.KillAllNearbyNPCs)
End Sub

Public Sub WriteLastIP(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.LastIP)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteChangeMOTD()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.ChangeMOTD)
End Sub

Public Sub WriteSetMOTD(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.SetMOTD)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteSystemMessage(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.SystemMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteCreateNPC(ByVal NPCIndex As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.CreateNPC)
        
        Call .WriteInteger(NPCIndex)
    End With
End Sub

Public Sub WriteCreateNPCWithRespawn(ByVal NPCIndex As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.CreateNPCWithRespawn)
        
        Call .WriteInteger(NPCIndex)
    End With
End Sub

Public Sub WriteServerOpenToUsersToggle()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.ServerOpenToUsersToggle)
End Sub

Public Sub WriteTurnOffServer()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.TurnOffServer)
End Sub

Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.RemoveCharFromGuild)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteRequestCharMail(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.RequestCharMail)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
    End With
End Sub

Public Sub WriteAlterPassword(ByVal UserName As String, ByVal CopyFrom As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.AlterPassword)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
        Call .WriteASCIIString(CopyFrom)
    End With
End Sub

Public Sub WriteAlterMail(ByVal UserName As String, ByVal newMail As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.AlterMail)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
        Call .WriteASCIIString(newMail)
    End With
End Sub

Public Sub WriteAlterName(ByVal UserName As String, ByVal newName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.AlterName)
        
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
        Call .WriteASCIIString(newName)
    End With
End Sub

Public Sub WriteToggleCentinelActivated()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.ToggleCentinelActivated)
End Sub

Public Sub WriteDoBackup()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.DoBackUp)
End Sub

Public Sub WriteShowGuildMessages(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.ShowGuildMessages)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteSaveMap()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.SaveMap)
End Sub

Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.ChangeMapInfoPK)
        
        Call .WriteBoolean(isPK)
    End With
End Sub

Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.ChangeMapInfoBackup)
        
        Call .WriteBoolean(backup)
    End With
End Sub

Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.ChangeMapInfoRestricted)
        
        Call .WriteASCIIString(restrict)
    End With
End Sub

Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.ChangeMapInfoNoMagic)
        
        Call .WriteBoolean(nomagic)
    End With
End Sub

Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.ChangeMapInfoNoInvi)
        
        Call .WriteBoolean(noinvi)
    End With
End Sub
                            
Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.ChangeMapInfoNoResu)
        
        Call .WriteBoolean(noresu)
    End With
End Sub

Public Sub WriteChangeMapInfoLand(ByVal land As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.ChangeMapInfoLand)
        
        Call .WriteASCIIString(land)
    End With
End Sub

Public Sub WriteChangeMapInfoZone(ByVal Zone As String)
'zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.ChangeMapInfoZone)
        
        Call .WriteASCIIString(Zone)
    End With
End Sub

Public Sub WriteSaveChars()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.SaveChars)
End Sub

Public Sub WriteCleanSOS()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.CleanSOS)
End Sub

Public Sub WriteShowServerForm()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.ShowServerForm)
End Sub

Public Sub WriteKickAllChars()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.KickAllChars)
End Sub

Public Sub WriteReloadNPCs()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.ReloadNPCs)
End Sub

Public Sub WriteReloadServidorIni()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.ReloadServidorIni)
End Sub

Public Sub WriteReloadSpells()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.ReloadSpells)
End Sub

Public Sub WriteReloadObjects()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.ReloadObjects)
End Sub

Public Sub WriteRestart()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.Restart)
End Sub

Public Sub WriteChatColor(ByVal r As Byte, ByVal g As Byte, ByVal B As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.ChatColor)
        
        Call .WriteByte(r)
        Call .WriteByte(g)
        Call .WriteByte(B)
    End With
End Sub

Public Sub WriteIgnored()
    Call outgoingData.WriteByte(ClientPacketID.GMCommand)
    Call outgoingData.WriteByte(GMPacketID.Ignored)
End Sub

Public Sub WriteCheckSlot(ByVal UserName As String, ByVal Slot As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.CheckSlot)
        Call .WriteASCIIString(StrConv(Trim$(UserName), vbProperCase))
        Call .WriteByte(Slot)
    End With
End Sub

Public Sub WritePing()
    'Prevent the timer from being cut
    If pingTime > 0 Then
        Exit Sub
    End If
    
    Call outgoingData.WriteByte(ClientPacketID.Ping)
    
    'Avoid computing Errors due to frame rate
    Call FlushBuffer
    DoEvents
    
    pingTime = GetTickCount
End Sub

Public Sub WriteSetIniVar(ByRef sLlave As String, ByRef sClave As String, ByRef sValor As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommand)
        Call .WriteByte(GMPacketID.SetIniVar)
        
        Call .WriteASCIIString(sLlave)
        Call .WriteASCIIString(sClave)
        Call .WriteASCIIString(sValor)
    End With
End Sub

Public Sub WriteAuctionCreate()
    Call outgoingData.WriteByte(ClientPacketID.AuctionCreate)
End Sub

Public Sub WriteAuctionBid(ByVal MoneyAmount As Long)
    With outgoingData
        Call .WriteByte(ClientPacketID.AuctionBid)
        
        Call .WriteLong(MoneyAmount)
    End With
End Sub

Public Sub WriteAuctionView()
    Call outgoingData.WriteByte(ClientPacketID.AuctionView)
End Sub

Public Sub WriteAniadirCompaniero(ByVal UserNameaAgregar As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.AniadirCompaniero)
        
        Call .WriteASCIIString(UserNameaAgregar)
    End With
End Sub

Public Sub WriteEliminarCompaniero(ByVal Slot As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.EliminarCompaniero)
        
        Call .WriteByte(Slot)
    End With
End Sub

Public Sub FlushBuffer()
'Sends all data existing in the buffer
    With outgoingData
        If .length = 0 Then
            Exit Sub
        End If
        
        Call SendData(.ReadASCIIStringFixed(.length))
    End With
End Sub

'Sends the data using the socket controls in the MainForm.
Private Sub SendData(ByRef sdData As String)
    
    'No enviamos nada si no estamos conectados
    If Not frmMain.Socket1.IsWritable Then
        'Put data back in the bytequeue
        Call outgoingData.WriteASCIIStringFixed(sdData)
        
        Exit Sub
    End If
    
    If Not frmMain.Socket1.Connected Then
        Exit Sub
    End If
    
    'Send data!
    Call frmMain.Socket1.Write(sdData, Len(sdData))

End Sub
