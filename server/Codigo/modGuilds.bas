Attribute VB_Name = "modGuilds"
Option Explicit

Private Const Max_Guilds As Byte = 255
'Cantidad Maxima de guilds en el servidor

Public CantidadDEGUILDAS As Byte
'Cantidad actual de guildas en el servidor

Public Guilds(1 To Max_Guilds) As clsGuild
'array global de guilds, se Indexa por UserList().Guild_Id

Public Const MaxASPIRANTES As Byte = 10
'Cantidad Maxima de aspirantes que puede tener una guilda acumulados a la vez

'numero de .wav del cliente
Public Enum SONIDOS_GUILD
    SND_CREACIONCLAN = 44
    SND_ACEPTADOCLAN = 43
    SND_DECLAREWAR = 45
End Enum

'estado entre guildas
Public Enum RELACIONES_GUILD
    GUERRA = -1
    PAZ = 0
    ALIADOS = 1
End Enum

Public GUILDPATH As String
Public GUILDINFOFILE As String

Public Sub LoadGuildsDB()

    Dim CantGuildas  As String
    Dim i           As Integer
    Dim TempStr     As String
    
    GUILDPATH = App.Path & "\GUILDS\"
    GUILDINFOFILE = GUILDPATH & "guildsinfo.inf"

    CantGuildas = GetVar(GUILDINFOFILE, "INIT", "nroGuilds")
    
    If IsNumeric(CantGuildas) Then
        CantidadDEGUILDAS = CByte(CantGuildas)
    Else
        CantidadDEGUILDAS = 0
    End If
    
    For i = 1 To CantidadDEGUILDAS
        Set Guilds(i) = New clsGuild
        TempStr = GetVar(GUILDINFOFILE, "GUILD" & i, "GuildName")
        Call Guilds(i).Inicializar(TempStr, i)
    Next i
    
End Sub

Public Sub m_DesconectarMiembroDelClan(ByVal UserIndex As Integer, ByVal Guild_Id As Integer)
    If UserList(UserIndex).Guild_Id > CantidadDEGUILDAS Then
        Exit Sub
    End If
    Call Guilds(Guild_Id).DesConectarMiembro(UserIndex)
End Sub

Private Function m_EsGuildLeader(ByRef PJ As String, ByVal Guild_Id As Integer) As Boolean
    m_EsGuildLeader = (UCase$(PJ) = UCase$(Trim$(Guilds(Guild_Id).GetLeader)))
End Function

Private Function m_EsGuildFounder(ByRef PJ As String, ByVal Guild_Id As Integer) As Boolean
    m_EsGuildFounder = (UCase$(PJ) = UCase$(Trim$(Guilds(Guild_Id).Fundador)))
End Function

Public Function m_EcharMiembroDeGuilda(ByVal Expulsador As Integer, ByVal Expulsado As String) As Integer
'UI echa a Expulsado del guilda de Expulsado
Dim UserIndex   As Integer
Dim GI          As Integer
    
    m_EcharMiembroDeGuilda = 0

    UserIndex = NameIndex(Expulsado)
    If UserIndex > 0 Then
        'pj online
        GI = UserList(UserIndex).Guild_Id
        If GI > 0 Then
            If m_PuedeSalirDeGuilda(Expulsado, GI, Expulsador) Then
                Call Guilds(GI).DesConectarMiembro(UserIndex)
                Call Guilds(GI).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " fue expulsado de " & Guilds(GI).GuildName & " Expulsador = " & Expulsador)
                UserList(UserIndex).Guild_Id = 0
                Call RefreshCharStatus(UserIndex)
                m_EcharMiembroDeGuilda = GI
            Else
                m_EcharMiembroDeGuilda = 0
            End If
        Else
            m_EcharMiembroDeGuilda = 0
        End If
    Else
        'pj offline
        GI = GetGuild_IdFromChar(Expulsado)
        If GI > 0 Then
            If m_PuedeSalirDeGuilda(Expulsado, GI, Expulsador) Then
                Call Guilds(GI).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " fue expulsado de " & Guilds(GI).GuildName & " Expulsador = " & Expulsador)
                m_EcharMiembroDeGuilda = GI
            Else
                m_EcharMiembroDeGuilda = 0
            End If
        Else
            m_EcharMiembroDeGuilda = 0
        End If
    End If

End Function

Public Sub ChangeDesc(ByRef Desc As String, ByVal Guild_Id As Integer)
    If Guild_Id < 1 Or Guild_Id > CantidadDEGUILDAS Then
        Exit Sub
    End If
    Call Guilds(Guild_Id).SetDesc(Desc)
End Sub

Public Sub ActualizarNoticias(ByVal UserIndex As Integer, ByRef Datos As String)
Dim GI              As Integer

    GI = UserList(UserIndex).Guild_Id
    
    If GI < 1 Or GI > CantidadDEGUILDAS Then
        Exit Sub
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        Exit Sub
    End If
    
    Call Guilds(GI).SetGuildNews(Datos)
        
End Sub

Public Function CrearNuevaGuilda(ByVal FundadorIndex As Integer, ByRef Desc As String, ByRef GuildName As String, ByRef refError As String) As Boolean
    
    Dim i               As Integer
    Dim DummyString     As String

    CrearNuevaGuilda = False
    If Not PuedeFundarUnaGuilda(FundadorIndex, DummyString) Then
        refError = DummyString
        Exit Function
    End If

    If GuildName = vbNullString Or Not GuildNameValido(GuildName) Then
        refError = "Nombre de guilda inválido."
        Exit Function
    End If
    
    If YaExiste(GuildName) Then
        refError = "Ya existe una guilda con ese nombre."
        Exit Function
    End If

    'tenemos todo para fundar ya
    If CantidadDEGUILDAS < UBound(Guilds) Then
        CantidadDEGUILDAS = CantidadDEGUILDAS + 1
        'ReDim Preserve Guilds(1 To CantidadDEGUILDAS) As clsGuild

        'constructor custom de la clase clan
        Set Guilds(CantidadDEGUILDAS) = New clsGuild
        
        With Guilds(CantidadDEGUILDAS)
            Call .Inicializar(GuildName, CantidadDEGUILDAS)
        
        'Damos de alta al guilda como nuevo inicializando sus archivos
            Call .InicializarNuevaGuilda(UserList(FundadorIndex).Name)
        
            Call .SetDesc(Desc)
            Call .SetLeader(UserList(FundadorIndex).Name)

        '"conectamos" al nuevo miembro a la lista de la clase
            Call .AceptarNuevoMiembro(UserList(FundadorIndex).Name)
            Call .ConectarMiembro(FundadorIndex)
        End With
        
        UserList(FundadorIndex).Guild_Id = CantidadDEGUILDAS
        Call RefreshCharStatus(FundadorIndex)
        
        For i = 1 To CantidadDEGUILDAS - 1
            Call Guilds(i).ProcesarFundacionDeOtraGuilda
        Next i
    Else
        refError = "No hay mas slots para fundar guildas. Consulte a un administrador."
        Exit Function
    End If
    
    'Open the database with an empty record and create the new user
    Call DB_RS_Open("SELECT * FROM guildas WHERE 0=1")
    DB_RS.AddNew
    
    'Put the data in the recordset
    DB_RS!Nombre = GuildName
    
    DB_RS.Update
    DB_RS_Close
    
    'UserList(FundadorIndex).Name
    
    CrearNuevaGuilda = True
End Function

Public Sub SendGuildNews(ByVal UserIndex As Integer)
Dim Guild_Id  As Integer
Dim i               As Integer
Dim go As Integer

    Guild_Id = UserList(UserIndex).Guild_Id
    
    If Guild_Id = 0 Then
        Exit Sub
    End If
    
    Dim enemies() As String
    
    With Guilds(Guild_Id)
        If .CantidadEnemys Then
            ReDim enemies(0 To .CantidadEnemys - 1) As String
    Else
        ReDim enemies(0)
    End If
    
    Dim allies() As String
    
        If .CantidadAllies Then
            ReDim allies(0 To .CantidadAllies - 1) As String
    Else
        ReDim allies(0)
    End If
    
        i = .Iterador_ProximaRelacion(RELACIONES_GUILD.GUERRA)
    go = 0
    
    While i > 0
        enemies(go) = Guilds(i).GuildName
            i = .Iterador_ProximaRelacion(RELACIONES_GUILD.GUERRA)
        go = go + 1
    Wend
    
        i = .Iterador_ProximaRelacion(RELACIONES_GUILD.ALIADOS)
    go = 0
    
    While i > 0
        allies(go) = Guilds(i).GuildName
            i = .Iterador_ProximaRelacion(RELACIONES_GUILD.ALIADOS)
    Wend

        Call WriteGuildNews(UserIndex, .GetGuildNews, enemies, allies)

        If .EleccionesAbiertas Then
        Call WriteConsoleMsg(UserIndex, "Hoy es la votacion para elegir un nuevo líder para la guilda!!.", FontTypeNames.FONTTYPE_GUILD)
        Call WriteConsoleMsg(UserIndex, "La eleccion durara 24 horas, se puede votar a cualquier miembro dla guilda.", FontTypeNames.FONTTYPE_GUILD)
        Call WriteConsoleMsg(UserIndex, "Para votar escribe /VOTO NICKNAME.", FontTypeNames.FONTTYPE_GUILD)
        Call WriteConsoleMsg(UserIndex, "Solo se computara un voto por miembro. Tu voto no puede ser cambiado.", FontTypeNames.FONTTYPE_GUILD)
    End If
    End With

End Sub

Public Function m_PuedeSalirDeGuilda(ByRef Nombre As String, ByVal Guild_Id As Integer, ByVal QuienLoEchaUI As Integer) As Boolean
'sale solo si no es fundador dla guilda.

    m_PuedeSalirDeGuilda = False
    
    If Guild_Id = 0 Then
        Exit Function
    End If
    
    'esto es un parche, si viene en -1 es porque la invoca la rutina de expulsion automatica de guildas x antifacciones
    If QuienLoEchaUI = -1 Then
        m_PuedeSalirDeGuilda = True
        Exit Function
    End If

    'cuando UI no puede echar a nombre?
    'si no es gm Y no es lider del guilda del pj Y no es el mismo que se va voluntariamente
    If UserList(QuienLoEchaUI).flags.Privilegios And PlayerType.User Then
        If Not m_EsGuildLeader(UCase$(UserList(QuienLoEchaUI).Name), Guild_Id) Then
            If UCase$(UserList(QuienLoEchaUI).Name) <> UCase$(Nombre) Then      'si no sale voluntariamente...
                Exit Function
            End If
        End If
    End If

    'Ahora el lider es el unico que no puede salir dla guilda
    m_PuedeSalirDeGuilda = UCase$(Guilds(Guild_Id).GetLeader) <> UCase$(Nombre)

End Function

Public Function PuedeFundarUnaGuilda(ByVal UserIndex As Integer, ByRef refError As String) As Boolean

    Exit Function

    PuedeFundarUnaGuilda = False
    If UserList(UserIndex).Guild_Id > 0 Then
        refError = "Ya perteneces a una guilda, no podés fundar otro."
        Exit Function
    End If
    
    If UserList(UserIndex).Stats.Elv < 35 Or UserList(UserIndex).Skills.Skill(eSkill.Liderazgo).Elv < 100 Then
        refError = "Para fundar una guilda tu nivel debe ser de 35 o más y tener 100 puntos de habilidad en Liderazgo."
        Exit Function
    End If
    
    If HasFound(UserList(UserIndex).Name) Then
        refError = "Ya has fundado una guilda, no podés fundar otra."
        Exit Function
    End If
    
    PuedeFundarUnaGuilda = True
    
End Function

Public Function Relacion2String(ByVal Relacion As RELACIONES_GUILD) As String
    Select Case Relacion
        Case RELACIONES_GUILD.ALIADOS
            Relacion2String = "A"
        Case RELACIONES_GUILD.GUERRA
            Relacion2String = "G"
        Case RELACIONES_GUILD.PAZ
            Relacion2String = "P"
        Case RELACIONES_GUILD.ALIADOS
            Relacion2String = "?"
    End Select
End Function

Public Function String2Relacion(ByVal S As String) As RELACIONES_GUILD
    Select Case UCase$(Trim$(S))
        Case vbNullString, "P"
            String2Relacion = RELACIONES_GUILD.PAZ
        Case "G"
            String2Relacion = RELACIONES_GUILD.GUERRA
        Case "A"
            String2Relacion = RELACIONES_GUILD.ALIADOS
        Case Else
            String2Relacion = RELACIONES_GUILD.PAZ
    End Select
End Function

Private Function GuildNameValido(ByVal cad As String) As Boolean
Dim car     As Byte
Dim i       As Integer

'old PUBLIC FUNCTION by morgo

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))

    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        GuildNameValido = False
        Exit Function
    End If
    
Next i

GuildNameValido = True

End Function

Private Function YaExiste(ByVal GuildName As String) As Boolean
    Dim i   As Integer
    
    YaExiste = False
    GuildName = UCase$(GuildName)
    
    For i = 1 To CantidadDEGUILDAS
        YaExiste = (UCase$(Guilds(i).GuildName) = GuildName)
        
        If YaExiste Then
            Exit Function
        End If
    Next i
End Function

Public Function HasFound(ByRef UserName As String) As Boolean
    Dim i As Long
    Dim Name As String
    
    Name = UCase$(UserName)
    
    For i = 1 To CantidadDEGUILDAS
        HasFound = (UCase$(Guilds(i).Fundador) = Name)
        If HasFound Then
            Exit Function
        End If
    Next i
End Function

Public Function v_AbrirElecciones(ByVal UserIndex As Integer, ByRef refError As String) As Boolean
    
    Dim Guild_Id As Integer

    v_AbrirElecciones = False
    Guild_Id = UserList(UserIndex).Guild_Id
    
    If Guild_Id = 0 Or Guild_Id > CantidadDEGUILDAS Then
        refError = "Tú no perteneces a ninguna guilda."
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, Guild_Id) Then
        refError = "No eres el líder de tu guilda."
        Exit Function
    End If
    
    If Guilds(Guild_Id).EleccionesAbiertas Then
        refError = "Las elecciones ya están abiertas."
        Exit Function
    End If
    
    v_AbrirElecciones = True
    Call Guilds(Guild_Id).AbrirElecciones
    
End Function

Public Function v_UsuarioVota(ByVal UserIndex As Integer, ByRef Votado As String, ByRef refError As String) As Boolean
    
    Dim Guild_Id As Integer
    Dim list() As String
    Dim i As Long

    v_UsuarioVota = False
    Guild_Id = UserList(UserIndex).Guild_Id
    
    If Guild_Id = 0 Or Guild_Id > CantidadDEGUILDAS Then
        refError = "Tú no perteneces a ninguna guilda."
        Exit Function
    End If

    With Guilds(Guild_Id)
        If Not .EleccionesAbiertas Then
            refError = "No hay elecciones abiertas en tu clan."
            Exit Function
        End If
        
        
        list = .GetMemberList()
        For i = 0 To UBound(list())
            If UCase$(Votado) = list(i) Then
                Exit For
            End If
        Next i
        
        If i > UBound(list()) Then
            refError = Votado & " no pertenece a la guilda."
            Exit Function
        End If
        
        
        If .YaVoto(UserList(UserIndex).Name) Then
            refError = "Ya has votado, no podés cambiar tu voto."
            Exit Function
        End If
        
        Call .ContabilizarVoto(UserList(UserIndex).Name, Votado)
        v_UsuarioVota = True
    End With

End Function

Public Sub v_RutinaElecciones()
    
    Dim i As Integer

On Error GoTo errh
    'Call SendData(SendTarget.ToAll, 0, Msg_ConsoleMsg("Servidor> Revisando elecciones", FontTypeNames.FONTTYPE_SERVER))
    For i = 1 To CantidadDEGUILDAS
        If Not Guilds(i) Is Nothing Then
            If Guilds(i).RevisarElecciones Then
                Call SendData(SendTarget.ToAll, 0, Msg_ConsoleMsg("Servidor> " & Guilds(i).GetLeader & " es el nuevo líder de " & Guilds(i).GuildName & ".", FontTypeNames.FONTTYPE_SERVER))
            End If
        End If
proximo:
    Next i
    'Call SendData(SendTarget.ToAll, 0, Msg_ConsoleMsg("Servidor> Elecciones revisadas.", FontTypeNames.FONTTYPE_SERVER))
Exit Sub
errh:
    Call LogError("modGuilds.v_RutinaElecciones():" & Err.description)
    Resume proximo
End Sub

Private Function GetGuild_IdFromChar(ByRef PlayerName As String) As Integer

'aca si que vamos a violar las capas deliveradamente ya que
'visual basic no permite declarar metodos de clase
Dim Temps   As String
    If InStrB(PlayerName, "/") > 0 Then
        PlayerName = Replace(PlayerName, "/", vbNullString)
    End If
    If InStrB(PlayerName, "/") > 0 Then
        PlayerName = Replace(PlayerName, "/", vbNullString)
    End If
    If InStrB(PlayerName, ".") > 0 Then
        PlayerName = Replace(PlayerName, ".", vbNullString)
    End If
    Temps = GetVar(CharPath & PlayerName & ".chr", "GUILD", "Guild_Id")
    If IsNumeric(Temps) Then
        GetGuild_IdFromChar = CInt(Temps)
    Else
        GetGuild_IdFromChar = 0
    End If
End Function

Public Function Guild_Id(ByRef GuildName As String) As Integer
'me da el indice del GuildName
Dim i As Integer

    Guild_Id = 0
    GuildName = UCase$(GuildName)
    For i = 1 To CantidadDEGUILDAS
        If UCase$(Guilds(i).GuildName) = GuildName Then
            Guild_Id = i
            Exit Function
        End If
    Next i
End Function

Public Function m_ListaDeMiembrosOnline(ByVal UserIndex As Integer, ByVal Guild_Id As Integer) As String
Dim i As Integer
    
    If Guild_Id > 0 And Guild_Id <= CantidadDEGUILDAS Then
        i = Guilds(Guild_Id).m_Iterador_ProximoUserIndex
        While i > 0
            'No mostramos dioses y admins
            If i <> UserIndex And ((UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) > 0 Or (UserList(UserIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) > 0)) Then _
                m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList(i).Name & ","
            i = Guilds(Guild_Id).m_Iterador_ProximoUserIndex
        Wend
    End If
    If Len(m_ListaDeMiembrosOnline) > 0 Then
        m_ListaDeMiembrosOnline = Left$(m_ListaDeMiembrosOnline, Len(m_ListaDeMiembrosOnline) - 1)
    End If
End Function

Public Function PrepareGuildsList() As String()
    Dim tStr() As String
    Dim i As Long
    
    If CantidadDEGUILDAS = 0 Then
        ReDim tStr(0) As String
    Else
        ReDim tStr(CantidadDEGUILDAS - 1) As String
        
        For i = 1 To CantidadDEGUILDAS
            tStr(i - 1) = Guilds(i).GuildName
        Next i
    End If
    
    PrepareGuildsList = tStr
End Function

Public Sub SendGuildDetails(ByVal UserIndex As Integer, ByRef GuildName As String)
    Dim GI      As Integer

    GI = Guild_Id(GuildName)
    
    If GI = 0 Then
        Exit Sub
    End If
    
    With Guilds(GI)
        Call Protocol.WriteGuildDetails(UserIndex, GuildName, .Fundador, .GetFechaFundacion, .GetLeader, _
                                    .CantidadDeMiembros, .EleccionesAbiertas, _
                                    .CantidadEnemys, .CantidadAllies, .GetDesc)
    End With
End Sub

Public Sub SendGuildLeaderInfo(ByVal UserIndex As Integer)

    Dim GI      As Integer
    Dim GuildList() As String
    Dim MemberList() As String
    Dim AspirantsList() As String

    With UserList(UserIndex)
        GI = .Guild_Id
        
        GuildList = PrepareGuildsList()
        
        If GI < 1 Or GI > CantidadDEGUILDAS Then
            'Send the guild list instead
            Call Protocol.WriteGuildList(UserIndex, GuildList)
            Exit Sub
        End If
        
        If Not m_EsGuildLeader(.Name, GI) Then
            'Send the guild list instead
            Call Protocol.WriteGuildList(UserIndex, GuildList)
            Exit Sub
        End If
        
        MemberList = Guilds(GI).GetMemberList()
        AspirantsList = Guilds(GI).GetAspirantes()
        
        Call WriteGuildLeaderInfo(UserIndex, GuildList, MemberList, Guilds(GI).GetGuildNews(), AspirantsList)
    End With
End Sub

Public Function m_Iterador_ProximoUserIndex(ByVal Guild_Id As Integer) As Integer
    'itera sobre los onlinemembers
    m_Iterador_ProximoUserIndex = 0
    If Guild_Id > 0 And Guild_Id <= CantidadDEGUILDAS Then
        m_Iterador_ProximoUserIndex = Guilds(Guild_Id).m_Iterador_ProximoUserIndex()
    End If
End Function

Public Function Iterador_ProximoGM(ByVal Guild_Id As Integer) As Integer
    'itera sobre los gms escuchando esta guilda
    Iterador_ProximoGM = 0
    If Guild_Id > 0 And Guild_Id <= CantidadDEGUILDAS Then
        Iterador_ProximoGM = Guilds(Guild_Id).Iterador_ProximoGM()
    End If
End Function

Public Function r_Iterador_ProximaPropuesta(ByVal Guild_Id As Integer, ByVal Tipo As RELACIONES_GUILD) As Integer
    'itera sobre las propuestas
    r_Iterador_ProximaPropuesta = 0
    If Guild_Id > 0 And Guild_Id <= CantidadDEGUILDAS Then
        r_Iterador_ProximaPropuesta = Guilds(Guild_Id).Iterador_ProximaPropuesta(Tipo)
    End If
End Function

Public Function GMEscuchaGuilda(ByVal UserIndex As Integer, ByVal GuildName As String) As Integer
Dim GI As Integer

    'listen to no guild at all
    If LenB(GuildName) = 0 And UserList(UserIndex).EscucheGuilda > 0 Then
        'Quit listening to previous guild!!
        Call WriteConsoleMsg(UserIndex, "Dejas de escuchar a: " & Guilds(UserList(UserIndex).EscucheGuilda).GuildName, FontTypeNames.FONTTYPE_GUILD)
        Guilds(UserList(UserIndex).EscucheGuilda).DesconectarGM (UserIndex)
        Exit Function
    End If
    
'devuelve el Guild_Id
    GI = Guild_Id(GuildName)
    If GI > 0 Then
        If UserList(UserIndex).EscucheGuilda > 0 Then
            If UserList(UserIndex).EscucheGuilda = GI Then
                'Already listening to them...
                Call WriteConsoleMsg(UserIndex, "Conectado a: " & GuildName, FontTypeNames.FONTTYPE_GUILD)
                GMEscuchaGuilda = GI
                Exit Function
            Else
                'Quit listening to previous guild!!
                Call WriteConsoleMsg(UserIndex, "Dejas de escuchar a: " & Guilds(UserList(UserIndex).EscucheGuilda).GuildName, FontTypeNames.FONTTYPE_GUILD)
                Guilds(UserList(UserIndex).EscucheGuilda).DesconectarGM (UserIndex)
            End If
        End If
        
        Call Guilds(GI).ConectarGM(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Conectado a: " & GuildName, FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaGuilda = GI
        UserList(UserIndex).EscucheGuilda = GI
    Else
        Call WriteConsoleMsg(UserIndex, "Error, el guilda '" & GuildName & "'no existe", FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaGuilda = 0
    End If
    
End Function

Public Sub GMDejaDeEscucharClan(ByVal UserIndex As Integer, ByVal Guild_Id As Integer)
'el Index lo tengo que tener de cuando me puse a escuchar
    UserList(UserIndex).EscucheGuilda = 0
    Call Guilds(Guild_Id).DesconectarGM(UserIndex)
End Sub
Public Function r_DeclararGuerra(ByVal UserIndex As Integer, ByRef GuildGuerra As String, ByRef refError As String) As Integer
Dim GI  As Integer
Dim GIG As Integer

    r_DeclararGuerra = 0
    GI = UserList(UserIndex).Guild_Id
    If GI < 1 Or GI > CantidadDEGUILDAS Then
        refError = "No eres miembro de ninguna guilda."
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu guilda."
        Exit Function
    End If
    
    If LenB(Trim$(GuildGuerra)) < 1 Then
        refError = "No has seleccionado ninguna guilda."
        Exit Function
    End If
    
    GIG = Guild_Id(GuildGuerra)
    If Guilds(GI).GetRelacion(GIG) = GUERRA Then
        refError = "Tu guilda ya está en guerra con " & GuildGuerra & "."
        Exit Function
    End If
        
    If GI = GIG Then
        refError = "No podés declarar la guerra a tu mismo clan"
        Exit Function
    End If

    If GIG < 1 Or GIG > CantidadDEGUILDAS Then
        Call LogError("ModGuilds.r_DeclararGuerra: " & GI & " declara a " & GuildGuerra)
        refError = "Inconsistencia en el sistema de guildas. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    Call Guilds(GI).AnularPropuestas(GIG)
    Call Guilds(GIG).AnularPropuestas(GI)
    Call Guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.GUERRA)
    Call Guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.GUERRA)
    
    r_DeclararGuerra = GIG

End Function

Public Function r_AceptarPropuestaDePaz(ByVal UserIndex As Integer, ByRef GuildPaz As String, ByRef refError As String) As Integer
'el guilda de UserIndex acepta la propuesta de paz de guildpaz, con quien esta en guerra
    Dim GI      As Integer
    Dim GIG     As Integer

    GI = UserList(UserIndex).Guild_Id
    If GI < 1 Or GI > CantidadDEGUILDAS Then
        refError = "No eres miembro de ninguna guilda."
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu guilda."
        Exit Function
    End If
    
    If LenB(Trim$(GuildPaz)) < 1 Then
        refError = "No has seleccionado ninguna guilda."
        Exit Function
    End If

    GIG = Guild_Id(GuildPaz)
    
    If GIG < 1 Or GIG > CantidadDEGUILDAS Then
        Call LogError("ModGuilds.r_AceptarPropuestaDePaz: " & GI & " acepta de " & GuildPaz)
        refError = "Inconsistencia en el sistema de guildas. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    If Guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.GUERRA Then
        refError = "No estás en guerra con esa guilda."
        Exit Function
    End If
    
    If Not Guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
        refError = "No hay ninguna propuesta de paz para aceptar."
        Exit Function
    End If

    Call Guilds(GI).AnularPropuestas(GIG)
    Call Guilds(GIG).AnularPropuestas(GI)
    Call Guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.PAZ)
    Call Guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.PAZ)
    
    r_AceptarPropuestaDePaz = GIG
End Function

Public Function r_RechazarPropuestaDeAlianza(ByVal UserIndex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
'devuelve el Index al guilda guildPro
Dim GI      As Integer
Dim GIG     As Integer

    r_RechazarPropuestaDeAlianza = 0
    GI = UserList(UserIndex).Guild_Id
    
    If GI < 1 Or GI > CantidadDEGUILDAS Then
        refError = "No eres miembro de ninguna guilda."
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu guilda."
        Exit Function
    End If
    
    If LenB(Trim$(GuildPro)) < 1 Then
        refError = "No has seleccionado ninguna guilda."
        Exit Function
    End If

    GIG = Guild_Id(GuildPro)
    
    If GIG < 1 Or GIG > CantidadDEGUILDAS Then
        Call LogError("ModGuilds.r_RechazarPropuestaDeAlianza: " & GI & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de guildas. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If
    
    If Not Guilds(GI).HayPropuesta(GIG, ALIADOS) Then
        refError = "No hay propuesta de alianza del guilda " & GuildPro
        Exit Function
    End If
    
    Call Guilds(GI).AnularPropuestas(GIG)
    'avisamos al otra guilda
    Call Guilds(GIG).SetGuildNews(Guilds(GI).GuildName & " ha rechazado nuestra propuesta de alianza. " & Guilds(GIG).GetGuildNews())
    r_RechazarPropuestaDeAlianza = GIG

End Function


Public Function r_RechazarPropuestaDePaz(ByVal UserIndex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
'devuelve el Index al guilda guildPro
Dim GI      As Integer
Dim GIG     As Integer

    r_RechazarPropuestaDePaz = 0
    GI = UserList(UserIndex).Guild_Id
    
    If GI < 1 Or GI > CantidadDEGUILDAS Then
        refError = "No eres miembro de ninguna guilda."
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu guilda."
        Exit Function
    End If
    
    If LenB(Trim$(GuildPro)) < 1 Then
        refError = "No has seleccionado ninguna guilda."
        Exit Function
    End If

    GIG = Guild_Id(GuildPro)
    
    If GIG < 1 Or GIG > CantidadDEGUILDAS Then
        Call LogError("ModGuilds.r_RechazarPropuestaDePaz: " & GI & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de guildas. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If
    
    If Not Guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
        refError = "No hay propuesta de paz del guilda " & GuildPro
        Exit Function
    End If
    
    Call Guilds(GI).AnularPropuestas(GIG)
    'avisamos al otra guilda
    Call Guilds(GIG).SetGuildNews(Guilds(GI).GuildName & " ha rechazado nuestra propuesta de paz. " & Guilds(GIG).GetGuildNews())
    r_RechazarPropuestaDePaz = GIG

End Function


Public Function r_AceptarPropuestaDeAlianza(ByVal UserIndex As Integer, ByRef GuildAllie As String, ByRef refError As String) As Integer
'el guilda de UserIndex acepta la propuesta de paz de guildpaz, con quien esta en guerra
Dim GI      As Integer
Dim GIG     As Integer

    r_AceptarPropuestaDeAlianza = 0
    GI = UserList(UserIndex).Guild_Id
    If GI < 1 Or GI > CantidadDEGUILDAS Then
        refError = "No eres miembro de ninguna guilda."
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu guilda."
        Exit Function
    End If
    
    If LenB(Trim$(GuildAllie)) < 1 Then
        refError = "No has seleccionado ninguna guilda."
        Exit Function
    End If

    GIG = Guild_Id(GuildAllie)
    
    If GIG < 1 Or GIG > CantidadDEGUILDAS Then
        Call LogError("ModGuilds.r_AceptarPropuestaDeAlianza: " & GI & " acepta de " & GuildAllie)
        refError = "Inconsistencia en el sistema de guildas. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    If Guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.PAZ Then
        refError = "No estás en paz con la guilda, solo puedes aceptar propuesas de alianzas con alguien que estes en paz."
        Exit Function
    End If
    
    If Not Guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.ALIADOS) Then
        refError = "No hay ninguna propuesta de alianza para aceptar."
        Exit Function
    End If

    Call Guilds(GI).AnularPropuestas(GIG)
    Call Guilds(GIG).AnularPropuestas(GI)
    Call Guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.ALIADOS)
    Call Guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.ALIADOS)
    
    r_AceptarPropuestaDeAlianza = GIG

End Function

Public Function r_ClanGeneraPropuesta(ByVal UserIndex As Integer, ByRef OtraGuilda As String, ByVal Tipo As RELACIONES_GUILD, ByRef Detalle As String, ByRef refError As String) As Boolean
    
    Dim OtraGuildaGI      As Integer
    Dim GI              As Integer

    r_ClanGeneraPropuesta = False
    
    GI = UserList(UserIndex).Guild_Id
    If GI < 1 Or GI > CantidadDEGUILDAS Then
        refError = "No eres miembro de ninguna guilda."
        Exit Function
    End If
    
    OtraGuildaGI = Guild_Id(OtraGuilda)
    
    If OtraGuildaGI = GI Then
        refError = "No podés declarar relaciones con tu propia guilda."
        Exit Function
    End If
    
    If OtraGuildaGI < 1 Or OtraGuildaGI > CantidadDEGUILDAS Then
        refError = "El sistema de guildas esta inconsistente, el otro guilda no existe!"
        Exit Function
    End If
    
    If Guilds(OtraGuildaGI).HayPropuesta(GI, Tipo) Then
        refError = "Ya hay propuesta de " & Relacion2String(Tipo) & " con " & OtraGuilda
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu guilda."
        Exit Function
    End If
    
    'de acuerdo al tipo procedemos validando las transiciones
    If Tipo = RELACIONES_GUILD.PAZ Then
        If Guilds(GI).GetRelacion(OtraGuildaGI) <> RELACIONES_GUILD.GUERRA Then
            refError = "No estás en guerra con " & OtraGuilda
            Exit Function
        End If
    ElseIf Tipo = RELACIONES_GUILD.GUERRA Then
        'por ahora no hay propuestas de guerra
    ElseIf Tipo = RELACIONES_GUILD.ALIADOS Then
        If Guilds(GI).GetRelacion(OtraGuildaGI) <> RELACIONES_GUILD.PAZ Then
            refError = "Para solicitar alianza no debes estar ni aliado ni en guerra con " & OtraGuilda
            Exit Function
        End If
    End If
    
    Call Guilds(OtraGuildaGI).SetPropuesta(Tipo, GI, Detalle)
    r_ClanGeneraPropuesta = True

End Function

Public Function r_VerPropuesta(ByVal UserIndex As Integer, ByRef OtroGuild As String, ByVal Tipo As RELACIONES_GUILD, ByRef refError As String) As String
Dim OtraGuildaGI      As Integer
Dim GI              As Integer
    
    r_VerPropuesta = vbNullString
    refError = vbNullString
    
    GI = UserList(UserIndex).Guild_Id
    If GI < 1 Or GI > CantidadDEGUILDAS Then
        refError = "No eres miembro de ninguna guilda."
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu guilda."
        Exit Function
    End If
    
    OtraGuildaGI = Guild_Id(OtroGuild)
    
    If Not Guilds(GI).HayPropuesta(OtraGuildaGI, Tipo) Then
        refError = "No existe la propuesta solicitada"
        Exit Function
    End If
    
    r_VerPropuesta = Guilds(GI).GetPropuesta(OtraGuildaGI, Tipo)
    
End Function

Public Function r_ListaDePropuestas(ByVal UserIndex As Integer, ByVal Tipo As RELACIONES_GUILD) As String()

    Dim GI  As Integer
    Dim i   As Integer
    Dim proposalCount As Integer
    Dim proposals() As String
    
    GI = UserList(UserIndex).Guild_Id
    
    If GI > 0 And GI <= CantidadDEGUILDAS Then
        With Guilds(GI)
            proposalCount = .CantidadPropuestas(Tipo)
            
            'Resize array to contain all proposals
            If proposalCount > 0 Then
                ReDim proposals(proposalCount - 1) As String
            Else
                ReDim proposals(0) As String
            End If
            
            'Store each guild name
            For i = 0 To proposalCount - 1
                proposals(i) = Guilds(.Iterador_ProximaPropuesta(Tipo)).GuildName
            Next i
        End With
    End If
    
    r_ListaDePropuestas = proposals
End Function

Public Sub a_RechazarAspiranteChar(ByRef Aspirante As String, ByVal Guild As Integer, ByRef Detalles As String)
    If InStrB(Aspirante, "/") > 0 Then
        Aspirante = Replace(Aspirante, "/", vbNullString)
    End If
    If InStrB(Aspirante, "/") > 0 Then
        Aspirante = Replace(Aspirante, "/", vbNullString)
    End If
    If InStrB(Aspirante, ".") > 0 Then
        Aspirante = Replace(Aspirante, ".", vbNullString)
    End If
    Call Guilds(Guild).InformarRechazoEnChar(Aspirante, Detalles)
End Sub

Public Function a_ObtenerRechazoDeChar(ByRef Aspirante As String) As String
    If InStrB(Aspirante, "/") > 0 Then
        Aspirante = Replace(Aspirante, "/", vbNullString)
    End If
    If InStrB(Aspirante, "/") > 0 Then
        Aspirante = Replace(Aspirante, "/", vbNullString)
    End If
    If InStrB(Aspirante, ".") > 0 Then
        Aspirante = Replace(Aspirante, ".", vbNullString)
    End If
    a_ObtenerRechazoDeChar = GetVar(CharPath & Aspirante & ".chr", "GUILD", "MotivoRechazo")
    Call WriteVar(CharPath & Aspirante & ".chr", "GUILD", "MotivoRechazo", vbNullString)
End Function

Public Function a_RechazarAspirante(ByVal UserIndex As Integer, ByRef Nombre As String, ByRef refError As String) As Boolean
Dim GI              As Integer
Dim NroAspirante    As Integer

    a_RechazarAspirante = False
    GI = UserList(UserIndex).Guild_Id
    If GI < 1 Or GI > CantidadDEGUILDAS Then
        refError = "No perteneces a ninguna guilda."
        Exit Function
    End If

    NroAspirante = Guilds(GI).NumeroDeAspirante(Nombre)

    If NroAspirante = 0 Then
        refError = Nombre & " no es aspirante a tu guilda."
        Exit Function
    End If

    Call Guilds(GI).RetirarAspirante(Nombre, NroAspirante)
    refError = "Fue rechazada tu solicitud de ingreso a " & Guilds(GI).GuildName
    a_RechazarAspirante = True

End Function

Public Function a_DetallesAspirante(ByVal UserIndex As Integer, ByRef Nombre As String) As String
Dim GI              As Integer
Dim NroAspirante    As Integer

    GI = UserList(UserIndex).Guild_Id
    If GI < 1 Or GI > CantidadDEGUILDAS Then
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        Exit Function
    End If
    
    NroAspirante = Guilds(GI).NumeroDeAspirante(Nombre)
    If NroAspirante > 0 Then
        a_DetallesAspirante = Guilds(GI).DetallesSolicitudAspirante(NroAspirante)
    End If
    
End Function

Public Sub SendDetallesPersonaje(ByVal UserIndex As Integer, ByVal Personaje As String)
    Dim GI          As Integer
    Dim NroAsp      As Integer
    Dim GuildName   As String
    Dim UserFile    As clsIniManager
    Dim Miembro     As String
    Dim GuildActual As Integer
    Dim list()      As String
    Dim i           As Long
    
    GI = UserList(UserIndex).Guild_Id
    
    Personaje = UCase$(Personaje)
    
    If GI < 1 Or GI > CantidadDEGUILDAS Then
        Call Protocol.WriteConsoleMsg(UserIndex, "No perteneces a ninguna guilda.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        Call Protocol.WriteConsoleMsg(UserIndex, "No eres el líder de tu guilda.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If InStrB(Personaje, "/") > 0 Then
        Personaje = Replace$(Personaje, "/", vbNullString)
    End If
    If InStrB(Personaje, "/") > 0 Then
        Personaje = Replace$(Personaje, "/", vbNullString)
    End If
    If InStrB(Personaje, ".") > 0 Then
        Personaje = Replace$(Personaje, ".", vbNullString)
    End If
    
    NroAsp = Guilds(GI).NumeroDeAspirante(Personaje)
    
    If NroAsp = 0 Then
        list = Guilds(GI).GetMemberList()
        
        For i = 0 To UBound(list())
            If Personaje = list(i) Then
                Exit For
            End If
        Next i
        
        If i > UBound(list()) Then
            Call Protocol.WriteConsoleMsg(UserIndex, "El personaje no es ni aspirante ni miembro dla guilda", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    'ahora traemos la info
    
    Set UserFile = New clsIniManager
    
    With UserFile
        .Initialize (CharPath & Personaje & ".chr")
        
        'Get the Char's current guild
        GuildActual = Val(.GetValue("GUILD", "Guild_Id"))
        If GuildActual > 0 And GuildActual <= CantidadDEGUILDAS Then
            GuildName = "<" & Guilds(GuildActual).GuildName & ">"
        Else
            GuildName = "Ninguno"
        End If
        
        'Get previous guilds
        Miembro = .GetValue("GUILD", "Miembro")
        If Len(Miembro) > 400 Then
            Miembro = ".." & Right$(Miembro, 400)
        End If
        
        Call Protocol.WriteCharInfo(UserIndex, Personaje, .GetValue("INIT", "Raza"), .GetValue("INIT", "Clase"), _
                                .GetValue("INIT", "Genero"), .GetValue("STATS", "ELV"), .GetValue("STATS", "GLD"), _
                                .GetValue("STATS", "Banco"), .GetValue("GUILD", "Pedidos"), _
                                GuildName, Miembro, .GetValue("STATS", "Matados"), .GetValue("STATS", "Muertes"))
    End With
    
    Set UserFile = Nothing
End Sub

Public Function a_NuevoAspirante(ByVal UserIndex As Integer, ByRef Guilda As String, ByRef Solicitud As String, ByRef refError As String) As Boolean
    
    Dim ViejoSolicitado     As String
    Dim ViejoGuild_Id     As Integer
    Dim ViejoNroAspirante   As Integer
    Dim NuevoGuild_Id     As Integer

    a_NuevoAspirante = False

    If UserList(UserIndex).Guild_Id > 0 Then
        refError = "Ya perteneces a una guilda, debes salir del mismo antes de solicitar ingresar a otra."
        Exit Function
    End If
    
    If EsPrincipiante(UserIndex) Then
        refError = "Los principiantes no tienen derecho a entrar a una guilda."
        Exit Function
    End If

    NuevoGuild_Id = Guild_Id(Guilda)
    If NuevoGuild_Id = 0 Then
        refError = "Esa guilda no existe! Avise a un administrador."
        Exit Function
    End If

    If Guilds(NuevoGuild_Id).CantidadAspirantes >= MaxASPIRANTES Then
        refError = "El guilda tiene demasiados aspirantes. Contáctate con un miembro para que procese las solicitudes."
        Exit Function
    End If

    ViejoSolicitado = GetVar(CharPath & UserList(UserIndex).Name & ".chr", "GUILD", "ASPIRANTEA")

    If LenB(ViejoSolicitado) > 0 Then
        'borramos la vieja solicitud
        ViejoGuild_Id = CInt(ViejoSolicitado)
        If ViejoGuild_Id > 0 Then
            ViejoNroAspirante = Guilds(ViejoGuild_Id).NumeroDeAspirante(UserList(UserIndex).Name)
            If ViejoNroAspirante > 0 Then
                Call Guilds(ViejoGuild_Id).RetirarAspirante(UserList(UserIndex).Name, ViejoNroAspirante)
            End If
        Else
            'RefError = "Inconsistencia en los guildas, avise a un administrador"
            'EXIT FUNCTION
        End If
    End If
    
    Call Guilds(NuevoGuild_Id).NuevoAspirante(UserList(UserIndex).Name, Solicitud)
    a_NuevoAspirante = True
End Function

Public Function a_AceptarAspirante(ByVal UserIndex As Integer, ByRef Aspirante As String, ByRef refError As String) As Boolean
Dim GI              As Integer
Dim NroAspirante    As Integer
Dim AspiranteUI     As Integer

    'un pj ingresa al guilda :D

    a_AceptarAspirante = False
    
    GI = UserList(UserIndex).Guild_Id
    If GI < 1 Or GI > CantidadDEGUILDAS Then
        refError = "No perteneces a ninguna guilda."
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu guilda."
        Exit Function
    End If
    
    NroAspirante = Guilds(GI).NumeroDeAspirante(Aspirante)
    
    If NroAspirante = 0 Then
        refError = "El Pj no es aspirante a la guilda"
        Exit Function
    End If
    
    AspiranteUI = NameIndex(Aspirante)
    If AspiranteUI > 0 Then
        'pj Online
        If Not UserList(AspiranteUI).Guild_Id = 0 Then
            refError = Aspirante & " ya es perteneces a otra guilda."
            Call Guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        End If
    Else
        If GetGuild_IdFromChar(Aspirante) Then
            refError = Aspirante & " ya pertenece a otra guilda."
            Call Guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        End If
    End If
    'el pj es aspirante al guilda y puede entrar
    
    Call Guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
    Call Guilds(GI).AceptarNuevoMiembro(Aspirante)
    
    'If player is online, update tag
    If AspiranteUI > 0 Then
        Call RefreshCharStatus(AspiranteUI)
    End If
    
    a_AceptarAspirante = True
End Function

Public Function GuildName(ByVal Guild_Id As Integer) As String
    If Guild_Id < 1 Or Guild_Id > CantidadDEGUILDAS Then _
        Exit Function
    
    GuildName = Guilds(Guild_Id).GuildName
End Function

Public Function GuildLeader(ByVal Guild_Id As Integer) As String
    If Guild_Id < 1 Or Guild_Id > CantidadDEGUILDAS Then _
        Exit Function
    
    GuildLeader = Guilds(Guild_Id).GetLeader
End Function

Public Function GuildFounder(ByVal Guild_Id As Integer) As String
    If Guild_Id < 1 Or Guild_Id > CantidadDEGUILDAS Then _
        Exit Function
    
    GuildFounder = Guilds(Guild_Id).Fundador
End Function
