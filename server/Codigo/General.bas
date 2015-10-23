Attribute VB_Name = "General"
Option Explicit

Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Const RunHighPriority As Boolean = True

Global LeerNpcs As New clsIniManager

Public Sub DarImagen(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        .Char.Heading = eHeading.SOUTH

        .Char.Head = .OrigChar.Head
            
        If .Inv.Head > 0 Then
            .Char.HeadAnim = ObjData(.Inv.Head).HeadAnim
        Else
            .Char.HeadAnim = NingunCasco
        End If
    
        If .Inv.Body > 0 Then
            .Char.Body = ObjData(.Inv.Body).BodyAnim
            .flags.Desnudo = False
        Else
            Call DarCuerpoDesnudo(UserIndex)
        End If

        .Char.WeaponAnim = GetWeaponAnim(UserIndex)

        If UsaEscudo(UserIndex) > 0 Then
            .Char.ShieldAnim = ObjData(.Inv.LeftHand).ShieldAnim
        Else
            .Char.ShieldAnim = NingunEscudo
        End If
                
        If .Inv.Ship > 0 Then
            If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then
                Call ToogleBoatBody(UserIndex)
                .flags.Navegando = True
                Call WriteNavigateToggle(UserIndex)
            Else
                .Inv.Ship = 0
                .flags.Navegando = False
            End If
        End If
    
    End With
End Sub

Public Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, Optional ByVal Mimetizado As Boolean = False)
'Da cuerpo desnudo a un usuario
    
    Dim CuerpoDesnudo As Integer
    
    With UserList(UserIndex)
    
        If .Genero = eGenero.Hombre Then

            Select Case .Raza
            
                Case eRaza.Humano
                    CuerpoDesnudo = 21
                Case eRaza.Drow
                    CuerpoDesnudo = 32
                Case eRaza.Elfo
                    CuerpoDesnudo = 210
                Case eRaza.Gnomo
                    CuerpoDesnudo = 222
                Case eRaza.Enano
                    CuerpoDesnudo = 53
            End Select
        Else
            Select Case .Raza
                Case eRaza.Humano
                    CuerpoDesnudo = 39
                Case eRaza.Drow
                    CuerpoDesnudo = 40
                Case eRaza.Elfo
                    CuerpoDesnudo = 259
                Case eRaza.Gnomo
                    CuerpoDesnudo = 260
                Case eRaza.Enano
                    CuerpoDesnudo = 60
            End Select
        End If
    
        If Mimetizado Then
            .CharMimetizado.Body = CuerpoDesnudo
        Else
            .Char.Body = CuerpoDesnudo
            
            If .Char.Head < 1 Then
                .Char.Head = .OrigChar.Head
            End If
        End If
    
        .flags.Desnudo = True
    End With

End Sub

Public Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal b As Boolean)
    'b ahora es boolean,
    'b=true bloquea el tile en (x,y)
    'b=false desbloquea el tile en (x,y)
    'toMap = true -> Envia los datos a todo el mapa
    'toMap = false -> Envia los datos al user
    'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
    'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s
    
    If toMap Then
        Call SendData(SendTarget.toMap, sndIndex, Msg_BlockPosition(X, Y, b))
    Else
        Call WriteBlockPosition(sndIndex, X, Y, b)
    End If

End Sub

Public Function HayAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    
    If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        With MapData(X, Y)
            If ((.Graphic(1) >= 1505 And .Graphic(1) <= 1520) Or _
            (.Graphic(1) >= 5665 And .Graphic(1) <= 5680) Or _
            (.Graphic(1) >= 13547 And .Graphic(1) <= 13562)) And _
               .Graphic(2) = 0 Then
                HayAgua = True
            Else
                HayAgua = False
        End If
        End With
    Else
      HayAgua = False
    End If

End Function

Private Function HayLava(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    
    If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        If MapData(X, Y).Graphic(1) >= 5837 And MapData(X, Y).Graphic(1) <= 5852 Then
            HayLava = True
        Else
            HayLava = False
        End If
    Else
        HayLava = False
    End If
    
End Function

Public Sub LimpiarMundo()

On Error Resume Next
    
    'Dim i As Integer
    'Dim d As New cGarbage
    
    'For i = TrashCollector.Count To 1 Step -1
    '    Set d = TrashCollector(i)
    '    Call EraseObj(d.Map, d.X, d.Y, -1)
    '    Call TrashCollector.Remove(i)
    '    Set d = Nothing
    'Next i

    Dim Map As Integer
    Dim X As Integer
    Dim Y As Integer
    
    For Map = 1 To NumMaps
        For Y = YMinMapSize To YMaxMapSize
            For X = XMinMapSize To XMaxMapSize
                If MapData(X, Y).UserIndex < 1 Then
                    If Not MapData(X, Y).Blocked Then
                        If MapData(X, Y).ObjInfo.index > 0 And MapData(X, Y).ObjInfo.index < NumObjDatas Then
                            If Not EsObjetoFijo(ObjData(MapData(X, Y).ObjInfo.index).Type) Then
                                Call EraseObj(Map, X, Y, -1)
                            End If
                        End If
                    End If
                End If
            Next X
        Next Y
    Next Map

End Sub

Public Sub EnviarSpawnList(ByVal UserIndex As Integer)

    Dim k As Long
    Dim npcNames() As String
    
    ReDim npcNames(1 To UBound(Spawn_List)) As String
    
    For k = 1 To UBound(Spawn_List)
        npcNames(k) = Spawn_List(k).NpcName
    Next k
    
    Call WriteSpawnList(UserIndex, npcNames())

End Sub

Public Sub Main()

On Error Resume Next

    Dim PacketKeys() As String
        
    'Directorios
    DatPath = App.Path & "\Dat\"
    MapPath = App.Path & "\Maps\"
    CharPath = App.Path & "\Chars\"

    ServidorIni = App.Path & "\Servidor.ini"
    
    If RunHighPriority Then
        SetThreadPriority GetCurrentThread, 2       'Reccomended you dont touch these values
        SetPriorityClass GetCurrentProcess, &H80    'unless you know what you're doing
    End If
    
    Call MySQL_Init
    
    If OptimizeDatabase Then
        MySQL_Optimize
    End If
    
    Call LoadMotd

    Prision.Map = 66
    Libertad.Map = 66
    
    Prision.X = 75
    Prision.Y = 47
    Libertad.X = 75
    Libertad.Y = 65
    
    LastBackup = Format(Now, "Short Time")
    Minutos = Format(Now, "Short Time")
    
    LevelSkill(1).LevelValue = 3
    LevelSkill(2).LevelValue = 5
    LevelSkill(3).LevelValue = 7
    LevelSkill(4).LevelValue = 10
    LevelSkill(5).LevelValue = 13
    LevelSkill(6).LevelValue = 15
    LevelSkill(7).LevelValue = 17
    LevelSkill(8).LevelValue = 20
    LevelSkill(9).LevelValue = 23
    LevelSkill(10).LevelValue = 25
    LevelSkill(11).LevelValue = 27
    LevelSkill(12).LevelValue = 30
    LevelSkill(13).LevelValue = 33
    LevelSkill(14).LevelValue = 35
    LevelSkill(15).LevelValue = 37
    LevelSkill(16).LevelValue = 40
    LevelSkill(17).LevelValue = 43
    LevelSkill(18).LevelValue = 45
    LevelSkill(19).LevelValue = 47
    LevelSkill(20).LevelValue = 50
    LevelSkill(21).LevelValue = 53
    LevelSkill(22).LevelValue = 55
    LevelSkill(23).LevelValue = 57
    LevelSkill(24).LevelValue = 60
    LevelSkill(25).LevelValue = 63
    LevelSkill(26).LevelValue = 65
    LevelSkill(27).LevelValue = 67
    LevelSkill(28).LevelValue = 70
    LevelSkill(29).LevelValue = 73
    LevelSkill(30).LevelValue = 75
    LevelSkill(31).LevelValue = 77
    LevelSkill(32).LevelValue = 80
    LevelSkill(33).LevelValue = 83
    LevelSkill(34).LevelValue = 85
    LevelSkill(35).LevelValue = 87
    LevelSkill(36).LevelValue = 90
    LevelSkill(37).LevelValue = 93
    LevelSkill(38).LevelValue = 95
    LevelSkill(39).LevelValue = 97
    LevelSkill(40).LevelValue = 100
    LevelSkill(41).LevelValue = 100
    LevelSkill(42).LevelValue = 100
    LevelSkill(43).LevelValue = 100
    LevelSkill(44).LevelValue = 100
    LevelSkill(45).LevelValue = 100
    LevelSkill(46).LevelValue = 100
    LevelSkill(47).LevelValue = 100
    LevelSkill(48).LevelValue = 100
    LevelSkill(49).LevelValue = 100
    LevelSkill(50).LevelValue = 100
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.Drow) = "Drow"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    
    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Pirat) = "Pirata"
    
    SkillName(eSkill.Magia) = "Magia"
    SkillName(eSkill.Robar) = "Robar"
    SkillName(eSkill.Tacticas) = "Evasion en combate"
    SkillName(eSkill.Armas) = "Combate con armas"
    SkillName(eSkill.Meditar) = "Meditar"
    SkillName(eSkill.Apuñalar) = "Apuñalar"
    SkillName(eSkill.Ocultarse) = "Ocultarse"
    SkillName(eSkill.Supervivencia) = "Supervivencia"
    SkillName(eSkill.Talar) = "Talar"
    SkillName(eSkill.Defensa) = "Defensa con escudos"
    SkillName(eSkill.Pesca) = "Pesca"
    SkillName(eSkill.Mineria) = "Mineria"
    SkillName(eSkill.Carpinteria) = "Carpinteria"
    SkillName(eSkill.Herreria) = "Herreria"
    SkillName(eSkill.Liderazgo) = "Liderazgo"
    SkillName(eSkill.Domar) = "Domar animales"
    SkillName(eSkill.Proyectiles) = "Combate a distancia"
    SkillName(eSkill.Wrestling) = "Combate sin armas"
    SkillName(eSkill.Navegacion) = "Navegacion"
    
    ListaAtributos(eAtributos.Fuerza) = "Fuerza"
    ListaAtributos(eAtributos.Agilidad) = "Agilidad"
    ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
    ListaAtributos(eAtributos.Carisma) = "Carisma"
    ListaAtributos(eAtributos.Constitucion) = "Constitucion"

    frmCargando.Show
        
    frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    
    'Bordes del mapa
    MinXBorder = 1 'XMinMapSize + (XWindow * 0.5)
    MaxXBorder = 100 * 5 'XMaxMapSize - (XWindow * 0.5)
    MinYBorder = 1 'YMinMapSize + (YWindow * 0.5)
    MaxYBorder = 100 * 5 'YMaxMapSize - (YWindow * 0.5)
    DoEvents

    Call LoadGuildsDB
    
    Call CargarSpawnList
    Call CargarForbidenWords
    Call LoadSini
    Call CargaApuestas
    Call CargaNpcsDat
    Call LoadOBJData
    Call CargarHechizos
    Call LoadArmasHerreria
    Call LoadArmadurasHerreria
    Call LoadObjCarpintero
    Call LoadBalance
    Call LoadPlataformas
    
    Call DB_RS_Open("SELECT * FROM people WHERE `logged` = 1")
    
    If Not DB_RS.EOF Then
        DB_RS!Logged = 0
        DB_RS.Update
    End If
    
    DB_RS_Close
    
    If BootDelBackUp Then
        frmCargando.Label1(2).Caption = "Cargando BackUp"
        Call CargarBackUp
    Else
        frmCargando.Label1(2).Caption = "Cargando Mapas"
        Call LoadMapData
    End If
    
    Call SonidosMapas.LoadSoundMapInfo

    Dim LoopC As Integer
    
    'Resetea las conexiones de los usuarios
    For LoopC = 1 To MaxPoblacion
        With UserList(LoopC)
           .ConnID = -1
           .ConnIDValida = False
            Set .incomingData = New clsByteQueue
            Set .outgoingData = New clsByteQueue
        End With
    Next LoopC

    frmServidor.Visible = True
    
    With frmMain
        .SaveTimer.Enabled = True
        .tLluvia.Enabled = True
        .tPiqueteC.Enabled = True
        .GameTimer.Enabled = True
        .tLluviaEvent.Enabled = True
        .FX.Enabled = True
        .Auditoria.Enabled = True
        .KillLog.Enabled = True
        .TIMER_AI.Enabled = True
        .TIMER_PET_AI.Enabled = True
        .NpcAtaca.Enabled = True
    End With
    
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    'Configuracion de los sockets
    
    Call IniciaWsApi(frmMain.hWnd)
    SockListen = ListenForConnect(Puerto, hWndMsg, vbNullString)
    
    If frmMain.Visible Then
        frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    End If
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    Unload frmCargando
    
    'Log
    Dim N As Integer
    N = FreeFile
    Open App.Path & "/logs/main.log" For Append Shared As #N
    Print #N, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
    Close #N
    
    frmMain.Show

    'Ocultar
    'If HideMe = 1 Then
    '    Call frmMain.InitMain(1)
    'Else
    '    Call frmMain.InitMain(1)
    'End If
    
    tInicioServer = GetTickCount() And &H7FFFFFFF
End Sub

Public Function FileExist(ByVal File As String, Optional FileNpcType As VbFileAttribute = vbNormal) As Boolean
'***
'Se fija si existe el archivo
'***
    FileExist = LenB(dir$(File, FileNpcType)) > 0
End Function

Public Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'***
'Gets a field from a string

'Last Modify Date: 11\15\2004
'Gets a field from a delimited string
'***
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Public Function MapaValido(ByVal Map As Integer) As Boolean
    MapaValido = Map > 0 And Map <= NumMaps
End Function

Public Sub LogCriticEvent(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile 'obtenemos un canal
Open App.Path & "/logs\Eventos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogIndex(ByVal index As Integer, ByVal Desc As String)
On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile 'obtenemos un canal
    Open App.Path & "/logs/" & index & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile
    
    Exit Sub

errhandler:

End Sub

Public Sub LogError(Desc As String)
On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile 'obtenemos un canal
    Open App.Path & "/logs/errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile
    
    Exit Sub

errhandler:

End Sub

Public Sub LogStatic(Desc As String)
On Error GoTo errhandler
    
    Dim nfile As Integer
    nfile = FreeFile 'obtenemos un canal
    Open App.Path & "/logs/stats.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile
    
    Exit Sub

errhandler:

End Sub

Public Sub LogTarea(Desc As String)
On Error GoTo errhandler
    
    Dim nfile As Integer
    nfile = FreeFile(1) 'obtenemos un canal
    Open App.Path & "/logs/haciendo.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile
    
    Exit Sub

errhandler:

End Sub

Public Sub LogClanes(ByVal str As String)
    
    Dim nfile As Integer
    nfile = FreeFile 'obtenemos un canal
    Open App.Path & "/logs/guildas.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub

Public Sub LogIP(ByVal str As String)

    Dim nfile As Integer
    nfile = FreeFile 'obtenemos un canal
    Open App.Path & "/logs/iP.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub

Public Sub LogDesarrollo(ByVal str As String)
    
    Dim nfile As Integer
    nfile = FreeFile 'obtenemos un canal
    Open App.Path & "/logs/desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile
    
End Sub

Public Sub LogGM(Nombre As String, texto As String)
On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile 'obtenemos un canal
    'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18\05\07
    Open App.Path & "/logs/" & Nombre & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile
    
Exit Sub

errhandler:

End Sub

Public Sub LogAsesinato(texto As String)
On Error GoTo errhandler

    Dim nfile As Integer
    
    nfile = FreeFile 'obtenemos un canal
    
    Open App.Path & "/logs/asesinatos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile
    
    Exit Sub

errhandler:

End Sub
Public Sub logVentaCasa(ByVal texto As String)
On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile 'obtenemos un canal
    
    Open App.Path & "/logs/propiedades.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub
    
errhandler:


End Sub
Public Sub LogHackAttemp(texto As String)
On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile 'obtenemos un canal
    Open App.Path & "/logs/hackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

errhandler:

End Sub

Public Sub LogCheating(texto As String)
On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile 'obtenemos un canal
    Open App.Path & "/logs/cH.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile
    
    Exit Sub
    
errhandler:

End Sub

Public Sub LogCriticalHackAttemp(texto As String)
On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile 'obtenemos un canal
    Open App.Path & "/logs/criticalHackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub
    
errhandler:

End Sub

Public Sub LogAntiCheat(texto As String)
On Error GoTo errhandler
    
    Dim nfile As Integer
    nfile = FreeFile 'obtenemos un canal
    Open App.Path & "/logs/antiCheat.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, vbNullString
    Close #nfile
    
    Exit Sub

errhandler:

End Sub

Public Function ValidInputNP(ByVal cad As String) As Boolean

    Dim Arg As String
    Dim i As Integer
    
    
    For i = 1 To 33
    
    Arg = ReadField(i, cad, 44)
    
    If LenB(Arg) = 0 Then
        Exit Function
    End If
    
    Next i
    
    ValidInputNP = True

End Function

Public Sub Restart()
    
    'Se asegura de que los sockets estan cerrados e ignora cualquier err
    On Error Resume Next
    
    If frmMain.Visible Then
        frmMain.txStatus.Caption = "Reiniciando."
    End If
    
    Dim LoopC As Long
    
    'Cierra el socket de escucha
    If SockListen >= 0 Then
        Call apiclosesocket(SockListen)
    End If
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, vbNullString)
    
    For LoopC = 1 To MaxPoblacion
        Call CloseSocket(LoopC)
    Next
    
    'Initialize statistics!!
    Call Statistics.Initialize
    
    For LoopC = 1 To UBound(UserList())
        Set UserList(LoopC).incomingData = Nothing
        Set UserList(LoopC).outgoingData = Nothing
    Next LoopC
    
    ReDim UserList(1 To MaxPoblacion) As User
    
    'Resetea las conexiones de los usuarios
    For LoopC = 1 To MaxPoblacion
        With UserList(LoopC)
           .ConnID = -1
           .ConnIDValida = False
            Set .incomingData = New clsByteQueue
            Set .outgoingData = New clsByteQueue
        End With
    Next LoopC
    
    LastUser = 0
    
    Poblacion = 0
    frmMain.Poblacion.Caption = "Población: " & Poblacion
    Call Base.OnlinePlayers
    
    Call FreeNpcs
    Call FreeCharIndexes
    
    Call LoadSini
    Call LoadOBJData
    
    Call LoadMapData
    
    Call CargarHechizos
    
    If frmMain.Visible Then
        frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    End If
    
    'Log it
    Dim N As Integer
    N = FreeFile
    Open App.Path & "/logs/main.log" For Append Shared As #N
    Print #N, Date & " " & Time & " servidor reiniciado."
    Close #N
    
    'Ocultar
    'If HideMe = 1 Then
    '    Call frmMain.InitMain(1)
    'Else
    '    Call frmMain.InitMain(0)
    'End If

End Sub


Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
    
    With UserList(UserIndex)
        If MapInfo(.Pos.Map).Zona <> Dungeon Then
            If MapData(.Pos.X, .Pos.Y).Trigger <> 1 And _
               MapData(.Pos.X, .Pos.Y).Trigger <> 4 Then
                Intemperie = True
            End If
        Else
            Intemperie = False
        End If
    End With
    
    'En las arenas no te afecta la lluvia
    If IsArena(UserIndex) Then
        Intemperie = False
    End If
End Function

Public Sub EfectoLluvia(ByVal UserIndex As Integer)

    On Error GoTo errhandler
    
    If UserList(UserIndex).flags.Logged Then
        If Intemperie(UserIndex) Then
                    Dim modifi As Long
                    modifi = Porcentaje(UserList(UserIndex).Stats.MaxSta, 3)
                    Call QuitarSta(UserIndex, modifi)
                    Call FlushBuffer(UserIndex)
        End If
    End If
    
    Exit Sub
errhandler:
 LogError ("Error en EfectoLluvia")
End Sub

Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
    Dim i As Integer
    
    For i = 1 To MaxPets
        With UserList(UserIndex)
            If .Pets.Pet(i).index > 0 Then
            
                Dim Tiempo As Integer
                
                Tiempo = NpcList(.Pets.Pet(i).index).Contadores.TiempoExistencia
                
                If Tiempo > 0 Then
                    
                    If Tiempo = Max_Integer_Value Then
                        Exit Sub
                    End If
                    
                    Tiempo = Tiempo - 1
                    
                    If Tiempo < 1 Then
                        Call MuereNpc(.Pets.Pet(i).index)
                    Else
                        NpcList(.Pets.Pet(i).index).Contadores.TiempoExistencia = Tiempo
                    End If
                End If
            End If
        End With
    Next i
End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)
'If user is naked and it's in a cold map, take health points from him

    Dim modifi As Integer
    
    With UserList(UserIndex)
        If .Counters.Frio < IntervaloFrio Then
            .Counters.Frio = .Counters.Frio + 1
    Else
            If MapInfo(.Pos.Map).terreno = Nieve Then
                modifi = Porcentaje(.Stats.MaxHP, 5)
                .Stats.MinHP = .Stats.MinHP - modifi
                    
                If .Stats.MinHP < 1 Then
                    Call WriteConsoleMsg(UserIndex, "Te moriste de frío.", FontTypeNames.FONTTYPE_INFO)
                    Call UserDie(UserIndex)
                Else
                    Call WriteUpdateHP(UserIndex)
                    'Call WriteConsoleMsg(UserIndex, "El frío te dañó en " & modifi & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
                End If
        Else
                modifi = Porcentaje(.Stats.MaxSta, 5)
            Call QuitarSta(UserIndex, modifi)
        End If
        
            .Counters.Frio = 0
    End If
    End With
End Sub

Public Sub EfectoLava(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        If .Counters.Lava < IntervaloFrio Then 'Usamos el mismo intervalo que el del frio
            .Counters.Lava = .Counters.Lava + 1
        Else
            If HayLava(.Pos.Map, .Pos.X, .Pos.Y) Then
                Call WriteConsoleMsg(UserIndex, "Te estás quemando.", FontTypeNames.FONTTYPE_INFO)
                .Stats.MinHP = .Stats.MinHP - Porcentaje(.Stats.MaxHP, 5)
            
                If .Stats.MinHP < 1 Then
                    Call WriteConsoleMsg(UserIndex, "Moriste quemado.", FontTypeNames.FONTTYPE_INFO)
                    Call UserDie(UserIndex)
                Else
                    Call WriteUpdateHP(UserIndex)
                End If
            End If
            .Counters.Lava = 0
        End If
    End With
End Sub

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)
'Maneja el tiempo y el efecto del mimetismo

    Dim Barco As ObjData
    
    With UserList(UserIndex)
        If .Counters.Mimetismo < IntervaloInvisible Then
            .Counters.Mimetismo = .Counters.Mimetismo + 1
        Else
            'restore old char
            Call WriteConsoleMsg(UserIndex, "Recuperaste tu apariencia normal.", FontTypeNames.FONTTYPE_INFO)
            
            If .flags.Navegando Then
                If Not .Stats.Muerto Then
                    Barco = ObjData(UserList(UserIndex).Inv.Ship)

                    If Barco.BodyAnim = iBarca Then
                        .Char.Body = iBarcaCiuda
                    End If
                    
                    If Barco.BodyAnim = iGalera Then
                        .Char.Body = iGaleraCiuda
                    End If
                    
                    If Barco.BodyAnim = iGaleon Then
                        .Char.Body = iGaleonCiuda
                    End If
                Else
                    .Char.Body = iFragataFantasmal
                End If
                
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.HeadAnim = NingunCasco
            Else
                .Char.Body = .CharMimetizado.Body
                .Char.Head = .CharMimetizado.Head
                .Char.HeadAnim = .CharMimetizado.HeadAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            End If
            
            With .Char
                Call ChangeUserChar(UserIndex, .Body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .HeadAnim)
            End With
            
            .Counters.Mimetismo = 0
            .flags.Mimetizado = False
        End If
    End With
End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        If .Counters.Invisibilidad < IntervaloInvisible Then
            .Counters.Invisibilidad = .Counters.Invisibilidad + 1
        Else
            .Counters.Invisibilidad = RandomNumber(-100, 100)
            .flags.Invisible = 0
            If .flags.Oculto < 1 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SetInvisible(.Char.CharIndex, False))
            End If
        End If
    End With

End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
        If .Counters.Ceguera > 0 Then
            .Counters.Ceguera = .Counters.Ceguera - 1
        Else
            If .flags.Ceguera = 1 Then
                    .flags.Ceguera = 0
                Call WriteBlindNoMore(UserIndex)
            End If
            If .flags.Estupidez = 1 Then
                .flags.Estupidez = 0
                Call WriteDumbNoMore(UserIndex)
            End If
        End If
    End With

End Sub

Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .Counters.Paralisis > 0 Then
            .Counters.Paralisis = .Counters.Paralisis - 1
        Else
            .flags.Paralizado = 0
            .flags.Inmovilizado = 0
            '.Flags.AdministrativeParalisis = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SetParalized(.Char.CharIndex, 0))
        End If
    End With
End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, ByVal Intervalo As Integer)

    With UserList(UserIndex)
        'If Maps(.Pos.Map).mapData( .Pos.X, .Pos.Y).Trigger = 1 And _
        '   Maps(.Pos.Map).mapData( .Pos.X, .Pos.Y).Trigger = 4 Then
        '   EXIT SUB
        'End If
        
        If .Stats.MinSta < .Stats.MaxSta Then
            If .Counters.STACounter < Intervalo Then
                .Counters.STACounter = .Counters.STACounter + 1
            Else
                .Counters.STACounter = 0
        
                .Stats.MinSta = .Stats.MinSta + .Stats.MaxSta * 0.01
                
                If .Stats.MinSta > .Stats.MaxSta Then
                    .Stats.MinSta = .Stats.MaxSta
                End If
                
                Call WriteUpdateSta(UserIndex)
            End If
        End If

    End With

End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        If .Counters.Veneno < IntervaloVeneno Then
            .Counters.Veneno = .Counters.Veneno + 1
        Else
            .Counters.Veneno = 0
            
            .Stats.MinHP = .Stats.MinHP - RandomNumber(3, 8) * 0.01 * .Stats.MaxHP
        
            If .Stats.MinHP < 1 Then
                Call UserDie(UserIndex)
                'Call WriteConsoleMsg(UserIndex, "Has muerto a causa del veneno.", FontTypeNames.FONTTYPE_VENENO)
            Else
                Call WriteUpdateHP(UserIndex)
                'Call Writedamage(UserIndex,
                'Call WriteConsoleMsg(UserIndex, "El veneno te daño en " & N & " puntos de vida.", FontTypeNames.FONTTYPE_VENENO)
            End If
            
        End If
    End With
End Sub

Public Sub DuracionPociones(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        'Controla la duracion de las pociones
        If .flags.DuracionEfecto > 0 Then
           .flags.DuracionEfecto = .flags.DuracionEfecto - 1
           If .flags.DuracionEfecto = 0 Then
                .flags.TomoPocion = False
                .flags.TipoPocion = 0
                'volvemos los atributos al estado normal
                Dim loopX As Integer
                
                For loopX = 1 To NUMATRIBUTOS
                    .Stats.Atributos(loopX) = .Stats.AtributosBackUP(loopX)
                Next loopX
                
                Call WriteUpdateStrenghtAndDexterity(UserIndex)
           End If
        End If
    End With

End Sub

Public Sub HambreYSed(ByVal UserIndex As Integer)

    Dim EnviarHyS As Boolean
    
    With UserList(UserIndex)
        If Not .flags.Privilegios And PlayerType.User Then
            Exit Sub
        End If
        
        'Sed
        If .Stats.MinSed > 0 Then
            If .Counters.AGUACounter < IntervaloSed Then
                .Counters.AGUACounter = .Counters.AGUACounter + 1
            Else
                .Counters.AGUACounter = 0
                .Stats.MinSed = .Stats.MinSed - 1
                
                If .Stats.MinSed < 1 Then
                    .Stats.MinSed = 0
                End If
            
                EnviarHyS = True
            End If
        End If
    
        'Hambre
        If .Stats.MinHam > 0 Then
            If .Counters.COMCounter < IntervaloHambre Then
                .Counters.COMCounter = .Counters.COMCounter + 1
            Else
                .Counters.COMCounter = 0
                .Stats.MinHam = .Stats.MinHam - 1
                
                If .Stats.MinHam < 1 Then
                    .Stats.MinHam = 0
                End If
                
                EnviarHyS = True
            End If
        End If
    End With
    
    If EnviarHyS Then
        Call WriteUpdateHungerAndThirst(UserIndex)
    End If
    
End Sub

Public Sub Sanar(ByVal UserIndex As Integer, ByVal Intervalo As Integer)

    With UserList(UserIndex)
        If MapData(.Pos.X, .Pos.Y).Trigger = 1 And _
            MapData(.Pos.X, .Pos.Y).Trigger = 4 Then
            Exit Sub
        End If
    
        Dim HpSana As Integer
        
        If .Stats.MinHP < .Stats.MaxHP Then
            If .Counters.HPCounter < Intervalo Then
                .Counters.HPCounter = .Counters.HPCounter + 1
            Else
                If .Stats.Atributos(eAtributos.Constitucion) - 18 > 0 Then
                    HpSana = Porcentaje(.Stats.MaxHP, 1) + 1 * (.Stats.Atributos(eAtributos.Constitucion) - 18)
                Else
                    HpSana = Porcentaje(.Stats.MaxHP, 1)
                End If
                
                If HpSana < 1 Then
                    HpSana = 1
                End If
                
                .Counters.HPCounter = 0
                .Stats.MinHP = .Stats.MinHP + HpSana
                
                If .Stats.MinHP > .Stats.MaxHP Then
                    .Stats.MinHP = .Stats.MaxHP
                End If

                Call WriteUpdateHP(UserIndex)
            End If
        End If
    End With

End Sub

Public Sub CargaNpcsDat()
    Dim NpcFile As String
    
    NpcFile = DatPath & "Npcs.dat"
    Call LeerNpcs.Initialize(NpcFile)
End Sub

Public Sub PasarSegundo()
On Error GoTo errhandler
    Dim i As Long
    
    For i = 1 To LastUser
        If UserList(i).flags.Logged Then
            'Cerrar usuario
            If UserList(i).Counters.Saliendo Then
                UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
                If UserList(i).Counters.Salir < 1 Then
                    Call FlushBuffer(i)
                    Call CloseSocket(i)
                'Else
                    'Call WriteConsoleMsg(i, UserList(i).Counters.Salir, FontTypeNames.FONTTYPE_INFO)
                End If
            End If
            
            'Respawnear usuario (mandarlo a su casa)
            If UserList(i).Stats.Muerto Then
                UserList(i).Counters.Respawn = UserList(i).Counters.Respawn - 1
                If UserList(i).Counters.Respawn < 1 Then
                    Call RespawnearUsuario(i)
                End If
            End If
            
            'Respawnear usuario (mandarlo a su casa)
            If UserList(i).Counters.EnPlataforma > 0 Then
                UserList(i).Counters.EnPlataforma = UserList(i).Counters.EnPlataforma + 1
                If UserList(i).Counters.EnPlataforma > 10 Then
                    Dim nPos As WorldPos
            
                    Call ClosestLegalPos(UserList(i).Pos, nPos)
            
                    If nPos.X > 0 And nPos.Y > 0 Then
                        Call WarpUserChar(i, UserList(i).Pos.Map, nPos.X, nPos.Y, False)
                        
                        UserList(i).Counters.EnPlataforma = 0
                    End If
                End If
            End If
        End If
    Next i
Exit Sub

errhandler:
    Call LogError("Error en PasarSegundo. Err: " & Err.description & " - " & Err.Number & " - UserIndex: " & i)
    Resume Next
End Sub
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
    'WorldSave
    Call ES.DoBackUp

    'commit experiencias
    Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios
    
    'If EjecutarLauncher Then
    '    Shell (App.Path & "/launcher.exe")
    'End If
    
    'Chauuu
    Unload frmMain

End Sub

Public Sub GuardarUsuarios()
    'haciendoBK = True
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.Logged Then
            Call SaveUser(i)
        End If
    Next i
    
    Call SendData(SendTarget.ToAll, 0, Msg_ConsoleMsg("Los personajes fueron grabados.", FontTypeNames.FONTTYPE_VENENO))
    'haciendoBK = False
End Sub

Public Sub FreeNpcs()
'Releases all Npc Indexes
    Dim LoopC As Long
    
    'Free all Npc Indexes
    For LoopC = 1 To MaxNpcs
        NpcList(LoopC).flags.NpcActive = False
    Next LoopC
End Sub

Public Sub FreeCharIndexes()
'Releases all char Indexes
    'Free all char Indexes (set them all to 0)
    Call ZeroMemory(CharList(1), MaxChars * Len(CharList(1)))
End Sub

Public Function Tilde(data As String) As String
    Tilde = Replace(Replace(Replace(Replace(Replace(UCase$(data), "Á", "A"), "É", "E"), "Í", "I"), "Ó", "O"), "Ú", "U")
End Function

Public Function EsCompaniero(ByVal UserIndex As Integer, ByVal CompaName As String) As Byte

On Error GoTo errhandler
    
    Dim j As Byte
    For j = 1 To MaxCompaSlots
        If LenB(UserList(UserIndex).Compas.Compa(j)) = LenB(CompaName) Then
            If UserList(UserIndex).Compas.Compa(j) = CompaName Then
                EsCompaniero = j
                Exit Function
            End If
        End If
    Next

Exit Function
errhandler:

End Function

Public Sub AgregarCompaniero(ByVal UserIndex As Integer, ByVal CompaName As String)
    
    Dim j As Byte
    Dim CompaIndex As Integer
    
    With UserList(UserIndex)
       
        CompaIndex = NameIndex(CompaName)
    
        If CompaIndex > 0 Then
            If .Pos.Map <> UserList(CompaIndex).Pos.Map Then
                If .flags.Privilegios And PlayerType.User Then
                    Exit Sub
                End If
            End If
        ElseIf .flags.Privilegios And PlayerType.User Then
            Exit Sub
        End If
        
        If .Compas.Nro = MaxCompaSlots Then
            Call WriteConsoleMsg(UserIndex, "No tenés espacio para más compañeros.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If EsAdmin(CompaName) Or EsDios(CompaName) Or _
        EsSemiDios(CompaName) Or EsConsejero(CompaName) Then
            If .flags.Privilegios And PlayerType.User Then
                Call WriteConsoleMsg(UserIndex, "No podés agregar como compañeros a un administrador.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        'Buscamos un slot vacio
        For j = 1 To MaxCompaSlots
            If LenB(.Compas.Compa(j)) < 1 Then
                Exit For
            End If
        Next j
        
        .Compas.Compa(j) = CompaName
        
        .Compas.Nro = .Compas.Nro + 1
    End With

    If CompaIndex > 0 Then
    
        Call WriteAddCompa(UserIndex, j, True)

        With UserList(CompaIndex)
            'Buscamos un slot vacio
            For j = 1 To MaxCompaSlots
                If LenB(.Compas.Compa(j)) < 1 Then
                    Exit For
                End If
            Next j
                
            If .Compas.Nro = MaxCompaSlots Then
                Call WriteConsoleMsg(CompaIndex, "No tenés espacio para más compañeros.", FontTypeNames.FONTTYPE_INFO)
            Else
                .Compas.Compa(j) = UserList(UserIndex).Name
                .Compas.Nro = .Compas.Nro + 1
                
                Call WriteAddCompa(CompaIndex, j, True, True)
            End If
        End With
    Else

        Call WriteAddCompa(UserIndex, j, False)

        Dim CompaStr As String
        
        Call DB_RS_Open("SELECT * from people WHERE `name`='" & CompaName & "'")
        
        CompaStr = DB_RS!Compas
        
        If LenB(CompaStr) > 0 Then
            CompaStr = CompaStr & vbNewLine & UserList(UserIndex).Name
        Else
            CompaStr = UserList(UserIndex).Name
        End If

        DB_RS!Compas = CompaStr
        
        DB_RS.Update
        DB_RS_Close
    End If
End Sub

Public Sub QuitarCompaniero(ByVal UserIndex As Integer, ByVal Slot As Byte)

    Dim CompaName As String
    Dim CompaIndex As Integer
    
    With UserList(UserIndex)
        CompaName = .Compas.Compa(Slot)
        
        .Compas.Compa(Slot) = vbNullString
                
        .Compas.Nro = .Compas.Nro - 1
    End With
    
    CompaIndex = NameIndex(CompaName)
    
    Call WriteQuitarCompa(UserIndex, Slot)

    If CompaIndex > 0 Then
        Dim j As Byte
        
        With UserList(CompaIndex)
            For j = 1 To MaxCompaSlots
                If LenB(.Compas.Compa(j)) = LenB(UserList(UserIndex).Name) Then
                    If .Compas.Compa(j) = UserList(UserIndex).Name Then
                        Exit For
                    End If
                End If
            Next j
                
            .Compas.Compa(j) = vbNullString
            .Compas.Nro = .Compas.Nro - 1
                
            Call WriteQuitarCompa(CompaIndex, j)
        End With
    Else
        Dim CompaStr As String
        
        Call DB_RS_Open("SELECT * from people WHERE `name`='" & CompaName & "'")
        
        CompaStr = DB_RS!Compas
        
        If InStr(CompaStr, vbNewLine & UserList(UserIndex).Name) > 0 Then
            CompaStr = Replace(CompaStr, vbNewLine & UserList(UserIndex).Name, vbNullString)
        ElseIf InStr(CompaStr, UserList(UserIndex).Name & vbNewLine) > 0 Then
            CompaStr = Replace(CompaStr, UserList(UserIndex).Name & vbNewLine, vbNullString)
        Else
            CompaStr = vbNullString
        End If

        DB_RS!Compas = CompaStr
        
        DB_RS.Update
        DB_RS_Close
    End If

End Sub

Public Sub RegistrarEstadisticas()
    Call DB_RS_Open("SELECT * from history WHERE 0=1")
    
    DB_RS.AddNew
    
    DB_RS!People = Poblacion
    DB_RS!Record = RecordPoblacion
    
    DB_RS.Update
    DB_RS_Close
End Sub

Public Function Calcular_ELU(ByVal Nivel As Byte) As Long
    
    Dim i As Byte
    
    Calcular_ELU = 100
    
    For i = 2 To Nivel
    
        Select Case Nivel
            Case Is < 15
                Calcular_ELU = Calcular_ELU * 1.25
            Case Is < 25
                Calcular_ELU = Calcular_ELU * 1.275
            Case Is < 30
                Calcular_ELU = Calcular_ELU * 1.3
            Case Is < 35
                Calcular_ELU = Calcular_ELU * 1.325
            Case Is < 40
                Calcular_ELU = Calcular_ELU * 1.35
            Case Is < 45
                Calcular_ELU = Calcular_ELU * 1.36
            Case Is < 48
                Calcular_ELU = Calcular_ELU * 1.375
            Case Is < 50
                Calcular_ELU = Calcular_ELU * 1.4
            Case Else
                Calcular_ELU = Calcular_ELU * 1.5
        End Select
    Next i
    
End Function

Public Function UsaArco(ByVal UserIndex As Integer) As Integer

    With UserList(UserIndex)
        If .Inv.LeftHand > 0 Then
            If ObjData(.Inv.LeftHand).Type = otArma Then
                UsaArco = .Inv.LeftHand
            End If
        End If
    End With
    
End Function

Public Function UsaArmaNoArco(ByVal UserIndex As Integer) As Integer

    With UserList(UserIndex)
        If .Inv.RightHand > 0 Then
            If ObjData(.Inv.RightHand).Type = otArma Then
                UsaArmaNoArco = .Inv.RightHand
            End If
        End If
    End With
    
End Function

Public Function UsaEscudo(ByVal UserIndex As Integer) As Integer

    With UserList(UserIndex)
        If .Inv.LeftHand > 0 Then
            If ObjData(.Inv.LeftHand).Type = otEscudo Then
                UsaEscudo = .Inv.LeftHand
            End If
        End If
    End With
    
End Function

Public Function ExistePlataforma(ByVal mapa As Integer) As Byte
    Dim i As Byte
    
    For i = 1 To MaxPlataformSlots
        If Plataforma(i).Map = mapa Then
            ExistePlataforma = i
            Exit For
        End If
    Next i
End Function

Public Function TienePlataforma(ByVal UserIndex As Integer, ByVal mapa As Integer) As Byte
    Dim i As Byte
    
    For i = 1 To MaxPlataformSlots
        If UserList(UserIndex).Plataformas.Plataforma(i).Map = mapa Then
            TienePlataforma = i
            Exit Function
        End If
    Next i
End Function

Public Sub AgregarPlataforma(ByVal UserIndex As Integer, ByVal mapa As Integer)
     If ExistePlataforma(mapa) > 0 Then
        If TienePlataforma(UserIndex, mapa) < 1 Then
            Dim i As Byte
            
            With UserList(UserIndex)
                For i = 1 To MaxPlataformSlots
                    If .Plataformas.Plataforma(i).Map = 0 Then
                        .Plataformas.Plataforma(i).Map = mapa
                        .Plataformas.Nro = .Plataformas.Nro + 1
                        Exit For
                    End If
                Next i
            End With
        End If
    End If
End Sub
