Attribute VB_Name = "ES"
Option Explicit

Public Sub CargarSpawnList()
    Dim N As Integer, LoopC As Integer
    N = Val(GetVar(DatPath & "invokar.dat", "INIT", "NumNpcs"))
    ReDim Spawn_List(N) As tCriaturasEntrenador
    For LoopC = 1 To N
        Spawn_List(LoopC).NpcIndex = Val(GetVar(DatPath & "invokar.dat", "LIST", "NI" & LoopC))
        Spawn_List(LoopC).NpcName = GetVar(DatPath & "invokar.dat", "LIST", "NN" & LoopC)
    Next LoopC
End Sub

Public Function EsAdmin(ByVal Name As String) As Boolean
    EsAdmin = (Val(Administradores.GetValue("Admin", Name)) = 1)
End Function

Public Function EsDios(ByVal Name As String) As Boolean
    EsDios = (Val(Administradores.GetValue("Dios", Name)) = 1)
End Function

Public Function EsSemiDios(ByVal Name As String) As Boolean
    EsSemiDios = (Val(Administradores.GetValue("SemiDios", Name)) = 1)
End Function

Public Function EsConsejero(ByVal Name As String) As Boolean
    EsConsejero = (Val(Administradores.GetValue("Consejero", Name)) = 1)
End Function

Public Function EsRolesMaster(ByVal Name As String) As Boolean
    EsRolesMaster = (Val(Administradores.GetValue("RM", Name)) = 1)
End Function

Public Sub LoadAdministrativeUsers()
'Admines     => Admin
'Dioses      => Dios
'SemiDioses  => SemiDios
'Consejeros  => Consejero
'RoleMasters => RM

    'Si esta mierda tuviese array asociativos el código sería tan lindo.
    Dim buf As Integer
    Dim i As Long
    Dim Name As String
       
    'Public container
    Set Administradores = New clsIniManager
    
    'Server ini info file
    Dim AdminsIni As clsIniManager
    Set AdminsIni = New clsIniManager
    
    Call AdminsIni.Initialize(ServidorIni)
    
    'Admines
    buf = Val(AdminsIni.GetValue("INIT", "Admines"))
    
    For i = 1 To buf
        Name = UCase$(AdminsIni.GetValue("Admines", "Admin" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then
            Name = Right$(Name, Len(Name) - 1)
        End If
        
        'Add key
        Call Administradores.ChangeValue("Admin", Name, "1")

    Next i
    
    'Dioses
    buf = Val(AdminsIni.GetValue("INIT", "Dioses"))
    
    For i = 1 To buf
        Name = UCase$(AdminsIni.GetValue("Dioses", "Dios" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then
            Name = Right$(Name, Len(Name) - 1)
        End If
        
        'Add key
        Call Administradores.ChangeValue("Dios", Name, "1")
        
    Next i
    
    'SemiDioses
    buf = Val(AdminsIni.GetValue("INIT", "SemiDioses"))
    
    For i = 1 To buf
        Name = UCase$(AdminsIni.GetValue("SemiDioses", "SemiDios" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then
            Name = Right$(Name, Len(Name) - 1)
        End If
        
        'Add key
        Call Administradores.ChangeValue("SemiDios", Name, "1")
        
    Next i
    
    'Consejeros
    buf = Val(AdminsIni.GetValue("INIT", "Consejeros"))
        
    For i = 1 To buf
        Name = UCase$(AdminsIni.GetValue("Consejeros", "Consejero" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then
            Name = Right$(Name, Len(Name) - 1)
        End If
        
        'Add key
        Call Administradores.ChangeValue("Consejero", Name, "1")
        
    Next i
    
    'RolesMasters
    buf = Val(AdminsIni.GetValue("INIT", "RolesMasters"))
        
    For i = 1 To buf
        Name = UCase$(AdminsIni.GetValue("RolesMasters", "RM" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then
            Name = Right$(Name, Len(Name) - 1)
        End If
        
        'Add key
        Call Administradores.ChangeValue("RM", Name, "1")
    Next i
    
    Set AdminsIni = Nothing
    
End Sub

Public Function TxtDimension(ByVal Name As String) As Long
    Dim N As Integer, cad As String, Tam As Long
    N = FreeFile(1)
    Open Name For Input As #N
    Tam = 0
    Do While Not EOF(N)
        Tam = Tam + 1
        Line Input #N, cad
    Loop
    Close N
    TxtDimension = Tam
End Function

Public Sub CargarForbidenWords()
    ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
    Dim N As Integer, i As Integer
    N = FreeFile(1)
    Open DatPath & "NombresInvalidos.txt" For Input As #N
    
    For i = 1 To UBound(ForbidenNames)
        Line Input #N, ForbidenNames(i)
    Next i
    
    Close N
End Sub

Public Sub CargarHechizos()

On Error GoTo errhandler

    If frmMain.Visible Then
        frmMain.txStatus.Caption = "Cargando Hechizos."
    End If
    
    Dim Hechizo As Integer
    Dim Leer As New clsIniManager
    
    Call Leer.Initialize(DatPath & "Hechizos.dat")
    
    'obtiene el numero de hechizos
    NumeroHechizos = Val(Leer.GetValue("INIT", "NumeroHechizos"))
    
    ReDim Hechizos(1 To NumeroHechizos) As tHechizo
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumeroHechizos
    frmCargando.cargar.value = 0
    
    'Llena la lista
    For Hechizo = 1 To NumeroHechizos
        With Hechizos(Hechizo)
            .Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
            .Desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
            .PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
    
            .HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
            .TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
            .PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
        
            .Tipo = Val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
            .WAV = Val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
            .FXgrh = Val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
        
            .Loops = Val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
        
        '.Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
        
            .SubeHP = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
            .MinHP = Val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
            .MaxHP = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
        
            .SubeMana = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
            .MiMana = Val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
            .MaMana = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
        
            .SubeSta = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
            .MinSta = Val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
            .MaxSta = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
        
            .SubeHam = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
            .MinHam = Val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
            .MaxHam = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
        
            .SubeSed = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
            .MinSed = Val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
            .MaxSed = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
        
            .SubeAgilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
            .MinAgilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
            .MaxAgilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
        
            .SubeFuerza = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
            .MinFuerza = Val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
            .MaxFuerza = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
        
            .SubeCarisma = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
            .MinCarisma = Val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
            .MaxCarisma = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
        
        
            .Invisibilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
            .Paraliza = Val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
            .Inmoviliza = Val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
            .RemoverParalisis = Val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
            .RemoverEstupidez = Val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
            .RemueveInvisibilidadParcial = Val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
        
        
            .CuraVeneno = Val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
            .Envenena = Val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
            .Maldicion = Val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
            .RemoverMaldicion = Val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
            .Bendicion = Val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
            .Revivir = Val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
        
            .Ceguera = Val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
            .Estupidez = Val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
        
            .Warp = Val(Leer.GetValue("Hechizo" & Hechizo, "Warp"))
        
            .Invoca = Val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
            .NumNpc = Val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
            .Cant = Val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
            .Mimetiza = Val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
        
        '.Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
        '.ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
        
            .MinSkill = Val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
            .ManaRequerido = Val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
        
            .StaRequerido = Val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
        
            .Target = Val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
        frmCargando.cargar.value = frmCargando.cargar.value + 1
        
            .NeedStaff = Val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
            .StaffAffected = CBool(Val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
        End With
    Next Hechizo

    Set Leer = Nothing
    
    Exit Sub

errhandler:
 MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.description
 
End Sub

Public Sub LoadMotd()
    Dim i As Integer
    
    MaxLines = Val(GetVar(DatPath & "motd.ini", "INIT", "NumLines"))
    
    ReDim Motd(1 To MaxLines)
    For i = 1 To MaxLines
        Motd(i).texto = GetVar(DatPath & "motd.ini", "Motd", "Line" & i)
        Motd(i).Formato = vbNullString
    Next i
End Sub

Public Sub DoBackUp()
'Call LogTarea("PUBLIC SUB DoBackUp")
    haciendoBK = True
    Dim i As Integer
    
    
    
    'Lo saco porque elimina elementales y Mascotas - Maraxus
    'lo pongo aca x sugernecia del yind
    'For i = 1 To LastNpc
    'If NpcList(i).flags.NpcActive Then
    'If NpcList(i).Contadores.TiempoExistencia > 0 Then
    'Call MuereNpc(i, 0)
    'End If
    'End If
    'Next i
    '/'lo pongo aca x sugernecia del yind

    Call SendData(SendTarget.ToAll, 0, Msg_PauseToggle())
    
    Call LimpiarMundo
    Call WorldSave
    Call modGuilds.v_RutinaElecciones
    Call ResetCentinelaInfo     'Reseteamos al centinela
    
    Call SendData(SendTarget.ToAll, 0, Msg_PauseToggle())
        
    haciendoBK = False
    
    'Log
    On Error Resume Next
    Dim nfile As Integer
    nfile = FreeFile 'obtenemos un canal
    Open App.Path & "/logs/backUps.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time
    Close #nfile
End Sub

Public Sub GrabarMapa(ByVal Map As Long, ByRef MapFile As String)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2011
'10/08/2010 - Pato: Implemento el clsByteBuffer para el grabado de mapas
'28/10/2010:ZaMa - Ahora no se hace backup de los pretorianos.
'12/01/2011 - Amraphen: Ahora no se hace backup de NPCs prohibidos (Pretorianos, Mascotas, Invocados y Centinela)
'***************************************************

On Error Resume Next
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim LoopC As Long
    Dim MapWriter As clsByteBuffer
    Dim InfWriter As clsByteBuffer
    Dim IniManager As clsIniManager
    Dim NpcInvalido As Boolean
    
    Set MapWriter = New clsByteBuffer
    Set InfWriter = New clsByteBuffer
    Set IniManager = New clsIniManager
    
    If FileExist(MapFile & ".map", vbNormal) Then
        Kill MapFile & ".map"
    End If
    
    If FileExist(MapFile & ".inf", vbNormal) Then
        Kill MapFile & ".inf"
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open MapFile & ".Map" For Binary As FreeFileMap
    
    Call MapWriter.initializeWriter(FreeFileMap)
    
    'Open .inf file
    FreeFileInf = FreeFile
    Open MapFile & ".Inf" For Binary As FreeFileInf
    
    Call InfWriter.initializeWriter(FreeFileInf)
    
    'map Header
    Call MapWriter.putInteger(MapInfo(Map).MapVersion)
        
    Call MapWriter.putString(MiCabecera.Desc, False)
    Call MapWriter.putLong(MiCabecera.Crc)
    Call MapWriter.putLong(MiCabecera.MagicWord)
    
    Call MapWriter.putDouble(0)
    
    'inf Header
    Call InfWriter.putDouble(0)
    Call InfWriter.putInteger(0)
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            With MapData(X, Y)
                ByFlags = 0
                
                If .Blocked Then ByFlags = ByFlags Or 1
                If .Graphic(2) Then ByFlags = ByFlags Or 2
                If .Graphic(3) Then ByFlags = ByFlags Or 4
                If .Graphic(4) Then ByFlags = ByFlags Or 8
                If .Trigger Then ByFlags = ByFlags Or 16
                
                Call MapWriter.putByte(ByFlags)
                
                Call MapWriter.putInteger(.Graphic(1))
                
                For LoopC = 2 To 4
                    If .Graphic(LoopC) Then _
                        Call MapWriter.putInteger(.Graphic(LoopC))
                Next LoopC
                
                If .Trigger Then _
                    Call MapWriter.putInteger(CInt(.Trigger))
                
                '.inf file
                ByFlags = 0
                
                If .ObjInfo.index > 0 Then
                   If ObjData(.ObjInfo.index).Type = otFogata Then
                        .ObjInfo.index = 0
                        .ObjInfo.Amount = 0
                    End If
                End If
    
                If .TileExit.Map Then ByFlags = ByFlags Or 1
                
                ' No hacer backup de los NPCs inválidos (Mascotas, Invocados y Centinela)
                If .NpcIndex Then
                    NpcInvalido = NpcList(.NpcIndex).MaestroUser > 0 Or EsCentinela(.NpcIndex)
                    
                    If Not NpcInvalido Then ByFlags = ByFlags Or 2
                End If
                
                If .ObjInfo.index Then ByFlags = ByFlags Or 4
                
                Call InfWriter.putByte(ByFlags)
                
                If .TileExit.Map Then
                    Call InfWriter.putInteger(.TileExit.Map)
                    Call InfWriter.putInteger(.TileExit.X)
                    Call InfWriter.putInteger(.TileExit.Y)
                End If
                
                If .NpcIndex And Not NpcInvalido Then _
                    Call InfWriter.putInteger(NpcList(.NpcIndex).Numero)
                
                If .ObjInfo.index Then
                    Call InfWriter.putInteger(.ObjInfo.index)
                    Call InfWriter.putInteger(.ObjInfo.Amount)
                End If
                
                NpcInvalido = False
            End With
        Next X
    Next Y
    
    Call MapWriter.saveBuffer
    Call InfWriter.saveBuffer
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf
    
    Set MapWriter = Nothing
    Set InfWriter = Nothing

    With MapInfo(Map)
        'write .dat file
        Call IniManager.ChangeValue("Mapa" & Map, "MusicNum", .Music)
        Call IniManager.ChangeValue("Mapa" & Map, "MagiaSinefecto", .MagiaSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "InviSinEfecto", .InviSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "ResuSinEfecto", .ResuSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "StartPos", .StartPos.Map & "-" & .StartPos.X & "-" & .StartPos.Y)
    
        Call IniManager.ChangeValue("Mapa" & Map, "Terreno", TerrainByteToString(.terreno))
        Call IniManager.ChangeValue("Mapa" & Map, "Zona", .Zona)
        Call IniManager.ChangeValue("Mapa" & Map, "Restringir", RestrictByteToString(.restringir))
        Call IniManager.ChangeValue("Mapa" & Map, "BackUp", str(.BackUp))
    
        If .PK Then
            Call IniManager.ChangeValue("Mapa" & Map, "Pk", "0")
        Else
            Call IniManager.ChangeValue("Mapa" & Map, "Pk", "1")
        End If
        
        Call IniManager.ChangeValue("Mapa" & Map, "OcultarSinEfecto", .OcultarSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "InvocarSinEfecto", .InvocarSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "RoboNpcsPermitido", .RoboNpcsPermitido)
    
        Call IniManager.DumpFile(MapFile & ".dat")
    End With
    
    Set IniManager = Nothing
End Sub

Public Sub LoadArmasHerreria()

    Dim N As Integer, lc As Integer
    
    N = Val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))
    
    ReDim Preserve ArmasHerrero(1 To N) As Integer
    
    For lc = 1 To N
        ArmasHerrero(lc) = Val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
    Next lc

End Sub

Public Sub LoadArmadurasHerreria()
    
    Dim N As Integer, lc As Integer
    
    N = Val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))
    
    ReDim Preserve ArmadurasHerrero(1 To N) As Integer
    
    For lc = 1 To N
        ArmadurasHerrero(lc) = Val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
    Next lc
    
End Sub

Public Sub LoadBalance()
    
    Dim i As Long
    
    'Modificadores de Clase
    For i = 1 To NUMCLASES
        With ModClase(i)
            .Evasion = Val(GetVar(DatPath & "Balance.dat", "MODEVASION", ListaClases(i)))
            .AtaqueArmas = Val(GetVar(DatPath & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
            .AtaqueProyectiles = Val(GetVar(DatPath & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
            .DanioArmas = Val(GetVar(DatPath & "Balance.dat", "MODDAÑOARMAS", ListaClases(i)))
            .DanioProyectiles = Val(GetVar(DatPath & "Balance.dat", "MODDAÑOPROYECTILES", ListaClases(i)))
            .DanioWrestling = Val(GetVar(DatPath & "Balance.dat", "MODDAÑOWRESTLING", ListaClases(i)))
            .Escudo = Val(GetVar(DatPath & "Balance.dat", "MODESCUDO", ListaClases(i)))
        End With
    Next i
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS
        With ModRaza(i)
            .Fuerza = Val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Fuerza"))
            .Agilidad = Val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Agilidad"))
            .Inteligencia = Val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Inteligencia"))
            .Carisma = Val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Carisma"))
            .Constitucion = Val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Constitucion"))
        End With
    Next i
    
    'Modificadores de Vida
    For i = 1 To NUMCLASES
        ModVida(i) = Val(GetVar(DatPath & "Balance.dat", "MODVIDA", ListaClases(i)))
    Next i
    
    'Distribución de Vida
    For i = 1 To 5
        DistribucionEnteraVida(i) = Val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "E" + CStr(i)))
    Next i
    For i = 1 To 4
        DistribucionSemienteraVida(i) = Val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "S" + CStr(i)))
    Next i
    
    'Extra
    'PorcentajeRecuperoMana = Val(GetVar(DatPath & "Balance.dat", "EXTRA", "PorcentajeRecuperoMana"))

    'Party
    ExponenteNivelParty = Val(GetVar(DatPath & "Balance.dat", "PARTY", "ExponenteNivelParty"))
End Sub

Public Sub LoadObjCarpintero()

    Dim N As Integer, lc As Integer
    
    N = Val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjCarpintero(1 To N) As Integer
    
    For lc = 1 To N
        ObjCarpintero(lc) = Val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
    Next lc

End Sub

Public Sub LoadPlataformas()

    Dim a As Byte, b As Byte
    
    a = Val(GetVar(DatPath & "Plataformas.dat", "INIT", "NumPlataformas"))
    
    For b = 1 To a
        Plataforma(b).Map = Val(GetVar(DatPath & "Plataformas.dat", "Plataforma" & b, "Mapa"))
        Plataforma(b).X = Val(GetVar(DatPath & "Plataformas.dat", "Plataforma" & b, "X"))
        Plataforma(b).Y = Val(GetVar(DatPath & "Plataformas.dat", "Plataforma" & b, "Y"))
    Next b
End Sub

Public Sub LoadOBJData()

On Error Resume Next

    If frmMain.Visible Then
        frmMain.txStatus.Caption = "Cargando objetos."
    End If
    
    'Carga la lista de objetos
    Dim Obj As Integer
    Dim Leer As New clsIniManager
    
    Call Leer.Initialize(DatPath & "Obj.dat")
    
    'obtiene el numero de obj
    NumObjDatas = Val(Leer.GetValue("INIT", "NumObjs"))
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumObjDatas
    frmCargando.cargar.value = 0
    
    ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
      
    'Llena la lista
    For Obj = 1 To NumObjDatas
        With ObjData(Obj)
            .Name = Leer.GetValue("OBJ" & Obj, "Name")

            .Log = Val(Leer.GetValue("OBJ" & Obj, "Log"))
            .NoLog = Val(Leer.GetValue("OBJ" & Obj, "NoLog"))

            .GrhIndex = Val(Leer.GetValue("OBJ" & Obj, "GrhIndex"))
        
            .Type = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))

            Select Case .Type
            
                Case otArmadura
                    .Guild = Val(Leer.GetValue("OBJ" & Obj, "Guild"))
                    .LingH = Round(Val(Leer.GetValue("OBJ" & Obj, "LingH")) / 10)
                    .LingP = Round(Val(Leer.GetValue("OBJ" & Obj, "LingP")) / 10)
                    .LingO = Round(Val(Leer.GetValue("OBJ" & Obj, "LingO")) / 10)
                    .SkHerreria = Val(Leer.GetValue("OBJ" & Obj, "SkHerreria"))
            
                Case otEscudo
                    .ShieldAnim = Val(Leer.GetValue("OBJ" & Obj, "Anim"))
                    .LingH = Round(Val(Leer.GetValue("OBJ" & Obj, "LingH")) / 10)
                    .LingP = Round(Val(Leer.GetValue("OBJ" & Obj, "LingP")) / 10)
                    .LingO = Round(Val(Leer.GetValue("OBJ" & Obj, "LingO")) / 10)
                    .SkHerreria = Val(Leer.GetValue("OBJ" & Obj, "SkHerreria"))
                    .Guild = Val(Leer.GetValue("OBJ" & Obj, "Guild"))
            
                Case otCasco
                    .HeadAnim = Val(Leer.GetValue("OBJ" & Obj, "Anim"))
                    .LingH = Round(Val(Leer.GetValue("OBJ" & Obj, "LingH")) / 10)
                    .LingP = Round(Val(Leer.GetValue("OBJ" & Obj, "LingP")) / 10)
                    .LingO = Round(Val(Leer.GetValue("OBJ" & Obj, "LingO")) / 10)
                    .SkHerreria = Val(Leer.GetValue("OBJ" & Obj, "SkHerreria"))
                    .Guild = Val(Leer.GetValue("OBJ" & Obj, "Guild"))
            
                Case otArma
                    .WeaponAnim = Val(Leer.GetValue("OBJ" & Obj, "Anim"))
                    .Apuñala = Val(Leer.GetValue("OBJ" & Obj, "Apuñala"))
                    .Envenena = Val(Leer.GetValue("OBJ" & Obj, "Envenena"))
                    .Paraliza = Val(Leer.GetValue("OBJ" & Obj, "Paraliza"))
                    .MinHit = Val(Leer.GetValue("OBJ" & Obj, "MinHit"))
                    .MaxHit = Val(Leer.GetValue("OBJ" & Obj, "MaxHit"))
                    
                    If .MaxHit < .MinHit Then
                        .MaxHit = .MinHit
                    End If
                    
                    .Proyectil = Val(Leer.GetValue("OBJ" & Obj, "Proyectil"))
                    .Municion = Val(Leer.GetValue("OBJ" & Obj, "Municiones"))
                    .StaffPower = Val(Leer.GetValue("OBJ" & Obj, "StaffPower"))
                    .StaffDamageBonus = Val(Leer.GetValue("OBJ" & Obj, "StaffDamageBonus"))
                    .Refuerzo = Val(Leer.GetValue("OBJ" & Obj, "Refuerzo"))
                    
                    .LingH = Round(Val(Leer.GetValue("OBJ" & Obj, "LingH")) / 10)
                    .LingP = Round(Val(Leer.GetValue("OBJ" & Obj, "LingP")) / 10)
                    .LingO = Round(Val(Leer.GetValue("OBJ" & Obj, "LingO")) / 10)
                    .SkHerreria = Val(Leer.GetValue("OBJ" & Obj, "SkHerreria"))
                    .Guild = Val(Leer.GetValue("OBJ" & Obj, "Guild"))
            
                Case otInstrumento
                    .Snd1 = Val(Leer.GetValue("OBJ" & Obj, "SND1"))
                    .Snd2 = Val(Leer.GetValue("OBJ" & Obj, "SND2"))
                    .Snd3 = Val(Leer.GetValue("OBJ" & Obj, "SND3"))
                    .Guild = Val(Leer.GetValue("OBJ" & Obj, "Guild"))
            
                Case otMineral
                    .MinSkill = Val(Leer.GetValue("OBJ" & Obj, "MinSkill"))
            
                Case otPuerta, otBotellaVacia, otBotellaLlena
                    .IndexAbierta = Val(Leer.GetValue("OBJ" & Obj, "IndexAbierta"))
                    .IndexCerrada = Val(Leer.GetValue("OBJ" & Obj, "IndexCerrada"))
                    .IndexCerradaLlave = Val(Leer.GetValue("OBJ" & Obj, "IndexCerradaLlave"))
            
                Case otPocion
                    .TipoPocion = Val(Leer.GetValue("OBJ" & Obj, "TipoPocion"))
                    .MaxModificador = Val(Leer.GetValue("OBJ" & Obj, "MaxModificador"))
                    .MinModificador = Val(Leer.GetValue("OBJ" & Obj, "MinModificador"))
                    .DuracionEfecto = Val(Leer.GetValue("OBJ" & Obj, "DuracionEfecto"))
                
                Case otBarco
                    .MinSkill = Val(Leer.GetValue("OBJ" & Obj, "MinSkill"))
                    .MaxHit = Val(Leer.GetValue("OBJ" & Obj, "MaxHit"))
                    .MinHit = Val(Leer.GetValue("OBJ" & Obj, "MinHit"))
            
                Case otFlecha
                    .MaxHit = Val(Leer.GetValue("OBJ" & Obj, "MaxHit"))
                    .MinHit = Val(Leer.GetValue("OBJ" & Obj, "MinHit"))
                    .Envenena = Val(Leer.GetValue("OBJ" & Obj, "Envenena"))
                    .Paraliza = Val(Leer.GetValue("OBJ" & Obj, "Paraliza"))
                
                Case otAnillo
                    .LingH = Round(Val(Leer.GetValue("OBJ" & Obj, "LingH")) / 10)
                    .LingP = Round(Val(Leer.GetValue("OBJ" & Obj, "LingP")) / 10)
                    .LingO = Round(Val(Leer.GetValue("OBJ" & Obj, "LingO")) / 10)
                    .SkHerreria = Val(Leer.GetValue("OBJ" & Obj, "SkHerreria"))
                
                Case otPortal
                    .Radio = Val(Leer.GetValue("OBJ" & Obj, "Radio"))
                
                Case otPasaje
                    .DesdeMap = Val(Leer.GetValue("OBJ" & Obj, "DesdeMap"))
                    .HastaMap = Val(Leer.GetValue("OBJ" & Obj, "HastaMap"))
                    .HastaX = Val(Leer.GetValue("OBJ" & Obj, "HastaX"))
                    .HastaY = Val(Leer.GetValue("OBJ" & Obj, "HastaY"))
                    
                Case otCinturon
                    .NumColumnas = Val(Leer.GetValue("OBJ" & Obj, "Columnas"))
                    .NumFilas = Val(Leer.GetValue("OBJ" & Obj, "Filas"))
                                        
            End Select
    
            .BodyAnim = Val(Leer.GetValue("OBJ" & Obj, "NumRopaje"))
            .SpellIndex = Val(Leer.GetValue("OBJ" & Obj, "HechizoIndex"))
        
            .LingoteIndex = Val(Leer.GetValue("OBJ" & Obj, "LingoteIndex"))
        
            .MineralIndex = Val(Leer.GetValue("OBJ" & Obj, "MineralIndex"))
        
            .MaxHP = Val(Leer.GetValue("OBJ" & Obj, "MaxHP"))
            .MinHP = Val(Leer.GetValue("OBJ" & Obj, "MinHP"))
        
            .Mujer = Val(Leer.GetValue("OBJ" & Obj, "Mujer"))
            .Hombre = Val(Leer.GetValue("OBJ" & Obj, "Hombre"))
        
            .MinHam = Val(Leer.GetValue("OBJ" & Obj, "MinHam"))
            .MinSed = Val(Leer.GetValue("OBJ" & Obj, "MinAgu"))
        
            .MinDef = Val(Leer.GetValue("OBJ" & Obj, "MINDEF"))
            .MaxDef = Val(Leer.GetValue("OBJ" & Obj, "MaxDEF"))
            
            .RazaEnana = Val(Leer.GetValue("OBJ" & Obj, "RazaEnana"))
            .RazaDrow = Val(Leer.GetValue("OBJ" & Obj, "RazaDrow"))
            .RazaElfa = Val(Leer.GetValue("OBJ" & Obj, "RazaElfa"))
            .RazaGnoma = Val(Leer.GetValue("OBJ" & Obj, "RazaGnoma"))
            .RazaHumana = Val(Leer.GetValue("OBJ" & Obj, "RazaHumana"))
            
            .Valor = Val(Leer.GetValue("OBJ" & Obj, "Valor"))
            
            .Crucial = Val(Leer.GetValue("OBJ" & Obj, "Crucial"))
            
            .Cerrada = (Val(Leer.GetValue("OBJ" & Obj, "abierta")) > 0)
            
            If .Cerrada Then
                .Llave = Val(Leer.GetValue("OBJ" & Obj, "Llave"))
                .clave = Val(Leer.GetValue("OBJ" & Obj, "Clave"))
            End If
    
            'Puertas y llaves
            .clave = Val(Leer.GetValue("OBJ" & Obj, "Clave"))
        
            .texto = Leer.GetValue("OBJ" & Obj, "Texto")
            .GrhSecundario = Val(Leer.GetValue("OBJ" & Obj, "VGrande"))
        
            .Agarrable = Val(Leer.GetValue("OBJ" & Obj, "Agarrable")) < 1
            .ForoID = Leer.GetValue("OBJ" & Obj, "ID")
    
            'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
            Dim i As Integer
            Dim N As Integer
            Dim S As String
            
            For i = 1 To NUMCLASES - 1
                S = UCase$(Leer.GetValue("OBJ" & Obj, "CP" & i))
     
                If S <> "TRABAJADOR" Then
                    N = 1
                      
                    Do While UCase$(ListaClases(N)) <> S
                        If N < 11 Then
                            N = N + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    
                    If LenB(S) > 0 Then
                        .ClaseProhibida(i) = N
                    End If
                End If
            Next i
    
            .MaxDefM = Val(Leer.GetValue("OBJ" & Obj, "MaxDefM"))
            .MinDefM = Val(Leer.GetValue("OBJ" & Obj, "MinDefM"))
        
            .SkCarpinteria = Val(Leer.GetValue("OBJ" & Obj, "SkCarpinteria"))
    
            .Madera = Round(Val(Leer.GetValue("OBJ" & Obj, "Madera")) / 10)
            .MaderaElfica = Round(Val(Leer.GetValue("OBJ" & Obj, "MaderaElfica")) / 10)
    
            'Bebidas
            .MinSta = Val(Leer.GetValue("OBJ" & Obj, "MinST"))
    
            .NoSeCae = Val(Leer.GetValue("OBJ" & Obj, "NoSeCae"))
    
            frmCargando.cargar.value = frmCargando.cargar.value + 1
        End With
    Next Obj

    Set Leer = Nothing
    
End Sub

Public Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String

    Dim sSpaces As String 'This will hold the input that the program will retrieve
    Dim szReturn As String 'This will be the defaul value if the string is not found
      
    szReturn = vbNullString
      
    sSpaces = Space$(EmptySpaces) 'This tells the computer how long the longest string can be
      
      
    GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
      
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Public Sub CargarBackUp()

    If frmMain.Visible Then
        frmMain.txStatus.Caption = "Cargando backup."
    End If

    Dim Map As Integer
    Dim TempInt As Integer
    Dim tFileName As String
    Dim NpcFile As String

On Error GoTo man
    
    NumMaps = Val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    PopulatePolyRects
       
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.value = 0
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
    
    For Map = 1 To NumMaps
        'If Val(GetVar(MapPath & "Mapa" & Map & ".Dat", "Mapa" & Map, "BackUp")) > 0 Then
        '    tFileName = App.Path & "/WorldBackUp/Mapa" & Map
        'Else
        '    tFileName = MapPath & "Mapa" & Map
        'End If
        
        'Map_Load Map
        Call CargarMapa(Map)
        
        frmCargando.cargar.value = frmCargando.cargar.value + 1
        DoEvents
    Next Map

    Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)
 
End Sub

Public Sub LoadMapData()

    If frmMain.Visible Then
        frmMain.txStatus.Caption = "Cargando mapas..."
    End If

    Dim Map As Integer
    Dim TempInt As Integer
    Dim tFileName As String
    Dim NpcFile As String

On Error GoTo man
    
    NumMaps = Val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.value = 0
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
    
    For Map = 1 To NumMaps
        Call CargarMapa(Map)
        
        frmCargando.cargar.value = frmCargando.cargar.value + 1
        DoEvents
    Next Map

    Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)

End Sub

Public Sub CargarMapa(ByVal Map As Integer)

On Error GoTo errh
    Dim MapString As String
    
    MapString = MapPath & "Mapa" & Map
    
    Dim hFile As Integer
    Dim X As Long
    Dim Y As Long
    Dim ByFlags As Byte
    Dim NpcFile As String
    Dim Leer As clsIniManager
    Dim MapReader As clsByteBuffer
    Dim InfReader As clsByteBuffer
    Dim Buff() As Byte
    
    Set MapReader = New clsByteBuffer
    Set InfReader = New clsByteBuffer
    Set Leer = New clsIniManager
    
    NpcFile = DatPath & "Npcs.dat"
    
    hFile = FreeFile

    Open MapString & ".map" For Binary As #hFile
        Seek hFile, 1

        ReDim Buff(LOF(hFile) - 1) As Byte
    
        Get #hFile, , Buff
    Close hFile
    
    Call MapReader.initializeReader(Buff)

    'inf
    Open MapString & ".inf" For Binary As #hFile
        Seek hFile, 1

        ReDim Buff(LOF(hFile) - 1) As Byte
    
        Get #hFile, , Buff
    Close hFile
    
    Call InfReader.initializeReader(Buff)
    
    'map Header
    'MapInfo(Map).MapVersion = MapReader.getInteger
    
    'MiCabecera.Desc = MapReader.getString(Len(MiCabecera.Desc))
    MiCabecera.Crc = MapReader.getLong
    MiCabecera.MagicWord = MapReader.getLong
    
    Call MapReader.getDouble

    'inf Header
    Call InfReader.getDouble
    Call InfReader.getInteger

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            With MapData(X, Y)
                '.map file
                'ByFlags = MapReader.getByte

                .Blocked = ByFlags And 1

                .Graphic(1) = 12080 'MapReader.getInteger

                'Layer 2 used?
                If ByFlags And 2 Then
                    '.Graphic(2) = MapReader.getInteger
                End If
                
                'Layer 3 used?
                If ByFlags And 4 Then
                    '.Graphic(3) = MapReader.getInteger
                End If
                
                'Layer 4 used?
                If ByFlags And 8 Then
                    '.Graphic(4) = MapReader.getInteger
                End If
                
                'Trigger used?
                If ByFlags And 16 Then
                    '.Trigger = MapReader.getInteger
                End If
                
                '.inf file
                'ByFlags = InfReader.getByte

                'If ByFlags And 1 Then
                '    .TileExit.Map = InfReader.getInteger
                '    .TileExit.X = InfReader.getInteger
                '    .TileExit.Y = InfReader.getInteger
                'End If

                If ByFlags And 2 Then
                    'Get and make NPC
                    '.NpcIndex = InfReader.getInteger

                    If .NpcIndex > 0 Then
                        'Si el npc debe hacer respawn en la pos
                        'original la guardamos
                        If Val(GetVar(NpcFile, "NPC" & .NpcIndex, "PosOrig")) = 1 Then
                            .NpcIndex = OpenNpc(.NpcIndex)
                            NpcList(.NpcIndex).Orig.Map = Map
                            NpcList(.NpcIndex).Orig.X = X
                            NpcList(.NpcIndex).Orig.Y = Y
                        Else
                            .NpcIndex = OpenNpc(.NpcIndex)
                        End If

'                        NpcList(.NpcIndex).Pos.Map = Map
'                        NpcList(.NpcIndex).Pos.X = X
'                        NpcList(.NpcIndex).Pos.Y = Y

                        Call MakeNpcChar(True, 0, .NpcIndex, Map, X, Y)
                    End If
                End If

                If ByFlags And 4 Then
                    'Get and make Object
                    '.ObjInfo.index = InfReader.getInteger
                    '.ObjInfo.Amount = InfReader.getInteger
                End If
            End With
        Next X
    Next Y
    
    'Call Leer.Initialize(MapString & ".dat")
    
    With MapInfo(Map)
        .Music = Leer.GetValue("Mapa" & Map, "MusicNum")
        .StartPos.Map = Val(ReadField(1, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
        .StartPos.X = Val(ReadField(2, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
        .StartPos.Y = Val(ReadField(3, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
                
        .MagiaSinEfecto = Val(Leer.GetValue("Mapa" & Map, "MagiaSinEfecto"))
        .InviSinEfecto = Val(Leer.GetValue("Mapa" & Map, "InviSinEfecto"))
        .ResuSinEfecto = Val(Leer.GetValue("Mapa" & Map, "ResuSinEfecto"))
        .OcultarSinEfecto = Val(Leer.GetValue("Mapa" & Map, "OcultarSinEfecto"))
        .InvocarSinEfecto = Val(Leer.GetValue("Mapa" & Map, "InvocarSinEfecto"))
        
        .RoboNpcsPermitido = Val(Leer.GetValue("Mapa" & Map, "RoboNpcsPermitido"))
        
        If Val(Leer.GetValue("Mapa" & Map, "Pk")) = 0 Then
            .PK = True
        Else
            .PK = False
        End If
        
        .terreno = TerrainStringToByte(Leer.GetValue("Mapa" & Map, "Terreno"))
        .Zona = Leer.GetValue("Mapa" & Map, "Zona")
        .restringir = RestrictStringToByte(Leer.GetValue("Mapa" & Map, "Restringir"))
        .BackUp = Val(Leer.GetValue("Mapa" & Map, "BACKUP"))
    End With
    
    Set MapReader = Nothing
    Set InfReader = Nothing
    Set Leer = Nothing
    
    Erase Buff
Exit Sub

errh:
    Call LogError("Error cargando mapa: " & Map & " - Pos: " & X & "," & Y & "." & Err.description)

    Set MapReader = Nothing
    Set InfReader = Nothing
    Set Leer = Nothing
End Sub

Public Sub LoadSini()

On Error Resume Next

    Dim Temporal As Long
    
    If frmMain.Visible Then
        frmMain.txStatus.Caption = "Cargando info de inicio del server."
    End If
    
    BootDelBackUp = Val(GetVar(ServidorIni, "INIT", "IniciarDesdeBackUp"))
    Puerto = Val(GetVar(ServidorIni, "INIT", "StartPort"))
    HideMe = Val(GetVar(ServidorIni, "INIT", "Hide"))
    AllowMultiLogins = Val(GetVar(ServidorIni, "INIT", "AllowMultiLogins"))
    IdleLimit = Val(GetVar(ServidorIni, "INIT", "IdleLimit"))
    
    PuedeCrearPersonajes = Val(GetVar(ServidorIni, "INIT", "PuedeCrearPersonajes"))
    ServerSoloGMs = Val(GetVar(ServidorIni, "INIT", "ServerSoloGMs"))
  
    'Intervalos
    SanaIntervaloSinDescansar = Val(GetVar(ServidorIni, "INTERVALOS", "SanaIntervaloSinDescansar"))
    FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar
    
    StaminaIntervaloSinDescansar = Val(GetVar(ServidorIni, "INTERVALOS", "StaminaIntervaloSinDescansar"))
    FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar
    
    SanaIntervaloDescansar = Val(GetVar(ServidorIni, "INTERVALOS", "SanaIntervaloDescansar"))
    FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar
    
    StaminaIntervaloDescansar = Val(GetVar(ServidorIni, "INTERVALOS", "StaminaIntervaloDescansar"))
    FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar
    
    IntervaloSed = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloSed"))
    FrmInterv.txtIntervaloSed.Text = IntervaloSed
    
    IntervaloHambre = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloHambre"))
    FrmInterv.txtIntervaloHambre.Text = IntervaloHambre
    
    IntervaloVeneno = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloVeneno"))
    FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno
    
    IntervaloParalizado = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloParalizado"))
    FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado
    
    IntervaloInvisible = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloInvisible"))
    FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible
    
    IntervaloFrio = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloFrio"))
    FrmInterv.txtIntervaloFrio.Text = IntervaloFrio
    
    IntervaloWavFx = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloWAVFX"))
    FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx
    
    IntervaloInvocacion = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloInvocacion"))
    FrmInterv.txtInvocacion.Text = IntervaloInvocacion
    
    IntervaloParaConexion = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloParaConexion"))
    FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion
    
    '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
    IntervaloPuedeSerAtacado = 3000 'Cargar desde balance.dat
    IntervaloOwnedNpc = 30000 'Cargar desde balance.dat
    
    IntervaloUserPuedeCastear = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloLanzaHechizo"))
    FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear
    
    frmMain.TIMER_AI.interval = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloNpcAI"))
    FrmInterv.txtAI.Text = frmMain.TIMER_AI.interval
    
    'frmMain.TIMER_PET_AI.Interval = frmMain.TIMER_AI.Interval / 1.5
    
    frmMain.NpcAtaca.interval = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloNpcPuedeAtacar"))
    FrmInterv.txtNPCPuedeAtacar.Text = frmMain.NpcAtaca.interval
    
    IntervaloUserPuedeTrabajar = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloTrabajo"))
    FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar
    
    IntervaloUserPuedeAtacar = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloUserPuedeAtacar"))
    FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar
    
    'TODO: Agregar estos intervalos al form!!!
    IntervaloMagiaGolpe = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloMagiaGolpe"))
    IntervaloGolpeMagia = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloGolpeMagia"))
    IntervaloGolpeUsar = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloGolpeUsar"))
    
    frmMain.tLluvia.interval = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloPerdidaStaminaLluvia"))
    FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.interval
    
    MinutosWs = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloWS"))
    
    IntervaloCerrarConexion = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloCerrarConexion"))
    IntervaloUserPuedeUsar = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloUserPuedeUsar"))
    IntervaloFlechasCazadores = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloFlechasCazadores"))
    
    IntervaloOculto = Val(GetVar(ServidorIni, "INTERVALOS", "IntervaloOculto"))
    
    '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
      
    '&&&&&&&&&&&&&&&&&&&&& MULTIPLICADORES &&&&&&&&&&&&&&&&&&&&
    MultiplicadorExp = Val(GetVar(ServidorIni, "MULTIPLICADORES", "MultiplicadorExp"))
    MultiplicadorGld = Val(GetVar(ServidorIni, "MULTIPLICADORES", "MultiplicadorGld"))
    '&&&&&&&&&&&&&&&&&&&&& FIN MULTIPLICADORES &&&&&&&&&&&&&&&&&&&&
      
    RecordPoblacion = Val(GetVar(ServidorIni, "INIT", "Record"))
      
    'Max users
    Temporal = Val(GetVar(ServidorIni, "INIT", "MaxPoblacion"))
    
    If MaxPoblacion = 0 Then
        MaxPoblacion = Temporal
        ReDim UserList(1 To MaxPoblacion) As User
    End If
    
    '&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    'Se agregó en LoadBalance y en el Balance.dat
    'PorcentajeRecuperoMana = val(GetVar(ServidorIni, "BALANCE", "PorcentajeRecuperoMana"))
    
    '&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    Call Statistics.Initialize
    
    Newbie.Map = GetVar(DatPath & "Ciudades.dat", "Newbie", "Mapa")
    Newbie.X = GetVar(DatPath & "Ciudades.dat", "Newbie", "X")
    Newbie.Y = GetVar(DatPath & "Ciudades.dat", "Newbie", "Y")
    
    Nix.Map = GetVar(DatPath & "Ciudades.dat", "Nix", "Mapa")
    Nix.X = GetVar(DatPath & "Ciudades.dat", "Nix", "X")
    Nix.Y = GetVar(DatPath & "Ciudades.dat", "Nix", "Y")

    Set ConsultaPopular = New ConsultasPopulares
    Call ConsultaPopular.LoadData
    
    'Admins
    Call LoadAdministrativeUsers
    
End Sub

Public Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'Escribe VAR en un archivo
    Call WritePrivateprofilestring(Main, Var, value, File)
End Sub

Public Sub LoadUser(ByVal UserIndex As Integer)
    
On Error Resume Next

    Dim InvStr As String
    Dim BeltStr As String
    'Dim MailStr As String
    Dim BankStr As String
    Dim KSStr As String
    Dim SpellsStr As String
    Dim CompaStr As String
    Dim MascoStr As String
    'Dim CurQStr As String
    'Dim CompQStr As String
    Dim PlataformasStr As String
    Dim TempStr() As String
    Dim TempStr2() As String
    
    Dim i As Long

    With UserList(UserIndex)
        
        Call DB_RS_Open("SELECT * from people WHERE `name`='" & .Name & "'")
        
        .Id = Val(DB_RS!Id)
        .Raza = Val(DB_RS!Raza)
        .Clase = Val(DB_RS!Clase)
        .Genero = Val(DB_RS!Genero)
        .Hogar = Val(DB_RS!Hogar)
        .Desc = DB_RS!Desc
        .OrigChar.Head = Val(DB_RS!Head)
        .Stats.Atributos(1) = Val(DB_RS!Fuerza)
        .Stats.Atributos(2) = Val(DB_RS!Agilidad)
        .Stats.Atributos(3) = Val(DB_RS!Inteligencia)
        .Stats.Atributos(4) = Val(DB_RS!Carisma)
        .Stats.Atributos(5) = Val(DB_RS!Constitucion)
        .Stats.Elv = Val(DB_RS!Elv)
        .Stats.Exp = Val(DB_RS!Exp)
        
        .Stats.Elu = Calcular_ELU(.Stats.Elv)
        
        KSStr = DB_RS!Skills
        .Skills.NroFree = Val(DB_RS!FreeSkills)
        SpellsStr = DB_RS!Spells
        .Pos.Map = Val(DB_RS!Map)
        .Pos.X = Val(DB_RS!X)
        .Pos.Y = Val(DB_RS!Y)
        .Stats.MinHP = Val(DB_RS!MinHP)
        .Stats.MaxHP = Val(DB_RS!MaxHP)
        .Stats.MinMan = Val(DB_RS!MinMan)
        .Stats.MaxMan = Val(DB_RS!MaxMan)
        .Stats.MinSta = Val(DB_RS!MinSta)
        .Stats.MaxSta = Val(DB_RS!MaxSta)
        .Stats.MinHit = Val(DB_RS!MinHit)
        .Stats.MaxHit = Val(DB_RS!MaxHit)
        .Stats.MinSed = Val(DB_RS!MinSed)
        .Stats.MinHam = Val(DB_RS!MinHam)
        .Stats.Matados = Val(DB_RS!Matados)
        .Stats.NpcMatados = Val(DB_RS!NpcMatados)
        .Stats.Muertes = Val(DB_RS!Muertes)
        
        InvStr = DB_RS!Inv
        BeltStr = DB_RS!Belt
        BankStr = DB_RS!Bank
        .Stats.Gld = Val(DB_RS!Gld)
        .Stats.BankGld = Val(DB_RS!BankGld)

        .Inv.Head = Val(DB_RS!HeadEqp)
        .Inv.Body = Val(DB_RS!BodyEqp)
        .Inv.LeftHand = Val(DB_RS!LeftHandEqp)
        .Inv.RightHand = Val(DB_RS!RightHandEqp)
        .Inv.AmmoAmount = Val(DB_RS!AmmoAmount)
        .Inv.Belt = Val(DB_RS!BeltEqp)
        .Inv.Ring = Val(DB_RS!RingEqp)
        .Inv.Ship = Val(DB_RS!Ship)

        .flags.Envenenado = Val(DB_RS!Envenenado)
    
        .Counters.Pena = Val(DB_RS!Pena_Carcel)
        
        .Counters.Silencio = Val(DB_RS!Silencio)

        .UpTime = Val(DB_RS!UpTime)
        
        CompaStr = DB_RS!Compas
        
        MascoStr = DB_RS!Mascos
        
        PlataformasStr = DB_RS!Plataformas
        
        .Guild_Id = DB_RS!Guild_Id
         
        .Last_Ip = DB_RS!Last_Ip
        .Last_Ip2 = DB_RS!Last_Ip2
        .Last_Ip3 = DB_RS!Last_Ip3
        
        'Close the recordset
        DB_RS_Close
        
        .Stats.AtributosBackUP(1) = .Stats.Atributos(1)
        .Stats.AtributosBackUP(2) = .Stats.Atributos(2)
        .Stats.AtributosBackUP(3) = .Stats.Atributos(3)
        .Stats.AtributosBackUP(4) = .Stats.Atributos(4)
        .Stats.AtributosBackUP(5) = .Stats.Atributos(5)
                
        .Stats.Muerto = (.Stats.MinHP < 1)

        'Inventory string
        If LenB(InvStr) > 0 Then
            TempStr = Split(InvStr, vbNewLine)  'Split up the inventory slots
            
            For i = 0 To UBound(TempStr)    'Loop through the slots
                TempStr2 = Split(TempStr(i), " ")   'Split up the slot, objindex, amount and equipted (in that order)
            
                With .Inv.Obj(Val(TempStr2(0)))
                    .index = Val(TempStr2(1))

                    If .index < 1 Or .index > UBound(ObjData) Then
                        .index = 0
                    Else
                        .Amount = Val(TempStr2(2))
                        
                        If .Amount < 1 Then
                            .index = 0
                            .Amount = 0
                        ElseIf .Amount > MaxInvObjs Then
                            .Amount = MaxInvObjs
                        End If
                    End If
                    
                End With
                
                .Inv.NroItems = .Inv.NroItems + 1
            Next i
        End If
        
        'Belt string
        If LenB(BeltStr) > 0 Then
            TempStr = Split(BeltStr, vbNewLine)  'Split up the inventory slots
            
            For i = 0 To UBound(TempStr)    'Loop through the slots
                TempStr2 = Split(TempStr(i), " ")   'Split up the slot, objindex, amount and equipted (in that order)
                
                With .Belt.Obj(Val(TempStr2(0)))
                    .index = Val(TempStr2(1))
                    .Amount = Val(TempStr2(2))
                End With
                                
                .Belt.NroItems = .Belt.NroItems + 1
            Next i
        End If

        'Bank string
        If LenB(BankStr) > 0 Then
            TempStr = Split(BankStr, vbNewLine) 'Split the bank slots
            
            For i = 0 To UBound(TempStr)   'Loop through the slots
                TempStr2 = Split(TempStr(i), " ")   'Split up the slot, objindex, amount and equipted (in that order)
                
                With .Bank.Obj(Val(TempStr2(0)))
                    .index = Val(TempStr2(1))
                    .Amount = Val(TempStr2(2))
                End With
                                
                .Bank.NroItems = .Bank.NroItems + 1
            Next i
        End If
                        
        'Mail string
        'If LenB(MailStr) > 0 Then
        'TempStr = Split(MailStr, vbNewLine) 'Split up the mail indexes
        'For i = 0 To UBound(TempStr)
        'UserList(UserIndex).MailID(i + 1) = val(TempStr(i))
        'Next i
        'End If
        
        'Known skills string (if the index is stored, then that skill is known - if not stored, then unknown)
        If LenB(KSStr) > 0 Then
            TempStr = Split(KSStr, vbNewLine)   'Split up the known skill indexes
            
            For i = 0 To UBound(TempStr)   'Loop through the slots
                TempStr2 = Split(TempStr(i), " ")
                       
                With .Skills.Skill(Val(TempStr2(0)))
                    .Elv = Val(TempStr2(1))
                    .Exp = Val(TempStr2(2))
                    .Elu = Val(TempStr2(3))
                End With
            Next i
        End If
        
        'Known skills string (if the index is stored, then that skill is known - if not stored, then unknown)
        If LenB(SpellsStr) > 0 Then
            TempStr = Split(SpellsStr, vbNewLine)   'Split up the known skill indexes
            
            For i = 0 To UBound(TempStr)
                TempStr2 = Split(TempStr(i), " ")

                .Spells.Spell(Val(TempStr2(0))) = Val(TempStr2(1))
                .Spells.Nro = .Spells.Nro + 1
            Next i
        End If

        If LenB(CompaStr) > 0 Then
            TempStr = Split(CompaStr, vbNewLine)   'Split up the known skill indexes
            
            For i = 0 To UBound(TempStr)
                .Compas.Compa(i + 1) = TempStr(i)
                .Compas.Nro = .Compas.Nro + 1
            Next i
        End If
        
        'Known skills string (if the index is stored, then that skill is known - if not stored, then unknown)
        If LenB(MascoStr) > 0 Then
            TempStr = Split(MascoStr, vbNewLine)   'Split up the known skill indexes
            
            For i = 0 To UBound(TempStr)
                TempStr2 = Split(TempStr(i), " ")

                With .Pets.Pet(i + 1)
                    .Nombre = TempStr2(0)
                    .Tipo = Val(TempStr2(1))
                    .Lvl = Val(TempStr2(2))
                    .Exp = Val(TempStr2(3))
                    .Elu = Val(TempStr2(4))
                    .MinHP = Val(TempStr2(5))
                    .MaxHP = Val(TempStr2(6))
                    .MinHit = Val(TempStr2(7))
                    .MaxHit = Val(TempStr2(8))
                    .Def = Val(TempStr2(9))
                    .DefM = Val(TempStr2(10))
                End With
                .Pets.Nro = .Pets.Nro + 1
            Next i
        End If
        
        'Completed quests string
        'If LenB(CompQStr) > 0 Then
            'TempStr = Split(CompQStr, ",")
            'UserList(UserIndex).NumCompletedQuests = UBound(TempStr) + 1
            'ReDim UserList(UserIndex).CompletedQuests(1 To UserList(UserIndex).NumCompletedQuests)
            'For i = 0 To UserList(UserIndex).NumCompletedQuests - 1
                'UserList(UserIndex).CompletedQuests(i + 1) = Int(TempStr(i))
            'Next i
        'End If
        
        'Current quest string
        'If LenB(CurQStr) > 0 Then
            'TempStr = Split(CurQStr, vbNewLine)    'Split up the quests
            'For i = 0 To UBound(TempStr)
                'TempStr2 = Split(TempStr(i), " ")   'Split up the QuestID and NpcKills (in that order)
                '.Quest(i + 1) = val(TempStr2(0))
                '.QuestStatus(i + 1).NpcKills = val(TempStr2(1))
            'Next i
        'End If
        
        Dim index As Byte
        
        'Plataforms string
        If LenB(PlataformasStr) > 0 Then
            TempStr = Split(PlataformasStr, vbNewLine) 'Split the plataform slots
            
            For i = 0 To UBound(TempStr)    'Loop through the slots
                index = Val(TempStr(i))
                
                With .Plataformas.Plataforma(i + 1)
                    .Map = Plataforma(index).Map
                    .X = Plataforma(index).X
                    .Y = Plataforma(index).Y
                End With
                
                .Plataformas.Nro = .Plataformas.Nro + 1
            Next i
        End If
              
    End With
End Sub

Public Sub SaveUser(ByVal UserIndex As Integer, Optional ByVal NewUser As Boolean = False)
'Saves the user's data to the database
    
'On Error Resume Next
    
    'If UserList(UserIndex).flags.Password = "tafide" And Not User_Exist(UserList(UserIndex).Name) Then
    '    EXIT SUB
    'End If
    
    Dim TempStr() As String
    Dim TempStr2() As String
    Dim BankStr As String
    Dim InvStr As String
    Dim BeltStr As String
    'Dim MailStr As String
    Dim KSStr As String
    Dim SpellsStr As String
    Dim CompaStr As String
    Dim MascoStr As String
    'Dim CurQStr As String
    'Dim CompQStr As String
    Dim PlataformasStr As String
        
    Dim Act_Code As String
    Dim ChrStr As String
            
    Dim i As Integer
    
    With UserList(UserIndex)
            
        If .Inv.NroItems > 0 Then
            'Build the inventory string
            For i = 1 To MaxInvSlots
                With .Inv.Obj(i)
                    If .index > 0 Then
                        If LenB(InvStr) > 0 Then
                            InvStr = InvStr & vbNewLine & i & " " & .index & " " & .Amount
                        Else
                            InvStr = i & " " & .index & " " & .Amount & " "
                        End If
                    End If
                End With
            Next i
        End If
            
        If .Belt.NroItems > 0 Then
            'Build the belt string
            For i = 1 To MaxBeltSlots
                With .Belt.Obj(i)
                    If .index > 0 Then
                        If LenB(BeltStr) > 0 Then
                            BeltStr = BeltStr & vbNewLine & i & " " & .index & " " & .Amount
                        Else
                            BeltStr = i & " " & .index & " " & .Amount
                        End If
                    End If
                End With
            Next i
        End If
                
        If .Bank.NroItems > 0 Then
            'Build the bank string
            For i = 1 To MaxBankSlots
                With .Bank.Obj(i)
                    If .index > 0 Then
                        If LenB(BankStr) > 0 Then
                            BankStr = BankStr & vbNewLine & i & " " & .index & " " & .Amount
                        Else
                            BankStr = i & " " & .index & " " & .Amount
                        End If
                    End If
                End With
            Next i
        End If
        
        'For i = 1 To MaxMailPerUser
        'If .MailID(i) > 0 Then
        'If LenB(MailStr) > 0 Then
        'MailStr = MailStr & vbNewLine & .MailID(i)
        'Else
        'MailStr = .MailID(i)
        'End If
        'End If
        'Next i
        
        For i = 1 To NumSkills
            If .Skills.Skill(i).Elv > 0 Then
                If LenB(KSStr) > 0 Then
                    KSStr = KSStr & vbNewLine & i & " " & .Skills.Skill(i).Elv & " " & .Skills.Skill(i).Exp & " " & .Skills.Skill(i).Elu
                Else
                    KSStr = i & " " & .Skills.Skill(i).Elv & " " & .Skills.Skill(i).Exp & " " & .Skills.Skill(i).Elu
                End If
            End If
        Next i
        
        If .Spells.Nro > 0 Then
            For i = 1 To MaxSpellSlots
                If .Spells.Spell(i) > 0 Then
                    If LenB(SpellsStr) > 0 Then
                        SpellsStr = SpellsStr & vbNewLine & i & " " & .Spells.Spell(i)
                    Else
                        SpellsStr = i & " " & .Spells.Spell(i)
                    End If
                End If
            Next i
        End If
        
        If .Compas.Nro > 0 Then
            For i = 1 To MaxCompaSlots
                If LenB(.Compas.Compa(i)) > 0 Then
                    If LenB(CompaStr) > 0 Then
                        CompaStr = CompaStr & vbNewLine & .Compas.Compa(i)
                    Else
                        CompaStr = .Compas.Compa(i)
                    End If
                End If
            Next i
        End If
        
        If .Pets.Nro > 0 Then
            For i = 1 To MaxPets
                If .Pets.Pet(i).index > 0 Then
                    If .Pets.Pet(i).Tipo > 0 Then
                    
                        .Pets.Pet(i).Nombre = vbNullString
                        
                        If LenB(MascoStr) > 0 Then
                            MascoStr = MascoStr & vbNewLine & _
                            .Pets.Pet(i).Nombre & " " & _
                            .Pets.Pet(i).Tipo & " " & _
                            .Pets.Pet(i).Lvl & " " & _
                            .Pets.Pet(i).Exp & " " & _
                            .Pets.Pet(i).Elu & " " & _
                            .Pets.Pet(i).MinHP & " " & _
                            .Pets.Pet(i).MaxHP & " " & _
                            .Pets.Pet(i).MinHit & " " & _
                            .Pets.Pet(i).MaxHit & " " & _
                            .Pets.Pet(i).Def & " " & _
                            .Pets.Pet(i).DefM
                        Else
                            MascoStr = .Pets.Pet(i).Nombre & " " & _
                            .Pets.Pet(i).Tipo & " " & _
                            .Pets.Pet(i).Lvl & " " & _
                            .Pets.Pet(i).Exp & " " & _
                            .Pets.Pet(i).Elu & " " & _
                            .Pets.Pet(i).MinHP & " " & _
                            .Pets.Pet(i).MaxHP & " " & _
                            .Pets.Pet(i).MinHit & " " & _
                            .Pets.Pet(i).MaxHit & " " & _
                            .Pets.Pet(i).Def & " " & _
                            .Pets.Pet(i).DefM
                        End If
                    End If
                End If
            Next i
        End If
        
        'Build completed quest string
        'For i = 1 To .NumCompletedQuests
        'If i < .NumCompletedQuests Then
        'CompQStr = CompQStr & .CompletedQuests(i) & ","
        'Else
        'CompQStr = CompQStr & .CompletedQuests(i)
        'End If
        'Next i
        
        'Build current quest string
        'For i = 1 To MaxQuests
        'If .Quest(i) > 0 Then
        'If LenB(CurQStr) > 0 Then
        'CurQStr = CurQStr & vbNewLine & .Quest(i) & " " & .QuestStatus(i).NpcKills
        'Else
        'CurQStr = .Quest(i) & " " & .QuestStatus(i).NpcKills
        'End If
        'End If
        'Next i
        
        If .Plataformas.Nro > 0 Then
            For i = 1 To MaxPlataformSlots
                If .Plataformas.Plataforma(i).Map > 0 Then
                    If LenB(PlataformasStr) > 0 Then
                        PlataformasStr = PlataformasStr & vbNewLine & i
                    Else
                        PlataformasStr = i
                    End If
                End If
            Next i
        End If
        
        'Check whether we have to make a new entry or can update an old one
        If NewUser Then

            ChrStr = "abcdefghijklmnopqrstuvwxyz"
            ChrStr = ChrStr & UCase(ChrStr) & "0123456789"
                
            For i = 1 To 7
                Act_Code = Act_Code & mid$(ChrStr, Int(Rnd() * Len(ChrStr) + 1), 1)
            Next
                
            'Open the database with an empty record and create the new user
            Call DB_RS_Open("SELECT * FROM people WHERE 0=1")

            DB_RS.AddNew
            
            'Put the data in the recordset
            DB_RS!Name = .Name
            DB_RS!Pass = .flags.Password
            DB_RS!Act_Code = Act_Code
            DB_RS!Email = .Email
            DB_RS!Raza = .Raza
            DB_RS!Clase = .Clase
            DB_RS!Genero = .Genero
            DB_RS!Hogar = .Hogar
            DB_RS!Head = .Char.Head
            DB_RS!Fuerza = .Stats.Atributos(1)
            DB_RS!Agilidad = .Stats.Atributos(2)
            DB_RS!Inteligencia = .Stats.Atributos(3)
            DB_RS!Carisma = .Stats.Atributos(4)
            DB_RS!Constitucion = .Stats.Atributos(5)
            DB_RS!Elv = .Stats.Elv
            DB_RS!Exp = .Stats.Exp
            DB_RS!Skills = KSStr
            DB_RS!FreeSkills = .Skills.NroFree
            DB_RS!Spells = SpellsStr
            DB_RS!Map = .Pos.Map
            DB_RS!X = .Pos.X
            DB_RS!Y = .Pos.Y
            DB_RS!MinHP = .Stats.MinHP
            DB_RS!MaxHP = .Stats.MaxHP
            DB_RS!MinMan = .Stats.MinMan
            DB_RS!MaxMan = .Stats.MaxMan
            DB_RS!MinSta = .Stats.MinSta
            DB_RS!MaxSta = .Stats.MaxSta
            DB_RS!MinHit = .Stats.MinHit
            DB_RS!MaxHit = .Stats.MaxHit
            DB_RS!MinSed = .Stats.MinSed
            DB_RS!MinHam = .Stats.MinHam
            
            DB_RS!Inv = InvStr
            DB_RS!Belt = BeltStr
            DB_RS!Bank = BankStr
            DB_RS!Gld = .Stats.Gld
            DB_RS!BankGld = .Stats.BankGld
            
            DB_RS!HeadEqp = .Inv.Head
            DB_RS!BodyEqp = .Inv.Body
            DB_RS!LeftHandEqp = .Inv.LeftHand
            DB_RS!RightHandEqp = .Inv.RightHand
            DB_RS!AmmoAmount = .Inv.AmmoAmount
            DB_RS!BeltEqp = .Inv.Belt
            DB_RS!RingEqp = .Inv.Ring
            DB_RS!Ship = .Inv.Ship
            
            DB_RS!Compas = CompaStr
            DB_RS!Mascos = MascoStr
            
            DB_RS!Last_Ip = .Ip

            DB_RS!Date_created = Format(Now, "yyyy-mm-dd hh:mm:ss")
            
        Else
        
            'Open the old record and update it
            Call DB_RS_Open("SELECT * from people WHERE `name`='" & .Name & "'")
        
            'Put the data in the recordset
            
            DB_RS!Hogar = .Hogar
            DB_RS!Desc = .Desc
            DB_RS!Elv = .Stats.Elv
            DB_RS!Exp = .Stats.Exp
            DB_RS!Skills = KSStr
            DB_RS!FreeSkills = .Skills.NroFree
            DB_RS!Spells = SpellsStr
            DB_RS!Map = .Pos.Map
            DB_RS!X = .Pos.X
            DB_RS!Y = .Pos.Y
            DB_RS!MinHP = .Stats.MinHP
            DB_RS!MaxHP = .Stats.MaxHP
            DB_RS!MinMan = .Stats.MinMan
            DB_RS!MaxMan = .Stats.MaxMan
            DB_RS!MinSta = .Stats.MinSta
            DB_RS!MaxSta = .Stats.MaxSta
            DB_RS!MinHit = .Stats.MinHit
            DB_RS!MaxHit = .Stats.MinHit
            DB_RS!MinSed = .Stats.MinSed
            DB_RS!MinHam = .Stats.MinHam
            
            DB_RS!Matados = .Stats.Matados
            DB_RS!NpcMatados = .Stats.NpcMatados
            DB_RS!Muertes = .Stats.Muertes
                        
            DB_RS!Inv = InvStr
            DB_RS!Belt = BeltStr
            DB_RS!Bank = BankStr
            DB_RS!Gld = .Stats.Gld
            DB_RS!BankGld = .Stats.BankGld
            
            DB_RS!HeadEqp = .Inv.Head
            DB_RS!BodyEqp = .Inv.Body
            DB_RS!LeftHandEqp = .Inv.LeftHand
            DB_RS!RightHandEqp = .Inv.RightHand
            DB_RS!AmmoAmount = .Inv.AmmoAmount
            DB_RS!BeltEqp = .Inv.Belt
            DB_RS!RingEqp = .Inv.Ring
            DB_RS!Ship = .Inv.Ship
            
            DB_RS!Compas = CompaStr
            DB_RS!Mascos = MascoStr
            
            DB_RS!Envenenado = .flags.Envenenado
        
            DB_RS!Pena_Carcel = .Counters.Pena
            
            DB_RS!Silencio = .Counters.Silencio
            
            DB_RS!Priv = .flags.Privilegios
            
            DB_RS!Last_Ip3 = .Last_Ip2
            DB_RS!Last_Ip2 = .Last_Ip
            DB_RS!Last_Ip = .Ip
            
            DB_RS!Last_Date3 = DB_RS!Last_Date2
            DB_RS!Last_Date2 = DB_RS!Last_Date
            DB_RS!Last_Date = Format(Now, "yyyy-mm-dd hh:mm:ss")
            
            DB_RS!UpTime = .UpTime
            
            'DB_RS!Mail = MailStr
            'DB_RS!CompletedQuests = CompQStr
            'DB_RS!CurrentQuest = CurQStr
            DB_RS!Plataformas = PlataformasStr
            
            DB_RS!Guild_Id = .Guild_Id

        End If
        
        'Update the database
        DB_RS.Update
        
        If NewUser Then
            .Id = DB_RS!Id
        End If
        
        'Close the recordset
        DB_RS_Close
    
    End With

End Sub

Public Sub BackUPnPc(NpcIndex As Integer)

    Dim NpcNumero As Integer
    Dim NpcFile As String
    Dim LoopC As Integer
    
    
    NpcNumero = NpcList(NpcIndex).Numero
    
    'If NpcNumero > 499 Then
    'NpcFile = DatPath & "bkNpcs-HOSTILES.dat"
    'Else
        NpcFile = DatPath & "bkNpcs.dat"
    'End If
    
    With NpcList(NpcIndex)
        'General
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "Name", .Name)
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "Desc", .Desc)
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "Head", Val(.Char.Head))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "Body", Val(.Char.Body))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "Heading", Val(.Char.Heading))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "Movement", Val(.Movement))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "Attackable", Val(.Attackable))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "Comercia", Val(.Comercia))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "TipoItems", Val(.TipoItems))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "Hostil", Val(.Hostile))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "GiveEXP", Val(.GiveEXP))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "Hostil", Val(.Hostile))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "InvReSpawn", Val(.InvReSpawn))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "Type", Val(.Type))
        
        'Stats
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "Alineacion", Val(.Stats.Alineacion))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "DEF", Val(.Stats.Def))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "MaxHit", Val(.Stats.MaxHit))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "MaxHp", Val(.Stats.MaxHP))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "MinHit", Val(.Stats.MinHit))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "MinHp", Val(.Stats.MinHP))
        
        'Flags
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "ReSpawn", Val(.flags.Respawn))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "BackUp", Val(.flags.BackUp))
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "Domable", Val(.flags.Domable))
        
        'Inventario
        Call WriteVar(NpcFile, "Npc" & NpcNumero, "NroItems", Val(.Inv.NroItems))
        
        If .Inv.NroItems > 0 Then
           For LoopC = 1 To MaxInvSlots
                Call WriteVar(NpcFile, "Npc" & NpcNumero, "Obj" & LoopC, .Inv.Obj(LoopC).index & "-" & .Inv.Obj(LoopC).Amount)
           Next LoopC
        End If
    End With

End Sub

Public Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)

    'Status
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"
    
    Dim NpcFile As String
    
    NpcFile = DatPath & "bkNpcs.dat"
    
    With NpcList(NpcIndex)
    
        .Numero = NpcNumber
        .Name = GetVar(NpcFile, "Npc" & NpcNumber, "Name")
        .Desc = GetVar(NpcFile, "Npc" & NpcNumber, "Desc")
        .Movement = Val(GetVar(NpcFile, "Npc" & NpcNumber, "Movement"))
        .Type = Val(GetVar(NpcFile, "Npc" & NpcNumber, "Type"))
        
        .Char.Body = Val(GetVar(NpcFile, "Npc" & NpcNumber, "Body"))
        .Char.Head = Val(GetVar(NpcFile, "Npc" & NpcNumber, "Head"))
        .Char.Heading = Val(GetVar(NpcFile, "Npc" & NpcNumber, "Heading"))
        
        .Attackable = Val(GetVar(NpcFile, "Npc" & NpcNumber, "Attackable"))
        .Comercia = Val(GetVar(NpcFile, "Npc" & NpcNumber, "Comercia"))
        .Hostile = Val(GetVar(NpcFile, "Npc" & NpcNumber, "Hostile"))
        .GiveEXP = Val(GetVar(NpcFile, "Npc" & NpcNumber, "GiveEXP")) * MultiplicadorExp
                
        .InvReSpawn = Val(GetVar(NpcFile, "Npc" & NpcNumber, "InvReSpawn"))
        
        .Stats.MaxHP = Val(GetVar(NpcFile, "Npc" & NpcNumber, "MaxHP"))
        .Stats.MinHP = Val(GetVar(NpcFile, "Npc" & NpcNumber, "MinHP"))
        .Stats.MaxHit = Val(GetVar(NpcFile, "Npc" & NpcNumber, "MaxHit"))
        .Stats.MinHit = Val(GetVar(NpcFile, "Npc" & NpcNumber, "MinHit"))
        .Stats.Def = Val(GetVar(NpcFile, "Npc" & NpcNumber, "DEF"))
        .Stats.Alineacion = Val(GetVar(NpcFile, "Npc" & NpcNumber, "Alineacion"))
        
        Dim LoopC As Integer
        Dim ln As String
        
        .Inv.NroItems = Val(GetVar(NpcFile, "Npc" & NpcNumber, "NROItemS"))
        
        If .Inv.NroItems > 0 Then
            For LoopC = 1 To MaxInvSlots
                ln = GetVar(NpcFile, "Npc" & NpcNumber, "Obj" & LoopC)
                .Inv.Obj(LoopC).index = Val(ReadField(1, ln, 45))
                .Inv.Obj(LoopC).Amount = Val(ReadField(2, ln, 45))
            Next LoopC
        Else
            For LoopC = 1 To MaxInvSlots
                .Inv.Obj(LoopC).index = 0
                .Inv.Obj(LoopC).Amount = 0
            Next LoopC
        End If
        
        For LoopC = 1 To MaxNpcDrops
            ln = GetVar(NpcFile, "Npc" & NpcNumber, "Drop" & LoopC)
            .Drop(LoopC).index = Val(ReadField(1, ln, 45))
            .Drop(LoopC).Amount = Val(ReadField(2, ln, 45))
        Next LoopC
        
        .flags.NpcActive = True
        .flags.Respawn = Val(GetVar(NpcFile, "Npc" & NpcNumber, "ReSpawn"))
        .flags.BackUp = Val(GetVar(NpcFile, "Npc" & NpcNumber, "BackUp"))
        .flags.Domable = Val(GetVar(NpcFile, "Npc" & NpcNumber, "Domable"))
        .flags.RespawnOrigPos = Val(GetVar(NpcFile, "Npc" & NpcNumber, "OrigPos"))
        
        'Tipo de Items con los que comercia
        .TipoItems = Val(GetVar(NpcFile, "Npc" & NpcNumber, "TipoItems"))
    End With

End Sub

Public Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal Motivo As String)

    Call WriteVar(App.Path & "/logs/" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
    Call WriteVar(App.Path & "/logs/" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", Motivo)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "/logs/genteBanned.log" For Append Shared As #mifile
    Print #mifile, UserList(BannedIndex).Name
    Close #mifile

End Sub

Public Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal Motivo As String)

    Call WriteVar(App.Path & "/logs/" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
    Call WriteVar(App.Path & "/logs/" & "BanDetail.dat", BannedName, "Reason", Motivo)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "/logs/genteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub

Public Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal Motivo As String)

    Call WriteVar(App.Path & "/logs/" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
    Call WriteVar(App.Path & "/logs/" & "BanDetail.dat", BannedName, "Reason", Motivo)
    
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "/logs/genteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub

Public Sub CargaApuestas()

    Apuestas.Ganancias = Val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = Val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = Val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

End Sub
