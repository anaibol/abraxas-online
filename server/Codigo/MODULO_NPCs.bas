Attribute VB_Name = "NPCs"
Option Explicit

Public Sub QuitarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    Dim i As Byte
    
    For i = 1 To MaxPets
    
        If UserList(UserIndex).Pets.Pet(i).index = NpcIndex Then
          
            With UserList(UserIndex).Pets.Pet(i)
                .index = 0
                .Tipo = 0
                .Lvl = 0
                .Elu = 0
                .Exp = 0
                .MinHP = 0
                .MaxHP = 0
                .MaxHP = 0
                .MaxHP = 0
                .MinHit = 0
                .MaxHit = 0
                .Def = 0
                .DefM = 0
            End With
            
            UserList(UserIndex).Pets.NroALaVez = UserList(UserIndex).Pets.NroALaVez - 1
            
            Exit For
             
        End If
        
    Next i

End Sub

Public Sub QuitarMascotaNpc(ByVal Maestro As Integer)
    NpcList(Maestro).Nro = NpcList(Maestro).Nro - 1
End Sub

Public Sub MuereNpc(ByVal NpcIndex As Integer, Optional ByVal UserIndexMatador As Integer = 0)
'Llamado cuando la vida de un Npc llega a cero.

    Dim MiNpc As Npc
    MiNpc = NpcList(NpcIndex)
    
    Call QuitarNpc(NpcIndex)
    
    If UserIndexMatador > 0 Then 'Lo mato un usuario?
        With UserList(UserIndexMatador)
        
            If MiNpc.flags.Snd3 > 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndexMatador, PrepareMessagePlayWave(MiNpc.flags.Snd3, MiNpc.Pos.x, MiNpc.Pos.y))
            End If
            
            .flags.TargetNpc = 0
            .flags.TargetNpcTipo = eNpcType.Comun
            
            'El user que lo mato tiene Mascotas?
            If .Pets.NroALaVez > 0 Then
                Dim T As Integer
                For T = 1 To MaxPets
                    If .Pets.Pet(T).index > 0 Then
                        If NpcList(.Pets.Pet(T).index).TargetNpc = NpcIndex Then
                            Call FollowAmo(.Pets.Pet(T).index)
                        End If
                    End If
                Next T
            End If
    
            If MiNpc.flags.ExpCount > 0 Then
                If .PartyIndex > 0 Then
                    Call mdParty.ObtenerExito(UserIndexMatador, MiNpc.flags.ExpCount, MiNpc.Pos.map, MiNpc.Pos.x, MiNpc.Pos.y)
                Else
                    Call CalcularDarExp(UserIndexMatador, NpcIndex)
                    .Stats.Exp = .Stats.Exp + MiNpc.flags.ExpCount
                    Call WriteUpdateExp(UserIndexMatador)
                End If
                MiNpc.flags.ExpCount = 0
            End If
        
            .Stats.NpcMatados = .Stats.NpcMatados + 1
        End With
    End If 'UserIndexMatador > 0
   
    If MiNpc.MaestroUser < 1 Then
        'Tiramos el inventario
        Call NpcTirarItems(MiNpc, UserIndexMatador)
        'ReSpawn o no
        If MiNpc.flags.Respawn > 0 Then
            Call CrearNpc(MiNpc.Numero, MiNpc.Pos.map, MiNpc.Orig)
        End If
        
    Else
        Dim i As Byte
        
        'Es Mascota, wardo la info (no me anda la g)
        For i = 1 To MaxPets
            With UserList(MiNpc.MaestroUser)
                If .Pets.Pet(i).index = NpcIndex Then
                    If MiNpc.Contadores.TiempoExistencia = Max_Integer_Value Then
                        .Pets.Pet(i).MinHP = 0
                    Else
                        Call QuitarMascota(MiNpc.MaestroUser, .Pets.Pet(i).index)
                    End If
                    Exit For
                End If
            End With
        Next i
    End If

End Sub

Private Sub ResetNpcFlags(ByVal NpcIndex As Integer)
    'Clear the npc's flags
    
    With NpcList(NpcIndex).flags
        .AfectaParalisis = 0
        .AguaValida = 0
        .BackUp = 0
        .Bendicion = 0
        .Domable = 0
        .Envenenado = 0
        .Faccion = 0
        .Follow = False
        .AtacaDoble = 0
        .LanzaSpells = 0
        .Invisible = 0
        .Maldicion = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Respawn = 0
        .RespawnOrigPos = 0
        .Snd1 = 0
        .Snd2 = 0
        .Snd3 = 0
        .TierraInvalida = 0
    End With
End Sub

Private Sub ResetNpcCounters(ByVal NpcIndex As Integer)
    With NpcList(NpcIndex).Contadores
        .Paralisis = 0
        .TiempoExistencia = 0
    End With
End Sub

Private Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)
    With NpcList(NpcIndex).Char
        .Body = 0
        .HeadAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .Heading = 0
        .Loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Private Sub ResetNpcCriatures(ByVal NpcIndex As Integer)
    Dim j As Long
    
    With NpcList(NpcIndex)
        For j = 1 To .NroCriaturas
            .Criaturas(j).NpcIndex = 0
            .Criaturas(j).NpcName = vbNullString
        Next j
        
        .NroCriaturas = 0
    End With
End Sub

Public Sub ResetExpresiones(ByVal NpcIndex As Integer)
    Dim j As Long
    
    With NpcList(NpcIndex)
        For j = 1 To .NroExpresiones
            .Expresiones(j) = vbNullString
        Next j
        
        .NroExpresiones = 0
    End With
End Sub

Private Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)
    With NpcList(NpcIndex)
        .Attackable = 0
        .CanAttack = 0
        .Comercia = 0
        .GiveEXP = 0
        .Hostile = 0
        .InvReSpawn = 0
                        
        If .MaestroUser > 0 Then
        
            'Guardo la data del pet
            Dim i As Byte
            
            For i = 1 To MaxPets
            
                If UserList(.MaestroUser).Pets.Pet(i).index = NpcIndex Then
              
                    If .Contadores.TiempoExistencia = 0 Then
                        UserList(.MaestroUser).Pets.Pet(i).MinHP = .Stats.MinHP
                    Else
                        UserList(.MaestroUser).Pets.Pet(i).Tipo = 0
                    End If
                    
                    Exit For
                End If
            Next i
            
            .MaestroUser = 0
        End If
            
        If .MaestroNpc > 0 Then
            Call QuitarMascotaNpc(.MaestroNpc)
            .MaestroNpc = 0
        End If
        
        .Nro = 0
        .Movement = 0
        .name = vbNullString
        .Type = 0
        .Numero = 0
        .Orig.map = 0
        .Orig.x = 0
        .Orig.y = 0
        .PoderAtaque = 0
        .PoderEvasion = 0
        .Pos.map = 0
        .Pos.x = 0
        .Pos.y = 0
        .TargetUser = 0
        .TargetNpc = 0
        .TipoItems = 0
        .Veneno = 0
        .Desc = vbNullString
        
        '.PFINFO.PathLenght = 0
        
        Dim j As Long
        For j = 1 To .Spells.Nro
            .Spells.Spell(j) = 0
        Next j

    End With
    
    Call ResetNpcCharInfo(NpcIndex)
    Call ResetNpcCriatures(NpcIndex)
    Call ResetExpresiones(NpcIndex)
End Sub

Public Sub QuitarNpc(ByVal NpcIndex As Integer)

    With NpcList(NpcIndex)
        .flags.NpcActive = False
        .TargetUser = 0
        
        If InMapBounds(.Pos.map, .Pos.x, .Pos.y) Then
            Call EraseNpcChar(NpcIndex)
        End If
    End With
        
    'Nos aseguramos de que el inventario sea removido...
    'asi los lobos no volveran a tirar armaduras ;))
    Call ResetNpcInv(NpcIndex)
    Call ResetNpcFlags(NpcIndex)

    Call ResetNpcMainInfo(NpcIndex)
    
    Call ResetNpcCounters(NpcIndex)
        
    If NpcIndex = LastNpc Then
        Do Until NpcList(LastNpc).flags.NpcActive
            LastNpc = LastNpc - 1
            If LastNpc < 1 Then
                Exit Do
            End If
        Loop
    End If
        
    If numNpcs > 0 Then
        numNpcs = numNpcs - 1
    End If

End Sub

Private Function TestSpawnTrigger(Pos As WorldPos, Optional PuedeAgua As Boolean = False) As Boolean
    If LegalPos(Pos.map, Pos.x, Pos.y, PuedeAgua) Then
        TestSpawnTrigger = maps(Pos.map).mapData(Pos.x, Pos.y).Trigger <> eTrigger.BAJOTECHO And _
        maps(Pos.map).mapData(Pos.x, Pos.y).Trigger <> eTrigger.EnPlataforma And _
        maps(Pos.map).mapData(Pos.x, Pos.y).Trigger <> eTrigger.POSINVALIDA
    End If
End Function

Public Sub CrearNpc(NroNpc As Integer, mapa As Integer, OrigPos As WorldPos)
    
    Dim Pos As WorldPos
    Dim Newpos As WorldPos
    Dim Altpos As WorldPos
    Dim nIndex As Integer
    Dim PosicionValida As Boolean
    Dim Iteraciones As Long
    Dim PuedeAgua As Boolean
    Dim PuedeTierra As Boolean
    
    Dim map As Integer
    Dim x As Byte
    Dim y As Byte

    nIndex = OpenNpc(NroNpc) 'Conseguimos un indice
    
    If nIndex > MaxNpcS Then
        Exit Sub
    End If
    
    With NpcList(nIndex)
        'Los Npc's suben de nivel!
        Dim r As Byte
        
        r = RandomNumber(0, 100)
            
        If r > 25 Then
            .Lvl = 2
            
            .Stats.MaxHP = Round(.Stats.MaxHP + (.Stats.MaxHP * RandomNumber(-5, 5) \ 100))
                
        ElseIf r > 5 Then
            .Lvl = 3
            .GiveEXP = Round(.GiveEXP * 1.75)
            .flags.ExpCount = .GiveEXP
            .Stats.MaxHP = Round(.Stats.MaxHP + (.Stats.MaxHP * RandomNumber(15, 35) \ 100))
            .Stats.MinHP = .Stats.MaxHP
            
            .Stats.MinHit = Round(.Stats.MinHit * 1.2)
            .Stats.MaxHit = Round(.Stats.MaxHit * 1.2)

        Else
            .Lvl = 4
            .GiveEXP = Round(.GiveEXP * 3)
            .flags.ExpCount = .GiveEXP
                .Stats.MaxHP = Round(.Stats.MaxHP + (.Stats.MaxHP * RandomNumber(75, 125) \ 100))
            .Stats.MinHP = .Stats.MaxHP
            
            .Stats.MinHit = Round(.Stats.MinHit * 1.5)
            .Stats.MaxHit = Round(.Stats.MaxHit * 1.5)
        End If
        
        PuedeAgua = .flags.AguaValida
        PuedeTierra = IIf(.flags.TierraInvalida = 1, False, True)
        
        'Necesita ser respawned en un lugar especifico
        If InMapBounds(OrigPos.map, OrigPos.x, OrigPos.y) Then
            
            map = OrigPos.map
            x = OrigPos.x
            y = OrigPos.y
            .Orig = OrigPos
            .Pos = OrigPos
        Else
            
            Pos.map = mapa
            Altpos.map = mapa
                    
            Do While Not PosicionValida
                Pos.x = RandomNumber(MinXBorder, MaxXBorder)    'Obtenemos posición al azar en x
                Pos.y = RandomNumber(MinYBorder, MaxYBorder)    'Obtenemos posición al azar en y
                
                Call ClosestLegalPos(Pos, Newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posición valida mas cercana
                If Newpos.x > 0 And Newpos.y > 0 Then
                    Altpos.x = Newpos.x
                    Altpos.y = Newpos.y     'posición alternativa (para evitar el anti respawn, pero intentando qeu si tenía que ser en el agua, sea en el agua.)
                Else
                    Call ClosestLegalPos(Pos, Newpos, PuedeAgua)
                    If Newpos.x > 0 And Newpos.y > 0 Then
                        Altpos.x = Newpos.x
                        Altpos.y = Newpos.y     'posición alternativa (para evitar el anti respawn)
                    End If
                End If
                'Si X e Y son iguales a 0 significa que no se encontro posición valida
                If LegalPosNpc(Newpos.map, Newpos.x, Newpos.y, PuedeAgua) And _
                   Not HayPCarea(Newpos) And TestSpawnTrigger(Newpos, PuedeAgua) Then
                    'Asignamos las nuevas coordenas solo si son validas
                    .Pos.map = Newpos.map
                    .Pos.x = Newpos.x
                    .Pos.y = Newpos.y
                    PosicionValida = True
                Else
                    Newpos.x = 0
                    Newpos.y = 0
                
                End If
    
                'for debug
                Iteraciones = Iteraciones + 1
                If Iteraciones > MaxSPAWNATTEMPS Then
                    If Altpos.x > 0 And Altpos.y > 0 Then
                        map = Altpos.map
                        x = Altpos.x
                        y = Altpos.y
                        .Pos.map = map
                        .Pos.x = x
                        .Pos.y = y
                        Call MakeNpcChar(True, nIndex, map, x, y)
                        Exit Sub
                    Else
                        Altpos.x = 50
                        Altpos.y = 50
                        Call ClosestLegalPos(Altpos, Newpos)
                        If Newpos.x > 0 And Newpos.y > 0 Then
                            .Pos.map = Newpos.map
                            .Pos.x = Newpos.x
                            .Pos.y = Newpos.y
                            Call MakeNpcChar(True, nIndex, Newpos.map, Newpos.x, Newpos.y)
                            Exit Sub
                        Else
                            Call QuitarNpc(nIndex)
                            Call LogError(MaxSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & mapa & " NroNpc:" & NroNpc)
                            Exit Sub
                        End If
                    End If
                End If
            Loop
            
            'asignamos las nuevas coordenas
            map = Newpos.map
            x = .Pos.x
            y = .Pos.y
        End If
    End With
    
    'Crea el Npc
    Call MakeNpcChar(True, nIndex, map, x, y)

End Sub

Public Sub MakeNpcChar(ByVal toMap As Boolean, NpcIndex As Integer, ByVal map As Integer, ByVal x As Byte, ByVal y As Byte, Optional ByVal UserIndex As Integer)

    If NpcIndex < 1 Then
        Exit Sub
    End If
    
    Dim CharIndex As Integer
    
    Dim i As Byte
    
    Dim MascoIndex As Byte
    
    With NpcList(NpcIndex)
        
        If .Char.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex
            .Char.CharIndex = CharIndex
            CharList(CharIndex) = NpcIndex
        End If
            
        maps(map).mapData(x, y).NpcIndex = NpcIndex
        
        If Not toMap Then
        
            If .MaestroUser > 0 Then
                If UserIndex > 0 Then
                    If UserList(UserIndex).Pets.NroALaVez > 0 Then
                        If .MaestroUser = UserIndex Then
                            For i = 1 To MaxPets
                                If UserList(UserIndex).Pets.Pet(i).index = NpcIndex Then
                                    MascoIndex = i
                                    Exit For
                                End If
                            Next i
                        End If
                    End If
                End If
            End If
        
            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageNpcCharCreate _
            (.Char.Body, .Char.Head, .Char.Heading, .Char.CharIndex, x, y, .name, .Lvl, MascoIndex))
                                                       
            Call FlushBuffer(UserIndex)

        Else
            .AreasInfo.AreaID = 0
            .AreasInfo.AreaPerteneceX = 0
            .AreasInfo.AreaPerteneceY = 0
            .AreasInfo.AreaReciveX = 0
            .AreasInfo.AreaReciveY = 0
            
            Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
        End If
    End With
End Sub

Private Sub EraseNpcChar(ByVal NpcIndex As Integer)

    If NpcList(NpcIndex).Char.CharIndex > 0 Then
        CharList(NpcList(NpcIndex).Char.CharIndex) = 0
    End If
    
    If NpcList(NpcIndex).Char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar < 2 Then
                Exit Do
            End If
        Loop
    End If
    
    'Quitamos del mapa
    maps(NpcList(NpcIndex).Pos.map).mapData(NpcList(NpcIndex).Pos.x, NpcList(NpcIndex).Pos.y).NpcIndex = 0
    
    'Actualizamos los clientes
    Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharRemove(NpcList(NpcIndex).Char.CharIndex))
    
    'Update la lista npc
    NpcList(NpcIndex).Char.CharIndex = 0

    'update NumChars
    NumChars = NumChars - 1

End Sub

Public Sub MoveNpc(ByVal NpcIndex As Integer, ByVal nHeading As eHeading)
    
    Dim nPos As WorldPos
    Dim UserIndex As Integer
    
    With NpcList(NpcIndex)
        
        nPos = .Pos
        
        Call HeadtoPos(nHeading, nPos)
        
        'Es una posición legal
        If LegalPosNpc(.Pos.map, nPos.x, nPos.y, .flags.AguaValida = 1) Then
            
            If .flags.AguaValida = 0 And HayAgua(.Pos.map, nPos.x, nPos.y) Then
                Exit Sub
            End If

            If .flags.TierraInvalida = 1 And Not HayAgua(.Pos.map, nPos.x, nPos.y) Then
                Exit Sub
            End If
            
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharMove(.Char.CharIndex, nPos.x, nPos.y))

            'Update map and user pos
            maps(.Pos.map).mapData(.Pos.x, .Pos.y).NpcIndex = 0
            .Pos = nPos
            .Char.Heading = nHeading
            maps(.Pos.map).mapData(nPos.x, nPos.y).NpcIndex = NpcIndex
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        
        ElseIf nPos.x > MinXBorder And nPos.x < MaxXBorder And _
        nPos.y > MinYBorder And nPos.y < MaxYBorder Then
        
            If maps(.Pos.map).mapData(nPos.x, nPos.y).Blocked Or _
            maps(.Pos.map).mapData(nPos.x, nPos.y).UserIndex > 0 Or _
            maps(.Pos.map).mapData(nPos.x, nPos.y).NpcIndex > 0 Then
            
                If .TargetUser > 0 Or .TargetNpc > 0 Then

                    If .TargetUser > 0 Then
                        If Distancia(.Pos, UserList(.TargetUser).Pos) < 2 Then
                            Call MoveNpc(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                            .PFINFO.PathLenght = 0
                            Exit Sub
                        End If
                    End If
                    
                    Call PathFindingAI(NpcIndex)
                    
                    If .PFINFO.PathLenght > 0 Then
                        Call FollowPath(NpcIndex)
                    'Else
                    '    Call MoveNpc(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                    
                Else
                    Call MoveNpc(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                End If
            End If
        End If
    End With

End Sub

Public Sub CheckNpcLevel(ByVal NpcIndex As Integer)

    'Dim Pts As Integer
    Dim AumentoHIT As Integer
    Dim AumentoHP As Integer
    Dim i As Byte
    
On Error GoTo ErrHandler
    
    With UserList(NpcList(NpcIndex).MaestroUser)
    
        For i = 1 To MaxPets
          If .Pets.Pet(i).index = NpcIndex Then
             Exit For
          End If
        Next i
                            
        While .Pets.Pet(i).Exp >= .Pets.Pet(i).Elu
            'Checkea si alcanzó el máximo nivel
            If .Pets.Pet(i).Lvl >= STAT_MaxELV Then
                .Pets.Pet(i).Exp = 0
                .Pets.Pet(i).Elu = 0
                Exit Sub
            End If
            
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_NIVEL, NpcList(NpcIndex).Pos.x, NpcList(NpcIndex).Pos.y))
            
            .Pets.Pet(i).Lvl = .Pets.Pet(i).Lvl + 1
            
            .Pets.Pet(i).Exp = .Pets.Pet(i).Exp - .Pets.Pet(i).Elu
            
            'Nueva subida de exp x lvl. Pablo (ToxicWaste)
            If .Pets.Pet(i).Lvl < 15 Then
                .Pets.Pet(i).Elu = .Pets.Pet(i).Elu * 1.2
            ElseIf .Pets.Pet(i).Lvl < 25 Then
                .Pets.Pet(i).Elu = .Pets.Pet(i).Elu * 1.25
            ElseIf .Pets.Pet(i).Lvl < 30 Then
                .Pets.Pet(i).Elu = .Pets.Pet(i).Elu * 1.33
            ElseIf .Pets.Pet(i).Lvl < 35 Then
                .Pets.Pet(i).Elu = .Pets.Pet(i).Elu * 1.35
            ElseIf .Pets.Pet(i).Lvl < 40 Then
                .Pets.Pet(i).Elu = .Pets.Pet(i).Elu * 1.4
            ElseIf .Pets.Pet(i).Lvl < 45 Then
                .Pets.Pet(i).Elu = .Pets.Pet(i).Elu * 1.45
            ElseIf .Pets.Pet(i).Lvl < 50 Then
                .Pets.Pet(i).Elu = .Pets.Pet(i).Elu * 1.5
            Else
                .Pets.Pet(i).Elu = .Pets.Pet(i).Elu * 1.6
            End If

            AumentoHIT = 5
            
            'Actualizamos HitPoints
            .Pets.Pet(i).MaxHP = .Pets.Pet(i).MaxHP + AumentoHP
            
            If .Pets.Pet(i).MaxHP > STAT_MaxHP Then
                .Pets.Pet(i).MaxHP = STAT_MaxHP
            End If
                        
            'Actualizamos Golpe Mínimo
            .Pets.Pet(i).MinHit = .Pets.Pet(i).MinHit + AumentoHIT
            
            'Actualizamos Golpe Máximo
            .Pets.Pet(i).MaxHit = .Pets.Pet(i).MaxHit + AumentoHIT
            
            NpcList(NpcIndex).Stats.MaxHP = .Pets.Pet(i).MaxHP
            NpcList(NpcIndex).Stats.MinHP = .Pets.Pet(i).MaxHP
            
            NpcList(NpcIndex).Stats.MinHit = .Pets.Pet(i).MinHit
            NpcList(NpcIndex).Stats.MaxHit = .Pets.Pet(i).MaxHit
        Wend
    End With
Exit Sub

ErrHandler:
    Call LogError("Error en CheckNpcLevel - Error: " & Err.Number & " - Description: " & Err.description)
End Sub

Public Function NextOpenNpc() As Integer
'Call LogTarea("PUBLIC SUB NextOpenNpc")

On Error GoTo ErrHandler
    Dim LoopC As Long
      
    For LoopC = 1 To MaxNpcS + 1
        If LoopC > MaxNpcS Then
            Exit For
        End If
        
        If Not NpcList(LoopC).flags.NpcActive Then
            Exit For
        End If
    Next LoopC
      
    NextOpenNpc = LoopC
    Exit Function

ErrHandler:
    Call LogError("Error en NextOpenNpc")
End Function

Public Sub NpcEnvenenarUser(ByVal UserIndex As Integer)
    
    If (UserList(UserIndex).flags.Privilegios And PlayerType.User) = 0 Then
        Exit Sub
    End If
    
    Dim N As Byte
    
    N = RandomNumber(1, 100)
    
    If N < 21 Then
        UserList(UserIndex).flags.Envenenado = 1
        Call WriteConsoleMsg(UserIndex, "La criatura te envenenó.", FontTypeNames.FONTTYPE_FIGHT)
    End If

End Sub

Public Function SpawnNpc(ByVal NpcIndex As Integer, Pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean, Optional ByVal SubeLvl As Boolean = True) As Integer

    Dim Newpos As WorldPos
    Dim Altpos As WorldPos
    Dim nIndex As Integer
    Dim PosicionValida As Boolean
    Dim PuedeAgua As Boolean
    Dim PuedeTierra As Boolean
    
    Dim map As Integer
    Dim x As Byte
    Dim y As Byte
    
    nIndex = OpenNpc(NpcIndex, Respawn, SubeLvl)  'Conseguimos un indice
    
    If nIndex > MaxNpcS Then
        SpawnNpc = 0
        Exit Function
    End If
    
    PuedeAgua = NpcList(nIndex).flags.AguaValida
    PuedeTierra = Not NpcList(nIndex).flags.TierraInvalida = 1
            
    Call ClosestLegalPos(Pos, Newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posición valida mas cercana
    Call ClosestLegalPos(Pos, Altpos, PuedeAgua)
    'Si X e Y son iguales a 0 significa que no se encontro posición valida
    
    If Newpos.x > 0 And Newpos.y > 0 Then
        'Asignamos las nuevas coordenas solo si son validas
        NpcList(nIndex).Pos.map = Newpos.map
        NpcList(nIndex).Pos.x = Newpos.x
        NpcList(nIndex).Pos.y = Newpos.y
        PosicionValida = True
    Else
        If Altpos.x > 0 And Altpos.y > 0 Then
            NpcList(nIndex).Pos.map = Altpos.map
            NpcList(nIndex).Pos.x = Altpos.x
            NpcList(nIndex).Pos.y = Altpos.y
            PosicionValida = True
        Else
            PosicionValida = False
        End If
    End If
    
    If Not PosicionValida Then
        Call QuitarNpc(nIndex)
        SpawnNpc = 0
        Exit Function
    End If
    
    'asignamos las nuevas coordenas
    map = Newpos.map
    x = NpcList(nIndex).Pos.x
    y = NpcList(nIndex).Pos.y
    
    'Crea el Npc
    Call MakeNpcChar(True, nIndex, map, x, y)
    
    If FX Then
        Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayWave(SND_WARP, x, y))
        Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(x, y, FXIDs.FX_WARP))
    End If
    
    SpawnNpc = nIndex

End Function

Public Function OpenNpc(ByVal NpcNumber As Integer, Optional ByVal Respawn = True, Optional ByVal SubeLvl As Boolean = True) As Integer

    Dim NpcIndex As Integer
    Dim Leer As clsIniManager
    Dim LoopC As Long
    Dim ln As String
    Dim aux As String
    
    Set Leer = LeerNpcs
    
    'If requested Index is invalid, abort
    If Not Leer.KeyExists("Npc" & NpcNumber) Then
        OpenNpc = MaxNpcS + 1
        Exit Function
    End If
    
    NpcIndex = NextOpenNpc
    
    If NpcIndex > MaxNpcS Then 'Limite de npcs
        OpenNpc = NpcIndex
        Exit Function
    End If
    
    With NpcList(NpcIndex)
        .Numero = NpcNumber
        .name = Leer.GetValue("Npc" & NpcNumber, "Name")
        .Desc = Leer.GetValue("Npc" & NpcNumber, "Desc")
        
        .Movement = Val(Leer.GetValue("Npc" & NpcNumber, "Movement"))
        
        .flags.AguaValida = Val(Leer.GetValue("Npc" & NpcNumber, "AguaValida"))
        .flags.TierraInvalida = Val(Leer.GetValue("Npc" & NpcNumber, "TierraInValida"))
        .flags.Faccion = Val(Leer.GetValue("Npc" & NpcNumber, "Faccion"))
        .flags.AtacaDoble = Val(Leer.GetValue("Npc" & NpcNumber, "AtacaDoble"))
        
        .Type = Val(Leer.GetValue("Npc" & NpcNumber, "NpcType"))

        .Char.Body = Val(Leer.GetValue("Npc" & NpcNumber, "Body"))
        .Char.Head = Val(Leer.GetValue("Npc" & NpcNumber, "Head"))
        .Char.Heading = Val(Leer.GetValue("Npc" & NpcNumber, "Heading"))
        
        .Attackable = Val(Leer.GetValue("Npc" & NpcNumber, "Attackable"))
        .Comercia = Val(Leer.GetValue("Npc" & NpcNumber, "Comercia"))
        .Hostile = Val(Leer.GetValue("Npc" & NpcNumber, "Hostile"))
        
        .GiveEXP = Val(Leer.GetValue("Npc" & NpcNumber, "GiveEXP")) * MultiplicadorExp
        
        .flags.ExpCount = .GiveEXP
        
        .Veneno = Val(Leer.GetValue("Npc" & NpcNumber, "Veneno"))
        
        .flags.Domable = Val(Leer.GetValue("Npc" & NpcNumber, "Domable"))
                
        .PoderAtaque = Val(Leer.GetValue("Npc" & NpcNumber, "PoderAtaque"))
        .PoderEvasion = Val(Leer.GetValue("Npc" & NpcNumber, "PoderEvasion"))
        
        .InvReSpawn = Val(Leer.GetValue("Npc" & NpcNumber, "InvReSpawn"))
        
        With .Stats
            .MaxHP = Val(Leer.GetValue("Npc" & NpcNumber, "MaxHP"))
            .MaxHP = .MaxHP * (1 + (0.01 * RandomNumber(-5, 5)))
            .MinHP = .MaxHP
            .MaxHit = Val(Leer.GetValue("Npc" & NpcNumber, "MaxHit"))
            .MinHit = Val(Leer.GetValue("Npc" & NpcNumber, "MinHit"))
            .Def = Val(Leer.GetValue("Npc" & NpcNumber, "DEF"))
            .DefM = Val(Leer.GetValue("Npc" & NpcNumber, "DEFm"))
            .Alineacion = Val(Leer.GetValue("Npc" & NpcNumber, "Alineacion"))
        End With
        
        .Inv.NroItems = Val(Leer.GetValue("Npc" & NpcNumber, "NROItemS"))
        For LoopC = 1 To .Inv.NroItems
            ln = Leer.GetValue("Npc" & NpcNumber, "Obj" & LoopC)
            .Inv.Obj(LoopC).index = Val(ReadField(1, ln, 45))
            .Inv.Obj(LoopC).Amount = Val(ReadField(2, ln, 45))
        Next LoopC
        
        For LoopC = 1 To MaxNpcDrops
            ln = Leer.GetValue("Npc" & NpcNumber, "Drop" & LoopC)
            .Drop(LoopC).index = Val(ReadField(1, ln, 45))
            .Drop(LoopC).Amount = Val(ReadField(2, ln, 45))
        Next LoopC
        
        .flags.LanzaSpells = Val(Leer.GetValue("Npc" & NpcNumber, "LanzaSpells"))
        
        For LoopC = 1 To .flags.LanzaSpells
            .Spells.Spell(LoopC) = Val(Leer.GetValue("Npc" & NpcNumber, "Sp" & LoopC))
            .Spells.Nro = .Spells.Nro + 1
        Next LoopC
                
        If .Type = eNpcType.Entrenador Then
            .NroCriaturas = Val(Leer.GetValue("Npc" & NpcNumber, "NroCriaturas"))
            ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador
            For LoopC = 1 To .NroCriaturas
                .Criaturas(LoopC).NpcIndex = Leer.GetValue("Npc" & NpcNumber, "CI" & LoopC)
                .Criaturas(LoopC).NpcName = Leer.GetValue("Npc" & NpcNumber, "CN" & LoopC)
            Next LoopC
        End If
        
        With .flags
            .NpcActive = True
            
            If Respawn Then
                .Respawn = Val(Leer.GetValue("Npc" & NpcNumber, "ReSpawn"))
            Else
                .Respawn = 0
            End If
            
            .BackUp = Val(Leer.GetValue("Npc" & NpcNumber, "BackUp"))
            .RespawnOrigPos = Val(Leer.GetValue("Npc" & NpcNumber, "OrigPos"))
            .AfectaParalisis = Val(Leer.GetValue("Npc" & NpcNumber, "AfectaParalisis"))
            
            .Snd1 = Val(Leer.GetValue("Npc" & NpcNumber, "Snd1"))
            .Snd2 = Val(Leer.GetValue("Npc" & NpcNumber, "Snd2"))
            .Snd3 = Val(Leer.GetValue("Npc" & NpcNumber, "Snd3"))
        End With
        
        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        .NroExpresiones = Val(Leer.GetValue("Npc" & NpcNumber, "NROEXP"))
        
        If .NroExpresiones > 0 Then
            ReDim .Expresiones(1 To .NroExpresiones) As String
        End If
        
        For LoopC = 1 To .NroExpresiones
            .Expresiones(LoopC) = Leer.GetValue("Npc" & NpcNumber, "Exp" & LoopC)
        Next LoopC
        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        
        'Tipo de Items con los que comercia
        .TipoItems = Val(Leer.GetValue("Npc" & NpcNumber, "TipoItems"))
        
        '.Ciudad = Val(Leer.GetValue("Npc" & NpcNumber, "Ciudad"))
        
        If .Hostile = 1 And SubeLvl Then
            'Los Npc's suben de nivel!
            Dim r As Byte
       
            r = RandomNumber(0, 100)
                                
            If r > 25 Then
                .Lvl = 2
                
                .Stats.MaxHP = Round(.Stats.MaxHP + (.Stats.MaxHP * RandomNumber(-5, 5) \ 100))

            ElseIf r > 5 Then
                .Lvl = 3
                
                .GiveEXP = Round(.GiveEXP * 1.75)
                .flags.ExpCount = .GiveEXP
                
                .Stats.MaxHP = Round(.Stats.MaxHP + (.Stats.MaxHP * RandomNumber(15, 35) \ 100))
                .Stats.MinHP = .Stats.MaxHP
                
                .Stats.MinHit = Round(.Stats.MinHit * 1.25)
                
                If .Stats.MinHit > .Stats.MinHit Then
                    .Stats.MinHit = .Stats.MaxHit
                End If
                
                '.Stats.MaxHit = Round(.Stats.MaxHit * 1.25)
            Else
                .Lvl = 4
                
                .GiveEXP = Round(.GiveEXP * 3)
                .flags.ExpCount = .GiveEXP
                
                .Stats.MaxHP = Round(.Stats.MaxHP + (.Stats.MaxHP * RandomNumber(75, 125) \ 100))
                .Stats.MinHP = .Stats.MaxHP
                
                .Stats.MinHit = Round(.Stats.MinHit * 1.75)
                
                If .Stats.MinHit > .Stats.MinHit Then
                    .Stats.MinHit = .Stats.MaxHit
                End If
                
                '.Stats.MaxHit = Round(.Stats.MaxHit * 1.75)
            End If
        Else
            .Lvl = 1
        End If
        
    End With
    
    'Update contadores de Npcs
    If NpcIndex > LastNpc Then
        LastNpc = NpcIndex
    End If
    numNpcs = numNpcs + 1
    
    'Devuelve el nuevo Indice
    OpenNpc = NpcIndex
End Function

Public Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    With NpcList(NpcIndex)
        If .flags.Follow Then
            .flags.Follow = False
            .TargetUser = 0
        Else
            .flags.Follow = True
            .Movement = TipoAI.NpcDefensa
            .TargetUser = UserIndex
            .TargetNpc = 0
            .Hostile = 0
        End If
    End With
End Sub

Public Sub FollowAmo(ByVal NpcIndex As Integer)
    With NpcList(NpcIndex)
        .flags.Follow = True
        .Movement = TipoAI.SigueAmo
        .TargetUser = .MaestroUser
        .TargetNpc = 0
        .Hostile = 0
    End With
End Sub
