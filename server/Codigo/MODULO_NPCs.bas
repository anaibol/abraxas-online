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
                Call SendData(SendTarget.ToPCArea, UserIndexMatador, Msg_SoundFX(MiNpc.flags.Snd3, MiNpc.Pos.X, MiNpc.Pos.Y))
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
                    Call mdParty.ObtenerExito(UserIndexMatador, MiNpc.flags.ExpCount, MiNpc.Pos.Map, MiNpc.Pos.X, MiNpc.Pos.Y)
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
            Call CrearNpc(MiNpc.Numero, MiNpc.Pos.Map, MiNpc.Orig)
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
        .UserIndex = 0
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
        .Name = vbNullString
        .Type = 0
        .Numero = 0
        .Orig.Map = 0
        .Orig.X = 0
        .Orig.Y = 0
        .PoderAtaque = 0
        .PoderEvasion = 0
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
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
        
        If InMapBounds(.Pos.Map, .Pos.X, .Pos.Y) Then
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
    If LegalPos(Pos.Map, Pos.X, Pos.Y, PuedeAgua) Then
        TestSpawnTrigger = MapData(Pos.X, Pos.Y).Trigger <> eTrigger.BAJOTECHO And _
        MapData(Pos.X, Pos.Y).Trigger <> eTrigger.EnPlataforma And _
        MapData(Pos.X, Pos.Y).Trigger <> eTrigger.POSINVALIDA
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
    
    Dim Map As Integer
    Dim X As Integer
    Dim Y As Integer

    nIndex = OpenNpc(NroNpc) 'Conseguimos un indice
    
    If nIndex > MaxNpcs Then
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
        If InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
            Map = OrigPos.Map
            X = OrigPos.X
            Y = OrigPos.Y
            .Orig = OrigPos
            .Pos = OrigPos
            
        Else
            Pos.Map = mapa
            Altpos.Map = mapa
                    
            Do While Not PosicionValida
                Pos.X = RandomNumber(MinXBorder, MaxXBorder)    'Obtenemos posición al azar en x
                Pos.Y = RandomNumber(MinYBorder, MaxYBorder)    'Obtenemos posición al azar en y
                
                Call ClosestLegalPos(Pos, Newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posición valida mas cercana
                If Newpos.X > 0 And Newpos.Y > 0 Then
                    Altpos.X = Newpos.X
                    Altpos.Y = Newpos.Y     'posición alternativa (para evitar el anti respawn, pero intentando qeu si tenía que ser en el agua, sea en el agua.)
                Else
                    Call ClosestLegalPos(Pos, Newpos, PuedeAgua)
                    If Newpos.X > 0 And Newpos.Y > 0 Then
                        Altpos.X = Newpos.X
                        Altpos.Y = Newpos.Y     'posición alternativa (para evitar el anti respawn)
                    End If
                End If
                'Si X e Y son iguales a 0 significa que no se encontro posición valida
                If LegalPosNpc(Newpos.Map, Newpos.X, Newpos.Y, PuedeAgua) And _
                   Not HayUserCerca(Newpos) And TestSpawnTrigger(Newpos, PuedeAgua) Then
                    'Asignamos las nuevas coordenas solo si son validas
                    .Pos.Map = Newpos.Map
                    .Pos.X = Newpos.X
                    .Pos.Y = Newpos.Y
                    PosicionValida = True
                Else
                    Newpos.X = 0
                    Newpos.Y = 0
                
                End If
    
                'for debug
                Iteraciones = Iteraciones + 1
                If Iteraciones > MaxSPAWNATTEMPS Then
                    If Altpos.X > 0 And Altpos.Y > 0 Then
                        Map = Altpos.Map
                        X = Altpos.X
                        Y = Altpos.Y
                        .Pos.Map = Map
                        .Pos.X = X
                        .Pos.Y = Y
                        Call MakeNpcChar(True, nIndex, Map, X, Y)
                        Exit Sub
                    Else
                        Altpos.X = 50
                        Altpos.Y = 50
                        Call ClosestLegalPos(Altpos, Newpos)
                        If Newpos.X > 0 And Newpos.Y > 0 Then
                            .Pos.Map = Newpos.Map
                            .Pos.X = Newpos.X
                            .Pos.Y = Newpos.Y
                            Call MakeNpcChar(True, nIndex, Newpos.Map, Newpos.X, Newpos.Y)
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
            Map = Newpos.Map
            X = .Pos.X
            Y = .Pos.Y
        End If
    End With
    
    'Crea el Npc
    Call MakeNpcChar(True, nIndex, Map, X, Y)

End Sub

Public Sub MakeNpcChar(ByVal toMap As Boolean, NpcIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal UserIndex As Integer)

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
            CharList(CharIndex).NpcIndex = NpcIndex
        End If
        
        MapData(X, Y).NpcIndex = NpcIndex
            
        plyrefs(CharIndex) = CharIndex
        nPolyRects = nPolyRects + 1
        
        PolyRects(CharIndex).bl.X = X - 25
        PolyRects(CharIndex).bl.Y = Y - 25
        PolyRects(CharIndex).tr.X = X + 25
        PolyRects(CharIndex).tr.Y = Y + 25
        
        tstQuad.CreateTree plyrefs(), nPolyRects, MinXBorder, MinYBorder, MaxXBorder, MaxYBorder
       
    ' To Get a ViewPort
        vp.bl.X = 0
        vp.bl.Y = 0
        vp.tr.X = 500
        vp.tr.Y = 500
        
        tstQuad.ViewPort
        If nQuadOutput > 0 Then
        
            Dim j As Integer
            For j = 0 To nQuadOutput '- 1
                If QuadOutput(j) > 0 Then
                    If CharList(QuadOutput(j)).UserIndex > 0 Then
                        Call UserList(CharList(QuadOutput(j)).UserIndex).outgoingData.WriteASCIIStringFixed(Msg_NpcCharCreate _
                        (.Char.Body, .Char.Head, .Char.Heading, CharIndex, X, Y, .Name, .Lvl, MascoIndex))
                        
                        Call FlushBuffer(CharList(QuadOutput(j)).UserIndex)
                    End If
                End If
            Next
            '
            ' objects have been returned
        Else
            ' nothing was found inside the viewport
        End If
        
'        Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
        
        Exit Sub
    
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
        
            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(Msg_NpcCharCreate _
            (.Char.Body, .Char.Head, .Char.Heading, CharIndex, X, Y, .Name, .Lvl, MascoIndex))
                                                       
            Call FlushBuffer(UserIndex)

        Else
            .Area.AreaID = 0
            .Area.AreaPerteneceX = 0
            .Area.AreaPerteneceY = 0
            .Area.AreaReciveX = 0
            .Area.AreaReciveY = 0
            
            Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
        End If
    End With
End Sub

Private Sub EraseNpcChar(ByVal NpcIndex As Integer)

    If NpcList(NpcIndex).Char.CharIndex > 0 Then
        CharList(NpcList(NpcIndex).Char.CharIndex).CharIndex = 0
        CharList(NpcList(NpcIndex).Char.CharIndex).NpcIndex = 0
    End If
    
    If NpcList(NpcIndex).Char.CharIndex = LastChar Then
        Do Until CharList(LastChar).CharIndex > 0
            LastChar = LastChar - 1
            If LastChar < 2 Then
                Exit Do
            End If
        Loop
    End If
    
    'Quitamos del mapa
    MapData(NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.Y).NpcIndex = 0
    
    'Actualizamos los clientes
    Call SendData(SendTarget.ToNpcArea, NpcIndex, Msg_CharRemove(NpcList(NpcIndex).Char.CharIndex))
    
    'Update la lista npc
    NpcList(NpcIndex).Char.CharIndex = 0

    plyrefs(NpcList(NpcIndex).Char.CharIndex) = 0
    nPolyRects = nPolyRects - 1

    'update NumChars
    NumChars = NumChars - 1

End Sub

Public Sub MoveNpc(ByVal NpcIndex As Integer, ByVal nHeading As eHeading)
    On Error Resume Next
    
    Dim nPos As WorldPos
    Dim UserIndex As Integer
    
    With NpcList(NpcIndex)
        
        nPos = .Pos
        
        Call HeadtoPos(nHeading, nPos)
        
        'Es una posición legal
        If LegalPosNpc(.Pos.Map, nPos.X, nPos.Y, .flags.AguaValida = 1) Then
            
            If .flags.AguaValida = 0 And HayAgua(.Pos.Map, nPos.X, nPos.Y) Then
                Exit Sub
            End If

            If .flags.TierraInvalida = 1 And Not HayAgua(.Pos.Map, nPos.X, nPos.Y) Then
                Exit Sub
            End If
            
            Call SendData(SendTarget.ToNpcArea, NpcIndex, Msg_CharMove(.Char.CharIndex, nPos.X, nPos.Y))

            'Update map and user pos
            MapData(.Pos.X, .Pos.Y).NpcIndex = 0
            .Pos = nPos
            .Char.Heading = nHeading
            MapData(nPos.X, nPos.Y).NpcIndex = NpcIndex
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        
        ElseIf nPos.X > MinXBorder And nPos.X < MaxXBorder And _
        nPos.Y > MinYBorder And nPos.Y < MaxYBorder Then
        
            If MapData(nPos.X, nPos.Y).Blocked Or _
            MapData(nPos.X, nPos.Y).UserIndex > 0 Or _
            MapData(nPos.X, nPos.Y).NpcIndex > 0 Then
            
                If .TargetUser > 0 Or .TargetNpc > 0 Then

                    If .TargetUser > 0 Then
                        If Distancia(.Pos, UserList(.TargetUser).Pos) < 2 Then
                            Call MoveNpc(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                            .PFINFO.PathLenght = 0
                            Exit Sub
                        End If
                    End If
                    Dim tHeading As Byte
                    
                    tHeading = FindDirection(.Pos, UserList(.TargetUser).Pos)
                    Call MoveNpc(NpcIndex, tHeading)
                    Exit Sub
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
    
On Error GoTo errhandler
    
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
            
            Call SendData(SendTarget.ToNpcArea, NpcIndex, Msg_SoundFX(SND_NIVEL, NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.Y))
            
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

errhandler:
    Call LogError("Error en CheckNpcLevel - Error: " & Err.Number & " - Description: " & Err.description)
End Sub

Public Function NextOpenNpc() As Integer
'Call LogTarea("PUBLIC SUB NextOpenNpc")

On Error GoTo errhandler
    Dim LoopC As Long
      
    For LoopC = 1 To MaxNpcs + 1
        If LoopC > MaxNpcs Then
            Exit For
        End If
        
        If Not NpcList(LoopC).flags.NpcActive Then
            Exit For
        End If
    Next LoopC
      
    NextOpenNpc = LoopC
    Exit Function

errhandler:
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
    
    Dim Map As Integer
    Dim X As Integer
    Dim Y As Integer
    
    nIndex = OpenNpc(NpcIndex, Respawn, SubeLvl)  'Conseguimos un indice
    
    If nIndex > MaxNpcs Then
        SpawnNpc = 0
        Exit Function
    End If
    
    PuedeAgua = NpcList(nIndex).flags.AguaValida
    PuedeTierra = Not NpcList(nIndex).flags.TierraInvalida = 1
            
    Call ClosestLegalPos(Pos, Newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posición valida mas cercana
    Call ClosestLegalPos(Pos, Altpos, PuedeAgua)
    'Si X e Y son iguales a 0 significa que no se encontro posición valida
    
    If Newpos.X > 0 And Newpos.Y > 0 Then
        'Asignamos las nuevas coordenas solo si son validas
        NpcList(nIndex).Pos.Map = Newpos.Map
        NpcList(nIndex).Pos.X = Newpos.X
        NpcList(nIndex).Pos.Y = Newpos.Y
        PosicionValida = True
    Else
        If Altpos.X > 0 And Altpos.Y > 0 Then
            NpcList(nIndex).Pos.Map = Altpos.Map
            NpcList(nIndex).Pos.X = Altpos.X
            NpcList(nIndex).Pos.Y = Altpos.Y
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
    Map = Newpos.Map
    X = NpcList(nIndex).Pos.X
    Y = NpcList(nIndex).Pos.Y
    
    'Crea el Npc
    Call MakeNpcChar(True, nIndex, Map, X, Y)
    
    If FX Then
        Call SendData(SendTarget.ToNpcArea, nIndex, Msg_SoundFX(SND_WARP, X, Y))
        Call SendData(SendTarget.ToNpcArea, nIndex, Msg_CreateFX(X, Y, FXIDs.FX_WARP))
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
        OpenNpc = MaxNpcs + 1
        Exit Function
    End If
    
    NpcIndex = NextOpenNpc
    
    If NpcIndex > MaxNpcs Then 'Limite de npcs
        OpenNpc = NpcIndex
        Exit Function
    End If
    
    With NpcList(NpcIndex)
        .Numero = NpcNumber
        .Name = Leer.GetValue("Npc" & NpcNumber, "Name")
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
