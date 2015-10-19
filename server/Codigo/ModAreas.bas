Attribute VB_Name = "ModAreas"
'ModAreas.bas - Module to allow the usage of areas instead of maps.
'Saves a lot of bandwidth.

Option Explicit

Public Type AreaInfo
    AreaPerteneceX As Integer
    AreaPerteneceY As Integer
    
    AreaReciveX As Integer
    AreaReciveY As Integer
    
    MinX As Integer
    MinY As Integer
    
    AreaID As Long
End Type

Public Type ConnGroup
    CountEntrys As Long
    OptValue As Long
    UserEntrys() As Long
End Type

Public Const USER_NUEVO As Byte = 255

'Cuidado:
'¡¡¡LAS AREAS ESTÁN HARDCODEADAS!!!
Private CurDay As Byte
Private CurHour As Byte

Private AreasInfo(1 To 100, 1 To 100) As Byte
Private PosToArea(1 To 100) As Byte

Private AreasRecive(12) As Integer

Public ConnGroups() As ConnGroup

Public Sub InitAreas()

    Dim LoopC As Long
    Dim loopX As Long

'Setup areas...
    For LoopC = 0 To 11
        AreasRecive(LoopC) = (2 ^ LoopC) Or IIf(LoopC > 0, 2 ^ (LoopC - 1), 0) Or IIf(LoopC <> 11, 2 ^ (LoopC + 1), 0)
    Next LoopC
    
    For LoopC = 1 To 100
        PosToArea(LoopC) = LoopC \ 9
    Next LoopC
    
    For LoopC = 1 To 100
        For loopX = 1 To 100
            'Usamos 121 IDs de area para saber si pasasamos de area "más rápido"
            AreasInfo(LoopC, loopX) = (LoopC \ 9 + 1) * (loopX \ 9 + 1)
        Next loopX
    Next LoopC

'Setup AutoOptimizacion de areas
    CurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
    CurHour = Fix(Hour(Time) \ 3) 'A ke parte de la hora pertenece
    
    ReDim ConnGroups(1 To NumMaps) As ConnGroup
    
    For LoopC = 1 To NumMaps
        ConnGroups(LoopC).OptValue = Val(GetVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, CurDay & "-" & CurHour))
        
        If ConnGroups(LoopC).OptValue = 0 Then ConnGroups(LoopC).OptValue = 1
        ReDim ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Long
    Next LoopC
End Sub

Public Sub AreasOptimizacion()
'Es la función de autooptimizacion.... la idea es no mandar redimensionando arrays grandes todo el tiempo

    Dim LoopC As Long
    Dim tCurDay As Byte
    Dim tCurHour As Byte
    Dim EntryValue As Long
    
    If (CurDay <> IIf(Weekday(Date) > 6, 1, 2)) Or (CurHour <> Fix(Hour(Time) \ 3)) Then
        
        tCurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
        tCurHour = Fix(Hour(Time) \ 3) 'A ke parte de la hora pertenece
        
        For LoopC = 1 To NumMaps
            EntryValue = Val(GetVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, CurDay & "-" & CurHour))
            Call WriteVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, CurDay & "-" & CurHour, CInt((EntryValue + ConnGroups(LoopC).OptValue) * 0.5))
            
            ConnGroups(LoopC).OptValue = Val(GetVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, tCurDay & "-" & tCurHour))
            If ConnGroups(LoopC).OptValue = 0 Then
                ConnGroups(LoopC).OptValue = 1
            End If
            
            If ConnGroups(LoopC).OptValue >= MapInfo(LoopC).Poblacion Then
                ReDim Preserve ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Long
            End If
        Next LoopC
        
        CurDay = tCurDay
        CurHour = tCurHour
    End If
End Sub

Public Sub CheckUpdateNeededUser(ByVal UserIndex As Integer, ByVal Head As Byte, Optional ByVal ButIndex As Boolean = False)
    
    If UserList(UserIndex).AreasInfo.AreaID = AreasInfo(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y) Then
        Exit Sub
    End If
    
    Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, x As Long, y As Long
    Dim TempInt As Long, map As Long
    
    With UserList(UserIndex)
        MinX = .AreasInfo.MinX
        MinY = .AreasInfo.MinY
        
        If Head = eHeading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - 9
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.SOUTH Then
            MaxY = MinY + 35
            MinY = MinY + 27
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY - 18)
        
        ElseIf Head = eHeading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - 9
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.EAST Then
            MaxX = MinX + 35
            MinX = MinX + 27
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX - 18)
            .AreasInfo.MinY = CInt(MinY)
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.Pos.y \ 9) - 1) * 9
            MaxY = MinY + 26
            
            MinX = ((.Pos.x \ 9) - 1) * 9
            MaxX = MinX + 26
            
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        End If
        
        If MinY < 1 Then
            MinY = 1
        End If
        
        If MinX < 1 Then
            MinX = 1
        End If
            
        If MaxY > 100 Then
            MaxY = 100
        End If
        
        If MaxX > 100 Then
            MaxX = 100
        End If
        
        map = .Pos.map
        
        'Esto es para ke el cliente elimine lo "fuera de area..."
        Call WriteAreaChanged(UserIndex)
        
        'Actualizamos!!!
        For x = MinX To MaxX
            For y = MinY To MaxY
                
                '<<< User >>>
                If maps(map).mapData(x, y).UserIndex > 0 Then
                    
                    TempInt = maps(map).mapData(x, y).UserIndex
                    
                    If UserIndex <> TempInt Then
                    
                        'If UserList(TempInt).Pos.Map <> Map Or UserList(TempInt).Pos.X <> X Or UserList(TempInt).Pos.Y <> Y Then
                        '    maps(map).mapData( X, Y).UserIndex = 0
                        
                        'Else
                            'Solo avisa al otro cliente si no es un admin invisible
                            If UserList(TempInt).flags.AdminInvisible < 1 Then
                                Call MakeUserChar(False, UserIndex, TempInt, map, x, y)
                                
                                'Si esta navegando, siempre esta visible
                                If Not UserList(TempInt).flags.Navegando Then
                                    'Si el user estaba invisible le avisamos al nuevo cliente de eso
                                    If UserList(TempInt).flags.Invisible > 0 Or UserList(TempInt).flags.Oculto > 0 Then
                                        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
                                            Call WriteSetInvisible(UserIndex, UserList(TempInt).Char.CharIndex, True)
                                        End If
                                    End If
                                End If
                            End If
                                         
                            'Solo avisa al otro cliente si no es un admin invisible
                            If .flags.AdminInvisible < 1 Then
                                Call MakeUserChar(False, TempInt, UserIndex, .Pos.map, .Pos.x, .Pos.y)
                                
                                'Si esta navegando, siempre esta visible
                                If Not .flags.Navegando Then
                                    If .flags.Invisible > 0 Or .flags.Oculto > 0 Then
                                        If UserList(TempInt).flags.Privilegios And PlayerType.User Then
                                            Call WriteSetInvisible(TempInt, .Char.CharIndex, True)
                                        End If
                                    End If
                                End If
                            End If
                    
                            If UserList(TempInt).flags.Paralizado > 0 Or UserList(TempInt).flags.Inmovilizado > 0 Then
                                Call WriteSetParalized(UserIndex, UserList(TempInt).Char.CharIndex, True)
                            End If
                    
                            Call FlushBuffer(TempInt)
                        'End If
                        
                    'ElseIf Head = USER_NUEVO Then
                    '    If Not ButIndex Then
                    '        Call MakeUserChar(False, UserIndex, UserIndex, Map, X, Y)
                    '    End If
                    End If
                Else
                
                    'ENEPECE
                    Dim TempNpc As Integer
                     
                    TempNpc = maps(map).mapData(x, y).NpcIndex
    
                    If TempNpc > 0 Then
                        Call MakeNpcChar(False, TempNpc, map, x, y, UserIndex)
                        If NpcList(TempNpc).flags.Paralizado > 0 Or NpcList(TempNpc).flags.Inmovilizado > 0 Then
                            Call WriteSetParalized(UserIndex, NpcList(TempNpc).Char.CharIndex, True)
                        End If
                    End If
                End If
                
                TempInt = maps(map).mapData(x, y).ObjInfo.index
    
                'ITEM
                If TempInt > 0 Then
                    Dim TempAmount As Long
                    
                    TempAmount = maps(map).mapData(x, y).ObjInfo.Amount
                    
                    If Not ItemEsDeMapa(ObjData(TempInt).Type) Then
                    
                        If ObjData(TempInt).Type = otGuita Then
                            Call WriteObjCreate(UserIndex, ObjData(TempInt).GrhIndex, ObjData(TempInt).Type, x, y, , TempAmount)
                            
                        ElseIf ObjData(TempInt).Type = otPortal Then
                            If maps(map).mapData(x, y).TileExit.map > 0 Then
                                Call WriteObjCreate(UserIndex, ObjData(TempInt).GrhIndex, ObjData(TempInt).Type, x, y, , 10000 + maps(map).mapData(x, y).TileExit.map)
                            End If

                        ElseIf ObjData(TempInt).Type = otAlijo Then
                            Call WriteObjCreate(UserIndex, ObjData(TempInt).GrhIndex, ObjData(TempInt).Type, x, y, ObjData(TempInt).name, 1)

                        ElseIf ObjData(TempInt).Type = otCuerpoMuerto Then
                            If TempAmount > 0 Then
                                If UserList(TempAmount).Stats.Muerto Then
                                    Call WriteObjCreate(UserIndex, ObjData(TempInt).GrhIndex, ObjData(TempInt).Type, x, y, UserList(TempAmount).name)
                                End If
                            End If
                            
                            Call EraseObj(map, x, y, -1)
                            
                        ElseIf ObjData(TempInt).Type = otPuerta Then
                            Call WriteObjCreate(UserIndex, ObjData(TempInt).GrhIndex, ObjData(TempInt).Type, x, y)
                            Call Bloquear(False, UserIndex, x, y, maps(map).mapData(x, y).Blocked)
                            Call Bloquear(False, UserIndex, x - 1, y, maps(map).mapData(x - 1, y).Blocked)
                        
                        ElseIf Not ObjData(TempInt).Agarrable Then
                            Call WriteObjCreate(UserIndex, ObjData(TempInt).GrhIndex, ObjData(TempInt).Type, x, y)
                        
                        Else
                            Call WriteObjCreate(UserIndex, ObjData(TempInt).GrhIndex, ObjData(TempInt).Type, x, y, ObjData(TempInt).name, TempAmount)
                        
                            If maps(map).mapData(x, y).Blocked Then
                                Call Bloquear(False, UserIndex, x, y, maps(map).mapData(x, y).Blocked)
                            End If
                        End If
                    End If
                End If
            
            Next y
        Next x
        
        'Precalculados :P
        TempInt = .Pos.x \ 9
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
        
        TempInt = .Pos.y \ 9
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
        .AreasInfo.AreaID = AreasInfo(.Pos.x, .Pos.y)
    End With
End Sub

Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)
    
    If NpcList(NpcIndex).AreasInfo.AreaID = AreasInfo(NpcList(NpcIndex).Pos.x, NpcList(NpcIndex).Pos.y) Then
        Exit Sub
    End If
    
    Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, x As Long, y As Long
    Dim TempInt As Long
    
    With NpcList(NpcIndex)
        MinX = .AreasInfo.MinX
        MinY = .AreasInfo.MinY
        
        If Head = eHeading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - 9
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.SOUTH Then
            MaxY = MinY + 35
            MinY = MinY + 27
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY - 18)
        
        ElseIf Head = eHeading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - 9
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        
        ElseIf Head = eHeading.EAST Then
            MaxX = MinX + 35
            MinX = MinX + 27
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX - 18)
            .AreasInfo.MinY = CInt(MinY)
        
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.Pos.y \ 9) - 1) * 9
            MaxY = MinY + 26
            
            MinX = ((.Pos.x \ 9) - 1) * 9
            MaxX = MinX + 26
            
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        End If
        
        If MinY < 1 Then
            MinY = 1
        End If
        
        If MinX < 1 Then
            MinX = 1
        End If
        
        If MaxY > 100 Then
            MaxY = 100
        End If
        
        If MaxX > 100 Then
            MaxX = 100
        End If
        
        Dim TempUserIndex As Integer
        
        'Actualizamos!!!
        If MapInfo(.Pos.map).Poblacion > 0 Then
            For x = MinX To MaxX
                For y = MinY To MaxY
                    TempUserIndex = maps(.Pos.map).mapData(x, y).UserIndex
                    If TempUserIndex > 0 Then
                        Call MakeNpcChar(False, NpcIndex, .Pos.map, .Pos.x, .Pos.y, TempUserIndex)
                        
                        If NpcList(NpcIndex).flags.Paralizado > 0 Then
                            Call WriteSetParalized(TempUserIndex, NpcList(NpcIndex).Char.CharIndex, True)
                        End If
                    End If
                Next y
            Next x
        End If
        
        'Precalculados :P
        TempInt = .Pos.x \ 9
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
            
        TempInt = .Pos.y \ 9
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
        .AreasInfo.AreaID = AreasInfo(.Pos.x, .Pos.y)
    End With
End Sub

Public Sub QuitarUser(ByVal UserIndex As Integer, ByVal map As Integer)

    Dim TempVal As Long
    Dim LoopC As Long
    
    'Search for the user
    For LoopC = 1 To ConnGroups(map).CountEntrys
        If ConnGroups(map).UserEntrys(LoopC) = UserIndex Then
            Exit For
        End If
    Next LoopC
    
    'Char not found
    If LoopC > ConnGroups(map).CountEntrys Then
        Exit Sub
    End If
    
    'Remove from old map
    ConnGroups(map).CountEntrys = ConnGroups(map).CountEntrys - 1
    TempVal = ConnGroups(map).CountEntrys
    
    'Move list back
    For LoopC = LoopC To TempVal
        ConnGroups(map).UserEntrys(LoopC) = ConnGroups(map).UserEntrys(LoopC + 1)
    Next LoopC
    
    If TempVal > ConnGroups(map).OptValue Then 'Nescesito Redim?
        ReDim Preserve ConnGroups(map).UserEntrys(1 To TempVal) As Long
    End If
End Sub

Public Sub AgregarUser(ByVal UserIndex As Integer, ByVal map As Integer)

    Dim TempVal As Long
    Dim EsNuevo As Boolean
    Dim i As Long
    
    If Not MapaValido(map) Then
        Exit Sub
    End If
    
    EsNuevo = True
    
    'Prevent adding repeated users
    For i = 1 To ConnGroups(map).CountEntrys
        If ConnGroups(map).UserEntrys(i) = UserIndex Then
            EsNuevo = False
            Exit For
        End If
    Next i
    
    If EsNuevo Then
        'Update map and connection groups data
        ConnGroups(map).CountEntrys = ConnGroups(map).CountEntrys + 1
        TempVal = ConnGroups(map).CountEntrys
        
        If TempVal > ConnGroups(map).OptValue Then 'Nescesito Redim
            ReDim Preserve ConnGroups(map).UserEntrys(1 To TempVal) As Long
        End If
        
        ConnGroups(map).UserEntrys(TempVal) = UserIndex
    End If
    
    'Update user
    UserList(UserIndex).AreasInfo.AreaID = 0
    
    UserList(UserIndex).AreasInfo.AreaPerteneceX = 0
    UserList(UserIndex).AreasInfo.AreaPerteneceY = 0
    UserList(UserIndex).AreasInfo.AreaReciveX = 0
    UserList(UserIndex).AreasInfo.AreaReciveY = 0
    
    Call CheckUpdateNeededUser(UserIndex, USER_NUEVO)
End Sub
