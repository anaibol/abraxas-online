Attribute VB_Name = "ModAreas"
'ModAreas.bas - Module to allow the usage of areas instead of maps.
'Saves a lot of bandwidth.

Option Explicit

Public Type Areas
    AreaPerteneceX As Double
    AreaPerteneceY As Double
    
    AreaReciveX As Double
    AreaReciveY As Double
    
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

Private AreasIds(1 To 500, 1 To 500) As Integer
Private AreasRecive(12) As Integer '12=NumeroDeAreas
Public ConnGroups() As ConnGroup

Public Sub InitAreas()

    Dim i As Byte
    Dim X As Long
    Dim Y As Long

    Dim NumeroDeAreasMenosUno As Byte
    NumeroDeAreasMenosUno = 11
    
'Setup areas...
    For i = 1 To NumeroDeAreasMenosUno
        AreasRecive(i) = (2 ^ i) Or IIf(i > 0, 2 ^ (i - 1), 0) Or IIf(i < NumeroDeAreasMenosUno, 2 ^ (i + 1), 0)
    Next i
        
    For X = MinXBorder To MaxXBorder
        For Y = MinYBorder To MaxYBorder
            'Usamos 121 IDs de area para saber si pasasamos de area "más rápido"
            AreasIds(X, Y) = (X \ 9 + 1) * (Y \ 9 + 1)
        Next Y
    Next X

'Setup AutoOptimizacion de areas
    CurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
    CurHour = Fix(Hour(Time) \ 3) 'A ke parte de la hora pertenece
    
    ReDim ConnGroups(1 To NumMaps) As ConnGroup
    
    For X = 1 To NumMaps
        ConnGroups(X).OptValue = Val(GetVar(DatPath & "AreasStats.dat", "Mapa" & X, CurDay & "-" & CurHour))
        Debug.Print ConnGroups(X).OptValue
        If ConnGroups(X).OptValue = 0 Then ConnGroups(X).OptValue = 1
        ReDim ConnGroups(X).UserEntrys(1 To ConnGroups(X).OptValue) As Long
    Next X
End Sub

Public Sub AreasOptimizacion()
'Es la función de autooptimizacion.... la idea es no mandar redimensionando arrays grandes todo el tiempo
    Exit Sub
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
            
            'If ConnGroups(LoopC).OptValue >= MapInfo(LoopC).Poblacion Then
            '    ReDim Preserve ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Long
            'End If
        Next LoopC
        
        CurDay = tCurDay
        CurHour = tCurHour
    End If
End Sub

Public Sub CheckUpdateNeededUser(ByVal UserIndex As Integer, ByVal Head As Byte, Optional ByVal ButIndex As Boolean = False)
    
    If UserList(UserIndex).Area.AreaID = AreasIds(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then
        Exit Sub
    End If
    
    Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long
    Dim TempInt As Long, Map As Long
    
    With UserList(UserIndex)
        MinX = .Area.MinX
        MinY = .Area.MinY
        
        If Head = eHeading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - 9
            MaxX = MinX + 26
            .Area.MinX = CInt(MinX)
            .Area.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.SOUTH Then
            MaxY = MinY + 35
            MinY = MinY + 27
            MaxX = MinX + 26
            .Area.MinX = CInt(MinX)
            .Area.MinY = CInt(MinY - 18)
        
        ElseIf Head = eHeading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - 9
            MaxY = MinY + 26
            .Area.MinX = CInt(MinX)
            .Area.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.EAST Then
            MaxX = MinX + 35
            MinX = MinX + 27
            MaxY = MinY + 26
            .Area.MinX = CInt(MinX - 18)
            .Area.MinY = CInt(MinY)
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.Pos.Y \ 9) - 1) * 9
            MaxY = MinY + 26
            
            MinX = ((.Pos.X \ 9) - 1) * 9
            MaxX = MinX + 26
            
            .Area.MinX = CInt(MinX)
            .Area.MinY = CInt(MinY)
        End If
        
        If MinX < MinXBorder Then
            MinX = MinXBorder
        End If

        If MaxX > MaxXBorder Then
            MaxX = MaxXBorder
        End If
        
        If MinY < MinYBorder Then
            MinY = MinYBorder
        End If
        
        If MaxY > MaxYBorder Then
            MaxY = MaxYBorder
        End If
        
        Map = .Pos.Map
        
        'Esto es para ke el cliente elimine lo "fuera de area..."
        Call WriteAreaChanged(UserIndex)
        
        'Actualizamos!!!
        For X = MinX To MaxX
            For Y = MinY To MaxY
                
                '<<< User >>>
                If MapData(X, Y).UserIndex > 0 Then
                    
                    TempInt = MapData(X, Y).UserIndex
                    
                    If UserIndex <> TempInt Then
                    
                        'If UserList(TempInt).Pos.Map <> Map Or UserList(TempInt).Pos.X <> X Or UserList(TempInt).Pos.Y <> Y Then
                        '    MapData(x, Y).UserIndex = 0
                        
                        'Else
                            'Solo avisa al otro cliente si no es un admin invisible
                            If UserList(TempInt).flags.AdminInvisible < 1 Then
                                Call MakeUserChar(False, UserIndex, TempInt, Map, X, Y)
                                
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
                                Call MakeUserChar(False, TempInt, UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
                                
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
                     
                    TempNpc = MapData(X, Y).NpcIndex
    
                    If TempNpc > 0 Then
                        Call MakeNpcChar(False, TempNpc, Map, X, Y, UserIndex)
                        If NpcList(TempNpc).flags.Paralizado > 0 Or NpcList(TempNpc).flags.Inmovilizado > 0 Then
                            Call WriteSetParalized(UserIndex, NpcList(TempNpc).Char.CharIndex, True)
                        End If
                    End If
                End If
                
                TempInt = MapData(X, Y).ObjInfo.index
    
                'ITEM
                If TempInt > 0 Then
                    Dim TempAmount As Long
                    
                    TempAmount = MapData(X, Y).ObjInfo.Amount
                    
                    If Not ItemEsDeMapa(ObjData(TempInt).Type) Then
                    
                        If ObjData(TempInt).Type = otGuita Then
                            Call WriteObjCreate(UserIndex, ObjData(TempInt).GrhIndex, ObjData(TempInt).Type, X, Y, , TempAmount)
                            
                        ElseIf ObjData(TempInt).Type = otPortal Then
                            If MapData(X, Y).TileExit.Map > 0 Then
                                Call WriteObjCreate(UserIndex, ObjData(TempInt).GrhIndex, ObjData(TempInt).Type, X, Y, , 10000 + MapData(X, Y).TileExit.Map)
                            End If

                        ElseIf ObjData(TempInt).Type = otAlijo Then
                            Call WriteObjCreate(UserIndex, ObjData(TempInt).GrhIndex, ObjData(TempInt).Type, X, Y, ObjData(TempInt).Name, 1)

                        ElseIf ObjData(TempInt).Type = otCuerpoMuerto Then
                            If TempAmount > 0 Then
                                If UserList(TempAmount).Stats.Muerto Then
                                    Call WriteObjCreate(UserIndex, ObjData(TempInt).GrhIndex, ObjData(TempInt).Type, X, Y, UserList(TempAmount).Name)
                                End If
                            End If
                            
                            Call EraseObj(Map, X, Y, -1)
                            
                        ElseIf ObjData(TempInt).Type = otPuerta Then
                            Call WriteObjCreate(UserIndex, ObjData(TempInt).GrhIndex, ObjData(TempInt).Type, X, Y)
                            Call Bloquear(False, UserIndex, X, Y, MapData(X, Y).Blocked)
                            Call Bloquear(False, UserIndex, X - 1, Y, MapData(X - 1, Y).Blocked)
                        
                        ElseIf Not ObjData(TempInt).Agarrable Then
                            Call WriteObjCreate(UserIndex, ObjData(TempInt).GrhIndex, ObjData(TempInt).Type, X, Y)
                        
                        Else
                            Call WriteObjCreate(UserIndex, ObjData(TempInt).GrhIndex, ObjData(TempInt).Type, X, Y, ObjData(TempInt).Name, TempAmount)
                        
                            If MapData(X, Y).Blocked Then
                                Call Bloquear(False, UserIndex, X, Y, MapData(X, Y).Blocked)
                            End If
                        End If
                    End If
                End If
            
            Next Y
        Next X
        
        'Precalculados :P
        TempInt = .Pos.X \ 9
        .Area.AreaReciveX = AreasRecive(TempInt)
        .Area.AreaPerteneceX = 2 ^ TempInt
        
        TempInt = .Pos.Y \ 9
        .Area.AreaReciveY = AreasRecive(TempInt)
        .Area.AreaPerteneceY = 2 ^ TempInt
        
        .Area.AreaID = AreasIds(.Pos.X, .Pos.Y)
    End With
End Sub

Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)
    
    If NpcList(NpcIndex).Area.AreaID = AreasIds(NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.Y) Then
        Exit Sub
    End If
    
    Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long
    Dim TempInt As Long
    
    With NpcList(NpcIndex)
        MinX = .Area.MinX
        MinY = .Area.MinY
        
        If Head = eHeading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - 9
            MaxX = MinX + 26
            .Area.MinX = CInt(MinX)
            .Area.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.SOUTH Then
            MaxY = MinY + 35
            MinY = MinY + 27
            MaxX = MinX + 26
            .Area.MinX = CInt(MinX)
            .Area.MinY = CInt(MinY - 18)
        
        ElseIf Head = eHeading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - 9
            MaxY = MinY + 26
            .Area.MinX = CInt(MinX)
            .Area.MinY = CInt(MinY)
        
        
        ElseIf Head = eHeading.EAST Then
            MaxX = MinX + 35
            MinX = MinX + 27
            MaxY = MinY + 26
            .Area.MinX = CInt(MinX - 18)
            .Area.MinY = CInt(MinY)
        
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.Pos.Y \ 9) - 1) * 9
            MaxY = MinY + 26
            
            MinX = ((.Pos.X \ 9) - 1) * 9
            MaxX = MinX + 26
            
            .Area.MinX = CInt(MinX)
            .Area.MinY = CInt(MinY)
        End If
        
        If MinY < 1 Then
            MinY = 1
        End If
        
        If MinX < 1 Then
            MinX = 1
        End If
        
        If MaxY > MaxYBorder Then
            MaxY = MaxYBorder
        End If
        
        If MaxX > MaxXBorder Then
            MaxX = MaxXBorder
        End If
        
        Dim TempUserIndex As Integer
        
        For X = MinX To MaxX
            For Y = MinY To MaxY
                TempUserIndex = MapData(X, Y).UserIndex
                If TempUserIndex > 0 Then
                    Call MakeNpcChar(False, NpcIndex, .Pos.Map, .Pos.X, .Pos.Y, TempUserIndex)
                    
                    If NpcList(NpcIndex).flags.Paralizado > 0 Then
                        Call WriteSetParalized(TempUserIndex, NpcList(NpcIndex).Char.CharIndex, True)
                    End If
                End If
            Next Y
        Next X
        
        'Precalculados :P
        TempInt = .Pos.X \ 9
        .Area.AreaReciveX = AreasRecive(TempInt)
        .Area.AreaPerteneceX = 2 ^ TempInt
            
        TempInt = .Pos.Y \ 9
        .Area.AreaReciveY = AreasRecive(TempInt)
        .Area.AreaPerteneceY = 2 ^ TempInt
        
        .Area.AreaID = AreasIds(.Pos.X, .Pos.Y)
    End With
End Sub

Public Sub QuitarUser(ByVal UserIndex As Integer, ByVal Map As Integer)

    Dim TempVal As Long
    Dim LoopC As Long
    
    'Search for the user
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        If ConnGroups(Map).UserEntrys(LoopC) = UserIndex Then
            Exit For
        End If
    Next LoopC
    
    'Char not found
    If LoopC > ConnGroups(Map).CountEntrys Then
        Exit Sub
    End If
    
    'Remove from old map
    ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys - 1
    TempVal = ConnGroups(Map).CountEntrys
    
    'Move list back
    For LoopC = LoopC To TempVal
        ConnGroups(Map).UserEntrys(LoopC) = ConnGroups(Map).UserEntrys(LoopC + 1)
    Next LoopC
    
    If TempVal > ConnGroups(Map).OptValue Then 'Nescesito Redim?
        ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Long
    End If
End Sub

Public Sub AgregarUser(ByVal UserIndex As Integer, ByVal Map As Integer)

    Dim TempVal As Long
    Dim EsNuevo As Boolean
    Dim i As Long
    
    If Not MapaValido(Map) Then
        Exit Sub
    End If
    
    EsNuevo = True
    
    'Prevent adding repeated users
    For i = 1 To ConnGroups(Map).CountEntrys
        If ConnGroups(Map).UserEntrys(i) = UserIndex Then
            EsNuevo = False
            Exit For
        End If
    Next i
    
    If EsNuevo Then
        'Update map and connection groups data
        ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys + 1
        TempVal = ConnGroups(Map).CountEntrys
        
        If TempVal > ConnGroups(Map).OptValue Then 'Nescesito Redim
            ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Long
        End If
        
        ConnGroups(Map).UserEntrys(TempVal) = UserIndex
    End If
    
    'Update user
    UserList(UserIndex).Area.AreaID = 0
    
    UserList(UserIndex).Area.AreaPerteneceX = 0
    UserList(UserIndex).Area.AreaPerteneceY = 0
    UserList(UserIndex).Area.AreaReciveX = 0
    UserList(UserIndex).Area.AreaReciveY = 0
    
    Call CheckUpdateNeededUser(UserIndex, USER_NUEVO)
End Sub
