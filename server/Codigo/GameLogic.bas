Attribute VB_Name = "Extra"
Option Explicit

Public Function EsPrincipiante(ByVal UserIndex As Integer) As Boolean
    EsPrincipiante = UserList(UserIndex).Stats.Elv <= LimitePrincipiante
End Function

Public Function EsGM(ByVal UserIndex As Integer) As Boolean
    EsGM = (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero))
End Function

Public Sub DoTileEvents(ByVal UserIndex As Integer)
'Handles the Map passage of Users. Allows the existance
'of exclusive maps for Principiantes
'and enables GMs to enter every map without restriction.
'Uses: Mapinfo(map).Restringir = "NEWBIE" (principiantes), "NO".
    
    Dim nPos As WorldPos
    Dim FxFlag As Boolean
    Dim TelepRadio As Integer
    Dim DestPos As WorldPos
    
    Dim Map As Integer
    Dim X As Integer
    Dim Y As Integer

    Map = UserList(UserIndex).Pos.Map
    X = UserList(UserIndex).Pos.X
    Y = UserList(UserIndex).Pos.Y

    If InMapBounds(Map, X, Y) Then
        With MapData(X, Y)
            If .ObjInfo.index > 0 Then
                FxFlag = (ObjData(.ObjInfo.index).Type = otPortal)
                TelepRadio = ObjData(.ObjInfo.index).Radio
            End If
            
            If .TileExit.Map > 0 Then
                
                If MapaValido(.TileExit.Map) Then
                    'Es un teleport, entra en una posición random, acorde al radio (si es 0, es pos fija)
                    If FxFlag And TelepRadio > 0 Then
                        DestPos.X = .TileExit.X + RandomNumber(TelepRadio * (-1), TelepRadio)
                        DestPos.Y = .TileExit.Y + RandomNumber(TelepRadio * (-1), TelepRadio)
                    'Posición fija
                    Else
                        DestPos.X = .TileExit.X
                        DestPos.Y = .TileExit.Y
                    End If
                    
                    DestPos.Map = .TileExit.Map
                    
                    If DestPos.Map = 286 Then
                        DestPos.Map = Newbie.Map
                        DestPos.X = Newbie.X
                        DestPos.Y = Newbie.Y
                    End If
                    
                    '¿Es mapa de principiantes?
                    If UCase$(MapInfo(DestPos.Map).restringir) = "NEWBIE" Then
                        '¿El usuario es un principiante?
                        If EsPrincipiante(UserIndex) Or EsGM(UserIndex) Then
                            If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(UserIndex)) Then
                                Call WarpUserChar(UserIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                            Else
                                Call ClosestLegalPos(DestPos, nPos)
                                If nPos.X > 0 And nPos.Y > 0 Then
                                    Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                                End If
                            End If
                        Else 'No es principiante
                            Call WriteConsoleMsg(UserIndex, "Esta zona es solo para principiantes.", FontTypeNames.FONTTYPE_INFO)
                            Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
            
                            If nPos.X > 0 And nPos.Y > 0 Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, False)
                            End If
                        End If
    
                    Else
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)
                            
                            If nPos.X > 0 And nPos.Y > 0 Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    End If
                    
                    'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
                    Dim aN As Integer
                    
                    aN = UserList(UserIndex).flags.AtacadoPorNpc
                    If aN > 0 Then
                       NpcList(aN).TargetUser = 0
                    End If
                
                    aN = UserList(UserIndex).flags.NpcAtacado
                    
                    If aN > 0 Then
                        If NpcList(aN).TargetUser = UserIndex Then
                            NpcList(aN).TargetUser = 0
                        End If
                    End If
                    
                    UserList(UserIndex).flags.AtacadoPorNpc = 0
                    UserList(UserIndex).flags.NpcAtacado = 0
                End If
                
            ElseIf .Trigger = eTrigger.EnPlataforma Then
                If UserList(UserIndex).Counters.EnPlataforma < 1 Then
                    Call AgregarPlataforma(UserIndex, Map)
                    
                    UserList(UserIndex).Counters.EnPlataforma = 1
                    
                    If UserList(UserIndex).Plataformas.Nro > 1 Then
                        Call WriteUserPlatforms(UserIndex)
                    End If
                End If
            
            ElseIf UserList(UserIndex).Counters.EnPlataforma > 0 Then
                UserList(UserIndex).Counters.EnPlataforma = 0
            End If
            
        End With
    End If

End Sub

Public Function InRangoVision(ByVal UserIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

    If X >= UserList(UserIndex).Pos.X - RangoVisionX And X <= UserList(UserIndex).Pos.X + RangoVisionX Then
        If Y >= UserList(UserIndex).Pos.Y - RangoVisionY And Y <= UserList(UserIndex).Pos.Y + RangoVisionY Then
            InRangoVision = True
            Exit Function
        End If
    End If
    
    InRangoVision = False

End Function

Public Function InMapBounds(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    If (Map < 1 Or Map > NumMaps) Or X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        InMapBounds = False
    Else
        InMapBounds = True
    End If
End Function

Private Function RhombLegalPos(ByRef Pos As WorldPos, ByRef vX As Long, ByRef vY As Long, _
                               ByVal Distance As Long, Optional PuedeAgua As Boolean = False, _
                               Optional PuedeTierra As Boolean = True, _
                               Optional ByVal CheckExitTile As Boolean = False) As Boolean
'walks all the perimeter of a rhomb of side  "distance + 1",
'which starts at Pos.x - Distance and Pos.y

    Dim i As Long
    
    vX = Pos.X - Distance
    vY = Pos.Y
    
    For i = 0 To Distance - 1
        If (LegalPos(Pos.Map, vX + i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX + i
            vY = vY - i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    vX = Pos.X
    vY = Pos.Y - Distance
    
    For i = 0 To Distance - 1
        If (LegalPos(Pos.Map, vX + i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX + i
            vY = vY + i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    vX = Pos.X + Distance
    vY = Pos.Y
    
    For i = 0 To Distance - 1
        If (LegalPos(Pos.Map, vX - i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX - i
            vY = vY + i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    vX = Pos.X
    vY = Pos.Y + Distance
    
    For i = 0 To Distance - 1
        If (LegalPos(Pos.Map, vX - i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX - i
            vY = vY - i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    RhombLegalPos = False
    
End Function

Public Function RhombLegalTilePos(ByRef Pos As WorldPos, ByRef vX As Long, ByRef vY As Long, _
                                  ByVal Distance As Long, ByVal ObjIndex As Integer, _
                                  ByVal PuedeAgua As Boolean, ByVal PuedeTierra As Boolean) As Boolean
'walks all the perimeter of a rhomb of side  "distance + 1",
'which starts at Pos.x - Distance and Pos.y
'and searchs for a valid position to drop Items

On Error GoTo errhandler

    Dim i As Long
    Dim HayObj As Boolean
    
    Dim X As Integer
    Dim Y As Integer
    Dim MapObjIndex As Integer
    
    vX = Pos.X - Distance
    vY = Pos.Y
    
    For i = 0 To Distance - 1
        
        X = vX + i
        Y = vY - i
        
        If LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True) Then
           If Not HayObjeto(Pos.Map, X, Y, ObjIndex) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    
    vX = Pos.X
    vY = Pos.Y - Distance
    
    For i = 0 To Distance - 1
        
        X = vX + i
        Y = vY + i
        
        If LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True) Then
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    
    vX = Pos.X + Distance
    vY = Pos.Y
    
    For i = 0 To Distance - 1
        
        X = vX - i
        Y = vY + i
    
        If LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True) Then
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    
    vX = Pos.X
    vY = Pos.Y + Distance
    
    For i = 0 To Distance - 1
        
        X = vX - i
        Y = vY - i
    
        If LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True) Then
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
        
    Exit Function
    
errhandler:
    Call LogError("Error en RhombLegalTilePos. Error: " & Err.Number & " - " & Err.description)
End Function

Public Function HayObjeto(ByVal Map As Byte, ByVal X As Integer, ByVal Y As Integer, _
                          ByVal ObjIndex As Integer) As Boolean
    
    Dim MapObjIndex As Integer
    MapObjIndex = MapData(X, Y).ObjInfo.index
            
    If MapObjIndex > 0 Then
        If MapObjIndex <> ObjIndex Then
            HayObjeto = True
        End If
    End If

End Function

Public Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos, Optional PuedeAgua As Boolean = False, _
                    Optional PuedeTierra As Boolean = True, Optional ByVal CheckExitTile As Boolean = False)
'Encuentra la posición legal mas cercana y la guarda en nPos

    Dim Found As Boolean
    Dim LoopC As Integer
    Dim tX As Long
    Dim tY As Long
    
    nPos = Pos
    tX = Pos.X
    tY = Pos.Y
    
    LoopC = 1
    
    'La primera posición es valida?
    If LegalPos(Pos.Map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra, CheckExitTile) Then
        Found = True
    
    'Busca en las demas posiciones, en forma de "rombo"
    Else
        While (Not Found) And LoopC < 22
            If RhombLegalPos(Pos, tX, tY, LoopC, PuedeAgua, PuedeTierra, CheckExitTile) Then
                nPos.X = tX
                nPos.Y = tY
                Found = True
            End If
        
            LoopC = LoopC + 1
        Wend
        
    End If
    
    If Not Found Then
        nPos.X = 0
        nPos.Y = 0
    End If

End Sub

Private Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
'Encuentra la posición legal mas cercana que no sea un portal y la guarda en nPos

    Dim Notfound As Boolean
    Dim LoopC As Integer
    Dim tX As Long
    Dim tY As Long
    
    nPos.Map = Pos.Map
    
    Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y)
        If LoopC > 12 Then
            Notfound = True
            Exit Do
        End If
        
        For tY = Pos.Y - LoopC To Pos.Y + LoopC
            For tX = Pos.X - LoopC To Pos.X + LoopC
                
                If LegalPos(nPos.Map, tX, tY) And MapData(tX, tY).TileExit.Map = 0 Then
                    nPos.X = tX
                    nPos.Y = tY
                    '¿Hay objeto?
                    
                    tX = Pos.X + LoopC
                    tY = Pos.Y + LoopC
      
                End If
            
            Next tX
        Next tY
        
        LoopC = LoopC + 1
        
    Loop
    
    If Notfound Then
        nPos.X = 0
        nPos.Y = 0
    End If

End Sub

Public Function NameIndex(ByVal Name As String) As Integer
    Dim UserIndex As Integer, i As Integer
     
    If InStrB(Name, "+") > 0 Then
        Name = UCase$(Replace(Name, "+", " "))
    End If
     
    If Len(Name) < 1 Then
        NameIndex = 0
        Exit Function
    End If
     
    UserIndex = 1
    
    If Right$(Name, 1) = "*" Then
        Name = Left$(Name, Len(Name) - 1)
        For i = 1 To LastUser
            If UCase$(UserList(i).Name) = UCase$(Name) Then
                NameIndex = i
                Exit Function
            End If
        Next
    Else
        For i = 1 To LastUser
            If UCase$(Left$(UserList(i).Name, Len(Name))) = UCase$(Name) Then
                NameIndex = i
                Exit Function
            End If
        Next
    End If
 
End Function

Public Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIp As String) As Boolean
    Dim LoopC As Long
    
    For LoopC = 1 To MaxPoblacion
        If UserList(LoopC).flags.Logged Then
            If UserList(LoopC).Ip = UserIp And UserIndex <> LoopC Then
                CheckForSameIP = True
                Exit Function
            End If
        End If
    Next LoopC
    
    CheckForSameIP = False
End Function

Public Function CheckForSameName(ByVal Name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
    Dim LoopC As Long
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.Logged Then
            
            'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
            'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
            'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
            'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ItemS
            'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
            
            If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
                CheckForSameName = True
                Exit Function
            End If
        End If
    Next LoopC
    
    CheckForSameName = False
End Function

Public Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)
'Toma una posición y se mueve hacia donde esta perfilado

    Select Case Head
        Case eHeading.NORTH
            Pos.Y = Pos.Y - 1
        
        Case eHeading.SOUTH
            Pos.Y = Pos.Y + 1
        
        Case eHeading.EAST
            Pos.X = Pos.X + 1
        
        Case eHeading.WEST
            Pos.X = Pos.X - 1
    End Select
End Sub

Public Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal CheckExitTile As Boolean = False) As Boolean
'Checks if the position is Legal.
    
    '¿Es un mapa válido?
    If (Map < 1 Or Map > NumMaps) Or _
       (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
                LegalPos = False
    Else
        With MapData(X, Y)
            If PuedeAgua And PuedeTierra Then
                LegalPos = Not .Blocked And _
                           .UserIndex < 1 And _
                           .NpcIndex < 1
            
            ElseIf PuedeTierra And Not PuedeAgua Then
                LegalPos = Not .Blocked And _
                           .UserIndex < 1 And _
                           .NpcIndex < 1 And _
                           (Not HayAgua(Map, X, Y))

            ElseIf PuedeAgua And Not PuedeTierra Then
                LegalPos = Not .Blocked And _
                           .UserIndex < 1 And _
                           .NpcIndex < 1 And _
                           (HayAgua(Map, X, Y))
            Else
                LegalPos = False
            End If
        End With
        
        If CheckExitTile Then
            LegalPos = LegalPos And (MapData(X, Y).TileExit.Map = 0)
        End If
        
    End If

End Function

Public Function MoveToLegalPos(ByVal UserMoving As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True) As Boolean

    Dim UserIndex As Integer
    Dim IsDeadChar As Boolean
    Dim IsAdminInvisible As Boolean
    
    '¿Es un mapa válido?
    If Map < 1 Or Map > NumMaps Or X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        MoveToLegalPos = False
    Else
        UserIndex = MapData(X, Y).UserIndex
        
        If UserIndex > 0 Then
            IsDeadChar = UserList(UserIndex).Stats.Muerto
            IsAdminInvisible = (UserList(UserIndex).flags.AdminInvisible > 0)
        Else
            IsDeadChar = False
            IsAdminInvisible = False
        End If
            
        If EsGM(UserMoving) Or UserList(UserMoving).Stats.Muerto Then
            MoveToLegalPos = (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                       (MapData(X, Y).NpcIndex = 0)
        ElseIf PuedeAgua And PuedeTierra Then
            MoveToLegalPos = Not MapData(X, Y).Blocked And _
                       (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                       (MapData(X, Y).NpcIndex = 0)
        ElseIf PuedeTierra And Not PuedeAgua Then
            MoveToLegalPos = Not MapData(X, Y).Blocked And _
                       (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                       (MapData(X, Y).NpcIndex = 0) And _
                       (Not HayAgua(Map, X, Y))
        ElseIf PuedeAgua And Not PuedeTierra Then
            MoveToLegalPos = Not MapData(X, Y).Blocked And _
                       (MapData(X, Y).NpcIndex = 0) And _
                       (HayAgua(Map, X, Y))
                        'ESTO O ALGO ACA PARECE NO TERMINADO
        Else
            MoveToLegalPos = False
        End If
    End If
End Function

Public Sub FindLegalPos(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'Search for a Legal pos for the user who is being teleported.

    If MapData(X, Y).UserIndex > 0 Or _
        MapData(X, Y).NpcIndex > 0 Then
                    
        'Se teletransporta a la misma pos a la que estaba
        If MapData(X, Y).UserIndex = UserIndex Then
           Exit Sub
        End If
        
        Dim FoundPlace As Boolean
        Dim tX As Integer
        Dim tY As Integer
        Dim Rango As Byte
        Dim OtherUserIndex As Integer
    
        For Rango = 1 To 5
            For tY = Y - Rango To Y + Rango
                For tX = X - Rango To X + Rango
                    'Reviso que no haya User ni Npc
                    If MapData(tX, tY).UserIndex = 0 Then
                        If MapData(tX, tY).NpcIndex = 0 Then
                            If MapData(tX, tY).Trigger <> eTrigger.EnPlataforma Then
                                If InMapBounds(Map, tX, tY) Then
                                    FoundPlace = True
                                End If
                                
                                Exit For
                            End If
                        End If
                    End If

                Next tX
        
                If FoundPlace Then _
                    Exit For
            Next tY
            
            If FoundPlace Then _
                    Exit For
        Next Rango

        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            X = tX
            Y = tY
        Else
            'Muy poco probable, pero..
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un Npc, lo pisamos.
            OtherUserIndex = MapData(X, Y).UserIndex
            If OtherUserIndex > 0 Then
                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If UserList(OtherUserIndex).ComUsu.DestUsu > 0 Then
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(OtherUserIndex).ComUsu.DestUsu).flags.Logged Then
                        Call FinComerciarUsu(UserList(OtherUserIndex).ComUsu.DestUsu)
                        Call WriteConsoleMsg(UserList(OtherUserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                        Call FlushBuffer(UserList(OtherUserIndex).ComUsu.DestUsu)
                    End If
                    'Lo sacamos.
                    If UserList(OtherUserIndex).flags.Logged Then
                        Call FinComerciarUsu(OtherUserIndex)
                        Call WriteErrorMsg(OtherUserIndex, "Alguien se conectó donde estabas, por favor reconectáte.")
                        Call FlushBuffer(OtherUserIndex)
                    End If
                End If
            
                Call CloseSocket(OtherUserIndex)
            End If
        End If
    End If

End Sub

Public Function LegalPosNpc(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte, Optional ByVal IsPet As Boolean = False) As Boolean

    Dim IsDeadChar As Boolean
    Dim UserIndex As Integer
    Dim IsAdminInvisible As Boolean
        
    If (Map < 1 Or Map > NumMaps) Or _
        (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPosNpc = False
        Exit Function
    End If

    With MapData(X, Y)
        UserIndex = .UserIndex
        If UserIndex > 0 Then
            IsDeadChar = UserList(UserIndex).Stats.Muerto
            IsAdminInvisible = (UserList(UserIndex).flags.AdminInvisible > 0)
        Else
            IsDeadChar = False
            IsAdminInvisible = False
        End If
    
        If AguaValida = 0 Then
            LegalPosNpc = Not .Blocked And _
            (.UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
            .NpcIndex = 0 And _
            (.Trigger <> eTrigger.POSINVALIDA And .Trigger <> eTrigger.EnPlataforma)
            'And Not HayAgua(Map, X, Y)
        Else
            LegalPosNpc = Not .Blocked And _
            (.UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
            .NpcIndex = 0 And _
            (.Trigger <> eTrigger.POSINVALIDA And .Trigger <> eTrigger.EnPlataforma)
        End If
    End With
End Function

Public Sub SendHelp(ByVal index As Integer)
    Dim NumHelpLines As Integer
    Dim LoopC As Integer
    
    NumHelpLines = Val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))
    
    For LoopC = 1 To NumHelpLines
        Call WriteConsoleMsg(index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), FontTypeNames.FONTTYPE_INFO)
    Next LoopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    If NpcList(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, NpcList(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, UserIndex, Msg_ChatOverHead(NpcList(NpcIndex).Expresiones(randomi), NpcList(NpcIndex).Char.CharIndex, vbWhite))
    End If
End Sub

Public Sub LookatTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal UsingSkill As Boolean = False, Optional ByVal RightClick As Boolean = False)

    'Responde al click del usuario sobre el mapa
    Dim FoundChar As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex As Integer
    Dim Stat As String
    Dim ft As FontTypeNames
    
    With UserList(UserIndex)
        '¿Rango Visión? (ToxicWaste)
        If (Abs(.Pos.Y - Y) > RangoVisionY) Or (Abs(.Pos.X - X) > RangoVisionX) Then
            Exit Sub
        End If
        
        If .flags.Comerciando Then
            Exit Sub
        End If
    
        '¿Posicion valida?
        If InMapBounds(Map, X, Y) Then
        
            .flags.TargetMap = Map
            .flags.TargetX = X
            .flags.TargetY = Y
            
            '¿Es un obj?
            If MapData(X, Y).ObjInfo.index > 0 Then
            
                .flags.TargetObjMap = Map
                .flags.TargetObjX = X
                .flags.TargetObjY = Y
                
                FoundSomething = 1
    
                Select Case ObjData(MapData(X, Y).ObjInfo.index).Type
                    Case otPuerta 'Es una puerta
                        Call AccionParaPuerta(Map, X, Y, UserIndex)
                    
                    Case otLeña    'Leña
                        If MapData(X, Y).ObjInfo.index = FOGATA_APAG And Not .Stats.Muerto Then
                            Call AccionParaRamita(Map, X, Y, UserIndex)
                        End If
                    
                    Case otCartel    'Cartel
                        If Len(ObjData(MapData(X, Y).ObjInfo.index).texto) > 0 Then
                            Call WriteShowSignal(UserIndex, MapData(X, Y).ObjInfo.index)
                        End If
                        
                    Case otAlijo
                        If .flags.Comerciando Then
                            Exit Sub
                        End If
                        
                        Dim Pos As WorldPos
                        
                        Pos.Map = .flags.TargetObjMap
                        Pos.X = .flags.TargetObjX
                        Pos.Y = .flags.TargetObjY
                            
                        If Distancia(Pos, .Pos) > 5 Then
                            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del alijo.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                            
                        Call WriteBank(UserIndex)
                        UserList(UserIndex).flags.Comerciando = True
                    
                        Call WriteObjCreate(UserIndex, ObjData(1055).GrhIndex, ObjData(1055).Type, Pos.X, Pos.Y, ObjData(1055).Name, 1)
                End Select
                
            ElseIf MapData(X + 1, Y).ObjInfo.index > 0 Then
                
                If ObjData(MapData(X + 1, Y).ObjInfo.index).Type = otPuerta Then
                    .flags.TargetObjMap = Map
                    .flags.TargetObjX = X + 1
                    .flags.TargetObjY = Y
                    
                    FoundSomething = 1
                    
                    Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
                End If
                
            ElseIf MapData(X + 1, Y + 1).ObjInfo.index > 0 Then
            
                If ObjData(MapData(X + 1, Y + 1).ObjInfo.index).Type = otPuerta Then
                    .flags.TargetObjMap = Map
                    .flags.TargetObjX = X + 1
                    .flags.TargetObjY = Y + 1
                    
                    FoundSomething = 1
                    
                    Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)
                End If
                
            ElseIf MapData(X, Y + 1).ObjInfo.index > 0 Then
            
                .flags.TargetObjMap = Map
                .flags.TargetObjX = X
                .flags.TargetObjY = Y + 1
                
                FoundSomething = 1
                
                If ObjData(MapData(X, Y + 1).ObjInfo.index).Type = otPuerta Then
                    Call AccionParaPuerta(Map, X, Y + 1, UserIndex)
                End If
            End If
            
            If FoundSomething = 1 Then
                .flags.TargetObjIndex = MapData(.flags.TargetObjX, .flags.TargetObjY).ObjInfo.index
            End If
            
            '¿Es un personaje?
            If Y + 1 <= YMaxMapSize Then
                If MapData(X, Y + 1).UserIndex > 0 Then
                    TempCharIndex = MapData(X, Y + 1).UserIndex
                    FoundChar = 1
                ElseIf MapData(X, Y + 1).NpcIndex > 0 Then
                    TempCharIndex = MapData(X, Y + 1).NpcIndex
                    FoundChar = 2
                End If
            End If
            
            '¿Es un personaje?
            If FoundChar = 0 Then
                If MapData(X, Y).UserIndex > 0 Then
                    TempCharIndex = MapData(X, Y).UserIndex
                    FoundChar = 1
                End If
                If MapData(X, Y).NpcIndex > 0 Then
                    TempCharIndex = MapData(X, Y).NpcIndex
                    FoundChar = 2
                End If
            End If
            
            'Reaccion al personaje
            If FoundChar = 1 Then '¿Encontro un Usuario?
            
                If UserList(TempCharIndex).flags.AdminInvisible < 1 Or .flags.Privilegios And PlayerType.Dios Then
                
                    .flags.TargetUser = TempCharIndex
                
                    FoundSomething = 1
                                        
                    If Not UsingSkill Then
                    
                        If RightClick Then
                            If .flags.SelectedChar > 0 Then
                                If PuedeAtacar(UserIndex, .flags.TargetUser) Then
                                    NpcList(.flags.SelectedChar).TargetUser = .flags.TargetUser
                                    NpcList(.flags.SelectedChar).TargetNpc = 0
                                    NpcList(.flags.SelectedChar).Movement = TipoAI.NpcDefensa
                                    NpcList(.flags.SelectedChar).Hostile = 1
                                End If
                            End If
                        
                        Else
                            If .flags.SelectedChar > 0 Then
                                .flags.SelectedChar = 0
                            End If
                            
                            With UserList(TempCharIndex)
                                
                                If LenB(.DescRM) = 0 And .ShowName Then 'No tiene descRM y quiere que se vea su nombre.
                                    
                                    Stat = .Name
                                    
                                    If EsPrincipiante(TempCharIndex) Then
                                        Stat = " <Principiante>"
                                    End If
                                                          
                                    If .Guild_Id > 0 Then
                                        Stat = Stat & " <" & modGuilds.GuildName(.Guild_Id) & ">"
                                    End If
                            
                                    If Len(.Desc) > 0 Then
                                        Stat = .Name & Stat & " - " & .Desc
                                    Else
                                        Stat = .Name & Stat
                                    End If
                                            
                                    If Not .flags.Privilegios And PlayerType.User Then
                                        Stat = Stat & " (Administrador)"
        
                                        'Elijo el color segun el rango del GM:
                                        'Dios
                                        If .flags.Privilegios = PlayerType.Dios Then
                                            ft = FontTypeNames.FONTTYPE_GM
                                        End If
                                    
                                    Else
                                        ft = FontTypeNames.FONTTYPE_TALK
                                    End If
                                      
                                    If .Counters.Silencio > 0 Then
                                       Stat = Stat & " <Silenciado>"
                                    End If
                                    
                                    If .flags.Oculto > 0 Then
                                       Stat = Stat & " <Oculto>"
                                    ElseIf .flags.Invisible > 0 Then
                                       Stat = Stat & " <Invisible>"
                                    End If
                                    
                                    If .flags.Inmovilizado > 0 Then
                                       Stat = Stat & " <Inmovilizado>"
                                    ElseIf .flags.Paralizado > 0 Then
                                       Stat = Stat & " <Paralizado>"
                                    End If
                                      
                                    If .flags.Envenenado > 0 Then
                                       Stat = Stat & " <Envenenado>"
                                    End If
                                
                                Else  'Si tiene descRM la muestro siempre.
                                    Stat = .DescRM
                                    ft = FontTypeNames.FONTTYPE_INFOBOLD
                                End If
                                
                            End With
                        
                            If LenB(Stat) > 0 Then
                                'Call WriteConsoleMsg(UserIndex, Stat, ft)
                            End If
                        End If
                    End If
                    
                    UserList(UserIndex).flags.TargetNpc = 0
                    UserList(UserIndex).flags.TargetNpcTipo = eNpcType.Comun
                    UserList(UserIndex).flags.SelectedChar = 0
                End If
            
            ElseIf FoundChar = 2 Then '¿Encontro un Npc?
                
                .flags.TargetNpcTipo = NpcList(TempCharIndex).Type
                .flags.TargetNpc = TempCharIndex
                '.flags.TargetUser = 0
                '.flags.TargetObjIndex = 0
                
                FoundSomething = 1
                
                If Not UsingSkill Then
                
                    If Len(NpcList(TempCharIndex).Desc) > 1 And NpcList(TempCharIndex).Comercia = 0 Then
                        If NpcList(TempCharIndex).Type <> eNpcType.Entrenador Then
                            Call WriteChatOverHead(UserIndex, NpcList(TempCharIndex).Desc, NpcList(TempCharIndex).Char.CharIndex, vbWhite)
                        End If

                    ElseIf NpcList(TempCharIndex).MaestroUser > 0 Then
                        If NpcList(TempCharIndex).Contadores.TiempoExistencia > 0 Then
                            If Not RightClick Then
                                If CInt(NpcList(TempCharIndex).Contadores.TiempoExistencia / frmMain.GameTimer.interval) > 0 Then
                                    If NpcList(TempCharIndex).MaestroUser = UserIndex Then
                                        Call WriteConsoleMsg(UserIndex, NpcList(TempCharIndex).Name & " (" & CInt(NpcList(TempCharIndex).Contadores.TiempoExistencia / frmMain.GameTimer.interval) & "/" & CInt(IntervaloInvocacion / frmMain.GameTimer.interval) & ")", FontTypeNames.FONTTYPE_INFO)
                                        
                                        If .flags.SelectedChar <> TempCharIndex Then
                                            .flags.SelectedChar = TempCharIndex
                                        Else
                                            .flags.SelectedChar = 0
                                        End If
                                    Else
                                        Call WriteConsoleMsg(UserIndex, NpcList(TempCharIndex).Name & ", invocado por " & UserList(NpcList(TempCharIndex).MaestroUser).Name, FontTypeNames.FONTTYPE_INFO)
                                        
                                        If .flags.SelectedChar > 0 Then
                                            .flags.SelectedChar = 0
                                        End If
                                    End If
                                End If
                            End If
                            
                        ElseIf NpcList(TempCharIndex).MaestroUser = UserIndex Then
                            If Not RightClick Then
                                Call WriteConsoleMsg(UserIndex, NpcList(TempCharIndex).Name & " (" & NpcList(TempCharIndex).Stats.MinHP & "/" & NpcList(TempCharIndex).Stats.MaxHP & ")", FontTypeNames.FONTTYPE_INFO)
                                
                                If .flags.SelectedChar <> TempCharIndex Then
                                    .flags.SelectedChar = TempCharIndex
                                End If
                            End If

                        Else
                        
                            If Not RightClick Then
                                Call WriteConsoleMsg(UserIndex, NpcList(TempCharIndex).Name & ", mascota de " & UserList(NpcList(TempCharIndex).MaestroUser).Name & ".", FontTypeNames.FONTTYPE_INFO)
                                
                                If .flags.SelectedChar > 0 Then
                                    .flags.SelectedChar = 0
                                End If
                            End If
                        End If
    
                    ElseIf RightClick Then
                    
                        If .flags.SelectedChar > 0 Then
                            If NpcList(.flags.SelectedChar).MaestroUser > 0 Then
                                If PuedeAtacarNpc(UserIndex, .flags.TargetNpc) Then
                                    NpcList(.flags.SelectedChar).TargetNpc = .flags.TargetNpc
                                    NpcList(.flags.SelectedChar).TargetUser = 0
                                End If
                            End If
                        End If
                      
                    ElseIf NpcList(TempCharIndex).Hostile = 1 Then
                    
                        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                            If NpcList(TempCharIndex).TargetUser > 0 Then
                                Call WriteConsoleMsg(UserIndex, NpcList(TempCharIndex).Name & " (" & NpcList(TempCharIndex).Stats.MinHP & "/" & NpcList(TempCharIndex).Stats.MaxHP & ") - Le pegó primero: " & UserList(NpcList(TempCharIndex).TargetUser).Name & ".", FontTypeNames.FONTTYPE_INFO)
                            Else
                                Call WriteConsoleMsg(UserIndex, NpcList(TempCharIndex).Name & " (" & NpcList(TempCharIndex).Stats.MinHP & "/" & NpcList(TempCharIndex).Stats.MaxHP & ")", FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                            
                    'COMERCIANTE?
                    ElseIf NpcList(TempCharIndex).Comercia > 0 Then
                    
                        If .flags.Comerciando Then
                            Exit Sub
                        End If
                            
                        If Distancia(NpcList(TempCharIndex).Pos, .Pos) > 5 Then
                            Call WriteChatOverHead(UserIndex, "Estás demasiado lejos.", NpcList(TempCharIndex).Char.CharIndex, vbWhite)
                            Exit Sub
                        End If
                            
                        Call WriteNpcInventory(UserIndex)
                        
                        UserList(UserIndex).flags.Comerciando = True
                        
                    'ENTRENADOR?
                    ElseIf NpcList(TempCharIndex).Type = eNpcType.Entrenador Then
                    
                        'Make sure it's close enough
                        If Distancia(NpcList(TempCharIndex).Pos, .Pos) > 10 Then
                            Call WriteChatOverHead(UserIndex, "Estás demasiado lejos.", NpcList(TempCharIndex).Char.CharIndex, vbWhite)
                            Exit Sub
                        End If
                        
                        Call WriteTrainerCreatureList(UserIndex, .flags.TargetNpc)
                        
                                
                    Else
                        Dim CentinelaIndex As Integer
                        
                        CentinelaIndex = EsCentinela(TempCharIndex)
                        
                        If CentinelaIndex <> 0 Then
                            'Enviamos nuevamente el texto del centinela según quien pregunta
                            Call modCentinela.CentinelaSendClave(UserIndex, CentinelaIndex)
                        End If
                    End If
                End If
            End If
            
            If FoundChar = 0 Then
                .flags.TargetNpc = 0
                .flags.TargetNpcTipo = eNpcType.Comun
                .flags.TargetUser = 0
                .flags.SelectedChar = 0
            End If
            
            If FoundSomething = 0 Then
                .flags.TargetNpc = 0
                .flags.TargetNpcTipo = eNpcType.Comun
                .flags.SelectedChar = 0
                .flags.TargetUser = 0
                .flags.TargetObjIndex = 0
                .flags.TargetObjMap = 0
                .flags.TargetObjX = 0
                .flags.TargetObjY = 0
            End If
        
        Else
            If FoundSomething = 0 Then
                .flags.TargetNpc = 0
                .flags.TargetNpcTipo = eNpcType.Comun
                .flags.SelectedChar = 0
                .flags.TargetUser = 0
                .flags.TargetObjIndex = 0
                .flags.TargetObjMap = 0
                .flags.TargetObjX = 0
                .flags.TargetObjY = 0
            End If
        End If
    
    End With
    
End Sub

Public Function FindDirection(Pos As WorldPos, Target As WorldPos) As eHeading
'***
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'***
    Dim X As Integer
    Dim Y As Integer
    
    X = Pos.X - Target.X
    Y = Pos.Y - Target.Y
    
    'NE
    If Sgn(X) = -1 And Sgn(Y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.NORTH, eHeading.EAST)

    'NW
    ElseIf Sgn(X) = 1 And Sgn(Y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.NORTH)
    
    'SW
    ElseIf Sgn(X) = 1 And Sgn(Y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.SOUTH)
    
    'SE
    ElseIf Sgn(X) = -1 And Sgn(Y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.SOUTH, eHeading.EAST)
    
    'Sur
    ElseIf Sgn(X) = 0 And Sgn(Y) = -1 Then
        FindDirection = eHeading.SOUTH
    
    'Norte
    ElseIf Sgn(X) = 0 And Sgn(Y) = 1 Then
        FindDirection = eHeading.NORTH
    
    'Oeste
    ElseIf Sgn(X) = 1 And Sgn(Y) = 0 Then
        FindDirection = eHeading.WEST
    
    'Este
    ElseIf Sgn(X) = -1 And Sgn(Y) = 0 Then
        FindDirection = eHeading.EAST
    
    'Misma
    ElseIf Sgn(X) = 0 And Sgn(Y) = 0 Then
        FindDirection = 0
    End If

End Function

Public Function EsObjetoFijo(ByVal ObjType As eObjType) As Boolean
        EsObjetoFijo = (ObjType = otPuerta Or _
                    ObjType = otForo Or _
                    ObjType = otCartel Or _
                    ObjType = otArbol Or _
                    ObjType = otArbolElfico Or _
                    ObjType = otYacimiento Or _
                    ObjType = otYunque Or _
                    ObjType = otFragua Or _
                    ObjType = otPortal Or _
                    ObjType = otAlijo)
End Function

Public Function ItemEsDeMapa(ByVal ObjType As eObjType) As Boolean
        ItemEsDeMapa = (ObjType = otForo Or _
                    ObjType = otCartel Or _
                    ObjType = otArbol Or _
                    ObjType = otArbol Or _
                    ObjType = otYacimiento Or _
                    ObjType = otAlijo)
End Function

Public Function RestrictStringToByte(ByRef restrict As String) As Byte
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 04/18/2011
'
'***************************************************
    restrict = UCase$(restrict)

    Select Case restrict
        Case "NEWBIE"
            RestrictStringToByte = 1
            
        Case "ARMADA"
            RestrictStringToByte = 2
            
        Case "CAOS"
            RestrictStringToByte = 3
            
        Case "FACCION"
            RestrictStringToByte = 4
            
        Case Else
            RestrictStringToByte = 0
    End Select
End Function

Public Function RestrictByteToString(ByVal restrict As Byte) As String
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 04/18/2011
'
'***************************************************
    Select Case restrict
        Case 1
            RestrictByteToString = "NEWBIE"
            
        Case 2
            RestrictByteToString = "ARMADA"
            
        Case 3
            RestrictByteToString = "CAOS"
            
        Case 4
            RestrictByteToString = "FACCION"
            
        Case 0
            RestrictByteToString = "NO"
    End Select
End Function

Public Function TerrainStringToByte(ByRef restrict As String) As Byte
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 04/18/2011
'
'***************************************************
    restrict = UCase$(restrict)

    Select Case restrict
        Case "NIEVE"
            TerrainStringToByte = 1
            
        Case "DESIERTO"
            TerrainStringToByte = 2
            
        Case "CIUDAD"
            TerrainStringToByte = 3
            
        Case "CAMPO"
            TerrainStringToByte = 4
            
        Case "DUNGEON"
            TerrainStringToByte = 5
            
        Case Else
            TerrainStringToByte = 0
    End Select
End Function

Public Function TerrainByteToString(ByVal restrict As Byte) As String
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 04/18/2011
'
'***************************************************
    Select Case restrict
        Case 1
            TerrainByteToString = "NIEVE"
            
        Case 2
            TerrainByteToString = "DESIERTO"
            
        Case 3
            TerrainByteToString = "CIUDAD"
            
        Case 4
            TerrainByteToString = "CAMPO"
            
        Case 5
            TerrainByteToString = "DUNGEON"
            
        Case 0
            TerrainByteToString = "BOSQUE"
    End Select
End Function

