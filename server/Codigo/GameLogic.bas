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
    
    Dim map As Integer
    Dim x As Byte
    Dim y As Byte

    map = UserList(UserIndex).Pos.map
    x = UserList(UserIndex).Pos.x
    y = UserList(UserIndex).Pos.y

    If InMapBounds(map, x, y) Then
        With maps(map).mapData(x, y)
            If .ObjInfo.index > 0 Then
                FxFlag = (ObjData(.ObjInfo.index).Type = otPortal)
                TelepRadio = ObjData(.ObjInfo.index).Radio
            End If
            
            If .TileExit.map > 0 Then
                
                If MapaValido(.TileExit.map) Then
                    'Es un teleport, entra en una posición random, acorde al radio (si es 0, es pos fija)
                    If FxFlag And TelepRadio > 0 Then
                        DestPos.x = .TileExit.x + RandomNumber(TelepRadio * (-1), TelepRadio)
                        DestPos.y = .TileExit.y + RandomNumber(TelepRadio * (-1), TelepRadio)
                    'Posición fija
                    Else
                        DestPos.x = .TileExit.x
                        DestPos.y = .TileExit.y
                    End If
                    
                    DestPos.map = .TileExit.map
                    
                    If DestPos.map = 286 Then
                        DestPos.map = Newbie.map
                        DestPos.x = Newbie.x
                        DestPos.y = Newbie.y
                    End If
                    
                    '¿Es mapa de principiantes?
                    If UCase$(MapInfo(DestPos.map).restringir) = "NEWBIE" Then
                        '¿El usuario es un principiante?
                        If EsPrincipiante(UserIndex) Or EsGM(UserIndex) Then
                            If LegalPos(DestPos.map, DestPos.x, DestPos.y, PuedeAtravesarAgua(UserIndex)) Then
                                Call WarpUserChar(UserIndex, DestPos.map, DestPos.x, DestPos.y, FxFlag)
                            Else
                                Call ClosestLegalPos(DestPos, nPos)
                                If nPos.x > 0 And nPos.y > 0 Then
                                    Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.y, FxFlag)
                                End If
                            End If
                        Else 'No es principiante
                            Call WriteConsoleMsg(UserIndex, "Esta zona es solo para principiantes.", FontTypeNames.FONTTYPE_INFO)
                            Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
            
                            If nPos.x > 0 And nPos.y > 0 Then
                                Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.y, False)
                            End If
                        End If
    
                    Else
                        If LegalPos(DestPos.map, DestPos.x, DestPos.y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, DestPos.map, DestPos.x, DestPos.y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)
                            
                            If nPos.x > 0 And nPos.y > 0 Then
                                Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.y, FxFlag)
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
                    Call AgregarPlataforma(UserIndex, map)
                    
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

Public Function InRangoVision(ByVal UserIndex As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean

    If x > UserList(UserIndex).Pos.x - MinXBorder And x < UserList(UserIndex).Pos.x + MinXBorder Then
        If y > UserList(UserIndex).Pos.y - MinYBorder And y < UserList(UserIndex).Pos.y + MinYBorder Then
            InRangoVision = True
            Exit Function
        End If
    End If
    
    InRangoVision = False

End Function

Public Function InRangoVisionNpc(ByVal NpcIndex As Integer, x As Integer, y As Integer) As Boolean

    If x > NpcList(NpcIndex).Pos.x - MinXBorder And x < NpcList(NpcIndex).Pos.x + MinXBorder Then
        If y > NpcList(NpcIndex).Pos.y - MinYBorder And y < NpcList(NpcIndex).Pos.y + MinYBorder Then
            InRangoVisionNpc = True
            Exit Function
        End If
    End If
    
    InRangoVisionNpc = False

End Function

Public Function InMapBounds(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    If (map < 1 Or map > NumMaps) Or x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
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
    
    vX = Pos.x - Distance
    vY = Pos.y
    
    For i = 0 To Distance - 1
        If (LegalPos(Pos.map, vX + i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX + i
            vY = vY - i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    vX = Pos.x
    vY = Pos.y - Distance
    
    For i = 0 To Distance - 1
        If (LegalPos(Pos.map, vX + i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX + i
            vY = vY + i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    vX = Pos.x + Distance
    vY = Pos.y
    
    For i = 0 To Distance - 1
        If (LegalPos(Pos.map, vX - i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX - i
            vY = vY + i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    vX = Pos.x
    vY = Pos.y + Distance
    
    For i = 0 To Distance - 1
        If (LegalPos(Pos.map, vX - i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
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

On Error GoTo ErrHandler

    Dim i As Long
    Dim HayObj As Boolean
    
    Dim x As Integer
    Dim y As Integer
    Dim MapObjIndex As Integer
    
    vX = Pos.x - Distance
    vY = Pos.y
    
    For i = 0 To Distance - 1
        
        x = vX + i
        y = vY - i
        
        If LegalPos(Pos.map, x, y, PuedeAgua, PuedeTierra, True) Then
           If Not HayObjeto(Pos.map, x, y, ObjIndex) Then
                vX = x
                vY = y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    
    vX = Pos.x
    vY = Pos.y - Distance
    
    For i = 0 To Distance - 1
        
        x = vX + i
        y = vY + i
        
        If LegalPos(Pos.map, x, y, PuedeAgua, PuedeTierra, True) Then
            If Not HayObjeto(Pos.map, x, y, ObjIndex) Then
                vX = x
                vY = y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    
    vX = Pos.x + Distance
    vY = Pos.y
    
    For i = 0 To Distance - 1
        
        x = vX - i
        y = vY + i
    
        If LegalPos(Pos.map, x, y, PuedeAgua, PuedeTierra, True) Then
            If Not HayObjeto(Pos.map, x, y, ObjIndex) Then
                vX = x
                vY = y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    
    vX = Pos.x
    vY = Pos.y + Distance
    
    For i = 0 To Distance - 1
        
        x = vX - i
        y = vY - i
    
        If LegalPos(Pos.map, x, y, PuedeAgua, PuedeTierra, True) Then
            If Not HayObjeto(Pos.map, x, y, ObjIndex) Then
                vX = x
                vY = y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
        
    Exit Function
    
ErrHandler:
    Call LogError("Error en RhombLegalTilePos. Error: " & Err.Number & " - " & Err.description)
End Function

Public Function HayObjeto(ByVal mapa As Integer, ByVal x As Long, ByVal y As Long, _
                          ByVal ObjIndex As Integer) As Boolean
    
    Dim MapObjIndex As Integer
    MapObjIndex = maps(mapa).mapData(x, y).ObjInfo.index
            
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
    tX = Pos.x
    tY = Pos.y
    
    LoopC = 1
    
    'La primera posición es valida?
    If LegalPos(Pos.map, nPos.x, nPos.y, PuedeAgua, PuedeTierra, CheckExitTile) Then
        Found = True
    
    'Busca en las demas posiciones, en forma de "rombo"
    Else
        While (Not Found) And LoopC < 22
            If RhombLegalPos(Pos, tX, tY, LoopC, PuedeAgua, PuedeTierra, CheckExitTile) Then
                nPos.x = tX
                nPos.y = tY
                Found = True
            End If
        
            LoopC = LoopC + 1
        Wend
        
    End If
    
    If Not Found Then
        nPos.x = 0
        nPos.y = 0
    End If

End Sub

Private Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
'Encuentra la posición legal mas cercana que no sea un portal y la guarda en nPos

    Dim Notfound As Boolean
    Dim LoopC As Integer
    Dim tX As Long
    Dim tY As Long
    
    nPos.map = Pos.map
    
    Do While Not LegalPos(Pos.map, nPos.x, nPos.y)
        If LoopC > 12 Then
            Notfound = True
            Exit Do
        End If
        
        For tY = Pos.y - LoopC To Pos.y + LoopC
            For tX = Pos.x - LoopC To Pos.x + LoopC
                
                If LegalPos(nPos.map, tX, tY) And maps(nPos.map).mapData(tX, tY).TileExit.map = 0 Then
                    nPos.x = tX
                    nPos.y = tY
                    '¿Hay objeto?
                    
                    tX = Pos.x + LoopC
                    tY = Pos.y + LoopC
      
                End If
            
            Next tX
        Next tY
        
        LoopC = LoopC + 1
        
    Loop
    
    If Notfound Then
        nPos.x = 0
        nPos.y = 0
    End If

End Sub

Public Function NameIndex(ByVal name As String) As Integer
    Dim UserIndex As Integer, i As Integer
     
    If InStrB(name, "+") > 0 Then
        name = UCase$(Replace(name, "+", " "))
    End If
     
    If Len(name) < 1 Then
        NameIndex = 0
        Exit Function
    End If
     
    UserIndex = 1
    
    If Right$(name, 1) = "*" Then
        name = Left$(name, Len(name) - 1)
        For i = 1 To LastUser
            If UCase$(UserList(i).name) = UCase$(name) Then
                NameIndex = i
                Exit Function
            End If
        Next
    Else
        For i = 1 To LastUser
            If UCase$(Left$(UserList(i).name, Len(name))) = UCase$(name) Then
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

Public Function CheckForSameName(ByVal name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
    Dim LoopC As Long
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.Logged Then
            
            'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
            'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
            'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
            'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ItemS
            'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
            
            If UCase$(UserList(LoopC).name) = UCase$(name) Then
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
            Pos.y = Pos.y - 1
        
        Case eHeading.SOUTH
            Pos.y = Pos.y + 1
        
        Case eHeading.EAST
            Pos.x = Pos.x + 1
        
        Case eHeading.WEST
            Pos.x = Pos.x - 1
    End Select
End Sub

Public Function LegalPos(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal CheckExitTile As Boolean = False) As Boolean
'Checks if the position is Legal.
    
    '¿Es un mapa válido?
    If (map < 1 Or map > NumMaps) Or _
       (x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) Then
                LegalPos = False
    Else
        With maps(map).mapData(x, y)
            If PuedeAgua And PuedeTierra Then
                LegalPos = Not .Blocked And _
                           .UserIndex < 1 And _
                           .NpcIndex < 1
            
            ElseIf PuedeTierra And Not PuedeAgua Then
                LegalPos = Not .Blocked And _
                           .UserIndex < 1 And _
                           .NpcIndex < 1 And _
                           (Not HayAgua(map, x, y))

            ElseIf PuedeAgua And Not PuedeTierra Then
                LegalPos = Not .Blocked And _
                           .UserIndex < 1 And _
                           .NpcIndex < 1 And _
                           (HayAgua(map, x, y))
            Else
                LegalPos = False
            End If
        End With
        
        If CheckExitTile Then
            LegalPos = LegalPos And (maps(map).mapData(x, y).TileExit.map = 0)
        End If
        
    End If

End Function

Public Function MoveToLegalPos(ByVal UserMoving As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True) As Boolean

    Dim UserIndex As Integer
    Dim IsDeadChar As Boolean
    Dim IsAdminInvisible As Boolean
    
    '¿Es un mapa válido?
    If map < 1 Or map > NumMaps Or x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        MoveToLegalPos = False
    Else
        UserIndex = maps(map).mapData(x, y).UserIndex
        
        If UserIndex > 0 Then
            IsDeadChar = UserList(UserIndex).Stats.Muerto
            IsAdminInvisible = (UserList(UserIndex).flags.AdminInvisible > 0)
        Else
            IsDeadChar = False
            IsAdminInvisible = False
        End If
            
        If EsGM(UserMoving) Or UserList(UserMoving).Stats.Muerto Then
            MoveToLegalPos = (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                       (maps(map).mapData(x, y).NpcIndex = 0)
        ElseIf PuedeAgua And PuedeTierra Then
            MoveToLegalPos = Not maps(map).mapData(x, y).Blocked And _
                       (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                       (maps(map).mapData(x, y).NpcIndex = 0)
        ElseIf PuedeTierra And Not PuedeAgua Then
            MoveToLegalPos = Not maps(map).mapData(x, y).Blocked And _
                       (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                       (maps(map).mapData(x, y).NpcIndex = 0) And _
                       (Not HayAgua(map, x, y))
        ElseIf PuedeAgua And Not PuedeTierra Then
            MoveToLegalPos = Not maps(map).mapData(x, y).Blocked And _
                       (maps(map).mapData(x, y).NpcIndex = 0) And _
                       (HayAgua(map, x, y))
                        'ESTO O ALGO ACA PARECE NO TERMINADO
        Else
            MoveToLegalPos = False
        End If
    End If
End Function

Public Sub FindLegalPos(ByVal UserIndex As Integer, ByVal map As Integer, ByVal x As Byte, ByVal y As Byte)
'Search for a Legal pos for the user who is being teleported.

    If maps(map).mapData(x, y).UserIndex > 0 Or _
        maps(map).mapData(x, y).NpcIndex > 0 Then
                    
        'Se teletransporta a la misma pos a la que estaba
        If maps(map).mapData(x, y).UserIndex = UserIndex Then
           Exit Sub
        End If
        
        Dim FoundPlace As Boolean
        Dim tX As Byte
        Dim tY As Byte
        Dim Rango As Byte
        Dim OtherUserIndex As Integer
    
        For Rango = 1 To 5
            For tY = y - Rango To y + Rango
                For tX = x - Rango To x + Rango
                    'Reviso que no haya User ni Npc
                    If maps(map).mapData(tX, tY).UserIndex = 0 Then
                        If maps(map).mapData(tX, tY).NpcIndex = 0 Then
                            If maps(map).mapData(tX, tY).Trigger <> eTrigger.EnPlataforma Then
                                If InMapBounds(map, tX, tY) Then
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
            x = tX
            y = tY
        Else
            'Muy poco probable, pero..
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un Npc, lo pisamos.
            OtherUserIndex = maps(map).mapData(x, y).UserIndex
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

Public Function LegalPosNpc(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal AguaValida As Byte, Optional ByVal IsPet As Boolean = False) As Boolean

    Dim IsDeadChar As Boolean
    Dim UserIndex As Integer
    Dim IsAdminInvisible As Boolean
        
    If (map < 1 Or map > NumMaps) Or _
        (x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) Then
        LegalPosNpc = False
        Exit Function
    End If

    With maps(map).mapData(x, y)
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
            (.Trigger <> eTrigger.POSINVALIDA And .Trigger <> eTrigger.EnPlataforma) _
            And Not HayAgua(map, x, y)
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
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(NpcList(NpcIndex).Expresiones(randomi), NpcList(NpcIndex).Char.CharIndex, vbWhite))
    End If
End Sub

Public Sub LookatTile(ByVal UserIndex As Integer, ByVal map As Integer, ByVal x As Byte, ByVal y As Byte, Optional ByVal UsingSkill As Boolean = False, Optional ByVal RightClick As Boolean = False)

    'Responde al click del usuario sobre el mapa
    Dim FoundChar As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex As Integer
    Dim Stat As String
    Dim ft As FontTypeNames
    
    With UserList(UserIndex)
        '¿Rango Visión? (ToxicWaste)
        If (Abs(.Pos.y - y) > RANGO_VISION_Y) Or (Abs(.Pos.x - x) > RANGO_VISION_X) Then
            Exit Sub
        End If
        
        If .flags.Comerciando Then
            Exit Sub
        End If
    
        '¿Posicion valida?
        If InMapBounds(map, x, y) Then
        
            .flags.TargetMap = map
            .flags.TargetX = x
            .flags.TargetY = y
            
            '¿Es un obj?
            If maps(map).mapData(x, y).ObjInfo.index > 0 Then
            
                .flags.TargetObjMap = map
                .flags.TargetObjX = x
                .flags.TargetObjY = y
                
                FoundSomething = 1
    
                Select Case ObjData(maps(map).mapData(x, y).ObjInfo.index).Type
                    Case otPuerta 'Es una puerta
                        Call AccionParaPuerta(map, x, y, UserIndex)
                    
                    Case otLeña    'Leña
                        If maps(map).mapData(x, y).ObjInfo.index = FOGATA_APAG And Not .Stats.Muerto Then
                            Call AccionParaRamita(map, x, y, UserIndex)
                        End If
                    
                    Case otCartel    'Cartel
                        If Len(ObjData(maps(map).mapData(x, y).ObjInfo.index).texto) > 0 Then
                            Call WriteShowSignal(UserIndex, maps(map).mapData(x, y).ObjInfo.index)
                        End If
                        
                    Case otAlijo
                        If .flags.Comerciando Then
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
                End Select
                
            ElseIf maps(map).mapData(x + 1, y).ObjInfo.index > 0 Then
                
                If ObjData(maps(map).mapData(x + 1, y).ObjInfo.index).Type = otPuerta Then
                    .flags.TargetObjMap = map
                    .flags.TargetObjX = x + 1
                    .flags.TargetObjY = y
                    
                    FoundSomething = 1
                    
                    Call AccionParaPuerta(map, x + 1, y, UserIndex)
                End If
                
            ElseIf maps(map).mapData(x + 1, y + 1).ObjInfo.index > 0 Then
            
                If ObjData(maps(map).mapData(x + 1, y + 1).ObjInfo.index).Type = otPuerta Then
                    .flags.TargetObjMap = map
                    .flags.TargetObjX = x + 1
                    .flags.TargetObjY = y + 1
                    
                    FoundSomething = 1
                    
                    Call AccionParaPuerta(map, x + 1, y + 1, UserIndex)
                End If
                
            ElseIf maps(map).mapData(x, y + 1).ObjInfo.index > 0 Then
            
                .flags.TargetObjMap = map
                .flags.TargetObjX = x
                .flags.TargetObjY = y + 1
                
                FoundSomething = 1
                
                If ObjData(maps(map).mapData(x, y + 1).ObjInfo.index).Type = otPuerta Then
                    Call AccionParaPuerta(map, x, y + 1, UserIndex)
                End If
            End If
            
            If FoundSomething = 1 Then
                .flags.TargetObjIndex = maps(map).mapData(.flags.TargetObjX, .flags.TargetObjY).ObjInfo.index
            End If
            
            '¿Es un personaje?
            If y + 1 <= YMaxMapSize Then
                If maps(map).mapData(x, y + 1).UserIndex > 0 Then
                    TempCharIndex = maps(map).mapData(x, y + 1).UserIndex
                    FoundChar = 1
                ElseIf maps(map).mapData(x, y + 1).NpcIndex > 0 Then
                    TempCharIndex = maps(map).mapData(x, y + 1).NpcIndex
                    FoundChar = 2
                End If
            End If
            
            '¿Es un personaje?
            If FoundChar = 0 Then
                If maps(map).mapData(x, y).UserIndex > 0 Then
                    TempCharIndex = maps(map).mapData(x, y).UserIndex
                    FoundChar = 1
                End If
                If maps(map).mapData(x, y).NpcIndex > 0 Then
                    TempCharIndex = maps(map).mapData(x, y).NpcIndex
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
                                    
                                    Stat = .name
                                    
                                    If EsPrincipiante(TempCharIndex) Then
                                        Stat = " <Principiante>"
                                    End If
                                                          
                                    If .Guild_Id > 0 Then
                                        Stat = Stat & " <" & modGuilds.GuildName(.Guild_Id) & ">"
                                    End If
                            
                                    If Len(.Desc) > 0 Then
                                        Stat = .name & Stat & " - " & .Desc
                                    Else
                                        Stat = .name & Stat
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
                                        Call WriteConsoleMsg(UserIndex, NpcList(TempCharIndex).name & " (" & CInt(NpcList(TempCharIndex).Contadores.TiempoExistencia / frmMain.GameTimer.interval) & "/" & CInt(IntervaloInvocacion / frmMain.GameTimer.interval) & ")", FontTypeNames.FONTTYPE_INFO)
                                        
                                        If .flags.SelectedChar <> TempCharIndex Then
                                            .flags.SelectedChar = TempCharIndex
                                        Else
                                            .flags.SelectedChar = 0
                                        End If
                                    Else
                                        Call WriteConsoleMsg(UserIndex, NpcList(TempCharIndex).name & ", invocado por " & UserList(NpcList(TempCharIndex).MaestroUser).name, FontTypeNames.FONTTYPE_INFO)
                                        
                                        If .flags.SelectedChar > 0 Then
                                            .flags.SelectedChar = 0
                                        End If
                                    End If
                                End If
                            End If
                            
                        ElseIf NpcList(TempCharIndex).MaestroUser = UserIndex Then
                            If Not RightClick Then
                                Call WriteConsoleMsg(UserIndex, NpcList(TempCharIndex).name & " (" & NpcList(TempCharIndex).Stats.MinHP & "/" & NpcList(TempCharIndex).Stats.MaxHP & ")", FontTypeNames.FONTTYPE_INFO)
                                
                                If .flags.SelectedChar <> TempCharIndex Then
                                    .flags.SelectedChar = TempCharIndex
                                End If
                            End If

                        Else
                        
                            If Not RightClick Then
                                Call WriteConsoleMsg(UserIndex, NpcList(TempCharIndex).name & ", mascota de " & UserList(NpcList(TempCharIndex).MaestroUser).name & ".", FontTypeNames.FONTTYPE_INFO)
                                
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
                                Call WriteConsoleMsg(UserIndex, NpcList(TempCharIndex).name & " (" & NpcList(TempCharIndex).Stats.MinHP & "/" & NpcList(TempCharIndex).Stats.MaxHP & ") - Le pegó primero: " & UserList(NpcList(TempCharIndex).TargetUser).name & ".", FontTypeNames.FONTTYPE_INFO)
                            Else
                                Call WriteConsoleMsg(UserIndex, NpcList(TempCharIndex).name & " (" & NpcList(TempCharIndex).Stats.MinHP & "/" & NpcList(TempCharIndex).Stats.MaxHP & ")", FontTypeNames.FONTTYPE_INFO)
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
    Dim x As Integer
    Dim y As Integer
    
    x = Pos.x - Target.x
    y = Pos.y - Target.y
    
    'NE
    If Sgn(x) = -1 And Sgn(y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.NORTH, eHeading.EAST)

    'NW
    ElseIf Sgn(x) = 1 And Sgn(y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.NORTH)
    
    'SW
    ElseIf Sgn(x) = 1 And Sgn(y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.SOUTH)
    
    'SE
    ElseIf Sgn(x) = -1 And Sgn(y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.SOUTH, eHeading.EAST)
    
    'Sur
    ElseIf Sgn(x) = 0 And Sgn(y) = -1 Then
        FindDirection = eHeading.SOUTH
    
    'Norte
    ElseIf Sgn(x) = 0 And Sgn(y) = 1 Then
        FindDirection = eHeading.NORTH
    
    'Oeste
    ElseIf Sgn(x) = 1 And Sgn(y) = 0 Then
        FindDirection = eHeading.WEST
    
    'Este
    ElseIf Sgn(x) = -1 And Sgn(y) = 0 Then
        FindDirection = eHeading.EAST
    
    'Misma
    ElseIf Sgn(x) = 0 And Sgn(y) = 0 Then
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

