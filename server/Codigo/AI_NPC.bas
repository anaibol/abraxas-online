Attribute VB_Name = "AI"
Option Explicit

Public Enum TipoAI
    Estatico = 1
    MueveAlAzar = 2
    NpcMaloAtacaUsersBuenos = 3
    NpcDefensa = 4
    NpcObjeto = 6
    SigueAmo = 8
    NpcPathfinding = 10
End Enum

Public Const ELEMENTALFUEGO As Integer = 93
Public Const ELEMENTALTIERRA As Integer = 94
Public Const ELEMENTALAGUA As Integer = 92

'Damos a los Npcs el mismo rango de visión que un PJ
Public Const RANGO_VISION_X As Byte = 8
Public Const RANGO_VISION_Y As Byte = 6

Private Sub BuscarUserCerca(ByVal NpcIndex As Integer)

    Dim nPos As WorldPos
    Dim tHeading As Byte
    Dim UserIndex As Integer
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    Dim i As Long
    Dim UserProtected As Boolean
    
    With NpcList(NpcIndex)
    
        Dim RangoVisionNpcx As Byte
        Dim RangoVisionNpcy As Byte
        
        If .TargetUser > 0 Or .TargetNpc > 0 Then
            RangoVisionNpcx = 10
            RangoVisionNpcy = 8
            
        ElseIf .flags.LanzaSpells > 0 Then
            RangoVisionNpcx = 9
            RangoVisionNpcy = 7
        
        Else
            Select Case .Lvl
                Case 2
                    RangoVisionNpcx = 7
                    RangoVisionNpcy = 6
                Case 3
                    RangoVisionNpcx = 8
                    RangoVisionNpcy = 6
                Case 4
                    RangoVisionNpcx = 10
                    RangoVisionNpcy = 8
            End Select
        End If
    
        'If .TargetUser > 0 Then
        '    UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
        '    UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado Or UserList(UserIndex).flags.EnConsulta
            
        '    If Abs(UserList(UserIndex).Pos.X - .Pos.X) > RangoVisionNpcx Or _
        '        Abs(UserList(UserIndex).Pos.Y - .Pos.Y) > RangoVisionNpcy Or _
        '        UserList(UserIndex).Stats.Muerto Or UserList(UserIndex).flags.Invisible > 0 Or UserList(UserIndex).flags.Oculto > 0 Or _
        '        Not UserList(UserIndex).flags.AdminPerseguible Then
                
        '        .TargetUser = 0
        '    End If
        'End If
        
        If .TargetUser < 1 Then
            For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
                UserIndex = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                
                If UserIndex > 0 Then
                    UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
                    UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado Or UserList(UserIndex).flags.EnConsulta
                    
                    If Abs(UserList(UserIndex).Pos.x - .Pos.x) > RangoVisionNpcx Or _
                        Abs(UserList(UserIndex).Pos.y - .Pos.y) > RangoVisionNpcy Or _
                        UserList(UserIndex).Stats.Muerto Or UserList(UserIndex).flags.Invisible > 0 Or UserList(UserIndex).flags.Oculto > 0 Or _
                        Not UserList(UserIndex).flags.AdminPerseguible Then
                        
                        .TargetUser = 0
                    
                    Else
                        .TargetUser = UserIndex
                        Exit For
                    End If
                End If
            Next i
        End If
        
        If .TargetUser < 1 Then
            If .flags.Inmovilizado < 1 Then
                
                For tHeading = eHeading.NORTH To eHeading.WEST
                    nPos = .Pos
                    
                    If .flags.Inmovilizado < 1 Or .Char.Heading = tHeading Then
                        Call HeadtoPos(tHeading, nPos)
                        
                        If InMapBounds(nPos.map, nPos.x, nPos.y) Then
                            UserIndex = maps(nPos.map).mapData(nPos.x, nPos.y).UserIndex
                            
                            If UserIndex > 0 Then
                                UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
                                UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado Or UserList(UserIndex).flags.EnConsulta
                                
                                If Not UserList(UserIndex).Stats.Muerto And UserList(UserIndex).flags.AdminPerseguible And (Not UserProtected) Then
                                    .TargetUser = UserIndex
                                End If
                            End If
                        End If
                    End If
                    
                Next tHeading
                
            End If
        End If

    End With
    
End Sub

Public Sub SeguirMaestro(ByVal NpcIndex As Integer)
    
    With UserList(NpcList(NpcIndex).MaestroUser)
    
        If Distancia(NpcList(NpcIndex).Pos, .Pos) > 20 Or .flags.Invisible > 0 Or .flags.Oculto > 0 Or .flags.AdminInvisible > 0 Then
            If RandomNumber(0, 25) = 0 Then
                Call MoveNpc(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
            End If
                        
        ElseIf Distancia(NpcList(NpcIndex).Pos, .Pos) = 1 Or Distancia(NpcList(NpcIndex).Pos, .Pos) = 2 Then
            If RandomNumber(0, 50) = 0 Then
                Call MoveNpc(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
            End If
            
        ElseIf Distancia(NpcList(NpcIndex).Pos, .Pos) > 2 Then
            If NpcList(NpcIndex).PFINFO.PathLenght < 1 Then
                Call MoveNpc(NpcIndex, FindDirection(NpcList(NpcIndex).Pos, .Pos))
            Else
                Call FollowPath(NpcIndex)
            End If
        End If
        
    End With
    
End Sub

Private Sub AiNpcAtacaUser(ByVal NpcIndex As Integer)

    Dim nPos As WorldPos
    Dim tHeading As Byte
    Dim UserIndex As Integer
    Dim UserProtected As Boolean
    Dim Ataca As Boolean
    
    With NpcList(NpcIndex)

        If Abs(.Pos.x - UserList(.TargetUser).Pos.x) > 10 Or Abs(.Pos.y - UserList(.TargetUser).Pos.y) > 8 Then
            Call BuscarUserCerca(NpcIndex)
        End If
        
        If Distancia(.Pos, UserList(.TargetUser).Pos) > 1 Then
            If UserList(.TargetUser).flags.Invisible > 0 Or UserList(.TargetUser).flags.Oculto > 0 Then
                Call BuscarUserCerca(NpcIndex)
                Exit Sub
            End If
            
            If NpcList(NpcIndex).PFINFO.PathLenght < 1 Then
                Call MoveNpc(NpcIndex, tHeading)
            Else
                Call FollowPath(NpcIndex)
            End If
        
        Else
            Dim AttackPos As WorldPos

            AttackPos = .Pos
            Call HeadtoPos(.Char.Heading, AttackPos)
            
            If maps(AttackPos.map).mapData(AttackPos.x, AttackPos.y).NpcIndex = .TargetUser Then
                Ataca = True
            
            ElseIf .flags.Inmovilizado < 1 Then
                Ataca = True
                
                tHeading = FindDirection(.Pos, UserList(.TargetUser).Pos)
                
                If tHeading <> .Char.Heading Then
                    .Char.Heading = tHeading
                    Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChangeCharHeading(.Char.CharIndex, tHeading))
                End If
            End If
        End If
        
        If Ataca Then
        
            UserProtected = Not IntervaloPermiteSerAtacado(.TargetUser) And UserList(.TargetUser).flags.NoPuedeSerAtacado
            UserProtected = UserProtected Or UserList(.TargetUser).flags.Ignorado Or UserList(.TargetUser).flags.EnConsulta
            
            If Not UserList(.TargetUser).Stats.Muerto And UserList(.TargetUser).flags.AdminPerseguible And (Not UserProtected) Then

                If .flags.LanzaSpells Then
                
                    If .Type = DRAGON Then
                        If Distancia(.Pos, UserList(.TargetUser).Pos) < 2 Then
                            If RandomNumber(0, 1) = 1 Then
                                Call NpcLanzaUnSpellSobreUser(NpcIndex, .TargetUser)
                            Else
                                Call SistemaCombate.NpcAtacaUser(NpcIndex, .TargetUser)
                            End If
                        Else
                            Call NpcLanzaUnSpellSobreUser(NpcIndex, .TargetUser)
                        End If
                    
                    ElseIf .flags.AtacaDoble = 1 Then
                        If RandomNumber(0, 1) = 0 Then
                            Call NpcAtacaUser(NpcIndex, .TargetUser)
                        End If
                        Call NpcLanzaUnSpellSobreUser(NpcIndex, .TargetUser)
                    
                    ElseIf RandomNumber(0, 2) = 0 Then
                        Call NpcLanzaUnSpellSobreUser(NpcIndex, .TargetUser)
                    End If
                
                Else
                    Call NpcAtacaUser(NpcIndex, .TargetUser)
                End If
            End If
        End If
    End With

End Sub

Private Sub AiNpcAtacaNpc(ByVal NpcIndex As Integer)
    
    Dim tHeading As Byte
    Dim x As Long
    Dim y As Long
        
    With NpcList(NpcIndex)
    
        If Abs(.Pos.x - NpcList(.TargetNpc).Pos.x) > 10 Or Abs(.Pos.y - NpcList(.TargetNpc).Pos.y) > 8 Then
            If .MaestroUser > 0 Then
                Call FollowAmo(NpcIndex)
            Else
                .TargetNpc = 0
            End If
            
            Exit Sub
        End If
        
        If .MaestroUser = NpcList(.TargetNpc).MaestroUser Then
            Call FollowAmo(NpcIndex)
            Exit Sub
        End If
    
        With NpcList(NpcIndex)
            If NpcList(.TargetNpc).MaestroUser > 0 Then
                If Distancia(.Pos, UserList(NpcList(.TargetNpc).MaestroUser).Pos) < 4 Then
                    .TargetUser = NpcList(.TargetNpc).MaestroUser
                    .TargetNpc = 0
                    
                    Exit Sub
                End If
            End If
        End With
        
        With NpcList(.TargetNpc)
            If .TargetUser > 0 Then
                If Distancia(.Pos, NpcList(NpcIndex).Pos) < 2 And .Stats.MinHP > NpcList(NpcIndex).Stats.MinHP And Distancia(.Pos, UserList(.TargetUser).Pos) > 3 Then
                    .TargetNpc = NpcIndex
                    .TargetUser = 0
                    
                    Exit Sub
                End If
            End If
        End With
        
        If .flags.LanzaSpells > 0 Then
            Call NpcLanzaUnSpellSobreNpc(NpcIndex, .TargetNpc)
            
        ElseIf Distancia(.Pos, NpcList(.TargetNpc).Pos) < 2 Then
            Dim AttackPos As WorldPos

            AttackPos = .Pos
            Call HeadtoPos(.Char.Heading, AttackPos)
            
            If maps(AttackPos.map).mapData(AttackPos.x, AttackPos.y).NpcIndex = .TargetNpc Then
                Call NpcAtacaNpc(NpcIndex, .TargetNpc)
            End If
        End If
            
        If .TargetNpc > 0 Then
            tHeading = FindDirection(.Pos, NpcList(.TargetNpc).Pos)
            
            If Distancia(.Pos, NpcList(.TargetNpc).Pos) > 1 Then
                If NpcList(NpcIndex).PFINFO.PathLenght < 1 Then
                    Call MoveNpc(NpcIndex, tHeading)
                Else
                    Call FollowPath(NpcIndex)
                End If
            
            ElseIf tHeading <> .Char.Heading Then
                .Char.Heading = tHeading
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChangeCharHeading(.Char.CharIndex, tHeading))
            End If
        End If
        
    End With
End Sub

Public Sub Resucitar(ByVal NpcIndex As Integer)

    Dim UserIndex As Integer
    Dim i As Long
    
    With NpcList(NpcIndex)
        For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
            UserIndex = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
            
            'Is it in it's range of vision??
            If Abs(UserList(UserIndex).Pos.x - .Pos.x) <= RANGO_VISION_X Then
                If Abs(UserList(UserIndex).Pos.y - .Pos.y) <= RANGO_VISION_Y Then
                    
                    With UserList(UserIndex)
                        If .Stats.Muerto Then
                            If maps(.Pos.map).mapData(.Pos.x, .Pos.y).Blocked Then
                                Exit Sub
                            End If
                            
                            Call RevivirUsuario(UserIndex)
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Pos.x, .Pos.y, 9))
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(18, .Pos.x, .Pos.y))
                        Else
                            Dim FX As Byte
                            
                            If .Stats.MinHP <> .Stats.MaxHP Then
                                .Stats.MinHP = .Stats.MaxHP
                                Call WriteUpdateHP(UserIndex)
                                FX = 9
                            End If
                            
                            If .Stats.MinMan <> .Stats.MaxMan Then
                                .Stats.MinMan = .Stats.MaxMan
                                Call WriteUpdateMana(UserIndex)
                                FX = 9
                            End If
                            
                            If .flags.Envenenado > 0 Then
                                .flags.Envenenado = 0
                                If FX <> 9 Then
                                    FX = 2
                                End If
                            End If
                            
                            If FX > 0 Then
                                If .Stats.MinSta <> .Stats.MaxSta Then
                                    .Stats.MinSta = .Stats.MaxSta
                                    Call WriteUpdateSta(UserIndex)
                                    FX = 9
                                End If
                            
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Pos.x, .Pos.y, FX))
                                
                                If FX = 9 Then
                                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(18, .Pos.x, .Pos.y))
                                Else
                                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(16, .Pos.x, .Pos.y))
                                End If
                                
                                Exit Sub
                                
                                Select Case RandomNumber(0, 1000)
                                
                                    Case 0
                                        Call WriteChatOverHead(UserIndex, "¡Aguante Black Sabbath!", NpcList(NpcIndex).Char.CharIndex, vbWhite)
                                    
                                    Case 1
                                        Call WriteChatOverHead(UserIndex, "¿Hey, tú sabes que la propiedad Privada es un robo?", NpcList(NpcIndex).Char.CharIndex, vbWhite)
                                
                                    Case 2
                                    'Call WriteChatOverHead(UserIndex, "Tú, ¿te interesaste interesado alguna ves en la metafísica?", NpcList(NpcIndex).Char.CharIndex, vbWhite)
                                
                                    Case 3
                                        Call WriteChatOverHead(UserIndex, "Antes, mi filosofía era la religión. Ahora, mi religión es la filosofía.", NpcList(NpcIndex).Char.CharIndex, vbWhite)
                                
                                    Case 4
                                        Call WriteChatOverHead(UserIndex, "Llámalo comunismo, anarquismo o socialismo, pero haz algo.", NpcList(NpcIndex).Char.CharIndex, vbWhite)
                                    
                                    Case 5
                                        Call WriteChatOverHead(UserIndex, "Charlar y hacer son cosas diferentes, más bien antagónicas.", NpcList(NpcIndex).Char.CharIndex, vbWhite)
                                
                                    Case 6
                                        Call WriteChatOverHead(UserIndex, "¿Has visto la película 'Fight Club'?", NpcList(NpcIndex).Char.CharIndex, vbWhite)
                                
                                    Case 7
                                        Call WriteChatOverHead(UserIndex, "A los sacerdotes nos gusta el metal pesado.", NpcList(NpcIndex).Char.CharIndex, vbWhite)
                                
                                    Case 8
                                        'Call WriteChatOverHead(UserIndex, "¿Te has preguntado qué es la realidad?", NpcList(NpcIndex).Char.CharIndex, vbWhite)
                                
                                    Case 9
                                        Call WriteChatOverHead(UserIndex, "Si las puertas de la percepción se limpiaran, todo aparecería ante el hombre como es: infInito.", NpcList(NpcIndex).Char.CharIndex, vbWhite)
                                End Select
                            End If
                        End If
                    End With
               End If
            End If
            
        Next i
    End With

End Sub

Public Sub NpcAI(ByVal NpcIndex As Integer)
    
    With NpcList(NpcIndex)

        If .TargetUser > 0 Then
            Call AiNpcAtacaUser(NpcIndex)
        
        ElseIf .TargetNpc > 0 Then
            Call AiNpcAtacaNpc(NpcIndex)
        
        ElseIf .Hostile > 0 Or .Movement = NpcMaloAtacaUsersBuenos Then
            Call BuscarUserCerca(NpcIndex)
        
        ElseIf .Movement = MueveAlAzar Then
            If .flags.Inmovilizado > 0 Then
                Exit Sub
            End If
            
            If RandomNumber(0, 12) = 0 Then
                Call MoveNpc(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
            End If
        End If
    
    End With

End Sub

Public Sub NpcPetAi()
    
    Dim NpcIndex As Long
    
    If Not haciendoBK And Not EnPausa Then
        'Update Npcs
        For NpcIndex = 1 To LastNpc
            
            With NpcList(NpcIndex)
            
                If NpcList(NpcIndex).MaestroUser > 0 Then
            
                    If NpcList(NpcIndex).flags.Paralizado > 0 Or NpcList(NpcIndex).flags.Inmovilizado > 0 Then
                        
                        If .Contadores.Paralisis > 0 Then
                            .Contadores.Paralisis = .Contadores.Paralisis - 1
                        Else
                            .flags.Paralizado = 0
                            .flags.Inmovilizado = 0
                            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageSetParalized(.Char.CharIndex, 0))
                        End If
                        
                    Else
                        If .TargetUser > 0 Then
                            If .TargetUser = .MaestroUser Then
                                Call SeguirMaestro(NpcIndex)
                            Else
                                Call AiNpcAtacaUser(NpcIndex)
                            End If
                            
                        ElseIf .TargetNpc > 0 Then
                            Call AiNpcAtacaNpc(NpcIndex)
                        End If
                    End If
                End If
            End With
        Next NpcIndex
    End If
    
End Sub

Public Function FollowPath(ByVal NpcIndex As Integer) As Boolean
'Moves the npc.

    Dim tmpPos As WorldPos
    Dim tHeading As Byte
    
    With NpcList(NpcIndex)

        If NpcList(NpcIndex).PFINFO.PathLenght = 0 Then 'Si no existe nos movemos al azar
            If RandomNumber(0, 25) = 0 Then
                Call MoveNpc(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
            End If
            
            Exit Function
        End If
            
        tmpPos.map = .Pos.map
        tmpPos.x = .PFINFO.Path(.PFINFO.CurPos).y 'invertimos las coordenadas
        tmpPos.y = .PFINFO.Path(.PFINFO.CurPos).x
        'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"
        
        tHeading = FindDirection(.Pos, tmpPos)
        
        Call MoveNpc(NpcIndex, tHeading)
        
        .PFINFO.CurPos = .PFINFO.CurPos + 1
        
        If .PFINFO.CurPos = .PFINFO.PathLenght Then
            .PFINFO.PathLenght = 0
        End If
        
    End With

End Function

Public Function PathFindingAI(ByVal NpcIndex As Integer) As Boolean
    With NpcList(NpcIndex)
        If .TargetUser > 0 Then
            .PFINFO.Target.x = UserList(.TargetUser).Pos.y
            .PFINFO.Target.y = UserList(.TargetUser).Pos.x
        ElseIf .TargetNpc > 0 Then
            .PFINFO.Target.x = NpcList(.TargetNpc).Pos.y
            .PFINFO.Target.y = NpcList(.TargetNpc).Pos.x
        End If
    End With
    
    Call SeekPath(NpcIndex)
End Function

Public Sub NpcLanzaUnSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    
    If NpcList(NpcIndex).MaestroUser > 0 Then
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.Invisible > 0 Or UserList(UserIndex).flags.Oculto > 0 Then
        Exit Sub
    End If
    
    With NpcList(NpcIndex)
        Dim k As Integer
        k = RandomNumber(1, .flags.LanzaSpells)
        
        If Hechizos(.Spells.Spell(k)).Inmoviliza > 0 And UserList(UserIndex).flags.Inmovilizado > 0 Then
            k = RandomNumber(1, .flags.LanzaSpells)
            
            If Hechizos(.Spells.Spell(k)).Inmoviliza > 0 Then
                k = RandomNumber(1, .flags.LanzaSpells)
            
                If Hechizos(.Spells.Spell(k)).Inmoviliza > 0 Then
                    Exit Sub
                End If
    
            End If
        End If
        
        If Hechizos(.Spells.Spell(k)).Paraliza > 0 And UserList(UserIndex).flags.Paralizado > 0 Then
            k = RandomNumber(1, .flags.LanzaSpells)
            
            If Hechizos(.Spells.Spell(k)).Paraliza > 0 Then
                k = RandomNumber(1, .flags.LanzaSpells)
            
                If Hechizos(.Spells.Spell(k)).Paraliza > 0 Then
                    Exit Sub
                End If
            End If
        End If
        
        Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, .Spells.Spell(k))
    End With
End Sub

Public Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNpc As Integer)
    Dim k As Integer
    k = RandomNumber(1, NpcList(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNpc, NpcList(NpcIndex).Spells.Spell(k))
End Sub

