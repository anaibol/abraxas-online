Attribute VB_Name = "Acciones"
Option Explicit

Public Sub AccionParaPuerta(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)

On Error Resume Next

    If Not IntervaloPermiteAtacar(UserIndex) Then
        Exit Sub
    End If
    
    If Distance(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y, x, y) < 3 Then
    
        With maps(map).mapData(x, y)
        
            If Not ObjData(.ObjInfo.index).Llave Then
                If ObjData(.ObjInfo.index).Cerrada Then
                    'Abre la puerta
                    If Not ObjData(.ObjInfo.index).Llave Then
                        
                        .ObjInfo.index = ObjData(.ObjInfo.index).IndexAbierta
                        
                        Call modSendData.SendToAreaByPos(map, x, y, PrepareMessageObjCreate(ObjData(.ObjInfo.index).GrhIndex, ObjData(.ObjInfo.index).Type, x, y))
                        
                        'Desbloquea
                        .Blocked = False
                        maps(map).mapData(x - 1, y).Blocked = False
                        
                        'Bloquea todos los mapas
                        Call Bloquear(True, map, x, y, 0)
                        Call Bloquear(True, map, x - 1, y, 0)
                          
                        'Sonido
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, x, y))
                        
                    Else
                         Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    'Cierra puerta
                    .ObjInfo.index = ObjData(.ObjInfo.index).IndexCerrada
                    
                    Call modSendData.SendToAreaByPos(map, x, y, PrepareMessageObjCreate(ObjData(.ObjInfo.index).GrhIndex, ObjData(.ObjInfo.index).Type, x, y))
                                    
                    .Blocked = True
                    maps(map).mapData(x - 1, y).Blocked = True
                    
                    
                    Call Bloquear(True, map, x - 1, y, 1)
                    Call Bloquear(True, map, x, y, 1)
                    
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, x, y))
                End If
                
                UserList(UserIndex).flags.TargetObjIndex = .ObjInfo.index
            Else
                Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
            End If
            
        End With
        
    Else
        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
    End If

End Sub

Public Sub AccionParaRamita(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

    Dim Suerte As Byte
    Dim exito As Byte
    Dim Obj As Obj
    
    Dim Pos As WorldPos
    Pos.map = map
    Pos.x = x
    Pos.y = y
    
    With UserList(UserIndex)
        If Distancia(Pos, .Pos) > 2 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If maps(map).mapData(x, y).Trigger = eTrigger.ZONASEGURA Or maps(map).mapData(x, y).Trigger = eTrigger.EnPlataforma Or MapInfo(map).PK = False Then
            Call WriteConsoleMsg(UserIndex, "No podés hacer fogatas en zona segura.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .Skills.Skill(eSkill.Supervivencia).Elv > 1 And .Skills.Skill(eSkill.Supervivencia).Elv < 6 Then
            Suerte = 3
        ElseIf .Skills.Skill(eSkill.Supervivencia).Elv >= 6 And .Skills.Skill(eSkill.Supervivencia).Elv < 20 Then
            Suerte = 2
        ElseIf .Skills.Skill(eSkill.Supervivencia).Elv >= 10 And .Skills.Skill(eSkill.Supervivencia).Elv Then
            Suerte = 1
        End If
        
        exito = RandomNumber(1, Suerte)
    
        If exito = 1 Then
            If MapInfo(.Pos.map).Zona <> Ciudad Then
                Obj.index = FOGATA
                Obj.Amount = 1
                
                Call WriteConsoleMsg(UserIndex, "Prendiste la fogata.", FontTypeNames.FONTTYPE_INFO)
                
                Call MakeObj(Obj, map, x, y)
                
                'Las fogatas prendidas se deben eliminar
                'Dim Fogatita As New cGarbage
                'Fogatita.Map = Map
                'Fogatita.X = X
                'Fogatita.Y = Y
                'Call TrashCollector.Add(Fogatita)
                    
                Call SubirSkill(UserIndex, eSkill.Supervivencia, True)
            Else
                Call WriteConsoleMsg(UserIndex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)
            Call SubirSkill(UserIndex, eSkill.Supervivencia, False)
        End If
    
    End With
End Sub
