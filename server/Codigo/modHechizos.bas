Attribute VB_Name = "modHechizos"
       
Option Explicit

Public Const SUPERANILLO As Integer = 700

Public Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer)

    If NpcList(NpcIndex).CanAttack = 0 Then
        Exit Sub
    End If

    If UserList(UserIndex).flags.Invisible > 0 Or UserList(UserIndex).flags.Oculto > 0 Then
        Exit Sub
    End If
    
    'Si no se peude usar magia en el mapa, no le deja hacerlo.
    If MapInfo(UserList(UserIndex).Pos.map).MagiaSinEfecto > 0 Then
        Exit Sub
    End If
    
    NpcList(NpcIndex).CanAttack = 0
    Dim Danio As Integer
    
    With UserList(UserIndex)
    
        If Hechizos(Spell).SubeHP = 1 Then
        
            If .Stats.MinHP <> .Stats.MaxHP Then
                Danio = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
                .Stats.MinHP = .Stats.MinHP + Danio
                Call WriteUserDamaged(UserIndex, NpcList(NpcIndex).Char.CharIndex, Danio, 2)
            End If
            
            Exit Sub
            
        ElseIf Hechizos(Spell).SubeHP = 2 Then
        
            Danio = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
            
            If .Inv.Head > 0 Then
                Danio = Danio - RandomNumber(ObjData(.Inv.Head).MinDefM, ObjData(.Inv.Head).MaxDefM)
            End If
            
            If .Inv.Ring > 0 Then
                Danio = Danio - RandomNumber(ObjData(.Inv.Ring).MinDefM, ObjData(.Inv.Ring).MaxDefM)
            End If
            
            If Danio < 0 Then
                Danio = 0
            End If
    
            If .flags.Privilegios And PlayerType.User Then
                .Stats.MinHP = .Stats.MinHP - Danio
            End If
            
            Call WriteUserDamaged(UserIndex, NpcList(NpcIndex).Char.CharIndex, Danio, 1)
                    
            If .Stats.MinHP < 1 Then
                If NpcList(NpcIndex).MaestroUser > 0 Then
                    Call UserDie(UserIndex, NpcList(NpcIndex).MaestroUser, NpcIndex)
                Else
                    Call UserDie(UserIndex, , NpcIndex)
                End If
            End If
        End If
    
        If Hechizos(Spell).Paraliza > 0 Then
            If .flags.Paralizado < 1 Then
                If .Inv.Ring = SUPERANILLO Then
                    Call WriteConsoleMsg(UserIndex, "Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                Else
                    .flags.Paralizado = 1
                    .flags.Inmovilizado = 0
                    
                    .Counters.Paralisis = IntervaloParalizado
                    
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, .Pos.x, .Pos.y))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Pos.x, .Pos.y, Hechizos(Spell).FXgrh, Hechizos(Spell).Loops))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetParalized(.Char.CharIndex, 1))
                    
                    'Call WritePosUpdate(UserIndex)
                End If
            End If
        ElseIf Hechizos(Spell).Inmoviliza > 0 Then
            If Not .flags.Inmovilizado > 0 Then
                If .Inv.Ring = SUPERANILLO Then
                    Call WriteConsoleMsg(UserIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                Else
                    .flags.Inmovilizado = 1
                    .flags.Paralizado = 0
                    
                    .Counters.Paralisis = IntervaloParalizado
                    
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, .Pos.x, .Pos.y))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Pos.x, .Pos.y, Hechizos(Spell).FXgrh, Hechizos(Spell).Loops))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetParalized(.Char.CharIndex, 1))
                    
                    'Call WritePosUpdate(UserIndex)
                End If
            End If
        End If
    
        If Hechizos(Spell).Estupidez > 0 Then 'turbacion
            If .flags.Estupidez = 0 Then
                If .Inv.Ring = SUPERANILLO Then
                    Call WriteConsoleMsg(UserIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                Else
                    .flags.Estupidez = 1
                    .Counters.Ceguera = IntervaloInvisible
                
                    Call WriteDumb(UserIndex)
                End If
            End If
        End If
    
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, .Pos.x, .Pos.y))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Pos.x, .Pos.y, Hechizos(Spell).FXgrh, Hechizos(Spell).Loops))
    
        Call CheckPets(NpcIndex, UserIndex)
    End With

End Sub


Public Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNpc As Integer, ByVal Spell As Integer)

    If NpcList(NpcIndex).CanAttack = 0 Then
        Exit Sub
    End If
    
    NpcList(NpcIndex).CanAttack = 0
    
    Dim Danio As Integer
    
    With NpcList(TargetNpc)
        
        If Hechizos(Spell).SubeHP = 2 Then
            Danio = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
            Call SendData(SendTarget.ToNPCArea, TargetNpc, PrepareMessagePlayWave(Hechizos(Spell).WAV, .Pos.x, .Pos.y))
            Call SendData(SendTarget.ToNPCArea, TargetNpc, PrepareMessageCreateFX(.Pos.x, .Pos.y, Hechizos(Spell).FXgrh, Hechizos(Spell).Loops))
            
            .Stats.MinHP = .Stats.MinHP - Danio
            
            If NpcList(NpcIndex).MaestroUser > 0 Then
                Call CalcularDarExp(NpcList(NpcIndex).MaestroUser, TargetNpc, Danio)
            End If
            
            'Muere
            If .Stats.MinHP < 1 Then
                If NpcList(NpcIndex).MaestroUser > 0 Then
                    Call MuereNpc(TargetNpc, NpcList(NpcIndex).MaestroUser)
                Else
                    Call MuereNpc(TargetNpc, 0)
                End If
            End If
    
        ElseIf Hechizos(Spell).Paraliza > 0 Then
            If .flags.AfectaParalisis = 0 Then
                If .flags.Paralizado < 1 Then
                    .flags.Paralizado = 1
                    .flags.Inmovilizado = 0
                    .Contadores.Paralisis = IntervaloParalizado
                    
                    Call SendData(SendTarget.ToNPCArea, TargetNpc, PrepareMessagePlayWave(Hechizos(Spell).WAV, .Pos.x, .Pos.y))
                    Call SendData(SendTarget.ToNPCArea, TargetNpc, PrepareMessageCreateFX(.Pos.x, .Pos.y, Hechizos(Spell).FXgrh, Hechizos(Spell).Loops))
                    Call SendData(SendTarget.ToNPCArea, TargetNpc, PrepareMessageSetParalized(.Char.CharIndex, 1))
                End If
            End If
            
        ElseIf Hechizos(Spell).Inmoviliza > 0 Then
            If .flags.AfectaParalisis = 0 Then
                If Not .flags.Inmovilizado > 0 Then
                    .flags.Inmovilizado = 1
                    .flags.Paralizado = 0
                    .Contadores.Paralisis = IntervaloParalizado
                    
                    Call SendData(SendTarget.ToNPCArea, TargetNpc, PrepareMessagePlayWave(Hechizos(Spell).WAV, .Pos.x, .Pos.y))
                    Call SendData(SendTarget.ToNPCArea, TargetNpc, PrepareMessageCreateFX(.Pos.x, .Pos.y, Hechizos(Spell).FXgrh, Hechizos(Spell).Loops))
                    Call SendData(SendTarget.ToNPCArea, TargetNpc, PrepareMessageSetParalized(.Char.CharIndex, 1))
                End If
            End If
        End If
    
    End With
End Sub

Public Function TieneHechizo(ByVal UserIndex As Integer, ByVal i As Integer) As Boolean
    Dim j As Integer
    
    For j = 1 To MaxSpellSlots
        If UserList(UserIndex).Spells.Spell(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next j
End Function

Public Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte)

    Dim hIndex As Integer
    Dim j As Integer
    
    With UserList(UserIndex)
        hIndex = ObjData(.Inv.Obj(Slot).index).SpellIndex
    
        If Not TieneHechizo(UserIndex, hIndex) Then
            'Buscamos un slot vacio
            For j = 1 To MaxSpellSlots
                If .Spells.Spell(j) = 0 Then
                    Exit For
                End If
            Next j
                
            If .Spells.Spell(j) > 0 Then
                Call WriteConsoleMsg(UserIndex, "No tenés espacio para más hechizos.", FontTypeNames.FONTTYPE_INFO)
            Else
                .Spells.Spell(j) = hIndex
               .Spells.Nro = .Spells.Nro + 1
                Call WriteSpellSlot(UserIndex, j)
                Call QuitarInvItem(UserIndex, Slot)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Ya tenés ese hechizo.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With

End Sub

Public Function PuedeLanzar(ByVal UserIndex As Integer, ByVal SpellIndex As Integer) As Boolean

    Dim DruidManaBonus As Single

    With UserList(UserIndex)
    If .Stats.Muerto Then
        PuedeLanzar = False
        Exit Function
    End If
        
    If Hechizos(SpellIndex).NeedStaff > 0 Then
            If .Clase = eClass.Mage Then
                If .Inv.RightHand > 0 Then
                    If ObjData(.Inv.RightHand).StaffPower < Hechizos(SpellIndex).NeedStaff Then
                    Call WriteConsoleMsg(UserIndex, "Para lanzar este hechizo necesitás un báculo más poderoso.", FontTypeNames.FONTTYPE_INFO)
                    PuedeLanzar = False
                    Exit Function
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Para lanzar este hechizo necesitás un báculo.", FontTypeNames.FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
            End If
        End If
    End If
    
    If .flags.Privilegios And PlayerType.User Then
    
        If .Skills.Skill(eSkill.Magia).Elv < Hechizos(SpellIndex).MinSkill Then
            Call WriteConsoleMsg(UserIndex, "Para lanzar este hechizo necesitás " & Hechizos(SpellIndex).MinSkill & " puntos de habilidad en Magia.", FontTypeNames.FONTTYPE_INFO)
            PuedeLanzar = False
            Exit Function
        End If
        
        If .Stats.MinSta < Hechizos(SpellIndex).StaRequerido Then
            Call WriteConsoleMsg(UserIndex, "No tenés suficiente energía para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
            PuedeLanzar = False
            Exit Function
        End If
    End If
    
    If .Clase = eClass.Druid Then
        If .Inv.Ring = FLAUTAMAGICA Then
            If Hechizos(SpellIndex).Mimetiza Then
                DruidManaBonus = 0.5
            ElseIf Hechizos(SpellIndex).Tipo = uInvocacion Then
                DruidManaBonus = 0.7
            Else
                DruidManaBonus = 1
            End If
        Else
            DruidManaBonus = 1
        End If
    Else
        DruidManaBonus = 1
    End If
    
    If .flags.Privilegios And PlayerType.User Then
        If .Stats.MinMan < Hechizos(SpellIndex).ManaRequerido * DruidManaBonus Then
            Call WriteConsoleMsg(UserIndex, "No tenés suficiente maná para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
            PuedeLanzar = False
            Exit Function
        ElseIf Hechizos(SpellIndex).Warp = 1 Then
            If .Stats.MinMan <> .Stats.MaxMan Then
                Call WriteConsoleMsg(UserIndex, "Para lanzar este hechizo necesitás tener el maná lleno.", FontTypeNames.FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
            End If
        End If
    End If
        
    End With
    PuedeLanzar = True
End Function

Public Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean)

    Dim PosCasteadaX As Integer
    Dim PosCasteadaY As Integer
    Dim PosCasteadaM As Integer
    Dim H As Integer
    Dim TempX As Integer
    Dim TempY As Integer
    
    With UserList(UserIndex)
    
        PosCasteadaX = .flags.TargetX
        PosCasteadaY = .flags.TargetY
        PosCasteadaM = .flags.TargetMap
        
        H = .Spells.Spell(.flags.Hechizo)
        
        If Hechizos(H).Revivir = 1 Then
        
            If ObjData(maps(PosCasteadaM).mapData(PosCasteadaX, PosCasteadaY).ObjInfo.index).Type <> otCuerpoMuerto Then
                Exit Sub
            End If
            
            If maps(PosCasteadaM).mapData(PosCasteadaX, PosCasteadaY).ObjInfo.Amount > Max_Integer_Value Then
                Exit Sub
            End If
            
            Dim TargetIndex As Integer

            TargetIndex = maps(PosCasteadaM).mapData(PosCasteadaX, PosCasteadaY).ObjInfo.Amount
            
            If TargetIndex < 1 Then
                HechizoCasteado = False
                Exit Sub
            End If
            
            If UserList(TargetIndex).Stats.Muerto Then
                'Seguro de resurreccion (solo afecta a los hechizos, no al sacerdote ni al comando de GM)
                'If UserList(TargetIndex).flags.SeguroResu Then
                '    Call WriteConsoleMsg(UserIndex, "¡El espíritu no tiene intenciones de regresar al mundo de los vivos!", FontTypeNames.FONTTYPE_INFO)
                '    HechizoCasteado = False
                '    EXIT SUB
                'End If
        
                'No usar resu en mapas con ResuSinEfecto
                If MapInfo(UserList(TargetIndex).Pos.map).ResuSinEfecto > 0 Then
                    Call WriteConsoleMsg(UserIndex, "¡Revivir no está permitido aquí! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                    Exit Sub
                End If
            
                'No podemos resucitar si nuestra barra de energía no está llena. (GD: 29\04\07)
                If .Stats.MaxSta <> .Stats.MinSta Then
                    Call WriteConsoleMsg(UserIndex, "No podés resucitar si no tenés tu barra de energía llena.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                    Exit Sub
                End If
            
                'revisamos si necesita vara
                If .Clase = eClass.Mage Then
                    If .Inv.RightHand > 0 Then
                        If ObjData(.Inv.RightHand).StaffPower < Hechizos(H).NeedStaff Then
                            Call WriteConsoleMsg(UserIndex, "Necesitás un báculo mejor para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                            HechizoCasteado = False
                            Exit Sub
                        End If
                    End If
                    
                ElseIf .Clase = eClass.Bard Then
                    If .Inv.Ring <> LAUDMAGICO Then
                        Call WriteConsoleMsg(UserIndex, "Necesitás un instrumento mágico para devolver la vida.", FontTypeNames.FONTTYPE_INFO)
                        HechizoCasteado = False
                        Exit Sub
                    End If
                    
                ElseIf .Clase = eClass.Druid Then
                    If .Inv.Ring <> FLAUTAMAGICA Then
                        Call WriteConsoleMsg(UserIndex, "Necesitás un instrumento mágico para devolver la vida.", FontTypeNames.FONTTYPE_INFO)
                        HechizoCasteado = False
                        Exit Sub
                    End If
                End If
    
                If maps(PosCasteadaM).mapData(PosCasteadaX, PosCasteadaY).Blocked Then
                    Call WriteConsoleMsg(UserIndex, "El objetivo no puede ser revivido porque está sobre espacio bloqueado.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
    
                With UserList(TargetIndex)
                    .Stats.MinSed = 0
                    .Stats.MinHam = 0
                    .Stats.MinMan = 0
                    .Stats.MinSta = 0
                End With
                
                Call WriteUpdateHungerAndThirst(TargetIndex)
                Call InfoHechizo(UserIndex)
                
                'Agregado para quitar la penalización de vida en el ring y cambio de ecuacion. (NicoNZ)
                If (TriggerZonaPelea(UserIndex, TargetIndex) <> TRIGGER6_PERMITE) Then
                    'Solo saco vida si es User. no quiero que exploten GMs por ahi.
                    .Stats.MinHP = .Stats.MinHP * (1 - UserList(TargetIndex).Stats.Elv * 0.015)
                End If
                    
                If .Stats.MinHP < 1 Then
                    Call UserDie(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar fue demasiado grande.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                Else
                    Call WriteUpdateHP(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar te debilitó.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = True
                End If
                
                Call RevivirUsuario(TargetIndex)
            Else
                HechizoCasteado = False
            End If
            
        ElseIf Hechizos(H).RemueveInvisibilidadParcial = 1 Then
        
            HechizoCasteado = True
            
            For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
                For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                    If InMapBounds(PosCasteadaM, TempX, TempY) Then
                        If maps(PosCasteadaM).mapData(TempX, TempY).UserIndex > 0 Then
                            'hay un user
                            If UserList(maps(PosCasteadaM).mapData(TempX, TempY).UserIndex).flags.Invisible > 0 And UserList(maps(PosCasteadaM).mapData(TempX, TempY).UserIndex).flags.AdminInvisible < 1 Then
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(TempX, TempY, Hechizos(H).FXgrh, Hechizos(H).Loops))
                            End If
                        End If
                    End If
                Next TempY
            Next TempX
        
            Call InfoHechizo(UserIndex)
        
        End If
    End With
    
End Sub

Public Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean)

    With UserList(UserIndex)
    'No permitimos se invoquen criaturas en zonas seguras
    If MapInfo(.Pos.map).PK = False Or maps(.Pos.map).mapData(.Pos.x, .Pos.y).Trigger = eTrigger.ZONASEGURA Or maps(.Pos.map).mapData(.Pos.x, .Pos.y).Trigger = eTrigger.EnPlataforma Then
        Call WriteConsoleMsg(UserIndex, "Acá no podés invocar criaturas.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    Dim SpellIndex As Integer, NroNpcs As Integer, NpcIndex As Integer, PetIndex As Byte
    Dim TargetPos As WorldPos

    TargetPos.map = .flags.TargetMap
    TargetPos.x = .flags.TargetX
    TargetPos.y = .flags.TargetY

    SpellIndex = .Spells.Spell(.flags.Hechizo)
    
    'Warp de Mascotas
    If Hechizos(SpellIndex).Warp = 1 Then
        If .Pets.NroALaVez > 0 Then
            PetIndex = FarthestPet(UserIndex)
        
            'La invoco cerca mio
            If PetIndex > 0 Then
                Call WarpMascota(UserIndex, PetIndex)
            End If
        End If
        
    'Invocacion normal
    Else
        If Hechizos(SpellIndex).NumNpc < 1 Then
            Exit Sub
        End If
        
        If Hechizos(SpellIndex).Cant = 1 Then
            If .Pets.NroALaVez < MaxPetsALaVez Then
                
                NpcIndex = SpawnNpc(Hechizos(SpellIndex).NumNpc, TargetPos, True, False, False)
                                
                If NpcIndex > 0 Then
                    
                    PetIndex = FreeMascotaIndex(UserIndex)
                    
                    .Pets.Pet(PetIndex).index = NpcIndex
                    .Pets.Pet(PetIndex).Tipo = NpcList(NpcIndex).Numero
                    
                    .Pets.Pet(PetIndex).Lvl = NpcList(NpcIndex).Lvl
                    
                    .Pets.Pet(PetIndex).MinHP = NpcList(NpcIndex).Stats.MinHP
                    .Pets.Pet(PetIndex).MaxHP = NpcList(NpcIndex).Stats.MaxHP
                    
                    .Pets.Pet(PetIndex).MinHit = NpcList(NpcIndex).Stats.MinHit
                    .Pets.Pet(PetIndex).MaxHit = NpcList(NpcIndex).Stats.MaxHit
                    
                    .Pets.Pet(PetIndex).Def = NpcList(NpcIndex).Stats.Def
                    .Pets.Pet(PetIndex).DefM = NpcList(NpcIndex).Stats.DefM
            
                    With NpcList(NpcIndex)
                        .MaestroUser = UserIndex
                        .Contadores.TiempoExistencia = IntervaloInvocacion
                    End With
            
                    Call FollowAmo(NpcIndex)
                Else
                    Exit Sub
                End If

                .Pets.NroALaVez = .Pets.NroALaVez + 1
            
            Else
                For NroNpcs = 1 To MaxPets
                    If .Pets.Pet(NroNpcs).index > 0 Then
                        If NpcList(.Pets.Pet(NroNpcs).index).Contadores.TiempoExistencia > 0 Then
                            If .Pets.Pet(NroNpcs).Tipo = Hechizos(SpellIndex).NumNpc Then
                                If PetIndex > 0 Then
                                    If NpcList(.Pets.Pet(NroNpcs).index).Contadores.TiempoExistencia < NpcList(.Pets.Pet(PetIndex).index).Contadores.TiempoExistencia Then
                                        PetIndex = NroNpcs
                                    End If
                                Else
                                    PetIndex = NroNpcs
                                End If
                            End If
                        End If
                    End If
                Next NroNpcs
                
                If PetIndex < 1 Then
                    Exit Sub
                End If
                
                Call QuitarNpc(.Pets.Pet(PetIndex).index)
                
                NpcIndex = SpawnNpc(Hechizos(SpellIndex).NumNpc, TargetPos, True, False, False)
                
                If NpcIndex > 0 Then
                    
                    PetIndex = FreeMascotaIndex(UserIndex)
                    
                    .Pets.Pet(PetIndex).index = NpcIndex
                    .Pets.Pet(PetIndex).Tipo = NpcList(NpcIndex).Numero
                    
                    .Pets.Pet(PetIndex).Lvl = NpcList(NpcIndex).Lvl
                    
                    .Pets.Pet(PetIndex).MinHP = NpcList(NpcIndex).Stats.MinHP
                    .Pets.Pet(PetIndex).MaxHP = NpcList(NpcIndex).Stats.MaxHP
                    
                    .Pets.Pet(PetIndex).MinHit = NpcList(NpcIndex).Stats.MinHit
                    .Pets.Pet(PetIndex).MaxHit = NpcList(NpcIndex).Stats.MaxHit
                    
                    .Pets.Pet(PetIndex).Def = NpcList(NpcIndex).Stats.Def
                    .Pets.Pet(PetIndex).DefM = NpcList(NpcIndex).Stats.DefM
            
                    With NpcList(NpcIndex)
                        .MaestroUser = UserIndex
                        .Contadores.TiempoExistencia = IntervaloInvocacion
                    End With
            
                    Call FollowAmo(NpcIndex)

                Else
                    .Pets.NroALaVez = .Pets.NroALaVez - 1
                    Exit Sub
                End If
            End If
            
        Else
          
            For NroNpcs = 1 To Hechizos(SpellIndex).Cant
                
                If .Pets.NroALaVez < MaxPetsALaVez Then
                
                    NpcIndex = SpawnNpc(Hechizos(SpellIndex).NumNpc, TargetPos, True, False, False)
                    
                    If NpcIndex > 0 Then
                        .Pets.NroALaVez = .Pets.NroALaVez + 1
                        
                        PetIndex = FreeMascotaIndex(UserIndex)
                        
                        .Pets.Pet(PetIndex).index = NpcIndex
                        .Pets.Pet(PetIndex).Tipo = NpcList(NpcIndex).Numero
                        
                        .Pets.Pet(PetIndex).Lvl = NpcList(NpcIndex).Lvl
                        
                        .Pets.Pet(PetIndex).MinHP = NpcList(NpcIndex).Stats.MinHP
                        .Pets.Pet(PetIndex).MaxHP = NpcList(NpcIndex).Stats.MaxHP
                        
                        .Pets.Pet(PetIndex).MinHit = NpcList(NpcIndex).Stats.MinHit
                        .Pets.Pet(PetIndex).MaxHit = NpcList(NpcIndex).Stats.MaxHit
                        
                        .Pets.Pet(PetIndex).Def = NpcList(NpcIndex).Stats.Def
                        .Pets.Pet(PetIndex).DefM = NpcList(NpcIndex).Stats.DefM
                
                        With NpcList(NpcIndex)
                            .MaestroUser = UserIndex
                            .Contadores.TiempoExistencia = IntervaloInvocacion
                        End With
                
                        Call FollowAmo(NpcIndex)
                    Else
                        Exit Sub
                    End If
                Else
                
                    Dim Numero As Byte
                    
                    For Numero = 1 To MaxPets
                        If .Pets.Pet(Numero).index > 0 Then
                            If NpcList(.Pets.Pet(Numero).index).Contadores.TiempoExistencia > 0 Then
                                If .Pets.Pet(Numero).Tipo = Hechizos(SpellIndex).NumNpc Then
                                    If PetIndex > 0 Then
                                        If NpcList(.Pets.Pet(Numero).index).Contadores.TiempoExistencia < NpcList(.Pets.Pet(PetIndex).index).Contadores.TiempoExistencia Then
                                            PetIndex = Numero
                                        End If
                                    Else
                                        PetIndex = Numero
                                    End If
                                End If
                            End If
                        End If
                    Next Numero
                    
                    If PetIndex < 1 Then
                        Exit Sub
                    End If
                    
                    Call QuitarNpc(.Pets.Pet(PetIndex).index)
                    
                    NpcIndex = SpawnNpc(Hechizos(SpellIndex).NumNpc, TargetPos, True, False, False)
                    
                    If NpcIndex > 0 Then
                        
                        PetIndex = FreeMascotaIndex(UserIndex)
                        
                        .Pets.Pet(PetIndex).index = NpcIndex
                        .Pets.Pet(PetIndex).Tipo = NpcList(NpcIndex).Numero
                        
                        .Pets.Pet(PetIndex).Lvl = NpcList(NpcIndex).Lvl
                        
                        .Pets.Pet(PetIndex).MinHP = NpcList(NpcIndex).Stats.MinHP
                        .Pets.Pet(PetIndex).MaxHP = NpcList(NpcIndex).Stats.MaxHP
                        
                        .Pets.Pet(PetIndex).MinHit = NpcList(NpcIndex).Stats.MinHit
                        .Pets.Pet(PetIndex).MaxHit = NpcList(NpcIndex).Stats.MaxHit
                        
                        .Pets.Pet(PetIndex).Def = NpcList(NpcIndex).Stats.Def
                        .Pets.Pet(PetIndex).DefM = NpcList(NpcIndex).Stats.DefM
                
                        With NpcList(NpcIndex)
                            .MaestroUser = UserIndex
                            .Contadores.TiempoExistencia = IntervaloInvocacion
                        End With
                
                        Call FollowAmo(NpcIndex)
    
                    Else
                        .Pets.NroALaVez = .Pets.NroALaVez - 1
                        Exit Sub
                    End If
                End If
                
                Next NroNpcs
            End If
        End If
    End With
    
    Call InfoHechizo(UserIndex)
    HechizoCasteado = True

End Sub

Public Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal SpellIndex As Integer)

    Dim HechizoCasteado As Boolean

    Select Case Hechizos(SpellIndex).Tipo
    
        Case TipoHechizo.uInvocacion '
            Call HechizoInvocacion(UserIndex, HechizoCasteado)
            
        Case TipoHechizo.uEstado
            Call HechizoTerrenoEstado(UserIndex, HechizoCasteado)
            
    End Select

    If HechizoCasteado Then
    
        With UserList(UserIndex)
            Call SubirSkill(UserIndex, eSkill.Magia, True)
            
            If Hechizos(SpellIndex).Warp = 1 Then 'Invoco un Mascota
                'Consume toda la mana
                .Stats.MinMan = 0
            Else
                If .Clase = eClass.Druid And .Inv.Ring = FLAUTAMAGICA Then
                    .Stats.MinMan = .Stats.MinMan - Hechizos(SpellIndex).ManaRequerido * 0.7
                Else
                    .Stats.MinMan = .Stats.MinMan - Hechizos(SpellIndex).ManaRequerido
                End If

                If .Stats.MinMan < 0 Then .Stats.MinMan = 0
            End If
            
                .Stats.MinSta = .Stats.MinSta - Hechizos(SpellIndex).StaRequerido
        End With
        Call WriteUpdateSta(UserIndex)
        Call WriteUpdateMana(UserIndex)
    End If

End Sub

Public Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal SpellIndex As Integer)

    Dim HechizoCasteado As Boolean
    Select Case Hechizos(SpellIndex).Tipo
        Case TipoHechizo.uEstado 'Afectan estados (por ejem: Envenenamiento)
            Call HechizoEstadoUsuario(UserIndex, HechizoCasteado)
        
        Case TipoHechizo.uPropiedades 'Afectan HP,MANA,STAMINA,ETC
            Call HechizoPropUsuario(UserIndex, HechizoCasteado)
    End Select

    If HechizoCasteado Then
        With UserList(UserIndex)
            'Agregado para que los druidas, al tener equipada la flauta magica, el coste de mana de mimetismo es de 50% menos.
            If .Clase = eClass.Druid And .Inv.Ring = FLAUTAMAGICA And Hechizos(SpellIndex).Mimetiza = 1 Then
                .Stats.MinMan = .Stats.MinMan - Hechizos(SpellIndex).ManaRequerido * 0.5
            Else
                .Stats.MinMan = .Stats.MinMan - Hechizos(SpellIndex).ManaRequerido
            End If
        
            .Stats.MinSta = .Stats.MinSta - Hechizos(SpellIndex).StaRequerido
            Call WriteUpdateSta(UserIndex)
            Call WriteUpdateMana(UserIndex)
            .flags.TargetUser = 0
        End With
        
        Call SubirSkill(UserIndex, eSkill.Magia, True)
    End If

End Sub

Public Sub HandleHechizoNpc(ByVal UserIndex As Integer, ByVal SpellIndex As Integer)

    Dim b As Boolean
    
    With UserList(UserIndex)
        Select Case Hechizos(SpellIndex).Tipo
            Case TipoHechizo.uEstado 'Afectan estados (por ejem: Envenenamiento)
                Call HechizoEstadoNpc(.flags.TargetNpc, SpellIndex, b, UserIndex)
            Case TipoHechizo.uPropiedades 'Afectan HP,MANA,STAMINA,ETC
                Call HechizoPropNpc(SpellIndex, .flags.TargetNpc, UserIndex, b)
        End Select
    
        If b Then
            Call SubirSkill(UserIndex, eSkill.Magia, True)
            .flags.TargetNpc = 0
            
            'Bonificación para druidas.
            If .Clase = eClass.Druid And .Inv.Ring = FLAUTAMAGICA And Hechizos(SpellIndex).Mimetiza = 1 Then
                .Stats.MinMan = .Stats.MinMan - Hechizos(SpellIndex).ManaRequerido * 0.5
            Else
                .Stats.MinMan = .Stats.MinMan - Hechizos(SpellIndex).ManaRequerido
            End If
        
            .Stats.MinSta = .Stats.MinSta - Hechizos(SpellIndex).StaRequerido
            Call WriteUpdateMana(UserIndex)
            Call WriteUpdateSta(UserIndex)
        End If
    End With

End Sub

Public Sub LanzarHechizo(Hechizo As Integer, UserIndex As Integer)

On Error GoTo ErrHandler

    Dim SpellIndex As Integer
    
    With UserList(UserIndex)
    
        SpellIndex = .Spells.Spell(Hechizo)
    
        If PuedeLanzar(UserIndex, SpellIndex) Then
        
            Select Case Hechizos(SpellIndex).Target
                
                Case TargetType.uUsuarios
                
                    If .flags.TargetUser > 0 Then
                        If Abs(UserList(.flags.TargetUser).Pos.y - .Pos.y) <= RANGO_VISION_Y Then
                            Call HandleHechizoUsuario(UserIndex, SpellIndex)
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "Este hechizo actúa sólo sobre personas.", FontTypeNames.FONTTYPE_INFO)
                    End If
                
                Case TargetType.uNpc
                
                    If .flags.TargetNpc > 0 Then
                        If Abs(NpcList(.flags.TargetNpc).Pos.y - .Pos.y) <= RANGO_VISION_Y Then
                            Call HandleHechizoNpc(UserIndex, SpellIndex)
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "Este hechizo sólo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)
                    End If
            
                Case TargetType.uUsuariosYnpc
                
                    If .flags.TargetUser > 0 Then
                        If Abs(UserList(.flags.TargetUser).Pos.y - .Pos.y) <= RANGO_VISION_Y Then
                            Call HandleHechizoUsuario(UserIndex, SpellIndex)
                        End If
                    ElseIf .flags.TargetNpc > 0 Then
                        If Abs(NpcList(.flags.TargetNpc).Pos.y - .Pos.y) <= RANGO_VISION_Y Then
                            Call HandleHechizoNpc(UserIndex, SpellIndex)
                        End If
                    End If
                
                Case TargetType.uTerreno
                    Call HandleHechizoTerreno(UserIndex, SpellIndex)
            
            End Select
            
        End If
    
        If .Counters.Trabajando Then
            .Counters.Trabajando = .Counters.Trabajando - 1
        End If
        
        If .Counters.Ocultando Then
            .Counters.Ocultando = .Counters.Ocultando - 1
        End If
        
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en LanzarHechizo. Error " & Err.Number & ": " & Err.description)
    
End Sub

Public Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean)

    Dim SpellIndex As Integer
    Dim TargetIndex As Integer

    With UserList(UserIndex)
        SpellIndex = .Spells.Spell(.flags.Hechizo)
        TargetIndex = .flags.TargetUser
    
        If Hechizos(SpellIndex).Invisibilidad = 1 Then
            If UserList(TargetIndex).Stats.Muerto Then
                HechizoCasteado = False
                Exit Sub
            End If
        
            If UserList(TargetIndex).Counters.Saliendo Then
                If UserIndex <> TargetIndex Then
                    Call WriteConsoleMsg(UserIndex, "El hechizo no tiene efecto.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                    Exit Sub
                Else
                    Call WriteConsoleMsg(UserIndex, "No podés hacerte invisible mientras estás saliendo.", FontTypeNames.FONTTYPE_WARNING)
                    HechizoCasteado = False
                    Exit Sub
                End If
            End If
        
            'No usar invi mapas InviSinEfecto
            If MapInfo(UserList(TargetIndex).Pos.map).InviSinEfecto > 0 Then
                Call WriteConsoleMsg(UserIndex, "La invisibilidad no funciona aquí.", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub
            End If
            
            'Si sos user, no uses este hechizo con GMS.
            If .flags.Privilegios And PlayerType.User Then
                If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then
                    Exit Sub
                End If
            End If
       
            UserList(TargetIndex).flags.Invisible = 1
            Call SendData(SendTarget.ToPCArea, TargetIndex, PrepareMessageSetInvisible(UserList(TargetIndex).Char.CharIndex, True))

            HechizoCasteado = True
        End If
    
        If Hechizos(SpellIndex).Mimetiza = 1 Then
            If UserList(TargetIndex).Stats.Muerto Then
                Exit Sub
            End If
        
            If UserList(TargetIndex).flags.Navegando Then
                Exit Sub
            End If
        
            If .flags.Navegando Then
                Exit Sub
            End If
        
            'Si sos user, no uses este hechizo con GMS.
            If .flags.Privilegios And PlayerType.User Then
                If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then
                    Exit Sub
                End If
            End If
        
            If .flags.Mimetizado Then
                Call WriteConsoleMsg(UserIndex, "Ya estás mimetizado. El hechizo no hizo efecto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .flags.AdminInvisible > 0 Then
                Exit Sub
            End If
        
            'copio el char original al mimetizado
        
            .CharMimetizado.Body = .Char.Body
            .CharMimetizado.Head = .Char.Head
            .CharMimetizado.HeadAnim = .Char.HeadAnim
            .CharMimetizado.ShieldAnim = .Char.ShieldAnim
            .CharMimetizado.WeaponAnim = .Char.WeaponAnim
            
            .flags.Mimetizado = True
            
            'ahora pongo local el del enemigo
            .Char.Body = UserList(TargetIndex).Char.Body
            .Char.Head = UserList(TargetIndex).Char.Head
            .Char.HeadAnim = UserList(TargetIndex).Char.HeadAnim
            .Char.ShieldAnim = UserList(TargetIndex).Char.ShieldAnim
            
            .Char.WeaponAnim = GetWeaponAnim(UserIndex)
            
            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.HeadAnim)
       
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True
        End If
    
        If Hechizos(SpellIndex).Envenena = 1 Then
            If UserIndex = TargetIndex Then
                Exit Sub
            End If
        
            If Not PuedeAtacar(UserIndex, TargetIndex) Then
                Exit Sub
            End If
            
            If UserIndex <> TargetIndex Then
                Call UserAtacadoPorUsuario(UserIndex, TargetIndex)
            End If
            
            UserList(TargetIndex).flags.Envenenado = 1
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True
        End If
    
        If Hechizos(SpellIndex).CuraVeneno > 0 Then
    
            'Verificamos que el usuario no este muerto
            If UserList(TargetIndex).Stats.Muerto Then
                Call WriteConsoleMsg(UserIndex, UserList(TargetIndex).name & " está muerto!", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub
            End If

            'Si sos user, no uses este hechizo con GMS.
            If .flags.Privilegios And PlayerType.User Then
                If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then
                    Exit Sub
                End If
            End If
            
            UserList(TargetIndex).flags.Envenenado = 0
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True
        End If
    
        If Hechizos(SpellIndex).Maldicion > 0 Then
            If UserIndex = TargetIndex Then
                Exit Sub
            End If
        
            If Not PuedeAtacar(UserIndex, TargetIndex) Then
                Exit Sub
            End If
            
            If UserIndex <> TargetIndex Then
                Call UserAtacadoPorUsuario(UserIndex, TargetIndex)
            End If
            
            UserList(TargetIndex).flags.Maldicion = 1
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True
        End If
    
        If Hechizos(SpellIndex).RemoverMaldicion > 0 Then
            UserList(TargetIndex).flags.Maldicion = 1
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True
        End If
    
        If Hechizos(SpellIndex).Bendicion > 0 Then
            UserList(TargetIndex).flags.Bendicion = 1
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True
        End If
    
        If Hechizos(SpellIndex).Paraliza = 1 Or Hechizos(SpellIndex).Inmoviliza = 1 Then
            If UserIndex = TargetIndex Then
                Exit Sub
            End If
        
             If UserList(TargetIndex).flags.Paralizado < 1 Then
                If Not PuedeAtacar(UserIndex, TargetIndex) Then
                    Exit Sub
                End If
                
                If UserIndex <> TargetIndex Then
                    Call UserAtacadoPorUsuario(UserIndex, TargetIndex)
                End If
                
                Call InfoHechizo(UserIndex)
                
                HechizoCasteado = True
                
                If UserList(TargetIndex).Inv.Ring = SUPERANILLO Then
                    Call WriteConsoleMsg(TargetIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
                    Call FlushBuffer(TargetIndex)
                    Exit Sub
                End If
            
                If Hechizos(SpellIndex).Inmoviliza = 1 Then
                    UserList(TargetIndex).flags.Inmovilizado = 1
                End If
                
                UserList(TargetIndex).flags.Paralizado = 1
                UserList(TargetIndex).Counters.Paralisis = IntervaloParalizado
                Call SendData(SendTarget.ToPCArea, TargetIndex, PrepareMessageSetParalized(UserList(TargetIndex).Char.CharIndex, 1))
                
                Call WritePosUpdate(TargetIndex)
            End If
        End If
        
        If Hechizos(SpellIndex).RemoverParalisis > 0 Then
            If UserList(TargetIndex).flags.Paralizado > 0 Then
                UserList(TargetIndex).flags.Inmovilizado = 0
                UserList(TargetIndex).flags.Paralizado = 0
                'no need to crypt this
                Call SendData(SendTarget.ToPCArea, TargetIndex, PrepareMessageSetParalized(UserList(TargetIndex).Char.CharIndex, 0))
                Call InfoHechizo(UserIndex)
                HechizoCasteado = True
            End If
        End If
    
        If Hechizos(SpellIndex).RemoverEstupidez > 0 Then
            If UserList(TargetIndex).flags.Estupidez > 0 Then
                UserList(TargetIndex).flags.Estupidez = 0
                'no need to crypt this
                Call WriteDumbNoMore(TargetIndex)
                Call FlushBuffer(TargetIndex)
                Call InfoHechizo(UserIndex)
                HechizoCasteado = True
            End If
        End If
    
        If Hechizos(SpellIndex).Ceguera = 1 Then
            If UserIndex = TargetIndex Then
                Exit Sub
            End If
        
            If Not PuedeAtacar(UserIndex, TargetIndex) Then
                Exit Sub
            End If
            
            If UserIndex <> TargetIndex Then
                Call UserAtacadoPorUsuario(UserIndex, TargetIndex)
            End If
            
            UserList(TargetIndex).flags.Ceguera = 1
            UserList(TargetIndex).Counters.Ceguera = IntervaloParalizado \ 3
    
            Call WriteBlind(TargetIndex)
            
            Call FlushBuffer(TargetIndex)
            
            Call InfoHechizo(UserIndex)
            
            HechizoCasteado = True
        End If
    
        If Hechizos(SpellIndex).Estupidez = 1 Then
            If UserIndex = TargetIndex Then
                Exit Sub
            End If
        
            If Not PuedeAtacar(UserIndex, TargetIndex) Then
                Exit Sub
            End If
            
            If UserIndex <> TargetIndex Then
                Call UserAtacadoPorUsuario(UserIndex, TargetIndex)
            End If
                    
            If UserList(TargetIndex).flags.Estupidez = 0 Then
                UserList(TargetIndex).flags.Estupidez = 1
                UserList(TargetIndex).Counters.Ceguera = IntervaloParalizado
            End If
                
            Call WriteDumb(TargetIndex)
            Call FlushBuffer(TargetIndex)
        
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True
        End If
    End With
End Sub

Public Sub HechizoEstadoNpc(ByVal NpcIndex As Integer, ByVal SpellIndex As Integer, ByRef HechizoCasteado As Boolean, ByVal UserIndex As Integer)
'Handles the Spells that afect the Stats of an Npc

    With NpcList(NpcIndex)
        If Hechizos(SpellIndex).Invisibilidad = 1 Then
            Call InfoHechizo(UserIndex)
            .flags.Invisible = 1
            HechizoCasteado = True
        End If

        If Hechizos(SpellIndex).Envenena = 1 Then
            If Not PuedeAtacarNpc(UserIndex, NpcIndex) Then
                HechizoCasteado = False
                Exit Sub
            End If
        
            Call NpcAtacado(NpcIndex, UserIndex)
        
            Call CheckPets(NpcIndex, UserIndex)
        
            Call InfoHechizo(UserIndex)
            .flags.Envenenado = 1
            HechizoCasteado = True
        End If

        If Hechizos(SpellIndex).CuraVeneno > 0 Then
            Call InfoHechizo(UserIndex)
            .flags.Envenenado = 0
            HechizoCasteado = True
        End If

        If Hechizos(SpellIndex).Maldicion > 0 Then
            If Not PuedeAtacarNpc(UserIndex, NpcIndex) Then
                HechizoCasteado = False
                Exit Sub
            End If
    
            Call NpcAtacado(NpcIndex, UserIndex)
    
            Call CheckPets(NpcIndex, UserIndex)
    
            Call InfoHechizo(UserIndex)
            .flags.Maldicion = 1
            HechizoCasteado = True
        End If

        If Hechizos(SpellIndex).RemoverMaldicion > 0 Then
            Call InfoHechizo(UserIndex)
            .flags.Maldicion = 0
            HechizoCasteado = True
        End If

        If Hechizos(SpellIndex).Bendicion > 0 Then
            Call InfoHechizo(UserIndex)
            .flags.Bendicion = 1
            HechizoCasteado = True
        End If

        If Hechizos(SpellIndex).Paraliza > 0 Then
            If .flags.AfectaParalisis = 0 Then
                If Not PuedeAtacarNpc(UserIndex, NpcIndex, True) Then
                    HechizoCasteado = False
                    Exit Sub
                End If
                
                Call NpcAtacado(NpcIndex, UserIndex)
                
                Call CheckPets(NpcIndex, UserIndex)
                
                Call InfoHechizo(UserIndex)
                .flags.Paralizado = 1
                .flags.Inmovilizado = 0
                .Contadores.Paralisis = IntervaloParalizado
                HechizoCasteado = True
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageSetParalized(.Char.CharIndex, 1))
            Else
                Call WriteConsoleMsg(UserIndex, "El Npc es inmune a este hechizo.", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub
            End If
        End If

        If Hechizos(SpellIndex).RemoverParalisis > 0 Then
            If .flags.Paralizado > 0 Or .flags.Inmovilizado > 0 Then
                If .MaestroUser = UserIndex Then
                    Call InfoHechizo(UserIndex)
                    .flags.Paralizado = 0
                    .Contadores.Paralisis = 0
                    HechizoCasteado = True
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Este Npc no está paralizado", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub
            End If
        End If
 
        If Hechizos(SpellIndex).Inmoviliza = 1 Then
            If .flags.AfectaParalisis = 0 Then
                If Not PuedeAtacarNpc(UserIndex, NpcIndex, True) Then
                    HechizoCasteado = False
                    Exit Sub
                End If
                
                Call NpcAtacado(NpcIndex, UserIndex)
                
                Call CheckPets(NpcIndex, UserIndex)
                
                .flags.Inmovilizado = 1
                .flags.Paralizado = 0
                .Contadores.Paralisis = IntervaloParalizado
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageSetParalized(.Char.CharIndex, 1))
                Call InfoHechizo(UserIndex)
                HechizoCasteado = True
            Else
                Call WriteConsoleMsg(UserIndex, "El Npc es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With

    If Hechizos(SpellIndex).Mimetiza = 1 Then
        With UserList(UserIndex)
            If .flags.Mimetizado Then
                Call WriteConsoleMsg(UserIndex, "Ya estás mimetizado. El hechizo no hizo efecto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .flags.AdminInvisible > 0 Then
                Exit Sub
            End If

            If .Clase = eClass.Druid Then
                'copio el char original al mimetizado
                    
                .CharMimetizado.Body = .Char.Body
                .CharMimetizado.Head = .Char.Head
                .CharMimetizado.HeadAnim = .Char.HeadAnim
                .CharMimetizado.ShieldAnim = .Char.ShieldAnim
                .CharMimetizado.WeaponAnim = .Char.WeaponAnim
                
                .flags.Mimetizado = True
                
                'ahora pongo lo del Npc.
                .Char.Body = NpcList(NpcIndex).Char.Body
                .Char.Head = NpcList(NpcIndex).Char.Head
                .Char.HeadAnim = NingunCasco
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
    
                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.HeadAnim)
            
            Else
                Call WriteConsoleMsg(UserIndex, "Sólo los druidas pueden mimetizarse con criaturas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            Call InfoHechizo(UserIndex)
            HechizoCasteado = True
        End With
    End If
End Sub

Public Sub HechizoPropNpc(ByVal SpellIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean)
'Handles the Spells that afect the Life Npc

    Dim Danio As Integer
    
    With NpcList(NpcIndex)
        'Salud
        If Hechizos(SpellIndex).SubeHP = 1 Then
        
            Danio = RandomNumber(Hechizos(SpellIndex).MinHP, Hechizos(SpellIndex).MaxHP)
            
            Danio = Danio + Porcentaje(Danio, 2 * UserList(UserIndex).Stats.Elv)
            
            Call InfoHechizo(UserIndex)
            
            .Stats.MinHP = .Stats.MinHP + Danio
            
            If .Stats.MinHP > .Stats.MaxHP Then
                .Stats.MinHP = .Stats.MaxHP
            End If
                        
            Call WriteDamage(UserIndex, .Char.CharIndex, Danio, .Stats.MinHP, .Stats.MaxHP, 2)
            
            HechizoCasteado = True
        
        ElseIf Hechizos(SpellIndex).SubeHP = 2 Then
            If Not PuedeAtacarNpc(UserIndex, NpcIndex) Then
                    HechizoCasteado = False
                Exit Sub
            End If
            
            Call NpcAtacado(NpcIndex, UserIndex)
            
            Call CheckPets(NpcIndex, UserIndex)
    
            Danio = RandomNumber(Hechizos(SpellIndex).MinHP, Hechizos(SpellIndex).MaxHP)
            Danio = Danio + Porcentaje(Danio, 3 * UserList(UserIndex).Stats.Elv)
        
            If Hechizos(SpellIndex).StaffAffected Then
                If UserList(UserIndex).Clase = eClass.Mage Then
                    If UserList(UserIndex).Inv.RightHand > 0 Then
                        Danio = (Danio * (ObjData(UserList(UserIndex).Inv.RightHand).StaffDamageBonus + 70)) \ 100
                        'Aumenta Danio segun el staff-
                        'Danio = (Danio* (70 + BonifBáculo)) \ 100
                    Else
                        Danio = Danio * 0.7 'Baja Danio a 70% del original
                    End If
                End If
            End If
            If UserList(UserIndex).Inv.Ring = LAUDMAGICO Or UserList(UserIndex).Inv.Ring = FLAUTAMAGICA Then
                Danio = Danio * 1.04  'laud magico de los bardos
            End If
        
            Call InfoHechizo(UserIndex)
            
            HechizoCasteado = True
            
            If .flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd2, .Pos.x, .Pos.y))
            End If
            
            Danio = Danio - .Stats.DefM
            
            If Danio < 0 Then
                Danio = 0
            End If
            
            .Stats.MinHP = .Stats.MinHP - Danio
            
            Call WriteDamage(UserIndex, .Char.CharIndex, Danio, .Stats.MinHP, .Stats.MaxHP, 1)
        
            Call CalcularDarExp(UserIndex, NpcIndex, Danio)
        
            If .Stats.MinHP < 1 Then
                Call MuereNpc(NpcIndex, UserIndex)
            End If
        End If
    End With

End Sub

Public Sub InfoHechizo(ByVal UserIndex As Integer)

On Error GoTo ErrorHandler

    Dim SpellIndex As Integer

    With UserList(UserIndex)
        SpellIndex = .Spells.Spell(.flags.Hechizo)
                               
        If .flags.TargetUser > 0 Then
        
            'Los admins invisibles no producen sonidos ni fx's
            If .flags.AdminInvisible > 0 And UserIndex = .flags.TargetUser Then
                Call EnviarDatosASlot(UserIndex, PrepareMessageCreateFX(UserList(.flags.TargetUser).Pos.x, UserList(.flags.TargetUser).Pos.y, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).Loops))
                Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, UserList(.flags.TargetUser).Pos.x, UserList(.flags.TargetUser).Pos.y))
                
                If Not Hechizos(SpellIndex).SubeHP Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessageChatOverHead(Hechizos(SpellIndex).PalabrasMagicas, .Char.CharIndex, vbWhite))
                End If
            Else
                Call SendData(SendTarget.ToPCArea, .flags.TargetUser, PrepareMessageCreateFX(UserList(.flags.TargetUser).Pos.x, UserList(.flags.TargetUser).Pos.y, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).Loops))
                Call SendData(SendTarget.ToPCArea, .flags.TargetUser, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, UserList(.flags.TargetUser).Pos.x, UserList(.flags.TargetUser).Pos.y))  'Esta linea faltaba. Pablo (ToxicWaste)
            
                If Hechizos(SpellIndex).SubeHP Then
                    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageChatOverHead(Hechizos(SpellIndex).PalabrasMagicas, .Char.CharIndex, vbCyan))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Hechizos(SpellIndex).PalabrasMagicas, .Char.CharIndex, vbCyan))
                End If
                    
                'Si estaba oculto o invisible, se vuelve visible
                If .flags.Invisible > 0 Or .flags.Oculto > 0 Then
                    .flags.Oculto = 0
                    .flags.Invisible = 0

                    .Counters.TiempoOculto = 0
                    .Counters.Invisibilidad = 0
                    
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                End If
            End If
            
        ElseIf .flags.TargetNpc > 0 Then
    
            'Los admins invisibles no producen sonidos ni fx's
            If .flags.AdminInvisible > 0 Then
                Call EnviarDatosASlot(UserIndex, PrepareMessageCreateFX(NpcList(.flags.TargetNpc).Pos.x, NpcList(.flags.TargetNpc).Pos.y, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).Loops))
                Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, NpcList(.flags.TargetNpc).Pos.x, NpcList(.flags.TargetNpc).Pos.y))
                
                If Not Hechizos(SpellIndex).SubeHP Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessageChatOverHead(Hechizos(SpellIndex).PalabrasMagicas, .Char.CharIndex, vbWhite))
                End If
            Else
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNpc, PrepareMessageCreateFX(NpcList(.flags.TargetNpc).Pos.x, NpcList(.flags.TargetNpc).Pos.y, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).Loops))
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNpc, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, NpcList(.flags.TargetNpc).Pos.x, NpcList(.flags.TargetNpc).Pos.y)) 'Esta linea faltaba. Pablo (ToxicWaste)
                    
                If Hechizos(SpellIndex).SubeHP Then
                    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageChatOverHead(Hechizos(SpellIndex).PalabrasMagicas, .Char.CharIndex, vbCyan))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Hechizos(SpellIndex).PalabrasMagicas, .Char.CharIndex, vbCyan))
                End If
                    
                'Si estaba oculto o invisible, se vuelve visible
                If .flags.Invisible > 0 Or .flags.Oculto > 0 Then
                    .flags.Oculto = 0
                    .flags.Invisible = 0
                    .Counters.TiempoOculto = 0
                    .Counters.Invisibilidad = 0
                    
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                End If
            End If
        
        Else
            'Los admins invisibles no producen sonidos ni fx's
            If .flags.AdminInvisible > 0 Then
                Call EnviarDatosASlot(UserIndex, PrepareMessageCreateFX(NpcList(.flags.TargetNpc).Pos.x, NpcList(.flags.TargetNpc).Pos.y, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).Loops))
                Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, NpcList(.flags.TargetNpc).Pos.x, NpcList(.flags.TargetNpc).Pos.y))
                
                If Not Hechizos(SpellIndex).SubeHP Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessageChatOverHead(Hechizos(SpellIndex).PalabrasMagicas, .Char.CharIndex, vbWhite))
                End If
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(NpcList(UserIndex).Pos.x, NpcList(UserIndex).Pos.y, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).Loops))
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, NpcList(UserIndex).Pos.x, NpcList(UserIndex).Pos.y)) 'Esta linea faltaba. Pablo (ToxicWaste)
                    
                If Hechizos(SpellIndex).SubeHP Then
                    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageChatOverHead(Hechizos(SpellIndex).PalabrasMagicas, .Char.CharIndex, vbCyan))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Hechizos(SpellIndex).PalabrasMagicas, .Char.CharIndex, vbCyan))
                End If
                    
                'Si estaba oculto o invisible, se vuelve visible
                If .flags.Invisible > 0 Or .flags.Oculto > 0 Then
                    .flags.Oculto = 0
                    .flags.Invisible = 0
                    .Counters.TiempoOculto = 0
                    .Counters.Invisibilidad = 0
                    
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                End If
            End If
        End If
    End With

ErrorHandler:

End Sub

Public Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean)
    
    Dim SpellIndex As Integer
    Dim Danio As Integer
    Dim tempChr As Integer
    
    SpellIndex = UserList(UserIndex).Spells.Spell(UserList(UserIndex).flags.Hechizo)
    tempChr = UserList(UserIndex).flags.TargetUser
          
    With UserList(tempChr)
        If .Stats.Muerto Then
            Exit Sub
        End If
          
        'Hambre
        If Hechizos(SpellIndex).SubeHam = 1 Then
        
            Call InfoHechizo(UserIndex)
        
            Danio = RandomNumber(Hechizos(SpellIndex).MinHam, Hechizos(SpellIndex).MaxHam)
        
            .Stats.MinHam = .Stats.MinHam + Danio
            
            If .Stats.MinHam > 100 Then
                .Stats.MinHam = 100
            End If
            
            If UserIndex <> tempChr Then
                Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Danio & " puntos de Hambre a " & .name, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " restauró " & Danio & " puntos de Hambre.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, "Te restauraste " & Danio & " puntos de Hambre.", FontTypeNames.FONTTYPE_FIGHT)
            End If
        
            Call WriteUpdateHungerAndThirst(tempChr)
            HechizoCasteado = True
        
        ElseIf Hechizos(SpellIndex).SubeHam = 2 Then
            If Not PuedeAtacar(UserIndex, tempChr) Then
                Exit Sub
            End If
            
            If UserIndex <> tempChr Then
                Call UserAtacadoPorUsuario(UserIndex, tempChr)
            Else
                Exit Sub
            End If
        
            Call InfoHechizo(UserIndex)
        
            Danio = RandomNumber(Hechizos(SpellIndex).MinHam, Hechizos(SpellIndex).MaxHam)
        
            .Stats.MinHam = .Stats.MinHam - Danio
            
            If UserIndex <> tempChr Then
                Call WriteConsoleMsg(UserIndex, "Le sacaste " & Danio & " puntos de Hambre a " & .name, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " daño en " & Danio & " puntos de Hambre.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, "Te sacaste " & Danio & " puntos de Hambre.", FontTypeNames.FONTTYPE_FIGHT)
            End If
        
            HechizoCasteado = True
        
            If .Stats.MinHam < 1 Then
                .Stats.MinHam = 0
            End If
        
            Call WriteUpdateHungerAndThirst(tempChr)
        End If
    
        'Sed
        If Hechizos(SpellIndex).SubeSed = 1 Then
        
            Call InfoHechizo(UserIndex)
        
            Danio = RandomNumber(Hechizos(SpellIndex).MinSed, Hechizos(SpellIndex).MaxSed)
        
            .Stats.MinSed = .Stats.MinSed + Danio
            
            If .Stats.MinSed > 100 Then
                .Stats.MinSed = 100
            End If
            
            Call WriteUpdateHungerAndThirst(tempChr)
             
            If UserIndex <> tempChr Then
                Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Danio & " puntos de Sed a " & .name, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " restauró " & Danio & " puntos de Sed.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, "Te restauraste " & Danio & " puntos de Sed.", FontTypeNames.FONTTYPE_FIGHT)
            End If
        
            HechizoCasteado = True
        
        ElseIf Hechizos(SpellIndex).SubeSed = 2 Then
        
            If Not PuedeAtacar(UserIndex, tempChr) Then
                Exit Sub
            End If
        
            If UserIndex <> tempChr Then
                Call UserAtacadoPorUsuario(UserIndex, tempChr)
            End If
        
            Call InfoHechizo(UserIndex)
        
            Danio = RandomNumber(Hechizos(SpellIndex).MinSed, Hechizos(SpellIndex).MaxSed)
        
            .Stats.MinSed = .Stats.MinSed - Danio
        
            If UserIndex <> tempChr Then
                Call WriteConsoleMsg(UserIndex, "Le sacaste " & Danio & " puntos de Sed a " & .name, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te sacó " & Danio & " puntos de Sed.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, "Te sacaste " & Danio & " puntos de Sed.", FontTypeNames.FONTTYPE_FIGHT)
            End If
        
            If .Stats.MinSed < 1 Then
                .Stats.MinSed = 0
            End If
        
            Call WriteUpdateHungerAndThirst(tempChr)
        
            HechizoCasteado = True
        End If
    
    '<-------- Agilidad ---------->
        If Hechizos(SpellIndex).SubeAgilidad = 1 Then
        
            If .Stats.Atributos(eAtributos.Agilidad) < MinimoInt(MaxAtributos, .Stats.AtributosBackUP(Agilidad) * 2) Then
                
                Danio = RandomNumber(Hechizos(SpellIndex).MinAgilidad, Hechizos(SpellIndex).MaxAgilidad)
            

                .Stats.Atributos(eAtributos.Agilidad) = .Stats.Atributos(eAtributos.Agilidad) + Danio
                
                If .Stats.Atributos(eAtributos.Agilidad) > MinimoInt(MaxAtributos, .Stats.AtributosBackUP(Agilidad) * 2) Then
                    .Stats.Atributos(eAtributos.Agilidad) = MinimoInt(MaxAtributos, .Stats.AtributosBackUP(Agilidad) * 2)
                End If
                
                Call WriteUpdateDexterity(tempChr)
            End If
            
            Call InfoHechizo(UserIndex)
            
            .flags.DuracionEfecto = 2000
            .flags.TomoPocion = True
            HechizoCasteado = True
            
        ElseIf Hechizos(SpellIndex).SubeAgilidad = 2 Then
        
            If Not PuedeAtacar(UserIndex, tempChr) Then
                Exit Sub
            End If
            
            If .Stats.Atributos(eAtributos.Agilidad) > MINATRIBUTOS Then
                Danio = RandomNumber(Hechizos(SpellIndex).MinAgilidad, Hechizos(SpellIndex).MaxAgilidad)
        
                .Stats.Atributos(eAtributos.Agilidad) = .Stats.Atributos(eAtributos.Agilidad) - Danio
                
                If .Stats.Atributos(eAtributos.Agilidad) < MINATRIBUTOS Then
                    .Stats.Atributos(eAtributos.Agilidad) = MINATRIBUTOS
                End If

                Call WriteUpdateDexterity(tempChr)
            End If
            
            Call UserAtacadoPorUsuario(UserIndex, tempChr)
            Call InfoHechizo(UserIndex)
            
            .flags.DuracionEfecto = 1000
            .flags.TomoPocion = True
            HechizoCasteado = True
            
        End If
    
        '<-------- Fuerza ---------->
        If Hechizos(SpellIndex).SubeFuerza = 1 Then
    
            If .Stats.Atributos(eAtributos.Fuerza) < MinimoInt(MaxAtributos, .Stats.AtributosBackUP(Fuerza) * 2) Then
                
                Danio = RandomNumber(Hechizos(SpellIndex).MinFuerza, Hechizos(SpellIndex).MaxFuerza)

                .Stats.Atributos(eAtributos.Fuerza) = .Stats.Atributos(eAtributos.Fuerza) + Danio
                
                If .Stats.Atributos(eAtributos.Fuerza) > MinimoInt(MaxAtributos, .Stats.AtributosBackUP(Fuerza) * 2) Then
                    .Stats.Atributos(eAtributos.Fuerza) = MinimoInt(MaxAtributos, .Stats.AtributosBackUP(Fuerza) * 2)
                End If
            
                Call WriteUpdateStrenght(tempChr)
            End If
            
            Call InfoHechizo(UserIndex)
            
            .flags.TomoPocion = True
            .flags.DuracionEfecto = 2000
            HechizoCasteado = True
            
        ElseIf Hechizos(SpellIndex).SubeFuerza = 2 Then
    
            If Not PuedeAtacar(UserIndex, tempChr) Then
                Exit Sub
            End If
                
            If .Stats.Atributos(eAtributos.Fuerza) > MINATRIBUTOS Then
                Danio = RandomNumber(Hechizos(SpellIndex).MinFuerza, Hechizos(SpellIndex).MaxFuerza)
                        
                .Stats.Atributos(eAtributos.Fuerza) = .Stats.Atributos(eAtributos.Fuerza) - Danio
            
                If .Stats.Atributos(eAtributos.Fuerza) < MINATRIBUTOS Then
                    .Stats.Atributos(eAtributos.Fuerza) = MINATRIBUTOS
                End If

                Call WriteUpdateStrenght(tempChr)
            End If
            
            Call UserAtacadoPorUsuario(UserIndex, tempChr)
            Call InfoHechizo(UserIndex)
            
            .flags.DuracionEfecto = 1000
            .flags.TomoPocion = True
            HechizoCasteado = True
        
        End If
        
        'Salud
        If Hechizos(SpellIndex).SubeHP = 1 Then
        
            'Verifica que el usuario no este muerto
            If UserList(tempChr).Stats.Muerto Then
                Exit Sub
            End If
        
            If .Stats.MinHP < .Stats.MaxHP Then
            
                Danio = RandomNumber(Hechizos(SpellIndex).MinHP, Hechizos(SpellIndex).MaxHP)
                Danio = Danio + Porcentaje(Danio, 3 * UserList(UserIndex).Stats.Elv)
            
                Call InfoHechizo(UserIndex)
        
                .Stats.MinHP = .Stats.MinHP + Danio
                
                If .Stats.MinHP > .Stats.MaxHP Then
                    .Stats.MinHP = .Stats.MaxHP
                End If
        
                If UserIndex <> tempChr Then
                    Call WriteUserDamaged(tempChr, UserList(UserIndex).Char.CharIndex, Danio, 2)
                End If
                
                Call WriteDamage(UserIndex, UserList(tempChr).Char.CharIndex, Danio, UserList(tempChr).Stats.MinHP, UserList(tempChr).Stats.MaxHP, 2)
            
                HechizoCasteado = True
            End If
            
        ElseIf Hechizos(SpellIndex).SubeHP = 2 Then
        
            If UserIndex = tempChr Then
                Exit Sub
            End If
        
            Danio = RandomNumber(Hechizos(SpellIndex).MinHP, Hechizos(SpellIndex).MaxHP)
        
            Danio = Danio + Porcentaje(Danio, 3 * UserList(UserIndex).Stats.Elv)
        
            If Hechizos(SpellIndex).StaffAffected Then
                If UserList(UserIndex).Clase = eClass.Mage Then
                    If UserList(UserIndex).Inv.RightHand > 0 Then
                        Danio = (Danio * (ObjData(UserList(UserIndex).Inv.RightHand).StaffDamageBonus + 70)) \ 100
                    Else
                        Danio = Danio * 0.7 'Baja Danio a 70% del original
                    End If
                End If
            End If
        
            If UserList(UserIndex).Inv.Ring = LAUDMAGICO Or UserList(UserIndex).Inv.Ring = FLAUTAMAGICA Then
                Danio = Danio * 1.04  'laud magico de los bardos
            End If
        
            'cascos antimagia
            If .Inv.Head > 0 Then
                Danio = Danio - RandomNumber(ObjData(.Inv.Head).MinDefM, ObjData(.Inv.Head).MaxDefM)
            End If
        
            'anillos
            If .Inv.Ring > 0 Then
                Danio = Danio - RandomNumber(ObjData(.Inv.Ring).MinDefM, ObjData(.Inv.Ring).MaxDefM)
            End If
        
            If Danio < 0 Then
                Danio = 0
            End If
        
            If Not PuedeAtacar(UserIndex, tempChr) Then
                Exit Sub
            End If
        
            If UserIndex <> tempChr Then
                Call UserAtacadoPorUsuario(UserIndex, tempChr)
            End If
          
            Call InfoHechizo(UserIndex)
       
            .Stats.MinHP = .Stats.MinHP - Danio
    
            Call WriteDamage(UserIndex, .Char.CharIndex, Danio, .Stats.MinHP, .Stats.MaxHP, 1)
            Call WriteUserDamaged(tempChr, UserList(UserIndex).Char.CharIndex, Danio, 1)
            
            'Muere
            If .Stats.MinHP < 1 Then
                Call UserDie(tempChr, UserIndex)
            End If
        
            HechizoCasteado = True
        End If
    
        'Mana
        If Hechizos(SpellIndex).SubeMana = 1 Then
        
            Call InfoHechizo(UserIndex)
            .Stats.MinMan = .Stats.MinMan + Danio
        
            Call WriteUpdateMana(tempChr)
        
            If UserIndex <> tempChr Then
                Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Danio & " puntos de maná a " & .name, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te restauró " & Danio & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, "Te restauraste " & Danio & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT)
            End If
        
            HechizoCasteado = True
        
        ElseIf Hechizos(SpellIndex).SubeMana = 2 Then
            If Not PuedeAtacar(UserIndex, tempChr) Then
                Exit Sub
            End If
        
            If UserIndex <> tempChr Then
                Call UserAtacadoPorUsuario(UserIndex, tempChr)
            End If
        
            Call InfoHechizo(UserIndex)
        
            If UserIndex <> tempChr Then
                Call WriteConsoleMsg(UserIndex, "Le sacaste " & Danio & " puntos de maná a " & .name, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te sacó " & Danio & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, "Te sacaste " & Danio & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT)
            End If
        
            .Stats.MinMan = .Stats.MinMan - Danio
            Call WriteUpdateMana(tempChr)
        
            HechizoCasteado = True
        End If
    
        'Stamina
        If Hechizos(SpellIndex).SubeSta = 1 Then
            Call InfoHechizo(UserIndex)
            .Stats.MinSta = .Stats.MinSta + Danio
        
            Call WriteUpdateSta(tempChr)
        
            If UserIndex <> tempChr Then
                Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Danio & " puntos de vitalidad a " & .name, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te restauró " & Danio & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, "Te restauraste " & Danio & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
            End If
            
            HechizoCasteado = True
            
        ElseIf Hechizos(SpellIndex).SubeSta = 2 Then
            If Not PuedeAtacar(UserIndex, tempChr) Then
                Exit Sub
            End If
        
            If UserIndex <> tempChr Then
                Call UserAtacadoPorUsuario(UserIndex, tempChr)
            End If
        
            Call InfoHechizo(UserIndex)
        
            If UserIndex <> tempChr Then
                Call WriteConsoleMsg(UserIndex, "Le sacaste " & Danio & " puntos de vitalidad a " & .name, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te sacó " & Danio & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, "Te sacaste " & Danio & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
            End If
        
            .Stats.MinSta = .Stats.MinSta - Danio
            
            Call WriteUpdateSta(tempChr)
        
            HechizoCasteado = True
        End If
        
    End With
    
    Call FlushBuffer(tempChr)

End Sub
