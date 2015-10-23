Attribute VB_Name = "SistemaCombate"
Option Explicit

Public Const MaxDISTANCIAARCO As Byte = 18
Public Const MaxDISTANCIAMAGIA As Byte = 18

Public Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        MinimoInt = b
    Else
        MinimoInt = a
    End If
End Function

Public Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        MaximoInt = a
    Else
        MaximoInt = b
    End If
End Function

Private Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long
    PoderEvasionEscudo = (UserList(UserIndex).Skills.Skill(eSkill.Defensa).Elv * ModClase(UserList(UserIndex).Clase).Evasion) * 0.5
End Function

Private Function PoderEvasion(ByVal UserIndex As Integer) As Long
    Dim lTemp As Long
    With UserList(UserIndex)
        lTemp = (.Skills.Skill(eSkill.Tacticas).Elv + _
          .Skills.Skill(eSkill.Tacticas).Elv \ 33 * .Stats.Atributos(eAtributos.Agilidad)) * ModClase(.Clase).Evasion
       
        PoderEvasion = (lTemp + (2.5 * MaximoInt(.Stats.Elv - 12, 0)))
    End With
End Function

Private Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        If .Skills.Skill(eSkill.Armas).Elv < 31 Then
            PoderAtaqueTemp = .Skills.Skill(eSkill.Armas).Elv * ModClase(.Clase).AtaqueArmas
        ElseIf .Skills.Skill(eSkill.Armas).Elv < 61 Then
            PoderAtaqueTemp = (.Skills.Skill(eSkill.Armas).Elv + .Stats.Atributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        ElseIf .Skills.Skill(eSkill.Armas).Elv < 91 Then
            PoderAtaqueTemp = (.Skills.Skill(eSkill.Armas).Elv + 2 * .Stats.Atributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        Else
           PoderAtaqueTemp = (.Skills.Skill(eSkill.Armas).Elv + 3 * .Stats.Atributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        End If
        
        PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.Elv - 12, 0)))
    End With
End Function

Private Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        If .Skills.Skill(eSkill.Proyectiles).Elv < 31 Then
            PoderAtaqueTemp = .Skills.Skill(eSkill.Proyectiles).Elv * ModClase(.Clase).AtaqueProyectiles
        ElseIf .Skills.Skill(eSkill.Proyectiles).Elv < 61 Then
            PoderAtaqueTemp = (.Skills.Skill(eSkill.Proyectiles).Elv + .Stats.Atributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueProyectiles
        ElseIf .Skills.Skill(eSkill.Proyectiles).Elv < 91 Then
            PoderAtaqueTemp = (.Skills.Skill(eSkill.Proyectiles).Elv + 2 * .Stats.Atributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueProyectiles
        Else
            PoderAtaqueTemp = (.Skills.Skill(eSkill.Proyectiles).Elv + 3 * .Stats.Atributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueProyectiles
        End If
        
        PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.Elv - 12, 0)))
    End With
End Function

Private Function PoderAtaqueWrestling(ByVal UserIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        If .Skills.Skill(eSkill.Wrestling).Elv < 31 Then
            PoderAtaqueTemp = .Skills.Skill(eSkill.Wrestling).Elv * ModClase(.Clase).AtaqueArmas
        ElseIf .Skills.Skill(eSkill.Wrestling).Elv < 61 Then
            PoderAtaqueTemp = (.Skills.Skill(eSkill.Wrestling).Elv + .Stats.Atributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        ElseIf .Skills.Skill(eSkill.Wrestling).Elv < 91 Then
            PoderAtaqueTemp = (.Skills.Skill(eSkill.Wrestling).Elv + 2 * .Stats.Atributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        Else
            PoderAtaqueTemp = (.Skills.Skill(eSkill.Wrestling).Elv + 3 * .Stats.Atributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        End If
        
        PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.Elv - 12, 0)))
    End With
End Function

Public Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
    Dim PoderAtaque As Long
    Dim Arma As Integer
    Dim Skill As eSkill
    Dim ProbExito As Long
    
    If UsaArco(UserIndex) > 0 Then
        PoderAtaque = PoderAtaqueProyectil(UserIndex)
        Skill = eSkill.Proyectiles
    
    ElseIf UsaArmaNoArco(UserIndex) > 0 Then
        PoderAtaque = PoderAtaqueArma(UserIndex)
        Skill = eSkill.Armas
        
    Else 'Peleando con puños
        PoderAtaque = PoderAtaqueWrestling(UserIndex)
        Skill = eSkill.Wrestling
    End If
    
    'Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - NpcList(NpcIndex).PoderEvasion) * 0.4)))
    
    'Temporal
        If UserList(UserIndex).Stats.Elv < 15 Then
        ProbExito = ProbExito * 1.25
    End If
    
    UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
    
    If UserImpactoNpc Then
        Call SubirSkill(UserIndex, Skill, True)
    Else
        Call SubirSkill(UserIndex, Skill, False)
    End If
End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
'*************************************************
'Author: Unknown
'Last modified: 03\15\2006
'Revisa si un Npc logra impactar a un user o no
'03\15\2006 Maraxus - Evité una división por cero que eliminaba Npcs
'*************************************************
    Dim Rechazo As Boolean
    Dim ProbRechazo As Long
    Dim ProbExito As Long
    Dim UserEvasion As Long
    Dim NpcPoderAtaque As Long
    Dim EvasionEscudo As Long
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long
    
    UserEvasion = PoderEvasion(UserIndex)
    NpcPoderAtaque = NpcList(NpcIndex).PoderAtaque
    EvasionEscudo = PoderEvasionEscudo(UserIndex)
    
    SkillTacticas = UserList(UserIndex).Skills.Skill(eSkill.Tacticas).Elv
    SkillDefensa = UserList(UserIndex).Skills.Skill(eSkill.Defensa).Elv

    'Está usando un escudo?
    If UsaEscudo(UserIndex) > 0 Then
        UserEvasion = UserEvasion + EvasionEscudo
    End If
    
    'Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))
    
    NpcImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    'Está usando un escudo?
    If UsaEscudo(UserIndex) > 0 Then
        If Not UserList(UserIndex).flags.Meditando Then
            If Not NpcImpacto Then
                If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
                    'Chances are rounded
                    ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa \ (SkillDefensa + SkillTacticas)))
                Else
                    ProbRechazo = 10 'Si no tiene skills le dejamos el 10% mínimo
                End If
                
                Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                
                If Rechazo Then
                    'Se rechazo el ataque con el escudo
                    Call SendData(SendTarget.ToPCArea, UserIndex, Msg_BlockedWithShield(UserList(UserIndex).Char.CharIndex))
                    Call SubirSkill(UserIndex, eSkill.Defensa, True)
                Else
                    Call SubirSkill(UserIndex, eSkill.Defensa, False)
                End If
            End If
        End If
    End If
End Function

Public Function CalcularDanio(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Integer
    
    Dim DanioArma As Integer
    Dim DanioUsuario As Long
    Dim Arma As ObjData
    Dim ModifClase As Single
    Dim Proyectil As ObjData
    Dim DanioMaxArma As Integer
    Dim DanioMinArma As Integer
    Dim ObjIndex As Integer
    
    With UserList(UserIndex)
       
        'Ataca a un npc?
        If NpcIndex > 0 Then
        
            If UsaArco(UserIndex) > 0 Then
                Arma = ObjData(.Inv.LeftHand)
                ModifClase = ModClase(.Clase).DanioProyectiles
                DanioArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                DanioMaxArma = Arma.MaxHit
                
                If Arma.Municion Then
                    Proyectil = ObjData(.Inv.RightHand)
                    DanioArma = DanioArma + RandomNumber(Proyectil.MinHit, Proyectil.MaxHit)
                    'For some reason this isn't done...
                    'DanioMaxArma = DanioMaxArma + proyectil.MaxHit
                End If
                
            ElseIf UsaArmaNoArco(UserIndex) > 0 Then
                Arma = ObjData(.Inv.RightHand)
                ModifClase = ModClase(.Clase).DanioArmas
                
                If .Inv.RightHand = EspadaMataDragonesIndex Then 'Usa la mata Dragones?
                    If NpcList(NpcIndex).Type = DRAGON Then 'Ataca Dragon?
                        DanioArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                        DanioMaxArma = Arma.MaxHit
                    Else 'Sino es Dragon Danio es 1
                        DanioArma = 1
                        DanioMaxArma = 1
                    End If
                Else
                    DanioArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                    DanioMaxArma = Arma.MaxHit
                End If
            End If
        
        Else 'Ataca usuario
            If UsaArco(UserIndex) > 0 Then
                Arma = ObjData(.Inv.LeftHand)
                ModifClase = ModClase(.Clase).DanioProyectiles
                DanioArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                DanioMaxArma = Arma.MaxHit
                 
                If Arma.Municion Then
                    Proyectil = ObjData(.Inv.RightHand)
                    DanioArma = DanioArma + RandomNumber(Proyectil.MinHit, Proyectil.MaxHit)
                    'For some reason this isn't done...
                    'DanioMaxArma = DanioMaxArma + proyectil.MaxHit
                End If
            ElseIf UsaArmaNoArco(UserIndex) > 0 Then
                Arma = ObjData(.Inv.RightHand)
                ModifClase = ModClase(.Clase).DanioArmas
                
                If .Inv.RightHand = EspadaMataDragonesIndex Then
                    ModifClase = ModClase(.Clase).DanioArmas
                    DanioArma = 1 'Si usa la espada mataDragones Danio es 1
                    DanioMaxArma = 1
                Else
                    DanioArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                    DanioMaxArma = Arma.MaxHit
                End If
            End If
        End If
        
        If UsaArco(UserIndex) < 1 And UsaArmaNoArco(UserIndex) < 1 Then
            ModifClase = ModClase(.Clase).DanioWrestling

            'Danio sin guantes
            DanioMinArma = 5
            DanioMaxArma = 8
            
            'Plus de guantes (en slot de anillo)
            ObjIndex = .Inv.Ring
            
            If ObjIndex > 0 Then
                If ObjData(ObjIndex).Guante = 1 Then
                    DanioMinArma = DanioMinArma + ObjData(ObjIndex).MinHit
                    DanioMaxArma = DanioMaxArma + ObjData(ObjIndex).MaxHit
                End If
            End If
            
            DanioArma = RandomNumber(DanioMinArma, DanioMaxArma)
        End If
        
        DanioUsuario = RandomNumber(.Stats.MinHit, .Stats.MaxHit) + 20 + (.Stats.Atributos(eAtributos.Fuerza) * 1.5)
        
        CalcularDanio = 3 * DanioArma + (DanioMaxArma \ 5 + DanioUsuario) * ModifClase
    End With
End Function

Public Function Apuñalar(ByVal UserIndex As Integer, ByVal Danio As Integer, Optional ByVal VictimNpcIndex As Integer = 0, Optional ByVal VictimUserIndex As Integer = 0)
'Simplifique la cuenta que hacia para sacar la suerte
'y arregle la cuenta que hacia para sacar el Danio

    Dim Suerte As Integer
    Dim Skill As Integer
    
    Skill = UserList(UserIndex).Skills.Skill(eSkill.Apuñalar).Elv
    
    Select Case UserList(UserIndex).Clase
        Case eClass.Assasin
            Suerte = Int(((0.00003 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)
        
        Case eClass.Cleric, eClass.Paladin
            Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
        
        Case eClass.Bard
            Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
        
        Case Else
            Suerte = Int(0.0361 * Skill + 4.39)
    End Select
    
    If RandomNumber(0, 100) < Suerte Then
        If VictimUserIndex > 0 Then
            If UserList(UserIndex).Clase = eClass.Assasin Then
                Apuñalar = Danio * 0.7
            Else
                Apuñalar = Danio * 0.8
            End If
                    
            Call FlushBuffer(VictimUserIndex)
        Else
            Apuñalar = Danio
        End If
        
        Call SubirSkill(UserIndex, eSkill.Apuñalar, True)
        
        Call SendData(SendTarget.ToUserAreaButIndex, UserIndex, Msg_SoundFX(SND_APU, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    Else
        Call SubirSkill(UserIndex, eSkill.Apuñalar, False)
    End If

End Function

Public Function GolpeCritico(ByVal UserIndex As Integer, ByVal Danio As Integer, Optional ByVal VictimNpcIndex As Integer = 0, Optional ByVal VictimUserIndex As Integer = 0)

    If VictimNpcIndex > 0 Then
        If UserList(UserIndex).Clase = eClass.Bandit Then
            If (RandomNumber(0, 100)) > 85 Then
                GolpeCritico = Danio * 0.2
            End If
        ElseIf (RandomNumber(0, 100)) > 90 Then
            GolpeCritico = Danio * 0.2
        End If
    ElseIf UserList(UserIndex).Clase = eClass.Bandit Then
        If RandomNumber(0, 100) > 90 Then
            GolpeCritico = 0.15
        End If
    ElseIf RandomNumber(0, 100) > 95 Then
        GolpeCritico = Danio * 0.15
    End If
    
End Function

Public Sub UserDanioNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    
    Dim Danio As Integer
    Dim DanioAntes As Integer
    Dim TipoGolpe As Byte '0:Golpe común | 1:Golpe crítico | 2:Golpe apuñalado | 3:Golpe crítico y apuñalado :D
    
    Danio = CalcularDanio(UserIndex, NpcIndex)
        
    If Danio < 1 Then
        Exit Sub
    End If
    
    'Esta navegando? si es asi le sumamos el Danio del barco
    If UserList(UserIndex).flags.Navegando Then
        If UserList(UserIndex).Inv.Ship > 0 Then
            Danio = Danio + RandomNumber(ObjData(UserList(UserIndex).Inv.Ship).MinHit, ObjData(UserList(UserIndex).Inv.Ship).MaxHit)
        End If
    End If
         
    DanioAntes = Danio
        
    'Golpe crítico
    Danio = Danio + GolpeCritico(UserIndex, Danio, NpcIndex)
    
    If Danio <> DanioAntes Then
        TipoGolpe = 3
    End If
    
    DanioAntes = Danio
    
    'Trata de apuñalar por la espalda al enemigo
    If PuedeApuñalar(UserIndex) Then
        Danio = Danio + Apuñalar(UserIndex, Danio, NpcIndex)
    End If
                                    
    If Danio <> DanioAntes Then
        If TipoGolpe = 3 Then
            TipoGolpe = 5
        Else
            TipoGolpe = 4
        End If
    End If
    
    Call UserParalizaGolpe(UserIndex, NpcIndex)
    
    With NpcList(NpcIndex)
        If TipoGolpe < 4 Or UserList(UserIndex).Clase <> eClass.Assasin Then
            Danio = Danio - .Stats.Def
        End If
    
        If Danio < 0 Then
            Danio = 0
        End If
        
        .Stats.MinHP = .Stats.MinHP - Danio

        Call WriteDamage(UserIndex, .Char.CharIndex, Danio, .Stats.MinHP, .Stats.MaxHP, TipoGolpe)

        Call CalcularDarExp(UserIndex, NpcIndex, Danio)

        If .Stats.MinHP < 1 Then
            'Si era un Dragon perdemos la espada mataDragones
            If .Type = DRAGON Then
                If .Stats.MaxHP > 100000 Then
                    Call LogDesarrollo(UserList(UserIndex).Name & " mató un dragón")
                End If
            End If
            
            Call AllFollowAmo(UserIndex)
            
            Call MuereNpc(NpcIndex, UserIndex)
        End If
    End With
End Sub

Public Sub NpcDanio(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    Dim Danio As Integer
    Dim Lugar As Integer
    Dim Absorbido As Integer
    Dim DefBarco As Integer
    Dim Obj As ObjData
    
    Danio = RandomNumber(NpcList(NpcIndex).Stats.MinHit, NpcList(NpcIndex).Stats.MaxHit)
    
    With UserList(UserIndex)
        If .flags.Navegando And .Inv.Ship > 0 Then
            Obj = ObjData(.Inv.Ship)
            DefBarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                'Si tiene casco absorbe el golpe
                If .Inv.Head > 0 Then
                   Obj = ObjData(.Inv.Head)
                   Absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                End If
          Case Else
                'Si tiene armadura absorbe el golpe
                If .Inv.Body > 0 Then
                    Dim Obj2 As ObjData
                    Obj = ObjData(.Inv.Body)
                    
                    If UsaEscudo(UserIndex) > 0 Then
                        Obj2 = ObjData(.Inv.LeftHand)
                        Absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
                    Else
                        Absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    End If
                End If
        End Select
        
        Absorbido = Absorbido + DefBarco
        Danio = Danio - Absorbido
        
        If Danio < 1 Then
            Danio = 1
        End If
        
        If .flags.Privilegios And PlayerType.User Then
            .Stats.MinHP = .Stats.MinHP - Danio
        End If
        
        Call WriteUserDamaged(UserIndex, NpcList(NpcIndex).Char.CharIndex, Danio, 1)
        
        If .flags.Meditando Then
            If Danio > Fix(.Stats.MinHP / 100 * .Stats.Atributos(eAtributos.Inteligencia) * .Skills.Skill(eSkill.Meditar).Elv / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
                .flags.Meditando = False
                .Char.FX = 0
                Call SendData(SendTarget.ToUserAreaButIndex, UserIndex, Msg_CreateCharFX(.Char.CharIndex))
            End If
        End If
        
        'Muere el usuario
        If .Stats.MinHP < 1 Then
            'Call WriteNpcKillUser(UserIndex) 'Le informamos que ha muerto ;)
            
            If NpcList(NpcIndex).MaestroUser > 0 Then
                Call AllFollowAmo(NpcList(NpcIndex).MaestroUser)
            Else
                'Al matarlo no lo sigue mas
                If NpcList(NpcIndex).Stats.Alineacion = 0 Then
                    NpcList(NpcIndex).TargetUser = 0
                End If
            End If
            Call UserDie(UserIndex, , NpcIndex)
        End If
    End With
End Sub

Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    Dim j As Integer
    
    If UserList(UserIndex).Pets.NroALaVez < 1 Then
        Exit Sub
    End If
    
    If Not PuedeAtacarNpc(UserIndex, NpcIndex) Then
        Exit Sub
    End If
    
    With UserList(UserIndex)
        For j = 1 To MaxPets
            If .Pets.Pet(j).index > 0 Then
                If .Pets.Pet(j).index <> NpcIndex Then
                    If NpcList(NpcIndex).MaestroUser <> UserIndex Then
                        If NpcList(.Pets.Pet(j).index).TargetNpc < 1 Then
                            NpcList(.Pets.Pet(j).index).TargetNpc = NpcIndex
                            NpcList(.Pets.Pet(j).index).TargetUser = 0
                        End If
                    End If
                End If
            End If
        Next j
    End With
End Sub

Public Sub AllFollowAmo(ByVal UserIndex As Integer)
    Dim j As Integer
    
    For j = 1 To MaxPets
        If UserList(UserIndex).Pets.Pet(j).index > 0 Then
            Call FollowAmo(UserList(UserIndex).Pets.Pet(j).index)
        End If
    Next j
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
    
    With UserList(UserIndex)
        If .flags.AdminInvisible > 0 Then
            Exit Function
        End If
    End With
    
    With NpcList(NpcIndex)
        If .CanAttack = 1 Then
            NpcAtacaUser = True
            Call CheckPets(NpcIndex, UserIndex)
            
            If UserList(UserIndex).flags.AtacadoPorNpc < 1 And UserList(UserIndex).flags.AtacadoPorUser < 1 Then
                UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex
            End If
        Else
            NpcAtacaUser = False
            Exit Function
        End If
        
        .CanAttack = 0
        
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNpcArea, NpcIndex, Msg_SoundFX(.flags.Snd1, .Pos.X, .Pos.Y))
        Else
            Call SendData(SendTarget.ToNpcArea, NpcIndex, Msg_SoundFX(SND_NpcSWING, .Pos.X, .Pos.Y))
        End If
    End With
    
    If NpcImpacto(NpcIndex, UserIndex) Then
        With UserList(UserIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_IMPACTO, .Pos.X, .Pos.Y))
            
            If Not .flags.Meditando Then
                If Not .flags.Navegando Then
                    Call SendData(SendTarget.ToUserAreaButIndex, UserIndex, Msg_CreateFX(.Pos.X, .Pos.Y, FXSANGRE))
                End If
            End If
            
            Call NpcDanio(NpcIndex, UserIndex)
            
            If UserList(UserIndex).Stats.MinHP > 0 Then
                '¿Puede envenenar?
                If NpcList(NpcIndex).Veneno = 1 Then
                    If UserList(UserIndex).flags.Envenenado < 1 Then
                        Call NpcEnvenenarUser(UserIndex)
                    End If
                End If
            End If
        End With
        
        Call SubirSkill(UserIndex, eSkill.Tacticas, False)
    Else
        Call WriteCharSwing(UserIndex, NpcList(NpcIndex).Char.CharIndex)
        Call SubirSkill(UserIndex, eSkill.Tacticas, True)
    End If
End Function

Private Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
    Dim PoderAtt As Long
    Dim PoderEva As Long
    Dim ProbExito As Long
    
    PoderAtt = NpcList(Atacante).PoderAtaque
    PoderEva = NpcList(Victima).PoderEvasion
    
    'Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtt - PoderEva) * 0.4))
    NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
End Function

Public Sub NpcDanioNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
    Dim Danio As Integer
    
    With NpcList(Atacante)
        Danio = RandomNumber(.Stats.MinHit, .Stats.MaxHit)
        NpcList(Victima).Stats.MinHP = NpcList(Victima).Stats.MinHP - Danio
        
        If .MaestroUser > 0 Then
            Call CalcularDarExp(.MaestroUser, Victima, Danio)
            Call EnviarDatosASlot(.MaestroUser, Msg_ChatOverHead(Danio, NpcList(Atacante).Char.CharIndex, RGB(200, 0, 0)))
        End If
        
        If NpcList(Victima).Stats.MinHP < 1 Then
            Call MuereNpc(Victima, .MaestroUser)
        End If
    End With
End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
    
    With NpcList(Atacante)

        'El npc puede atacar ???
        If .CanAttack = 1 Then
            .CanAttack = 0
        Else
            Exit Sub
        End If
        
        If NpcList(Victima).MaestroUser > 0 Then
            Call CheckPets(Atacante, NpcList(Victima).MaestroUser)
        End If
        
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNpcArea, Atacante, Msg_SoundFX(.flags.Snd1, .Pos.X, .Pos.Y))
        End If
        
        If NpcImpactoNpc(Atacante, Victima) Then
            If NpcList(Victima).flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNpcArea, Victima, Msg_SoundFX(NpcList(Victima).flags.Snd2, NpcList(Victima).Pos.X, NpcList(Victima).Pos.Y))
            Else
                Call SendData(SendTarget.ToNpcArea, Victima, Msg_SoundFX(SND_IMPACTO2, NpcList(Victima).Pos.X, NpcList(Victima).Pos.Y))
            End If
        
            If .MaestroUser > 0 Then
                Call SendData(SendTarget.ToNpcArea, Atacante, Msg_SoundFX(SND_IMPACTO, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToNpcArea, Victima, Msg_SoundFX(SND_IMPACTO, NpcList(Victima).Pos.X, NpcList(Victima).Pos.Y))
            End If
            
            Call NpcDanioNpc(Atacante, Victima)
        Else
            If .MaestroUser > 0 Then
                Call WriteCharSwing(.MaestroUser, .Char.CharIndex)
                Call SendData(SendTarget.ToNpcArea, Atacante, Msg_SoundFX(SND_NpcSWING, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToNpcArea, Victima, Msg_SoundFX(SND_NpcSWING, NpcList(Victima).Pos.X, NpcList(Victima).Pos.Y))
            End If
        End If
    End With
End Sub

Public Function UserAtacaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean

    If Not PuedeAtacarNpc(UserIndex, NpcIndex) Then
        Exit Function
    End If
    
    Call NpcAtacado(NpcIndex, UserIndex)
    
    Call CheckPets(NpcIndex, UserIndex)
    
    If UserImpactoNpc(UserIndex, NpcIndex) Then
        If NpcList(NpcIndex).flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNpcArea, NpcIndex, Msg_SoundFX(NpcList(NpcIndex).flags.Snd2, NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_IMPACTO2, NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.Y))
        End If
        
        Call UserDanioNpc(UserIndex, NpcIndex)
    Else
        If RandomNumber(1, 2) = 1 Then
            Call SendData(SendTarget.ToUserAreaButIndex, UserIndex, Msg_SoundFX(SND_SWING, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Else
            Call SendData(SendTarget.ToUserAreaButIndex, UserIndex, Msg_SoundFX(SND_SWING2, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        End If
        
        Call WriteUserSwing(UserIndex)
    End If
    
    'Reveló su condición de usuario al atacar, los npcs lo van a atacar
    UserList(UserIndex).flags.Ignorado = False
    
    UserAtacaNpc = True

End Function

Public Sub UserAtaca(ByVal UserIndex As Integer)
    Dim index As Integer
    Dim AttackPos As WorldPos
    Dim QuitaSt As Byte
    
    'Check bow's interval
    If Not IntervaloPermiteUsarArcos(UserIndex, False) Then
        Exit Sub
    End If
    
    'Check Spell-Magic interval
    If Not IntervaloPermItemagiaGolpe(UserIndex) Then
        'Check Attack interval
        If Not IntervaloPermiteAtacar(UserIndex) Then
            Exit Sub
        End If
    End If
    
    With UserList(UserIndex)
    
        QuitaSt = .Stats.MinSta * 0.05
    
        'Quitamos stamina
        If .Stats.MinSta >= QuitaSt Then
            Call QuitarSta(UserIndex, QuitaSt)
        Else
            Call WriteConsoleMsg(UserIndex, "No tenés suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'ENVÍA ANIMACIÓN ARMA
        Call SendData(SendTarget.ToPCArea, UserIndex, Msg_AnimAttack(UserList(UserIndex).Char.CharIndex))

        AttackPos = .Pos
        Call HeadtoPos(.Char.Heading, AttackPos)
        
        'Exit if not legal
        If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
            If RandomNumber(1, 2) = 1 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_SWING, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_SWING2, .Pos.X, .Pos.Y))
            End If
            Exit Sub
        End If
        
        If MapData(AttackPos.X, AttackPos.Y).UserIndex > 0 Then
            index = MapData(AttackPos.X, AttackPos.Y).UserIndex
            Call UserAtacaUser(UserIndex, index)
            
            Exit Sub
            
        ElseIf MapData(AttackPos.X, AttackPos.Y).NpcIndex > 0 Then
            index = MapData(AttackPos.X, AttackPos.Y).NpcIndex
            Call UserAtacaNpc(UserIndex, index)
  
            Exit Sub
        End If
        
        If RandomNumber(1, 2) = 1 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_SWING, .Pos.X, .Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_SWING2, .Pos.X, .Pos.Y))
        End If
        
        If .Counters.Trabajando Then
            .Counters.Trabajando = .Counters.Trabajando - 1
        End If
            
        If .Counters.Ocultando Then
            .Counters.Ocultando = .Counters.Ocultando - 1
        End If
    End With
End Sub

Public Function UserImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
    Dim ProbRechazo As Long
    Dim Rechazo As Boolean
    Dim ProbExito As Long
    Dim PoderAtaque As Long
    Dim UserPoderEvasion As Long
    Dim UserPoderEvasionEscudo As Long
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long
    Dim ProbEvadir As Long
    Dim Skill As eSkill
    
    SkillTacticas = UserList(VictimaIndex).Skills.Skill(eSkill.Tacticas).Elv
    SkillDefensa = UserList(VictimaIndex).Skills.Skill(eSkill.Defensa).Elv
    
    'Calculamos el poder de evasion...
    UserPoderEvasion = PoderEvasion(VictimaIndex)
    
    If UsaEscudo(VictimaIndex) > 0 Then
       UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
       UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
    Else
        UserPoderEvasionEscudo = 0
    End If
    
    'Está usando un arma?
    If UsaArco(VictimaIndex) > 0 Then
        PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
        Skill = eSkill.Proyectiles
    
    ElseIf UsaArmaNoArco(VictimaIndex) > 0 Then
        PoderAtaque = PoderAtaqueArma(AtacanteIndex)
        Skill = eSkill.Armas
    
    Else
        PoderAtaque = PoderAtaqueWrestling(AtacanteIndex)
        Skill = eSkill.Wrestling
    End If
    
    'Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtaque - UserPoderEvasion) * 0.4))
    
    'Se reduce la evasion un 25%
    If UserList(VictimaIndex).flags.Meditando Then
        ProbEvadir = (100 - ProbExito) * 0.75
        ProbExito = MinimoInt(90, 100 - ProbEvadir)
    End If
    
    UserImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    'Está usando un escudo?
    If UsaEscudo(VictimaIndex) > 0 Then
        If Not UserList(VictimaIndex).flags.Meditando Then
            'Fallo ???
            If Not UserImpacto Then
                'Chances are rounded
                ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa \ (SkillDefensa + SkillTacticas)))
                Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                If Rechazo Then
                    'Se rechazo el ataque con el escudo
                    Call WriteBlockedWithShieldOther(AtacanteIndex)
                    Call SendData(SendTarget.ToPCArea, VictimaIndex, Msg_BlockedWithShield(UserList(VictimaIndex).Char.CharIndex))
                    Call SubirSkill(VictimaIndex, eSkill.Defensa, True)
                Else
                    Call SubirSkill(VictimaIndex, eSkill.Defensa, False)
                End If
            End If
        End If
    End If
    
    If Not UserImpacto Then
        Call SubirSkill(AtacanteIndex, Skill, False)
    End If
    
    Call FlushBuffer(VictimaIndex)
End Function

Public Function UserAtacaUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean

On Error GoTo errhandler

    If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then
        Exit Function
    End If
    
    With UserList(AtacanteIndex)
        If Distancia(.Pos, UserList(VictimaIndex).Pos) > MaxDISTANCIAARCO Then
           Exit Function
        End If
        
        Call UserAtacadoPorUsuario(AtacanteIndex, VictimaIndex)
        
        If UserImpacto(AtacanteIndex, VictimaIndex) Then
            Call SendData(SendTarget.ToPCArea, AtacanteIndex, Msg_SoundFX(SND_IMPACTO, .Pos.X, .Pos.Y))
            
            If Not UserList(VictimaIndex).flags.Navegando Then
                Call SendData(SendTarget.ToPCArea, VictimaIndex, Msg_CreateFX(UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y, FXSANGRE))
            End If
            
            'Guantes de Hurto del Bandido en acción
            If .Clase = eClass.Bandit Then
                Call DoHurtar(AtacanteIndex, VictimaIndex)
            'y ahora, el ladrón puede llegar a paralizar con el golpe.
            ElseIf .Clase = eClass.Thief Then
                Call DoHandInmo(AtacanteIndex, VictimaIndex)
            End If
            
            Call SubirSkill(VictimaIndex, eSkill.Tacticas, False)
            Call UserDanioUser(AtacanteIndex, VictimaIndex)
        Else
            'Invisible admins doesn't make sound to other clients except itself
            If .flags.AdminInvisible < 1 Then
                If RandomNumber(1, 2) = 1 Then
                    Call SendData(SendTarget.ToUserAreaButIndex, AtacanteIndex, Msg_SoundFX(SND_SWING, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToUserAreaButIndex, AtacanteIndex, Msg_SoundFX(SND_SWING2, .Pos.X, .Pos.Y))
                End If
            End If
            
            Call WriteUserSwing(AtacanteIndex)
            Call WriteCharSwing(VictimaIndex, UserList(AtacanteIndex).Char.CharIndex)
            Call SubirSkill(VictimaIndex, eSkill.Tacticas, True)
        End If
        
        If .Clase = eClass.Thief Then
            Call Desarmar(AtacanteIndex, VictimaIndex)
        End If
    End With
    
    UserAtacaUser = True
    
    Exit Function
    
errhandler:
    Call LogError("Error en UserAtacaUser. Error " & Err.Number & ": " & Err.description)
End Function

Public Sub UserDanioUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    Dim Danio As Integer
    Dim DanioAntes
    Dim TipoGolpe As Byte '0:Golpe común | 1:Golpe crítico | 2:Golpe apuñalado | 3:Golpe crítico y apuñalado :D
    
    Dim Lugar As Byte
    Dim Defensa As Integer
    Dim Obj As ObjData
    Dim Refuerzo As Byte
           
    Danio = CalcularDanio(AtacanteIndex)
       
    If Danio < 1 Then
        Exit Sub
    End If
    
    Call UserEnvenenaGolpe(AtacanteIndex, VictimaIndex)
    
    Call UserParalizaGolpe(AtacanteIndex, , VictimaIndex)
    
    With UserList(AtacanteIndex)
        
        If .flags.Navegando And .Inv.Ship > 0 Then
             Obj = ObjData(.Inv.Ship)
             Danio = Danio + RandomNumber(Obj.MinHit, Obj.MaxHit)
        End If
                  
        DanioAntes = Danio

        Danio = Danio + GolpeCritico(AtacanteIndex, Danio, , VictimaIndex)
        
        If Danio <> DanioAntes Then
            TipoGolpe = 3
        End If
    
        DanioAntes = Danio
        
        'Trata de apuñalar por la espalda al enemigo
        If PuedeApuñalar(AtacanteIndex) Then
            Danio = Danio + Apuñalar(AtacanteIndex, Danio, , VictimaIndex)
        End If

        If Danio <> DanioAntes Then
            If TipoGolpe = 1 Then
                TipoGolpe = 4
            Else
                TipoGolpe = 5
            End If
        End If
        
        If UserList(VictimaIndex).flags.Navegando And UserList(VictimaIndex).Inv.Ship > 0 Then
             Obj = ObjData(UserList(VictimaIndex).Inv.Ship)
             Defensa = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        If UsaArco(AtacanteIndex) > 0 Then
            Refuerzo = ObjData(.Inv.LeftHand).Refuerzo
        ElseIf UsaArmaNoArco(AtacanteIndex) > 0 Then
            Refuerzo = ObjData(.Inv.RightHand).Refuerzo
        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                'Si tiene casco absorbe el golpe
                If UserList(VictimaIndex).Inv.Head > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Inv.Head)
                    Defensa = Defensa + RandomNumber(Obj.MinDef, Obj.MaxDef)
                End If

            Case Else
                'Si tiene armadura absorbe el golpe
                If UserList(VictimaIndex).Inv.Body > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Inv.Body)
                    Defensa = Defensa + RandomNumber(Obj.MinDef, Obj.MaxDef)
                End If
                
                'SI NO PEGA EN LAS PIERNAS, EL ESCUDO PROTEJE
                If Lugar <> PartesCuerpo.bPiernaIzquierda Then
                    If Lugar <> PartesCuerpo.bPiernaDerecha Then
                        If UsaEscudo(VictimaIndex) > 0 Then
                            Obj = ObjData(UserList(VictimaIndex).Inv.LeftHand)
                            Defensa = Defensa + RandomNumber(Obj.MinDef, Obj.MaxDef)
                        End If
                    End If
                End If
        End Select
             
        Defensa = Defensa - Refuerzo
        
        If Defensa < 0 Then
            Defensa = 0
        End If
        
        Danio = Danio - Defensa
        
        If Danio < 0 Then
            Danio = 1
        End If

        UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - Danio
        
        Call WriteDamage(AtacanteIndex, UserList(VictimaIndex).Char.CharIndex, Danio, UserList(VictimaIndex).Stats.MinHP, UserList(VictimaIndex).Stats.MaxHP, TipoGolpe)
        Call WriteUserDamaged(VictimaIndex, UserList(AtacanteIndex).Char.CharIndex, Danio, TipoGolpe)
        
        If .Stats.MinHam > 0 And .Stats.MinSed > 0 Then
            If UsaArco(AtacanteIndex) > 0 Then
                Call SubirSkill(AtacanteIndex, eSkill.Proyectiles, True)
            
            ElseIf UsaArmaNoArco(AtacanteIndex) > 0 Then
                Call SubirSkill(AtacanteIndex, eSkill.Armas, True)
            
            Else
                Call SubirSkill(AtacanteIndex, eSkill.Wrestling, True)
            End If
        End If
        
        If UserList(VictimaIndex).Stats.MinHP < 1 Then
            'Para que las Mascotas no sigan intentando luchar y
            'comiencen a seguir al amo
            Dim j As Integer
            For j = 1 To MaxPets
                If .Pets.Pet(j).index > 0 Then
                    If NpcList(.Pets.Pet(j).index).TargetUser = VictimaIndex Then
                        NpcList(.Pets.Pet(j).index).TargetUser = 0
                        Call FollowAmo(.Pets.Pet(j).index)
                    End If
                End If
            Next j

            Call UserDie(VictimaIndex, AtacanteIndex)
        End If
    End With
    
    Call FlushBuffer(VictimaIndex)
End Sub

Public Sub UserAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)

    If TriggerZonaPelea(AttackerIndex, VictimIndex) = TRIGGER6_PERMITE Then
        Exit Sub
    End If
    
    If UserList(VictimIndex).flags.Meditando Then
        UserList(VictimIndex).flags.Meditando = False
        UserList(VictimIndex).Char.FX = 0
        Call SendData(SendTarget.ToUserAreaButIndex, VictimIndex, Msg_CreateFX(UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y))
    End If
    
    Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)
    
    'Si la victima esta saliendo se cancela la salida
    Call CancelExit(VictimIndex)
    Call FlushBuffer(VictimIndex)
End Sub

Public Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
    'Reaccion de los Mascotas
    Dim iCount As Integer
    
    For iCount = 1 To MaxPets
        If UserList(Maestro).Pets.Pet(iCount).index > 0 Then
            NpcList(UserList(Maestro).Pets.Pet(iCount).index).TargetUser = victim
            NpcList(UserList(Maestro).Pets.Pet(iCount).index).TargetNpc = 0
            NpcList(UserList(Maestro).Pets.Pet(iCount).index).Movement = TipoAI.NpcDefensa
            NpcList(UserList(Maestro).Pets.Pet(iCount).index).Hostile = 1
        End If
    Next iCount
End Sub

Public Function PuedeAtacar(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
'Returns true if the AttackerIndex is allowed to attack the VictimIndex.

On Error GoTo errhandler

    'MUY importante el orden de estos "IF"...
    
    'Estás muerto no podes atacar
    If UserList(AttackerIndex).Stats.Muerto Then
        Exit Function
    End If
    
    'No podes atacar a alguien muerto
    If UserList(VictimIndex).Stats.Muerto Then
        Exit Function
    End If
    
    'No podés atacarte a vos mismo
    If AttackerIndex = VictimIndex Then
        Exit Function
    End If

    'Estamos en una Arena? o un trigger zona segura?
    Select Case TriggerZonaPelea(AttackerIndex, VictimIndex)
        Case eTrigger6.TRIGGER6_PERMITE
            PuedeAtacar = True
            Exit Function
        
        Case eTrigger6.TRIGGER6_PROHIBE
            Exit Function
        
        Case eTrigger6.TRIGGER6_AUSENTE
            'Si no estamos en el Trigger 6 entonces es imposible atacar un gm
            If (UserList(VictimIndex).flags.Privilegios And PlayerType.User) = 0 Then
                Exit Function
            End If
    End Select
    
    'Estas en un Mapa Seguro?
    If Not MapInfo(UserList(VictimIndex).Pos.Map).PK Then
        Call WriteConsoleMsg(AttackerIndex, "Esta es una zona segura, aqui no podes atacar otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
        Exit Function
    End If
    
    'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
    If MapData(UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).Trigger = eTrigger.ZONASEGURA Or _
        MapData(UserList(AttackerIndex).Pos.X, UserList(AttackerIndex).Pos.Y).Trigger = eTrigger.ZONASEGURA Then
        Call WriteConsoleMsg(AttackerIndex, "No podes pelear aqui.", FontTypeNames.FONTTYPE_WARNING)
        Exit Function
    End If
    
    PuedeAtacar = True
    
    Exit Function

errhandler:
    Call LogError("Error en PuedeAtacar. Error " & Err.Number & ": " & Err.description)
End Function

Public Function PuedeAtacarNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, Optional ByVal Paraliza As Boolean = False) As Boolean
'Returns True if UserIndex can attack the NpcIndex

    With NpcList(NpcIndex)
        'Estás muerto?
        If UserList(UserIndex).Stats.Muerto Then
            Exit Function
        End If
        
        'Sos consejero?
        If UserList(UserIndex).flags.Privilegios And PlayerType.Consejero Then
            Exit Function
        End If
        
        'Es una criatura atacable?
        If .Attackable = 0 Then
            Exit Function
        End If
        
        'Es valida la distancia a la cual estamos atacando?
        If Distancia(UserList(UserIndex).Pos, .Pos) >= MaxDISTANCIAARCO Then
           Exit Function
        End If
        
        If .MaestroUser > 0 Then
            If .MaestroUser = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No podés atacar a tu mascota.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Function
            End If
            
            If Not MapInfo(.Pos.Map).PK Then
                Exit Function
            End If
            
            Dim iCount As Byte
            For iCount = 1 To MaxPets
                If UserList(NpcList(NpcIndex).MaestroUser).Pets.Pet(iCount).index > 0 Then
                    With NpcList(UserList(NpcList(NpcIndex).MaestroUser).Pets.Pet(iCount).index)
                        .TargetUser = UserIndex
                        .TargetNpc = 0
                        .Movement = TipoAI.NpcDefensa
                        .Hostile = 1
                    End With
                End If
            Next iCount
            
        End If

    End With
    
    PuedeAtacarNpc = True
End Function

Private Function SameGuild(ByVal UserIndex As Integer, ByVal OtherUserIndex As Integer) As Boolean
    SameGuild = (UserList(UserIndex).Guild_Id = UserList(OtherUserIndex).Guild_Id) And _
                UserList(UserIndex).Guild_Id > 0
End Function

Private Function SameParty(ByVal UserIndex As Integer, ByVal OtherUserIndex As Integer) As Boolean
    SameParty = UserList(UserIndex).PartyIndex = UserList(OtherUserIndex).PartyIndex And _
                UserList(UserIndex).PartyIndex > 0
End Function

Public Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, Optional ByVal Danio As Integer = 0)

On Error Resume Next

    Dim ExpADar As Long
    
    If NpcList(NpcIndex).Stats.MaxHP < 1 Then
        Exit Sub
    End If
    
    ExpADar = 0.66 * CLng(Danio) * NpcList(NpcIndex).GiveEXP / CLng(NpcList(NpcIndex).Stats.MaxHP)
    
    If NpcList(NpcIndex).Stats.MinHP < 1 Then
        ExpADar = ExpADar + 0.34 * NpcList(NpcIndex).GiveEXP
    End If
    
    If ExpADar < 1 Then
        Exit Sub
    End If
    
    If ExpADar > 0 Then
        If UserList(UserIndex).PartyIndex > 0 Then
            Call mdParty.ObtenerExito(UserIndex, ExpADar, NpcList(NpcIndex).Pos.Map, NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.Y)
        Else
            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpADar
            Call WriteUpdateExp(UserIndex)
        End If
    End If
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6

On Error GoTo errhandler
    Dim tOrg As eTrigger
    Dim tDst As eTrigger
    
    tOrg = MapData(UserList(Origen).Pos.X, UserList(Origen).Pos.Y).Trigger
    tDst = MapData(UserList(Destino).Pos.X, UserList(Destino).Pos.Y).Trigger
    
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If

Exit Function
errhandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & Err.description)
End Function

Public Sub UserEnvenenaGolpe(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    
    Dim ObjInd As Integer

    If UsaArco(AtacanteIndex) > 0 Then
        ObjInd = UserList(AtacanteIndex).Inv.LeftHand
    
    ElseIf UsaArmaNoArco(AtacanteIndex) > 0 Then
        ObjInd = UserList(AtacanteIndex).Inv.RightHand
    End If
    
    If ObjInd > 0 Then
        If ObjData(ObjInd).Envenena = 1 Then
            
            If RandomNumber(1, 100) < 60 Then
                UserList(VictimaIndex).flags.Envenenado = 1
                Call WriteConsoleMsg(VictimaIndex, UserList(AtacanteIndex).Name & " te envenenó.", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(AtacanteIndex, "Has envenenado a " & UserList(VictimaIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(VictimaIndex)
            End If
        End If
    End If

End Sub

Public Sub UserParalizaGolpe(ByVal UserIndex As Integer, Optional ByVal VictimNpcIndex As Integer = 0, Optional ByVal VictimUserIndex As Integer = 0)

    Dim ObjInd As Integer
    
    If UsaArco(UserIndex) > 0 Then
        ObjInd = UserList(UserIndex).Inv.LeftHand
    
    ElseIf UsaArmaNoArco(UserIndex) > 0 Then
        ObjInd = UserList(UserIndex).Inv.RightHand
    End If
       
    If ObjInd > 0 Then
        If ObjData(ObjInd).Paraliza > 0 Then
            
            If VictimNpcIndex > 0 Then
                If RandomNumber(0, 100) < 20 Then
                    NpcList(VictimNpcIndex).flags.Paralizado = 1
                    NpcList(VictimNpcIndex).Contadores.Paralisis = IntervaloParalizado

                    Call SendData(SendTarget.ToNpcArea, VictimNpcIndex, Msg_SetParalized(NpcList(VictimNpcIndex).Char.CharIndex, 1))
                    Call SendData(SendTarget.ToNpcArea, VictimNpcIndex, Msg_CreateFX(NpcList(VictimNpcIndex).Pos.X, NpcList(VictimNpcIndex).Pos.Y, 8))
                End If
                
            ElseIf RandomNumber(0, 100) < 15 Then
                UserList(VictimUserIndex).flags.Paralizado = 1
                UserList(VictimUserIndex).Counters.Paralisis = IntervaloParalizado * 0.5

                Call SendData(SendTarget.ToPCArea, VictimUserIndex, Msg_SetParalized(UserList(VictimUserIndex).Char.CharIndex, 1))
                Call SendData(SendTarget.ToPCArea, VictimUserIndex, Msg_CreateFX(UserList(VictimUserIndex).Pos.X, UserList(VictimUserIndex).Pos.Y, 8))
                
                Call WriteConsoleMsg(VictimUserIndex, "¡El golpe te paralizó!", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(UserIndex, "¡Tu golpe lo paralizó!", FontTypeNames.FONTTYPE_FIGHT)
            End If
        End If
    End If

End Sub
