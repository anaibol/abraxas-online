Attribute VB_Name = "Usuarios"
Option Explicit

Public Sub RespawnearUsuario(ByVal UserIndex As Integer)
    
    Dim x As Byte
    Dim y As Byte
    
    If UserList(UserIndex).Hogar < 1 Then
        If EsPrincipiante(UserIndex) Then
             UserList(UserIndex).Hogar = Newbie.map
        Else
            UserList(UserIndex).Hogar = Ullathorpe.map
        End If
    End If
    
    Select Case UserList(UserIndex).Hogar
        
        Case Newbie.map
            x = Newbie.x
            y = Newbie.y
        
        Case Nix.map
            x = Nix.x
            y = Nix.y
            
        Case Ullathorpe.map
            x = Ullathorpe.x
            y = Ullathorpe.y
            
        Case Banderbill.map
            x = Banderbill.x
            y = Banderbill.y
            
        Case Lindos.map
            x = Lindos.x
            y = Lindos.y
            
        Case Arghal.map
            x = Arghal.x
            y = Arghal.y
            
        Case Else
            UserList(UserIndex).Hogar = Ullathorpe.map
            x = Ullathorpe.x
            y = Ullathorpe.y
    End Select
    
    If UserList(UserIndex).Stats.Muerto Then
        Call RevivirUsuario(UserIndex)
    End If
    
    Call WarpUserChar(UserIndex, UserList(UserIndex).Hogar, x, y, True)
End Sub

Public Sub RevivirUsuario(ByVal UserIndex As Integer)

    With UserList(UserIndex)
    
        If maps(.Pos.map).mapData(.Pos.x, .Pos.y).ObjInfo.index = iObjCuerpoMuerto Then
            Call EraseObj(.Pos.map, .Pos.x, .Pos.y, -1)
            Call WriteObjDelete(UserIndex, .Pos.x, .Pos.y)
        End If
    
        .Stats.Muerto = False
        .flags.Envenenado = 0
        .Stats.MinHP = .Stats.MaxHP
        .Stats.MinMan = .Stats.MaxMan
        .Stats.MinSta = .Stats.MaxSta
        
        Call WriteUpdateHP(UserIndex)
        Call WriteUpdateMana(UserIndex)
        Call WriteUpdateSta(UserIndex)
        
        If LegalPos(.Pos.map, .Pos.x, .Pos.y, PuedeAtravesarAgua(UserIndex)) Then
            Call MakeUserChar(True, .Pos.map, UserIndex, .Pos.map, .Pos.x, .Pos.y)
        Else
            Dim nPos As WorldPos
            
            Call ClosestLegalPos(.Pos, nPos)
            
            If nPos.x > 0 And nPos.y > 0 Then
                Call MakeUserChar(True, nPos.map, UserIndex, .Pos.map, nPos.x, nPos.y)
            End If
        End If
        
        Call RevivirMascotas(UserIndex)
        
    End With
End Sub

Public Sub RevivirMascotas(ByVal UserIndex As Integer)

    With UserList(UserIndex)
    
        If .Pets.Nro > 0 Then
            If MapInfo(.Pos.map).PK Then
                Dim i As Integer
                Dim Nro As Byte
                
                For i = 1 To MaxPets
                
                    If .Pets.Pet(i).Tipo > 0 Then
                    
                        If Nro < .Pets.Nro Then
                            
                            .Pets.Pet(i).index = SpawnNpc(.Pets.Pet(i).Tipo, .Pos, True, False, False)
                            
                            If .Pets.Pet(i).index > 0 Then
                            
                                Nro = Nro + 1
                                
                                NpcList(.Pets.Pet(i).index).name = .Pets.Pet(i).Nombre
                                
                                NpcList(.Pets.Pet(i).index).MaestroUser = UserIndex
                                
                                NpcList(.Pets.Pet(i).index).Stats.MaxHP = .Pets.Pet(i).MaxHP
                                NpcList(.Pets.Pet(i).index).Stats.MinHP = NpcList(.Pets.Pet(i).index).Stats.MaxHP
                                
                                NpcList(.Pets.Pet(i).index).Stats.MinHit = .Pets.Pet(i).MinHit
                                NpcList(.Pets.Pet(i).index).Stats.MaxHit = .Pets.Pet(i).MaxHit
                                
                                NpcList(.Pets.Pet(i).index).Stats.Def = .Pets.Pet(i).Def
                                NpcList(.Pets.Pet(i).index).Stats.DefM = .Pets.Pet(i).DefM
                                
                                Call FollowAmo(.Pets.Pet(i).index)
                            Else
                                .Pets.Pet(i).index = 0
                            End If
                        End If
                    End If
                Next i
                
                If Nro > 0 Then
                    If Nro > 1 Then
                        Call WarpMascotas(UserIndex)
                        Call WriteConsoleMsg(UserIndex, .Pets.Pet(i).Nombre & " (nivel " & .Pets.Pet(i).Lvl & ") fue revivido." & .name, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WarpMascotas(UserIndex)
                        Call WriteConsoleMsg(UserIndex, "Tus mascotas han sido revividas.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub ToogleBoatBody(ByVal UserIndex As Integer)
'Gives boat body depending on user alignment.

    Dim BodyAnim As Integer
    
    With UserList(UserIndex)
        
        .Char.Head = 0
    
        BodyAnim = ObjData(.Inv.Ship).BodyAnim
        
        Select Case BodyAnim
            Case iBarca
                .Char.Body = iBarcaCiuda
            
            Case iGalera
                .Char.Body = iGaleraCiuda
            
            Case iGaleon
                .Char.Body = iGaleonCiuda
        End Select
        
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        .Char.HeadAnim = NingunCasco
    End With

End Sub

Public Sub ChangeUserChar(ByVal UserIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
    With UserList(UserIndex).Char
    
        If .CharIndex < 1 Then
            Exit Sub
        End If
        
        .Body = Body
        .Head = Head
        .Heading = Heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .HeadAnim = Casco
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharChange(Body, Head, Heading, .CharIndex, Arma, Escudo, Casco))
    End With
End Sub

Public Function GetWeaponAnim(ByVal UserIndex As Integer) As Integer

    Dim ObjIndex As Integer
    Dim Tmp As Integer

    With UserList(UserIndex)
        If UsaArco(UserIndex) > 0 Then
            ObjIndex = .Inv.LeftHand
        
        ElseIf UsaArmaNoArco(UserIndex) > 0 Then
            ObjIndex = .Inv.RightHand
        
        Else
            GetWeaponAnim = NingunArma
            
            Exit Function
        End If
        
        Tmp = ObjData(ObjIndex).WeaponRazaEnanaAnim
            
        If Tmp > 0 Then
            If .Raza = eRaza.Enano Or .Raza = eRaza.Gnomo Then
                GetWeaponAnim = Tmp
                Exit Function
            End If
        End If
        
        GetWeaponAnim = ObjData(ObjIndex).WeaponAnim
    End With
End Function

Public Sub EraseUserChar(ByVal UserIndex As Integer, ByVal IsAdminInvisible As Boolean)

On Error GoTo ErrorHandler
    
    With UserList(UserIndex)
        maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex = 0
    
        If .Char.CharIndex = LastChar Then
            Do Until CharList(LastChar) > 0
                LastChar = LastChar - 1
                If LastChar < 2 Then
                    Exit Do
                End If
            Loop
        End If
        
        'Si esta invisible, solo el sabe de su propia existencia, es innecesario borrarlo en los demas clientes
        If IsAdminInvisible Then
            Call EnviarDatosASlot(UserIndex, PrepareMessageCharRemove(.Char.CharIndex))
        Else
            'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharRemove(.Char.CharIndex))
        End If
        
        Call QuitarUser(UserIndex, .Pos.map)
    
        CharList(.Char.CharIndex) = 0
        
        NumChars = NumChars - 1
        
        .Char.CharIndex = 0
    End With
Exit Sub
    
ErrorHandler:
    Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.description)
End Sub

Public Sub RefreshCharStatus(ByVal UserIndex As Integer)
'Refreshes the status and tag of UserIndex

    Dim Guild As String
    Dim guildalineacion As Byte
    
    Dim Barco As ObjData
    
    With UserList(UserIndex)
    
        If .Guild_Id > 0 Then
            Guild = modGuilds.GuildName(.Guild_Id)
            'FALTA GUILD RELACIONSHIP
        End If
        
        If .ShowName Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, .name, Guild, guildalineacion))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, vbNullString, vbNullString, 0))
        End If
        
        'Si esta navengando, se cambia la barca.
        If .flags.Navegando Then
            Call ToogleBoatBody(UserIndex)
            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.HeadAnim)
        End If
    End With
End Sub

Public Sub MakeUserChar(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal UserIndex As Integer, ByVal map As Integer, ByVal x As Byte, ByVal y As Byte)

On Error Resume Next

    Dim CharIndex As Integer
    
    With UserList(UserIndex)

        If InMapBounds(map, x, y) Then
            'If needed make a new Char in list
            If .Char.CharIndex = 0 Then
                CharIndex = NextOpenCharIndex
                .Char.CharIndex = CharIndex
                CharList(CharIndex) = UserIndex
            End If
            
            'Place Char on map if needed
            If toMap Then
                maps(map).mapData(x, y).UserIndex = UserIndex
            End If
            
            'Send make Char command to clients
            Dim bNick As String
            Dim bGuild As String
            Dim bGuildAlineacion As Byte
            Dim bPriv As Byte
            
            bPriv = .flags.Privilegios
            
            'Preparo el nick
            If .ShowName Then
                bNick = .name
                If .Guild_Id > 0 Then
                    bGuild = modGuilds.GuildName(.Guild_Id)
                    'FALTA GUILD RELATIONSHIP
                End If
            Else
                bNick = vbNullString
                bGuild = vbNullString
            End If
            
            Dim NW As Boolean
            
            NW = EsPrincipiante(UserIndex)
            
            If Not toMap Then
                Call UserList(sndIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharCreate _
                    (.Char.Body, .Char.Head, .Char.Heading, .Char.CharIndex, x, y, _
                    .Char.WeaponAnim, .Char.ShieldAnim, .Char.HeadAnim, _
                    bNick, bGuild, bPriv, .Stats.Elv))

            Else
                'Hide the name and guilda - set Privs as normal user
                 Call AgregarUser(UserIndex, .Pos.map)
            End If
        End If
    End With
    
    Exit Sub

End Sub

Public Sub CheckUserLevel(ByVal UserIndex As Integer)
'Chequea que el usuario no halla alcanzado el siguiente nivel,
'de lo contrario le da la vida, mana, etc, correspodiente.
    
    Dim Pts As Integer
    Dim AumentoHIT As Integer
    Dim AumentoMANA As Integer
    Dim AumentoSTA As Integer
    Dim AumentoHP As Integer
    Dim WasPrincipiante As Boolean
    Dim Promedio As Double
    Dim aux As Byte
    Dim DistVida(1 To 5) As Integer
    Dim GI As Integer 'Guild Index
    
    WasPrincipiante = EsPrincipiante(UserIndex)
    
    With UserList(UserIndex)
                            
        While .Stats.Exp >= .Stats.Elu
            'Checkea si alcanzó el máximo nivel
            If .Stats.Elv >= STAT_MaxELV Then
                .Stats.Exp = 0
                .Stats.Elu = 0
                Exit Sub
            End If
            
            'Store it!
            Call Statistics.UserLevelUp(UserIndex)
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, .Pos.x, .Pos.y))
            
            'For multiple levels being riSed at once
            Pts = Pts + 5
            
            .Stats.Elv = .Stats.Elv + 1
            
            .Stats.Exp = .Stats.Exp - .Stats.Elu

            .Stats.Elu = Calcular_ELU(.Stats.Elv)
            
            'Calculo subida de vida
            Promedio = ModVida(.Clase) - (21 - .Stats.Atributos(eAtributos.Constitucion)) * 0.5
            aux = RandomNumber(0, 100)
            
            If Promedio - Int(Promedio) = 0.5 Then
                'Es promedio semientero
                DistVida(1) = DistribucionSemienteraVida(1)
                DistVida(2) = DistVida(1) + DistribucionSemienteraVida(2)
                DistVida(3) = DistVida(2) + DistribucionSemienteraVida(3)
                DistVida(4) = DistVida(3) + DistribucionSemienteraVida(4)
                
                If aux <= DistVida(1) Then
                    AumentoHP = Promedio + 1.5
                ElseIf aux <= DistVida(2) Then
                    AumentoHP = Promedio + 0.5
                ElseIf aux <= DistVida(3) Then
                    AumentoHP = Promedio - 0.5
                Else
                    AumentoHP = Promedio - 1.5
                End If
            Else
                'Es promedio entero
                
                DistVida(1) = DistribucionSemienteraVida(1)
                DistVida(2) = DistVida(1) + DistribucionEnteraVida(2)
                DistVida(3) = DistVida(2) + DistribucionEnteraVida(3)
                DistVida(4) = DistVida(3) + DistribucionEnteraVida(4)
                DistVida(5) = DistVida(4) + DistribucionEnteraVida(5)
                
                If aux <= DistVida(1) Then
                    AumentoHP = Promedio + 2
                ElseIf aux <= DistVida(2) Then
                    AumentoHP = Promedio + 1
                ElseIf aux <= DistVida(3) Then
                    AumentoHP = Promedio
                ElseIf aux <= DistVida(4) Then
                    AumentoHP = Promedio - 1
                Else
                    AumentoHP = Promedio - 2
                End If
            End If
        
            Select Case .Clase
                Case eClass.Warrior
                    AumentoHIT = IIf(.Stats.Elv > 35, 2, 3)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Hunter
                    AumentoHIT = IIf(.Stats.Elv > 35, 2, 3)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Pirat
                    AumentoHIT = 3
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Paladin
                    AumentoHIT = IIf(.Stats.Elv > 35, 1, 3)
                    AumentoMANA = .Stats.Atributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Thief
                    AumentoHIT = 1
                    AumentoSTA = AumentoSTLadron
                
                Case eClass.Mage
                    AumentoHIT = 1
                    AumentoMANA = 3 * .Stats.Atributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTMago
                                
                Case eClass.Cleric
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.Atributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Druid
                    AumentoHIT = 2
                    AumentoMANA = 2.3 * .Stats.Atributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Assasin
                    AumentoHIT = IIf(.Stats.Elv > 35, 1, 3)
                    AumentoMANA = .Stats.Atributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Bard
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.Atributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Bandit
                    AumentoHIT = IIf(.Stats.Elv > 35, 1, 3)
                    
                    AumentoMANA = IIf(.Stats.MaxMan >= 300, 0, .Stats.Atributos(eAtributos.Inteligencia))
                    
                    If AumentoMANA > 0 Then
                        If .Stats.MaxMan + AumentoMANA >= 300 Then
                            AumentoMANA = 300 - .Stats.MaxMan
                        End If
                    End If
                    
                    AumentoSTA = AumentoStBandido
            
                Case Else
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTDef
            End Select
            
            'Actualizamos HitPoints
            .Stats.MaxHP = .Stats.MaxHP + AumentoHP
            
            If .Stats.MaxHP > STAT_MaxHP Then
                .Stats.MaxHP = STAT_MaxHP
            End If
            
            'Actualizamos Stamina
            .Stats.MaxSta = .Stats.MaxSta + AumentoSTA
            
            If .Stats.MaxSta > STAT_MaxSta Then
                .Stats.MaxSta = STAT_MaxSta
            End If
            
            'Actualizamos Mana
            .Stats.MaxMan = .Stats.MaxMan + AumentoMANA
            
            If .Stats.MaxMan > STAT_MaxMan Then
                .Stats.MaxMan = STAT_MaxMan
            End If
            
            If .Clase = eClass.Bandit Then 'mana del bandido restringido hasta 300
                If .Stats.MaxMan > 300 Then
                    .Stats.MaxMan = 300
                End If
            End If
            
            'Actualizamos Golpe Máximo
            .Stats.MaxHit = .Stats.MaxHit + AumentoHIT
            If .Stats.Elv < 35 Then
                If .Stats.MaxHit > STAT_MaxHit_UNDER35 Then
                    .Stats.MaxHit = STAT_MaxHit_UNDER35
                End If
            Else
                If .Stats.MaxHit > STAT_MaxHit_OVER35 Then
                    .Stats.MaxHit = STAT_MaxHit_OVER35
                End If
            End If
            
            'Actualizamos Golpe Mínimo
            .Stats.MinHit = .Stats.MinHit + AumentoHIT
            
            If .Stats.Elv < 35 Then
                If .Stats.MinHit > STAT_MaxHit_UNDER35 Then
                    .Stats.MinHit = STAT_MaxHit_UNDER35
                End If
            Else
                If .Stats.MinHit > STAT_MaxHit_OVER35 Then
                    .Stats.MinHit = STAT_MaxHit_OVER35
                End If
            End If
            
            .Stats.MinHP = .Stats.MaxHP
            .Stats.MinMan = .Stats.MaxMan
            .Stats.MinSta = .Stats.MaxSta
                        
            'If it ceaSed to be a principiante, remove principiante Items and get char away from principiante dungeon
            If Not EsPrincipiante(UserIndex) And WasPrincipiante Then
                .Stats.Gld = .Stats.Gld + (5000)
                Call WriteUpdateGold(UserIndex)
                Call WarpUserChar(UserIndex, 1, 50, 50, True)
            End If
                
            'Send all gained skill points at once (if any)
            If Pts > 0 Then
                Call WriteLevelUp(UserIndex, Pts, AumentoHP, AumentoSTA, AumentoMANA, AumentoHIT)
                .Skills.NroFree = .Skills.NroFree + Pts
            End If
            
            Call LogDesarrollo(.name & " pasó a nivel " & .Stats.Elv & " ganó HP: " & AumentoHP)
            
            Call SaveUser(UserIndex)
            
            'If user is in a party, we modify the variable p_sumaniveleselevados
            Call mdParty.ActualizarSumaNivelesElevados(UserIndex)
            'If user reaches lvl 25 and he is in a guild, we check the guild's alignment and expulses the user if guild has factionary alignment
        
        Wend
        
    End With
    
End Sub

Public Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean
    PuedeAtravesarAgua = (UserList(UserIndex).flags.Navegando = True) Or Not (UserList(UserIndex).flags.Privilegios And PlayerType.User)
End Function

Public Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading)
'Moves the char, sending the Message to everyone in range.
        
    Dim nPos As WorldPos
    Dim Sailing As Boolean
    
    Sailing = PuedeAtravesarAgua(UserIndex)
    nPos = UserList(UserIndex).Pos
    Call HeadtoPos(nHeading, nPos)
        
    If MoveToLegalPos(UserIndex, UserList(UserIndex).Pos.map, nPos.x, nPos.y, Sailing, Not Sailing) Then
        'si no estoy solo en el mapa...
        If MapInfo(UserList(UserIndex).Pos.map).Poblacion > 1 Then
            'Si es un admin invisible, no se avisa a los demas clientes
            If UserList(UserIndex).flags.AdminInvisible < 1 Then
                Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharMove(UserList(UserIndex).Char.CharIndex, nPos.x, nPos.y))
            End If
        End If
        
        Dim oldUserIndex As Integer
        
        oldUserIndex = maps(UserList(UserIndex).Pos.map).mapData(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y).UserIndex
        
        'Si no hay intercambio de pos con nadie
        If oldUserIndex = UserIndex Then
            maps(UserList(UserIndex).Pos.map).mapData(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y).UserIndex = 0
        End If
        
        UserList(UserIndex).Pos = nPos
        UserList(UserIndex).Char.Heading = nHeading
        maps(UserList(UserIndex).Pos.map).mapData(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y).UserIndex = UserIndex
        
        'Actualizamos las áreas de ser necesario
        Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading)
        
    Else
        Call WritePosUpdate(UserIndex)
    End If
    
    If UserList(UserIndex).Counters.Trabajando Then
        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1
    End If
    
    If UserList(UserIndex).Counters.Ocultando Then
        UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
    End If
End Sub

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading
'Returns the Heading opposite to the one pasSed by val.
    
    Select Case nHeading
        Case eHeading.EAST
            InvertHeading = WEST
        Case eHeading.WEST
            InvertHeading = EAST
        Case eHeading.SOUTH
            InvertHeading = NORTH
        Case eHeading.NORTH
            InvertHeading = SOUTH
    End Select
End Function

Public Function NextOpenCharIndex() As Integer
    Dim LoopC As Long

    For LoopC = 1 To MaxCHARS
        If CharList(LoopC) = 0 Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            
            If LoopC > LastChar Then
                LastChar = LoopC
            End If
            
            Exit Function
        End If
    Next LoopC
End Function

Public Function NextOpenUser() As Integer
    Dim LoopC As Integer
    
    For LoopC = 1 To MaxPoblacion + 1
        If LoopC > MaxPoblacion Then
            Exit For
        End If
        
        If (UserList(LoopC).ConnID = -1 And Not UserList(LoopC).flags.Logged) Then
            Exit For
        End If
    Next LoopC
    
    NextOpenUser = LoopC
End Function

Public Sub SendUserStatsTxt(ByVal SendIndex As Integer, ByVal UserIndex As Integer)
    Dim GuildI As Integer
    
    With UserList(UserIndex)
        If UsaArco(UserIndex) > 0 Then
            Call WriteConsoleMsg(SendIndex, "Daño: " & .Stats.MinHit & "/" & .Stats.MaxHit & " (" & ObjData(.Inv.LeftHand).MinHit & "/" & ObjData(.Inv.LeftHand).MaxHit & ")", FontTypeNames.FONTTYPE_INFO)
        
        ElseIf UsaArmaNoArco(UserIndex) > 0 Then
            Call WriteConsoleMsg(SendIndex, "Daño: " & .Stats.MinHit & "/" & .Stats.MaxHit & " (" & ObjData(.Inv.RightHand).MinHit & "/" & ObjData(.Inv.RightHand).MaxHit & ")", FontTypeNames.FONTTYPE_INFO)
        
        Else
            Call WriteConsoleMsg(SendIndex, "Daño: " & .Stats.MinHit & "/" & .Stats.MaxHit, FontTypeNames.FONTTYPE_INFO)
        End If
        
        If .Inv.Body > 0 Then
            If UsaEscudo(UserIndex) > 0 Then
                Call WriteConsoleMsg(SendIndex, "Defensa (torso): " & ObjData(.Inv.Body).MinDef + ObjData(.Inv.LeftHand).MinDef & "/" & ObjData(.Inv.Body).MaxDef + ObjData(.Inv.LeftHand).MaxDef, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(SendIndex, "Defensa (torso): " & ObjData(.Inv.Body).MinDef & "/" & ObjData(.Inv.Body).MaxDef, FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(SendIndex, "Defensa: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        
        If .Inv.Head > 0 Then
            Call WriteConsoleMsg(SendIndex, "Defensa (cabeza): " & ObjData(.Inv.Head).MinDef & "/" & ObjData(.Inv.Head).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(SendIndex, "Defensa (cabeza): 0", FontTypeNames.FONTTYPE_INFO)
        End If
        
        GuildI = .Guild_Id
        If GuildI > 0 Then
            Call WriteConsoleMsg(SendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_INFO)
            If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(.name) Then
                Call WriteConsoleMsg(SendIndex, "Status: Lider", FontTypeNames.FONTTYPE_INFO)
            End If
            'guildpts no tienen objeto
        End If
        
        Dim TempDate As Date
        Dim TempSecs As Long
        Dim TempStr As String
        TempDate = Now - .LogOnTime
        TempSecs = (.UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
        TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " segundos."
        Call WriteConsoleMsg(SendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Total: " & TempStr, FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(SendIndex, "Oro: " & .Stats.Gld & "  Posicion: " & .Pos.x & "," & .Pos.y & " en mapa " & .Pos.map, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Atributos: " & .Stats.Atributos(eAtributos.Fuerza) & ", " & .Stats.Atributos(eAtributos.Agilidad) & ", " & .Stats.Atributos(eAtributos.Inteligencia) & ", " & .Stats.Atributos(eAtributos.Carisma) & ", " & .Stats.Atributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Public Sub SendUserMiniStatsTxt(ByVal SendIndex As Integer, ByVal UserIndex As Integer)
'Shows the users Stats when the user is online.
    
    With UserList(UserIndex)
        Call WriteConsoleMsg(SendIndex, "Pj: " & .name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Matados: " & .Stats.Matados, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Muertes: " & .Stats.Muertes, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Npcs matados: " & .Stats.NpcMatados, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Clase: " & ListaClases(.Clase), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Pena: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Silencio: " & .Counters.Silencio, FontTypeNames.FONTTYPE_INFO)
        
        If .Guild_Id > 0 Then
            Call WriteConsoleMsg(SendIndex, "Clan: " & GuildName(.Guild_Id), FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Public Sub SendUserMiniStatsTxtFromChar(ByVal SendIndex As Integer, ByVal CharName As String)
'Shows the users Stats when the user is offline.

    Dim CharFile As String
    Dim Ban As String
    Dim BanDetailPath As String
    
    BanDetailPath = App.Path & "/logs/" & "BanDetail.dat"
    CharFile = CharPath & CharName & ".chr"
    
    If FileExist(CharFile) Then
        Call WriteConsoleMsg(SendIndex, "Pj: " & CharName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Matados: " & GetVar(CharFile, "STATS", "MATADOS"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Muertes: " & GetVar(CharFile, "STATS", "MUERTES"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Npcs matados: " & GetVar(CharFile, "STATS", "NpcMatados"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Clase: " & ListaClases(GetVar(CharFile, "INIT", "Clase")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Pena: " & GetVar(CharFile, "COUNTERS", "PENA"), FontTypeNames.FONTTYPE_INFO)
        
        If IsNumeric(GetVar(CharFile, "Guild", "Guild_Id")) Then
            Call WriteConsoleMsg(SendIndex, "Clan: " & modGuilds.GuildName(CInt(GetVar(CharFile, "Guild", "Guild_Id"))), FontTypeNames.FONTTYPE_INFO)
        End If
        
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call WriteConsoleMsg(SendIndex, "Ban: " & Ban, FontTypeNames.FONTTYPE_INFO)
        
        If Ban = "1" Then
            Call WriteConsoleMsg(SendIndex, "Ban por: " & GetVar(CharFile, CharName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, CharName, "Reason"), FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(SendIndex, "El pj no existe: " & CharName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Public Sub SendUserInvTxt(ByVal SendIndex As Integer, ByVal UserIndex As Integer)

On Error Resume Next

    Dim j As Long
    
    With UserList(UserIndex)
        Call WriteConsoleMsg(SendIndex, .name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Tiene " & .Inv.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MaxInvSlots
            If .Inv.Obj(j).index > 0 Then
                Call WriteConsoleMsg(SendIndex, "Objeto " & j & " " & ObjData(.Inv.Obj(j).index).name & " Cantidad:" & .Inv.Obj(j).Amount, FontTypeNames.FONTTYPE_INFO)
            End If
        Next j
    End With
End Sub

Public Sub SendUserInvTxtFromChar(ByVal SendIndex As Integer, ByVal CharName As String)
On Error Resume Next

    Dim j As Long
    Dim CharFile As String, Tmp As String
    Dim ObjInd As Long, ObjCant As Long
    
    CharFile = CharPath & CharName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(SendIndex, CharName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, " Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MaxInvSlots
            Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))
            If ObjInd > 0 Then
                Call WriteConsoleMsg(SendIndex, " Objeto " & j & " " & ObjData(ObjInd).name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)
            End If
        Next j
    Else
        Call WriteConsoleMsg(SendIndex, "Personaje inexistente: " & CharName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Public Sub SendSkillsTxt(ByVal SendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
    Dim j As Integer
    
    Call WriteConsoleMsg(SendIndex, UserList(UserIndex).name, FontTypeNames.FONTTYPE_INFO)
    
    For j = 1 To NumSkills
        Call WriteConsoleMsg(SendIndex, SkillName(j) & " = " & UserList(UserIndex).Skills.Skill(j).Elv, FontTypeNames.FONTTYPE_INFO)
    Next j
    
    Call WriteConsoleMsg(SendIndex, " SkillLibres:" & UserList(UserIndex).Skills.NroFree, FontTypeNames.FONTTYPE_INFO)
End Sub

Private Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
    If NpcList(NpcIndex).MaestroUser > 0 Then
        'EsMascotaCiudadano = Not UserList(NpcList(NpcIndex).MaestroUser).Criminal
        
        If EsMascotaCiudadano Then
            Call WriteConsoleMsg(NpcList(NpcIndex).MaestroUser, "¡" & UserList(UserIndex).name & " está atacando a " & NpcList(NpcIndex).name & "!", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
End Function

Public Sub NpcAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    
    'Guardamos el usuario que ataco el npc.
    NpcList(NpcIndex).TargetUser = UserIndex
    NpcList(NpcIndex).TargetNpc = 0

    'Guarda el Npc que estas atacando ahora.
    UserList(UserIndex).flags.NpcAtacado = NpcIndex
    
    If NpcList(NpcIndex).MaestroUser > 0 Then
        Call AllMascotasAtacanUser(UserIndex, NpcList(NpcIndex).MaestroUser)
    End If

End Sub
Public Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean

    If UsaArmaNoArco(UserIndex) > 0 Then
        If ObjData(UserList(UserIndex).Inv.RightHand).Apuñala = 1 Then
            PuedeApuñalar = True
        End If
    End If
End Function

Public Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer, ByVal Acerto As Boolean)
    
    If UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSed < 1 Then
        Exit Sub
    End If
    
    With UserList(UserIndex).Skills.Skill(Skill)

        If .Elv = MaxSkillPoints Then
            Exit Sub
        End If
    
        Dim Lvl As Integer
        Lvl = .Elv
    
        If Lvl > UBound(LevelSkill) Then
            Lvl = UBound(LevelSkill)
        End If
    
        If .Elv >= LevelSkill(Skill).LevelValue Then
            Exit Sub
        End If
    
        If Acerto Then
            .Exp = .Exp + EXP_ACIERTO_SKILL
        Else
            .Exp = .Exp + EXP_FALLO_SKILL
        End If
    
        If .Exp >= .Elu Then
            .Elv = .Elv + 1
            
            Call WriteSkillUp(UserIndex, Skill)
            
            .Exp = .Exp + 10 * .Elv * MultiplicadorExp
                            
            Call WriteUpdateExp(UserIndex)
            Call CheckEluSkill(UserIndex, Skill, False)
        End If
    End With
End Sub

Public Sub UserDie(ByVal UserIndex As Integer, Optional ByVal UserMatador As Integer = 0, Optional ByVal NpcMatador As Integer = 0)
'Muere un usuario

On Error Resume Next

    If UserMatador > 0 Then
        If UserList(UserIndex).Pos.map <> UserList(UserMatador).Pos.map Then
            Exit Sub
        End If
        
    ElseIf NpcMatador > 0 Then
        If UserList(UserIndex).Pos.map <> NpcList(NpcMatador).Pos.map Then
            Exit Sub
        End If
    End If
    
    Dim i As Long
    Dim aN As Integer
    
    With UserList(UserIndex)
        If Not .flags.Privilegios And PlayerType.User Then
            Exit Sub
        End If
                
        'Sonido
        If .Genero = eGenero.Mujer Then
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MUERTE_MUJER)
        Else
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MUERTE_HOMBRE)
        End If
                
        .Stats.MinHP = 0
        .Stats.MinMan = 0
        .Stats.MinSta = 0
        .flags.Envenenado = 0
        .Stats.Muerto = True
        .Counters.Respawn = 60
        .flags.AtacadoPorUser = 0
        
        'No se activa en arenas
        If TriggerZonaPelea(UserIndex, UserIndex) <> TRIGGER6_PERMITE Then
            .flags.SeguroResu = True
            'Call WriteResuscitationSafeOn(UserIndex)
        Else
            .flags.SeguroResu = False
            'Call WriteResuscitationSafeOff(UserIndex)
        End If
        
        aN = .flags.AtacadoPorNpc
        
        If aN > 0 Then
            NpcList(aN).TargetUser = 0
            .flags.AtacadoPorNpc = 0
        End If
        
        aN = .flags.NpcAtacado
        
        If aN > 0 Then
            If NpcList(aN).TargetUser = UserIndex Then
                NpcList(aN).TargetUser = 0
            End If
            .flags.NpcAtacado = 0
        End If
                
        If UserMatador > 0 Then
            Call Muerte(UserIndex, UserMatador)
        End If

        .flags.Paralizado = 0
        .flags.Estupidez = 0
        .flags.Descansando = False
        .flags.Meditando = False
        
        'Invisible
        If .flags.Invisible > 0 Or .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .flags.Invisible = 0
            .Counters.TiempoOculto = 0
            .Counters.Invisibilidad = 0
        End If
        
        'Restauramos el mimetismo
        If .flags.Mimetizado Then
        '    .Char.Body = .CharMimetizado.Body
        '    .Char.Head = .CharMimetizado.Head
        '    .Char.HeadAnim = .CharMimetizado.HeadAnim
        '    .Char.ShieldAnim = .CharMimetizado.ShieldAnim
        '    .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            .Counters.Mimetismo = 0
            .flags.Mimetizado = False
        End If
        
        '<< Restauramos los atributos >>
        If .flags.TomoPocion Then
            For i = 1 To 5
                .Stats.Atributos(i) = .Stats.AtributosBackUP(i)
            Next i
        End If
        
        Call DesinvocarMascotas(UserIndex)
                
        Call EraseUserChar(UserIndex, False)
                
        Dim ObjCuerpoMuerto As Obj
        
        ObjCuerpoMuerto.index = iObjCuerpoMuerto
        ObjCuerpoMuerto.Amount = UserIndex
        
        If maps(.Pos.map).mapData(.Pos.x, .Pos.y).ObjInfo.index > 0 Then
            Dim Pos As WorldPos
        
            Dim NuevaPos As WorldPos

            Pos.map = .Pos.map
            Pos.x = .Pos.x
            Pos.y = .Pos.y
            
            Call Tilelibre(Pos, NuevaPos, ObjCuerpoMuerto.index, False, True)
    
            If NuevaPos.x > 0 And NuevaPos.y > 0 Then
                Call MakeObj(ObjCuerpoMuerto, NuevaPos.map, NuevaPos.x, NuevaPos.y)
                .Pos.x = NuevaPos.x
                .Pos.y = NuevaPos.y
                
                Call EnviarDatosASlot(UserIndex, PrepareMessageCharMove(UserList(UserIndex).Char.CharIndex, .Pos.x, .Pos.y))
            End If

        Else
            Call MakeObj(ObjCuerpoMuerto, .Pos.map, .Pos.x, .Pos.y)
        End If
        
        If TriggerZonaPelea(UserIndex, UserIndex) <> eTrigger6.TRIGGER6_PERMITE Then
            'Si es principiante no pierde el inventario
            If Not EsPrincipiante(UserIndex) Then
                If maps(.Pos.map).mapData(.Pos.x, .Pos.y).Trigger <> 6 Then
                    'Si es pirata y usa un Galeón entonces no explota los Items. (en el agua)
                    If Not (.Clase = eClass.Pirat And .Inv.Ship = 476) Then
                        Call TirarItemsAlMorir(UserIndex)
                        Call TirarOro(CLng((RandomNumber(1, 3) / 100 * .Stats.Gld)), UserIndex)
                    ElseIf UserMatador > 0 Then
                        If UserList(UserMatador).Clase = eClass.Pirat Then
                            Call TirarItemsAlMorir(UserIndex)
                            Call TirarOro(CLng((RandomNumber(1, 3) / 100 * .Stats.Gld)), UserIndex)
                        End If
                    End If
                End If
            End If
        End If

        '<<Castigos por party>>
        If .PartyIndex > 0 Then
            Call mdParty.ObtenerExito(UserIndex, .Stats.Elv * -10 * mdParty.CantMiembros(UserIndex), .Pos.map, .Pos.x, .Pos.y)
        End If
        
        '<<Cerramos comercio seguro>>
        Call LimpiarComercioSeguro(UserIndex)
    End With

End Sub

Public Sub Muerte(ByVal Muerto As Integer, ByVal Atacante As Integer)

    With UserList(Atacante)
        Call WriteConsoleMsg(Atacante, "Mataste a " & UserList(Muerto).name & "!", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(Muerto, "¡" & .name & " te mató!", FontTypeNames.FONTTYPE_FIGHT)
        
        Call FlushBuffer(Muerto)
        
        'Log
        Call LogAsesinato(.name & " mató a " & UserList(Muerto).name)
        
        If EsPrincipiante(Muerto) Then
            Exit Sub
        End If
        
        If .Stats.Elv > UserList(Muerto).Stats.Elv + 10 Then
            Exit Sub
        End If
    
        If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then
            Exit Sub
        End If
        
        If EsCompaniero(Atacante, UserList(Muerto).name) > 0 Then
            Exit Sub
        End If
    
        If LenB(.flags.UltimoMatado) = LenB(UserList(Muerto).name) Then
            If .flags.UltimoMatado = UserList(Muerto).name Then
                Exit Sub
            End If
        End If
        
        .flags.UltimoMatado = UserList(Muerto).name
        .Stats.Matados = .Stats.Matados + 1
        .Stats.MatadosSinMorir = .Stats.MatadosSinMorir + 1

        Dim ExpADar As Integer
        
        ExpADar = 1000 '* MultiplicadorExp
        
        If UserList(Muerto).Stats.Elv < .Stats.Elv + 10 Then
            ExpADar = ExpADar * UserList(Muerto).Stats.Elv \ .Stats.Elv
        End If
        
        If .Stats.MatadosSinMorir > 1 Then
            ExpADar = ExpADar * .Stats.MatadosSinMorir
        End If
        
        If UserList(Muerto).Stats.MatadosSinMorir > 2 Then
            ExpADar = ExpADar * UserList(Muerto).Stats.MatadosSinMorir / 2
        End If
        
        .Stats.Exp = .Stats.Exp + ExpADar
        
        Call WriteUpdateExp(Atacante)
        
        UserList(Muerto).Stats.Muertes = UserList(Muerto).Stats.Muertes + 1
        UserList(Muerto).Stats.MatadosSinMorir = 0
    End With
                    
End Sub

Public Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByVal ObjIndex As Integer, _
              ByVal PuedeAgua As Boolean, ByVal PuedeTierra As Boolean)

On Error GoTo ErrHandler

    Dim Found As Boolean
    Dim LoopC As Integer
    Dim tX As Long
    Dim tY As Long
    
    nPos = Pos
    tX = Pos.x
    tY = Pos.y
    
    LoopC = 1
    
    'La primera posición es valida?
    If LegalPos(Pos.map, nPos.x, nPos.y, PuedeAgua, PuedeTierra, True) Then
        If Not HayObjeto(Pos.map, nPos.x, nPos.y, ObjIndex) Then
            Found = True
        End If
    End If
    
    'Busca en las demas posiciones, en forma de "rombo"
    If Not Found Then
        While (Not Found) And LoopC < 26
            If RhombLegalTilePos(Pos, tX, tY, LoopC, ObjIndex, PuedeAgua, PuedeTierra) Then
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
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error en Tilelibre. Error: " & Err.Number & " - " & Err.description)
End Sub

Public Sub WarpUserChar(ByVal UserIndex As Integer, ByVal map As Integer, ByVal x As Byte, ByVal y As Byte, ByVal FX As Boolean)

    Dim OldMap As Integer
    Dim OldX As Byte
    Dim OldY As Byte
    Dim i As Byte
    
    With UserList(UserIndex)
    
        OldMap = .Pos.map
        OldX = .Pos.x
        OldY = .Pos.y
                
        If .Char.CharIndex > 0 Then
            Call EraseUserChar(UserIndex, .flags.AdminInvisible > 0)
        End If
        
        .Pos.x = x
        .Pos.y = y
        .Pos.map = map
            
        If MapInfo(.Pos.map).Zona = Ciudad Or .Pos.map = Newbie.map Then
            .Hogar = .Pos.map
        End If
        
        If OldMap <> map Then
            Call WriteChangeMap(UserIndex, map)
            
            'Update new Map Users
            MapInfo(map).Poblacion = MapInfo(map).Poblacion + 1
            
            'Update old Map Users
            MapInfo(OldMap).Poblacion = MapInfo(OldMap).Poblacion - 1
            If MapInfo(OldMap).Poblacion < 0 Then
                MapInfo(OldMap).Poblacion = 0
            End If
        End If
        
        Call MakeUserChar(True, map, UserIndex, map, x, y)
        
        Call WriteUpdateUserStats(UserIndex)
 
        'Force a flush, so user Index is in there before it's destroyed for teleporting
        Call FlushBuffer(UserIndex)
        
        'Seguis invisible al pasar de mapa
        If .flags.Invisible > 0 Or .flags.Oculto > 0 And .flags.AdminInvisible < 1 Then
            If Not .flags.Navegando Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
            End If
        End If
        
        'NO Seguis paralizado al pasar de mapa
        If .flags.Paralizado > 0 Then
            .flags.Paralizado = 0
        End If
        
        If .flags.Inmovilizado > 0 Then
            .flags.Inmovilizado = 0
        End If
        
        If FX And .flags.AdminInvisible < 1 Then 'FX
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_WARP, x, y))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(x, y, FXIDs.FX_WARP))
        End If
        
        If .Pets.Nro > 0 Then
            Call WarpMascotas(UserIndex)
        End If
        
        'No puede ser atacado cuando cambia de mapa, por cierto tiempo
        Call IntervaloPermiteSerAtacado(UserIndex, True)
    
    End With
    
End Sub

Public Sub WarpMascotas(ByVal UserIndex As Integer, Optional ByVal Enviar As Boolean = False)

    If UserList(UserIndex).Pets.NroALaVez > 0 Then
        'If Not MapInfo(UserList(UserIndex).Pos.Map).Pk Then
            Call DesinvocarMascotas(UserIndex)
        '    Exit Sub
        'End If
    End If
    
    Dim i As Byte
    Dim PetTiempoDeVida As Integer
    Dim index As Integer
    
    With UserList(UserIndex)
    
        For i = 1 To MaxPets
            
            If .Pets.Pet(i).Tipo > 0 Then
                    
                If .Pets.Pet(i).MinHP > 0 Then
                    
                    If .Pets.Pet(i).Tipo > 0 Then
                        PetTiempoDeVida = 0
                        
                    ElseIf .Pets.NroALaVez > 0 Then
                        index = .Pets.Pet(i).index
            
                        If index > 0 Then
                            'Store data and remove Npc to recreate it after warp
                            PetTiempoDeVida = NpcList(index).Contadores.TiempoExistencia
                            
                            Call QuitarNpc(index)
                        End If
                    End If
                
                    index = SpawnNpc(.Pets.Pet(i).Tipo, .Pos, False, False, False)
            
                    If index < 1 Then
                        Call WriteConsoleMsg(UserIndex, "Tus mascotas no pueden pasar acá.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        .Pets.NroALaVez = .Pets.NroALaVez + 1
                        
                        .Pets.Pet(i).index = index
        
                        NpcList(index).name = .Pets.Pet(i).Nombre
                        
                        'Nos aseguramos de que conserve el hp, si estaba dañado
                        NpcList(index).Stats.MinHP = .Pets.Pet(i).MinHP
                        NpcList(index).Stats.MaxHP = .Pets.Pet(i).MaxHP
                        
                        NpcList(index).Stats.MinHit = .Pets.Pet(i).MinHit
                        NpcList(index).Stats.MaxHit = .Pets.Pet(i).MaxHit
                        
                        NpcList(index).Stats.Def = .Pets.Pet(i).Def
                        NpcList(index).Stats.DefM = .Pets.Pet(i).DefM
                    
                        NpcList(index).Lvl = .Pets.Pet(i).Lvl
                    
                        NpcList(index).MaestroUser = UserIndex
                        NpcList(index).Movement = TipoAI.SigueAmo
                        NpcList(index).TargetUser = UserIndex
                        NpcList(index).TargetNpc = 0
                        NpcList(index).Contadores.TiempoExistencia = PetTiempoDeVida
        
                        Call FollowAmo(index)
                        
                        If Enviar Then
                            With NpcList(index)
                                Call EnviarDatosASlot(UserIndex, PrepareMessageNpcCharCreate(.Char.Body, .Char.Head, .Char.Heading, .Char.CharIndex, .Pos.x, .Pos.y, .name, .Lvl, i))
                            End With
                        End If
                    End If
                    
                    If .Pets.NroALaVez > MaxPets Then
                        Exit Sub
                    End If
                End If
            End If
        Next i
    
    End With
    
End Sub

Public Sub WarpMascota(ByVal UserIndex As Integer, ByVal PetIndex As Integer)
'Warps a pet without changing its stats

    Dim NpcIndex As Integer
    Dim TargetPos As WorldPos

    With UserList(UserIndex)
        
        TargetPos.map = .flags.TargetMap
        TargetPos.x = .flags.TargetX
        TargetPos.y = .flags.TargetY
        
        NpcIndex = .Pets.Pet(PetIndex).index
            
        Call QuitarNpc(NpcIndex)
        
        NpcIndex = SpawnNpc(.Pets.Pet(PetIndex).Tipo, TargetPos, True, False, False)
    
        'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
        'Exception: Pets don't spawn in water if they can't swim
        If NpcIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, "Tus mascotas no pueden estar por acá.", FontTypeNames.FONTTYPE_INFO)
            .Pets.NroALaVez = .Pets.NroALaVez = 0
        Else
            .Pets.NroALaVez = .Pets.NroALaVez + 1
        
            .Pets.Pet(PetIndex).index = NpcIndex

            'Nos aseguramos de que conserve el hp, si estaba dañado
            NpcList(NpcIndex).Stats.MinHP = .Pets.Pet(PetIndex).MinHP
        
            With NpcList(NpcIndex)
                .MaestroUser = UserIndex
                .Movement = TipoAI.SigueAmo
                .TargetUser = UserIndex
                .TargetNpc = 0
            End With
            
            Call FollowAmo(NpcIndex)
        End If
    End With
End Sub

Public Sub DesinvocarMascotas(ByVal UserIndex As Integer)
    
    Dim i As Byte

    With UserList(UserIndex)
        
        If .Pets.NroALaVez < 1 Then
            Exit Sub
        End If
        
        For i = 1 To MaxPets
            If .Pets.Pet(i).index > 0 Then
                Call QuitarNpc(.Pets.Pet(i).index)
            End If
        Next i
        
        .Pets.NroALaVez = 0
    End With

End Sub

Public Sub CerrarUsuario(ByVal UserIndex As Integer)

    Dim HiddenPirat As Boolean
    
    With UserList(UserIndex)
        If .flags.Logged And Not .Counters.Saliendo Then
            If (.flags.Privilegios And PlayerType.User) And MapInfo(.Pos.map).PK And Not UserList(UserIndex).Stats.Muerto Then
                .Counters.Saliendo = True
                .Counters.Salir = IntervaloCerrarConexion
            Else
                Call FlushBuffer(UserIndex)
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
            .Counters.Saliendo = True
            .Counters.Salir = IIf((.flags.Privilegios And PlayerType.User) And MapInfo(.Pos.map).PK And Not UserList(UserIndex).Stats.Muerto, IntervaloCerrarConexion, 0)
            
            If .flags.Invisible > 0 Or .flags.Oculto > 0 Then
                .flags.Invisible = 0
                
                If .flags.Oculto > 0 Then
                    If .flags.Navegando Then
                        If .Clase = eClass.Pirat Then
                            'Pierde la apariencia de fragata fantasmal
                            Call ToogleBoatBody(UserIndex)
                            Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, _
                                                NingunEscudo, NingunCasco)
                            HiddenPirat = True
                        End If
                    End If
                    
                    .flags.Oculto = 0
                End If
                
                'Si esta navegando ya esta visible
                If Not .flags.Navegando Then
                    If .flags.Privilegios And PlayerType.User Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                    End If
                End If
                
                If .Counters.Salir > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Saliendo...", FontTypeNames.FONTTYPE_INFO)
                End If
                
            End If
        End If
    End With
End Sub

Public Sub CancelExit(ByVal UserIndex As Integer)
    If UserList(UserIndex).Counters.Saliendo Then
        'Is the user still connected?
        If UserList(UserIndex).ConnIDValida Then
            UserList(UserIndex).Counters.Saliendo = False
            UserList(UserIndex).Counters.Salir = 0
            Call WriteConsoleMsg(UserIndex, "Salida cancelada.", FontTypeNames.FONTTYPE_WARNING)
        Else
            'Simply reset
            UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(UserIndex).Pos.map).PK, IntervaloCerrarConexion, 0)
        End If
    End If
End Sub

Public Sub CambiarNick(ByVal UserIndex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
'CambiarNick: Cambia el Nick de un slot.

    Dim ViejoNick As String
    Dim ViejoCharBackup As String
    
    If Not UserList(UserIndexDestino).flags.Logged Then
        Exit Sub
    End If
    
    ViejoNick = UserList(UserIndexDestino).name
    
    If User_Exist(ViejoNick) Then
        'hace un backup del char
        ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
        Name CharPath & ViejoNick & ".chr" As ViejoCharBackup
    End If
End Sub

Public Sub SendUserStatsTxtOFF(ByVal SendIndex As Integer, ByVal Nombre As String)
    If Not User_Exist(Nombre) Then
        Call WriteConsoleMsg(SendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(SendIndex, "Estadisticas de: " & Nombre, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Nivel: " & GetVar(CharPath & Nombre & ".chr", "STATS", "elv") & "  EXP: " & GetVar(CharPath & Nombre & ".chr", "STATS", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", "STATS", "elu"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Vitalidad: " & GetVar(CharPath & Nombre & ".chr", "STATS", "MinSta") & "/" & GetVar(CharPath & Nombre & ".chr", "STATS", "MaxSta"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Salud: " & GetVar(CharPath & Nombre & ".chr", "STATS", "MinHP") & "/" & GetVar(CharPath & Nombre & ".chr", "STATS", "MaxHP") & "  Mana: " & GetVar(CharPath & Nombre & ".chr", "STATS", "MinMan") & "/" & GetVar(CharPath & Nombre & ".chr", "STATS", "MaxMan"), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(SendIndex, "Menor Golpe/mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "STATS", "MaxHit"), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(SendIndex, "Oro: " & GetVar(CharPath & Nombre & ".chr", "STATS", "GLD"), FontTypeNames.FONTTYPE_INFO)
        
        Dim TempSecs As Long
        Dim TempStr As String
        TempSecs = GetVar(CharPath & Nombre & ".chr", "INIT", "UpTime")
        TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(SendIndex, "Tiempo Logeado: " & TempStr, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Public Sub SendUserOROTxtFromChar(ByVal SendIndex As Integer, ByVal CharName As String)
    Dim CharFile As String
    
On Error Resume Next
    CharFile = CharPath & CharName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(SendIndex, CharName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, " Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(SendIndex, "Personaje inexistente: " & CharName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Public Function BodyIsBoat(ByVal Body As Integer) As Boolean
'Checks if a given body Index is a boat

'TODO: This should be checked somehow else. This is nasty....
    If Body = iBarcaPk Or _
            Body = iGaleraPk Or Body = iGaleonPk Or Body = iBarcaCiuda Or _
            Body = iGaleraCiuda Or Body = iGaleonCiuda Or Body = iFragataFantasmal Then
        BodyIsBoat = True
    End If
End Function

Public Function IsArena(ByVal UserIndex As Integer) As Boolean
'Returns true if the user is in an Arena
    IsArena = TriggerZonaPelea(UserIndex, UserIndex) = TRIGGER6_PERMITE
End Function

Public Function GetDireccion(ByVal UserIndex As Integer, ByVal OtherUserIndex As Integer) As String
'Devuelve la direccion hacia donde esta el usuario

    Dim x As Byte
    Dim y As Byte
    
    x = UserList(UserIndex).Pos.x - UserList(OtherUserIndex).Pos.x
    y = UserList(UserIndex).Pos.y - UserList(OtherUserIndex).Pos.y
    
    If x = 0 And y > 0 Then
        GetDireccion = "Sur"
    ElseIf x = 0 And y < 0 Then
        GetDireccion = "Norte"
    ElseIf x > 0 And y = 0 Then
        GetDireccion = "Este"
    ElseIf x < 0 And y = 0 Then
        GetDireccion = "Oeste"
    ElseIf x > 0 And y < 0 Then
        GetDireccion = "NorEste"
    ElseIf x < 0 And y < 0 Then
        GetDireccion = "NorOeste"
    ElseIf x > 0 And y > 0 Then
        GetDireccion = "SurEste"
    ElseIf x < 0 And y > 0 Then
        GetDireccion = "SurOeste"
    End If

End Function

Public Function FarthestPet(ByVal UserIndex As Integer) As Byte
'Devuelve el indice de la Mascota mas lejana.

On Error GoTo ErrHandler
    
    Dim PetIndex As Byte
    Dim Distancia As Integer
    Dim OtraDistancia As Integer
    
    With UserList(UserIndex)
        
        For PetIndex = 1 To MaxPets
            'Solo pos invocar criaturas que exitan!
            If .Pets.Pet(PetIndex).index > 0 Then
                'Solo aplica a Mascota, nada de element4ales..
                If NpcList(.Pets.Pet(PetIndex).index).Contadores.TiempoExistencia = 0 Then
                    If FarthestPet = 0 Then
                        'Por si tiene 1 sola Mascota
                        FarthestPet = PetIndex
                        Distancia = Abs(.Pos.x - NpcList(.Pets.Pet(PetIndex).index).Pos.x) + _
                                    Abs(.Pos.y - NpcList(.Pets.Pet(PetIndex).index).Pos.y)
                    Else
                        'La distancia de la próxima Mascota
                        OtraDistancia = Abs(.Pos.x - NpcList(.Pets.Pet(PetIndex).index).Pos.x) + _
                                        Abs(.Pos.y - NpcList(.Pets.Pet(PetIndex).index).Pos.y)
                        'Está más lejos?
                        If OtraDistancia > Distancia Then
                            Distancia = OtraDistancia
                            FarthestPet = PetIndex
                        End If
                    End If
                End If
            End If
        Next PetIndex
    End With

    Exit Function
    
ErrHandler:
    Call LogError("Error en FarthestPet")
End Function

Public Sub CheckEluSkill(ByVal UserIndex As Integer, ByVal Skill As Byte, ByVal Allocation As Boolean)

    With UserList(UserIndex).Skills.Skill(Skill)
        If .Elv < MaxSkillPoints Then
            If Allocation Then
                .Exp = 0
            Else
                .Exp = 0
            End If
            
            .Elu = ELU_SKILL_INICIAL * 1.03 ^ .Elv
        Else
            .Exp = 0
            .Elu = 0
        End If
    End With

End Sub

Public Function HasEnoughItems(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Amount As Long) As Boolean

    Dim Slot As Long
    Dim ItemInvAmount As Long
    
    For Slot = 1 To UserList(UserIndex).Inv.NroItems
        'Si es el Item que busco
        If UserList(UserIndex).Inv.Obj(Slot).index = ObjIndex Then
            'Lo sumo a la Cantidad total
            ItemInvAmount = ItemInvAmount + UserList(UserIndex).Inv.Obj(Slot).Amount
        End If
    Next Slot

    HasEnoughItems = Amount <= ItemInvAmount
End Function

Public Function TotalOfferItems(ByVal ObjIndex As Integer, ByVal UserIndex As Integer) As Long

    Dim Slot As Byte
    
    For Slot = 1 To UserList(UserIndex).Inv.NroItems
            'Si es el Item que busco
        If UserList(UserIndex).ComUsu.Objeto(Slot) = ObjIndex Then
            'Lo sumo a la Cantidad total
            TotalOfferItems = TotalOfferItems + UserList(UserIndex).ComUsu.Cant(Slot)
        End If
    Next Slot

End Function
