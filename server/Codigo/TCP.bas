Attribute VB_Name = "TCP"
Option Explicit

Private Function ValidarCabeza(ByVal UserRaza As Byte, ByVal UserGenero As Byte, ByVal Head As Integer) As Boolean

    Select Case UserGenero
    
        Case eGenero.Hombre
        
            Select Case UserRaza
                Case eRaza.Humano
                    ValidarCabeza = (Head >= HUMANO_H_PRIMER_CABEZA And _
                                    Head <= HUMANO_H_ULTIMA_CABEZA)
                Case eRaza.Elfo
                    ValidarCabeza = (Head >= ELFO_H_PRIMER_CABEZA And _
                                    Head <= ELFO_H_ULTIMA_CABEZA)
                Case eRaza.Drow
                    ValidarCabeza = (Head >= DROW_H_PRIMER_CABEZA And _
                                    Head <= DROW_H_ULTIMA_CABEZA)
                Case eRaza.Enano
                    ValidarCabeza = (Head >= ENANO_H_PRIMER_CABEZA And _
                                    Head <= ENANO_H_ULTIMA_CABEZA)
                Case eRaza.Gnomo
                    ValidarCabeza = (Head >= GNOMO_H_PRIMER_CABEZA And _
                                    Head <= GNOMO_H_ULTIMA_CABEZA)
            End Select
        
        Case eGenero.Mujer
        
            Select Case UserRaza
                Case eRaza.Humano
                    ValidarCabeza = (Head >= HUMANO_M_PRIMER_CABEZA And _
                                    Head <= HUMANO_M_ULTIMA_CABEZA)
                Case eRaza.Elfo
                    ValidarCabeza = (Head >= ELFO_M_PRIMER_CABEZA And _
                                    Head <= ELFO_M_ULTIMA_CABEZA)
                Case eRaza.Drow
                    ValidarCabeza = (Head >= DROW_M_PRIMER_CABEZA And _
                                    Head <= DROW_M_ULTIMA_CABEZA)
                Case eRaza.Enano
                    ValidarCabeza = (Head >= ENANO_M_PRIMER_CABEZA And _
                                    Head <= ENANO_M_ULTIMA_CABEZA)
                Case eRaza.Gnomo
                    ValidarCabeza = (Head >= GNOMO_M_PRIMER_CABEZA And _
                                    Head <= GNOMO_M_ULTIMA_CABEZA)
            End Select
    End Select
        
End Function

Public Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Public Function Numeric(ByVal cad As String) As Boolean
    
    Dim car As Byte
    Dim i As Integer
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If (car < 48 Or car > 57) Then
            Numeric = False
            Exit Function
        End If
        
    Next i
    
    Numeric = True

End Function

Public Function NombrePermitido(ByVal Nombre As String) As Boolean
    
    Dim i As Integer
    
    For i = 1 To UBound(ForbidenNames)
        If InStr(Nombre, ForbidenNames(i)) Then
                NombrePermitido = False
                Exit Function
        End If
    Next i
    
    NombrePermitido = True

End Function

Public Function ValidateSkills(ByVal UserIndex As Integer) As Boolean
    
    Dim LoopC As Integer
     
    For LoopC = 1 To NumSkills
        If UserList(UserIndex).Skills.Skill(LoopC).Elv < 0 Then
            Exit Function
            If UserList(UserIndex).Skills.Skill(LoopC).Elv > MaxSkillPoints Then
                UserList(UserIndex).Skills.Skill(LoopC).Elv = MaxSkillPoints
            End If
        End If
    Next LoopC
    
    ValidateSkills = True
        
End Function

Public Sub ConnectNewUser(ByVal UserIndex As Integer, ByRef name As String, ByRef Password As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal UserClase As eClass, _
                    ByRef Atributos() As Byte, ByRef UserEmail As String, ByVal Head As Integer)
'Conecta un nuevo Usuario
        
    With UserList(UserIndex)
    
        If .flags.Logged Then
            Exit Sub
        End If
        
        If Len(name) < 3 Then
            Exit Sub
        End If
        
        Dim LoopC As Long
        Dim totalatri As Long
        
        '¿Existe el personaje?
        If User_Exist(name) Then
            Call WriteErrorMsg(UserIndex, "Ya existe alguien con ese nombre.")
            Exit Sub
        End If
        
        If Not ValidarCabeza(UserRaza, UserSexo, Head) Then
            Call LogCheating(name & " ha seleccionado la cabeza " & Head & " desde la Ip " & .Ip)
            
            Call WriteErrorMsg(UserIndex, "Cabeza inválida, seleccione una cabeza.")
            Exit Sub
        End If
        
        .flags.Password = Password
                                        
        .name = name
        .Clase = UserClase
        .Raza = UserRaza
        .Genero = UserSexo
        .Email = UserEmail
        
        For LoopC = 1 To NumSkills
            .Skills.Skill(LoopC).Elv = 0
            .Skills.Skill(LoopC).Elu = ELU_SKILL_INICIAL
        Next LoopC
        
        .Skills.NroFree = 10
        
        '%%%%%%%%%%%%% PREVENIR HACKEO DE LOS ATRIBUTOS %%%%%%%%%%%%%
        For LoopC = 1 To NUMATRIBUTOS
            UserList(UserIndex).Stats.Atributos(LoopC) = Atributos(LoopC - 1)
            totalatri = totalatri + Abs(UserList(UserIndex).Stats.Atributos(LoopC))
        Next LoopC
    
        If totalatri <> 76 Then
            Call LogHackAttemp(.name & " intentó hackear los atributos.")
            Call BorrarUsuario(.name)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
                                   
        .Stats.Atributos(eAtributos.Fuerza) = .Stats.Atributos(eAtributos.Fuerza) + ModRaza(UserRaza).Fuerza
        .Stats.Atributos(eAtributos.Agilidad) = .Stats.Atributos(eAtributos.Agilidad) + ModRaza(UserRaza).Agilidad
        .Stats.Atributos(eAtributos.Inteligencia) = .Stats.Atributos(eAtributos.Inteligencia) + ModRaza(UserRaza).Inteligencia
        .Stats.Atributos(eAtributos.Carisma) = .Stats.Atributos(eAtributos.Carisma) + ModRaza(UserRaza).Carisma
        .Stats.Atributos(eAtributos.Constitucion) = .Stats.Atributos(eAtributos.Constitucion) + ModRaza(UserRaza).Constitucion
        
        .Stats.MinSed = 100
        .Stats.MinHam = 100
        
        '<-----------------MANA----------------------->
        Select Case UserClase
            Case eClass.Mage
            
                If .Stats.Atributos(eAtributos.Constitucion) > 18 Then
                    .Stats.MaxHP = 20 + (.Stats.Atributos(eAtributos.Constitucion) - 18) * 1
                Else
                    .Stats.MaxHP = 20
                End If
            
                .Stats.MaxMan = 3 * .Stats.Atributos(eAtributos.Inteligencia)
                
                .Spells.Spell(1) = 2
                .Spells.Spell(2) = 6
                
               .Spells.Nro = 2
    
            Case eClass.Druid
            
                If .Stats.Atributos(eAtributos.Constitucion) > 18 Then
                    .Stats.MaxHP = 20 + (.Stats.Atributos(eAtributos.Constitucion) - 18) * 2
                Else
                    .Stats.MaxHP = 20
                End If
            
                .Stats.MaxMan = 2.3 * .Stats.Atributos(eAtributos.Inteligencia)
                .Spells.Spell(1) = 2
                .Spells.Spell(2) = 1
                
               .Spells.Nro = 2
    
            Case eClass.Cleric
            
                If .Stats.Atributos(eAtributos.Constitucion) > 18 Then
                    .Stats.MaxHP = 20 + (.Stats.Atributos(eAtributos.Constitucion) - 18) * 2
                Else
                    .Stats.MaxHP = 20
                End If
            
                .Stats.MaxMan = 2 * .Stats.Atributos(eAtributos.Inteligencia)
                .Spells.Spell(1) = 2
                .Spells.Spell(2) = 3
            
               .Spells.Nro = 2
                
            Case eClass.Bard
            
                If .Stats.Atributos(eAtributos.Constitucion) > 18 Then
                    .Stats.MaxHP = 20 + (.Stats.Atributos(eAtributos.Constitucion) - 18) * 2
                Else
                    .Stats.MaxHP = 20
                End If
                
                .Stats.MaxMan = 2 * .Stats.Atributos(eAtributos.Inteligencia)
                .Spells.Spell(1) = 2
                
               .Spells.Nro = 1
            
            Case eClass.Assasin Or eClass.Paladin
            
                If .Stats.Atributos(eAtributos.Constitucion) > 18 Then
                    .Stats.MaxHP = 20 + (.Stats.Atributos(eAtributos.Constitucion) - 18) * 3
                Else
                    .Stats.MaxHP = 20
                End If
            
                .Stats.MaxMan = 50
                .Spells.Spell(1) = 2
                
               .Spells.Nro = 1
                
            Case eClass.Bandit
            
                If .Stats.Atributos(eAtributos.Constitucion) > 18 Then
                    .Stats.MaxHP = 20 + (.Stats.Atributos(eAtributos.Constitucion) - 18) * 3
                Else
                    .Stats.MaxHP = 20
                End If
                
                .Stats.MaxMan = 25
                .Spells.Spell(1) = 2
                
               .Spells.Nro = 1
                
            Case eClass.Warrior
                If .Stats.Atributos(eAtributos.Constitucion) > 18 Then
                    .Stats.MaxHP = 25 + (.Stats.Atributos(eAtributos.Constitucion) - 18) * 3
                Else
                    .Stats.MaxHP = 25
                End If
                
                .Stats.MaxMan = 0
                
            Case Else
                If .Stats.Atributos(eAtributos.Constitucion) > 18 Then
                    .Stats.MaxHP = 20 + (.Stats.Atributos(eAtributos.Constitucion) - 18) * 3
                Else
                    .Stats.MaxHP = 20
                End If
                
                .Stats.MaxMan = 0
                
        End Select
        
        .Stats.MaxSta = 5 * .Stats.Atributos(eAtributos.Constitucion)
    
        .Stats.MinHP = .Stats.MaxHP
        .Stats.MinMan = .Stats.MaxMan
        .Stats.MinSta = .Stats.MaxSta
        
        .Stats.MinHit = 1
        .Stats.MaxHit = 2
        
        .Stats.Gld = 0
        
        .Stats.Exp = 0
        .Stats.Elu = 300
        .Stats.Elv = 1
        
        .OrigChar.Head = Head
        
        Dim Ropa As Integer
        
        Select Case UserRaza
            Case eRaza.Humano
                Ropa = 31
            Case eRaza.Elfo
                Ropa = 32
            Case eRaza.Drow
                Ropa = 35
            Case eRaza.Enano
                Ropa = 466
            Case eRaza.Gnomo
                Ropa = 466
        End Select
        
        .Inv.Body = Ropa 'VESTIMENTAS COMÚNES
        
        Select Case UserClase
        
            Case eClass.Assasin 'ASESINO
                .Inv.Obj(1).index = 139 'PESCADO
                .Inv.Obj(1).Amount = 20
                
                .Inv.Obj(2).index = 42 'VINO
                .Inv.Obj(2).Amount = 20
                
                .Inv.Obj(3).index = 38 'POCIÓN ROJA
                .Inv.Obj(3).Amount = 100
    
                .Inv.Obj(4).index = 36 'POCIÓN AMARILLA
                .Inv.Obj(4).Amount = 50
                
                .Inv.RightHand = 15 'DAGA
                            
                .Inv.NroItems = 4
                
            Case eClass.Bandit 'BANDIDO
                .Inv.Obj(1).index = 139 'PESCADO
                .Inv.Obj(1).Amount = 20
                
                .Inv.Obj(2).index = 42 'VINO
                .Inv.Obj(2).Amount = 20
                
                .Inv.Obj(3).index = 38 'POCIÓN ROJA
                .Inv.Obj(3).Amount = 100
            
                .Inv.Obj(4).index = 39 'POCIÓN VERDE
                .Inv.Obj(4).Amount = 50
                
                .Inv.RightHand = 2 'ESPADA LARGA
                            
                .Inv.NroItems = 4
    
            Case eClass.Bard 'BARDO
                .Inv.Obj(1).index = 139 'PESCADO
                .Inv.Obj(1).Amount = 20
                
                .Inv.Obj(2).index = 42 'VINO
                .Inv.Obj(2).Amount = 20
                
                .Inv.Obj(3).index = 38 'POCIÓN ROJA
                .Inv.Obj(3).Amount = 100
            
                .Inv.Obj(4).index = 37 'POCIÓN AZUL
                .Inv.Obj(4).Amount = 100
                
                .Inv.RightHand = 15 'DAGA
                       
                .Inv.NroItems = 4
                
            Case eClass.Cleric 'CLERIGO
                .Inv.Obj(1).index = 139 'PESCADO
                .Inv.Obj(1).Amount = 20
                
                .Inv.Obj(2).index = 42 'VINO
                .Inv.Obj(2).Amount = 20
                
                .Inv.Obj(3).index = 38 'POCIÓN ROJA
                .Inv.Obj(3).Amount = 100
                
                .Inv.Obj(4).index = 37 'POCIÓN AZUL
                .Inv.Obj(4).Amount = 100
                
                .Inv.RightHand = 2 'ESPADA LARGA
                       
                .Inv.NroItems = 4
                
            Case eClass.Druid 'DRUIDA
                .Inv.Obj(1).index = 139 'PESCADO
                .Inv.Obj(1).Amount = 20
                
                .Inv.Obj(2).index = 42 'VINO
                .Inv.Obj(2).Amount = 20
                
                .Inv.Obj(3).index = 38 'POCIÓN ROJA
                .Inv.Obj(3).Amount = 100
                
                .Inv.Obj(4).index = 37 'POCIÓN AZUL
                .Inv.Obj(4).Amount = 150
                
                .Inv.RightHand = 15 'DAGA
                
                .Inv.NroItems = 4
            
            Case eClass.Hunter 'CAZADOR
                .Inv.Obj(1).index = 139 'PESCADO
                .Inv.Obj(1).Amount = 20
                
                .Inv.Obj(2).index = 42 'VINO
                .Inv.Obj(2).Amount = 20
                
                .Inv.Obj(3).index = 38 'POCIÓN ROJA
                .Inv.Obj(3).Amount = 100
                
                .Inv.Obj(4).index = 15 'DAGA
                .Inv.Obj(4).Amount = 1
    
                .Inv.LeftHand = 478  'ARCO SIMPLE
                .Inv.RightHand = 480 'FLECHA
                .Inv.AmmoAmount = 1000
                
                .Inv.NroItems = 4
                
            Case eClass.Mage 'MAGO
                .Inv.Obj(1).index = 139 'PESCADO
                .Inv.Obj(1).Amount = 20
                
                .Inv.Obj(2).index = 42 'VINO
                .Inv.Obj(2).Amount = 20
                
                .Inv.Obj(3).index = 38 'POCIÓN ROJA
                .Inv.Obj(3).Amount = 100
            
                .Inv.Obj(4).index = 37 'POCION AZUL
                .Inv.Obj(4).Amount = 200
                
                .Inv.RightHand = 658 'VARA DE FRESNO
                            
                .Inv.NroItems = 4
                
            Case eClass.Paladin 'PALADÍN
                .Inv.Obj(1).index = 139 'PESCADO
                .Inv.Obj(1).Amount = 20
                
                .Inv.Obj(2).index = 42 'VINO
                .Inv.Obj(2).Amount = 20
                
                .Inv.Obj(3).index = 38 'POCIÓN ROJA
                .Inv.Obj(3).Amount = 100
                
                .Inv.LeftHand = 404 'ESCUDO DE TORTUGA
                .Inv.RightHand = 2 'ESPADA LARGA
                            
                .Inv.NroItems = 3
                
            Case eClass.Pirat 'PIRATA
                .Inv.Obj(1).index = 139 'PESCADO
                .Inv.Obj(1).Amount = 20
                
                .Inv.Obj(2).index = 42 'VINO
                .Inv.Obj(2).Amount = 20
                
                .Inv.Obj(3).index = 38 'POCIÓN ROJA
                .Inv.Obj(3).Amount = 100
    
                .Inv.RightHand = 2 'ESPADA LARGA
                            
                .Inv.NroItems = 3
    
            Case eClass.Thief 'LADRÓN
                .Inv.Obj(1).index = 139 'PESCADO
                .Inv.Obj(1).Amount = 20
                
                .Inv.Obj(2).index = 42 'VINO
                .Inv.Obj(2).Amount = 20
                
                .Inv.Obj(3).index = 38 'POCIÓN ROJA
                .Inv.Obj(3).Amount = 100
                
                .Inv.Obj(4).index = 36 'POCIÓN AMARILLA
                .Inv.Obj(4).Amount = 50
                
                .Inv.RightHand = 15 'DAGA
                        
                .Inv.NroItems = 4
                
            Case eClass.Warrior 'GUERRERO
                .Inv.Obj(1).index = 139 'PESCADO
                .Inv.Obj(1).Amount = 20
                
                .Inv.Obj(2).index = 42 'VINO
                .Inv.Obj(2).Amount = 20
                
                .Inv.Obj(3).index = 38 'POCIÓN ROJA
                .Inv.Obj(3).Amount = 200
            
                .Inv.Obj(4).index = 39 'POCIÓN VERDE
                .Inv.Obj(4).Amount = 50

                .Inv.Head = 132 'CASCO DE HIERRO
                .Inv.RightHand = 3 'HACHA
                           
                .Inv.NroItems = 4
        End Select
                                                
        Call DarImagen(UserIndex)
        
        .LogOnTime = Now
        .UpTime = 0
                        
        .Pos.map = Newbie.map
        .Pos.x = Newbie.x
        .Pos.y = Newbie.y
        
        .Hogar = .Pos.map
        
        Call AgregarPlataforma(UserIndex, .Hogar)

        .flags.Password = Password
        
        Call SaveUser(UserIndex, True)
        
        'Valores Default de facciones al Activar nuevo usuario
        'Call ResetFacciones(UserIndex)
          
        'Vemos que clase de user es (se lo usa para setear los Privilegios al loguear el PJ)
        If EsAdmin(name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
            Call LogGM(name, "Se conecto con ip:" & .Ip)
            .flags.Ignorado = True
            Call DoAdminInvisible(UserIndex)
        ElseIf EsDios(name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
            Call LogGM(name, "Se conecto con ip:" & .Ip)
            .flags.Ignorado = True
            .flags.AdminPerseguible = True
        ElseIf EsSemiDios(name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
            Call LogGM(name, "Se conecto con ip:" & .Ip)
            .flags.Ignorado = True
        ElseIf EsConsejero(name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Consejero
            Call LogGM(name, "Se conecto con ip:" & .Ip)
            .flags.Ignorado = True
        Else
            .flags.Privilegios = .flags.Privilegios Or PlayerType.User
            .flags.AdminPerseguible = True
        End If
        
        .flags.Ignorado = False
        
        'Add RM flag if needed
        If EsRolesMaster(name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster
        End If
                                
        'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
        If maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex > 0 Or maps(.Pos.map).mapData(.Pos.x, .Pos.y).NpcIndex > 0 Or maps(.Pos.map).mapData(.Pos.x, .Pos.y).ObjInfo.Amount > 0 Then
            
            Dim FoundPlace As Boolean
            Dim esAgua As Boolean
            Dim tX As Long
            Dim tY As Long
            
            FoundPlace = False
            esAgua = HayAgua(.Pos.map, .Pos.x, .Pos.y)
            
            For tY = .Pos.y - 1 To .Pos.y + 1
                For tX = .Pos.x - 1 To .Pos.x + 1
                    If esAgua Then
                        'reviso que sea pos legal en agua, que no haya User ni Npc para poder loguear.
                        If LegalPos(.Pos.map, tX, tY, True, False) Then
                            FoundPlace = True
                            Exit For
                        End If
                    Else
                        'reviso que sea pos legal en tierra, que no haya User ni Npc para poder loguear.
                        If LegalPos(.Pos.map, tX, tY, False, True) Then
                            FoundPlace = True
                            Exit For
                        End If
                    End If
                Next tX
                
                If FoundPlace Then
                    Exit For
                End If
            Next tY
            
            If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
                .Pos.x = tX
                .Pos.y = tY
            ElseIf maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex > 0 Then
                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If UserList(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex).ComUsu.DestUsu > 0 Then
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex).ComUsu.DestUsu).flags.Logged Then
                        Call FinComerciarUsu(UserList(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex).ComUsu.DestUsu)
                        Call WriteConsoleMsg(UserList(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                        Call FlushBuffer(UserList(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex).ComUsu.DestUsu)
                    End If
                    'Lo sacamos.
                    If UserList(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex).flags.Logged Then
                        Call FinComerciarUsu(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex)
                        Call WriteErrorMsg(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex, "Alguien se conectó donde estabas, por favor reconectáte.")
                        Call FlushBuffer(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex)
                    End If
                End If
                
                Call CloseSocket(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex)
            End If
        End If
        
        .ShowName = True
                
        If .flags.Privilegios <> PlayerType.User Then
            .flags.ChatColor = &H80FF80
        ElseIf EsPrincipiante(UserIndex) Then
            .flags.ChatColor = &HC0FFFF
        Else
            .flags.ChatColor = &HFFC0C0
        End If
        
        Call WriteChangeMap(UserIndex, .Pos.map)

        Call MakeUserChar(True, .Pos.map, UserIndex, .Pos.map, .Pos.x, .Pos.y)
        
        If .Pets.Nro > 0 Then
            Call WarpMascotas(UserIndex, True)
        End If
                
        Call CheckUserLevel(UserIndex)

        Call WriteUpdateUserStats(UserIndex)

        .flags.Logged = True
        
        If .flags.Privilegios = PlayerType.User Then
            Poblacion = Poblacion + 1
            
            Call DB_RS_Open("SELECT * FROM people WHERE `name`='" & .name & "'")
            DB_RS.Update
            DB_RS.Close
        End If
        
        MapInfo(.Pos.map).Poblacion = MapInfo(.Pos.map).Poblacion + 1
        
        If Poblacion > RecordPoblacion Then
            frmMain.Poblacion.Caption = "Población: " & Poblacion & "!"
            frmMain.Poblacion.ForeColor = RGB(200, 0, 0)
        
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Record de población! " & "Hay " & Poblacion & " habitantes.", FontTypeNames.FONTTYPE_INFO))
            RecordPoblacion = Poblacion
            Call WriteVar(ServidorIni, "Init", "Record", str(RecordPoblacion))
            Call RegistrarEstadisticas
        Else
            frmMain.Poblacion.Caption = "Población: " & Poblacion
            frmMain.Poblacion.ForeColor = vbWhite
        End If
                
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePopulation())
        
        Call Base.OnlinePlayers
    
        Call WriteUpdateHungerAndThirst(UserIndex)
        Call WriteUpdateStrenghtAndDexterity(UserIndex)
    
        Call WriteInventory(UserIndex)
        
        If .Belt.NroItems > 0 Then
            Call WriteBeltInv(UserIndex)
        End If
        
        If .Stats.MaxMan > 0 And .Spells.Nro > 0 Then
            Call WriteSpells(UserIndex)
        End If
        
        If .Skills.NroFree > 0 Then
            Call WriteFreeSkills(UserIndex)
        End If
        
        Call WriteSkills(UserIndex)
        
        'Companieros
        If .Compas.Nro > 0 Then
            Call WriteCompas(UserIndex)
        End If
        
        .flags.Seguro = True
        
        Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCreateFX(.Pos.x, .Pos.y, FXIDs.FX_WARP))
        
        Call EnviarDatosASlot(UserIndex, PrepareMessageWeather)
        
        Call WriteLogged(UserIndex)
          
        'Call SendMOTD(UserIndex)
  
        If haciendoBK Then
            Call WritePauseToggle(UserIndex)
            Call WriteConsoleMsg(UserIndex, "Servidor> Espera algunos segundos, WorldSave está ejecutándose.", FontTypeNames.FONTTYPE_SERVER)
        End If
        
        If EnPausa Then
            Call WritePauseToggle(UserIndex)
            Call WriteConsoleMsg(UserIndex, "Servidor> Abraxas está en pausa. Probá ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)
        End If
        
        'If Not .flags.Privilegios And PlayerType.User Then
            'Call WriteShowGMPanelForm(UserIndex)
        'End If

        'Esta protegido del ataque de npcs por unos segundos, si no realiza ninguna accion
        Call IntervaloPermiteSerAtacado(UserIndex, True)
        
        'If Lloviendo Then Call WriteRainToggle(UserIndex)
        
        'Load the user statistics
        Call Statistics.UserConnected(UserIndex)

    End With
    
End Sub

Public Sub CloseSocket(ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        If UserIndex = LastUser Then
            Do Until UserList(LastUser).flags.Logged
                LastUser = LastUser - 1
                If LastUser < 1 Then
                    Exit Do
                End If
            Loop
        End If
        
        If .ConnID <> -1 Then
            Call CloseSocketSL(UserIndex)
        End If
        
        Dim CentinelaIndex As Byte
        CentinelaIndex = .flags.CentinelaIndex
        
        If CentinelaIndex <> 0 Then
            Call modCentinela.CentinelaUserLogout(CentinelaIndex)
        End If
        
        'mato los comercios seguros
        If .ComUsu.DestUsu > 0 Then
            If UserList(.ComUsu.DestUsu).flags.Logged Then
                If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(.ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(.ComUsu.DestUsu)
                    Call FlushBuffer(.ComUsu.DestUsu)
                End If
            End If
        End If
        
        'Empty buffer for reuse
        Call .incomingData.ReadASCIIStringFixed(.incomingData.length)
        
        If .flags.Logged Then
            Call CloseUser(UserIndex)
        Else
            Call ResetUserSlot(UserIndex)
        End If
        
        .ConnID = -1
        .ConnIDValida = False
        
        Exit Sub

    End With
    
ErrHandler:
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False
    Call ResetUserSlot(UserIndex)

    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.description & " - UserIndex = " & UserIndex)
End Sub

Public Sub CloseSocketSL(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
        Call BorraSlotsock(UserList(UserIndex).ConnID)
        Call WSApiCloseSocket(UserList(UserIndex).ConnID)
        UserList(UserIndex).ConnIDValida = False
        Call Revisar_Subasta(UserIndex)
    End If
    
End Sub

'Send an string to a Slot
Public Function EnviarDatosASlot(ByVal UserIndex As Integer, ByRef Datos As String) As Long

    On Error GoTo Err
    
    Dim ret As Long
    
    ret = WsApiEnviar(UserIndex, Datos)
    
    If ret > 0 And ret <> WSAEWOULDBLOCK Then
        'Close the socket avoiding any critical error
        Call CloseSocketSL(UserIndex)
        Call CerrarUsuario(UserIndex)
    End If
    
    Exit Function
    
Err:

End Function
Public Function EstaPCarea(index As Integer, Index2 As Integer) As Boolean

    Dim x As Byte
    Dim y As Byte
    For y = UserList(index).Pos.y - MinYBorder + 1 To UserList(index).Pos.y + MinYBorder - 1
            For x = UserList(index).Pos.x - MinXBorder + 1 To UserList(index).Pos.x + MinXBorder - 1
    
                If maps(UserList(index).Pos.map).mapData(x, y).UserIndex = Index2 Then
                    EstaPCarea = True
                    Exit Function
                End If
            
            Next x
    Next y

    EstaPCarea = False
End Function

Public Function HayPCarea(Pos As WorldPos) As Boolean


Dim x As Integer, y As Integer
For y = Pos.y - MinYBorder + 1 To Pos.y + MinYBorder - 1
        For x = Pos.x - MinXBorder + 1 To Pos.x + MinXBorder - 1
            If x > 0 And y > 0 And x < 101 And y < 101 Then
                If maps(Pos.map).mapData(x, y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next x
Next y
HayPCarea = False
End Function

Public Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean

    Dim x As Integer, y As Integer
    
    For y = Pos.y - MinYBorder + 1 To Pos.y + MinYBorder - 1
            For x = Pos.x - MinXBorder + 1 To Pos.x + MinXBorder - 1
                If maps(Pos.map).mapData(x, y).ObjInfo.index = ObjIndex Then
                    HayOBJarea = True
                    Exit Function
                End If
            
            Next x
    Next y
    
    HayOBJarea = False

End Function

Public Sub ConnectUser(ByVal UserIndex As Integer, ByVal name As String, ByVal Pass As String, ByVal SoyYo As Boolean)
    
    Dim N As Integer
    Dim tStr As String
    
    With UserList(UserIndex)
        .flags.Password = Pass
        
        If .flags.Logged Then
            'Kick player ( and leave Char inside :D )!
            Call CloseSocketSL(UserIndex)
            Call CerrarUsuario(UserIndex)
            Exit Sub
        End If
        
        'Controlamos no pasar el máximo de población
        If Poblacion >= MaxPoblacion Then
            Call WriteErrorMsg(UserIndex, "El servidor ha alcanzado el máximo de población soportado, por favor vuelva a intertarlo mas tarde.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
                
        'Vemos que clase de user es (se lo usa para setear los Privilegios al loguear el PJ)
        If EsAdmin(name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
            Call LogGM(name, "Se conecto con ip:" & .Ip)
            .flags.Ignorado = True
        ElseIf EsDios(name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
            Call LogGM(name, "Se conecto con ip:" & .Ip)
            .flags.Ignorado = True
            .flags.AdminPerseguible = True
        ElseIf EsSemiDios(name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
            Call LogGM(name, "Se conecto con ip:" & .Ip)
            .flags.Ignorado = True
        ElseIf EsConsejero(name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Consejero
            Call LogGM(name, "Se conecto con ip:" & .Ip)
            .flags.Ignorado = True
        Else
            .flags.Privilegios = .flags.Privilegios Or PlayerType.User
            .flags.AdminPerseguible = True
        End If
        
        .flags.Ignorado = False
        
        'Add RM flag if needed
        If EsRolesMaster(name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster
        End If
        
        If ServerSoloGMs > 0 Then
            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
                Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
                Call FlushBuffer(UserIndex)
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
        End If
        
        'SEGUIR CON ESTOOOO-----
        If name = "Anaibol" Then
            .ShowName = False
        Else
            .ShowName = True
        End If
            
        'If SoyYo And Not User_Exist(Name) Then
        '    .Name = "Anaibol"
        '    Call LoadUser(UserIndex)
'            .Name = Name
        'Else
             .name = name
            Call LoadUser(UserIndex)
        'End If
        'SEGUIR CON ESTOOOO-----
        
        If Not ValidateSkills(UserIndex) Then
            Call WriteErrorMsg(UserIndex, "Error en los skills del personaje.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If

        If Not MapaValido(.Pos.map) Or .Pos.x < MinXBorder Or .Pos.x > MaxXBorder Or .Pos.y < MinYBorder Or .Pos.y > MaxYBorder Then
            Call WriteErrorMsg(UserIndex, "Estabas en un lugar inválido.")
            Call RespawnearUsuario(UserIndex)
        End If

        'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
        If maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex > 0 Or maps(.Pos.map).mapData(.Pos.x, .Pos.y).NpcIndex > 0 Or maps(.Pos.map).mapData(.Pos.x, .Pos.y).ObjInfo.Amount > 0 Then
            
            Dim FoundPlace As Boolean
            Dim esAgua As Boolean
            Dim tX As Long
            Dim tY As Long
            
            FoundPlace = False
            esAgua = HayAgua(.Pos.map, .Pos.x, .Pos.y)
            
            For tY = .Pos.y - 1 To .Pos.y + 1
                For tX = .Pos.x - 1 To .Pos.x + 1
                    If esAgua Then
                        'reviso que sea pos legal en agua, que no haya User ni Npc para poder loguear.
                        If LegalPos(.Pos.map, tX, tY, True, False) Then
                            FoundPlace = True
                            Exit For
                        End If
                    Else
                        'reviso que sea pos legal en tierra, que no haya User ni Npc para poder loguear.
                        If LegalPos(.Pos.map, tX, tY, False, True) Then
                            FoundPlace = True
                            Exit For
                        End If
                    End If
                Next tX
                
                If FoundPlace Then
                    Exit For
                End If
            Next tY
            
            If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
                .Pos.x = tX
                .Pos.y = tY
            
            ElseIf maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex > 0 Then
                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If UserList(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex).ComUsu.DestUsu > 0 Then
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex).ComUsu.DestUsu).flags.Logged Then
                        Call FinComerciarUsu(UserList(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex).ComUsu.DestUsu)
                        Call WriteConsoleMsg(UserList(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se desconecó.", FontTypeNames.FONTTYPE_TALK)
                        Call FlushBuffer(UserList(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex).ComUsu.DestUsu)
                    End If
                    'Lo sacamos.
                    If UserList(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex).flags.Logged Then
                        Call FinComerciarUsu(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex)
                        Call WriteErrorMsg(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex, "Alguien se conectó donde estabas, por favor reconectáte.")
                        Call FlushBuffer(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex)
                    End If
                End If
                
                Call CloseSocket(maps(.Pos.map).mapData(.Pos.x, .Pos.y).UserIndex)
            End If
        
        End If
        
        If .flags.Privilegios <> PlayerType.User Then
            .flags.ChatColor = &H80FF80
        ElseIf EsPrincipiante(UserIndex) Then
            .flags.ChatColor = &HC0FFFF
        Else
            .flags.ChatColor = &HFFC0C0
        End If
        
        .LogOnTime = Now

        Call DarImagen(UserIndex)
        
        Call WriteChangeMap(UserIndex, .Pos.map)
        
        Call MakeUserChar(True, .Pos.map, UserIndex, .Pos.map, .Pos.x, .Pos.y)
        
        If .Pets.Nro > 0 Then
            Call WarpMascotas(UserIndex, True)
        End If
        
        'If .flags.Privilegios And PlayerType.Admin Then
        '    UserList(UserIndex).flags.Ignorado = True
        '    Call DoAdminInvisible(UserIndex)
        'End If
        
        Call CheckUserLevel(UserIndex)
        
        Call WriteUpdateUserStats(UserIndex)
        
        .flags.Logged = True
        
        Call DB_RS_Open("SELECT * FROM people WHERE `name`='" & .name & "'")
        DB_RS!Logged = 1
        DB_RS.Update
        DB_RS.Close
        
        MapInfo(.Pos.map).Poblacion = MapInfo(.Pos.map).Poblacion + 1
        
        If Poblacion > 0 Then
            frmMain.Poblacion.Caption = "Población: " & Poblacion
                    
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePopulation())
            
            If Poblacion > RecordPoblacion Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Record de población." & "Hay " & Poblacion & " habitantes.", FontTypeNames.FONTTYPE_INFO))
                RecordPoblacion = Poblacion
                Call WriteVar(ServidorIni, "INIT", "Record", str(RecordPoblacion))
            End If
        End If
                
        Call Base.OnlinePlayers

        If .flags.Privilegios = PlayerType.User Then
            Poblacion = Poblacion + 1
        End If

        Call WriteUpdateHungerAndThirst(UserIndex)
        Call WriteUpdateStrenghtAndDexterity(UserIndex)
        
        Call WriteInventory(UserIndex)
        
        If .Belt.NroItems > 0 Then
            Call WriteBeltInv(UserIndex)
        End If
        
        If .Stats.MaxMan > 0 And .Spells.Nro > 0 Then
            Call WriteSpells(UserIndex)
        End If
        
        If .Skills.NroFree > 0 Then
            Call WriteFreeSkills(UserIndex)
        End If
        
        Call WriteSkills(UserIndex)
        
        'Companieros
        If .Compas.Nro > 0 Then
            Call WriteCompas(UserIndex)
        End If
        
        Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCreateFX(.Pos.x, .Pos.y, FXIDs.FX_WARP))
        
        Call EnviarDatosASlot(UserIndex, PrepareMessageWeather)
        
        Call WriteLogged(UserIndex)
        
        'Call SendMOTD(UserIndex)
        
        'If .Guild_Id > 0 Then
        'Call modGuilds.SendGuildNews(UserIndex)
        'End If
        
        If haciendoBK Then
            Call WritePauseToggle(UserIndex)
            Call WriteConsoleMsg(UserIndex, "Servidor> Espera algunos segundos, WorldSave está ejecutándose.", FontTypeNames.FONTTYPE_SERVER)
        End If
        
        If EnPausa Then
            Call WritePauseToggle(UserIndex)
            Call WriteConsoleMsg(UserIndex, "Servidor> Abraxas está en pausa. Probá ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)
        End If
        
        'If Not .flags.Privilegios And PlayerType.User Then
            'Call WriteShowGMPanelForm(UserIndex)
        'End If

        'Esta protegido del ataque de npcs por unos segundos, si no realiza ninguna accion
        Call IntervaloPermiteSerAtacado(UserIndex, True)
        
        'If Lloviendo Then Call WriteRainToggle(UserIndex)
        
        tStr = modGuilds.a_ObtenerRechazoDeChar(.name)
        
        If LenB(tStr) > 0 Then
            Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al guilda fue rechazada. El guilda te explica que: " & tStr)
        End If
        
        'Load the user statistics
        Call Statistics.UserConnected(UserIndex)
                        
        'N = FreeFile
        'Log
        'Open App.Path & "/logs/connect.log" For Append Shared As #N
        'Print #N, .Name & " ha entrado al juego. UserIndex:" & UserIndex & " " & Time & " " & Date
        'Close #N
    End With
    
End Sub

Public Sub SendMOTD(ByVal UserIndex As Integer)
    Dim j As Long
    
    For j = 1 To MaxLines
        Call WriteConsoleMsg(UserIndex, Motd(j).texto, FontTypeNames.FONTTYPE_INFOBOLD)
    Next j
End Sub

Public Sub ResetContadores(ByVal UserIndex As Integer)
'Resetea todos los valores generales y las stats
    
    With UserList(UserIndex).Counters
        .AGUACounter = 0
        .AttackCounter = 0
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Paralisis = 0
        .Pena = 0
        .Silencio = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Trabajando = 0
        .Ocultando = 0
        .bPuedeMeditar = False
        .Lava = 0
        .Mimetismo = 0
        .Saliendo = False
        .Salir = 0
        .TiempoOculto = 0
        .TimerMagiaGolpe = 0
        .TimerGolpeMagia = 0
        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeUsarArco = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0
        .Respawn = 0
        .EnPlataforma = 0
    End With
End Sub

Public Sub ResetCharInfo(ByVal UserIndex As Integer)
'Resetea todos los valores generales y las stats
    With UserList(UserIndex).Char
        .Body = 0
        .HeadAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .Loops = 0
        .Heading = 0
        .Loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Public Sub ResetBasicUserInfo(ByVal UserIndex As Integer)
'Resetea todos los valores generales y las stats
    With UserList(UserIndex)
        '.Id = 0
        .name = vbNullString
        .Desc = vbNullString
        .DescRM = vbNullString
        .Pos.map = 0
        .Pos.x = 0
        .Pos.y = 0
        .Ip = vbNullString
        .Clase = 0
        .Email = vbNullString
        .Genero = 0
        .Raza = 0
        
        .PartyIndex = 0
        .PartySolicitud = 0
        
        .Skills.NroFree = 0
        
        With .Stats
            .BankGld = 0
            .Elv = 0
            .Elu = 0
            .Exp = 0
            .Def = 0
            .NpcMatados = 0
            .Muertes = 0
            .Matados = 0
            .Gld = 0
            .Atributos(1) = 0
            .Atributos(2) = 0
            .Atributos(3) = 0
            .Atributos(4) = 0
            .Atributos(5) = 0
            .AtributosBackUP(1) = 0
            .AtributosBackUP(2) = 0
            .AtributosBackUP(3) = 0
            .AtributosBackUP(4) = 0
            .AtributosBackUP(5) = 0
        End With
        
    End With
End Sub

Public Sub ResetGuildInfo(ByVal UserIndex As Integer)
    If UserList(UserIndex).EscucheGuilda > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(UserIndex, UserList(UserIndex).EscucheGuilda)
        UserList(UserIndex).EscucheGuilda = 0
    End If
    If UserList(UserIndex).Guild_Id > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(UserIndex, UserList(UserIndex).Guild_Id)
    End If
    UserList(UserIndex).Guild_Id = 0
End Sub

Public Sub ResetUserFlags(ByVal UserIndex As Integer)
'Resetea todos los valores generales y las stats
    With UserList(UserIndex).flags
        .Comerciando = False
        .Ban = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .StatsChanged = 0
        .TargetUser = 0
        .TargetNpc = 0
        .TargetNpcTipo = eNpcType.Comun
        .TargetObjIndex = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Descansando = False
        .Navegando = False
        .Oculto = 0
        .Envenenado = 0
        .Invisible = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = False
        .Privilegios = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .Hechizo = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .CentinelaOK = False
        .AdminPerseguible = False
        .UltimoMatado = vbNullString
        .UltimoMensaje = 0
        .Mimetizado = False
    End With
End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
    With UserList(UserIndex).ComUsu
        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(UserIndex)
        End If
    End With
End Sub

Public Sub ResetUserSlot(ByVal UserIndex As Integer)

    Dim i As Byte
    
    With UserList(UserIndex)
        .ConnIDValida = False
        .ConnID = -1
        
        Call LimpiarComercioSeguro(UserIndex)
        Call ResetContadores(UserIndex)
        Call ResetGuildInfo(UserIndex)
        Call ResetCharInfo(UserIndex)
        Call ResetBasicUserInfo(UserIndex)
        Call ResetUserFlags(UserIndex)
        Call ResetUserInventario(UserIndex)
        Call ResetUserCinturon(UserIndex)
        Call ResetUserHechizos(UserIndex)
        Call ResetUserCompanieros(UserIndex)
        Call ResetUserMascotas(UserIndex)
        Call ResetUserBanco(UserIndex)
        Call ResetUserPlataformas(UserIndex)
    End With

    With UserList(UserIndex).ComUsu
        .Acepto = False
        
        For i = 1 To Max_OFFER_SLOTS
            .Cant(i) = 0
            .Objeto(i) = 0
        Next i
        
        .GoldAmount = 0
        .DestNick = vbNullString
        .DestUsu = 0
    End With

End Sub

Public Sub CloseUser(ByVal UserIndex As Integer)
    
On Error Resume Next

    Dim N As Integer
    Dim LoopC As Integer
    Dim map As Integer
    Dim name As String
    Dim i As Integer
    
    With UserList(UserIndex)
    
        If .Stats.Muerto Then
            Call RespawnearUsuario(UserIndex)
        End If
        
        Dim aN As Integer
        
        aN = .flags.AtacadoPorNpc
        
        If aN > 0 Then
              NpcList(aN).TargetUser = 0
        End If
        
        aN = .flags.NpcAtacado
        
        If aN > 0 Then
            If NpcList(aN).TargetUser = UserIndex Then
                NpcList(aN).TargetUser = 0
            End If
        End If
        
        .flags.AtacadoPorNpc = 0
        .flags.NpcAtacado = 0
        
        map = .Pos.map
        name = .name
        
        .Char.FX = 0
                
        .flags.Logged = False
        
        Call DB_RS_Open("SELECT * FROM people WHERE `name`='" & name & "'")
        DB_RS!Logged = 0
        DB_RS.Update
        DB_RS.Close
        
        If Poblacion > 0 Then
            If .flags.Privilegios = PlayerType.User Then
                Poblacion = Poblacion - 1
                frmMain.Poblacion.Caption = "Población: " & Poblacion
                Call SendData(SendTarget.ToAll, 0, PrepareMessagePopulation())
                Call Base.OnlinePlayers
            End If
        End If
        
        'Le devolvemos el body y head originales
        If .flags.AdminInvisible > 0 Then
            Call DoAdminInvisible(UserIndex)
        End If
        
        'si esta en party le devolvemos la experiencia
        If .PartyIndex > 0 Then
            Call mdParty.SalirDeParty(UserIndex)
        End If
        
        Call DesinvocarMascotas(UserIndex)
        
        'Grabamos el personaje del usuario
        Call SaveUser(UserIndex)

        'Companieros
        If .Compas.Nro > 0 Then
            Dim CompaIndex As Integer
        
            For LoopC = 1 To MaxCompaSlots
                If LenB(.Compas.Compa(LoopC)) > 0 Then
                
                    CompaIndex = NameIndex(.Compas.Compa(LoopC))
                    
                    If CompaIndex > 0 Then
                        Dim LoopC2 As Byte
                        
                        For LoopC2 = 1 To MaxCompaSlots
                            If LenB(UserList(CompaIndex).Compas.Compa(LoopC2)) > 0 Then
                                If UserList(CompaIndex).Compas.Compa(LoopC2) = .name Then
                                    Call WriteCompaDisconnected(CompaIndex, LoopC2)
                                    Exit For
                                End If
                            End If
                        Next LoopC2
                    End If
                End If
            Next LoopC
        End If
        
        Call EraseUserChar(UserIndex, .flags.AdminInvisible > 0)
        
        'Update Map Users
        MapInfo(map).Poblacion = MapInfo(map).Poblacion - 1
        
        If MapInfo(map).Poblacion < 0 Then
            MapInfo(map).Poblacion = 0
        End If
        
        'Si el usuario habia dejado un msg en la gm's queue lo borramos
        If Ayuda.Existe(.name) Then
            Call Ayuda.Quitar(.name)
        End If
        
        Call ResetUserSlot(UserIndex)
        
        N = FreeFile(1)
        Open App.Path & "/logs/connect.log" For Append Shared As #N
        Print #N, name & " ha dejado el juego. " & "User Index:" & UserIndex & " " & Time & " " & Date
        Close #N
        
    End With
    
End Sub

Public Sub ReloadSokcet()
On Error GoTo ErrHandler

    Call LogApiSock("ReloadSokcet() " & Poblacion & " " & LastUser & " " & MaxPoblacion)
    
    If Poblacion < 1 Then
        Call WSApiReiniciarSockets
    Else
'Call apiclosesocket(SockListen)
'SockListen = ListenForConnect(Puerto, hWndMsg, vbnullstring)
    End If

Exit Sub
ErrHandler:
    Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.description)

End Sub

Public Sub EcharPjsNoPrivilegiados()
Dim LoopC As Long

For LoopC = 1 To LastUser
    If UserList(LoopC).flags.Logged And UserList(LoopC).ConnID > 1 And UserList(LoopC).ConnIDValida Then
        If UserList(LoopC).flags.Privilegios And PlayerType.User Then
            Call CloseSocket(LoopC)
        End If
    End If
Next LoopC

End Sub
