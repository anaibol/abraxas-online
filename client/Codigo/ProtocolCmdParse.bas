Attribute VB_Name = "modProtocolCmdParse"
Option Explicit

Public Enum eNumber_Types
    ent_Byte
    ent_Integer
    ent_Long
    ent_Trigger
End Enum

Public Sub ParseUserCommand(ByVal RawCommand As String)
'Interpreta, valida y ejecuta el comando ingresado.

On Error Resume Next

    RawCommand = RTrim$(RawCommand)

    Dim TmpArgos() As String
    
    Dim Comando As String
    Dim Mensaje As String
    Dim ArgumentosAll() As String
    Dim ArgumentosRaw As String
    Dim Argumentos2() As String
    Dim Argumentos3() As String
    Dim Argumentos4() As String
    Dim CantidadArgumentos As Long
    Dim notNullArguments As Boolean
    
    Dim tmpArr() As String
    Dim tmpInt As Integer
    
    Dim Name As String
    'TmpArgs: Un array de a lo sumo dos elementos,
    'el primero es el comando (hasta el primer espacio)
    'y el segundo elemento es el resto. Si no hay argumentos
    'devuelve un array de un solo elemento
    TmpArgos = Split(RawCommand, " ", 2)
    
    Comando = UCase$(TmpArgos(0))
    
    If UBound(TmpArgos) > 0 Then
        'El string en crudo que este despues del primer espacio
        ArgumentosRaw = TmpArgos(1)
        
        'veo que los argumentos no sean nulos
        notNullArguments = LenB(Trim$(ArgumentosRaw))
        
        'Un array separado por blancos, con tantos elementos como
        'se pueda
        ArgumentosAll = Split(TmpArgos(1), " ")
        
        'Cantidad de argumentos. En ESTE PUNTO el minimo es 1
        CantidadArgumentos = UBound(ArgumentosAll) + 1
        
        'Los siguientes arrays tienen A LO SUMO, COMO MAXIMO
        '2, 3 y 4 elementos respectivamente. Eso significa
        'que pueden tener menos, por lo que es imperativo
        'preguntar por CantidadArgumentos.
        
        Argumentos2 = Split(TmpArgos(1), " ", 2)
        Argumentos3 = Split(TmpArgos(1), " ", 3)
        Argumentos4 = Split(TmpArgos(1), " ", 4)
    Else
        CantidadArgumentos = 0
    End If
    
    If Left$(Comando, 1) = "/" Then
        'Comando normal
        
        LastParsedString = RawCommand
        
        Select Case Comando
        
            Case "/BORRARME"
            
                UserPassword = InputBox("Ingresa tu clave.")
                UserEmail = InputBox("Ingresa tu mail.")
                
                Call WriteKillChar
            
            Case "/SUBASTAR"
            
                If MapData(UserPos.x, UserPos.y).Obj.Amount > 0 Then
                    If MapData(UserPos.x, UserPos.y).Obj.ObjType = otGuita Then
                        Call ShowConsoleMsg("No podés subastar monedas de oro.")
                    Else
                        Call WriteAuctionCreate
                    End If
                Else
                    Call ShowConsoleMsg("Para subastar objetos primero tenés que pararte sobre ellos.")
                End If
                
            Case "/OFERTAR"
                If notNullArguments And ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
                    Call WriteAuctionBid(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Debes introducir la Cantidad de monedas de oro a ofertar. Escribe /OFERTAR X, siendo X la Cantidad de monedas.")
                End If
                
            Case "/SUBASTA"
                Call WriteAuctionView
                                        
            Case "/SALIR"
                Call WriteQuit
                
            Case "/DEJARGUILDA"
                Call WriteGuildLeave
                
            Case "/BALANCE"
                If UserMuerto Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                    Exit Sub
                End If
                
                Call WriteRequestAccountState
                
            Case "/QUIETO"
            
                If UserMuerto Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                    Exit Sub
                End If
                
                If SelectedCharIndex < 1 Or Charlist(SelectedCharIndex).MascoIndex < 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Primero tenés que seleccionar la mascota.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                
                    Exit Sub
                Else
                    Call WritePetStand(Charlist(SelectedCharIndex).MascoIndex)
                    
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(Charlist(SelectedCharIndex).Nombre & " se quedará quieto.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                End If
                
            Case "/SEGUIR"
            
                If UserMuerto Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                    Exit Sub
                End If
                
                If SelectedCharIndex < 1 Or Charlist(SelectedCharIndex).MascoIndex < 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Primero tenés que seleccionar la mascota.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                
                    Exit Sub
                Else
                    Call WritePetFollow(Charlist(SelectedCharIndex).MascoIndex)
                    
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(Charlist(SelectedCharIndex).Nombre & " te seguirá.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                End If
                                
            Case "/LIBERAR"
            
                Exit Sub
                
                If UserMuerto Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                    Exit Sub
                End If
                
                If SelectedCharIndex < 0 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Primero tenés que seleccionar la mascota.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                    Exit Sub
                Else
                    Call WriteReleasePet
                End If
                                
            Case "/DESCANSAR"
            
                If UserMuerto Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                    Exit Sub
                End If
                
                If UserMoving Then
                    Exit Sub
                End If
                        
                If Descansando = False Then
                    If UserMinHP <> UserMaxHP Or UserMinSTA <> UserMaxSTA Then
                        Descansando = True
                    End If
                Else
                    Descansando = False
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Dejás de descansar.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                End If
                                      
                Call WriteRest
                
            Case "/EST"
                Call WriteRequestStats
            
            Case "/AYUDA"
                Call WriteHelp
                
            Case "/COMERCIAR"
                If UserMuerto Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                    Exit Sub
                
                ElseIf Comerciando Then 'Comerciando
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Ya estás comerciando", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                    Exit Sub
                End If
                
                Call WriteCommerceStart
                
            Case "/BOVEDA"
                If UserMuerto Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                    Exit Sub
                End If
                Call WriteBankStart
                                                                    
            Case "/MOTD"
                Call WriteRequestMOTD
                
            Case "/UPTIME"
                Call WriteUpTime
                
            Case "/SALIRPARTY"
                Call WritePartyLeave
                
            Case "/CREARPARTY"
                If UserMuerto Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                    Exit Sub
                End If
                Call WritePartyCreate
                
            Case "/PARTY"
                If UserMuerto Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                    Exit Sub
                End If
                Call WritePartyJoin
                
            Case "/ENCUESTA"
                If CantidadArgumentos = 0 Then
                    'Version sin argumentos: Inquiry
                    Call WriteInquiry
                Else
                    'Version con argumentos: InquiryVote
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Byte) Then
                        Call WriteInquiryVote(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Para votar una opcion, escribe /encuesta NUMERODEOPCION, por ejemplo para votar la opcion 1, escribe /encuesta 1.")
                    End If
                End If
            
            Case "/CENTINELA"
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                        Call WriteCentinelReport(CInt(ArgumentosRaw))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("El código de verificación debe ser numerico. Utilice /centinela X, siendo X el código de verificación.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /centinela X, siendo X el código de verificación.")
                End If
        
            Case "/ONLINE"
                Call WriteOnline
        
            Case "/ONLINECLAN"
                Call WriteGuildOnline
                                
            Case "/ONLINEPARTY"
                Call WritePartyOnline
                                
            Case "/ROL"
                If notNullArguments Then
                    Call WriteRoleMasterRequest(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba una pregunta.")
                End If
                
            Case "/GM"
                Call WriteGMRequest
                
            Case "/BUG"
                If notNullArguments Then
                    Call WriteBugReport(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba una descripción del bug.")
                End If
            
            Case "/DESC"
                If UserMuerto Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                    Exit Sub
                End If
                
                If Len(ArgumentosRaw) < 0 Or Len(ArgumentosRaw) > 50 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Tu descripción debe tener entre 1 y 50 carácteres.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                    Exit Sub
                End If
                
                Call WriteChangeDescription(ArgumentosRaw)
            
            Case "/VOTO"
                If notNullArguments Then
                    Call WriteGuildVote(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /voto NICKNAME.")
                End If
               
            Case "/PENAS"
                If notNullArguments Then
                    Call WritePunishments(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /penas NICKNAME.")
                End If
                
            Case "/clave"
                Call frmChangePassword.Show(vbModal, frmMain)
            
            Case "/APOSTAR"
                If UserMuerto Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
                    End With
                    Exit Sub
                End If
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                        Call WriteGamble(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Cantidad incorrecta. Utilice /apostar Cantidad.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /apostar Cantidad.")
                End If
                
            Case "/DENUNCIAR"
                If notNullArguments Then
                    Call WriteDenounce(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Formule su denuncia.")
                End If
                
            Case "/FUNDARGUILDA"
                
                Call ShowConsoleMsg("Las guildas se encuentran deshabilitadas.")
            
                'If UserLvl >= 35 Then
                '    Call WriteGuildFundate
                'Else
                '    Call ShowConsoleMsg("Para fundar una guilda tu nivel debe ser de 35 o más y tener 100 puntos de habilidad en Liderazgo.")
                'End If
                        
            Case "/EPARTY"
                If notNullArguments Then
                    Call WritePartyKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /Eparty NICKNAME.")
                End If
                
            Case "/PARTYLIDER"
                If notNullArguments Then
                    Call WritePartySetLeader(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /partylider NICKNAME.")
                End If
                
            Case "/ACCEPTPARTY"
                If notNullArguments Then
                    Call WritePartyAcceptMember(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /acceptparty NICKNAME.")
                End If

            'BEGIN GM COMMANDS
            Case "/GMSG"
                If notNullArguments Then
                    Call WriteGMMessage(ArgumentosRaw)
                End If
                
            Case "/NAME"
                Call WriteShowName
                
            Case "/IRCERCA"
                If notNullArguments Then
                    Call WriteGoNearby(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ircerca NICKNAME.")
                End If
                
            Case "/REM"
                If notNullArguments Then
                    Call WriteComment(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un comentario.")
                End If
            
            Case "/HORA"
                Call WriteServerTime
            
            Case "/DONDE"
                If notNullArguments Then
                    Call WriteWhere(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /donde NICKNAME.")
                End If
                
            Case "/NENE"
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                        Call WriteCreaturesInMap(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Mapa incorrecto. Utilice /nene MAPA.")
                    End If
                Else
                    'Por default, toma el mapa en el que esta
                    Call WriteCreaturesInMap(UserMap)
                End If
                
            Case "/TELEPLOC"
                Call WriteWarpMeToTarget
                
            Case "/T"
                If notNullArguments And CantidadArgumentos >= 4 Then
                    If ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) Then
                        Call WriteWarpChar(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /T NICKNAME MAPA X Y.")
                    End If
                ElseIf CantidadArgumentos = 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        'Por defecto, si no se indica el nombre, se teletransporta el mismo usuario
                        Call WriteWarpChar("YO", ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                    ElseIf ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        'Por defecto, si no se indica el mapa, se teletransporta al mismo donde esta el usuario
                        Call WriteWarpChar(ArgumentosAll(0), UserMap, ArgumentosAll(1), ArgumentosAll(2))
                    Else
                        'No uso ningun formato por defecto
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /T NICKNAME MAPA X Y.")
                    End If
                ElseIf CantidadArgumentos = 2 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) Then
                        'Por defecto, se considera que se quiere unicamente cambiar las coordenadas del usuario, en el mismo mapa
                        Call WriteWarpChar("YO", UserMap, ArgumentosAll(0), ArgumentosAll(1))
                    Else
                        'No uso ningun formato por defecto
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /T NICKNAME MAPA X Y.")
                    End If
                
                'SI PONGO PONEMOS SOLO EL MAPA, NOS MANDA AL MAPA, 50, 50 :D
                ElseIf CantidadArgumentos = 1 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                        Call WriteWarpChar("YO", ArgumentosAll(0), 50, 50)
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /T NICKNAME MAPA X Y.")
                End If
                
            Case "/SILENCIAR"
                If notNullArguments Then
                    If CantidadArgumentos = 1 Then
                        Call WriteSilence(ArgumentosAll(0), ArgumentosAll(1))
                    Else
                        Call WriteSilence(ArgumentosAll(0), 10)
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /silenciar NICKNAME (MINUTOS).")
                End If
                
            Case "/SHOW"
                If notNullArguments Then
                    Select Case UCase$(ArgumentosAll(0))
                        Case "SOS"
                            Call WriteSOSShowList
                            
                        Case "INT"
                            Call WriteShowServerForm
                    End Select
                End If
                
            Case "/I"
                If notNullArguments Then
                    Call WriteGoToChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /I NICKNAME.")
                End If
        
            Case "/INVI"
                Call WriteInvisible
                
            Case "/P"
                Call WriteGMPanel
                
            Case "/TRABAJANDO"
                Call WriteWorking
                
            Case "/OCULTANDO"
                Call WriteHiding
                
            Case "/CARCEL"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@")
                    If UBound(tmpArr) = 2 Then
                        If ValidNumber(tmpArr(2), eNumber_Types.ent_Byte) Then
                            Call WriteJail(tmpArr(0), tmpArr(1), tmpArr(2))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Tiempo incorrecto. Utilice /carcel NICKNAME@MOTIVO@TIEMPO.")
                        End If
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /carcel NICKNAME@MOTIVO@TIEMPO.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /carcel NICKNAME@MOTIVO@TIEMPO.")
                End If
                
            Case "/RMATA"
                Call WriteKillNPC
                
            Case "/ADVERTENCIA"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        Call WriteWarnUser(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /advertencia NICKNAME@MOTIVO.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /advertencia NICKNAME@MOTIVO.")
                End If
                
            Case "/MOD"
                If notNullArguments And CantidadArgumentos >= 3 Then
                    Select Case UCase$(ArgumentosAll(1))
                        Case "BODY"
                            tmpInt = eEditOptions.eo_Body
                        
                        Case "HEAD"
                            tmpInt = eEditOptions.eo_Head
                        
                        Case "ORO"
                            tmpInt = eEditOptions.eo_Gold
                        
                        Case "LEVEL"
                            tmpInt = eEditOptions.eo_Level
                        
                        Case "SKILLS"
                            tmpInt = eEditOptions.eo_Skills
                        
                        Case "SKILLSLIBRES"
                            tmpInt = eEditOptions.eo_SkillPointsLeft
                        
                        Case "CLASE"
                            tmpInt = eEditOptions.eo_Class
                        
                        Case "EXP"
                            tmpInt = eEditOptions.eo_Experience
                                                
                        Case "SEX"
                            tmpInt = eEditOptions.eo_Sex
                            
                        Case "RAZA"
                            tmpInt = eEditOptions.eo_Raza
                        
                        Case "AGREGAR"
                            tmpInt = eEditOptions.eo_addGold
                        
                        Case Else
                            tmpInt = -1
                    End Select
                    
                    If tmpInt > 0 Then
                        If CantidadArgumentos = 3 Then
                            Call WriteEditChar(ArgumentosAll(0), tmpInt, ArgumentosAll(2), vbNullString)
                        Else
                            Call WriteEditChar(ArgumentosAll(0), tmpInt, ArgumentosAll(2), ArgumentosAll(3))
                        End If
                    Else
                        'Avisar que no exite el comando
                        Call ShowConsoleMsg("Comando incorrecto.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros.")
                End If
            
            Case "/INFO"
                If notNullArguments Then
                    Call WriteRequestCharInfo(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /info NICKNAME.")
                End If
                
            Case "/STAT"
                If notNullArguments Then
                    Call WriteRequestCharStats(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /stat NICKNAME.")
                End If
                
            Case "/BAL"
                If notNullArguments Then
                    Call WriteRequestCharGold(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /bal NICKNAME.")
                End If
                
            Case "/INV"
                If notNullArguments Then
                    Call WriteRequestCharInv(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /inv NICKNAME.")
                End If
                
            Case "/BOV"
                If notNullArguments Then
                    Call WriteRequestCharBank(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /bov NICKNAME.")
                End If
                
            Case "/SKILLS"
                If notNullArguments Then
                    Call WriteRequestCharSkills(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /SKILLS NICKNAME.")
                End If
                
            Case "/R"
                If notNullArguments Then
                    Call WriteReviveChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /R NICKNAME.")
                End If
                
            Case "/ONGM"
                Call WriteOnlineGM
                
            Case "/ONMAP"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                        Call WriteOnlineMap(ArgumentosAll(0))
                    Else
                        Call ShowConsoleMsg("Mapa incorrecto.")
                    End If
                Else
                    Call WriteOnlineMap(UserMap)
                End If
                  
            Case "/E"
                If notNullArguments Then
                    Call WriteKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /E NICKNAME.")
                End If
                
            Case "/EJECUTAR"
                If notNullArguments Then
                    Call WriteExecute(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ejecutar NICKNAME.")
                End If
                
            Case "/BAN"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        Call WriteBanChar(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /ban NICKNAME@MOTIVO.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ban NICKNAME@MOTIVO.")
                End If
                
            Case "/UNBAN"
                If notNullArguments Then
                    Call WriteUnbanChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /unban NICKNAME.")
                End If
                
            Case "/FOLLOW"
                Call WriteNPCFollow
                
            Case "/S"
                If notNullArguments Then
                    Call WriteSummonChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /S NICKNAME.")
                End If
                
            Case "/CC"
                Call WriteSpawnListRequest
                
            Case "/RESETINV"
                Call WriteResetNpcInv
                
            Case "/LIMPIAR"
                Call WriteCleanWorld
                
            Case "/RMSG"
                If notNullArguments Then
                    Call WriteServerMessage(ArgumentosRaw)
                End If
                
            Case "/NICK2IP"
                If notNullArguments Then
                    Call WriteNickToIP(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /nick2ip NICKNAME.")
                End If
                
            Case "/IP2NICK"
                If notNullArguments Then
                    If validipv4str(ArgumentosRaw) Then
                        Call WriteIPToNick(str2ipv4l(ArgumentosRaw))
                    Else
                        'No es una IP
                        Call ShowConsoleMsg("IP incorrecta. Utilice /ip2nick IP.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ip2nick IP.")
                End If
                
            Case "/ONGUILDA"
                If notNullArguments Then
                    Call WriteGuildOnlineMembers(ArgumentosRaw)
                Else
                    'Avisar sintaxis incorrecta
                    Call ShowConsoleMsg("Utilice /ONGUILDA nombre de la guilda.")
                End If
                
            Case "/CT"
                If notNullArguments And CantidadArgumentos >= 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And _
                        ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        
                        If CantidadArgumentos = 3 Then
                        Call WriteTeleportCreate(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                    Else
                            If ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) Then
                                Call WriteTeleportCreate(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                            Else
                                'No es numerico
                                Call ShowConsoleMsg("Valor incorrecto. Utilice /ct MAPA X Y RADIO(Opcional).")
                            End If
                        End If
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /ct MAPA X Y RADIO(Opcional).")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ct MAPA X Y RADIO(Opcional).")
                End If
                
            Case "/DT"
                Call WriteTeleportDestroy
                
            Case "/LLUVIA"
                Call WriteRainToggle
                
            Case "/SETDESC"
                Call WriteSetCharDescription(ArgumentosRaw)
            
            Case "/FORCEMP3MAP"
                If notNullArguments Then
                    'elegir el mapa es opcional
                    If CantidadArgumentos = 1 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                            'eviamos un mapa nulo para que tome el del usuario.
                            Call WriteForceMP3ToMap(ArgumentosAll(0), 0)
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("MP3 incorrecto. Utilice /FORCEMP3map MP3 MAPA, siendo el mapa opcional.")
                        End If
                    Else
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                            Call WriteForceMP3ToMap(ArgumentosAll(0), ArgumentosAll(1))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Valor incorrecto. Utilice /FORCEMP3MAP MP3 MAPA, siendo el mapa opcional.")
                        End If
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Utilice /FORCEMP3MAP MP3 MAPA, siendo el mapa opcional.")
                End If
                
            Case "/FORCEWAVMAP"
                If notNullArguments Then
                    'elegir la posicion es opcional
                    If CantidadArgumentos = 1 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                            'eviamos una posicion nula para que tome la del usuario.
                            Call WriteForceWAVEToMap(ArgumentosAll(0), 0, 0, 0)
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
                        End If
                    ElseIf CantidadArgumentos = 4 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) Then
                            Call WriteForceWAVEToMap(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
                        End If
                    Else
                        'Avisar que falta el parametro
                        Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
                End If
            
            Case "/TALKAS"
                If notNullArguments Then
                    Call WriteTalkAsNPC(ArgumentosRaw)
                End If
        
            Case "/MASSDEST"
                Call WriteDestroyAllItemsInArea
                    
            Case "/ESTUPIDO"
                If notNullArguments Then
                    Call WrItemakeDumb(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /estupido NICKNAME.")
                End If
                
            Case "/NOESTUPIDO"
                If notNullArguments Then
                    Call WrItemakeDumbNoMore(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /noestupido NICKNAME.")
                End If
                
            Case "/DUMPSECURITY"
                Call WriteDumpIPTables
                
            Case "/TRIGGER"
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Trigger) Then
                        Call WriteSetTrigger(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Numero incorrecto. Utilice /trigger NUMERO.")
                    End If
                Else
                    'Version sin parametro
                    Call WriteAskTrigger
                End If
                
            Case "/BANIPLIST"
                Call WriteBannedIPList
                
            Case "/BANIPRELOAD"
                Call WriteBannedIPReload
                
            Case "/MIEMBROSGUILDA"
                If notNullArguments Then
                    Call WriteGuildMemberList(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /MIEMBROSGUILDA GUILDNAME.")
                End If
                
            Case "/BANGUILDA"
                If notNullArguments Then
                    Call WriteGuildBan(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /BANGUILDA GUILDNAME.")
                End If
                
            Case "/BANIP"
                If CantidadArgumentos >= 2 Then
                    If validipv4str(ArgumentosAll(0)) Then
                        Call WriteBanIP(True, str2ipv4l(ArgumentosAll(0)), vbNullString, Right$(ArgumentosRaw, Len(ArgumentosRaw) - Len(ArgumentosAll(0)) - 1))
                    Else
                        'No es una IP, es un nick
                        Call WriteBanIP(False, str2ipv4l("0.0.0.0"), ArgumentosAll(0), Right$(ArgumentosRaw, Len(ArgumentosRaw) - Len(ArgumentosAll(0)) - 1))
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /banip IP motivo o /banip nick motivo.")
                End If
                
            Case "/UNBANIP"
                If notNullArguments Then
                    If validipv4str(ArgumentosRaw) Then
                        Call WriteUnbanIP(str2ipv4l(ArgumentosRaw))
                    Else
                        'No es una IP
                        Call ShowConsoleMsg("IP incorrecta. Utilice /unbanip IP.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /unbanip IP.")
                End If
                
            Case "/CI"
                If notNullArguments Then
                    If CantidadArgumentos = 1 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                            Call WriteCreateItem(ArgumentosAll(0), 1)
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Objeto incorrecto. Utilice /CI OBJETO (Cantidad).")
                        End If
                    ElseIf CantidadArgumentos = 2 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                            If ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Long) And ArgumentosAll(1) > 0 Then
                                Call WriteCreateItem(ArgumentosAll(0), ArgumentosAll(1))
                            Else
                                'No es numerico
                                Call ShowConsoleMsg("Cantidad incorrecta. Utilice /CI OBJETO (Cantidad).")
                            End If
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Objeto incorrecto. Utilice /CI OBJETO (Cantidad).")
                        End If
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /CI OBJETO (Cantidad).")
                End If
                
            Case "/DEST"
                Call WriteDestroyItems
                    
            Case "/FORCEMP3"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                        Call WriteForceMP3All(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("MP3 incorrecto. Utilice /FORCEMP3 MP3.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /FORCEMP3 MP3.")
                End If
    
            Case "/FORCEWAV"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                        Call WriteForceWAVEAll(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Wav incorrecto. Utilice /forcewav WAV.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /forcewav WAV.")
                End If
                
            Case "/BORRARPENA"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 3)
                    If UBound(tmpArr) = 2 Then
                        Call WriteRemovePunishment(tmpArr(0), tmpArr(1), tmpArr(2))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /borrarpena NICK@PENA@NuevaPena.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /borrarpena NICK@PENA@NuevaPena.")
                End If
                
            Case "/BLOQ", "/BLOQUEAR"
                Call WriteTileBlockedToggle
                
            Case "/M"
                Call WriteKillNPCNoRespawn
        
            Case "/MASSKILL"
                Call WriteKillAllNearbyNPCs
                
            Case "/LASTIP"
                If notNullArguments Then
                    Call WriteLastIP(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /lastip NICKNAME.")
                End If
    
            Case "/MOTDCAMBIA"
                Call WriteChangeMOTD
                
            Case "/SMSG"
                If notNullArguments Then
                    Call WriteSystemMessage(ArgumentosRaw)
                End If
                
            Case "/ACC"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                        Call WriteCreateNPC(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Npc incorrecto. Utilice /acc NPC.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /acc NPC.")
                End If
                
            Case "/RACC"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                        Call WriteCreateNPCWithRespawn(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Npc incorrecto. Utilice /racc NPC.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /racc NPC.")
                End If
                                
            Case "/HABILITAR"
                Call WriteServerOpenToUsersToggle
            
            Case "/APAGAR"
                Call WriteTurnOffServer
                                
            Case "/RAJARCLAN"
                If notNullArguments Then
                    Call WriteRemoveCharFromGuild(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /rajarclan NICKNAME.")
                End If
                
            Case "/LASTEMAIL"
                If notNullArguments Then
                    Call WriteRequestCharMail(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /lastemail NICKNAME.")
                End If
                
            Case "/APASS"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        Call WriteAlterPassword(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /apass PJSINPASS@PJCONPASS.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /apass PJSINPASS@PJCONPASS.")
                End If
                
            Case "/AEMAIL"
                If notNullArguments Then
                    tmpArr = AEMAILSplit(ArgumentosRaw)
                    If LenB(tmpArr(0)) = 0 Then
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /aemail NICKNAME-NUEVOMAIL.")
                    Else
                        Call WriteAlterMail(tmpArr(0), tmpArr(1))
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /aemail NICKNAME-NUEVOMAIL.")
                End If
                
            Case "/ANAME"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        Call WriteAlterName(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /aname ORIGEN@DESTINO.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /aname ORIGEN@DESTINO.")
                End If
                
            Case "/SLOT"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        If ValidNumber(tmpArr(1), eNumber_Types.ent_Byte) Then
                            Call WriteCheckSlot(tmpArr(0), tmpArr(1))
                        Else
                            'Faltan o sobran los parametros con el formato propio
                            Call ShowConsoleMsg("Formato incorrecto. Utilice /Slot NICK@Slot.")
                        End If
                    Else
                        'Faltan o sobran los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /Slot NICK@Slot.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /Slot NICK@Slot.")
                End If
                
            Case "/CENTINELAACTIVADO"
                Call WriteToggleCentinelActivated
                
            Case "/DOBACKUP"
                Call WriteDoBackup
                
            Case "/SHOWCMSG"
                If notNullArguments Then
                    Call WriteShowGuildMessages(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /showcmsg GUILDNAME.")
                End If
                
            Case "/B"
                If notNullArguments And LenB(ArgumentosAll(0)) > 7 Then
                    Call WriteSearchObj(ArgumentosAll(0))
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /B OBJETO.")
                End If
                
            Case "/GUARDARMAPA", "/GRABARMAPA"
                Call WriteSaveMap
                
            Case "/MODMAPINFO" 'PK, BACKUP
                If CantidadArgumentos > 1 Then
                    Select Case UCase$(ArgumentosAll(0))
                        Case "PK" '"/MODMAPINFO PK"
                            Call WriteChangeMapInfoPK(ArgumentosAll(1) = "1")
                        
                        Case "BACKUP" '"/MODMAPINFO BACKUP"
                            Call WriteChangeMapInfoBackup(ArgumentosAll(1) = "1")
                        
                        Case "RESTRINGIR" '/MODMAPINFO RESTRINGIR
                            Call WriteChangeMapInfoRestricted(ArgumentosAll(1))
                        
                        Case "MAGIASINEFECTO" '/MODMAPINFO MAGIASINEFECTO
                            Call WriteChangeMapInfoNoMagic(ArgumentosAll(1))
                        
                        Case "INVISINEFECTO" '/MODMAPINFO INVISINEFECTO
                            Call WriteChangeMapInfoNoInvi(ArgumentosAll(1))
                        
                        Case "RESUSINEFECTO" '/MODMAPINFO RESUSINEFECTO
                            Call WriteChangeMapInfoNoResu(ArgumentosAll(1))
                        
                        Case "TERRENO" '/MODMAPINFO TERRENO
                            Call WriteChangeMapInfoLand(ArgumentosAll(1))
                        
                        Case "ZONA" '/MODMAPINFO ZONA
                            Call WriteChangeMapInfoZone(ArgumentosAll(1))
                    End Select
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parametros. Opciones: PK, BACKUP, RESTRINGIR, MAGIASINEFECTO, INVISINEFECTO, RESUSINEFECTO, TERRENO, ZONA")
                End If
                
            Case "/GUARDAR", "/GRABAR", "/G"
                Call WriteSaveChars
                
            Case "/BORRAR"
                If notNullArguments Then
                    Select Case UCase$(ArgumentosAll(0))
                        Case "SOS" '"/BORRAR SOS"
                            Call WriteCleanSOS
                            
                    End Select
                End If
                

            Case "/TIEMPO"
                If notNullArguments Then
                    Call WriteWeather(ArgumentosAll(0))
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /TIEMPO X, siendo X el estado de tiempo.")
                End If
                
            Case "/ETODOSPJS"
                Call WriteKickAllChars
                
            Case "/RELOADNPCS"
                Call WriteReloadNPCs
                
            Case "/RELOADSINI"
                Call WriteReloadServidorIni
                
            Case "/RELOADHECHIZOS"
                Call WriteReloadSpells
                
            Case "/RELOADOBJ"
                Call WriteReloadObjects
                 
            Case "/REINICIAR"
                Call WriteRestart
                
            Case "/CHATCOLOR"
                If notNullArguments And CantidadArgumentos >= 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        Call WriteChatColor(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /Chatcolor R G B.")
                    End If
                    
                ElseIf Not notNullArguments Then    'Go back to default!
                    Call WriteChatColor(0, 255, 0)
                    
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /Chatcolor R G B.")
                End If
            
            Case "/IGNORADO"
                Call WriteIgnored
            
            Case "/P", "/PING"
                Call WritePing
                
            Case "/SETINIVAR"
                If CantidadArgumentos = 3 Then
                    ArgumentosAll(2) = Replace(ArgumentosAll(2), "+", " ")
                    Call WriteSetIniVar(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                Else
                    Call ShowConsoleMsg("Prámetros incorrectos. Utilice /SETINIVAR LLAVE CLAVE VALOR")
                End If

        End Select
        
    'Mensaje a Companiero
    ElseIf Left$(Comando, 1) = ":" Then
    
        If UserMuerto Then 'Muerto
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
            End With
            Exit Sub
        End If
    
        If Len(Comando) < 5 Then
            Call UserChat(RawCommand, eChatType.Norm)
            
        ElseIf notNullArguments Then
                    
            Name = mid$(Comando, 2)
                  
            If LenB(Name) < 3 Then
                Exit Sub
            End If
    
            If InStrB(Name, "+") > 0 Then
                Name = Replace$(Name, "+", " ")
            End If
            
            Name = StrConv(Name, vbProperCase)
            
            If Name <> UserName Then
            
                Dim Slot As Byte
                
                Slot = EsCompaniero(Name)
                
                If Slot > 0 Then
                
                    If Compa(Slot).Online Then
                        RawCommand = Trim$(mid$(RawCommand, Len(":" & Name & " ")))
                                            
                        If LenB(RawCommand) > 0 Then
                            Call UserChat(RawCommand, eChatType.Komp, Slot)
                            
                            LastParsedString = ":" & Name & " "
                        End If
                    Else
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg(Name & " no está.", .Red, .Green, .Blue, .Bold, .Italic)
                        End With
                        Exit Sub
                    End If
                Else
                    'For i = 1 To LastChar
                    'If Not Charlist(i).EsUser Then
                    'If Not Charlist(i).Invisible Then
                    'If LenB(Charlist(i).Nombre) = LenB(Name) Then
                    'If Charlist(i).Nombre = Name Then
                                                                               
                    RawCommand = Trim$(mid$(RawCommand, Len(":" & Name & " ")))

                    If LenB(RawCommand) > 0 Then
                        If (InStrB(Name, " ") > 0) Then
                            Name = Replace$(Name, " ", "+")
                        End If
         
                        If Len(Name) > 2 Then
                            Call UserChat(RawCommand, eChatType.Priv, , Name)
                            LastParsedString = ":" & Name & " "
                        End If
                    End If
                    
                    'Exit For
                    'End If
                    'End If
                    'End If
                    'End If
                    'Next i
                End If
            End If
        End If
                
    ElseIf Left$(Comando, 1) = "*" Then
    
        If UserMuerto Then 'Muerto
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
            End With
            Exit Sub
        End If
    
        Mensaje = Trim$(mid$(RawCommand, 2))
        
        If LenB(Mensaje) < 1 Then
            Exit Sub
        End If
        
        If UserLvl < 15 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("No podés comunicarte en modo global siendo principiante.", .Red, .Green, .Blue, .Bold, .Italic)
            End With
            Exit Sub
        End If
        
        Call UserChat(Mensaje, eChatType.Glob)

    ElseIf Left$(Comando, 1) = "." Then

        If Comando = ".." Or Comando = "..." Then
            Call UserChat(Comando, eChatType.Norm)
        Else
            Mensaje = Trim$(mid$(RawCommand, 2))
            
            If LenB(Mensaje) < 1 Then
                Exit Sub
            End If

            Call UserChat(Mensaje, eChatType.Guil)
        End If
        
        
    ElseIf Left$(Comando, 1) = "+" Then
            
        If UserMuerto Then 'Muerto
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
            End With
            Exit Sub
        End If
        
        Name = mid$(RawCommand, 2)
                                
        If Name = UserName Then
            Exit Sub
        End If
        
        If EsCompaniero(Name) > 0 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg(Name & " ya es tu compañero/a.", .Red, .Green, .Blue, .Bold, .Italic)
            End With
            Exit Sub
        End If
        
        If Not AsciiValidos(Name) Then
            Exit Sub
        End If
        
        If Len(Name) < 3 Then
            Exit Sub
        End If

        Call WriteAniadirCompaniero(Name)
        
        LastParsedString = vbNullString
        
    ElseIf Left$(Comando, 1) = "-" Then
    
        If UserMuerto Then 'Muerto
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
            End With
            Exit Sub
        End If
                
        Dim k As Byte
        
        Name = StrConv(mid$(RawCommand, 2), vbProperCase)
                
        If Name = UserName Then
            Exit Sub
        End If
        
        Dim Nro As Byte
        
        Nro = EsCompaniero(Name)
        
        If Nro < 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg(Name & " no es tu compañero/a.", .Red, .Green, .Blue, .Bold, .Italic)
            End With
            Exit Sub
        End If

        Call WriteEliminarCompaniero(Nro)
        
        LastParsedString = vbNullString

    ElseIf Left$(Comando, 1) = ">" Then

        'Ojo, no usar notNullArguments porque se usa el string vacio para borrar cartel.
        If CantidadArgumentos > 0 Then
            Call WritePartyMessage(ArgumentosRaw)
        End If
        
    Else
        If UserMuerto Then 'Muerto
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
            End With
            Exit Sub
        End If
        
        Call UserChat(RawCommand, eChatType.Norm)
    End If
End Sub

Public Function ValidNumber(ByVal Numero As String, ByVal TIPO As eNumber_Types) As Boolean
'Returns whether the number is correct.
    Dim Minimo As Long
    Dim Maximo As Long
    
    If Not IsNumeric(Numero) Then
        Exit Function
    End If
    
    Select Case TIPO
        Case eNumber_Types.ent_Byte
            Minimo = 0
            Maximo = 255

        Case eNumber_Types.ent_Integer
            Minimo = -32768
            Maximo = 32767

        Case eNumber_Types.ent_Long
            Minimo = -2147483648#
            Maximo = 2147483647
        
        Case eNumber_Types.ent_Trigger
            Minimo = 0
            Maximo = 6
    End Select
    
    If Val(Numero) >= Minimo And Val(Numero) <= Maximo Then ValidNumber = True
End Function

Private Function validipv4str(ByVal Ip As String) As Boolean
'Returns whether the ip format is correct.
    Dim tmpArr() As String
    
    tmpArr = Split(Ip, ".")
    
    If UBound(tmpArr) <> 3 Then _
        Exit Function

    If Not ValidNumber(tmpArr(0), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(1), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(2), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(3), eNumber_Types.ent_Byte) Then _
        Exit Function
    
    validipv4str = True
End Function

Private Function str2ipv4l(ByVal Ip As String) As Byte()
'Converts a string into the correct ip format.
    Dim tmpArr() As String
    Dim bArr(3) As Byte
    
    tmpArr = Split(Ip, ".")
    
    bArr(0) = CByte(tmpArr(0))
    bArr(1) = CByte(tmpArr(1))
    bArr(2) = CByte(tmpArr(2))
    bArr(3) = CByte(tmpArr(3))

    str2ipv4l = bArr
End Function

Private Function AEMAILSplit(ByRef Text As String) As String()
'Do an Split() in the /AEMAIL in onother way
    Dim tmpArr(0 To 1) As String
    Dim Pos As Byte
    
    Pos = InStr(1, Text, "-")
    
    If Pos > 0 Then
        tmpArr(0) = mid$(Text, 1, Pos - 1)
        tmpArr(1) = mid$(Text, Pos + 1)
    Else
        tmpArr(0) = vbNullString
    End If
    
    AEMAILSplit = tmpArr
End Function
