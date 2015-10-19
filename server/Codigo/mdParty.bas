Attribute VB_Name = "mdParty"
'mdParty.bas - Library of functions to manipulate parties.

Option Explicit

'SOPORTES PARA LAS PARTIES
'(Ver este modulo como una clase abstracta "PartyManager")

'Cantidad Maxima de parties en el servidor
Public Const Max_PARTIES As Integer = 300

'nivel minimo para crear party
Public Const MINPARTYLEVEL As Byte = 15

'Cantidad Maxima de gente en la party
Public Const PARTY_MaxMEMBERS As Byte = 5

'Si esto esta en True, la exp sale por cada golpe que le da
'Si no, la exp la recibe al salirse de la party (pq las partys, floodean)
Public Const PARTY_EXPERIENCIAPORGOLPE As Boolean = True

'Maxima diferencia de niveles permitida en una party
Public Const MaxPARTYDELTALEVEL As Byte = 7

'distancia al Leader para que este acepte el ingreso
Public Const MaxDISTANCIAINGRESOPARTY As Byte = 2

'Maxima distancia a un exito para obtener su experiencia
Public Const PARTY_MaxDISTANCIA As Byte = 18

'restan las muertes de los miembros?
Public Const CASTIGOS As Boolean = False

'Numero al que elevamos el nivel de cada miembro de la party
'Esto es usado para calcular la distribución de la experiencia entre los miembros
'Se lee del archivo de balance
Public ExponenteNivelParty As Single

Public Type tPartyMember
    UserIndex As Integer
    Experiencia As Double
End Type

Public Function NextParty() As Integer
    Dim i As Integer
    
    NextParty = -1
    
    For i = 1 To Max_PARTIES
        If Parties(i) Is Nothing Then
            NextParty = i
            Exit Function
        End If
    Next i
End Function

Public Function PuedeCrearParty(ByVal UserIndex As Integer) As Boolean
    With UserList(UserIndex)
        PuedeCrearParty = True
    'If .Stats.ELV < MINPARTYLEVEL Then
        
        If CInt(.Stats.Atributos(eAtributos.Carisma)) * .Skills.Skill(eSkill.Liderazgo).Elv < 100 Then
            Call WriteConsoleMsg(UserIndex, "Con tu nivel de Carisma (" & .Stats.Atributos(eAtributos.Carisma) & ") necesitás " & 100 \ .Stats.Atributos(eAtributos.Carisma) & " puntos de habilidad en Liderazgo para crear ", FontTypeNames.FONTTYPE_PARTY)
            PuedeCrearParty = False
        ElseIf .Stats.Muerto Then
            Call WriteConsoleMsg(UserIndex, "Estás muerto.", FontTypeNames.FONTTYPE_PARTY)
            PuedeCrearParty = False
        ElseIf .PartyIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Ya perteneces a una party.", FontTypeNames.FONTTYPE_PARTY)
            PuedeCrearParty = False
        End If
    End With
End Function

Public Sub CrearParty(ByVal UserIndex As Integer)
    Dim tInt As Integer

    With UserList(UserIndex)

        tInt = mdParty.NextParty
        If tInt = -1 Then
            Call WriteConsoleMsg(UserIndex, "Por el momento no se pueden crear mas parties", FontTypeNames.FONTTYPE_PARTY)
            Exit Sub
        Else
            Set Parties(tInt) = New clsParty
            If Not Parties(tInt).NuevoMiembro(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "La party está llena, no podés entrar", FontTypeNames.FONTTYPE_PARTY)
                Set Parties(tInt) = Nothing
                Exit Sub
            Else
                Call WriteConsoleMsg(UserIndex, "¡Has formado una party!", FontTypeNames.FONTTYPE_PARTY)
                    .PartyIndex = tInt
                    .PartySolicitud = 0
                If Not Parties(tInt).HacerLeader(UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "No podés hacerte líder.", FontTypeNames.FONTTYPE_PARTY)
                Else
                    Call WriteConsoleMsg(UserIndex, "¡ Te convertiste en líder de la party !", FontTypeNames.FONTTYPE_PARTY)
                End If
            End If
        End If
    End With
End Sub

Public Sub SolicitarIngresoAParty(ByVal UserIndex As Integer)
'ESTO ES enviado por el PJ para solicitar el ingreso a la party
    
    Dim tInt As Integer

    With UserList(UserIndex)
        If .PartyIndex > 0 Then
            'si ya esta en una party
            Call WriteConsoleMsg(UserIndex, "Ya perteneces a una party, escribe /SALIRPARTY para abandonarla", FontTypeNames.FONTTYPE_PARTY)
                .PartySolicitud = 0
            Exit Sub
        End If
            If .Stats.Muerto Then
                Call WriteConsoleMsg(UserIndex, "Estás muerto.", FontTypeNames.FONTTYPE_INFO)
                .PartySolicitud = 0
                    Exit Sub
            End If
            tInt = .flags.TargetUser
        If tInt > 0 Then
            If UserList(tInt).PartyIndex > 0 Then
                    .PartySolicitud = UserList(tInt).PartyIndex
                Call WriteConsoleMsg(UserIndex, " El fundador decidirá si te acepta en la party", FontTypeNames.FONTTYPE_PARTY)
            Else
                Call WriteConsoleMsg(UserIndex, UserList(tInt).Name & " no es fundador de ninguna party.", FontTypeNames.FONTTYPE_INFO)
                    .PartySolicitud = 0
                Exit Sub
            End If
        Else
            Call WriteConsoleMsg(UserIndex, " Para ingresar a una party debes hacer click sobre el fundador y luego escribir /pARTY", FontTypeNames.FONTTYPE_PARTY)
                .PartySolicitud = 0
        End If
    End With

End Sub

Public Sub SalirDeParty(ByVal UserIndex As Integer)
    Dim PI As Integer
    PI = UserList(UserIndex).PartyIndex
    If PI > 0 Then
        If Parties(PI).SaleMiembro(UserIndex) Then
            'sale el Leader
            Set Parties(PI) = Nothing
        Else
            UserList(UserIndex).PartyIndex = 0
        End If
    Else
        Call WriteConsoleMsg(UserIndex, " No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Public Sub ExpulsarDeParty(ByVal Leader As Integer, ByVal OldMember As Integer)
    Dim PI As Integer
    
    PI = UserList(Leader).PartyIndex
    
    If PI = UserList(OldMember).PartyIndex Then
        If Parties(PI).SaleMiembro(OldMember) Then
            'si la funcion me da true, entonces la party se disolvio
            'y los partyIndex fueron reseteados a 0
            Set Parties(PI) = Nothing
        Else
            UserList(OldMember).PartyIndex = 0
        End If
    Else
        Call WriteConsoleMsg(Leader, LCase(UserList(OldMember).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
    End If

End Sub

Public Function UserPuedeEjecutarComandos(ByVal User As Integer) As Boolean
'Determines if a user can use party commands like /acceptparty or not.

    Dim PI As Integer
    
    PI = UserList(User).PartyIndex
    
    If PI > 0 Then
        If Parties(PI).EsPartyLeader(User) Then
            UserPuedeEjecutarComandos = True
        Else
            Call WriteConsoleMsg(User, "¡No eres el líder de tu Party!", FontTypeNames.FONTTYPE_PARTY)
            Exit Function
        End If
    Else
        Call WriteConsoleMsg(User, "No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    
End Function

Public Sub AprobarIngresoAParty(ByVal Leader As Integer, ByVal NewMember As Integer)
    
    'el UI es el Leader
    Dim PI As Integer
    Dim razon As String
    
    PI = UserList(Leader).PartyIndex
    
    If UserList(NewMember).PartySolicitud = PI Then
        If Not UserList(NewMember).Stats.Muerto Then
            If UserList(NewMember).PartyIndex = 0 Then
                If Parties(PI).PuedeEntrar(NewMember, razon) Then
                    If Parties(PI).NuevoMiembro(NewMember) Then
                        Call Parties(PI).MandarMensajeAConsola(UserList(Leader).Name & " ha aceptado a " & UserList(NewMember).Name & " en la party.", "Servidor")
                        UserList(NewMember).PartyIndex = PI
                        UserList(NewMember).PartySolicitud = 0
                    Else
                        'no pudo entrar
                        'ACA UNO PUEDE CODIFICAR OTRO TIPO DE ERRORES...
                        Call SendData(SendTarget.ToAdmins, Leader, PrepareMessageConsoleMsg(" Servidor> CATASTROFE EN PARTIES, NUEVOMIEMBRO DIO FALSE! :S ", FontTypeNames.FONTTYPE_PARTY))
                        End If
                    Else
                    'no debe entrar
                    Call WriteConsoleMsg(Leader, razon, FontTypeNames.FONTTYPE_PARTY)
                End If
            Else
                Call WriteConsoleMsg(Leader, UserList(NewMember).Name & " ya es miembro de otra party.", FontTypeNames.FONTTYPE_PARTY)
                Exit Sub
            End If
        Else
            Call WriteConsoleMsg(Leader, "¡Está muerto, no podés aceptar miembros en ese estado!", FontTypeNames.FONTTYPE_PARTY)
            Exit Sub
        End If
    Else
        Call WriteConsoleMsg(Leader, LCase(UserList(NewMember).Name) & " no ha solicitado ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
        Exit Sub
    End If

End Sub

Public Sub BroadCastParty(ByVal UserIndex As Integer, ByRef texto As String)
    Dim PI As Integer
    
    PI = UserList(UserIndex).PartyIndex
    
    If PI > 0 Then
        Call Parties(PI).MandarMensajeAConsola(texto, UserList(UserIndex).Name)
    End If
End Sub

Public Sub OnlineParty(ByVal UserIndex As Integer)

    Dim i As Integer
    Dim PI As Integer
    Dim Text As String
    Dim MembersOnline(1 To PARTY_MaxMEMBERS) As Integer
    
    PI = UserList(UserIndex).PartyIndex
    
    If PI > 0 Then
        Call Parties(PI).ObtenerMiembrosOnline(MembersOnline)
        Text = "Nombre(Exp): "
        For i = 1 To PARTY_MaxMEMBERS
            If MembersOnline(i) > 0 Then
                Text = Text & " - " & UserList(MembersOnline(i)).Name & " (" & Fix(Parties(PI).MiExperiencia(MembersOnline(i))) & ")"
            End If
        Next i
        Text = Text & ". Experiencia total: " & Parties(PI).ObtenerExperienciaTotal
        Call WriteConsoleMsg(UserIndex, Text, FontTypeNames.FONTTYPE_PARTY)
    End If

End Sub

Public Sub TransformarEnLider(ByVal OldLeader As Integer, ByVal NewLeader As Integer)
    
    Dim PI As Integer
    
    If OldLeader = NewLeader Then
        Exit Sub
    End If
    
    PI = UserList(OldLeader).PartyIndex
    
    If PI = UserList(NewLeader).PartyIndex Then
        If Not UserList(NewLeader).Stats.Muerto Then
            If Parties(PI).HacerLeader(NewLeader) Then
                Call Parties(PI).MandarMensajeAConsola("El nuevo líder de la party es " & UserList(NewLeader).Name, UserList(OldLeader).Name)
            Else
                Call WriteConsoleMsg(OldLeader, "¡No se ha hecho el cambio de mando!", FontTypeNames.FONTTYPE_PARTY)
            End If
        Else
            Call WriteConsoleMsg(OldLeader, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(OldLeader, LCase(UserList(NewLeader).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
    End If
    
End Sub

Public Sub ActualizaExperiencias()
'esta funcion se invoca antes de worlsaves, y apagar servidores
'en caso que la experiencia sea acumulada y no por golpe
'para que grabe los datos en los charfiles
    Dim i As Integer
    
    If Not PARTY_EXPERIENCIAPORGOLPE Then
        haciendoBK = True
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        
        'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Distribuyendo experiencia en parties.", FontTypeNames.FONTTYPE_SERVER))
        For i = 1 To Max_PARTIES
            If Not Parties(i) Is Nothing Then
                Call Parties(i).FlushExperiencia
            End If
        Next i
        'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Experiencia distribuida.", FontTypeNames.FONTTYPE_SERVER))
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        haciendoBK = False
    End If

End Sub

Public Sub ObtenerExito(ByVal UserIndex As Integer, ByVal Exp As Long, mapa As Integer, X As Integer, Y As Integer)
    If Exp < 1 Then
        If Not CASTIGOS Then
            Exit Sub
        End If
    End If
    
    Call Parties(UserList(UserIndex).PartyIndex).ObtenerExito(Exp, mapa, X, Y)
End Sub

Public Function CantMiembros(ByVal UserIndex As Integer) As Integer
    CantMiembros = 0
    If UserList(UserIndex).PartyIndex > 0 Then
        CantMiembros = Parties(UserList(UserIndex).PartyIndex).CantMiembros
    End If
End Function

Public Sub ActualizarSumaNivelesElevados(ByVal UserIndex As Integer)
'Sets the new p_sumaniveleselevados to the party.
'When a user level up and he is in a party, we call this PUBLIC SUB to don't desestabilice the party exp formula
  
  If UserList(UserIndex).PartyIndex > 0 Then
        Call Parties(UserList(UserIndex).PartyIndex).UpdateSumaNivelesElevados(UserList(UserIndex).Stats.Elv)
    End If
End Sub
