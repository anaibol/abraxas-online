Attribute VB_Name = "Admin"
Option Explicit

Public Type tMotd
    texto As String
    Formato As String
End Type

Public MaxLines As Integer
Public Motd() As tMotd

Public Type tAPuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type

Public Apuestas As tAPuestas

Public tInicioServer As Long

'INTERVALOS
Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public IntervaloHambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizado As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloLanzaHechizo As Integer
Public IntervaloNpcPuedeAtacar As Integer
Public IntervaloNpcAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloOculto As Integer
Public IntervaloUserPuedeAtacar As Long
Public IntervaloGolpeUsar As Long
Public IntervaloMagiaGolpe As Long
Public IntervaloGolpeMagia As Long
Public IntervaloUserPuedeCastear As Long
Public IntervaloUserPuedeTrabajar As Long
Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long
Public IntervaloUserPuedeUsar As Long
Public IntervaloFlechasCazadores As Long
Public IntervaloPuedeSerAtacado As Long
Public IntervaloAtacable As Long
Public IntervaloOwnedNpc As Long

'MULTIPLICADORES
Public MultiplicadorExp As Integer
Public MultiplicadorGld As Integer

'BALANCE
'Public PorcentajeRecuperoMana As Integer

Public MinutosWs As Long
Public Puerto As Integer

Public BootDelBackUp As Byte
Public Lloviendo As Boolean

Public Sub WorldSave()
    On Error Resume Next
    
    Dim loopX As Integer
    Dim Porc As Long
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Iniciando WorldSave", FontTypeNames.FONTTYPE_SERVER))
        
    'Dim j As Integer , k As Integer
    
    'For j = 1 To NumMaps
    '    If MapInfo(j).BackUp = 1 Then
    '        k = k + 1
    '    End If
    'Next j
    
    FrmStat.ProgressBar1.min = 0
    FrmStat.ProgressBar1.max = NumMaps 'k
    FrmStat.ProgressBar1.value = 0
    
    For loopX = 1 To NumMaps
        'DoEvents
        
        'If MapInfo(loopX).BackUp = 1 Then
            Call GrabarMapa(loopX, App.Path & "/WorldBackUp/Mapa" & loopX)
            FrmStat.ProgressBar1.value = FrmStat.ProgressBar1.value + 1
        'End If
    Next loopX
    
    FrmStat.Visible = False
    
    If FileExist(DatPath & "/bkNpc.dat", vbNormal) Then
        Kill (DatPath & "bkNpc.dat")
    End If
    
    'If FileExist(DatPath & "/bkNpcs-HOSTILES.dat", vbNormal) Then Kill (DatPath & "bkNpcs-HOSTILES.dat")
    
    For loopX = 1 To LastNpc
        If NpcList(loopX).flags.BackUp = 1 Then
            Call BackUPnPc(loopX)
        End If
    Next
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> WorldSave ha concluído.", FontTypeNames.FONTTYPE_SERVER))

End Sub

Public Sub PurgarPenas()
    Dim i As Long
    
    For i = 1 To LastUser
        If UserList(i).flags.Logged Then
            If UserList(i).Counters.Pena > 0 Then
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                If UserList(i).Counters.Pena < 1 Then
                    Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                    Call WriteConsoleMsg(i, "¡Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
                    
                    Call FlushBuffer(i)
                End If
            End If
            
            If UserList(i).Counters.Silencio > 0 Then
                UserList(i).Counters.Silencio = UserList(i).Counters.Silencio - 1
                
                If UserList(i).Counters.Silencio < 1 Then
                    Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                    Call WriteShowMessageBox(i, "El efecto del silencio ha desaparecido.")
                    Call FlushBuffer(i)
                End If
            End If
        End If
    Next i
End Sub

Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = vbNullString)
        
    UserList(UserIndex).Counters.Pena = Minutos
    
    Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
    
    If LenB(GmName) < 1 Then
        Call WriteConsoleMsg(UserIndex, "Has sido encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, GmName & " te encarceló, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    End If
     
End Sub

Public Sub BorrarUsuario(ByVal UserName As String)
    If User_Exist(UserName) Then
        Kill CharPath & UCase$(UserName) & ".chr"
    End If
End Sub

Public Sub KillCharInfo(ByVal UserName As String)

On Error Resume Next
     
    Dim c As String
    Dim d As String
    Dim f As String
    Dim g As String
    Dim H As Byte
    Dim i As String
    Dim j As String
     
     
    c = GetVar(CharPath & UserName & ".chr", "GUILD", "Guild_Id")
    d = GetVar(GUILDINFOFILE, "GUILD" & c, "Founder")
    f = GetVar(GUILDINFOFILE, "GUILD" & c, "GuildName")
    g = GetVar(GUILDPATH & f & "-members.mem", "INIT", "NroMembers")
    j = GetVar(GUILDPATH & f & "-members.mem", "Members", "Member" & g)
        
    If c = vbNullString Then
        Kill (CharPath & UserName & ".chr")
    Else
        If d <> UserName Then
            Guilds(c).ExpulsarMiembro (UserName)
        Else
            For H = 1 To g
               i = GetVar(GUILDPATH & f & "-members.mem", "Members", "Member" & H)
    
               If i = UserName Then
                   Call WriteVar(GUILDPATH & f & "-members.mem", "Members", "Member" & H, j): Call WriteVar(GUILDPATH & f & "-members.mem", "INIT", "NroMembers", g - 1)
               End If
                   
               Call WriteVar(GUILDINFOFILE, "GUILD" & c, "EleccionesAbiertas", "1")
               Call WriteVar(GUILDINFOFILE, "GUILD" & c, "EleccionesFinalizan", DateAdd("d", 1, Now))
               Call WriteVar(GUILDPATH & f & "-votaciones.vot", "INIT", "NumVotos", "0")
           Next H
              
       End If
    
    End If
     
End Sub

Public Function UnBan(ByVal Name As String) As Boolean
    'Unban the Char
    Call WriteVar(CharPath & Name & ".chr", "FLAGS", "Ban", "0")
    
    'Remove it from the banned people database
    Call WriteVar(App.Path & "/logs/" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
    Call WriteVar(App.Path & "/logs/" & "BanDetail.dat", Name, "Reason", "NO REASON")
End Function

Public Sub BanIpAgrega(ByVal Ip As String)
    BanIps.Add Ip
    
    Call BanIpGuardar
End Sub

Public Function BanIpBuscar(ByVal Ip As String) As Long
    Dim Dale As Boolean
    Dim LoopC As Long
    
    Dale = True
    LoopC = 1
    Do While LoopC <= BanIps.Count And Dale
        Dale = (BanIps.Item(LoopC) <> Ip)
        LoopC = LoopC + 1
    Loop
    
    If Dale Then
        BanIpBuscar = 0
    Else
        BanIpBuscar = LoopC - 1
    End If
End Function

Public Function BanIpQuita(ByVal Ip As String) As Boolean

    On Error Resume Next
    
    Dim N As Long
    
    N = BanIpBuscar(Ip)
    If N > 0 Then
        BanIps.Remove N
        BanIpGuardar
        BanIpQuita = True
    Else
        BanIpQuita = False
    End If

End Function

Public Sub BanIpGuardar()
    Dim ArchivoBanIp As String
    Dim ArchN As Long
    Dim LoopC As Long
    
    ArchivoBanIp = DatPath & "banIps.dat"
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Output As #ArchN
    
    For LoopC = 1 To BanIps.Count
        Print #ArchN, BanIps.Item(LoopC)
    Next LoopC
    
    Close #ArchN
End Sub

Public Sub BanIpCargar()
    Dim ArchN As Long
    Dim Tmp As String
    Dim ArchivoBanIp As String
    
    ArchivoBanIp = DatPath & "banIps.dat"
    
    Do While BanIps.Count > 0
        BanIps.Remove 1
    Loop
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Input As #ArchN
    
    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        BanIps.Add Tmp
    Loop
    
    Close #ArchN

End Sub

Public Function UserDarPrivilegioLevel(ByVal Name As String) As PlayerType
    If EsAdmin(Name) Then
        UserDarPrivilegioLevel = PlayerType.Admin
    ElseIf EsDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.Dios
    ElseIf EsSemiDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.SemiDios
    ElseIf EsConsejero(Name) Then
        UserDarPrivilegioLevel = PlayerType.Consejero
    Else
        UserDarPrivilegioLevel = PlayerType.User
    End If
End Function

Public Sub BanCharacter(ByVal bannerUserIndex As Integer, ByVal UserName As String, ByVal Reason As String)

    Dim tUser As Integer
    Dim userPriv As Byte
    Dim cantPenas As Byte
    Dim rank As Integer
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")
    End If
    
    tUser = NameIndex(UserName)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)
        If tUser < 1 Then
            Call WriteConsoleMsg(bannerUserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_TALK)
            
            If User_Exist(UserName) Then
                userPriv = UserDarPrivilegioLevel(UserName)
                
                If (userPriv And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(bannerUserIndex, "No podés banear a al alguien de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban") <> "0" Then
                        Call WriteConsoleMsg(bannerUserIndex, "El personaje ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call LogBanFromName(UserName, bannerUserIndex, Reason)
                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserName & ".", FontTypeNames.FONTTYPE_SERVER))
                        
                        'ponemos el flag de ban a 1
                        Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                        'ponemos la pena
                        cantPenas = Val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(Reason) & " " & Date & " " & Time)
                        
                        If (userPriv And rank) = (.flags.Privilegios And rank) Then
                            .flags.Ban = 1
                            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                            Call CloseSocket(bannerUserIndex)
                        End If
                        
                        Call LogGM(.Name, "BAN a " & UserName)
                    End If
                End If
            Else
                Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                Call WriteConsoleMsg(bannerUserIndex, "No podés banear a al alguien de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            Call LogBan(tUser, bannerUserIndex, Reason)
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_SERVER))
            
            'Ponemos el flag de ban a 1
            UserList(tUser).flags.Ban = 1
            
            If (UserList(tUser).flags.Privilegios And rank) = (.flags.Privilegios And rank) Then
                .flags.Ban = 1
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                Call CloseSocket(bannerUserIndex)
            End If
            
            Call LogGM(.Name, "BAN a " & UserName)
            
            'ponemos el flag de ban a 1
            Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
            'ponemos la pena
            cantPenas = Val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(Reason) & " " & Date & " " & Time)
            
            Call CloseSocket(tUser)
        End If
    End With
End Sub

