Attribute VB_Name = "ModGeneral"
Option Explicit

'Private FontSmooth As Boolean

Private Const RunHighPriority As Boolean = False

Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Declare Function BitBlt Lib "gdi32" _
    (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, _
     ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, _
     ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Integer, ByVal uParam As Integer, ByVal lpvParam As Integer, ByVal fuWinIni As Integer) As Integer

Const SPI_GETFONTSMOOTHING As Integer = &H4A
Const SPI_SETFONTSMOOTHING As Integer = &H4B

Const SPI_GETFONTSMOOTHINGTYPE As Integer = &H200A
Const SPI_SETFONTSMOOTHINGTYPE As Integer = &H200B

Const ClearType As Integer = &H2
Const StandardType As Integer = &H1
Const NoType As Integer = &H0

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Public bPortal As Boolean

Public bFogata As Boolean

Public bLluvia() As Byte 'Array para determinar si
'debemos mostrar la animacion de la lluvia

Private lFrameTimer As Long


'FUNCIONES API

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Para ejecutar el browser y programas externos
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)

Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Property Get TitleBarHeight() As Long
    TitleBarHeight = GetSystemMetrics(4)
End Property

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Sub CargarAnimArmas()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = DataPath & "Armas.dat"
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc
End Sub

Public Sub CargarMiniMapa()

End Sub

Public Sub CargarAnimEscudos()

On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = DataPath & "Escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
    Next loopc
End Sub

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

Public Sub UnloadAllForms()
On Error Resume Next
    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Public Function LegalChar(ByVal KeyAscii As Integer) As Boolean
'Only allow Chars that are Win 95 filename compatible

    'if backspace allow
    If KeyAscii = 8 Then
        LegalChar = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special Chars
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special Chars in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalChar = True
End Function

Public Sub SetConnected()
'Sets the client to "Connect" mode

On Error Resume Next

    'Variable initialization
    EngineRun = True
    Nombres = True

    WriteVar DataPath & "Game.ini", "INIT", "Name", UserName
    
    If frmConnect.SavePassImg.Visible Then
        WriteVar DataPath & "Game.ini", "INIT", "Pass", UserPassword
    Else
        WriteVar DataPath & "Game.ini", "INIT", "Pass", vbNullString
    End If
            
    'Guardar los valores del array en el archivo de historial
    'Call saveValues
        
    'Unload the connect form
    Unload frmCrearPersonaje
    Unload frmConnect
    
    If Not ChangeResolution And Not ResolucionActual Then
        frmMain.BorderStyle = 3
        frmMain.Caption = frmMain.Caption
    End If
                      
    Call SetMusicInfo("Jugando Abraxas | Nombre: " & UserName & " | Jugadores en línea: " & frmMain.LblPoblacion.Caption, "abraxas-online.com", "Games", "{1}{0}")

    Call Audio.MusicMP3Stop

    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg("Bienvenido a ", .Red, .Green, .Blue, .Bold, .Italic, True)
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_NUMERO)
        Call ShowConsoleMsg("Abraxas", .Red, .Green, .Blue, .Bold, .Italic, True)
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg(".", .Red, .Green, .Blue, .Bold, .Italic)
    End With
    
    'Load main form
    frmMain.Visible = True
    
    Call SetWindowPos(frmMain.hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
        
    UserLogged = True
End Sub

Public Sub MoveTo(ByVal Direccion As eHeading)

    Dim LegalOk As Boolean
    
    If Cartel Then
        Cartel = False
    End If
    
    LegalOk = MoveToLegalPos(Direccion)

    If LegalOk And Not UserParalizado Then
        If Meditando Then
            Meditando = False
            Charlist(UserCharIndex).FxIndex = 0
            
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("Dejás de meditar.", .Red, .Green, .Blue, .Bold, .Italic)
            End With
        End If
        
        If Descansando Then
            Descansando = False
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("Dejás de descansar.", .Red, .Green, .Blue, .Bold, .Italic)
            End With
        End If
        
        Call WriteWalk(Direccion)
        
        Call MoveCharbyHead(Direccion)
  
    Else
        If Charlist(UserCharIndex).Heading <> Direccion Then
            Call WriteChangeHeading(Direccion)
            
            If Not UserParalizado Then
                Charlist(UserCharIndex).Heading = Direccion
            End If
        End If
    End If
    
    'If frmMain.MacroTrabajo.Enabled Then
    'frmMain.DesactivarMacroTrabajo
    'End If
        
    End Sub

Private Sub CheckKeys()
'Checks keys and respond

    Static lastMovement As Long
    
    'No input allowed while Argentum is not the active window
    If Not modApplication.IsAppActive() Then
        Exit Sub
    End If
    
    'No walking when in commerce or banking.
    If Comerciando Then
        Exit Sub
    End If
    
    'Don't allow any these keys during movement..
    If UserMuerto Then
        Exit Sub
    End If
    
    'Don't allow any these keys during movement..
    If UserMoving Then
        Exit Sub
    End If
    
    'If game is paused, abort movement.
    If Pausa Then
        Exit Sub
    End If
    
    'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
    If GetTickCount - lastMovement > 56 Then
        lastMovement = GetTickCount
    Else
        Exit Sub
    End If
    
    'Meditate
    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyMeditate)) < 0 Then
        If Not Meditando Then
            If Not MainTimer.Check(TimersIndex.Medit) Then
                Exit Sub
            End If
            
            Call DoMeditar
        End If
    Else
        Meditando = False
        Charlist(UserCharIndex).FxIndex = 0
        
        Call Audio.mSound_StopWav(PortalBufferIndex)
        PortalBufferIndex = 0
    End If

    If Not UserEstupido Then
        'Move Up
        If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
            If frmMain.TrainingMacro.Enabled Then
                frmMain.DesactivarMacroHechizos
            End If
            Call MoveTo(NORTH)
            Exit Sub
        End If
        
        'Move Right
        If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
            If frmMain.TrainingMacro.Enabled Then
                frmMain.DesactivarMacroHechizos
            End If
            Call MoveTo(EAST)
            Exit Sub
        End If
    
        'Move down
        If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
            If frmMain.TrainingMacro.Enabled Then
                frmMain.DesactivarMacroHechizos
            End If
            Call MoveTo(SOUTH)
            Exit Sub
        End If
    
        'Move left
        If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
            If frmMain.TrainingMacro.Enabled Then
                frmMain.DesactivarMacroHechizos
            End If
            Call MoveTo(WEST)
            Exit Sub
        End If
    
    Else
        Dim kp As Boolean
        kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
            GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
            GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
            GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
        
        If kp Then
            Call MoveTo(RandomNumber(NORTH, WEST))
        End If
        
        If frmMain.TrainingMacro.Enabled Then
            frmMain.DesactivarMacroHechizos
        End If
    End If

End Sub
Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'Gets a field from a delimited string

    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'Gets the number of fields in a delimited string

    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos > 0
    
    FieldCount = Count
End Function

Function FileExist(ByVal file As String, ByVal FileObjType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileObjType) <> vbNullString)
End Function

Public Sub Main()

On Error Resume Next

    PacketName(ServerPacketID.SpawnList) = "SpawnList"  'SPL
    PacketName(ServerPacketID.ShowSOSForm) = "ShowSOSForm" 'MSOS
    PacketName(ServerPacketID.ShowMOTDEditionForm) = "ShowMOTDEditionForm" 'ZMOTD
    PacketName(ServerPacketID.ShowGMPanelForm) = "ShowGMPanelForm" 'ABPANEL
    PacketName(ServerPacketID.UserNameList) = "UserNameList" 'LISTUSU
    
    PacketName(ServerPacketID.Logged) = "Logged" 'LOGGED
    PacketName(ServerPacketID.RandomName) = "RandomName"
    PacketName(ServerPacketID.NavigateToggle) = "NavigateToggle" 'NAVEG
    PacketName(ServerPacketID.UserCommerceInit) = "UserCommerceInit" 'INITCOMUSU
    PacketName(ServerPacketID.UserCommerceEnd) = "UserCommerceEnd" 'FINCOMUSUOK
    PacketName(ServerPacketID.UserOfferConfirm) = "UserOfferConfirm"
    PacketName(ServerPacketID.CommerceChat) = "CommerceChat"
    PacketName(ServerPacketID.CharSwing) = "CharSwing"
    PacketName(ServerPacketID.NPCKillUser) = "NPCKillUser"
    PacketName(ServerPacketID.BlockedWithShield) = "BlockedWithShield"
    PacketName(ServerPacketID.BlockedWithShieldOther) = "BlockedWithShieldOther"
    PacketName(ServerPacketID.UserSwing) = "UserSwing"
    PacketName(ServerPacketID.ResuscitationSafeOn) = "ResuscitationSafeOn"
    PacketName(ServerPacketID.ResuscitationSafeOff) = "ResuscitationSafeOff"
    PacketName(ServerPacketID.UpdateSta) = "UpdateSta"
    PacketName(ServerPacketID.UpdateMana) = "UpdateMana"
    PacketName(ServerPacketID.UpdateHP) = "UpdateHP"
    PacketName(ServerPacketID.UpdateGold) = "UpdateGold"
    PacketName(ServerPacketID.UpdateExp) = "UpdateExp"
    PacketName(ServerPacketID.ChangeMap) = "ChangeMap"
    PacketName(ServerPacketID.PosUpdate) = "PosUpdate"
    PacketName(ServerPacketID.Damage) = "Damage"
    PacketName(ServerPacketID.UserDamaged) = "UserDamaged"
    PacketName(ServerPacketID.ChatOverHead) = "ChatOverHead"
    PacketName(ServerPacketID.DeleteChatOverHead) = "DeleteChatOverHead"
    PacketName(ServerPacketID.ConsoleMsg) = "ConsoleMsg"
    PacketName(ServerPacketID.ChatNormal) = "ChatNormal"
    'PacketName(ServerPacketID.ChatGM) = "ChatGM"
    PacketName(ServerPacketID.ChatGuild) = "ChatGuild"
    PacketName(ServerPacketID.ChatCompa) = "ChatCompa"
    PacketName(ServerPacketID.ChatPrivate) = "ChatPrivate"
    PacketName(ServerPacketID.ChatGlobal) = "ChatGlobal"
    PacketName(ServerPacketID.ShowMessageBox) = "ShowMessageBox"
    PacketName(ServerPacketID.CharCreate) = "CharCreate"
    PacketName(ServerPacketID.NpcCharCreate) = "NpcCharCreate"
    PacketName(ServerPacketID.CharRemove) = "CharRemove"
    PacketName(ServerPacketID.CharChangeNick) = "CharChangeNick"
    PacketName(ServerPacketID.CharMove) = "CharMove"
    PacketName(ServerPacketID.ForceCharMove) = "ForceCharMove"
    PacketName(ServerPacketID.CharChange) = "CharChange"
    PacketName(ServerPacketID.ChangeCharHeading) = "ChangeCharHeading"
    PacketName(ServerPacketID.ObjCreate) = "ObjCreate"
    PacketName(ServerPacketID.ObjectDelete) = "ObjectDelete"
    PacketName(ServerPacketID.BlockPosition) = "BlockPosition"
    PacketName(ServerPacketID.PlayMP3) = "PlayMP3"
    PacketName(ServerPacketID.PlayWav) = "PlayWav"
    PacketName(ServerPacketID.GuildList) = "GuildList"
    PacketName(ServerPacketID.AreaChanged) = "AreaChanged"
    PacketName(ServerPacketID.PauseToggle) = "PauseToggle"
    PacketName(ServerPacketID.RainToggle) = "RainToggle"
    PacketName(ServerPacketID.Weather) = "Weather"
    PacketName(ServerPacketID.CreateFX) = "CreateFX"
    PacketName(ServerPacketID.CreateCharFX) = "CreateCharFX"
    PacketName(ServerPacketID.UpdateUserStats) = "UpdateUserStats"
    PacketName(ServerPacketID.SlotMenosUno) = "SlotMenosUno"
    PacketName(ServerPacketID.Inventory) = "Inventory"
    PacketName(ServerPacketID.BeltInv) = "BeltInv"
    PacketName(ServerPacketID.BankSlot) = "Bank"
    PacketName(ServerPacketID.NpcInventory) = "NpcInventory"
    PacketName(ServerPacketID.InventorySlot) = "InventorySlot"
    PacketName(ServerPacketID.BankSlot) = "BankSlot"
    PacketName(ServerPacketID.NpcInventorySlot) = "NpcInventorySlot"
    PacketName(ServerPacketID.Spells) = "Spells"
    PacketName(ServerPacketID.SpellSlot) = "SpellSlot"
    PacketName(ServerPacketID.Compas) = "Compas"
    PacketName(ServerPacketID.AddCompa) = "AddCompa"
    PacketName(ServerPacketID.QuitarCompa) = "QuitarCompa"
    PacketName(ServerPacketID.CompaConnected) = "CompaConnected"
    PacketName(ServerPacketID.CompaDisconnected) = "CompaDisconnected"
    PacketName(ServerPacketID.Attributes) = "Attributes"
    PacketName(ServerPacketID.UserPlatforms) = "UserPlatforms"
    PacketName(ServerPacketID.BlacksmithWeapons) = "BlacksmithWeapons"
    PacketName(ServerPacketID.BlacksmithArmors) = "BlacksmithArmors"
    PacketName(ServerPacketID.CarpenterObjects) = "CarpenterObjects"
    PacketName(ServerPacketID.RestOK) = "RestOK"
    PacketName(ServerPacketID.ErrorMsg) = "ErrorMsg"
    PacketName(ServerPacketID.Blind) = "Blind"
    PacketName(ServerPacketID.Dumb) = "Dumb"
    PacketName(ServerPacketID.ShowSignal) = "ShowSignal"
    PacketName(ServerPacketID.UpdateHungerAndThirst) = "UpdateHungerAndThirst"
    PacketName(ServerPacketID.MiniStats) = "MiniStats"
    PacketName(ServerPacketID.SkillUp) = "SkillUp"
    PacketName(ServerPacketID.LevelUp) = "LevelUp" 'SUNI
    PacketName(ServerPacketID.SetInvisible) = "SetInvisible" 'NOVER
    PacketName(ServerPacketID.SetParalized) = "SetParalized" 'NOVER
    PacketName(ServerPacketID.BlindNoMore) = "BlindNoMore" 'NSEGUE
    PacketName(ServerPacketID.DumbNoMore) = "DumbNoMore" 'NESTUP
    PacketName(ServerPacketID.Skills) = "Skills" 'SKILLS
    PacketName(ServerPacketID.FreeSkillPts) = "FreeSkillPts"
    PacketName(ServerPacketID.TrainerCreatureList) = "TrainerCreatureList"
    PacketName(ServerPacketID.GuildNews) = "GuildNews"  'GUILDNE
    PacketName(ServerPacketID.OfferDetails) = "OfferDetails" 'PEACEDE & ALLIEDE
    PacketName(ServerPacketID.AlianceProposalsList) = "AlianceProposalsList" 'ALLIEPR
    PacketName(ServerPacketID.PeaceProposalsList) = "PeaceProposalsList" 'PEACEPR
    PacketName(ServerPacketID.CharInfo) = "CharInfo"  'CHRINFO
    PacketName(ServerPacketID.GuildLeaderInfo) = "GuildLeaderInfo" 'LEADERI
    PacketName(ServerPacketID.GuildMemberInfo) = "GuildMemberInfo" 'LEADERI
    PacketName(ServerPacketID.GuildDetails) = "GuildDetails" 'GuildaDET
    PacketName(ServerPacketID.ShowGuildFundationForm) = "ShowGuildFundationForm" 'SHOWFUN
    PacketName(ServerPacketID.ShowUserRequest) = "ShowUserRequest " 'PETICIO
    PacketName(ServerPacketID.ChangeUserTradeSlot) = "ChangeUserTradeSlot"
    PacketName(ServerPacketID.Pong) = "Pong"
    PacketName(ServerPacketID.UpdateTagAndStatus) = "UpdateTagAndStatus"
    PacketName(ServerPacketID.Population) = "Population"
    PacketName(ServerPacketID.AnimAttack) = "AnimAttack"
    PacketName(ServerPacketID.CharMeditate) = "CharMeditate"
    PacketName(ServerPacketID.ShowPartyForm) = "ShowPartyForm"
    PacketName(ServerPacketID.UpdateStrenghtAndDexterity) = "UpdateStrenghtAndDexterity"
    PacketName(ServerPacketID.UpdateStrenght) = "UpdateStrenght"
    PacketName(ServerPacketID.UpdateDexterity) = "UpdateDexterity"
    PacketName(ServerPacketID.StopWorking) = "StopWorking"
    PacketName(ServerPacketID.CancelOfferItem) = "CancelOfferItem"

    'If FindPreviousInstance Then
    '    Call MsgBox("Abraxas ya está abierto!")
    '    End
    'End If

    If RunHighPriority Then
        SetThreadPriority GetCurrentThread, 2       'Reccomended you dont touch these values
        SetPriorityClass GetCurrentProcess, &H80    'unless you know what you're doing
    End If

    Shell "regsvr32 /s dx7vb.dll"
    Shell "regsvr32 /s %WinDir%\system32\dx7vb.dll"
    
    Dim PacketKeys() As String

    'Inicialización de variables globales
    prgRun = True
    Pausa = False
    EngineRun = False
    
    DataPath = App.path & "\Data\"
    GrhPath = App.path & "\Grh\"
    MapPath = App.path & "\Maps\"
    MusicPath = App.path & "\Music\"
    SfxPath = App.path & "\Sfx\"
        
    ServerIP = "127.0.0.1"
    ServerPort = 7666
    
    'Call LoadRessources
    
    'FontSmooth = True
    'MsgBox SystemParametersInfo(SPI_GETFONTSMOOTHING, 0&, b, 0&)
    
    Call SystemParametersInfo(SPI_SETFONTSMOOTHING, 1&, &H0, &H1 Or &H2)
    Call SystemParametersInfo(SPI_SETFONTSMOOTHINGTYPE, &H0, ClearType, &H1 Or &H2)
    
    Call LoadConfigIni
    
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEItemU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEItemDCK)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Drop, INT_DROP)
    Call MainTimer.SetInterval(TimersIndex.BuySell, INT_BUY_SELL)
    Call MainTimer.SetInterval(TimersIndex.PublicMessage, INT_PUB_MSG)
    Call MainTimer.SetInterval(TimersIndex.Talk, INT_TALK)
    Call MainTimer.SetInterval(TimersIndex.Medit, INT_MEDIT)
    Call MainTimer.SetInterval(TimersIndex.RandomName, INT_RAND_NAME)
    frmMain.MacroTrabajo.Interval = INT_MACRO_TRABAJO
    frmMain.MacroTrabajo.Enabled = False
    
    'Load the form for screenshots
    Call Load(frmScreenshots)
    
    'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    Call MainTimer.Start(TimersIndex.Drop)
    Call MainTimer.Start(TimersIndex.BuySell)
    Call MainTimer.Start(TimersIndex.PublicMessage)
    Call MainTimer.Start(TimersIndex.Talk)
    Call MainTimer.Start(TimersIndex.Medit)
    Call MainTimer.Start(TimersIndex.RandomName)
    
    'Set the dialog's font
    Dialogos.font = frmCharge.font
    
    Call InicializarNombres
    
    'Initialize FONTTYPES
    Call modFonts.InitFonts

    Call modResolution.SetResolution
                         
    'Set resolution BEFORE the loading form is displayed, therefore it will be centeRed.
    If ChangeResolution Or ResolucionActual Then
        If Not InitTileEngine(frmMain.hWnd, frmMain.MainViewShp.Top, frmMain.MainViewShp.Left, 13, 17, 9, 8, 8, 0.018) Then
            Call CloseClient
        End If
    Else
        If Not InitTileEngine(frmMain.hWnd, frmMain.MainViewShp.Top + TitleBarHeight, frmMain.MainViewShp.Left + 3, 13, 17, 9, 8, 8, 0.018) Then
            Call CloseClient
        End If
    End If
    
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarMiniMapa

    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(frmMain.PicInv, MaxInvSlots)
    
    Call Cinturon.Initialize(frmMain.PicBelt, MaxBeltSlots)
    
    Call Hechizos.Initialize(frmMain.PicSpellInv, MaxSpellSlots)
    
    'Call InvCompanieros.Initialize(DirectDraw, frmMain.PicCompaInv, MaxCompaSlots)
    
    'If Not ChangeResolution And Not ResolucionActual Then
    'frmConnect.BorderStyle = 3
    'frmConnect.Caption = frmConnect.Caption
    'End If
    
    Call Audio.mSound_PlayWav(SND_INTRO)

    'Lo.pongo.antes
    frmConnect.Visible = True

    Call SetWindowPos(frmConnect.hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
                
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    
    Call Audio.MusicMP3Play(SND_MUSIC_INTRO)
        
    Do While prgRun
        'Sólo dibujamos si la ventana no está minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
            
            'Play ambient sounds
            Call RenderSounds
            
            Call CheckKeys
        End If
        
        'If there is anything to be sent, we send it
        If outgoingData.length > 0 Then
            Call FlushBuffer
        End If
        
        DoEvents
    Loop
    
    Call CloseClient
End Sub

Public Sub DoMeditar()
    If UserMuerto Then 'Muerto
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
        End With
        Exit Sub
    End If

    If UserMinMan = UserMaxMan Then
        If Charlist(UserCharIndex).Priv < 2 Then
            Exit Sub
        End If
    End If
    
    If UserMoving Then
        Exit Sub
    End If
               
    Meditando = True

    With Charlist(UserCharIndex)
        Select Case UserLvl
            'Show proper FX according to level
            Case Is < 15
                .FxIndex = FX_MEDITARCHICO
            Case Is < 25
                .FxIndex = FX_MEDITARMEDIANO
            Case Is < 35
                .FxIndex = FX_MEDITARGRANDE
            Case Is < 40
                .FxIndex = FX_MEDITARXGRANDE
            Case Else
                .FxIndex = FX_MEDITARXXGRANDE
        End Select
        
        .fX.Loops = -1
    
        Call InitGrh(.fX, FxData(.FxIndex).Animacion)
        
        If PortalBufferIndex = 0 Then
            PortalBufferIndex = Audio.mSound_PlayWav(SND_PORTAL, 1)
        End If
    End With

End Sub
Public Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'Writes a var to a text file
    WritePrivateProfileString Main, Var, Value, file
End Sub

Public Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
'Gets a Var from a text file
    Dim sSpaces As String 'This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) 'This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    GetPrivateProfileString Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Function CheckMailString(ByVal sString As String) As Boolean
'Función para chequear el email

On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos > 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then
            Exit Function
        End If
        
        '3er test: Recorre todos los carácteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then
                    Exit Function
                End If
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Public Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562)) And _
                MapData(X, Y).Graphic(2).GrhIndex = 0
                
End Function

Public Sub LoadConfigIni()
On Error GoTo Error
    If FileExist(DataPath & "Config.ini", vbArchive) Then
        MusicActivated = (GetVar(DataPath & "Config.ini", "INIT", "MusicActivated")) > 0
        SoundEffectsActivated = (GetVar(DataPath & "Config.ini", "INIT", "SoundEffectsActivated")) > 0
        ChangeResolution = (GetVar(DataPath & "Config.ini", "INIT", "ChangeResolution")) > 0
        AlphaBActivated = (GetVar(DataPath & "Config.ini", "INIT", "AlphaBActivated")) > 0
    Else
        MusicActivated = True
        SoundEffectsActivated = True
        ChangeResolution = True
        AlphaBActivated = True
    End If
Error:
End Sub

Public Sub SaveConfigIni()
    If FileExist(DataPath & "Config.ini", vbArchive) Then
        WriteVar DataPath & "Config.ini", "INIT", "MusicActivated", CStr(IIf(MusicActivated, 1, 0))
        WriteVar DataPath & "Config.ini", "INIT", "SoundEffectsActivated", CStr(IIf(SoundEffectsActivated, 1, 0))
        WriteVar DataPath & "Config.ini", "INIT", "ChangeResolution", CStr(IIf(ChangeResolution, 1, 0))
        WriteVar DataPath & "Config.ini", "INIT", "AlphaBActivated", CStr(IIf(AlphaBActivated, 1, 0))
    End If
End Sub

Private Sub InicializarNombres()
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.

    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cLindos) = "Lindos"
    Ciudades(eCiudad.cArghal) = "Arghâl"
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"

    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clérigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladrón"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladín"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Pirat) = "Pirata"
    
    SkillName(eSkill.Magia) = "Magia"
    SkillName(eSkill.Robar) = "Robar"
    SkillName(eSkill.Tacticas) = "Evasión en combate"
    SkillName(eSkill.Armas) = "Combate cuerpo a cuerpo"
    SkillName(eSkill.Meditar) = "Meditar"
    SkillName(eSkill.Apuñalar) = "Apuñalar"
    SkillName(eSkill.Ocultarse) = "Ocultarse"
    SkillName(eSkill.Supervivencia) = "Supervivencia"
    SkillName(eSkill.Talar) = "Talar árboles"
    SkillName(eSkill.Defensa) = "Defensa con escudos"
    SkillName(eSkill.Pescar) = "Pescar"
    SkillName(eSkill.Minar) = "Minar"
    SkillName(eSkill.Carpinteria) = "Carpintería"
    SkillName(eSkill.Herreria) = "Herreria"
    SkillName(eSkill.Liderazgo) = "Liderazgo"
    SkillName(eSkill.Domar) = "Domar animales"
    SkillName(eSkill.Proyectiles) = "Combate a distancia"
    SkillName(eSkill.Wrestling) = "Combate sin armas"
    SkillName(eSkill.Navegacion) = "Navegación"

    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"
End Sub

Public Sub CloseClient()
'Frees all used resources, cleans up and leaves
    
    EngineRun = False

    Call ResetResolution
    
    'Stop tile engine
    Call DeinitTileEngine
        
    'Destruimos los objetos públicos creados
    Set CustomMessages = Nothing
    Set CustomKeys = Nothing
    Set Dialogos = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set Hechizos = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    
    Call UnloadAllForms
    
    'If FontSmooth Then
        'Call SystemParametersInfo(SPI_SETFONTSMOOTHING, 0&, &H0, &H1 Or &H2)
        'Call SystemParametersInfo(SPI_SETFONTSMOOTHINGTYPE, &H0, NoType, &H1 Or &H2)
    'End If
    
    'Call UnloadRessources
    
    Call SetMusicInfo(vbNullString, vbNullString, vbNullString)
    End
End Sub

Public Function getTagPosition(ByVal Nick As String) As Integer
    Dim buf As Integer
    
    buf = InStr(Nick, "<")
    
    If buf > 0 Then
        getTagPosition = buf
        Exit Function
    End If
    
    buf = InStr(Nick, "[")
    
    If buf > 0 Then
        getTagPosition = buf
        Exit Function
    End If
    
    getTagPosition = Len(Nick) + 2
End Function

Public Sub checkText(ByVal Text As String)
    Dim Nivel As Integer
    
    If Right(Text, Len(MENSAJE_FRAGSHOOTER_TE_HA_MATADO)) = MENSAJE_FRAGSHOOTER_TE_HA_MATADO Then
        Call ScreenCapture(True)
        Exit Sub
    End If
    
    If Left(Text, Len(MENSAJE_FRAGSHOOTER_HAS_MATADO)) = MENSAJE_FRAGSHOOTER_HAS_MATADO Then
        EsperandoLevel = True
        Exit Sub
    End If
    
    If EsperandoLevel Then
        If Right(Text, Len(MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA)) = MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA Then
            'If CInt(mid(Text, Len(MENSAJE_FRAGSHOOTER_HAS_GANADO), (Len(Text) - (Len(MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA) + Len(MENSAJE_FRAGSHOOTER_HAS_GANADO))))) * 0.5 > ClientSetup.byMurdeRedLevel Then
                Call ScreenCapture(True)
            'End If
        End If
    End If
    
    EsperandoLevel = False
End Sub

Public Function getStrenghtColor() As Long
    Dim M As Long
    M = 255 / MAXATRIBUTOS
    getStrenghtColor = RGB(255 - (M * UserFuerza), (M * UserFuerza), 0)
End Function
Public Function getDexterityColor() As Long
    Dim M As Long
    M = 255 / MAXATRIBUTOS
    getDexterityColor = RGB(255, M * UserAgilidad, 0)
End Function

Public Sub Make_Transparent_Richtext(ByVal hWnd As Long)
    Call SetWindowLong(hWnd, -20, &H20&)
End Sub

Public Sub Make_Transparent_Form(ByVal hWnd As Long, ByVal Transparencia As Byte)
'If TransparenciaEnForms Then
    Call SetWindowLong(hWnd, -20, GetWindowLong(hWnd, -20) Or &H80000)
    Call SetLayeredWindowAttributes(hWnd, 0, Transparencia, &H2)
'End If
End Sub

'MOVER FORMULARIOS
Public Sub Auto_Drag(ByVal hWnd As Long)
    Call ReleaseCapture
    Call SendMessage(hWnd, &HA1, 2, ByVal 0&)
End Sub

Public Sub DibujarConsola(ByRef Mensaje As String, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal bCrLf As Boolean = True)

    Dim i As Byte
    
    With Consola
        For i = 2 To 10
            .MensajeConsola(i - 1) = .MensajeConsola(i)
            .Color_Red(i - 1) = .Color_Red(i)
            .Color_Green(i - 1) = .Color_Green(i)
            .Color_Blue(i - 1) = .Color_Blue(i)
        Next i
            i = 10
            .Color_Red(10) = Red
            .Color_Green(10) = Green
            .Color_Blue(10) = Blue
            .MensajeConsola(10) = Mensaje
    End With
    
End Sub

Public Function PonerPuntos(ByVal Numero As Long) As String

    Dim i As Integer
    Dim Cifra As String
    
    Cifra = str(Numero)
    Cifra = Right$(Cifra, Len(Cifra) - 1)
    For i = 0 To 4
        If Len(Cifra) - 3 * i >= 3 Then
            If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> vbNullString Then
                PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos
            End If
        Else
            If Len(Cifra) - 3 * i > 0 Then
                PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos
            End If
            Exit For
        End If
    Next
    
    PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)

End Function

'SubNick del MSN:
Public Sub SetMusicInfo(ByRef r_sArtist As String, ByRef r_sAlbum As String, ByRef r_sTitle As String, Optional ByRef r_sWMContentID As String = vbNullString, Optional ByRef r_sFormat As String = "{0} - {1}", Optional ByRef r_bShow As Boolean = True)
    Dim udtData As COPYDATASTRUCT
    Dim sBuffer As String
    Dim hMSGRUI As Long
    sBuffer = "/0Games\0" & Abs(r_bShow) & "/0" & r_sFormat & "/0" & r_sArtist & "/0" & r_sTitle & "/0" & r_sAlbum & "/0" & r_sWMContentID & "/0" & vbNullChar
    udtData.dwData = &H547
    udtData.lpData = StrPtr(sBuffer)
    udtData.cbData = LenB(sBuffer)
    Do
    hMSGRUI = FindWindowEx(0&, hMSGRUI, "MsnMsgrUIManager", vbNullString)
    If (hMSGRUI > 0) Then
    Call SendMessage(hMSGRUI, &H4A, 0, VarPtr(udtData))
    End If
    Loop Until (hMSGRUI = 0)
End Sub

Public Function CalculateBuyPrice(ByRef ObjValue As Long, ByVal ObjAmount As Integer) As Long
On Error GoTo Error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateBuyPrice = ObjValue * ObjAmount
    
    Exit Function
Error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.Number
End Function

Public Function CalculateSellPrice(ByRef ObjValue As Long, ByVal ObjAmount As Integer) As Long
On Error GoTo Error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateSellPrice = ObjValue * ObjAmount * 0.5
    
    Exit Function
Error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.Number
End Function

Public Function EsCompaniero(ByVal CharName As String) As Byte

On Error GoTo ErrHandler
    
    Dim i As Byte
    
    For i = 1 To MaxCompaSlots
        If LenB(Compa(i).Nombre) = LenB(CharName) Then
            If Compa(i).Nombre = CharName Then
                EsCompaniero = i
                Exit Function
            End If
        End If
    Next i
    Exit Function
    
ErrHandler:

End Function

Public Sub AddtoCommerceRecTxt(ByVal Text As String, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
    
    With frmComerciarUsu.CommerceConsole
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = Bold
        .SelItalic = Italic
        
        If Not Red = -1 Then
            .SelColor = RGB(Red, Green, Blue)
        End If
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
    End With

End Sub

Public Sub MeterEnInventario(ByRef Objeto As tInv)

    Dim Slot As Byte
    
    If NroItems > 0 Then
        For Slot = 1 To MaxInvSlots
            If Inv(Slot).ObjIndex = Objeto.ObjIndex And Inv(Slot).Amount < MaxInvObjs Then
                Call Inventario.SetSlotAmount(Slot, Inv(Slot).Amount + 1)
                Exit Sub
            End If
        Next Slot
    End If
    
    For Slot = 1 To MaxInvSlots
        If Inv(Slot).ObjIndex < 1 Then
            Exit For
            
        ElseIf Slot = MaxInvSlots Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("No podés llevar nada más.", .Red, .Green, .Blue, .Bold, .Italic)
            End With
            Exit Sub
        End If
    Next Slot
    
    Call Inventario.SetSlot(Slot, Objeto.ObjIndex, Objeto.Amount, Objeto.GrhIndex, Objeto.ObjType, _
    Objeto.MinHit, Objeto.MaxHit, Objeto.MinDef, Objeto.MaxDef, Objeto.Valor, Objeto.Name, _
    Objeto.PuedeUsar, Objeto.Proyectil)
    
    NroItems = NroItems + 1
End Sub

Public Sub Equipar()

    If UserMuerto Then 'Muerto
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("Estás muerto.", .Red, .Green, .Blue, .Bold, .Italic)
        End With
        Exit Sub
    End If
    
    If Not MainTimer.Check(TimersIndex.UseItemWithU) Then
        Exit Sub
    End If
    
    If InvSelSlot < 1 Or InvSelSlot > MaxInvSlots Then
        Exit Sub
    End If
    
    Dim Objeto As tInv
    
    With Inv(InvSelSlot)
        Objeto.ObjIndex = .ObjIndex
        Objeto.GrhIndex = .GrhIndex
        Objeto.ObjType = .ObjType
        Objeto.MinHit = .MinHit
        Objeto.MaxHit = .MaxHit
        Objeto.MinDef = .MinDef
        Objeto.MaxDef = .MaxDef
        Objeto.Valor = .Valor
        Objeto.Name = .Name
        
        If .ObjType = otFlecha Then
            Objeto.Amount = .Amount
        End If
        
        Objeto.PuedeUsar = .PuedeUsar
        Objeto.Proyectil = .Proyectil
    End With

    With Objeto
        If .ObjType <> otArma And _
        .ObjType <> otArmadura And _
        .ObjType <> otCasco And _
        .ObjType <> otEscudo And _
        .ObjType <> otFlecha And _
        .ObjType <> otFlecha And _
        .ObjType <> otAnillo And _
        .ObjType <> otBarco Then
            Exit Sub
        End If
        
        If .ObjType = otBarco Then
            Call WriteUseItem(InvSelSlot)
            Exit Sub
        End If
        
        If Not .PuedeUsar Then
            Call ShowConsoleMsg("Tu clase, género o raza no puede usar este objeto.", FontTypes(FontTypeNames.FONTTYPE_INFO).Red, FontTypes(FontTypeNames.FONTTYPE_INFO).Green, FontTypes(FontTypeNames.FONTTYPE_INFO).Blue, FontTypes(FontTypeNames.FONTTYPE_INFO).Bold, FontTypes(FontTypeNames.FONTTYPE_INFO).Italic)
            Exit Sub
        End If
    
        Dim SR As RECT
        Dim DR As RECT
    
        SR.Right = 32
        SR.bottom = 32
     
        DR.Right = 32
        DR.bottom = 32
                
        Select Case .ObjType
        
            Case otArma
            
                If .Proyectil Then
                    If LeftHandEqp.ObjIndex > 0 Then
                        If LeftHandEqp.ObjIndex = Objeto.ObjIndex Then
                            Exit Sub
                        End If
                        
                        Call DesEquipar(LeftHandEqp)
                    End If
                    
                    If RightHandEqp.ObjIndex > 0 Then
                        If RightHandEqp.ObjType <> otFlecha Then
                            Call DesEquipar(RightHandEqp)
                        End If
                    End If
                    
                    LeftHandEqp = Objeto
                    
                    frmMain.lblLeftHandEqp.Caption = .MinHit & "/" & .MaxHit
                    frmMain.picLeftHandEqp.Picture = frmMain.picLeftHandEqp.Picture
                    
                    Call DrawTransparentGrhtoHdc(frmMain.picLeftHandEqp.hdc, 0, 0, .GrhIndex, SR)
    
                    frmMain.picLeftHandEqp.Refresh
                    
                Else
                    If LeftHandEqp.ObjIndex > 0 Then
                        If LeftHandEqp.Proyectil Then
                            Call DesEquipar(LeftHandEqp)
                        End If
                    End If
                    
                    If RightHandEqp.ObjIndex > 0 Then
                        If RightHandEqp.ObjIndex = Objeto.ObjIndex Then
                            Exit Sub
                        End If
                        
                        Call DesEquipar(RightHandEqp)
                    End If
                    
                    RightHandEqp = Objeto
                    
                    frmMain.lblRightHandEqp.Caption = .MinHit & "/" & .MaxHit
                    
                    frmMain.picRightHandEqp.Picture = frmMain.picRightHandEqp.Picture
                    
                    Call DrawTransparentGrhtoHdc(frmMain.picRightHandEqp.hdc, 0, 0, .GrhIndex, SR)
    
                    frmMain.picRightHandEqp.Refresh
                End If
                
            Case otArmadura
                
                If BodyEqp.ObjIndex > 0 Then
                    Call DesEquipar(BodyEqp)
                End If
                
                BodyEqp = Objeto
                
                frmMain.lblBodyEqp.Caption = .MinDef & "/" & .MaxDef
                frmMain.picBodyEqp.Picture = frmMain.picBodyEqp.Picture
                
                Call DrawTransparentGrhtoHdc(frmMain.picBodyEqp.hdc, 0, 0, .GrhIndex, SR)
    
                frmMain.picBodyEqp.Refresh
            
            Case otCasco
            
                If HeadEqp.ObjIndex > 0 Then
                    Call DesEquipar(HeadEqp)
                End If
                
                HeadEqp = Objeto
                
                frmMain.lblHeadEqp.Caption = .MinDef & "/" & .MaxDef
                frmMain.picHeadEqp.Picture = frmMain.picHeadEqp.Picture
                
                Call DrawTransparentGrhtoHdc(frmMain.picHeadEqp.hdc, 0, 0, .GrhIndex, SR)
    
                frmMain.picHeadEqp.Refresh
                                    
            Case otEscudo
                
                If LeftHandEqp.ObjIndex > 0 Then
                    Call DesEquipar(LeftHandEqp)
                End If
                
                LeftHandEqp = Objeto
                
                frmMain.lblLeftHandEqp.Caption = .MinDef & "/" & .MaxDef
                frmMain.picLeftHandEqp.Picture = frmMain.picLeftHandEqp.Picture
                
                Call DrawTransparentGrhtoHdc(frmMain.picLeftHandEqp.hdc, 0, 0, .GrhIndex, SR)
    
                frmMain.picLeftHandEqp.Refresh
                
            Case otAnillo
        
                If RingEqp.ObjIndex > 0 Then
                    Call DesEquipar(RingEqp)
                End If
                
                RingEqp = Objeto
                
                frmMain.lblRingEqp.Caption = .MinDef & "/" & .MaxDef
                frmMain.picRingEqp.Picture = frmMain.picRingEqp.Picture
                
                Call DrawTransparentGrhtoHdc(frmMain.picRingEqp.hdc, 0, 0, .GrhIndex, SR)
    
                frmMain.picRingEqp.Refresh
                
            Case otFlecha
                
                If RightHandEqp.ObjIndex > 0 Then
                    Call DesEquipar(RightHandEqp)
                End If
                
                RightHandEqp = Objeto
                
                frmMain.lblRightHandEqp.Caption = .Amount
                 
                frmMain.picRightHandEqp.Picture = frmMain.picRightHandEqp.Picture
                
                Call DrawTransparentGrhtoHdc(frmMain.picRightHandEqp.hdc, 0, 0, .GrhIndex, SR)

                frmMain.picRightHandEqp.Refresh
                
        End Select
        
        If Inv(InvSelSlot).ObjType = otFlecha Or Inv(InvSelSlot).Amount < 2 Then
            Call Inventario.UnSetSlot(InvSelSlot)
            NroItems = NroItems - 1
        Else
            Call Inventario.SetSlotAmount(InvSelSlot, Inv(InvSelSlot).Amount - 1)
        End If
    End With
    
    Call WriteEquipItem(InvSelSlot)

End Sub

Public Sub DesEquipar(ByRef Objeto As tInv)
    
    If Objeto.ObjIndex < 1 Then
        Exit Sub
    End If
    
    Call MeterEnInventario(Objeto)

    Call WriteUnEquipItem(Objeto.ObjType)

    Select Case Objeto.ObjType

        Case otArmadura
            BodyEqp.ObjIndex = 0
            frmMain.picBodyEqp.Picture = frmMain.picBodyEqp.Picture
            frmMain.lblBodyEqp.Caption = vbNullString
                
        Case otCasco
            HeadEqp.ObjIndex = 0
            frmMain.picHeadEqp.Picture = frmMain.picHeadEqp.Picture
            frmMain.lblHeadEqp.Caption = vbNullString
        
        Case otArma
            If Objeto.Proyectil Then
                LeftHandEqp.ObjIndex = 0
                frmMain.picLeftHandEqp.Picture = frmMain.picLeftHandEqp.Picture
                frmMain.lblLeftHandEqp.Caption = vbNullString
            Else
                RightHandEqp.ObjIndex = 0
                frmMain.picRightHandEqp.Picture = frmMain.picRightHandEqp.Picture
                frmMain.lblRightHandEqp.Caption = vbNullString
            End If
                        
        Case otEscudo
            LeftHandEqp.ObjIndex = 0
            frmMain.picLeftHandEqp.Picture = frmMain.picLeftHandEqp.Picture
            frmMain.lblLeftHandEqp.Caption = vbNullString
         
        Case otFlecha
            RightHandEqp.ObjIndex = 0
            frmMain.picRightHandEqp.Picture = frmMain.picRightHandEqp.Picture
            frmMain.lblRightHandEqp.Caption = vbNullString
            
        Case otCinturon
            BeltEqp.ObjIndex = 0
            frmMain.picBeltEqp.Picture = frmMain.picRingEqp.Picture
            frmMain.lblBeltEqp.Caption = vbNullString

        Case otAnillo
            RingEqp.ObjIndex = 0
            frmMain.picRingEqp.Picture = frmMain.picRingEqp.Picture
            frmMain.lblRingEqp.Caption = vbNullString
    
        Case otBarco
            Ship.ObjIndex = 0
            frmMain.picShip.Picture = frmMain.picShip.Picture
            frmMain.lblShip.Caption = vbNullString
        
    End Select

End Sub

Public Sub Morir()

On Error Resume Next

    Call EraseChar(UserCharIndex)
        
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg("Estás muerto. Tenés un minuto para ser revivido, de lo contrario volverás a tu hogar. Para volver ahora, presioná la tecla 'Espacio'.", .Red, .Green, .Blue, .Bold, .Italic)
    End With
    
    UserMuerto = True
    Descansando = False
    Meditando = False
    UserCiego = False
    UserEstupido = False
    UserParalizado = False
    Charlist(UserCharIndex).Invisible = False
    Charlist(UserCharIndex).Paralizado = False
    
    UserMinHP = 0
    UserMinMan = 0
    UserMinSTA = 0
        
    UsingSkill = 0
    frmMain.MousePointer = vbDefault
    
    If frmMain.TrainingMacro Then
        frmMain.DesactivarMacroHechizos
    End If
    
    If frmMain.MacroTrabajo Then
        frmMain.DesactivarMacroTrabajo
    End If

    Call RemoveDamage
    
    Dim i As Byte
    
    For i = 1 To MaxSpellSlots
        If Spell(i).Grh > 0 Then
            If Spell(i).PuedeLanzar Then
                Spell(i).PuedeLanzar = False
                Call Hechizos.DrawSpellSlot(i)
            End If
        End If
    Next i

End Sub
Public Sub UserChat(ByVal Chat As String, ByVal TipoChat As eChatType, Optional ByVal CompaSlot As Byte, Optional ByVal Name As String)

On Error Resume Next

    If Not MainTimer.Check(TimersIndex.Talk) Then
        Exit Sub
    End If
    
    'Dim Text As String
    
    'Text = Chat
    
    'Text = Replace(LCase$(Text), "pelotudo", "********")
    'Text = Replace(LCase$(Text), "hijo de puta", "* * *")
    'Text = Replace(LCase$(Text), "hdp", "***")
    'Text = Replace(LCase$(Text), "alkon", "*")
    'Text = Replace(LCase$(Text), "imperium", "*")
    'Text = Replace(LCase$(Text), " iao", " *")
    'Text = Replace(LCase$(Text), "tpao", "*")
    'Text = Replace(LCase$(Text), "tds", "*")
    'Text = Replace(LCase$(Text), "t d s", "*")
    'Text = Replace(LCase$(Text), "i a o", "*")
    'Text = Replace(LCase$(Text), "a l k o n", "*")
    'Text = Replace(LCase$(Text), "i m p e r i u m", "*")
    'Text = Replace(LCase$(Text), "puto", "****")
    'Text = Replace(LCase$(Text), "p u t o", "****")
    'Text = Replace(LCase$(Text), "pt", "**")
    'Text = Replace(LCase$(Text), "p t", "**")
    'Text = Replace(LCase$(Text), " ao", " Abraxas")
    'Text = Replace(LCase$(Text), "sv", "juego")
    'Text = Replace(LCase$(Text), "server", "juego")
    'Text = Replace(LCase$(Text), "nw", "**")
    'Text = Replace(LCase$(Text), "n w", "*")
    'Text = Replace(LCase$(Text), "newbie", "*")
    'Text = Replace(LCase$(Text), "n e w b i e", "*")

    'If LCase$(Chat) <> LCase$(Text) Then
    '    Chat = Text
    'End If
    
    'If Chat = "e" Or Chat = "ee" Or Chat = "eee" Then
    '    Chat = "Eh"
    'End If
    
    Select Case TipoChat
    
        Case eChatType.Norm
            
            Call WriteTalk(Chat)
    
            If Charlist(UserCharIndex).Priv < 2 Then
            
                'If Charlist(UserCharIndex).Lvl < 15 Then
                '    With FontTypes(FontTypeNames.FONTTYPE_PRINCIPIANTE)
                '        Call ShowConsoleMsg(UserName & ": ", .Red, .Green, .Blue, .Italic, .Bold, True, True)
                '        Call Dialogos.CreateDialog(Chat, UserCharIndex, RGB(.Red, .Green, .Blue))
                '    End With
                'Else
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(UserName & ": ", .Red + 15, .Green + 15, .Blue + 15, .Italic, .Bold, True, True)
                    End With
                        
                    With FontTypes(FontTypeNames.FONTTYPE_TALK)
                        Call Dialogos.CreateDialog(Chat, UserCharIndex, RGB(.Red, .Green, .Blue))
                    End With
                'End If
                
            Else
                With FontTypes(FontTypeNames.FONTTYPE_GM)
                    Call ShowConsoleMsg(UserName & ": ", .Red, .Green, .Blue, .Italic, .Bold, True, True)
                    Call Dialogos.CreateDialog(Chat, UserCharIndex, RGB(.Red, .Green, .Blue))
                End With
            End If
            
            If Right$(Chat, 1) = "!" Then
                If Charlist(UserCharIndex).Priv < 2 Then
                    With FontTypes(FontTypeNames.FONTTYPE_YELL)
                        Call ShowConsoleMsg(Chat, .Red, .Green, .Blue, .Bold, .Italic, , True)
                        Call Dialogos.CreateDialog(Chat, UserCharIndex, RGB(.Red, .Green, .Blue))
                    End With
                Else
                    With FontTypes(FontTypeNames.FONTTYPE_YELLGM)
                        Call ShowConsoleMsg(Chat, .Red, .Green, .Blue, .Bold, .Italic, , True)
                        Call Dialogos.CreateDialog(Chat, UserCharIndex, RGB(.Red, .Green, .Blue))
                    End With
                End If
                
            Else
                If Charlist(UserCharIndex).Priv < 2 Then
                    With FontTypes(FontTypeNames.FONTTYPE_TALK)
                        Call ShowConsoleMsg(Chat, .Red, .Green, .Blue, .Bold, .Italic, , True)
                        Call Dialogos.CreateDialog(Chat, UserCharIndex, RGB(.Red, .Green, .Blue))
                    End With
                             
                Else
                    With FontTypes(FontTypeNames.FONTTYPE_TALKGM)
                        Call ShowConsoleMsg(Chat, .Red, .Green, .Blue, .Bold, .Italic, , True)
                        Call Dialogos.CreateDialog(Chat, UserCharIndex, RGB(.Red, .Green, .Blue))
                    End With
                End If
            End If
        
            LastParsedString = Chat

        Case eChatType.Guil
            Call WriteGuildMessage(Chat)
            
        Case eChatType.Komp
        
            Call WriteCompaMessage(CompaSlot, Chat)
                            
            With FontTypes(FontTypeNames.FONTTYPE_COMPAMESSAGE)
                Call ShowConsoleMsg(UserName & ": ", .Red, .Green, .Blue, .Bold, .Italic, True, True)
                
                Dim i As Integer
                
                For i = 1 To LastChar
                    If Charlist(i).EsUser Then
                        If LenB(Charlist(i).Nombre) = LenB(Compa(CompaSlot).Nombre) Then
                            If Charlist(i).Nombre = Compa(CompaSlot).Nombre Then
                                Call Dialogos.CreateDialog(Chat, UserCharIndex, RGB(.Red, .Green, .Blue))
                                Exit For
                            End If
                        End If
                    End If
                Next i
            End With
            
            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                Call ShowConsoleMsg(Chat, .Red, .Green, .Blue, .Bold, .Italic, , True)
            End With
            
        Case eChatType.Priv
        
            Call WritePrivateMessage(Name, Chat)

        Case eChatType.Glob
        
            If Not MainTimer.Check(TimersIndex.PublicMessage) Then
                Exit Sub
            End If
        
            Call WritePublicMessage(Chat)
            
            With FontTypes(FontTypeNames.FONTTYPE_PUBLICMESSAGE)
                Call ShowConsoleMsg(UserName & ": ", .Red, .Green, .Blue, .Bold, .Italic, True, True)
            End With
        
            Call ShowConsoleMsg(Chat, 170, 170, 170, False, True, , True)
    
            LastParsedString = vbNullString
    End Select
    
End Sub

Public Function browseName(gh As Integer) As Integer
    Dim i As Integer
    For i = 1 To frmMain.MouseImage.ListImages.Count
        If frmMain.MouseImage.ListImages(i).Key = "g" & CStr(gh) Then
            browseName = i
            Exit For
        End If
    Next i
End Function

Public Sub SetBeltSlot(ByVal Slot As Byte, ByVal ObjIndex As Integer, ByVal Amount As Integer, _
                       ByVal GrhIndex As Integer, ByVal Valor As Long, ByVal Name As String, ByVal PuedeUsar As Boolean)
    
        With Belt(Slot)
            .Amount = Amount
            .Name = Name
            .ObjIndex = ObjIndex
            .GrhIndex = GrhIndex
            .Valor = Valor
            .PuedeUsar = PuedeUsar
        End With

    Call Cinturon.DrawBeltSlot(Slot)
    
End Sub

Public Sub SetSpellSlot(ByVal Slot As Byte, ByVal Grh As Integer, ByVal Nombre As String, ByVal MinSkill As Boolean, _
    ByVal ManaRequerido As Integer, ByVal StaRequerido As Integer, ByVal NeedStaff As Integer)
    
    With Spell(Slot)
        .Grh = Grh
        .Nombre = Nombre
        .MinSkill = MinSkill
        .ManaRequerido = ManaRequerido
        .StaRequerido = StaRequerido
        .NeedStaff = NeedStaff
        
        If .Grh > 24032 Then
            .Grh = 24032
        End If
            
        If .Grh = 24031 Then
            .Grh = 24032
        End If
        
        If .Grh = 24003 Then
            .Grh = 24002
        End If

        If .ManaRequerido > UserMinMan Or .StaRequerido > UserMinSTA Then
           .PuedeLanzar = False
        Else
           .PuedeLanzar = True
        End If
        
        'Call frmMain.lstSpells.AddItem(.Nombre)
    End With

    Call Hechizos.DrawSpellSlot(Slot)
End Sub

Public Sub SetCompaSlot(ByVal Slot As Byte, ByVal Nombre As String, ByVal Online As Boolean)
    With Compa(Slot)
        .Nombre = Nombre
        .Online = Online
    End With
End Sub

Public Sub UnSetCompaSlot(ByVal Slot As Byte)
    With Compa(Slot)
        .Nombre = vbNullString
        .Online = False
    End With
End Sub

