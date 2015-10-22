Attribute VB_Name = "ModTileEngine"
Option Explicit

Private Const GrhFogata As Integer = 1521

Private Const GrhPortal As Integer = 669

'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Encabezado bmp
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    x As Long
    y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    map As Integer
    x As Integer
    y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    PixelWidth As Integer
    PixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    Speed As Single
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
End Type

'List   a de cuerpos
Public Type BodyData
    Walk(eHeading.NORTH To eHeading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(eHeading.NORTH To eHeading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(eHeading.NORTH To eHeading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(eHeading.NORTH To eHeading.WEST) As Grh
End Type


'Apariencia del personaje
Public Type Char
    Active As Byte
    Heading As eHeading
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    
    fX As Grh
    FxIndex As Integer
        
    Nombre As String
    
    Guilda As String
    AlineacionGuilda As Byte
    
    ScrollDirectionX As Integer
    ScrollDirectionY As Integer
    
    Moving As Boolean
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    ScreenX As Single
    ScreenY As Single
    
    Pie As Boolean
    Muerto As Boolean
    Invisible As Boolean
    Paralizado As Boolean
    Priv As Byte
    Lvl As Byte
    CompaIndex As Byte
    MascoIndex As Byte
    Quieto As Boolean
    EsUser As Boolean
End Type

'Info de un objeto
Public Type ObjInfo
    Index As Integer
    Amount As Long
    Grh As Grh
    Name As String
    ObjType As eObjType
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    
    Obj As ObjInfo
    TileExit As WorldPos
    Blocked As Boolean
    
    Trigger As Integer
    
    fX As Grh
    FxIndex As Integer
End Type

'Info de cada mapa
Public Type MapInfoBlock
    Name As String
    Version As Integer
    Zone As String
    Music As Byte
    Top As Byte
    Left As Byte
End Type

'DX7 Objects
Public DirectX As New DirectX7
Public DirectDraw As DirectDraw7
Private PrimarySurface As DirectDrawSurface7
Private PrimaryClipper As DirectDrawClipper
Private BackBufferSurface As DirectDrawSurface7

'Bordes del mapa
Public MinXBorder As Integer
Public MaxXBorder As Integer
Public MinYBorder As Integer
Public MaxYBorder As Integer

Public RangoVisionX As Byte
Public RangoVisionY As Byte

'Status del user
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public PrevUserPos As Position 'Si se mueve
Public UserCharIndex As Integer
Public UserMap As Integer

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long
Private fpsLastCheck As Long

'Tamaño del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer

'Offset del desde 0,0 del main view
Private MainViewTop As Integer
Private MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Dim timerElapsedTime As Single
Dim timerTicksPerFrame As Single
Dim engineBaseSpeed As Single


Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Private MainViewWidth As Integer
Private MainViewHeight As Integer

Public MouseTileX As Integer
Public MouseTileY As Integer

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock 'Mapa
Public MapInfo() As MapInfoBlock 'Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?

Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer

Public Charlist(1 To 10000) As Char

'Used by GetTextExtentPoint32
Private Type size
    cX As Long
    cY As Long
End Type

Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum

Private Declare Function BltAlphaFast Lib "vbabdx" (ByRef lpDDSDest As Any, ByRef lpDDSSource As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchSrc As Long, ByVal pitchDst As Long, ByVal dwMode As Long) As Long
Private Declare Function BltEfectoNoche Lib "vbabdx" (ByRef lpDDSDest As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchDst As Long, ByVal dwMode As Long) As Long

Private Declare Function vbDABLalphablend16 Lib "vbDABL" (ByVal iMode As Integer, ByVal bColorKey As Integer, _
        ByRef sPtr As Any, ByRef dPtr As Any, ByVal iAlphaVal As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, _
ByVal isPitch As Integer, ByVal idPitch As Integer, ByVal iColorKey As Integer) As Integer
        Public Declare Function vbDABLcolorblend16555 Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
        Public Declare Function vbDABLcolorblend16565 Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
        Public Declare Function vbDABLcolorblend16555ck Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
        Public Declare Function vbDABLcolorblend16565ck Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'Text width computation. Needed to center text.
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As size) As Long

Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

'RENDERCHARNAME
Private CharName As String
Private CharColor As Long

'RENDEROBJNAME
Private ObjName As String
Private ObjType As eObjType
Private ObjX As Integer
Private ObjY As Integer

'RENDERDAMAGE
Private Damage As String
Private startTime As Long
Private x As Integer
Private y As Integer
Private SUBe As Integer

'RENDERCHARDAMAGE
Private startTime4 As Long
Private SUBe4 As Integer
Private CharX As Integer
Private CharY As Integer

Public AttackerCharIndex As Integer
Public CharDamage As String
Public CharDamage2 As String

'RENDERCHARHP
Public CharMinHP As Byte
Public TempCharHP As Integer
Private CharX2 As Integer
Private CharY2 As Integer

Public AttackedCharIndex As Integer

Public SelectedCharIndex As Integer

'RENDEREXP
Private Exp As String
Private StartTime2 As Long
Private Y2 As Integer
Private X2 As Integer
Private SUBe2 As Integer

'RENDERGLD
Private Gld As Long
Private StartTime3 As Long
Private Y3 As Integer
Private X3 As Integer
Private SUBe3 As Integer

Public DamageType As Byte

Public CharDamageType As Byte

Public Sub CargarCabezas()
    Dim n As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    n = FreeFile()
    Open DataPath & "Cabezas.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #n, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #n
End Sub

Public Sub CargarCascos()
    Dim n As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    n = FreeFile()
    Open DataPath & "Cascos.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #n, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #n
End Sub

Public Sub CargarCuerpos()
    Dim n As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    n = FreeFile()
    Open DataPath & "Cuerpos.ind" For Binary Access Read As #n
    
    Get #n, , MiCabecera
    
    Get #n, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #n, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #n
End Sub

Public Sub CargarFxs()
    Dim n As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    n = FreeFile()
    Open DataPath & "Fx.ind" For Binary Access Read As #n
    
    Get #n, , MiCabecera
    
    Get #n, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #n, , FxData(i)
    Next i
    
    Close #n
End Sub

Public Sub CargarArrayLluvia()
    Dim n As Integer
    Dim i As Long
    Dim Nu As Integer
    
    n = FreeFile()
    Open DataPath & "fk.ind" For Binary Access Read As #n
    
    Get #n, , MiCabecera
    
    Get #n, , Nu
    
    'Resize array
    ReDim bLluvia(1 To Nu) As Byte
    
    For i = 1 To Nu
        Get #n, , bLluvia(i)
    Next i
    
    Close #n
End Sub

Public Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Integer, ByRef tY As Integer)
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.

    tX = UserPos.x + viewPortX / 32 - WindowTileWidth * 0.5
    tY = UserPos.y + viewPortY / 32 - WindowTileHeight * 0.5
End Sub

Public Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal x As Integer, ByVal y As Integer, Optional ByVal Arma As Integer = 0, Optional ByVal Escudo As Integer = 0, Optional ByVal Casco As Integer = 0)

On Error Resume Next

    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then
        LastChar = CharIndex
    End If
    
    With Charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .Active = False Then
            NumChars = NumChars + 1
        End If
        
        If Arma < 1 Then
            Arma = 2
        End If
        
        If Escudo < 1 Then
            Escudo = 2
        End If
        
        If Casco < 1 Then
            Casco = 2
        End If
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = False
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.x = x
        .Pos.y = y
        
        'Make active
        .Active = True
    End With
    
    'Plot on map
    MapData(x, y).CharIndex = CharIndex

End Sub

Public Sub ResetCharInfo(ByVal CharIndex As Integer)
    With Charlist(CharIndex)
        .Active = False
        .FxIndex = 0
        .Invisible = False
        .Paralizado = False
        .Moving = False
        .Nombre = vbNullString
        .Pie = 0
        .Pos.x = 0
        .Pos.y = 0
        .Lvl = 0
        .Priv = 0
        .CompaIndex = 0
        .MascoIndex = 0
        .Quieto = False
        .EsUser = False
    End With
End Sub

Public Sub EraseChar(ByVal CharIndex As Integer)
'Erases a Char from CharList and map

On Error Resume Next
    Charlist(CharIndex).Active = False

    'Update lastchar
    If CharIndex = LastChar Then
        Do Until Charlist(LastChar).Active = True
            LastChar = LastChar - 1
            If LastChar = 0 Then
                Exit Do
            End If
        Loop
    End If
    
    If Charlist(CharIndex).Pos.x > 0 And Charlist(CharIndex).Pos.y > 0 Then
        MapData(Charlist(CharIndex).Pos.x, Charlist(CharIndex).Pos.y).CharIndex = 0
    End If
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    
    Call ResetCharInfo(CharIndex)
    
    If AttackedCharIndex = CharIndex Then
        AttackedCharIndex = 0
        CharMinHP = 0
    End If
    
    If AttackerCharIndex = CharIndex Then
        AttackerCharIndex = 0
        CharDamage2 = vbNullString
        DamageType = 1
    End If
    
    'Update NumChars
    NumChars = NumChars - 1

End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'Sets up a grh. MUST be done before rendering
    Grh.GrhIndex = GrhIndex
    
    If Grh.GrhIndex < 1 Or Grh.GrhIndex > 32000 Then
        Exit Sub
    End If
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then
            Started = 0
        End If
        
        Grh.Started = Started
    End If
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed * 1.3
End Sub

Public Sub MoveCharbyHead(ByVal nHeading As eHeading)
'Starts the movement of a Char in nHeading direction

On Error Resume Next

    If UserCharIndex < 1 Then
        Exit Sub
    End If
    
    Dim addX As Integer
    Dim addY As Integer
    Dim x As Integer
    Dim y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With Charlist(UserCharIndex)
        x = .Pos.x
        y = .Pos.y
        
        'Figure out which way to move
        Select Case nHeading
            Case eHeading.NORTH
                addY = -1
        
            Case eHeading.EAST
                addX = 1
        
            Case eHeading.SOUTH
                addY = 1
            
            Case eHeading.WEST
                addX = -1
        End Select
        
        nX = x + addX
        nY = y + addY
        
        'If nY < MinLimiteY Or nY > MaxLimiteY Or nX < MinLimiteX Or nX > MaxLimiteX Then
        '    Exit Sub
        'End If
        
        MapData(nX, nY).CharIndex = UserCharIndex
        .Pos.x = nX
        .Pos.y = nY
        MapData(x, y).CharIndex = 0

        .MoveOffsetX = -32 * addX
        .MoveOffsetY = -32 * addY
        
        .Moving = True
        .Heading = nHeading
        
        .ScrollDirectionX = addX
        .ScrollDirectionY = addY
    End With
    
    Call MoveScreen(nHeading)

    Call DoPasosFx(UserCharIndex)

    Call DibujarMiniMapa

End Sub

Public Sub DoPortalFx()
    Dim location As Position
    
    If bPortal Then
        bPortal = HayPortal(location)
        If Not bPortal Then
            Call Audio.StopWave(PortalBufferIndex)
            PortalBufferIndex = 0
        End If
    Else
        bPortal = HayPortal(location)
        If bPortal And PortalBufferIndex = 0 Then
            PortalBufferIndex = Audio.Play(SND_PORTAL, location.x, location.y, LoopStyle.Enabled)
        End If
    End If
End Sub

Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        If bFogata And FogataBufferIndex = 0 Then
            FogataBufferIndex = Audio.Play(SND_FOGATA, location.x, location.y, LoopStyle.Enabled)
        End If
    End If
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With Charlist(CharIndex).Pos
        EstaPCarea = .x > UserPos.x - MinXBorder And .x < UserPos.x + MinXBorder And .y > UserPos.y - MinYBorder And .y < UserPos.y + MinYBorder
    End With
End Function

Public Sub DoPasosFx(ByVal CharIndex As Integer)
    
    With Charlist(CharIndex)
        
        If .Priv > 1 And UserCharIndex <> CharIndex Then
            Exit Sub
        End If
    
        If EstaPCarea(CharIndex) And .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
    
            If Not UserNavegando Then
                .Pie = Not .Pie
                
                If .Pie Then
                    Call Audio.Play(SND_PASOS1, .Pos.x, .Pos.y)
                Else
                    Call Audio.Play(SND_PASOS2, .Pos.x, .Pos.y)
                End If
            Else
                Call Audio.Play(SND_NAVEGANDO, .Pos.x, .Pos.y)
            End If
            
        End If
    End With
End Sub

Public Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
    
On Error Resume Next

    If UserCharIndex < 1 Then
        Exit Sub
    End If
    
    If CharIndex = UserCharIndex Then
        If UserParalizado Then
            Exit Sub
        End If
    End If
    
    Dim x As Integer
    Dim y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As eHeading
    
    With Charlist(CharIndex)
        x = .Pos.x
        y = .Pos.y
        
        If x < 1 Or y < 1 Then
            Exit Sub
        End If
                        
        addX = nX - x
        addY = nY - y
        
        If Sgn(addX) = 1 Then
            nHeading = eHeading.EAST
        ElseIf Sgn(addX) = -1 Then
            nHeading = eHeading.WEST
        ElseIf Sgn(addY) = -1 Then
            nHeading = eHeading.NORTH
        ElseIf Sgn(addY) = 1 Then
            nHeading = eHeading.SOUTH
        End If

        If MapData(nX, nY).CharIndex = UserCharIndex Then
            With Charlist(UserCharIndex)
                Charlist(UserCharIndex).Pos = PrevUserPos
                MapData(.Pos.x, .Pos.y).CharIndex = UserCharIndex
                UserPos = PrevUserPos
            End With
        End If

        MapData(x, y).CharIndex = 0

        MapData(nX, nY).CharIndex = CharIndex
                
        .Pos.x = nX
        .Pos.y = nY
        
        .MoveOffsetX = -1 * (32 * addX)
        .MoveOffsetY = -1 * (32 * addY)
        
        .Moving = True
        .Heading = nHeading
        
        .ScrollDirectionX = Sgn(addX)
        .ScrollDirectionY = Sgn(addY)
    End With
    
    If Not EstaPCarea(CharIndex) Then
        Call Dialogos.RemoveDialog(CharIndex)
    End If
    
    If nY < MinLimiteY Or nY > MaxLimiteY Or nX < MinLimiteX Or nX > MaxLimiteX Then
        CharIndex = 0
        Debug.Print MaxLimiteY
        'Call EraseChar(CharIndex)
    End If
    
    Call DibujarMiniMapa
End Sub

Public Sub MoveScreen(ByVal nHeading As eHeading)
'Starts the screen moving in a direction

On Error GoTo Error

    Dim x As Integer
    Dim y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case eHeading.NORTH
            y = -1
        
        Case eHeading.EAST
            x = 1
        
        Case eHeading.SOUTH
            y = 1
        
        Case eHeading.WEST
            x = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.y + y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        PrevUserPos = UserPos

        'Start moving... MainLoop does the rest
        AddtoUserPos.x = x
        UserPos.x = tX
        AddtoUserPos.y = y
        UserPos.y = tY
        UserMoving = True
        
        Call Audio.MoveListener(UserPos.x, UserPos.y)
        
        bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
            MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
            MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
    End If
    
    Exit Sub
Error: MsgBox Err.Description
End Sub

Private Function HayPortal(ByRef location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.x - 8 To UserPos.x + 8
        For k = UserPos.y - 6 To UserPos.y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).Obj.Grh.GrhIndex = GrhPortal Then
                    location.x = j
                    location.y = k
                    
                    HayPortal = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.x - 8 To UserPos.x + 8
        For k = UserPos.y - 6 To UserPos.y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).Obj.Grh.GrhIndex = GrhFogata Then
                    location.x = j
                    location.y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function

Public Function NextOpenChar() As Integer
'Finds next open char Slot in CharList

    Dim loopc As Long
    Dim Dale As Boolean
    
    loopc = 1
    Do While Charlist(loopc).Active And Dale
        loopc = loopc + 1
        Dale = (loopc <= UBound(Charlist))
    Loop
    
    NextOpenChar = loopc
End Function

Private Function LoadGrhData() As Boolean

On Error GoTo ErrorHandler

    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    
    'Open files
    handle = FreeFile()
    
    Open DataPath & GrhFile For Binary Access Read As handle
    Seek #1, 1
    
    'Get file version
    Get handle, , fileVersion
    
    'Get number of grhs
    Get handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    While Not EOF(handle)
        Get handle, , Grh
        
        If Grh = 0 Then
            Grh = 1
        End If
        
        With GrhData(Grh)
            'Get number of frames
            Get handle, , .NumFrames
            
            If .NumFrames <= 0 Then
                GoTo ErrorHandler
            End If
            
            ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next Frame
                
                Get handle, , .Speed
                
                If .Speed <= 0 Then
                    GoTo ErrorHandler
                End If
                
                'Compute width and height
                .PixelHeight = GrhData(.Frames(1)).PixelHeight
                
                If .PixelHeight <= 0 Then
                    GoTo ErrorHandler
                End If
                
                .PixelWidth = GrhData(.Frames(1)).PixelWidth
                
                If .PixelWidth <= 0 Then
                    GoTo ErrorHandler
                End If
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                
                If .TileWidth <= 0 Then
                    GoTo ErrorHandler
                End If
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                
                If .TileHeight <= 0 Then
                    GoTo ErrorHandler
                End If
                
            Else
                'Read in normal GRH data
                Get handle, , .FileNum
                
                If .FileNum <= 0 Then
                    GoTo ErrorHandler
                End If
                
                Get handle, , GrhData(Grh).sX
                
                If .sX < 0 Then
                    GoTo ErrorHandler
                End If
                
                Get handle, , .sY
                
                If .sY < 0 Then
                    GoTo ErrorHandler
                End If
                
                Get handle, , .PixelWidth
                
                If .PixelWidth <= 0 Then
                    GoTo ErrorHandler
                End If
                
                Get handle, , .PixelHeight
                
                If .PixelHeight <= 0 Then
                    GoTo ErrorHandler
                End If
                
                'Compute width and height
                .TileWidth = .PixelWidth / 32
                .TileHeight = .PixelHeight / 32
                
                .Frames(1) = Grh
            End If
        End With
    Wend
    
    Close handle
        
    LoadGrhData = True
Exit Function

ErrorHandler:
    LoadGrhData = False
End Function

Public Function LegalPos(ByVal x As Integer, ByVal y As Integer) As Boolean
'Checks to see if a tile position is legal

    'Limites del mapa
    If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(x, y).Blocked Then
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(x, y).CharIndex > 0 Then
        Exit Function
    End If
    
    If UserNavegando <> HayAgua(x, y) Then
        Exit Function
    End If
    
    LegalPos = True
End Function

Public Function MoveToLegalPos(ByVal Direccion As Byte) As Boolean
'Checks to see if a tile position is legal, including if there is a casper in the tile

    Dim x As Integer
    Dim y As Integer
    Dim CharIndex As Integer
    
    Select Case Direccion
        Case eHeading.NORTH
            x = UserPos.x
            y = UserPos.y - 1
        Case eHeading.EAST
            x = UserPos.x + 1
            y = UserPos.y
        Case eHeading.SOUTH
            x = UserPos.x
            y = UserPos.y + 1
        Case eHeading.WEST
            x = UserPos.x - 1
            y = UserPos.y
    End Select
    
    'Limites del mapa
    If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(x, y).Blocked Then
        If Charlist(UserCharIndex).Priv < 2 Then
            Exit Function
        End If
    End If
    
    CharIndex = MapData(x, y).CharIndex
    
    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.x, UserPos.y).Blocked Then
            Exit Function
        End If
        
        With Charlist(CharIndex)
            'Si no es casper, no puede pasar
            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else
                If HayAgua(UserPos.x, UserPos.y) Then
                    If Not HayAgua(x, y) Then
                        Exit Function
                    End If
                ElseIf HayAgua(x, y) Then
                    Exit Function
                End If
                
                'Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If Charlist(UserCharIndex).Priv > 1 Then
                    If Charlist(UserCharIndex).Invisible Then
                        Exit Function
                    End If
                End If
            End If
        End With
    End If
      
    If UserNavegando <> HayAgua(x, y) Then
        If Charlist(UserCharIndex).Priv < 2 Then
            Exit Function
        End If
    End If
    
    MoveToLegalPos = True
End Function

Public Function InMapBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
'Checks to see if a tile position is in the maps bounds

    If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Private Sub DDrawGrhtoSurface(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Byte)
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
On Error GoTo Error
    
    If Grh.GrhIndex < 1 Then
        Exit Sub
    End If

    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
     
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * 32 * 0.5) + 32 * 0.5
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * 32) + 32
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .PixelWidth
        SourceRect.Bottom = SourceRect.Top + .PixelHeight
        
        'Draw
        Call BackBufferSurface.BltFast(x, y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_WAIT)
    End With
Exit Sub

Error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un Error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del Error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub

Public Sub DDrawTransGrhIndextoSurface(ByVal GrhIndex As Integer, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * 32 * 0.5) + 32 * 0.5
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * 32) + 32
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .PixelWidth
        SourceRect.Bottom = SourceRect.Top + .PixelHeight
        
        'Draw
        Call BackBufferSurface.BltFast(x, y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    End With
End Sub

Public Sub DDrawTransGrhtoSurface(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Byte)
'Draws a GRH transparently to a X and Y position

    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    
On Error GoTo Error
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * 32 * 0.5) + 32 * 0.5
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * 32) + 32
            End If
        End If
                
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .PixelWidth
        SourceRect.Bottom = SourceRect.Top + .PixelHeight
        
        'Draw
        Call BackBufferSurface.BltFast(x, y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    End With
Exit Sub

Error:

    Exit Sub
    
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un Error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del Error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub

Public Sub DDrawTransGrhtoSurfaceAlpha(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Byte)
'Draws a GRH transparently to a X and Y position

On Error Resume Next
    
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    Dim Src As DirectDrawSurface7
    Dim rDest As RECT
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * 32 * 0.5) + 32 * 0.5
            End If
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * 32) + 32
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .PixelWidth
        SourceRect.Bottom = SourceRect.Top + .PixelHeight
        
        Set Src = SurfaceDB.Surface(.FileNum)
        
        Src.GetSurfaceDesc ddsdSrc
        BackBufferSurface.GetSurfaceDesc ddsdDest
        
        rDest.Left = x
        rDest.Top = y
        rDest.Right = x + .PixelWidth
        rDest.Bottom = y + .PixelHeight
        
        If rDest.Right > ddsdDest.lWidth Then
            rDest.Right = ddsdDest.lWidth
        End If
        If rDest.Bottom > ddsdDest.lHeight Then
            rDest.Bottom = ddsdDest.lHeight
        End If
    End With
        
    Dim SrcLock As Boolean
    Dim DstLock As Boolean
        
On Error GoTo HayErrorAlpha

    If x < 0 Then
        Exit Sub
    End If
    
    If y < 0 Then
        Exit Sub
    End If
    
    Call Src.Lock(SourceRect, ddsdSrc, DDLOCK_WAIT, 0)
    SrcLock = True
    Call BackBufferSurface.Lock(rDest, ddsdDest, DDLOCK_WAIT, 0)
    DstLock = True
    
    Call BackBufferSurface.GetLockedArray(dArray())
    Call Src.GetLockedArray(sArray())
    
    Call BltAlphaFast(ByVal VarPtr(dArray(x + x, y)), ByVal VarPtr(sArray(SourceRect.Left + SourceRect.Left, SourceRect.Top)), rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 1)
    
    BackBufferSurface.Unlock rDest
    DstLock = False
    Src.Unlock SourceRect
    SrcLock = False
Exit Sub

HayErrorAlpha:
    'Grh.Started = 0
    'Grh.FrameCounter = 0
    'Grh.Loops = 0
    'Grh.Speed = 0
    'Grh.GrhIndex = 0

    BackBufferSurface.Unlock rDest
    Src.Unlock SourceRect
End Sub

Public Sub DrawGrhtoHdc(ByVal hdc As Long, ByVal GrhIndex As Integer, ByRef SourceRect As RECT, ByRef DestRect As RECT)
    'Draws a Grh's portion to the given area of any Device Context
    Call SurfaceDB.Surface(GrhData(GrhIndex).FileNum).BltToDC(hdc, SourceRect, DestRect)
End Sub

Public Sub DrawTransparentGrhtoHdc(ByVal dsthdc As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal GrhIndex As Integer, ByRef SourceRect As RECT)
'This method is SLOW... Don't use in a loop if you care about
'speed!
    Dim Color As Long
    Dim x As Long
    Dim y As Long
    Dim srchdc As Long
    Dim Surface As DirectDrawSurface7
    
    Set Surface = SurfaceDB.Surface(GrhData(GrhIndex).FileNum)
    
    srchdc = Surface.GetDC
    
    For x = SourceRect.Left To SourceRect.Right - 1
        For y = SourceRect.Top To SourceRect.Bottom - 1
            Color = GetPixel(srchdc, x, y)
            
            If Color <> vbBlack Then
                Call SetPixel(dsthdc, dstX + (x - SourceRect.Left), dstY + (y - SourceRect.Top), Color)
            End If
        Next y
    Next x
    
    Call Surface.ReleaseDC(srchdc)
End Sub

Public Sub DrawImageInPicture(ByRef PictureBox As PictureBox, ByRef Picture As StdPicture, ByVal X1 As Single, ByVal Y1 As Single, Optional Width1, Optional Height1, Optional X2, Optional Y2, Optional Width2, Optional Height2)
'Draw Picture in the PictureBox
    Call PictureBox.PaintPicture(Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2)
End Sub

Public Sub RenderScreen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'Renders everything to the viewport
    
    On Error Resume Next
    
    Dim x           As Integer     'Keeps track of where on map we are
    Dim y           As Integer     'Keeps track of where on map we are
    Dim ScreenMinX  As Integer  'Start Y pos on current screen
    Dim ScreenMaxX  As Integer  'End Y pos on current screen
    Dim ScreenMinY  As Integer  'Start X pos on current screen
    Dim ScreenMaxY  As Integer  'End X pos on current screen
    Dim MinX       As Integer  'Start Y pos on current map
    Dim MaxX        As Integer  'End Y pos on current map
    Dim MinY        As Integer  'Start X pos on current map
    Dim MaxY        As Integer  'End X pos on current map
    Dim ScreenX     As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY     As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset  As Integer
    Dim minYOffset  As Integer
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    
    'Figure out Ends and Starts of screen

    ScreenMinX = TileX - HalfWindowTileWidth
    ScreenMaxX = TileX + HalfWindowTileWidth
    ScreenMinY = TileY - HalfWindowTileHeight
    ScreenMaxY = TileY + HalfWindowTileHeight
        
    MinX = ScreenMinX - 9 'TileBufferSize
    MaxX = ScreenMaxX + 9 'TileBufferSize
    MinY = ScreenMinY - 9 'TileBufferSize
    MaxY = ScreenMaxY + 9 'TileBufferSize

    ScreenMinX = ScreenMinX - 1
    ScreenMaxX = ScreenMaxX + 1
    ScreenMinY = ScreenMinY - 1
    ScreenMaxY = ScreenMaxY + 1

    Dim map As Byte
    Dim MapX As Integer
    Dim MapY As Integer
        
    'Draw floor layer
    For y = ScreenMinY To ScreenMaxY
        For x = ScreenMinX To ScreenMaxX
                        
            map = UserMap
            MapX = x
            MapY = y
                        
            'If MapX < MinXBorder And MapY < MinYBorder Then
            '    map = map - 5
            '    MapX = MapX + 100
            '    MapY = MapY + 100
                
            'ElseIf MapX < MinXBorder And MapY > MaxYBorder Then
            '    map = map + 3
            '    MapX = MapX + 100
            '    MapY = MapY - 100
                
            'ElseIf MapX < MinXBorder Then
            '    map = map - 1
            '    MapX = MapX + 100
                
            'ElseIf MapY < MinYBorder Then
            '    map = map - 4
            '    MapY = MapY + 100
                
            'ElseIf MapX > MaxXBorder And MapY < MinYBorder Then
            '    map = map - 3
            '    MapX = MapX - 100
            '    MapY = MapY + 100
                
            'ElseIf MapX > MaxXBorder And MapY > MaxYBorder Then
            '    map = map + 5
            '    MapX = MapX - 100
            '    MapY = MapY - 100
                
            'ElseIf MapX > MaxXBorder Then
            '    map = map + 1
            '    MapX = MapX - 100
                
            'ElseIf MapY > MaxYBorder Then
            '    map = map + 4
            '    MapY = MapY - 100
            'End If
            
            'Layer 1
            If MapData(MapX, MapY).Graphic(1).GrhIndex > 1 Then
                Call DDrawGrhtoSurface(MapData(MapX, MapY).Graphic(1), _
                    (ScreenX - 1) * 32 + PixelOffsetX + TileBufferPixelOffsetX, _
                    (ScreenY - 1) * 32 + PixelOffsetY + TileBufferPixelOffsetY, _
                    0, 1)
            End If
            
            ScreenX = ScreenX + 1
        Next x
        
        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + ScreenMinX
        ScreenY = ScreenY + 1
    Next y
    
    'Draw floor layer 2
    ScreenY = minYOffset
    
    For y = MinY To MaxY
        ScreenX = minXOffset
        For x = MinX To MaxX
            
            map = UserMap
            MapX = x
            MapY = y
                        
            If MapX < MinXBorder And MapY < MinYBorder Then
                map = map - 5
                MapX = MapX + 100
                MapY = MapY + 100
                
            ElseIf MapX < MinXBorder And MapY > MaxYBorder Then
                map = map + 3
                MapX = MapX + 100
                MapY = MapY - 100
                
            ElseIf MapX < MinXBorder Then
                map = map - 1
                MapX = MapX + 100
                
            ElseIf MapY < MinYBorder Then
                map = map - 4
                MapY = MapY + 100
                
            ElseIf MapX > MaxXBorder And MapY < MinYBorder Then
                map = map - 3
                MapX = MapX - 100
                MapY = MapY + 100
                
            ElseIf MapX > MaxXBorder And MapY > MaxYBorder Then
                map = map + 5
                MapX = MapX - 100
                MapY = MapY - 100
                
            ElseIf MapX > MaxXBorder Then
                map = map + 1
                MapX = MapX - 100
                
            ElseIf MapY > MaxYBorder Then
                map = map + 4
                MapY = MapY - 100
            End If
            
            'Layer 2
            If MapData(MapX, MapY).Graphic(2).GrhIndex > 1 Then
                Call DDrawTransGrhtoSurface(MapData(MapX, MapY).Graphic(2), _
                    (ScreenX - 1) * 32 + PixelOffsetX, _
                    (ScreenY - 1) * 32 + PixelOffsetY, _
                    1, 1)
            End If
                        
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
    
    'Draw Transparent Layers
    ScreenY = minYOffset
    
    For y = MinY To MaxY
        ScreenX = minXOffset
        For x = MinX To MaxX
            
            map = UserMap
            MapX = x
            MapY = y
                        
            If MapX < MinXBorder And MapY < MinYBorder Then
                map = map - 5
                MapX = MapX + 100
                MapY = MapY + 100
                
            ElseIf MapX < MinXBorder And MapY > MaxYBorder Then
                map = map + 3
                MapX = MapX + 100
                MapY = MapY - 100
                
            ElseIf MapX < MinXBorder Then
                map = map - 1
                MapX = MapX + 100
                
            ElseIf MapY < MinYBorder Then
                map = map - 4
                MapY = MapY + 100
                
            ElseIf MapX > MaxXBorder And MapY < MinYBorder Then
                map = map - 3
                MapX = MapX - 100
                MapY = MapY + 100
                
            ElseIf MapX > MaxXBorder And MapY > MaxYBorder Then
                map = map + 5
                MapX = MapX - 100
                MapY = MapY - 100
                
            ElseIf MapX > MaxXBorder Then
                map = map + 1
                MapX = MapX - 100
                
            ElseIf MapY > MaxYBorder Then
                map = map + 4
                MapY = MapY - 100
            End If
            
            PixelOffsetXTemp = (ScreenX - 1) * 32 + PixelOffsetX
            PixelOffsetYTemp = (ScreenY - 1) * 32 + PixelOffsetY
            
            With MapData(MapX, MapY)

                'Object Layer
                If .Obj.Grh.GrhIndex > 0 Then
                
                    If AlphaBActivated Then
                        If .Obj.Grh.GrhIndex = GrhPortal Or _
                        (.Obj.ObjType = otPuerta And Charlist(UserCharIndex).Priv > 1) Then
                            Call DDrawTransGrhtoSurfaceAlpha(.Obj.Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                        Else
                            Call DDrawTransGrhtoSurface(.Obj.Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                        End If
                    Else
                        Call DDrawTransGrhtoSurface(.Obj.Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                    End If

                    If MouseTileX = MapX Then
                        If MouseTileY = MapY Then
                        
                            If MouseTileX > 9 And MouseTileX < 92 And MouseTileY > 7 And MouseTileY < 94 Then
                        
                                If .Obj.Amount > 0 Then
                                
                                    Call InitObjName(.Obj.Name, .Obj.ObjType, PixelOffsetXTemp + 30, PixelOffsetYTemp - 10)
                                    
                                    If AlphaBActivated Then
                                        Call SurfaceColor(.Obj.Grh, PixelOffsetXTemp, PixelOffsetYTemp, 225, 225, 50)
                                    End If
                                    
                                    If UsingSkill = 0 Then
                                        If MapX = UserPos.x And MapY = UserPos.y Then
                                            frmMain.MousePointer = 5
                                        End If
                                    End If
                                    
                                ElseIf .Obj.ObjType = otCuerpoMuerto Then
                                    Call InitObjName(.Obj.Name, .Obj.ObjType, PixelOffsetXTemp + 30, PixelOffsetYTemp - 10)
                                    
                                    If AlphaBActivated Then
                                        Call SurfaceColor(.Obj.Grh, PixelOffsetXTemp, PixelOffsetYTemp, 225, 225, 50)
                                    End If
                                
                                ElseIf .Obj.ObjType = otTeleport Then
                                    Call InitObjName(.Obj.Name, .Obj.ObjType, PixelOffsetXTemp + 30, PixelOffsetYTemp - 10)
                            
                                    If AlphaBActivated Then
                                        'If .Obj.Grh.Started = 2 Then
                                            Call SurfaceColor(.Obj.Grh, PixelOffsetXTemp, PixelOffsetYTemp, 225, 225, 50)
                                        'End If
                                    End If
                                Else
                                    Call RemoveObjName
                                End If
                            Else
                                Call RemoveObjName
                            End If
                        End If
                    End If
                Else
                    If MouseTileX = MapX And MouseTileY = MapY Then
                        Call RemoveObjName
                        
                        If UsingSkill = 0 Then
                            frmMain.MousePointer = vbDefault
                        End If
                    End If
                End If
                
                'Char layer
                If .CharIndex > 0 Then
                    If Charlist(.CharIndex).Pos.x <> MapX Or Charlist(.CharIndex).Pos.y <> MapY Then
                        .CharIndex = 0
                        'Call EraseChar(.CharIndex)
                        
                    Else
                        Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                        
                        If MouseTileX = MapX Then
                            If MouseTileY = MapY Then
                                If .CharIndex <> UserCharIndex Then
                                    If Charlist(.CharIndex).EsUser Then
                                        If Not Charlist(.CharIndex).Invisible Then
                                            If Charlist(.CharIndex).Priv < 2 Then
                                                Call InitCharName(.CharIndex)
                                            End If
                                        End If
                                    Else
                                        Call InitCharName(.CharIndex)
                                    End If
                                End If
                            End If
                        End If
                    End If
                ElseIf MouseTileX = MapX And MouseTileY = MapY Then
                    Call RemoveCharName
                End If
                
                'Layer 3 *
                If .Graphic(3).GrhIndex > 0 Then
                    If MapX > MinX And MapY > MinY And MapX < MaxX And MapY < MaxY Then
                        'Draw
                        If AlphaBActivated Then
                        
                            If Charlist(UserCharIndex).Priv > 1 Then
                                Call DDrawTransGrhtoSurfaceAlpha(.Graphic(3), _
                                    PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                            
                            ElseIf .Blocked And .CharIndex > 0 Then
                                Call DDrawTransGrhtoSurfaceAlpha(.Graphic(3), _
                                    PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                            
                            ElseIf GrhData(.Graphic(3).GrhIndex).FileNum > 5999 And GrhData(.Graphic(3).GrhIndex).FileNum < 6999 Then
                                If (Abs(UserPos.x - MapX) < 3 And Abs(UserPos.x - MapX) >= 0 And _
                                    UserPos.y - MapY < 0 And UserPos.y - MapY > -6) Or _
                                    Charlist(UserCharIndex).Priv > 1 Then
                                    'Abs(MouseTileX - MapX) < 3 And
                                    'MouseTileY - MapY < 0 And MouseTileY - MapY > -6)
                                    
                                    Call DDrawTransGrhtoSurfaceAlpha(.Graphic(3), _
                                        PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                                Else
                                    Call DDrawTransGrhtoSurface(.Graphic(3), _
                                        PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                                End If
                            
                            ElseIf MapX = MouseTileX Then
                            
                                If MapY = MouseTileY Then
                                    'If mapdata(MapX, MapY).CharIndex > 0 Then
                                    'If Not Charlist(mapdata(MapX, MapY).CharIndex).Invisible Then
                                    'Call DDrawTransGrhtoSurfaceAlpha(.Graphic(3), _
                                    'PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                                    'End If
                                        
                                    'ElseIf mapdata(X, MapY).Amount > 0 Or mapdata(MapX - 1, MapY).Amount > 0 Or mapdata(MapX - 2, MapY).Amount > 0 Then
                                    'Call DDrawTransGrhtoSurfaceAlpha(.Graphic(3), _
                                    'PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                                    'Else
                                        Call DDrawTransGrhtoSurface(.Graphic(3), _
                                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                                    'End If
                                Else
                                    Call DDrawTransGrhtoSurface(.Graphic(3), _
                                        PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                                End If
                                
                            Else
                                Call DDrawTransGrhtoSurface(.Graphic(3), _
                                    PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                            End If
                            
                        Else
                            Call DDrawTransGrhtoSurface(.Graphic(3), _
                                PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                        End If
                    End If
                End If
                
                If .FxIndex > 0 Then
                    If AlphaBActivated Then
                        Call DDrawTransGrhtoSurfaceAlpha(.fX, PixelOffsetXTemp + FxData(.FxIndex).OffsetX, PixelOffsetYTemp + FxData(.FxIndex).OffsetY, 1, 1)
                    Else
                        Call DDrawTransGrhtoSurface(.fX, PixelOffsetXTemp + FxData(.FxIndex).OffsetX, PixelOffsetYTemp + FxData(.FxIndex).OffsetY, 1, 1)
                    End If
                    
                    'Check if animation is over
                    If .fX.Started = 0 Then
                        .FxIndex = 0
                    End If
                End If
                
            End With
            
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
    
    'Draw blocked tiles and grid
    ScreenY = minYOffset
    
    For y = MinY To MaxY
        ScreenX = minXOffset
        For x = MinX To MaxX

            map = UserMap
            MapX = x
            MapY = y
                        
            If MapX < MinXBorder And MapY < MinYBorder Then
                map = map - 5
                MapX = MapX + 100
                MapY = MapY + 100
                
            ElseIf MapX < MinXBorder And MapY > MaxYBorder Then
                map = map + 3
                MapX = MapX + 100
                MapY = MapY - 100
                
            ElseIf MapX < MinXBorder Then
                map = map - 1
                MapX = MapX + 100
                
            ElseIf MapY < MinYBorder Then
                map = map - 4
                MapY = MapY + 100
                
            ElseIf MapX > MaxXBorder And MapY < MinYBorder Then
                map = map - 3
                MapX = MapX - 100
                MapY = MapY + 100
                
            ElseIf MapX > MaxXBorder And MapY > MaxYBorder Then
                map = map + 5
                MapX = MapX - 100
                MapY = MapY - 100
                
            ElseIf MapX > MaxXBorder Then
                map = map + 1
                MapX = MapX - 100
                
            ElseIf MapY > MaxYBorder Then
                map = map + 4
                MapY = MapY - 100
            End If
                        
            'Layer 4
            If MapData(MapX, MapY).Graphic(4).GrhIndex > 0 Then
                'Draw
                If AlphaBActivated And Not bTecho And Charlist(UserCharIndex).Priv < 2 Then
                    Call DDrawTransGrhtoSurfaceAlpha(MapData(MapX, MapY).Graphic(4), _
                        (ScreenX - 1) * 32 + PixelOffsetX, _
                        (ScreenY - 1) * 32 + PixelOffsetY, _
                        1, 0)
                
                ElseIf Not bTecho And Charlist(UserCharIndex).Priv < 2 Then
                    Call DDrawTransGrhtoSurface(MapData(MapX, MapY).Graphic(4), _
                        (ScreenX - 1) * 32 + PixelOffsetX, _
                        (ScreenY - 1) * 32 + PixelOffsetY, _
                        1, 0)
                End If
            End If
            
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y

'TODO : Check this!
    'If bLluvia(Map) = 1 Then
    'If bRain Then
            'Figure out what frame to draw
    'If llTick < DirectX.TickCount - 50 Then
    'iFrameIndex = iFrameIndex + 1
    'If iFrameIndex > 7 Then iFrameIndex = 0
    'llTick = DirectX.TickCount
    'End If

    'For Y = 0 To 4
    'For X = 0 To 4
    'Call BackBufferSurface.BltFast(LTLluvia(Y), LTLluvia(X), SurfaceDB.Surface(15168), RLluvia(iFrameIndex), DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
    'Next X
    'Next Y
    'End If
    'End If

    If MapInfo(UserMap).Zone <> "DUNGEON" Then
        If AlphaBActivated Then
            Select Case Tiempo
                Case 1
                    EfectoAmanecer BackBufferSurface
                'Case 2
                '    EfectoMañana BackBufferSurface
                'Case 3
                '    EfectoMediodía BackBufferSurface
                'Case 4
                '    EfectoTarde BackBufferSurface
                'Case 5
                '    EfectoAnochecer BackBufferSurface
                Case 6
                    EfectoNoche BackBufferSurface
            End Select
        End If
    End If
End Sub

Public Function RenderSounds()
'Actualiza todos los sonidos del mapa.

On Error Resume Next

    If bLluvia(UserMap) = 1 Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                    If RainBufferIndex Then
                        Call Audio.StopWave(RainBufferIndex)
                    End If
                    
                    RainBufferIndex = Audio.Play("lluviain.wav", 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviain
                End If
            Else
                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                    If RainBufferIndex Then
                        Call Audio.StopWave(RainBufferIndex)
                    End If
                    
                    RainBufferIndex = Audio.Play("lluviaout.wav", 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviaout
                End If
            End If
        End If
    End If
    
    Call DoPortalFx
    
    Call DoFogataFx
End Function

Public Sub LoadGraphics()
'Initializes the SurfaceDB and sets up the rain rects

    'New surface manager :D
    Call SurfaceDB.Initialize(DirectDraw, GrhPath, 64)
    
    'Set up te rain rects
    RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
    RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
    RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
    RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128
    
    RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
    RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
    RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
    RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
    
End Sub

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setMainViewTop As Integer, ByVal setMainViewLeft As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer, ByVal engineSpeed As Single) As Boolean
'Creates all DX objects and configures the engine to start running.

    Dim SurfaceDesc As DDSURFACEDESC2
    Dim ddck As DDCOLORKEY
    
    'Fill startup variables
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    HalfWindowTileHeight = setWindowTileHeight * 0.5
    HalfWindowTileWidth = setWindowTileWidth * 0.5
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    TileBufferPixelOffsetX = ((TileBufferSize - 1) * 32)
    TileBufferPixelOffsetY = ((TileBufferSize - 1) * 32)
    
    engineBaseSpeed = engineSpeed
    
    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    MinXBorder = 1 'XMinMapSize + (XWindow * 0.5)
    MaxXBorder = 100 * 5 'XMaxMapSize - (XWindow * 0.5)
    MinYBorder = 1 'YMinMapSize + (YWindow * 0.5)
    MaxYBorder = 100 * 5 'YMaxMapSize - (YWindow * 0.5)
    
    MainViewWidth = 32 * WindowTileWidth
    MainViewHeight = 32 * WindowTileHeight
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To 28) As MapInfoBlock
    
    'Set intial user position
    UserPos.x = MinXBorder
    UserPos.y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
    'Set the view rect
    With MainViewRect
        .Left = MainViewLeft
        .Top = MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    'Set the dest rect
    With MainDestRect
        .Left = 32 * TileBufferSize - 32
        .Top = 32 * TileBufferSize - 32
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
On Error Resume Next
    Set DirectX = New DirectX7
    
    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If

    
    'INIT DirectDraw
    'Create the root DirectDraw object
    Set DirectDraw = DirectX.DirectDrawCreate(vbNullString)
    
    If Err Then
        MsgBox "No se puede iniciar DirectDraw. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If
    
On Error GoTo 0
    Call DirectDraw.SetCooperativeLevel(setDisplayFormhWnd, DDSCL_NORMAL)
    
    'Primary Surface
    'Fill the surface description structure
    With SurfaceDesc
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With
    'Create the surface
    Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)
    
    'Create Primary Clipper
    Set PrimaryClipper = DirectDraw.CreateClipper(0)
    Call PrimaryClipper.SetHWnd(frmMain.hWnd)
    Call PrimarySurface.SetClipper(PrimaryClipper)
    
    With BackBufferRect
        .Left = 0
        .Top = 0
        .Right = 32 * (WindowTileWidth + 2 * TileBufferSize)
        .Bottom = 32 * (WindowTileHeight + 2 * TileBufferSize)
    End With
    
    With SurfaceDesc
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .lHeight = BackBufferRect.Bottom
        .lWidth = BackBufferRect.Right
    End With
    
    'Create surface
    Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)
    
    'Set color key
    ddck.low = 0
    ddck.high = 0
    Call BackBufferSurface.SetColorKey(DDCKEY_SRCBLT, ddck)
        
    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    
    LTLluvia(0) = 224
    LTLluvia(1) = 352
    LTLluvia(2) = 480
    LTLluvia(3) = 608
    LTLluvia(4) = 736
    
    Call LoadGraphics
    
    InitTileEngine = True
End Function

Public Sub DeinitTileEngine()
'Destroys all DX objects

On Error Resume Next
    Set PrimarySurface = Nothing
    Set PrimaryClipper = Nothing
    Set BackBufferSurface = Nothing
    
    Set DirectDraw = Nothing
    
    Set DirectX = Nothing
End Sub

Public Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)
'Updates the game's model and renders everything.

    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    
    'Set main view rectangle
    MainViewRect.Left = (DisplayFormLeft / Screen.TwipsPerPixelX) + MainViewLeft
    MainViewRect.Top = (DisplayFormTop / Screen.TwipsPerPixelY) + MainViewTop
    MainViewRect.Right = MainViewRect.Left + MainViewWidth
    MainViewRect.Bottom = MainViewRect.Top + MainViewHeight
    
    If EngineRun Then
        If UserMoving Then
            'Move screen Left and Right if needed
            If AddtoUserPos.x <> 0 Then
                
                If UserMuerto Then
                    OffsetCounterX = OffsetCounterX - (ScrollPixelsPerFrameX + ScrollPixelsPerFrameX) * AddtoUserPos.x * timerTicksPerFrame
                Else
                    OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.x * timerTicksPerFrame
                End If
                
                If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.x) Then
                    OffsetCounterX = 0
                    AddtoUserPos.x = 0
                    UserMoving = False
                End If
            End If
            
            'Move screen Up and Down if needed
            If AddtoUserPos.y <> 0 Then
                
                If UserMuerto Then
                    OffsetCounterY = OffsetCounterY - (ScrollPixelsPerFrameY + ScrollPixelsPerFrameY) * AddtoUserPos.y * timerTicksPerFrame
                Else
                    OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.y * timerTicksPerFrame
                End If
                
                If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.y = 0
                    UserMoving = False
                End If
            End If
        End If
        
        'Update mouse position within view area
        'Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
        
        'Update screen
        If UserCiego Then
            Call CleanViewPort
        Else
            Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
        End If
        
        'Dim con As Byte
        
        'With Consola
        'For con = 1 To 10
        'Call RenderText(260, 181 + (con * 15), .MensajeConsola(con), RGB(.Color_Red(con), .Color_Green(con), .Color_Blue(con)), frmMain.font)
        'Next con
        'End With
                
        If UserPasarNivel > 0 Then
            If frmMain.ImgExp.Width <> CInt(((UserExp * 0.01) / (UserPasarNivel * 0.01)) * 171) Then
                If frmMain.ImgExp.Width < CInt(((UserExp * 0.01) / (UserPasarNivel * 0.01)) * 171) Then
                    frmMain.ImgExp.Width = frmMain.ImgExp.Width + 1
                Else
                    If frmMain.ImgExp.Width > 1 Then
                        frmMain.ImgExp.Width = frmMain.ImgExp.Width - 2
                    Else
                        frmMain.ImgExp.Width = frmMain.ImgExp.Width - 1
                    End If
                End If
                 
                'frmMain.ExpLbl.Caption = UserExp & " / " & UserPasarNivel
            End If
        End If
                
        If UserMaxHP > 0 Then
            If frmMain.ImgHP.Width <> CInt(((UserMinHP * 0.01) / (UserMaxHP * 0.01)) * 88) Then
                If frmMain.ImgHP.Width < CInt(((UserMinHP * 0.01) / (UserMaxHP * 0.01)) * 88) Then
                    frmMain.ImgHP.Width = frmMain.ImgHP.Width + 1
                Else
                    frmMain.ImgHP.Width = frmMain.ImgHP.Width - 1
                End If
            End If
            
            frmMain.HPLbl.Caption = UserMinHP & " / " & UserMaxHP
        End If
        
        If UserMaxMan > 0 Then
            If frmMain.ImgMana.Width <> CInt(((UserMinMan * 0.01) / (UserMaxMan * 0.01)) * 88) Then
                If frmMain.ImgMana.Width < CInt(((UserMinMan * 0.01) / (UserMaxMan * 0.01)) * 88) Then
                    frmMain.ImgMana.Width = frmMain.ImgMana.Width + 1
                Else
                    frmMain.ImgMana.Width = frmMain.ImgMana.Width - 1
                End If
            End If
            
            frmMain.MANLbl.Caption = UserMinMan & " / " & UserMaxMan
        End If
        
        If UserMaxSTA > 0 Then
            If frmMain.ImgSta.Width <> CInt(((UserMinSTA * 0.01) / (UserMaxSTA * 0.01)) * 88) Then
                If frmMain.ImgSta.Width < CInt(((UserMinSTA * 0.01) / (UserMaxSTA * 0.01)) * 88) Then
                    frmMain.ImgSta.Width = frmMain.ImgSta.Width + 1
                Else
                    frmMain.ImgSta.Width = frmMain.ImgSta.Width - 1
                End If
            End If
            
            frmMain.STALbl.Caption = UserMinSTA & " / " & UserMaxSTA
        End If

        If FPSFLAG Then
            Call RenderFPS
        End If
        
        Call RenderExp
        Call RenderGld
        Call RenderDamage
        Call RenderCharHP
        Call RenderCharDamage
        Call DibujarCartel
        Call Dialogos.Render
        Call RenderObjName
        Call RenderCharName
        Call RenderCoord
        
        'Display front-buffer!
        Call PrimarySurface.Blt(MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT)
        
        'Limit FPS to 100 (an easy number higher than monitor's vertical refresh rates)
        While (DirectX.TickCount - fpsLastCheck) / 10 < FramesPerSecCounter
            Sleep 5
        Wend
        
        'FPS update
        If fpsLastCheck + 1000 < DirectX.TickCount Then
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            fpsLastCheck = DirectX.TickCount
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
                
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
    End If
    
End Sub

Public Sub RenderText(ByVal lngXPos As Integer, ByVal lngYPos As Integer, ByRef strText As String, ByVal lngColor As Long, ByRef font As StdFont)
    If LenB(strText) > 0 Then
        With BackBufferSurface
            Call .SetForeColor(vbBlack)
            Call .SetFont(font)
            Call .DrawText(lngXPos - 1, lngYPos, strText, False)
            Call .DrawText(lngXPos + 1, lngYPos, strText, False)
            Call .DrawText(lngXPos, lngYPos + 1, strText, False)
            Call .DrawText(lngXPos, lngYPos - 1, strText, False)
            
            Call .DrawText(lngXPos - 1, lngYPos - 1, strText, False)
            Call .DrawText(lngXPos - 1, lngYPos + 1, strText, False)
            Call .DrawText(lngXPos + 1, lngYPos + 1, strText, False)
            Call .DrawText(lngXPos + 1, lngYPos - 1, strText, False)
            
            Call .SetForeColor(lngColor)
            Call .DrawText(lngXPos, lngYPos, strText, False)
        End With
    End If
End Sub

Public Sub RenderTextCenteRed(ByVal lngXPos As Integer, ByVal lngYPos As Integer, ByRef strText As String, ByVal lngColor As Long, ByRef font As StdFont)
    Dim hdc As Long
    Dim Ret As size
    
    If LenB(strText) > 0 Then
        With BackBufferSurface
            Call .SetFont(font)
            
            'Get width of text once rendeRed
            hdc = .GetDC()
            Call GetTextExtentPoint32(hdc, strText, Len(strText), Ret)
            Call .ReleaseDC(hdc)
            
            lngXPos = lngXPos - Ret.cX * 0.5
            
            Call .SetForeColor(vbBlack)
            Call .SetFont(font)
            Call .DrawText(lngXPos, lngYPos - 1, strText, False)
            Call .DrawText(lngXPos, lngYPos + 1, strText, False)
            Call .DrawText(lngXPos - 1, lngYPos, strText, False)
            Call .DrawText(lngXPos + 1, lngYPos, strText, False)
            
            Call .DrawText(lngXPos - 1, lngYPos - 1, strText, False)
            Call .DrawText(lngXPos - 1, lngYPos + 1, strText, False)
            Call .DrawText(lngXPos + 1, lngYPos + 1, strText, False)
            Call .DrawText(lngXPos + 1, lngYPos - 1, strText, False)
            
            Call .SetForeColor(lngColor)
            Call .DrawText(lngXPos, lngYPos, strText, False)
        End With
    End If
End Sub

Public Sub RenderObjName()

    If LenB(ObjName) < 1 Then
        Exit Sub
    End If
    
    Dim Color As Long
    
    Select Case ObjType
        Case otGuita
            Color = RGB(255, 255, 200)
        
        Case otCasco, otEscudo, otArmadura, otArma, otArma
            Color = RGB(230, 230, 150)
        
        Case otLlave
            Color = RGB(220, 220, 100)
        
        Case otMineral, otLeña
            Color = RGB(200, 200, 0)
        
        Case otBarco
            Color = RGB(200, 200, 0)
        
        Case otPergamino
            Color = RGB(200, 255, 0)
        
        Case otPocion
            Color = RGB(200, 200, 0)
        
        Case otBebida, otBotellaLlena, otBotellaVacia, otUseOnce
            Color = RGB(150, 150, 100)
        
        Case otAnillo
            Color = RGB(100, 200, 50)
                                        
        Case otTeleport
            Color = RGB(200, 200, 255)
                           
        Case otCuerpoMuerto
            Color = RGB(225, 225, 225)
            
    End Select
    
    With BackBufferSurface
        Call .SetFont(frmMSG.font)
        Call .SetForeColor(RGB(20, 20, 50))
        Call .DrawText(ObjX - 1, ObjY, ObjName, False)
        Call .DrawText(ObjX, ObjY - 1, ObjName, False)
        Call .DrawText(ObjX + 1, ObjY, ObjName, False)
        Call .DrawText(ObjX, ObjY + 1, ObjName, False)
             
        Call .DrawText(ObjX - 1, ObjY - 1, ObjName, False)
        Call .DrawText(ObjX - 1, ObjY + 1, ObjName, False)
        Call .DrawText(ObjX + 1, ObjY + 1, ObjName, False)
        Call .DrawText(ObjX + 1, ObjY - 1, ObjName, False)
                    
        Call .SetForeColor(Color)
        Call .DrawText(ObjX, ObjY, ObjName, False)
    End With
End Sub

Public Sub InitObjName(ByVal Name As String, ByVal T As eObjType, ByVal x As Integer, ByVal y As Integer)
    
    If Name = ObjName Then
        If ObjX = x Then
            If ObjY = y Then
                If ObjType = T Then
                    Exit Sub
                End If
            End If
        End If
    End If
    
    ObjName = Name
    ObjType = T
    ObjX = x
    ObjY = y
End Sub

Public Sub RemoveObjName()
    If LenB(ObjName) > 0 Then
        ObjName = vbNullString
    End If
End Sub

Public Sub RenderCharName()

    If LenB(CharName) < 1 Then
        Exit Sub
    End If
    
    Dim x As Integer
    Dim y As Integer
    
    x = 260
    y = 260
    
    With BackBufferSurface
        Call .SetFont(frmCarp.font)
        Call .SetForeColor(RGB(20, 20, 50))
        Call .DrawText(x - 1, y, CharName, False)
        Call .DrawText(x, y - 1, CharName, False)
        Call .DrawText(x + 1, y, CharName, False)
        Call .DrawText(x, y + 1, CharName, False)
             
        Call .DrawText(x - 1, y - 1, CharName, False)
        Call .DrawText(x - 1, y + 1, CharName, False)
        Call .DrawText(x + 1, y + 1, CharName, False)
        Call .DrawText(x + 1, y - 1, CharName, False)
                    
        Call .SetForeColor(CharColor)
        Call .DrawText(x, y, CharName, False)
    End With
End Sub

Public Sub InitCharName(ByVal CharIndex As Integer)

    With Charlist(CharIndex)
    
        CharName = .Nombre
            
        If .EsUser Then
            If .Priv < 2 Then
                CharColor = RGB(230 - .Lvl, 230 - .Lvl, 100)
            Else
                CharColor = RGB(120, 210, 0)
            End If
            
            CharName = .Nombre
        
            If .Priv < 2 Then
                CharName = CharName & " (Nv. " & .Lvl & ")"
            End If
        
        Else
            Select Case .Lvl
                Case 1
                    Exit Sub
                Case 2
                    CharColor = RGB(200, 150, 85)
                Case 3
                    CharColor = RGB(200, 150, 85)
                Case 4
                    CharColor = RGB(200, 150, 85)
            End Select
            
            If .Lvl > 2 Then
                CharName = CharName & " (Nv. " & .Lvl - 1 & ")"
            End If
        End If
    
    End With
    
End Sub

Public Sub RemoveCharName()
    If LenB(CharName) > 0 Then
        CharName = vbNullString
    End If
End Sub

Public Sub RenderDamage()
    
    If LenB(Damage) < 1 Then
        Exit Sub
    End If
    
    If Val(Damage) < 1 And Damage <> "Fallás" Then
        Exit Sub
    End If
    
    If GetTickCount() - startTime > 2250 Then
        Damage = vbNullString
        Exit Sub
    End If
            
    If SUBe > 0 Then
        SUBe = SUBe - 1
    End If
    
    y = 408 + Charlist(UserCharIndex).Body.HeadOffset.y + SUBe
    
    Select Case Len(Damage)
        Case 1
            x = 524
        Case 2
            x = 521
        Case 3
            x = 516
        Case 6
            x = 512
    End Select
    
    If DamageType = 3 Or DamageType = 5 Then
        If Right$(Damage, 1) <> "!" Then
            Damage = Damage & "!"
            x = x + 2
        End If
    End If
    
    If DamageType > 3 And SUBe = 19 Then
        Call Audio.Play(SND_APU)
    End If
                
    With BackBufferSurface
        Call .SetForeColor(RGB(40, 15, 15))
        
        If DamageType < 3 Then
            Call .SetFont(frmMensaje.font)
        Else
            Call .SetFont(frmMapa.font)
        End If
        
        Call .DrawText(x + 1, y, Damage, False)
        Call .DrawText(x - 1, y, Damage, False)
        Call .DrawText(x, y + 1, Damage, False)
        Call .DrawText(x, y - 1, Damage, False)
        Call .DrawText(x + 1, y - 1, Damage, False)
        Call .DrawText(x - 1, y + 1, Damage, False)
        Call .DrawText(x + 1, y + 1, Damage, False)
        Call .DrawText(x - 1, y - 1, Damage, False)
                
        Call .DrawText(x - 2, y, Damage, False)
        Call .DrawText(x + 2, y, Damage, False)
                
        Select Case DamageType
            Case 2
                Call .SetForeColor(RGB(0, 180, 0))
                
                If AttackedCharIndex = UserCharIndex Then
                    CharMinHP = 0
                End If
                
            Case Is > 3
                Call .SetForeColor(RGB(115, 0, 0))
            Case Else
                Call .SetForeColor(RGB(190, 0, 0))
        End Select
                
        Call .DrawText(x, y, Damage, False)
    End With
End Sub

Public Sub InitDamage(ByVal D As String)
    Damage = D
    startTime = GetTickCount
    SUBe = 20
    
    Gld = 0
    Call Dialogos.RemoveDialog(UserCharIndex)
    
    If RightHandEqp.ObjType = otFlecha Then
        RightHandEqp.Amount = RightHandEqp.Amount - 1
        frmMain.lblRightHandEqp.Caption = RightHandEqp.Amount
    End If
End Sub

Public Sub RemoveDamage()
    Damage = vbNullString
End Sub

Public Sub InitCharDamage(ByVal x As Integer, ByVal y As Integer)
    CharX = x
    CharY = y
    
    If LenB(CharDamage) > 0 Then
        CharDamage2 = CharDamage
        CharDamage = vbNullString
        startTime4 = GetTickCount
        SUBe4 = 20
    End If
End Sub

Public Sub RenderCharDamage()

On Error Resume Next
    
    If AttackerCharIndex < 1 Then
        Exit Sub
    End If
    
    If GetTickCount() - startTime4 > 3000 Then
        CharDamage2 = vbNullString
        CharDamageType = 1
        Exit Sub
    End If
    
    If LenB(CharDamage2) < 1 Then
        Exit Sub
    End If
    
    If SUBe4 > 0 Then
        SUBe4 = SUBe4 - 1
        CharY = CharY + SUBe4
    End If
                
    If Charlist(AttackerCharIndex).Head.Head(Charlist(AttackerCharIndex).Heading).GrhIndex > 0 Then
        CharY = CharY - GrhData(Charlist(AttackerCharIndex).Body.Walk(1).GrhIndex).PixelHeight + Charlist(AttackerCharIndex).Body.HeadOffset.y
    Else
        CharY = CharY - GrhData(Charlist(AttackerCharIndex).Body.Walk(1).GrhIndex).PixelHeight
    End If
        
    If AttackedCharIndex = AttackerCharIndex Then
        If CharMinHP > 0 Then
            CharY = CharY - 8
        End If
    End If
    
    CharY = CharY + 8
        
    Select Case Len(CharDamage2)
        Case 1
            CharX = CharX + 12
        Case 2
            CharX = CharX + 8
        Case 3
            CharX = CharX + 4
        Case 5
            CharX = CharX + 1
    End Select
        
    With BackBufferSurface
        Call .SetForeColor(RGB(40, 15, 15))
        
        If CharDamageType < 3 Then
            Call .SetFont(frmMensaje.font)
        Else
            Call .SetFont(frmMapa.font)
        End If
            
        Call .DrawText(CharX + 1, CharY, CharDamage2, False)
        Call .DrawText(CharX - 1, CharY, CharDamage2, False)
        Call .DrawText(CharX, CharY + 1, CharDamage2, False)
        Call .DrawText(CharX, CharY - 1, CharDamage2, False)
              
        Call .DrawText(CharX - 1, CharY - 1, CharDamage2, False)
        Call .DrawText(CharX - 1, CharY + 1, CharDamage2, False)
        Call .DrawText(CharX + 1, CharY + 1, CharDamage2, False)
        Call .DrawText(CharX + 1, CharY - 1, CharDamage2, False)
        
        Select Case CharDamageType
            Case 2
                Call .SetForeColor(RGB(0, 180, 0))
            Case Else
                Call .SetForeColor(RGB(180, 0, 0))
        End Select
        
        Call .DrawText(CharX, CharY, CharDamage2, False)
    End With
End Sub

Public Sub RenderCharHP()
    
    If AttackedCharIndex = 0 Then
        Exit Sub
    End If
    
    If GetTickCount() - startTime > 5000 Then
        CharMinHP = 0
        Exit Sub
    End If
    
    If CharMinHP < 1 Then
        Exit Sub
    End If
    
    If Charlist(AttackedCharIndex).Head.Head(Charlist(AttackedCharIndex).Heading).GrhIndex > 0 Then
        CharY2 = CharY2 - GrhData(Charlist(AttackedCharIndex).Body.Walk(1).GrhIndex).PixelHeight + Charlist(AttackedCharIndex).Body.HeadOffset.y
    Else
        CharY2 = CharY2 - GrhData(Charlist(AttackedCharIndex).Body.Walk(1).GrhIndex).PixelHeight
    End If
    
    CharY2 = CharY2 + 18
    
    BackBufferSurface.SetForeColor RGB(80, 0, 0)
    BackBufferSurface.SetFillColor RGB(140, 0, 0)
    BackBufferSurface.SetFillStyle 0
    BackBufferSurface.DrawRoundedBox CharX2 + 1, CharY2, CharX2 + 31, CharY2 + 4, 0, 0
    
    BackBufferSurface.SetForeColor RGB(0, 50, 0)
    BackBufferSurface.SetFillColor RGB(0, 100, 0)
    
    If 0.3 * CharMinHP < 2 Then
        BackBufferSurface.DrawRoundedBox CharX2 + 1, CharY2, CharX2 + 3, CharY2 + 4, 0, 0
    ElseIf 0.3 * CharMinHP > 1 Then
        BackBufferSurface.DrawRoundedBox CharX2 + 1, CharY2, CharX2 + 1 + 30 * CharMinHP / 100, CharY2 + 4, 0, 0
    End If
        
    BackBufferSurface.SetForeColor vbBlack
    BackBufferSurface.SetFillStyle 1
    BackBufferSurface.DrawRoundedBox CharX2, CharY2 - 1, CharX2 + 32, CharY2 + 5, 4, 4
End Sub

Public Sub InitCharHP(ByVal x As Integer, ByVal y As Integer)
    CharX2 = x
    CharY2 = y
End Sub

Public Sub RenderExp()

    If LenB(Exp) < 1 Or Exp = "0" Then
        Exit Sub
    End If

    If (GetTickCount() - StartTime2) >= 1000 Then
        Exp = vbNullString
        Exit Sub
    End If
                
    If SUBe2 > 0 Then
        SUBe2 = SUBe2 - 1
    End If
    
    If LenB(Damage) < 1 Then
        Y2 = 408 + Charlist(UserCharIndex).Body.HeadOffset.y + SUBe2
    Else
        Y2 = 393 + Charlist(UserCharIndex).Body.HeadOffset.y + SUBe2
    End If
    
    X2 = 528 - 4 * Len(Exp)
    
    Select Case Len(Exp)
        Case 1
            X2 = 523
        Case 2
            X2 = 520
        Case 3
            X2 = 517
        Case 4
            X2 = 513
        Case 5
            X2 = 509
    End Select
    
    With BackBufferSurface
        Call .SetFont(frmSpawnList.font)
        Call .SetForeColor(RGB(20, 20, 50))
        Call .DrawText(X2 - 1, Y2, Exp, False)
        Call .DrawText(X2, Y2 - 1, Exp, False)
        Call .DrawText(X2 + 1, Y2, Exp, False)
        Call .DrawText(X2, Y2 + 1, Exp, False)
            
        Call .DrawText(X2 - 1, Y2 - 1, Exp, False)
        Call .DrawText(X2 - 1, Y2 + 1, Exp, False)
        Call .DrawText(X2 + 1, Y2 + 1, Exp, False)
        Call .DrawText(X2 + 1, Y2 - 1, Exp, False)
                    
        Call .SetForeColor(RGB(0, 90, 150))
        Call .DrawText(X2, Y2, Exp, False)
    End With
End Sub

Public Sub InitExp(ByVal E As String)
    Exp = E
    StartTime2 = GetTickCount
    SUBe2 = 30
   
    Gld = 0
    Call Dialogos.RemoveDialog(UserCharIndex)
End Sub

Public Sub RemoveExp()
   Exp = vbNullString
End Sub

Public Sub RenderGld()
    If Gld = 0 Then
        Exit Sub
    End If
    
    If GetTickCount() - StartTime3 >= 2000 Then
        Gld = 0
        Exit Sub
    End If
        
    If SUBe3 > 0 Then
        SUBe3 = SUBe3 - 1
    End If
    
    Y3 = 401 + Charlist(UserCharIndex).Body.HeadOffset.y + SUBe3
        
    Dim Gold As String
    
    If Gld > 0 Then
        Gold = "+ " & PonerPuntos(CStr(Gld))
        X3 = 520 - 4 * Len(CStr(Gld))
    Else
        Gold = "- " & PonerPuntos(CStr(-Gld))
        X3 = 520 - 4 * Len(CStr(-Gld))
    End If
    
    With BackBufferSurface
        Call .SetFont(frmMensaje.font)
        Call .SetForeColor(RGB(30, 30, 10))

        Call .DrawText(X3 - 1, Y3, Gold, False)
        Call .DrawText(X3, Y3 - 1, Gold, False)
        Call .DrawText(X3 + 1, Y3, Gold, False)
        Call .DrawText(X3, Y3 + 1, Gold, False)
                
        Call .DrawText(X3 - 1, Y3 - 1, Gold, False)
        Call .DrawText(X3 - 1, Y3 + 1, Gold, False)
        Call .DrawText(X3 + 1, Y3 + 1, Gold, False)
        Call .DrawText(X3 + 1, Y3 - 1, Gold, False)

        If Gld > 0 Then
            Call .SetForeColor(RGB(200, 160, 0))
            Call .DrawText(X3, Y3, Gold, False)
        Else
            Call .SetForeColor(RGB(140, 100, 0))
            Call .DrawText(X3, Y3, Gold, False)
        End If
    End With
End Sub

Public Sub InitGld(ByVal g As Long)
    Gld = g
    StartTime3 = GetTickCount
    SUBe3 = 20
   
    Damage = vbNullString
    ObjName = 0

    Call Dialogos.RemoveDialog(UserCharIndex)
End Sub

Public Sub RemoveGld()
    Gld = 0
End Sub

Public Sub RenderCoord()
    With BackBufferSurface
    
        Const posX = 259
        Const posY = 640
        
        .SetFont frmOpciones.font
        .SetForeColor RGB(30, 30, 30)
        .DrawText posX, posY, "X: " & UserPos.x & " Y: " & UserPos.y, False
        .DrawText posX + 1, posY + 1, "X: " & UserPos.x & " Y: " & UserPos.y, False
        .DrawText posX + 2, posY, "X: " & UserPos.x & " Y: " & UserPos.y, False
        .DrawText posX + 2, posY - 1, "X: " & UserPos.x & " Y: " & UserPos.y, False
        
        .DrawText posX, posY - 1, "X: " & UserPos.x & " Y: " & UserPos.y, False
        .DrawText posX, posY + 1, "X: " & UserPos.x & " Y: " & UserPos.y, False
        .DrawText posX + 2, posY + 1, "X: " & UserPos.x & " Y: " & UserPos.y, False
        .DrawText posX + 2, posY - 1, "X: " & UserPos.x & " Y: " & UserPos.y, False
    
        .DrawText posX, posY + 15, "Mapa: " & UserMap, False
        .DrawText posX + 1, posY + 14, "Mapa: " & UserMap, False
        .DrawText posX + 2, posY + 15, "Mapa: " & UserMap, False
        .DrawText posX + 1, posY + 16, "Mapa: " & UserMap, False
          
        .DrawText posX, posY + 14, "Mapa: " & UserMap, False
        .DrawText posX, posY + 16, "Mapa: " & UserMap, False
        .DrawText posX + 2, posY + 16, "Mapa: " & UserMap, False
        .DrawText posX + 2, posY + 14, "Mapa: " & UserMap, False
          
        .SetForeColor RGB(200, 190, 150)
        .DrawText posX + 1, posY, "X: " & UserPos.x & " Y: " & UserPos.y, False
        .DrawText posX + 1, posY + 15, "Mapa: " & UserMap, False
    End With
End Sub

Public Sub RenderFPS()
    With BackBufferSurface
        .SetFont frmOpciones.font

        Const posX = 775
        Const posY = 258

        .SetForeColor RGB(30, 30, 30)
        .DrawText posX - 1, posY, ModTileEngine.FPS, False
        .DrawText posX, posY - 1, ModTileEngine.FPS, False
        .DrawText posX + 1, posY, ModTileEngine.FPS, False
        .DrawText posX, posY + 1, ModTileEngine.FPS, False
        
        .DrawText posX - 1, posY - 1, ModTileEngine.FPS, False
        .DrawText posX - 1, posY + 1, ModTileEngine.FPS, False
        .DrawText posX + 1, posY + 1, ModTileEngine.FPS, False
        .DrawText posX + 1, posY - 1, ModTileEngine.FPS, False
                  
        .SetForeColor RGB(200, 190, 150)
        .DrawText posX, posY, ModTileEngine.FPS, False
    End With
End Sub

Private Function GetElapsedTime() As Single
'Gets the time that past since the last call

    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

On Error GoTo Errorcito

    Dim Moved As Boolean
    Dim Pos As Integer
    Dim Color As Long
    
    With Charlist(CharIndex)
    
        If .Heading < 1 Then
            Exit Sub
        End If
        
        If .Moving Then
            'If needed, move left and right
            If .ScrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.ScrollDirectionX) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos ObjNameloten al moverse! REVER!
                If .Body.Walk(.Heading).Speed > 0 Then
                    .Body.Walk(.Heading).Started = 1
                    .Arma.WeaponWalk(.Heading).Started = 1
                    .Escudo.ShieldWalk(.Heading).Started = 1
                End If
                
                'Char moved
                Moved = True
                
                'Check if we already got there
                If (Sgn(.ScrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.ScrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .ScrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .ScrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * 1 * Sgn(.ScrollDirectionY) * timerTicksPerFrame
                
                'Start animations
                If .Body.Walk(.Heading).Speed > 0 Then
                    .Body.Walk(.Heading).Started = 1
                    .Arma.WeaponWalk(.Heading).Started = 1
                    .Escudo.ShieldWalk(.Heading).Started = 1
                End If
                
                'Char moved
                Moved = True
                
                'Check if we already got there
                If (Sgn(.ScrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.ScrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .ScrollDirectionY = 0
                End If
            End If
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
                
        If .Head.Head(.Heading).GrhIndex > 0 Then
            If Not .Invisible Then
                'Call SurfaceSombra(BackBufferSurface, .Body.Walk(.Heading), PixelOffsetX + .Body.HeadOffset.X + 5, PixelOffsetY + .Body.HeadOffset.Y - 15, 1, 0)
                'http://blisse-games.com.ar/sombra-en-dx7-t1556.html
                If AlphaBActivated Then

                    'Draw Body
                    If .Body.Walk(.Heading).GrhIndex > 0 Then
                        If .iHead = CASPER_HEAD Then
                            Call DDrawTransGrhtoSurfaceAlpha(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                        ElseIf CharIndex = SelectedCharIndex Then
                            Call SurfaceColor(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 225, 225, 50)
                        ElseIf .Paralizado Then
                            Call SurfaceColor(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 200, 200, 255, True)
                        'ElseIf .Lvl = 3 Then
                        '    Call SurfaceColor(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 255, 230, 0)
                        'ElseIf .Lvl = 4 Then
                        '    Call SurfaceColor(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 255, 0, 0)
                        Else
                            Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                        End If
                    End If
                
                    'Draw Head
                    If .Head.Head(.Heading).GrhIndex > 0 Then
                        If .iHead = CASPER_HEAD Then
                            Call DDrawTransGrhtoSurfaceAlpha(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0)
                        ElseIf CharIndex = SelectedCharIndex Then
                            Call SurfaceColor(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 225, 225, 50)
                        ElseIf .Paralizado Then
                            Call SurfaceColor(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 200, 200, 255, True)
                        'ElseIf .Lvl = 3 Then
                        '    Call SurfaceColor(.Head.Head(.Heading), PixelOffsetX, PixelOffsetY, 255, 230, 0)
                        'ElseIf .Lvl = 4 Then
                        '    Call SurfaceColor(.Head.Head(.Heading), PixelOffsetX, PixelOffsetY, 255, 0, 0)
                        Else
                            Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0)
                        End If
                    
                    
                        'Draw Helmet
                        If .Casco.Head(.Heading).GrhIndex > 0 Then
                            If .Paralizado Then
                                Call SurfaceColor(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y - 34, 200, 200, 255, True)
                            Else
                                Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y - 34, 1, 0)
                            End If
                        End If
                        
                        'Draw Weapon
                        If .Arma.WeaponWalk(.Heading).GrhIndex > 0 Then
                            If .Paralizado Then
                                Call SurfaceColor(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 200, 200, 255, True)
                            Else
                                Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                            End If
                        End If
                        
                        'Draw Shield
                        If .Escudo.ShieldWalk(.Heading).GrhIndex > 0 Then
                            If .Paralizado Then
                                Call SurfaceColor(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 200, 200, 255, True)
                            Else
                                Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                            End If
                        End If
                    End If
                
                Else
                    'Draw Body
                    If .Body.Walk(.Heading).GrhIndex > 0 Then
                        Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                    End If
                
                    'Draw Head
                    Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0)
                    
                    'Draw Helmet
                    If .Casco.Head(.Heading).GrhIndex > 0 Then
                        Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y - 34, 1, 0)
                    End If
                        
                    'Draw Weapon
                    If .Arma.WeaponWalk(.Heading).GrhIndex > 0 Then
                        Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                    End If
                        
                    'Draw Shield
                    If .Escudo.ShieldWalk(.Heading).GrhIndex > 0 Then
                        Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                    End If
                End If
                    
                'Draw name over head
                If .EsUser Then
                    If Nombres Then
                        Pos = getTagPosition(.Nombre)
                        
                        If CharIndex = SelectedCharIndex Then
                            Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 30, .Nombre, RGB(255, 225, 50), frmCharge.font)
                        Else
                            If .Priv < 2 Then
                                'If .Lvl > 14 Then
                                    Color = RGB(230 - .Lvl * 2.5, 230 - .Lvl * 2.5, 100)
                                'Else
                                '    Color = RGB(230, 230, 100)
                                'End If
                            Else
                                Color = RGB(120, 210, 0)
                            End If
                                    
                            'Nick
                            Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 30, .Nombre, Color, frmCharge.font)
                        End If
                        
                        'Guilda
                        If LenB(.Guilda) > 0 Then
                            If CharIndex = SelectedCharIndex Then
                                Color = RGB(255, 225, 50)
                            Else
                                
                                Select Case .AlineacionGuilda
                                    Case 1
                                        Color = RGB(120, 0, 0)
                                    Case 2
                                        Color = RGB(255, 80, 80)
                                    Case 3
                                        Color = RGB(167, 167, 167)
                                    Case 4
                                        Color = RGB(0, 0, 150)
                                    Case 5
                                        Color = RGB(0, 70, 150)
                                    Case 6
                                        Color = RGB(120, 255, 120)
                                End Select
                            End If
                            
                            Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 45, "<" & .Guilda & ">", Color, frmCharge.font)
                        End If
                    End If
                    
                ElseIf CharIndex = SelectedCharIndex Then
                    Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 32, .Nombre, RGB(255, 225, 50), frmCharge.font)
                
                'ElseIf Abs(MouseTileX - .Pos.x) < 1 And Abs(MouseTileY - .Pos.y) < 1 Then
                
                '    Select Case .Lvl
                '        Case 1
                '            Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 32, .Nombre, RGB(220, 220, 220), frmCharge.font)
                '        Case 2
                '            Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 32, .Nombre, RGB(200, 150, 120), frmCharge.font)
                '        Case 3
                '            Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 32, .Nombre & " (Nv. 2)", RGB(200, 150, 85), frmCharge.font)
                '        Case 4
                '            Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 32, .Nombre & " (Nv. 3)", RGB(200, 150, 30), frmCharge.font)
                '    End Select
                End If
                
            'ALPHA SI ESTA INVISIBLE
            ElseIf CharIndex = UserCharIndex Or _
                LenB(.Guilda) > 0 And .Guilda = Charlist(UserCharIndex).Guilda Or _
                (Charlist(UserCharIndex).Priv > 1) Then
                                                              
                'CON ALPHABACTIVATED
                If AlphaBActivated Then
                
                    'Draw Body
                    If .Body.Walk(.Heading).GrhIndex > 0 Then
                        Call DDrawTransGrhtoSurfaceAlpha(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                    End If
                
                    'Draw Head
                    If .Head.Head(.Heading).GrhIndex > 0 Then
                        Call DDrawTransGrhtoSurfaceAlpha(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0)
                    
                        'Draw Helmet
                        If .Casco.Head(.Heading).GrhIndex > 0 Then
                            Call DDrawTransGrhtoSurfaceAlpha(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y - 34, 1, 0)
                        End If
                            
                        'Draw Weapon
                        If .Arma.WeaponWalk(.Heading).GrhIndex > 0 Then
                            Call DDrawTransGrhtoSurfaceAlpha(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                        End If
                        
                        'Draw Shield
                        If .Escudo.ShieldWalk(.Heading).GrhIndex > 0 Then
                            Call DDrawTransGrhtoSurfaceAlpha(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                        End If
                    Else
                        Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0)
                    End If
                    
                'SIN ALPHABACTIVATED
                Else
                    'Draw Body
                    If .Body.Walk(.Heading).GrhIndex > 0 Then
                        Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                    End If
                
                    'Draw Head
                    Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0)
                    
                    'Draw Helmet
                    If .Casco.Head(.Heading).GrhIndex > 0 Then
                        Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y - 34, 1, 0)
                    End If
                        
                    'Draw Weapon
                    If .Arma.WeaponWalk(.Heading).GrhIndex > 0 Then
                        Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                    End If
                        
                    'Draw Shield
                    If .Escudo.ShieldWalk(.Heading).GrhIndex > 0 Then
                        Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                    End If
                End If

                'Draw name over head
                If Nombres Then
                    Pos = getTagPosition(.Nombre)
                
                    If .Priv < 2 Then
                        'If .Lvl > 14 Then
                            Color = RGB(230 - .Lvl * 2.5, 230 - .Lvl * 2.5, 100)
                        'Else
                        '    Color = RGB(230, 230, 100)
                        'End If
                    Else
                        Color = RGB(120, 210, 0)
                    End If
                            
                    'Nick
                    Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 30, .Nombre, Color, frmCharge.font)

                    'Guilda
                    If LenB(.Guilda) > 0 Then
                        Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 45, "<" & .Guilda & ">", &HC0FFFF, frmCharge.font)
                    End If
                End If
            End If

        'Draw Body
        ElseIf .Body.Walk(.Heading).GrhIndex > 0 Then
  
            If AlphaBActivated Then
                If CharIndex = SelectedCharIndex Then
                    Call SurfaceColor(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 225, 225, 50)
                ElseIf .Paralizado Then
                    Call SurfaceColor(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 200, 200, 255, True)
                'ElseIf .Lvl = 3 Then
                '    Call SurfaceColor(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 230, 150, 0)
                'ElseIf .Lvl = 4 Then
                '    Call SurfaceColor(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 255, 70, 0)
                Else
                    Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                End If
            Else
                Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
            End If
            
            'Draw name over head
            If Nombres Then
                If .EsUser Then
                    Pos = getTagPosition(.Nombre)
                                  
                    If .Priv < 2 Then
                        'If .Lvl > 14 Then
                            Color = RGB(230 - .Lvl * 2.5, 230 - .Lvl * 2.5, 100)
                        'Else
                        '    Color = RGB(230, 230, 100)
                        'End If
                    Else
                        Color = RGB(120, 210, 0)
                    End If
                         
                    Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 30, .Nombre, Color, frmCharge.font)
                        
                    'Guilda
                    If LenB(.Guilda) > 0 Then
                        Select Case .AlineacionGuilda
                            Case 1
                                Color = RGB(120, 0, 0)
                            Case 2
                                Color = RGB(255, 80, 80)
                            Case 3
                                Color = RGB(167, 167, 167)
                            Case 4
                                Color = RGB(0, 0, 150)
                            Case 5
                                Color = RGB(0, 70, 150)
                            Case 6
                                Color = RGB(120, 255, 120)
                        End Select
                        
                        Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 45, "<" & .Guilda & ">", Color, frmCharge.font)
                    End If
                    
                'NOMBRE DE NPC
                ElseIf CharIndex = SelectedCharIndex Then
                    Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 32, .Nombre, RGB(255, 225, 50), frmCharge.font)
                
                'ElseIf Abs(MouseTileX - .Pos.x) < 1 And Abs(MouseTileY - .Pos.y) < 1 Then
                '    Select Case .Lvl
                '        Case 1
                '            Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 32, .Nombre, RGB(220, 220, 220), frmCharge.font)
                '        Case 2
                '            Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 32, .Nombre, RGB(200, 150, 120), frmCharge.font)
                '        Case 3
                '            Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 32, .Nombre & " (Nv. 2)", RGB(200, 150, 85), frmCharge.font)
                '        Case 4
                '            Call RenderTextCenteRed(PixelOffsetX + 32 * 0.5 + 5, PixelOffsetY + 32, .Nombre & " (Nv. 3)", RGB(200, 150, 30), frmCharge.font)
                '    End Select
                End If
            End If
        End If
        
        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y - 34, CharIndex)
        
        'Draw FX
        If .FxIndex > 0 Then
            If AlphaBActivated Then
                Call DDrawTransGrhtoSurfaceAlpha(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, 1)
            Else
                Call DDrawTransGrhtoSurface(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, 1)
            End If
            
            'Check if animation is over
            If .fX.Started < 1 Then
                .FxIndex = 0
            End If
        End If
                        
        If CharIndex = AttackedCharIndex Then
            If CharIndex <> UserCharIndex Then
                Call InitCharHP(PixelOffsetX, PixelOffsetY)
            End If
        End If
        
        If CharIndex = AttackerCharIndex Then
            Call InitCharDamage(PixelOffsetX, PixelOffsetY)
        End If
        
    End With
 
Errorcito:
End Sub

Private Sub CleanViewPort()
'Fills the viewport with black.

    Dim r As RECT
    Call BackBufferSurface.BltColorFill(r, vbBlack)
End Sub

Public Sub SurfaceColor(Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal r As Byte, ByVal g As Byte, ByVal B As Byte, Optional ByVal Paralized As Boolean = False)

    Dim iGrhIndex As Integer
    Dim SourceRect As RECT
     
    If Grh.GrhIndex = 0 Then
        Exit Sub
    End If
     
    If Not Paralized And Grh.Started = 1 Then
        Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
                
        If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
            Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                    
            If Grh.Loops <> INFINITE_LOOPS Then
                If Grh.Loops > 0 Then
                    Grh.Loops = Grh.Loops - 1
                Else
                    Grh.Started = 0
                End If
            End If
        End If
    End If
    
    iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
     
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        y = y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
     
    With SourceRect
        .Left = GrhData(iGrhIndex).sX + IIf(x < 0, Abs(x), 0)
        .Top = GrhData(iGrhIndex).sY + IIf(y < 0, Abs(y), 0)
        .Right = .Left + GrhData(iGrhIndex).PixelWidth
        .Bottom = .Top + GrhData(iGrhIndex).PixelHeight
    End With
     
    Dim Src As DirectDrawSurface7
    Dim rDest As RECT
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
     
    Set Src = SurfaceDB.Surface(GrhData(iGrhIndex).FileNum)
     
    Src.GetSurfaceDesc ddsdSrc
    BackBufferSurface.GetSurfaceDesc ddsdDest
    
    With rDest
        .Left = x
        .Top = y
        .Right = x + GrhData(iGrhIndex).PixelWidth
        .Bottom = y + GrhData(iGrhIndex).PixelHeight
       
        If .Right > ddsdDest.lWidth Then
            .Right = ddsdDest.lWidth
        End If
        If .Bottom > ddsdDest.lHeight Then
            .Bottom = ddsdDest.lHeight
        End If
    End With
     
    Dim SrcLock As Boolean, DstLock As Boolean
     
On Error GoTo HayErrorAlpha
     
    If x < 0 Then
        Exit Sub
    End If
    
    If y < 0 Then
        Exit Sub
    End If
    
    Src.Lock SourceRect, ddsdSrc, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
    BackBufferSurface.Lock rDest, ddsdDest, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
     
    BackBufferSurface.GetLockedArray dArray()
    Src.GetLockedArray sArray()
           
    Call vbDABLcolorblend16565ck(ByVal VarPtr(sArray(SourceRect.Left + SourceRect.Left, SourceRect.Top)), ByVal VarPtr(dArray(x + x, y)), 65, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, r, g, B)
    BackBufferSurface.Unlock rDest
    Src.Unlock SourceRect
     
    Exit Sub
     
HayErrorAlpha:
        'Grh.Started = 0
        'Grh.FrameCounter = 0
        'Grh.Loops = 0
        'Grh.Speed = 0
        'Grh.GrhIndex = 0
                        
        BackBufferSurface.Unlock rDest
        Src.Unlock SourceRect
End Sub
 
Public Sub EfectoAmanecer(ByRef Surface As DirectDrawSurface7)
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
    Dim rRect As RECT
     
    Surface.GetSurfaceDesc ddsdDest
     
    With rRect
    .Left = 0
    .Top = 0
    .Right = ddsdDest.lWidth
    .Bottom = ddsdDest.lHeight
    End With
     
    If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
        Modo = 0
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
        Modo = 1
    Else
        Modo = 2
    End If
     
    Dim DstLock As Boolean
    DstLock = False
    
On Error GoTo HayErrorAlpha
    
    Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
    DstLock = True
     
    Surface.GetLockedArray dArray()
     
    Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), 70, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, 150, 80, 120)
    
HayErrorAlpha:
     
    If DstLock Then
        Surface.Unlock rRect
        DstLock = False
    End If
End Sub

Public Sub EfectoMañana(ByRef Surface As DirectDrawSurface7)
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
    Dim rRect As RECT
     
    Surface.GetSurfaceDesc ddsdDest
     
    With rRect
    .Left = 0
    .Top = 0
    .Right = ddsdDest.lWidth
    .Bottom = ddsdDest.lHeight
    End With
     
    If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 0
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 1
    Else
    Modo = 2
    End If
     
    Dim DstLock As Boolean
    DstLock = False
      
    On Error GoTo HayErrorAlpha
    
    Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
    DstLock = True
     
    Surface.GetLockedArray dArray()
     
    Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), 70, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, 120, 110, 110)

HayErrorAlpha:
     
    If DstLock Then
        Surface.Unlock rRect
        DstLock = False
    End If
End Sub


Public Sub EfectoMediodía(ByRef Surface As DirectDrawSurface7)
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
    Dim rRect As RECT
     
    Surface.GetSurfaceDesc ddsdDest
     
    With rRect
    .Left = 0
    .Top = 0
    .Right = ddsdDest.lWidth
    .Bottom = ddsdDest.lHeight
    End With
     
    If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 0
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 1
    Else
    Modo = 2
    End If
     
    Dim DstLock As Boolean
    DstLock = False
      
    On Error GoTo HayErrorAlpha
    
    Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
    DstLock = True
     
    Surface.GetLockedArray dArray()
     
    Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), 70, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, 125, 120, 120)

HayErrorAlpha:
     
    If DstLock Then
        Surface.Unlock rRect
        DstLock = False
    End If

End Sub

Public Sub EfectoTarde(ByRef Surface As DirectDrawSurface7)
 
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
    Dim rRect As RECT
     
    Surface.GetSurfaceDesc ddsdDest
     
    With rRect
    .Left = 0
    .Top = 0
    .Right = ddsdDest.lWidth
    .Bottom = ddsdDest.lHeight
    End With
     
    If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 0
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 1
    Else
    Modo = 2
    End If
     
    Dim DstLock As Boolean
    DstLock = False
    
    Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
    DstLock = True
     
    On Error GoTo HayErrorAlpha
    
    Surface.GetLockedArray dArray()
     
    Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), 70, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, 150, 100, 80)

HayErrorAlpha:
     
    If DstLock Then
        Surface.Unlock rRect
        DstLock = False
    End If

End Sub

Public Sub EfectoAnochecer(ByRef Surface As DirectDrawSurface7)
 
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
    Dim rRect As RECT
     
    Surface.GetSurfaceDesc ddsdDest
     
    With rRect
    .Left = 0
    .Top = 0
    .Right = ddsdDest.lWidth
    .Bottom = ddsdDest.lHeight
    End With
     
    If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 0
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 1
    Else
    Modo = 2
    End If
     
    Dim DstLock As Boolean
    DstLock = False
    
    On Error GoTo HayErrorAlpha
    
    Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
    DstLock = True
     
    Surface.GetLockedArray dArray()
     
    Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), 70, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, 90, 50, 50)

HayErrorAlpha:
     
    If DstLock Then
        Surface.Unlock rRect
        DstLock = False
    End If

End Sub

  
Public Sub EfectoNoche(ByRef Surface As DirectDrawSurface7)
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
    Dim rRect As RECT
     
    Surface.GetSurfaceDesc ddsdDest
     
    With rRect
    .Left = 0
    .Top = 0
    .Right = ddsdDest.lWidth
    .Bottom = ddsdDest.lHeight
    End With
     
    If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 0
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 1
    Else
    Modo = 2
    End If
     
    Dim DstLock As Boolean
    DstLock = False
    
    On Error GoTo HayErrorAlpha
    
    Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
    DstLock = True
     
    Surface.GetLockedArray dArray()
    Call BltEfectoNoche(ByVal VarPtr(dArray(0, 0)), _
    ddsdDest.lWidth, ddsdDest.lHeight, ddsdDest.lPitch, _
    Modo)
    
HayErrorAlpha:
     
    If DstLock Then
        Surface.Unlock rRect
        DstLock = False
    End If

End Sub

Public Sub SurfaceSombra(Surface As DirectDrawSurface7, Grh As Grh, ByVal x As Integer, ByVal y As Integer, Center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
 
    On Error GoTo Errorcito
     
    Dim iGrhIndex As Integer
    Dim SourceRect As RECT
    Dim QuitarAnimacion As Boolean
     
    If Animate Then
        If Grh.Started = 1 Then
            If Grh.Speed > 0 Then
                Grh.Speed = Grh.Speed - 1
                If Grh.Speed = 0 Then
                    Grh.Speed = GrhData(Grh.GrhIndex).Speed
                    Grh.FrameCounter = Grh.FrameCounter + 1
                    If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                        Grh.FrameCounter = 1
                        If KillAnim Then
                            If Charlist(KillAnim).fX.Loops <> LoopAdEternum Then
     
                                If Charlist(KillAnim).fX.Loops > 0 Then Charlist(KillAnim).fX.Loops = Charlist(KillAnim).fX.Loops - 1
                                If Charlist(KillAnim).fX.Loops < 1 Then 'Matamos la anim del fx )
                                    Charlist(KillAnim).fX.Loops = 0
                                    Exit Sub
                                End If
     
                            End If
                        End If
                   End If
                End If
            End If
        End If
    End If
     
    If Grh.GrhIndex = 0 Then
        Exit Sub
    End If
    
    iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
     
    If Center Then
        If GrhData(iGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
        End If
        If GrhData(iGrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
        End If
    End If
     
    With SourceRect
        .Left = GrhData(iGrhIndex).sX + IIf(x < 0, Abs(x), 0)
        .Top = GrhData(iGrhIndex).sY + IIf(y < 0, Abs(y), 0)
        .Right = .Left + GrhData(iGrhIndex).PixelWidth
        .Bottom = .Top + GrhData(iGrhIndex).PixelHeight
    End With
     
    Dim Src As DirectDrawSurface7
    Dim rDest As RECT
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
     
    Set Src = SurfaceDB.Surface(GrhData(iGrhIndex).FileNum)
     
    Src.GetSurfaceDesc ddsdSrc
    Surface.GetSurfaceDesc ddsdDest
    
    With rDest
        .Left = x
        .Top = y
        .Right = x + GrhData(iGrhIndex).PixelWidth
        .Bottom = y + GrhData(iGrhIndex).PixelHeight
       
        If .Right > ddsdDest.lWidth Then
            .Right = ddsdDest.lWidth
        End If
        If .Bottom > ddsdDest.lHeight Then
            .Bottom = ddsdDest.lHeight
        End If
    End With
     
    Dim SrcLock As Boolean, DstLock As Boolean
     
    On Error GoTo Errorcito
     
    Src.Lock SourceRect, ddsdSrc, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
    Surface.Lock rDest, ddsdDest, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
     
    Surface.GetLockedArray dArray()
    Src.GetLockedArray sArray()
           
     
    Call vbDABLcolorblend16565ck(ByVal VarPtr(sArray(SourceRect.Left + SourceRect.Left, SourceRect.Top)), ByVal VarPtr(dArray(x + x, y)), 255, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 50, 50, 50)
     
    Surface.Unlock rDest
    Src.Unlock SourceRect
    Exit Sub
     
Errorcito:
    Src.Unlock SourceRect
    Surface.Unlock rDest
End Sub

Public Sub GenerarMiniMapa()
    Dim SR As RECT
    
    SR.Left = MapInfo(UserMap).Left * 204
    SR.Top = MapInfo(UserMap).Top * 201
    SR.Right = SR.Left + 204 + 32 * 2
    SR.Bottom = SR.Top + 201 + 32 * 2

    SupBMiniMap.Blt SR, SurfaceDB.Surface(GrhData(4).FileNum), SR, DDBLT_DONOTWAIT
    
End Sub

Public Sub DibujarMiniMapa()

    Dim i As Integer
    Dim DR As RECT
    
    Dim CurrentGrhIndex As Integer

    DR.Top = UserPos.y * 2 - 16 * 2
    DR.Left = UserPos.x * 2 - 16 * 2
    DR.Bottom = DR.Top + 201 + 16 * 2
    DR.Right = DR.Left + 204 + 16 * 2
    
    SupMiniMap.BltFast 1, 1, SupBMiniMap, DR, DDBLTFAST_WAIT
    
    'For i = 1 To LastChar
    
    Dim MinX As Integer
    Dim MinY As Integer
    Dim MaxX As Integer
    Dim MaxY As Integer
    
    MinX = UserPos.x - 35
    MaxX = UserPos.x + 35
    MinY = UserPos.y - 35
    MaxY = UserPos.y + 35
    
    If MinX < MinXBorder Then
        MinX = MinXBorder
    End If

    If MaxX > MaxXBorder Then
        MaxX = MaxXBorder
    End If
    
    If MinY < MinYBorder Then
        MinY = MinYBorder
    End If
    
    If MaxY > MaxYBorder Then
        MaxY = MaxYBorder
    End If
    
    For x = MinX To MaxX
        For y = MinY To MaxY

            If MapData(x, y).CharIndex > 0 Then
            
                i = MapData(x, y).CharIndex
                
                With Charlist(i)
                    
                    'If .Active Then
                        If .EsUser Then
                                            
                            'SI SOY YO
                            If i = UserCharIndex Then
                                DR.Left = 58
                                DR.Top = 58
                                DR.Bottom = DR.Top + 2
                                DR.Right = DR.Left + 2
                                
                                SupMiniMap.BltColorFill DR, &HC0FFFF
                            
                            'SI ESTA INVISIBLE
                            ElseIf .Invisible > 0 Then
                                'SI SOY GM LO VEO
                                If Charlist(UserCharIndex).Priv > 1 Then
                                    DR.Left = 58 - (UserPos.x - x) * 2
                                    DR.Top = 58 - (UserPos.y - y) * 2
                                    DR.Bottom = DR.Top + 2
                                    DR.Right = DR.Left + 2
                                    
                                    SupMiniMap.BltColorFill DR, vbBlack
                                End If
                            
                            'SI ES GM
                            ElseIf .Priv > 1 Then
                                DR.Left = 58 - (UserPos.x - x) * 2
                                DR.Top = 58 - (UserPos.y - y) * 2
                                DR.Bottom = DR.Top + 2
                                DR.Right = DR.Left + 2
                                
                                SupMiniMap.BltColorFill DR, vbGreen
                            
                            'USER COMUN, LO VEO SI SOY GM (AGRE GAR Items O SPELL PARA VER USERS DEL AREA, Y OTRO PARA VER INVIS TMB)
                            ElseIf Charlist(UserCharIndex).Priv > 1 Then
                                DR.Left = 58 - (UserPos.x - x) * 2
                                DR.Top = 58 - (UserPos.y - y) * 2
                                DR.Bottom = DR.Top + 2
                                DR.Right = DR.Left + 2
                                
                                SupMiniMap.BltColorFill DR, &HC0FFFF
                            End If
                        
                        'ES NPC
                        ElseIf .Lvl > 1 Then
                            DR.Left = 58 - (UserPos.x - x) * 2
                            DR.Top = 58 - (UserPos.y - y) * 2
                            DR.Bottom = DR.Top + 2
                            DR.Right = DR.Left + 2
        
                            'SupMiniMap.BltColorFill DR, RGB(220, 220, 220)
                            'SupMiniMap.BltColorFill DR, RGB(255, 255, 100)
                            SupMiniMap.BltColorFill DR, RGB(0, 255, 255)
                        
                        'NO ES HOSTIL
                        Else
                            DR.Left = 58 - (UserPos.x - x) * 2
                            DR.Top = 58 - (UserPos.y - y) * 2
                            DR.Bottom = DR.Top + 2
                            DR.Right = DR.Left + 2
                            
                            'SupMiniMap.BltColorFill DR, RGB(200, 150, 120)
                            SupMiniMap.BltColorFill DR, RGB(0, 150, 150)
                        End If
                    'End If
               
                End With
           
            End If
                
        Next
    Next

    DR.Left = 0
    DR.Top = 0
    DR.Bottom = 300
    DR.Right = 300
    
    SupMiniMap.BltToDC frmMain.Minimap.hdc, DR, DR
    
    frmMain.Minimap.Refresh
    
End Sub



'--------Public function RotateSurface-----
'RotateSurface , takes a surface to rotate, the surface to draw the rotation on, the angle to rotate on,
'the x and y destination to draw on, an optional transparency value
'and optionally it can return the width and height of the surface
Public Function RotateSurface(lngAngle As Long, XDest As Long, YDest As Long, Optional rgbTransparency As Long = -1, Optional ByRef Width As Long, Optional ByRef Height As Long)

On Error Resume Next

Dim pi As Double
pi = 3.141592654
    Dim surfSource As DirectDrawSurface7
    Dim surfDestination As DirectDrawSurface7
    
    Set surfSource = SurfaceDB.Surface(5) 'GrhData(iGrhIndex).FileNum
    
    
        Set surfDestination = SurfaceDB.Surface(5)
    
    Dim ddsdOriginal As DDSURFACEDESC2, ddsdDestination As DDSURFACEDESC2
    Dim iX As Long, iY As Long
    Dim iXDest As Long, iYDest As Long
    Dim rEmpty As RECT, rEmpty2 As RECT
    Dim sngA As Single, SinA As Single, CosA As Single
    Dim dblRMax As Long
    Dim lngXO As Long, lngYO As Long
    Dim lngColor As Long
    Dim lWidth As Long, lHeight As Long
    
    sngA = lngAngle * pi / 180
    SinA = Sin(sngA)
    CosA = Cos(sngA)
    
    surfSource.GetSurfaceDesc ddsdOriginal
    lWidth = ddsdOriginal.lWidth
    lHeight = ddsdOriginal.lHeight
    dblRMax = Sqr(lWidth ^ 2 + lHeight ^ 2)
    surfSource.Lock rEmpty, ddsdOriginal, DDLOCK_WAIT, 0
    surfDestination.GetSurfaceDesc ddsdDestination
    surfDestination.Lock rEmpty2, ddsdDestination, DDLOCK_WAIT, 0
    
    XDest = XDest + lWidth / 2
    YDest = YDest + lHeight / 2
    For iX = -dblRMax To dblRMax
        For iY = -dblRMax To dblRMax
            lngXO = lWidth / 2 - (iX * CosA + iY * SinA)
            lngYO = lHeight / 2 - (iX * SinA - iY * CosA)
            If lngXO >= 0 Then
                If lngYO >= 0 Then
                    If lngXO < lWidth Then
                        If lngYO < lHeight Then
                            lngColor = surfSource.GetLockedPixel(lngXO, lngYO)
                            If rgbTransparency = -1 Or lngColor <> rgbTransparency Then
                                surfDestination.SetLockedPixel XDest + iX, YDest + iY, lngColor
                            End If
                        End If
                    End If
                End If
            End If
        Next iY
    Next iX
    surfSource.Unlock rEmpty
    surfDestination.Unlock rEmpty2
    Width = lWidth / 2 - (iX * CosA + iY * SinA)
    Height = lHeight / 2 - (iX * SinA - iY * CosA)
End Function



