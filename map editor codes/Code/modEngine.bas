Attribute VB_Name = "modEngine"
Option Explicit

Public Type tLight
    active As Boolean
    map_x As Integer
    map_y As Integer
    color As Long
    range As Byte
End Type

Public MiMouse As Boolean

Public light_list() As tLight

Public LastPostCliked As ClikedPos
Public ShowTriggers As Boolean
Public ShowBlocked As Boolean
Public ShowGrill As Boolean
Public ShowTrans As Boolean
Public ShowLayer1 As Boolean
Public ShowLayer2 As Boolean
Public ShowLayer3 As Boolean
Public ShowNpcs As Boolean
Public ShowObjs As Boolean
Public ShowLuces As Boolean
Public ShowAuto As Boolean
Public ShowChar As Boolean

Rem Mannakia
Public PutTrans     As Boolean
Public PutNPC       As Boolean
Public PutAuto      As Boolean
Public PutLight     As Boolean
Public Audio        As New clsAudio
Rem Mannakia
Public bRunning As Boolean

Public PutBlock As Boolean
Public PutSurface As Boolean 'Surface = Superficie Mannakia
Public PutTrigger As Boolean
Public PutParticles As Boolean
Public PutObjs As Boolean

Public TmpRGB(3) As Long
Public GrillRGB(3) As Long
Type ClikedPos
    x As Integer
    y As Integer
End Type

'Map sizes in tiles
Public Enum eClass
    Mage = 1    'Mago
    Cleric      'Clérigo
    Warrior     'Guerrero
    Assasin     'Asesino
    Thief       'Ladrón
    Bard        'Bardo
    Druid       'Druida
    Bandit      'Bandido
    Paladin     'Paladín
    Hunter      'Cazador
    Fisher      'Pescador
    Blacksmith  'Herrero
    Lumberjack  'Leñador
    Miner       'Minero
    Carpenter   'Carpintero
    Pirat       'Pirata
End Enum

Public Enum eObjType
    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLeña = 14
    otFogata = 15
    otESCUDO = 16
    otCASCO = 17
    otAnillo = 18
    otTeleport = 19
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otCualquiera = 1000
End Enum

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'




Public Enum eTrigger
    Nada = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum



'Posicion en un mapa
Public Type Position
    x As Long
    dsd As Byte
    XX As String
    asd As Double
    yy As Long
    y As Long
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    active As Boolean
    MiniMap_color As Long
    
    Speed As Single
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
    angle As Single
End Type

Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type


'Apariencia del personaje
Public Type Char
    active As Byte
    Heading As E_Heading

    alpha As Integer
    alphacounter As Long
    
    alpha_sentido As Boolean
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
    
    MANp As Byte
    VIDp As Byte
    Barras As Boolean
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    
    fX As Grh
    FxIndex As Integer
    
    Criminal As Byte
    
    Nombre As String
    
    Pos As Position
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single

End Type

'Info de un objeto
Public Type Obj
    ObjIndex As Integer
    amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    gView(1 To 4) As Grh
    
    selec As Grh
    
    CharIndex As Integer
    ObjGrh As Grh
    objView As Grh
    
    Luz As Long
    light_value(3) As Long
    
    Particle_Group As Integer
    Particle_index As Integer
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
    
    'autoSelect As Byte
End Type

Type tCopy
    dX As Integer
    dY As Integer
    
    cObj As Boolean
    cNpc As Boolean
    cTrig As Boolean
    cBloq As Boolean
    cCap(1 To 4) As Boolean
    
    copied() As MapBlock
End Type

Public cData As tCopy
Public cp As tCopy

'Info del mapa
Type tMapInfo
    NumUsers As Integer
    Music As String
    name As String
    
    dX As Long
    dY As Long
    
    PK As Boolean
    MagiaSinEfecto As Byte
    InviSinEfecto As Byte
    ResuSinEfecto As Byte
    
    terreno As eTerreno
    restringir As eRestringir
End Type

Public Enum eTerreno
    Desconocido
    Nieve
    Bosque
    Ciudad
    Campo
    Desierto
    Dungeon
End Enum

Public Enum eRestringir
    Nada
    Newbie
    Armada
    Caos
    Faccion
End Enum


Public MapInfo As tMapInfo
'Bordes del mapa
Public MinXBorder As Integer
Public MaxXBorder As Integer
Public MinYBorder As Integer
Public MaxYBorder As Integer

'Status del user
Public UserMoving As Byte
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve


'Tamaño del la vista en Tiles


'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer


'Tamaño de los tiles en pixels


'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Public NumChars As Integer
Public LastChar As Integer

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type

Public Type D3D8Textures
    texture As Direct3DTexture8
    Dimension As Long
End Type

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public CascoAnimData() As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapData2() As MapBlock ' Mapa

Public bTecho       As Byte 'hay techo?

Public charlist(1 To 10000) As Char

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type


Public MiCabecera As tCabecera


Rem ##################################################################

Public SurfaceDB As clsTexManager

Public dX As DirectX8
Public D3D As Direct3D8
Public D3DDevice As Direct3DDevice8
Public D3DX As D3DX8

Public Type TLVERTEX
    x As Single
    y As Single
    z As Single
    rhw As Single
    color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Public Engine As New clsMotor
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const PI As Single = 3.14159265358979
Public Const INFINITE_LOOPS As Integer = -1

Public base_light As Long

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Integer, ByRef tY As Integer)
    tX = UserPos.x + viewPortX \ 32 - frmMain.ScaleWidth \ 64
    tY = UserPos.y + viewPortY \ 32 - (frmMain.ScaleHeight + 24) \ 64
    LastPostCliked.x = tX
    LastPostCliked.y = tY
End Sub

Public Sub Grh_Init(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.GrhIndex = GrhIndex
    If GrhIndex = 0 Then Exit Sub
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub
Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = -1
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed * 2
End Sub


Public Function General_RGB_Color_to_Long(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal a As Long) As Long
        
    Dim c As Long
        
    If a > 127 Then
        a = a - 128
        c = a * 2 ^ 24 Or &H80000000
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    Else
        c = a * 2 ^ 24
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    End If
    
    General_RGB_Color_to_Long = c

End Function

Function NextOpenChar() As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim loopc As Integer

loopc = 1
Do While charlist(loopc).active
    loopc = loopc + 1
Loop

NextOpenChar = loopc

End Function
