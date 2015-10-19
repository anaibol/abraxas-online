VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMotor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Private rectMain As RECT
'Particulas
    Private Type BCOLOR
        b As Byte
        g As Byte
        r As Byte
        a As Byte
    End Type
    
    Private Const pStatAlive As Byte = 0
    Private Const pStatDead As Byte = 1
    Private Const pStatDelete As Byte = 2
    Private Const pStatNone As Byte = 3

    Private Type Particle
        x As Single
        y As Single
        
        vX As Single
        vY As Single
        
        sColor As BCOLOR
        eColor As BCOLOR
        cColor As BCOLOR
        
        lifeTime As Long
        created As Long
        status As Byte
    End Type
    
    Private Type LVERTEX
        x As Single
        y As Single
        z As Single
        rhw As Single
        color As Long
        tu As Single
        tv As Single
    End Type
    
    Private Type pa_Gro
        Dead As Byte
        nDelCount As Long
        
        alpha As Byte
        
        equation As Byte
        ratio As Integer
        
        tipe As Integer
        nParticles As Long
        
        size As Long
        Gravity As Single
        
        xWind As Single
        yWind As Single
        
        vX As Single
        vY As Single
        
        x1 As Single
        x2 As Single
        y1 As Single
        y2 As Single
        
        Activated As Byte

        Life As Integer
        PrtData() As Particle
        prtVertList() As LVERTEX
        
        textures(1 To 10) As Integer
        numTexs As Integer
        
        'Pointer to Dest ¡New!
        dest As Long
    End Type
    
    Dim pList() As pa_Gro
    Dim pCount As Integer
    Dim pLast As Integer
    
    Private Const pTextureCount As Integer = 5
    Private pTexture(1 To pTextureCount) As Direct3DTexture8
    
    Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (destination As Any, ByVal Length As Long)
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef Source As Any, ByVal Length As Long)
    Private Declare Sub LerpColor Lib "mEngine.dll" (ByRef o As BCOLOR, ByRef a As BCOLOR, ByRef b As BCOLOR, ByVal F As Byte)
'Particulas
'Luces

Dim light_count As Long

'Luces

'Textos
Private Type Fuente
    Tamanio As Integer
    Caracteres(0 To 255) As Integer 'indice de cada letra
    tex As D3D8Textures
End Type
    
Private Fuentes(1 To 1) As Fuente


Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public engineBaseSpeed As Single
Public FPS As Integer

Private FramesPerSecCounter As Integer
Private timerElapsedTime As Single
Private timerTicksPerFrame As Double

Private lFrameTimer As Long

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer

'Sets a Grh animation to loop indefinitely.
Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
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

Public Sub Engine_Init()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
On Error GoTo ErrHandler:

    Dim DispMode As D3DDISPLAYMODE
    Dim DispModeBK As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    
    Set SurfaceDB = New clsTexManager
    
    Set dX = New DirectX8
    Set D3D = dX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
    
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK

    With D3DWindow
        .Windowed = True

        .SwapEffect = D3DSWAPEFFECT_COPY

        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = 2000
        .BackBufferHeight = 2000
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.hwnd
    End With
    
    DispMode.Format = D3DFMT_X8R8G8B8


    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                    D3DWindow)
                                                            
    HalfWindowTileHeight = (frmMain.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.ScaleWidth / 32) \ 2
    
    TileBufferSize = 5
    
    D3DDevice.SetVertexShader FVF

    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
    Call SurfaceDB.Init(D3DX, D3DDevice, 1000000, 700)

    engineBaseSpeed = 0.028
    
    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    ScrollPixelsPerFrameX = 32 'Modificar
    ScrollPixelsPerFrameY = 32


    UserPos.x = 50
    UserPos.y = 50
    
    Font_Initializate
    Particle_Save_Load
    Particle_Texture_Load
    
    'partículas
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
    frmMain.ScaleMode = 3
    
    rectMain.Right = frmMain.ScaleWidth
    rectMain.bottom = frmMain.ScaleHeight
 
    bRunning = True
    
    Exit Sub
    
ErrHandler:
    Debug.Print "Error Number Returned: " & err.Number
    bRunning = False
End Sub

Public Sub Engine_Reset()
        
    HalfWindowTileHeight = (frmMain.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.ScaleWidth / 32) \ 2
    
    frmMain.ScaleMode = 3
    
    rectMain.Right = frmMain.ScaleWidth
    rectMain.bottom = frmMain.ScaleHeight
    
    Debug.Print frmMain.Width, frmMain.Height

    Exit Sub
    
End Sub

Public Sub Engine_Deinit()
    Erase MapData
    Erase charlist
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set dX = Nothing
    End
End Sub

Private Sub Draw_Grh(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal Animate As Byte, ByRef rgb_list() As Long, Optional ByVal alpha As Boolean, Optional ByVal map_x As Integer = 1, Optional ByVal map_y As Integer = 1, Optional ByVal angle As Single, Optional ByVal Shadow As Byte)
    Dim CurrentGrhIndex As Integer
    If Grh.GrhIndex = 0 Then Exit Sub
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> -1 Then
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
    If GrhData(Grh.GrhIndex).NumFrames > 0 Then
        CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    Else
        Exit Sub
    End If

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
        End If

        If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
        End If
    End If
    


    Device_Box_Textured_Render CurrentGrhIndex, x, y, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, rgb_list(), GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, alpha, angle, Shadow

End Sub

Public Sub Render()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************


D3DDevice.BeginScene
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0

    RenderScreen

    If GetTickCount - lFrameTimer > 1000 Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount
    End If
    
    Text_Render "FPS:" & FPS, 5, 5, TmpRGB, -1, 1
    Text_Render "Ultimo Click : " & IIf(UltimoClickX > 9, UltimoClickX, "0" & UltimoClickX) & "," & IIf(UltimoClickY > 9, UltimoClickY, "0" & UltimoClickY), 5, 15, TmpRGB, -1, 1
    Text_Render "Posicion :      " & IIf(UserPos.x > 9, UserPos.x, "0" & UserPos.x) & "," & IIf(UserPos.y > 9, UserPos.y, "0" & UserPos.y), 5, 25, TmpRGB, -1, 1

D3DDevice.EndScene
D3DDevice.Present rectMain, rectMain, frmMain.hwnd, ByVal 0

FramesPerSecCounter = FramesPerSecCounter + 1
timerElapsedTime = GetElapsedTime()
timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

End Sub


Sub RenderScreen()
    Dim y As Long, x As Long
    Dim GrhP As Grh
    Dim screenminY, screenmaxY, screenminX, screenmaxX, minY, maxY, minX, maxX As Integer
    Dim ScreenX, ScreenY, minXOffset, minYOffset, PixelOffsetXTemp, PixelOffsetYTemp, CurrentGrhIndex As Integer
    Dim tilex As Integer, tiley As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer
    
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    If UserMoving Then
        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.x <> 0 Then
            OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.x * timerTicksPerFrame
            If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.x) Then
                OffsetCounterX = 0
                AddtoUserPos.x = 0
                UserMoving = False
            End If
        End If
        
        '****** Move screen Up and Down if needed ******
        If AddtoUserPos.y <> 0 Then
            OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.y * timerTicksPerFrame
            If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.y) Then
                OffsetCounterY = 0
                AddtoUserPos.y = 0
                UserMoving = False
            End If
        End If
    End If
    
    tilex = UserPos.x - AddtoUserPos.x
    tiley = UserPos.y - AddtoUserPos.y
    PixelOffsetX = OffsetCounterX
    PixelOffsetY = OffsetCounterY

    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < 1 Then
        minYOffset = 1 - minY
        minY = 1
    End If
    
    If maxY > MapInfo.dY Then maxY = MapInfo.dY
    
    If minX < 1 Then
        minXOffset = 1 - minX
        minX = 1
    End If
    
    If maxX > MapInfo.dX Then maxX = MapInfo.dX
    
    'If we can, we render around the view area to make it smoother
    If screenminY > 1 Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < MapInfo.dY Then screenmaxY = screenmaxY + 1
    
    If screenminX > 1 Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < MapInfo.dX Then screenmaxX = screenmaxX + 1
    
    If screenmaxY > MapInfo.dY Then screenmaxY = MapInfo.dY

    If screenmaxX > MapInfo.dX Then screenmaxX = MapInfo.dX
    
    
    'Draw floor layer
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            With MapData(x, y)
                
                'Layer 1 **********************************
                If .Graphic(1).GrhIndex And ShowLayer1 Then
                        If .Graphic(1).Started = 1 Then
                            .Graphic(1).FrameCounter = .Graphic(1).FrameCounter + (timerElapsedTime * GrhData(.Graphic(1).GrhIndex).NumFrames / .Graphic(1).Speed)
                            If .Graphic(1).FrameCounter > GrhData(.Graphic(1).GrhIndex).NumFrames Then
                                .Graphic(1).FrameCounter = (.Graphic(1).FrameCounter Mod GrhData(.Graphic(1).GrhIndex).NumFrames) + 1
                
                                If .Graphic(1).Loops <> -1 Then
                                    If .Graphic(1).Loops > 0 Then
                                        .Graphic(1).Loops = .Graphic(1).Loops - 1
                                    Else
                                        .Graphic(1).Started = 0
                                    End If
                                End If
                            End If
                        End If
                    If GrhData(.Graphic(1).GrhIndex).NumFrames <> 0 Then
                     CurrentGrhIndex = GrhData(.Graphic(1).GrhIndex).Frames(.Graphic(1).FrameCounter)
    
                    Device_Box_Textured_Render CurrentGrhIndex, _
                        (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, _
                        GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
                        .light_value, _
                        GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
                        False _
                        , 0
                    End If
                End If
                
                If PutSurface = True Then
                    If Val(frmMode.lstCapa.Text) = 1 Then
                        If frmMode.ckView.value = vbChecked Then
                            If MapData(x, y).gView(1).GrhIndex <> 0 Then
                                Call Draw_Grh(.gView(1), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 0, 1, AlphaRGB, False, x, y)
                            End If
                        End If
                   End If
                End If
                        
                
                'Layer 2 **********************************
                    If .Graphic(2).GrhIndex <> 0 And ShowLayer2 Then _
                        Call Draw_Grh(.Graphic(2), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 1, 1, .light_value, False, x, y)
                
                If PutSurface = True Then
                    If Val(frmMode.lstCapa.Text) = 2 Then
                        If frmMode.ckView.value = vbChecked Then
                            If MapData(x, y).gView(2).GrhIndex <> 0 Then
                                Call Draw_Grh(.gView(2), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 1, 1, AlphaRGB, False, x, y)
                            End If
                        End If
                   End If
                End If
                
                ScreenX = ScreenX + 1
            End With
        Next x
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next y
            
                
    ScreenY = minYOffset - TileBufferSize
    For y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For x = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            With MapData(x, y)
                'Object Layer **********************************
                     If .ObjGrh.GrhIndex <> 0 And ShowObjs Then
                         Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, .light_value, False, x, y)
                         If .Graphic(3).GrhIndex <> 0 Then
                            Call Text_Render("OBJ&3", PixelOffsetXTemp, PixelOffsetYTemp, TmpRGB, -1, 1)
                        Else
                            Call Text_Render("OBJ", PixelOffsetXTemp, PixelOffsetYTemp, TmpRGB, -1, 1)
                        End If
                     End If
                     
                     If .objView.GrhIndex <> 0 And PutObjs Then
                         Call Draw_Grh(.objView, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, AlphaRGB, False, x, y)
                     End If
                 '***********************************************
                             
                 'Char layer ************************************
                     If .CharIndex <> 0 And ShowNpcs Then
                         Call Char_Render(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp, x, y)
                     End If
                 '*************************************************
                 
                If ShowChar Then
                    If x = screenminX + 1 + HalfWindowTileWidth Then
                        If y = screenminY + 1 + HalfWindowTileWidth Then
                            Char_Render 9990, PixelOffsetXTemp, PixelOffsetYTemp, x, y
                        End If
                    End If
                End If
                 
                 If ShowLayer3 Then
                     'Layer 3 *****************************************
                         If .Graphic(3).GrhIndex <> 0 Then _
                             Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, .light_value, False, x, y)
                     '************************************************
                 End If
                 
                 If PutSurface = True Then
                     If Val(frmMode.lstCapa.Text) = 3 Then
                         If frmMode.ckView.value = vbChecked Then
                             If MapData(x, y).gView(3).GrhIndex <> 0 Then
                                 Call Draw_Grh(.gView(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, AlphaRGB, False, x, y)
                             End If
                         End If
                    End If
                 End If
                
            End With
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
    ScreenY = minYOffset - 5
    
    Particle_Set_Alpha

    ScreenY = minYOffset - TileBufferSize
    For y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For x = minX To maxX
            With MapData(x, y)
                '***********************************************
                If .Particle_Group > 0 Then
                    Particle_Update .Particle_Group, ScreenX * 32 + PixelOffsetX + 16, ScreenY * 32 + PixelOffsetY + 8
                    Particle_Render .Particle_Group
                End If
            End With
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y

    Particle_Unset_Alpha
    
    
    'Draw Layer 4 And Particles
    ScreenY = minYOffset - TileBufferSize
    For y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For x = minX To maxX
            With MapData(x, y)
                'Layer 4******************************************
                If MapData(x, y).Graphic(4).GrhIndex Then
                    If bTecho = 1 Then
                        Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, .light_value(), False, x, y)
                    ElseIf bTecho = 2 Then
                        Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, .light_value(), True, x, y)
                    End If
                End If
                
                If PutSurface = True Then
                    If Val(frmMode.lstCapa.Text) = 4 Then
                        If frmMode.ckView.value = vbChecked Then
                            If MapData(x, y).gView(4).GrhIndex <> 0 Then
                                Call Draw_Grh(.gView(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, AlphaRGB, False, x, y)
                            End If
                        End If
                    End If
                End If
                
                If MapData(x, y).selec.GrhIndex <> 0 Then
                    Engine.Draw_GrhIndex MapData(x, y).selec.GrhIndex, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, TmpRGB
                End If
                
                'If MapData(X, Y).autoSelect > 0 Then
                '    Engine.Draw_GrhIndex 2, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, TmpRGB
                'End If
            End With
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
                
    ScreenY = minYOffset
    For y = minY To maxY
        ScreenX = minXOffset
        For x = minX To maxX
            InitGrh GrhP, 1
                If ShowTriggers And MapData(x, y).Trigger <> 0 Then
                    Engine.Text_Render CStr(MapData(x, y).Trigger), (ScreenX - 5) * 32 + PixelOffsetX, (ScreenY - 5) * 32 + PixelOffsetY, TmpRGB, -1, 1
                End If
                
                If MapData(x, y).Blocked = 1 And ShowBlocked = True Then
                    GrhP.GrhIndex = 4
                    Draw_Grh GrhP, (ScreenX - 5) * 32 + PixelOffsetX, (ScreenY - 5) * 32 + PixelOffsetY, 1, 0, TmpRGB(), False, x, y
                End If
                
                If MapData(x, y).TileExit.map <> 0 And ShowTrans = True Then
                    GrhP.GrhIndex = 3
                    Grh_Render_Advance GrhP, (ScreenX - 5) * 32 + PixelOffsetX + 12, (ScreenY - 5) * 32 + PixelOffsetY + 12, 20, 20, TmpRGB()
                    
                    If frmMain.ViewX = x And frmMain.ViewY = y Then
                        Engine.Text_Render MapData(x, y).TileExit.map & ":" & MapData(x, y).TileExit.x & "," & MapData(x, y).TileExit.y, (ScreenX - 5) * 32 + PixelOffsetX, (ScreenY - 5) * 32 + PixelOffsetY, TmpRGB, -1, 1
                    End If
                
                End If
                
                If (x = LastPostCliked.x) And (y = LastPostCliked.y) Then
                    GrhP.GrhIndex = 2
                    Draw_Grh GrhP, (ScreenX - 5) * 32 + PixelOffsetX, (ScreenY - 5) * 32 + PixelOffsetY, 1, 0, TmpRGB(), False, x, y
                    Text_Render "Mouse :        " & IIf(x > 9, x, "0" & x) & "," & IIf(y > 9, y, "0" & y), 5, 35, TmpRGB, -1, 1
                End If
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
        
    If ShowGrill Then Device_Box_Textured_Render 4, 0, 0, 800, 600, GrillRGB(), 0, 0, False, 0
End Sub

Public Sub Draw_GrhIndex(ByVal grh_index As Integer, ByVal x As Integer, ByVal y As Integer, ByRef rgb_list() As Long, Optional alpha As Boolean = False, Optional angle As Single = 0)
    If grh_index <= 0 Then Exit Sub
    Device_Box_Textured_Render grh_index, x, y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list(), GrhData(grh_index).sX, GrhData(grh_index).sY, alpha, angle
End Sub
Private Function Geometry_Create_TLVertex(ByVal x As Single, ByVal y As Single, ByVal z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.x = x
    Geometry_Create_TLVertex.y = y
    Geometry_Create_TLVertex.z = z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function
Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_size As Long, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.bottom - dest.Top) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    
    '0 - Bottom left vertex
    If Textures_size Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_size, (src.bottom + 1) / Textures_size)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    
    '1 - Top left vertex
    If Textures_size Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_size, src.Top / Textures_size)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If Textures_size Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_size, (src.bottom + 1) / Textures_size)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    
    '3 - Top right vertex
    If Textures_size Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_size, src.Top / Textures_size)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 0, 0)
    End If

End Sub

Public Sub Device_Box_Textured_Render_Advance(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, ByVal dest_width As Integer, Optional ByVal dest_height As Integer, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'Copies the Textures allowing resizing
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long

    If GrhIndex = 0 Then Exit Sub
    Set d3dTextures.texture = SurfaceDB.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.Dimension)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)

    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .Left = dest_x
        .Right = dest_x + dest_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.Dimension, angle
        
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub

Public Sub Device_Box_Textured_Render(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single, _
                                            Optional ByVal Shadow As Byte)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/12/2004
'Just copies the Textures
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long
    
    If GrhIndex = 0 Then Exit Sub
    Set d3dTextures.texture = SurfaceDB.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.Dimension)
    
    If ShowLuces Then
        light_value(0) = rgb_list(0)
        light_value(1) = rgb_list(1)
        light_value(2) = rgb_list(2)
        light_value(3) = rgb_list(3)
        
        If (light_value(0) = 0) Then light_value(0) = base_light
        If (light_value(1) = 0) Then light_value(1) = base_light
        If (light_value(2) = 0) Then light_value(2) = base_light
        If (light_value(3) = 0) Then light_value(3) = base_light
    Else
        light_value(0) = -1
        light_value(1) = -1
        light_value(2) = -1
        light_value(3) = -1
    End If
    
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.Dimension, angle
    
    If Shadow Then
        temp_verts(1).x = temp_verts(1).x + src_width / 2
        temp_verts(1).y = temp_verts(1).y - src_height / 2
    
        temp_verts(3).x = temp_verts(3).x + src_width
        temp_verts(3).y = temp_verts(3).y - src_width
    End If
    
    
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub


Public Sub Engine_MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim x As Integer
    Dim y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            y = -1
        
        Case E_Heading.EAST
            x = 1
        
        Case E_Heading.SOUTH
            y = 1
        
        Case E_Heading.WEST
            x = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.y + y
    
    If ShowChar = True Then
        If Map_Legal_Pos(tX, tY) = False Then Exit Sub
    End If
    
    'Check to see if its out of bounds
    If tX < MinXBorder + 1 Or tX > MaxXBorder Or tY < MinYBorder + 1 Or tY > MaxYBorder Then
        Exit Sub
    Else
        If ShowChar Then
            bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 4, 0, 1)
                
            charlist(9990).Heading = nHeading
        End If
    
        'Start moving... MainLoop does the rest
        AddtoUserPos.x = x
        UserPos.x = tX
        AddtoUserPos.y = y
        UserPos.y = tY
        UserMoving = 1
    End If
End Sub
Function Map_Legal_Pos(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    'Limites del mapa
    If x < 1 Or x > MapInfo.dX Or y < 1 Or y > MapInfo.dY Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(x, y).Blocked = 1 Then
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(x, y).CharIndex > 0 Then
        Exit Function
    End If
    
    Map_Legal_Pos = True
End Function
Public Sub Char_Erase(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next
    If CharIndex = 0 Then Exit Sub
    
    charlist(CharIndex).active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until charlist(LastChar).active = 1
            LastChar = LastChar - 1
            If LastChar <= 0 Then Exit Do
        Loop
    End If
    
    If LastChar <= 0 Then Exit Sub
      
    MapData(charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.y).CharIndex = 0
  
    'Remove char's dialog
    Call Char_Reset(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
End Sub
Public Sub Char_Make(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal x As Integer, ByVal y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    If x = 0 Or y = 0 Then Exit Sub
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .muerto = (Head = 500 Or Head = 501)

        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.x = x
        .Pos.y = y
        
        'Make active
        .active = 1
    End With
    
    'Plot on map
    MapData(x, y).CharIndex = CharIndex
End Sub

Public Sub Char_Reset(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        .active = 0
        .Criminal = 0
        .FxIndex = 0
        .invisible = False
        .Moving = 0
        .muerto = False
        .Nombre = ""
        .pie = False
        .Pos.x = 0
        .Pos.y = 0
        .UsandoArma = False
    End With
End Sub

Private Sub Char_Render(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal x As Long, y As Long)
    With charlist(CharIndex)
        .Heading = 3
        
        .Body.Walk(.Heading).Started = 0
        .Body.Walk(.Heading).FrameCounter = 1

        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        If .Head.Head(.Heading).GrhIndex Then
            If .Body.Walk(.Heading).GrhIndex Then
                Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, MapData(.Pos.x, .Pos.y).light_value, False, x, y)
            End If
                            
            If .Head.Head(.Heading).GrhIndex Then
                Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, MapData(.Pos.x, .Pos.y).light_value, False, x, y)
            End If
        Else
            If .Body.Walk(.Heading).GrhIndex Then
                Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, MapData(.Pos.x, .Pos.y).light_value, , x, y)
            End If
        End If
        
        Call Text_Render("NPC", PixelOffsetX, PixelOffsetY + 30, TmpRGB, -1, 1)
    End With
End Sub

Private Function Engine_FToDW(F As Single) As Long
    Dim Buf As D3DXBuffer
    Set Buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData Buf, 0, 4, 1, F
    D3DX.BufferGetData Buf, 0, 4, 1, Engine_FToDW
End Function

Private Sub Grh_Render_Advance(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal Height As Integer, ByVal Width As Integer, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 11/19/2003
'Similar to Grh_Render, but let´s you resize the Grh
'**************************************************************
    Dim grh_index As Long
    
    'Animation
    If Grh.Started Then
        Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.Speed)
        If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = 1
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    grh_index = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    'Center Grh over X, Y pos
    If GrhData(Grh.GrhIndex).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(Grh.GrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
    End If
    
    If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(Grh.GrhIndex).TileHeight * 32) + 32
    End If

    'Draw it to device
    Device_Box_Textured_Render_Advance grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list(), _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        Width, Height, alpha_blend, Grh.angle
End Sub


Sub Text_Render(Texto As String, ByVal x As Integer, ByVal y As Integer, ByRef tempRGB() As Long, ByVal color As Long, Optional ByVal font_index As Integer = 2, Optional multi_line As Boolean = False)
On Error Resume Next
Dim a, b, d, e, F As Integer
Dim graf As Grh
Dim temp_array(3) As Long 'Si le queres dar color a la letra pasa este parametro dsp xD

If Not color = -58 Then
    temp_array(0) = color
    temp_array(1) = color
    temp_array(2) = color
    temp_array(3) = color
Else
    temp_array(0) = tempRGB(0)
    temp_array(1) = tempRGB(1)
    temp_array(2) = tempRGB(2)
    temp_array(3) = tempRGB(3)
End If

If (Len(Texto) = 0) Then Exit Sub

d = 0
If multi_line = False Then
    For a = 1 To Len(Texto)
        b = Asc(mid(Texto, a, 1))
        graf.GrhIndex = Fuentes(font_index).Caracteres(b)
        If b <> 32 Then
            If graf.GrhIndex <> 0 Then
                graf.GrhIndex = Fuentes(font_index).Caracteres(b)
                Font_Render graf.GrhIndex, 1 + (x + d), y, temp_array, font_index
                d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 2
            Else
                Debug.Print Chr(b) & " and " & b
            End If
        Else
            d = d + 4
        End If
    Next a
Else
    e = 0
    F = 0
    For a = 1 To Len(Texto)
        b = Asc(mid(Texto, a, 1))
        graf.GrhIndex = Fuentes(font_index).Caracteres(b)
        If b = 32 Or b = 13 Then
            If e >= 21 Then 'reemplazar por lo que os plazca
                F = F + 1
                e = 0
                d = 0
            Else
                If b = 32 Then d = d + 4
            End If
        Else
            If graf.GrhIndex > 12 Then
                graf.GrhIndex = Fuentes(font_index).Caracteres(b)
                Font_Render graf.GrhIndex, (x + d), y + F * 14, temp_array, font_index '14 es el height de esta fuente dsp lo hacemos dinamico
                d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
            End If
       End If
       e = e + 1
    Next a
End If


End Sub
Function Font_Render(ByVal grh_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, ByVal font_index As Byte)
    Dim light_value(0 To 3) As Long
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    
    grh_index = GrhData(grh_index).Frames(1)
    
    'Set up the source rectangle
    With src_rect
        .bottom = GrhData(grh_index).sY + GrhData(grh_index).pixelHeight
        .Left = GrhData(grh_index).sX
        .Right = GrhData(grh_index).sX + GrhData(grh_index).pixelWidth
        .Top = GrhData(grh_index).sY
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .bottom = screen_y + GrhData(grh_index).pixelHeight
        .Left = screen_x
        .Right = screen_x + GrhData(grh_index).pixelWidth
        .Top = screen_y
    End With
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), 256
    
    'Set Textures
    D3DDevice.SetTexture 0, Fuentes(font_index).tex.texture
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
End Function
Public Function Font_Initializate()
    Dim j As Long, h As Long
    h = 24031
    For j = 32 To 255
        Fuentes(1).Caracteres(j) = h
        h = h + 1
    Next j
    
    Set Fuentes(1).tex.texture = D3DX.CreateTextureFromFileEx(D3DDevice, DirGraficos & "font.bmp", _
                D3DX_DEFAULT, D3DX_DEFAULT, _
                D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                D3DX_FILTER_POINT, &HFF000000, _
                ByVal 0, ByVal 0)
    
    Dim surface_desc As D3DSURFACE_DESC
    Fuentes(1).tex.texture.GetLevelDesc 0, surface_desc
    
    Fuentes(1).tex.Dimension = surface_desc.Width
End Function
Public Function Light_Create(ByVal map_x As Integer, ByVal map_y As Integer, Optional range As Byte = 3, Optional ByVal color As Long = -1)
    
    light_count = light_count + 1
    ReDim Preserve light_list(1 To light_count) As tLight
    
    With light_list(light_count)
        .color = color
        .range = range
        .active = True
        .map_x = map_x
        .map_y = map_y
    End With
    
    MapData(map_x, map_y).Luz = light_count
    
    Call LightRender(light_count)
End Function
Public Function Light_Delete_From_Pos(ByVal x As Integer, ByVal y As Integer)
    Dim i As Long
    For i = 1 To light_count
        If light_list(i).active Then
            If light_list(i).map_x = x And light_list(i).map_y = y Then
                Light_Delete i
                Exit Function
            End If
        End If
    Next i
End Function
Public Function Light_Delete_All()
    Dim i As Long
    For i = 1 To light_count
        Light_Delete i
    Next i
    Erase light_list
    light_count = 0
End Function
Public Function Light_Delete(ByVal light_index As Long)
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim x As Integer
    Dim y As Integer
    Dim light_trigger As Integer
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    light_trigger = MapData(light_list(light_index).map_x, light_list(light_index).map_y).Trigger
    
    light_list(light_index).active = False
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).Luz = 0
    
    If light_index = light_count And light_count <> 0 Then
        Do While light_list(light_count).active = False
            light_count = light_count - 1
            If light_count = 0 Then Exit Do
        Loop
    End If
    
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        If MapData(min_x, min_y).Trigger = light_trigger Then _
            MapData(min_x, min_y).light_value(2) = 0
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        If MapData(max_x, min_y).Trigger = light_trigger Then _
            MapData(max_x, min_y).light_value(0) = 0
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        If MapData(max_x, max_y).Trigger = light_trigger Then _
            MapData(max_x, max_y).light_value(1) = 0
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        If MapData(min_x, max_y).Trigger = light_trigger Then _
            MapData(min_x, max_y).light_value(3) = 0
    End If
    
    'Arrange borders
    'Upper border
    For x = min_x + 1 To max_x - 1
        If InMapBounds(x, min_y) Then
            If MapData(x, min_y).Trigger = light_trigger Then
                MapData(x, min_y).light_value(0) = 0
                MapData(x, min_y).light_value(2) = 0
            End If
        End If
    Next x
    
    'Lower border
    For x = min_x + 1 To max_x - 1
        If InMapBounds(x, max_y) Then
            If MapData(x, max_y).Trigger = light_trigger Then
                MapData(x, max_y).light_value(1) = 0
                MapData(x, max_y).light_value(3) = 0
            End If
        End If
    Next x
    
    'Left border
    For y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, y) Then
            If MapData(min_x, y).Trigger = light_trigger Then
                MapData(min_x, y).light_value(2) = 0
                MapData(min_x, y).light_value(3) = 0
            End If
        End If
    Next y
    
    'Right border
    For y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, y) Then
            If MapData(max_x, y).Trigger = light_trigger Then
                MapData(max_x, y).light_value(0) = 0
                MapData(max_x, y).light_value(1) = 0
            End If
        End If
    Next y
    
    'Set the inner part of the light
    For x = min_x + 1 To max_x - 1
        For y = min_y + 1 To max_y - 1
            If InMapBounds(x, y) Then
                If MapData(x, y).Trigger = light_trigger Then
                    MapData(x, y).light_value(0) = 0
                    MapData(x, y).light_value(1) = 0
                    MapData(x, y).light_value(2) = 0
                    MapData(x, y).light_value(3) = 0
                End If
            End If
        Next y
    Next x
   
End Function
 
Private Sub LightRender(ByVal light_index As Long)
 
    If light_index = 0 Then Exit Sub
    If light_list(light_index).active = False Then Exit Sub
   
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim x As Integer
    Dim y As Integer
    Dim color As Long
    Dim light_trigger As Integer
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    light_trigger = MapData(light_list(light_index).map_x, light_list(light_index).map_y).Trigger
    
    'Set color
    color = light_list(light_index).color
    
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        If MapData(min_x, min_y).Trigger = light_trigger Then _
            MapData(min_x, min_y).light_value(2) = color
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        If MapData(max_x, min_y).Trigger = light_trigger Then _
            MapData(max_x, min_y).light_value(0) = color
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        If MapData(max_x, max_y).Trigger = light_trigger Then _
            MapData(max_x, max_y).light_value(1) = color
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        If MapData(min_x, max_y).Trigger = light_trigger Then _
            MapData(min_x, max_y).light_value(3) = color
    End If
    
    'Arrange borders
    'Upper border
    For x = min_x + 1 To max_x - 1
        If InMapBounds(x, min_y) Then
            If MapData(x, min_y).Trigger = light_trigger Then
                MapData(x, min_y).light_value(0) = color
                MapData(x, min_y).light_value(2) = color
            End If
        End If
    Next x
    
    'Lower border
    For x = min_x + 1 To max_x - 1
        If InMapBounds(x, max_y) Then
            If MapData(x, max_y).Trigger = light_trigger Then
                MapData(x, max_y).light_value(1) = color
                MapData(x, max_y).light_value(3) = color
            End If
        End If
    Next x
    
    'Left border
    For y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, y) Then
            If MapData(min_x, y).Trigger = light_trigger Then
                MapData(min_x, y).light_value(2) = color
                MapData(min_x, y).light_value(3) = color
            End If
        End If
    Next y
    
    'Right border
    For y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, y) Then
            If MapData(max_x, y).Trigger = light_trigger Then
                MapData(max_x, y).light_value(0) = color
                MapData(max_x, y).light_value(1) = color
            End If
        End If
    Next y
    
    'Set the inner part of the light
    For x = min_x + 1 To max_x - 1
        For y = min_y + 1 To max_y - 1
            If InMapBounds(x, y) Then
                If MapData(x, y).Trigger = light_trigger Then
                    MapData(x, y).light_value(0) = color
                    MapData(x, y).light_value(1) = color
                    MapData(x, y).light_value(2) = color
                    MapData(x, y).light_value(3) = color
                End If
            End If
        Next y
    Next x
End Sub
'################################################
'#                 Particulas                   #
'################################################
Public Sub Particle_Create(ByVal pX1 As Integer, ByVal pY1 As Integer, ByVal pX2 As Integer, ByVal pY2 As Integer, _
                            ByVal vX As Single, ByVal vY As Single, _
                            ByVal sA As Byte, ByVal sR As Byte, ByVal sG As Byte, ByVal sB As Byte, _
                            ByVal eA As Byte, ByVal eR As Byte, ByVal eG As Byte, ByVal eB As Byte, _
                            ByVal nParticles As Integer, ByVal Gravity As Single, _
                            ByRef texList() As Integer, ByVal numTexs As Integer, ByVal size As Single, ByVal Life As Integer, _
                            ByVal eq As Byte, ByVal Dead As Byte, ByRef dest As Integer, _
                            ByVal ratio As Integer)
                                
    Dim i As Long
    Dim nowA As Long
    Dim pindex As Integer
    
    pCount = pCount + 1
    If pCount > pLast Then
        pLast = pCount
        pindex = pCount
        ReDim Preserve pList(0 To pLast)
    Else
        For i = 1 To pLast
            If pList(i).Activated = 0 Then
                pindex = i
                Exit For
            End If
        Next i
    End If
    
    'Particle System Mannakia Mod
    With pList(pindex)
        ReDim Preserve .PrtData(0 To nParticles)
        ReDim Preserve .prtVertList(0 To nParticles)
        
        'The Particle Pos
        .x1 = pX1
        .y1 = pY1
        .x2 = pX2
        .y2 = pY2
        
        .Activated = 1

        '¡New!
        .Dead = Dead
        .nDelCount = 0
        .equation = eq
        .dest = VarPtr(dest)
        .ratio = ratio
        '¡New!

        .Life = Life
        .size = Particle_SingleToLong(size)
        .Gravity = Gravity
        
        'Sets alive counter
        .nParticles = nParticles
        
        .numTexs = numTexs
        For i = 1 To .numTexs
            .textures(i) = texList(i)
        Next
        
        .vX = vX
        .vY = vY

        nowA = GetTickCount
        For i = 0 To .nParticles
            .PrtData(i).status = pStatAlive
            .PrtData(i).lifeTime = (Rnd() * .Life / 2)
            .PrtData(i).created = nowA
            
            .PrtData(i).vX = (Rnd * vX) - (vX / 2)
            .PrtData(i).vY = (Rnd * vY) - (vY / 3) + .Gravity * (i * (.size / .nParticles))
            
            If .equation = 0 Then 'Primera ecuacion:D REDONDA :D
                .PrtData(i).x = General_Random_Number(.x1, .x2)
                .PrtData(i).y = General_Random_Number(.y1, .y2)
            ElseIf .equation = 1 Then
                Dim a  As Single
                a = Rnd * 360 * 0.0174532925
                .PrtData(i).x = Sin(a) * (.size * 1.5)
                .PrtData(i).y = Cos(a) * (.size * 1.5)
            End If
            
            .PrtData(i).sColor.a = sA
            .PrtData(i).sColor.r = sR
            .PrtData(i).sColor.g = sG
            .PrtData(i).sColor.b = sB
            
            .PrtData(i).eColor.a = eA
            .PrtData(i).eColor.r = eR
            .PrtData(i).eColor.g = eG
            .PrtData(i).eColor.b = eB
            
            .PrtData(i).cColor = .PrtData(i).sColor
            
            .prtVertList(i).x = .PrtData(i).x
            .prtVertList(i).y = .PrtData(i).y
            .prtVertList(i).color = D3DColorARGB(.PrtData(i).cColor.a, .PrtData(i).cColor.r, .PrtData(i).cColor.g, .PrtData(i).cColor.b)
        Next i
        
    End With

    CopyMemory ByVal VarPtr(dest), pindex, 2
End Sub
Public Sub Particle_Update(ByVal ii As Integer, ByVal offx As Integer, ByVal offy As Integer)
    Static i As Long, nowA As Long, r As Single
    nowA = GetTickCount
        
    If ii > pLast Then Exit Sub
    If pList(ii).Activated = 0 Then Exit Sub
    
    With pList(ii)
        For i = 0 To .nParticles - 1
            If .PrtData(i).status <> pStatNone Then
                If nowA - .PrtData(i).created > .PrtData(i).lifeTime Then
                    If .Dead = 0 Then
                        .PrtData(i).status = pStatDead
                    Else
                        .PrtData(i).status = pStatDelete
                    End If
                End If
            End If
            
            
            If .PrtData(i).status = pStatAlive Then
                .PrtData(i).x = .PrtData(i).x + ((.PrtData(i).vX / 5)) * (timerElapsedTime * 0.05)
                .PrtData(i).y = .PrtData(i).y + ((.PrtData(i).vY / 5)) * (timerElapsedTime * 0.05)
                
                .PrtData(i).vY = .PrtData(i).vY + (.Gravity / 5) * (timerElapsedTime * 0.05)
                
                On Error Resume Next
                LerpColor .PrtData(i).cColor, .PrtData(i).sColor, .PrtData(i).eColor, _
                    ((nowA - .PrtData(i).created) / .PrtData(i).lifeTime) * 255
                On Local Error GoTo 0
                .prtVertList(i).x = .PrtData(i).x + offx
                .prtVertList(i).y = .PrtData(i).y + offy
                .prtVertList(i).z = 0
                .prtVertList(i).rhw = 1
                .prtVertList(i).color = D3DColorARGB(.PrtData(i).cColor.a, .PrtData(i).cColor.r, .PrtData(i).cColor.g, .PrtData(i).cColor.b)
            ElseIf .PrtData(i).status = pStatDead Then
                If .Dead = 0 Then
                    .PrtData(i).status = pStatAlive
                    .PrtData(i).lifeTime = .Life + (Rnd * .Life / 2)
                    .PrtData(i).created = nowA
                        
                    If .equation = 0 Then
                        .PrtData(i).x = General_Random_Number(.x1, .x2)
                        .PrtData(i).y = General_Random_Number(.y1, .y2)
                    ElseIf .equation = 1 Then
                        r = Rnd * 6.2831853
                        .PrtData(i).x = Sin(r) * .ratio   'Este ultimo es el radio
                        .PrtData(i).y = Cos(r) * .ratio
                    End If
    
                    .PrtData(i).vX = (Rnd * .vX) - (.vX * Rnd)
                    .PrtData(i).vY = (Rnd * .vY) - (.vY * Rnd)
                    .PrtData(i).cColor = .PrtData(i).sColor
                    .prtVertList(i).x = .PrtData(i).x + offx
                    .prtVertList(i).y = .PrtData(i).y + offy
                    .prtVertList(i).color = D3DColorARGB(.PrtData(i).cColor.a, .PrtData(i).cColor.r, .PrtData(i).cColor.g, .PrtData(i).cColor.b)
                Else
                    .PrtData(i).status = pStatNone
                    .nDelCount = .nDelCount + 1
                    .prtVertList(i).color = 0
                End If
            ElseIf .PrtData(i).status = pStatDelete Then
                .PrtData(i).status = pStatNone
                .nDelCount = .nDelCount + 1
                .prtVertList(i).color = 0
            End If
        Next i
        
        If .nDelCount >= .nParticles Then
            Particle_Delete ii
        End If
    End With
End Sub
Public Sub Particle_Delete(ByVal i As Integer, Optional ByVal all As Byte, Optional ByVal effect As Byte)
        
    If all = 1 Then
        pCount = 0
        pLast = 0
        Erase pList
        Exit Sub
    End If
    
    If i > pLast Then Exit Sub
    If i = 0 Then Exit Sub
         
    If effect = 1 Then
        pList(i).Dead = 1
        Exit Sub
    End If
    
    If i = pLast Then
        pLast = pLast - 1
    End If
    
    pCount = pCount - 1

    pList(i).Activated = 0
    
    Dim nullPoint As Integer
    CopyMemory ByVal pList(i).dest, nullPoint, 2
    
    ZeroMemory pList(i), LenB(pList(i))
    
End Sub

Private Sub Particle_Set_Alpha()
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
End Sub

Private Sub Particle_Unset_Alpha()
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub

Public Function Particle_SingleToLong(ByVal v As Single) As Long
    Static t As Long
    
    CopyMemory t, v, 4
    
    Particle_SingleToLong = t
End Function

Public Sub Particle_Render(ByVal i As Long)
On Error Resume Next
    If i > pLast Then Exit Sub
    If pList(i).alpha = 1 Then Particle_Set_Alpha
    
    D3DDevice.SetRenderState D3DRS_POINTSIZE, pList(i).size
    
    If pList(i).numTexs = 1 Then
        D3DDevice.SetTexture 0, pTexture(pList(i).textures(1))
        D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, pList(i).nParticles, pList(i).prtVertList(0), 28
    Else
        Dim l As Long, d As Long
        
        d = pList(i).nParticles / pList(i).numTexs
        
        For l = 1 To pList(i).numTexs
            D3DDevice.SetTexture 0, pTexture(pList(i).textures(l))
            D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, d, _
                                    pList(i).prtVertList(d * (l - 1)), 28
        Next
    End If
    
    If pList(i).alpha = 1 Then Particle_Unset_Alpha
End Sub

Sub Particle_Save_Create(ByVal i As Integer, ByRef dest As Integer, Optional ByVal Dead As Byte)
    If i > NumParticles Then Exit Sub
    
    With pSaves(i)
        Call Particle_Create(.x1, .y1, .x2, .y2, .vX, .vY, _
                         .sColor.a, .sColor.r, .sColor.g, .sColor.b, _
                         .eColor.a, .eColor.r, .eColor.g, .eColor.b, _
                         .nParticles, .Gravity, .textures, .numTexs, .size, .Life, _
                         .equation, Dead, dest, .ratio)   '¡New!
    End With
End Sub
Private Sub Particle_Texture_Load()
    Dim i As Long
    
    For i = 1 To pTextureCount
        Set pTexture(i) = D3DX.CreateTextureFromFileEx(D3DDevice, DirGraficos & "p" & i & ".png", _
                D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, &HFF000000, ByVal 0, ByVal 0)
    Next
End Sub
Private Sub Particle_Save_Load()
    Dim F As Integer
    Dim i As Long
    Dim tmpStr As String * 32
    
    F = FreeFile
    Open DirInit & "pm.ind" For Binary Access Read As #F
        Get #F, , NumParticles
        ReDim pSaves(1 To NumParticles) As ParticleSave
        For i = 1 To NumParticles
            Get #F, , pSaves(i).name
            Get #F, , pSaves(i).ratio
            
            Get #F, , pSaves(i).sColor
            Get #F, , pSaves(i).eColor
            
            Get #F, , pSaves(i).Gravity
            Get #F, , pSaves(i).Life
            Get #F, , pSaves(i).nParticles
            Get #F, , pSaves(i).size
            
            Get #F, , pSaves(i).textures
            Get #F, , pSaves(i).numTexs
            
            Get #F, , pSaves(i).vX
            Get #F, , pSaves(i).vY
            Get #F, , pSaves(i).x1
            Get #F, , pSaves(i).x2
            Get #F, , pSaves(i).y1
            Get #F, , pSaves(i).y2
            
            Get #F, , pSaves(i).alpha
            
            Get #F, , pSaves(i).equation
            
            frmMode.lstParticles.AddItem i & " - " & RTrim$(pSaves(i).name)
        Next i
    Close #F
End Sub

Sub SaveBackBuffer(ByVal FileName As String, ByVal x As Integer, ByVal y As Integer)
    Dim RECT As RECT
    Dim PAL As PALETTEENTRY
    
    PAL.blue = 255
    PAL.green = 255
    PAL.red = 255
    
    RECT.Right = x
    RECT.bottom = y
    D3DX.SaveSurfaceToFile FileName, D3DXIFF_BMP, D3DDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO), PAL, RECT
End Sub
Sub Set_Map(ByVal x As Integer, ByVal y As Integer)
    HalfWindowTileHeight = y
    HalfWindowTileWidth = x
End Sub
Sub RenderCopy()
    Dim y As Long, x As Long
    Dim GrhP As Grh
    Dim screenminY, screenmaxY, screenminX, screenmaxX As Integer
    
    screenminY = 1
    screenmaxY = cp.dY
    screenminX = 1
    screenmaxX = cp.dX
    
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            With cp.copied(x, y)
                'Layer 1 **********************************
                If .Graphic(1).GrhIndex <> 0 And ShowLayer2 Then
                    Call Draw_Grh(.Graphic(1), (x - 1) * 32, (y - 1) * 32, 0, 1, .light_value, False, x, y)
                End If
                
                'Layer 2 **********************************
                If .Graphic(2).GrhIndex <> 0 And ShowLayer2 Then
                    Call Draw_Grh(.Graphic(2), (x - 1) * 32, (y - 1) * 32, 1, 1, .light_value, False, x, y)
                End If
            End With
        Next x
    Next y
                    
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            With cp.copied(x, y)
                'Object Layer **********************************
                If .ObjGrh.GrhIndex <> 0 And ShowObjs Then
                    Call Draw_Grh(.ObjGrh, (x - 1) * 32, (y - 1) * 32, 1, 1, .light_value, False, x, y)
                End If
                '***********************************************
                             
                'Char layer ************************************
                If .CharIndex <> 0 And ShowNpcs Then
                    Call Char_Render(.CharIndex, (x - 1) * 32, (y - 1) * 32, x, y)
                End If
                 '*************************************************
                 
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 And ShowLayer3 Then
                    Call Draw_Grh(.Graphic(3), (x - 1) * 32, (y - 1) * 32, 1, 1, .light_value, False, x, y)
                End If
                '************************************************
                
            End With
        Next x
    Next y
    
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            With cp.copied(x, y)
                'Layer 4 *****************************************
                If .Graphic(4).GrhIndex <> 0 And bTecho Then
                    Call Draw_Grh(.Graphic(4), (x - 1) * 32, (y - 1) * 32, 1, 1, .light_value, False, x, y)
                End If
                '************************************************
            End With
        Next x
    Next y

    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            With cp.copied(x, y)
                If ShowTriggers And MapData(x, y).Trigger <> 0 Then
                    Engine.Text_Render CStr(MapData(x, y).Trigger), (x - 1) * 32, (y - 1) * 32, TmpRGB, -1, 1
                End If
                
                If MapData(x, y).Blocked = 1 And ShowBlocked = True Then
                    GrhP.GrhIndex = 4
                    Draw_Grh GrhP, (x - 1) * 32, (y - 1) * 32, 1, 0, TmpRGB(), False, x, y
                End If
            End With
        Next x
    Next y
End Sub

