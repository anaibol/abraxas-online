Attribute VB_Name = "modDX8_Draw"
Option Explicit

'Lineas
Private Type LINEVERTEX
    X As Single
    Y As Single
    Z As Single
    Rhw As Single
    
    color As Long
End Type

Private Const LFVF As Long = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE
'Linea

'DX8
    Private dX As DirectX8
    Private D3D As Direct3D8
    Public D3DX As D3DX8
    Public D3DDevice As Direct3DDevice8
    
    Private Type LVERTEX
        X As Single
        Y As Single
        Z As Single
        Rhw As Single
        color As Long
        tU As Single
        tV As Single
    End Type

    Private Const FVF As Long = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
    Private Const FVFSize As Long = 28
    
    Private Const hSIZE As Long = 337
    
    Private Type sEntry
        file As Long
        d As Long
        tex As Direct3DTexture8
    End Type
    
    Private Type hNodex
        sCount As Integer
        s() As sEntry
    End Type
    
    Private hList(hSIZE - 1) As hNodex
    
    'Static Declaration Memory
    Dim rvt(0 To 3) As LVERTEX
    Dim rt As Direct3DTexture8
    Dim RD As Long
'DX8

'Textos
    Private Type tChar
        w As Long
        h As Long
        
        tu1 As Single
        tv1 As Single
        tu2 As Single
        tv2 As Single
    End Type
    
    Private chars(32 To 255) As tChar
    Private texChar As Direct3DTexture8
    
    'Static Declaration Memory
    Private Type tQuad
        tV(0 To 3) As LVERTEX
    End Type
    
    Private quads(0 To 1024) As tQuad
    
    Private Type tLine
        T() As Byte
    End Type

    Private SegOn_Msg() As Byte
    Private SegOff_Msg() As Byte
    Private ModCom_Msg() As Byte
'Textos

'Otros
Private LastTexture As Integer
Private AmbientColor As Long

Public sDefaultColor(3) As Long
Public sAlphaColor(3) As Long
'Otros

'#######################################################################################################
'#######################################################################################################
'#######################################################################################################
Public Sub Textured_Preload(ByVal file As Long)
    Dim da As Long
    
    Textured_Get file, da
End Sub
Private Function Textured_Get(ByVal file As Long, ByRef d As Long) As Direct3DTexture8
On Error GoTo Err
    If file = 0 Then Exit Function

    Dim i As Long
    
    ' Search the index on the list
    With hList(file Mod hSIZE)
        For i = 1 To .sCount
            If .s(i).file = file Then
                d = .s(i).d
                Set Textured_Get = .s(i).tex
                Exit Function
            End If
        Next i
    End With

    Textured_Load file, d
    Set Textured_Get = hList(file Mod hSIZE).s(hList(file Mod hSIZE).sCount).tex
    
    Exit Function
Err:
    If Err.Number = 49 Then Resume Next
    
End Function
Private Sub Textured_Load(ByVal file As Long, ByRef d As Long)
    Dim tInfo As D3DXIMAGE_INFO
    Dim Index As Long

    Index = file Mod hSIZE
    
    With hList(Index)
        .sCount = .sCount + 1
        
        ReDim Preserve .s(1 To .sCount) As sEntry
        
        With .s(.sCount)
            'Nombre
            .file = file
            
            Set .tex = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\Grh\" & file & ".png", _
                D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, &HFF000000, tInfo, ByVal 0)
             
            .d = tInfo.Width
            d = .d
        End With
    End With
End Sub

Public Sub Draw_Quad(ByVal file As Long, ByVal dX As Integer, ByVal dY As Integer, _
                        ByVal sW As Integer, ByVal sH As Integer, ByVal sX As Integer, ByVal sY As Integer, _
                        ByRef color() As Long)
    If file = 0 Then Exit Sub
    
    If LastTexture <> file Then
        Set rt = Textured_Get(file, RD)
    
        D3DDevice.SetTexture 0, rt
        LastTexture = file
    End If
    
    rvt(0).X = dX
    rvt(0).Y = dY + sH
    rvt(0).Z = 0
    rvt(0).Rhw = 1
    rvt(0).color = IIf(color(0) = 0, AmbientColor, color(0))
    rvt(0).tU = sX / RD
    rvt(0).tV = (sY + sH + 1) / RD
    
    rvt(1).X = dX
    rvt(1).Y = dY
    rvt(1).Z = 0
    rvt(1).Rhw = 1
    rvt(1).color = IIf(color(1) = 0, AmbientColor, color(1))
    rvt(1).tU = sX / RD
    rvt(1).tV = sY / RD
    
    rvt(2).X = dX + sW
    rvt(2).Y = dY + sH
    rvt(2).Z = 0
    rvt(2).Rhw = 1
    rvt(2).color = IIf(color(2) = 0, AmbientColor, color(2))
    rvt(2).tU = (sX + sW + 1) / RD
    rvt(2).tV = (sY + sH + 1) / RD
    
    rvt(3).X = dX + sW
    rvt(3).Y = dY
    rvt(3).Z = 0
    rvt(3).Rhw = 1
    rvt(3).color = IIf(color(3) = 0, AmbientColor, color(3))
    rvt(3).tU = (sX + sW + 1) / RD
    rvt(3).tV = sY / RD
    
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, rvt(0), FVFSize
End Sub

Public Sub Draw_Quad_Ex(ByVal file As Long, ByVal dX As Integer, ByVal dY As Integer, _
                        ByVal sW As Integer, ByVal sH As Integer, ByVal sX As Integer, ByVal sY As Integer, _
                        ByRef color() As Long, ByVal angle As Single)
    If file = 0 Then Exit Sub
    
    If LastTexture <> file Then
        Set rt = Textured_Get(file, RD)
    
        D3DDevice.SetTexture 0, rt
        LastTexture = file
    End If
    
    Dim x_center As Single, y_center As Single
    Dim radius As Single, temp As Single
    Dim right_point As Single, left_point As Single
    
    x_center = dX + (sW - dX) / 2
    y_center = dY + (sH - dY) / 2
    
    radius = Sqr((sW - x_center) ^ 2 + (sH - y_center) ^ 2)
    
    temp = (sW - x_center) / radius
    right_point = Atn(temp / Sqr(-temp * temp + 1))
    left_point = 3.1416 - right_point
    
    rvt(0).X = x_center + Cos(-left_point - angle) * radius + (sW / 4)
    rvt(0).Y = y_center - Sin(-left_point - angle) * radius + (sH / 4)
    rvt(0).Z = 0
    rvt(0).Rhw = 1
    rvt(0).color = IIf(color(0) = 0, AmbientColor, color(0))
    rvt(0).tU = sX / RD
    rvt(0).tV = (sY + sH + 1) / RD
        
    rvt(1).X = x_center + Cos(left_point - angle) * radius + (sW / 4)
    rvt(1).Y = y_center - Sin(left_point - angle) * radius + (sH / 4)
    rvt(1).Z = 0
    rvt(1).Rhw = 1
    rvt(1).color = IIf(color(1) = 0, AmbientColor, color(1))
    rvt(1).tU = sX / RD
    rvt(1).tV = sY / RD
    
    rvt(2).X = x_center + Cos(-right_point - angle) * radius + (sW / 4)
    rvt(2).Y = y_center - Sin(-right_point - angle) * radius + (sH / 4)
    rvt(2).Z = 0
    rvt(2).Rhw = 1
    rvt(2).color = IIf(color(2) = 0, AmbientColor, color(2))
    rvt(2).tU = (sX + sW + 1) / RD
    rvt(2).tV = (sY + sH + 1) / RD
    
    rvt(3).X = x_center + Cos(right_point - angle) * radius + (sW / 4)
    rvt(3).Y = y_center - Sin(right_point - angle) * radius + (sH / 4)
    rvt(3).Z = 0
    rvt(3).Rhw = 1
    rvt(3).color = IIf(color(3) = 0, AmbientColor, color(3))
    rvt(3).tU = (sX + sW + 1) / RD
    rvt(3).tV = sY / RD
    
     D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, rvt(0), FVFSize
End Sub
'#######################################################################################################
'#######################################################################################################
'#######################################################################################################
Public Sub Device_Init(ByVal hWnd As Long, ByVal w As Integer, ByVal h As Integer, Optional acc As Integer = D3DCREATE_SOFTWARE_VERTEXPROCESSING)
    'DX8
        Dim DispMode As D3DDISPLAYMODE
        Dim D3DWindow As D3DPRESENT_PARAMETERS

        Set dX = New DirectX8
        Set D3D = dX.Direct3DCreate()
        Set D3DX = New D3DX8
        
        D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
        With D3DWindow
            .Windowed = True
            .SwapEffect = D3DSWAPEFFECT_COPY
            .BackBufferFormat = DispMode.Format
            .BackBufferWidth = w
            .BackBufferHeight = h
            .hDeviceWindow = hWnd
        End With
        
        Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, _
                                hWnd, acc, D3DWindow)
                                                                
        'Audio.mSound_InitDirect dX, frmMain.hWnd
        
        D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
        D3DDevice.SetVertexShader FVF
        
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
        
        D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    'DX8
    
    sDefaultColor(0) = &HFFFFFFFF
    sDefaultColor(1) = &HFFFFFFFF
    sDefaultColor(2) = &HFFFFFFFF
    sDefaultColor(3) = &HFFFFFFFF

    sAlphaColor(0) = &H64646464
    sAlphaColor(1) = &H64646464
    sAlphaColor(2) = &H64646464
    sAlphaColor(3) = &H64646464
    
    On Error Resume Next
    Text_Init
    
End Sub

Public Sub Device_Render_Init()
On Error GoTo Err
    Dim r As Long
    
Inicio:
    D3DDevice.BeginScene
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0

    Exit Sub
Err:
    If r = 0 Then
        D3DDevice.EndScene
        r = 1
        GoTo Inicio
    Else
        Exit Sub
    End If
    
End Sub

Public Sub Device_Render_End(Optional ByVal hWnd As Long = 0, Optional ByVal w As Long, Optional ByVal h As Long)
    Dim a As RECT
    a.Left = 0
    a.Top = 0
    a.bottom = h
    a.Right = w
    
    D3DDevice.EndScene
    
    If hWnd = 0 Then
        D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 9
    Else
        D3DDevice.Present a, ByVal 0, hWnd, ByVal 0
    End If
End Sub

'############################
'#          Textos          #
'############################
Private Sub Text_Init()
    Dim tInfo As D3DXIMAGE_INFO

    Set texChar = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\Grh\font.bmp", _
                            D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, _
                            D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                            D3DX_FILTER_NONE, &HFF000000, tInfo, ByVal 0)

    Dim fD As Long, i As Long, Grh As Long
    fD = tInfo.Width

    Grh = 24031
    For i = 32 To 255
        With GrhData(Grh)
            chars(i).w = .PixelWidth
            chars(i).h = .PixelHeight
            chars(i).tu1 = .sX / fD
            chars(i).tv1 = .sY / fD
            chars(i).tu2 = (.sX + chars(i).w + 1) / fD
            chars(i).tv2 = (.sY + chars(i).h + 1) / fD
        End With
        Grh = Grh + 1
    Next i
    
    SegOn_Msg = Text_Convert("Seguro Activado")
    SegOff_Msg = Text_Convert("Seguro Desactivado")
    ModCom_Msg = Text_Convert("Modo Combate")
    Exit Sub
End Sub

Public Function Text_Width(ByRef txt As String) As Long
    Dim i As Long
    Dim Text() As Byte
    
    Text = StrConv(txt, vbFromUnicode)
    If Text(0) = 0 Then Exit Function
        
    For i = 0 To UBound(Text)
        If Text(i) <> 32 Then
            Text_Width = Text_Width + chars(Text(i)).w - 2
        Else
            Text_Width = Text_Width + 4
        End If
    Next i
End Function
Public Sub Text_Render(ByRef txt As String, ByVal dX As Long, ByVal dY As Long, ByVal color As Long)
    Dim i As Long, C As Byte
    Dim ii As Long
    Dim Text() As Byte
    
    Text = StrConv(txt, vbFromUnicode)
    
    For i = 0 To UBound(Text)
        C = Text(i)
        
        quads(ii).tV(0).X = dX
        
        quads(ii).tV(0).Y = dY + chars(C).h
        quads(ii).tV(0).Z = 0
        quads(ii).tV(0).Rhw = 1
        quads(ii).tV(0).color = color
        quads(ii).tV(0).tU = chars(C).tu1
        quads(ii).tV(0).tV = chars(C).tv2

        quads(ii).tV(1).X = dX
        quads(ii).tV(1).Y = dY
        quads(ii).tV(1).Z = 0
        quads(ii).tV(1).Rhw = 1
        quads(ii).tV(1).color = color
        quads(ii).tV(1).tU = chars(C).tu1
        quads(ii).tV(1).tV = chars(C).tv1

        quads(ii).tV(2).X = dX + chars(C).w
        quads(ii).tV(2).Y = dY + chars(C).h
        quads(ii).tV(2).Z = 0
        quads(ii).tV(2).Rhw = 1
        quads(ii).tV(2).color = color
        quads(ii).tV(2).tU = chars(C).tu2
        quads(ii).tV(2).tV = chars(C).tv2

        quads(ii).tV(3).X = dX + chars(C).w
        quads(ii).tV(3).Y = dY
        quads(ii).tV(3).Z = 0
        quads(ii).tV(3).Rhw = 1
        quads(ii).tV(3).color = color
        quads(ii).tV(3).tU = chars(C).tu2
        quads(ii).tV(3).tV = chars(C).tv1
        
        ii = ii + 1
        dX = dX + chars(C).w - 2
    Next i
    
    If LastTexture <> -65 Then
        LastTexture = -65
        D3DDevice.SetTexture 0, texChar
    End If
    
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, (4 * ii) - 2, quads(0).tV(0), FVFSize

End Sub
Public Function Text_Convert(ByRef Text As String) As Byte()
    Dim Ret() As Byte
    
    If Text = "" Then
        ReDim Ret(0)
    Else
        Ret = StrConv(Text, vbFromUnicode)
    End If
    
    Text_Convert = Ret
End Function

Public Sub Box_Draw_Porc(ByVal X As Integer, ByVal Y As Integer, _
                    ByVal w As Integer, ByVal h As Integer, _
                    ByVal porcW As Integer, _
                    ByVal outColor As Long, ByVal inColor As Long)
                    
    Dim outVertex(4) As LINEVERTEX

    D3DDevice.SetTexture 0, Nothing
    
    outVertex(0).X = X - 1
    outVertex(0).Y = Y + h + 1
    outVertex(0).Z = 1
    outVertex(0).Rhw = 1
    outVertex(0).color = outColor
    
    outVertex(1).X = X - 1
    outVertex(1).Y = Y - 1
    outVertex(1).Z = 1
    outVertex(1).Rhw = 1
    outVertex(1).color = outColor
    
    outVertex(2).X = X + w + 1
    outVertex(2).Y = Y - 1
    outVertex(2).Z = 1
    outVertex(2).Rhw = 1
    outVertex(2).color = outColor
    
    outVertex(3).X = X + w + 1
    outVertex(3).Y = Y + h + 1
    outVertex(3).Z = 1
    outVertex(3).Rhw = 1
    outVertex(3).color = outColor
    
    outVertex(4).X = X - 1
    outVertex(4).Y = Y + h + 1
    outVertex(4).Z = 1
    outVertex(4).Rhw = 1
    outVertex(4).color = outColor
    
    D3DDevice.SetVertexShader LFVF
    D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, 4, outVertex(0), Len(outVertex(0))
    
    Dim inVertex(3) As LINEVERTEX
    
    w = (w * porcW) / 100
    
    inVertex(0).X = X
    inVertex(0).Y = Y + h
    inVertex(0).Z = 0
    inVertex(0).Rhw = 1
    inVertex(0).color = inColor
    
    inVertex(1).X = X
    inVertex(1).Y = Y
    inVertex(1).Z = 0
    inVertex(1).Rhw = 1
    inVertex(1).color = inColor
    
    inVertex(2).X = X + w
    inVertex(2).Y = Y + h
    inVertex(2).Z = 0
    inVertex(2).Rhw = 1
    inVertex(2).color = inColor
    
    inVertex(3).X = X + w
    inVertex(3).Y = Y
    inVertex(3).Z = 0
    inVertex(3).Rhw = 1
    inVertex(3).color = inColor
    
    D3DDevice.SetVertexShader FVF
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, inVertex(0), Len(inVertex(0))
    
    If LastTexture > 0 Then
        D3DDevice.SetTexture 0, Textured_Get(LastTexture, 0)
    End If
End Sub


Public Sub Box_Draw(ByVal X As Integer, ByVal Y As Integer, _
                    ByVal w As Integer, ByVal h As Integer, _
                    ByVal color As Long)
                    
    Dim inVertex(3) As LINEVERTEX

    inVertex(0).X = X
    inVertex(0).Y = Y + h
    inVertex(0).Z = 0
    inVertex(0).Rhw = 1
    inVertex(0).color = color
    
    inVertex(1).X = X
    inVertex(1).Y = Y
    inVertex(1).Z = 0
    inVertex(1).Rhw = 1
    inVertex(1).color = color
    
    inVertex(2).X = X + w
    inVertex(2).Y = Y + h
    inVertex(2).Z = 0
    inVertex(2).Rhw = 1
    inVertex(2).color = color
    
    inVertex(3).X = X + w
    inVertex(3).Y = Y
    inVertex(3).Z = 0
    inVertex(3).Rhw = 1
    inVertex(3).color = color
    
    D3DDevice.SetTexture 0, Nothing
    D3DDevice.SetVertexShader FVF
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, inVertex(0), Len(inVertex(0))
    
    If LastTexture > 0 Then
        D3DDevice.SetTexture 0, Textured_Get(LastTexture, 0)
    End If
End Sub



