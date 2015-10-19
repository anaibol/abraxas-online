Attribute VB_Name = "modLoadData"
Option Explicit
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public SelecOX As Integer
Public SelecOY As Integer
Public SelecDX As Integer
Public SelecDY As Integer
Public SelecAC As Integer

Public numGrh As Long
'Saved
Public Function General_Random_Number(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    Randomize Timer
    General_Random_Number = (UpperBound - LowerBound) * Rnd + LowerBound
End Function
Public Function LoadGrhData() As Boolean
On Error Resume Next
    Dim handle As Integer
    Dim Grh As Long
    Dim Frame As Integer
    Dim fileVer As Long

    'Open files
    handle = FreeFile()
    Open DirInit & "Grh.ind" For Binary Access Read As handle
        Seek #1, 1
        
        Get handle, , fileVer
        
        Get handle, , numGrh
        'Resize arrays
        ReDim GrhData(0 To numGrh) As GrhData
        
        'Get first Grh Number
        Get handle, , Grh
        While Grh < numGrh And Grh > 0
            With GrhData(Grh)
                'Get number of frames
                .active = True
                
                Get handle, , .NumFrames
                If .NumFrames <= 0 Then Resume Next
                
                ReDim .Frames(1 To GrhData(Grh).NumFrames)
                
                If .NumFrames > 1 Then
                    'Read a animation GRH set
                    For Frame = 1 To .NumFrames
                        Get handle, , .Frames(Frame)
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > UBound(GrhData) Then
                            Resume Next
                        End If
                    Next Frame
                    
                    Get handle, , .Speed

                    'Compute width and height
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then Resume Next
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then Resume Next
                    
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then Resume Next
                    
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then Resume Next
                Else
                    'Read in normal GRH data
                    Get handle, , .FileNum
                    If .FileNum <= 0 Then Resume Next
                    
                    Get handle, , .sX
                    If .sX < 0 Then Resume Next
                    
                    Get handle, , .sY
                    If .sY < 0 Then Resume Next
                    
                    Get handle, , .pixelWidth
                    If .pixelWidth <= 0 Then Resume Next
                    
                    Get handle, , .pixelHeight
                    If .pixelHeight <= 0 Then Resume Next
                    
                    'Compute width and height
                    .TileWidth = .pixelWidth / 32
                    .TileHeight = .pixelHeight / 32
                    
                    .Frames(1) = Grh
                End If
            End With
            
            Get handle, , Grh
        Wend
        
    Close handle
    

Dim Count As Long

Open DirInit & "minimap.ind" For Binary As #1
    Seek #1, 1
    For Count = 1 To numGrh
        Get #1, , GrhData(Count).MiniMap_color
    Next Count
Close #1
   
    LoadGrhData = True
Exit Function

ErrorHandler:
    LoadGrhData = False
End Function
Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open DirInit & "Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call Grh_Init(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call Grh_Init(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call Grh_Init(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call Grh_Init(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCascos()
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open DirInit & "Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call Grh_Init(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call Grh_Init(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call Grh_Init(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call Grh_Init(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open DirInit & "Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            Grh_Init BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            Grh_Init BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            Grh_Init BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            Grh_Init BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
End Sub

Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    N = FreeFile()
    Open DirInit & "Fx.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
End Sub



Public Sub DibujarMiniMapa()
On Error Resume Next
Dim map_x As Long, map_y As Long
frmMinimap.MiniMap.Cls

    For map_y = 1 To 100
        For map_x = 1 To 100
            If MapData(map_x, map_y).Graphic(1).GrhIndex > 0 Then
                If GrhData(MapData(map_x, map_y).Graphic(1).GrhIndex).MiniMap_color <> 0 Then SetPixel frmMinimap.MiniMap.hdc, map_x, map_y, GrhData(MapData(map_x, map_y).Graphic(1).GrhIndex).MiniMap_color
            End If

            If MapData(map_x, map_y).Graphic(2).GrhIndex > 0 Then
                If GrhData(MapData(map_x, map_y).Graphic(2).GrhIndex).MiniMap_color <> 0 Then SetPixel frmMinimap.MiniMap.hdc, map_x, map_y, GrhData(MapData(map_x, map_y).Graphic(2).GrhIndex).MiniMap_color
            End If
            
            If MapData(map_x, map_y).ObjGrh.GrhIndex <> 0 Then
                If GrhData(MapData(map_x, map_y).ObjGrh.GrhIndex).MiniMap_color <> 0 Then SetPixel frmMinimap.MiniMap.hdc, map_x, map_y, GrhData(MapData(map_x, map_y).ObjGrh.GrhIndex).MiniMap_color
            End If
            
            If MapData(map_x, map_y).Graphic(3).GrhIndex <> 0 Then
                If GrhData(MapData(map_x, map_y).Graphic(3).GrhIndex).MiniMap_color <> 0 Then SetPixel frmMinimap.MiniMap.hdc, map_x, map_y, GrhData(MapData(map_x, map_y).Graphic(4).GrhIndex).MiniMap_color
            End If
            
            If MapData(map_x, map_y).Graphic(4).GrhIndex <> 0 And Not bTecho Then
                If GrhData(MapData(map_x, map_y).Graphic(4).GrhIndex).MiniMap_color <> 0 Then SetPixel frmMinimap.MiniMap.hdc, map_x, map_y, GrhData(MapData(map_x, map_y).Graphic(4).GrhIndex).MiniMap_color
            End If
            
        Next map_x
    Next map_y
    
    frmMinimap.MiniMap.Refresh
    
End Sub
Public Sub CargarIndicesSuperficie()

On Error GoTo Fallo
    If FileExist(App.Path & "\indices.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'indices.ini'", vbCritical
        End
    End If
    
    Dim Leer As New clsIniReader
    Dim i As Integer
    
    Leer.Initialize App.Path & "\indices.ini"
    MaxSup = Leer.GetValue("INIT", "Referencias")
    
    ReDim SupData(MaxSup) As SupData
    
    For i = 0 To MaxSup
        SupData(i).name = Leer.GetValue("REFERENCIA" & i, "Name")
        SupData(i).Grh = Val(Leer.GetValue("REFERENCIA" & i, "GrhIndice"))
        SupData(i).Width = Val(Leer.GetValue("REFERENCIA" & i, "Ancho"))
        SupData(i).Height = Val(Leer.GetValue("REFERENCIA" & i, "Alto"))
        SupData(i).Block = IIf(Val(Leer.GetValue("REFERENCIA" & i, "Bloquear")) = 1, True, False)
        SupData(i).Capa = Val(Leer.GetValue("REFERENCIA" & i, "Capa"))

        frmMode.lstSuperfices.AddItem SupData(i).name & " - #" & i
    Next
    
    frmMode.lstCapa.ListIndex = 0

    DoEvents
    
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el indice " & i & " de " & App.Path & "\indices.ini" & vbCrLf & "Err: " & err.Number & " - " & err.Description, vbCritical + vbOKOnly
End Sub
